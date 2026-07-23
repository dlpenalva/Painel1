"""Geradores de documentos administrativos/juridicos em DOCX.

Gera o Despacho Saneador e a Minuta de Termo de Apostilamento a partir dos
dados canonicos do Objeto Processo de Reajuste. Nao recalcula valores — apenas
apresenta os dados ja consolidados pelos motores oficiais, em LINGUAGEM
ADMINISTRATIVA (nunca expoe vocabulario de implementacao do XLS/Python).

Campos manuais ausentes recebem o marcador [PREENCHER: <descricao>] com
destaque amarelo. Ausencia de dado automatico nunca vira zero.

Nenhum arquivo entregue pode conter emoji/pictograma (sanitizacao no output).

Interface publica:
    gerar_despacho_saneador(leitura_ou_objeto, identificacao, campos_manuais) -> bytes
    gerar_termo_apostila(leitura_ou_objeto, identificacao, campos_manuais) -> bytes
    diagnosticar_campos_manuais(leitura_ou_objeto, identificacao, campos_manuais) -> list[dict]
"""
from __future__ import annotations

from io import BytesIO
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from _sumario_executivo import (
    NAO_INFORMADO,
    formatar_moeda,
    montar_dados_sumario_executivo,
    _num_ou_none,
)
from _objeto_processo_reajuste import obter_objeto_processo_reajuste
from _sanitizacao_documental import remover_emojis_leve

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

PREENCHER_TAG = "[PREENCHER: {}]"
COR_NEGATIVO = RGBColor(0xC0, 0x00, 0x00)
_LETRAS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def _fmt_pct_doc(valor: Any) -> str:
    """Formata percentual documental com exatamente duas casas decimais.

    Entrada no formato decimal canonico (0.0421 -> '4,21%').
    """
    numero = _num_ou_none(valor)
    if numero is None:
        return NAO_INFORMADO
    texto = f"{numero * 100:.2f}".replace(".", ",")
    return f"{texto}%"


def _indice_amigavel_doc(indice: Any) -> str | None:
    """Nome amigavel do indice, sem expor codigo tecnico (SGS-433/189/DIMAC).

    Retorna None quando indefinido, para que o chamador use marcador manual.
    """
    texto = remover_emojis_leve(indice).strip()
    if not texto or texto == NAO_INFORMADO:
        return None
    norm = texto.upper()
    if norm.startswith("IST"):
        return "IST (Série Local)"
    if norm.startswith("ICTI"):
        return "ICTI (Ipeadata)"
    if norm.startswith("IPCA"):
        return "IPCA"
    if norm.startswith("IGP"):
        return "IGP-M"
    if norm.startswith("INPC"):
        return "INPC"
    import re as _re
    limpo = _re.sub(r"\s*\[[^\]]*\]\s*", " ", texto)      # remove "[SGS-433]"
    limpo = _re.sub(r"\s*\(\s*\d+\s*\)\s*$", "", limpo)   # remove "(433)"
    return limpo.strip() or None


CAMPOS_MANUAIS_DESPACHO = [
    ("contrato", "Numero do contrato", "despacho"),
    ("processo_pleito", "Referencias do pleito da contratada", "despacho"),
    ("data_proposta", "Data da proposta (para verificacao da anualidade)", "despacho"),
    ("referencia_analise", "Referencia onde o resultado da analise consta", "despacho"),
    ("data_corte_descricao", "Descricao da data/posicao de corte adotada", "despacho"),
    ("adequacao_orcamentaria_ref", "Referencia da adequacao orcamentaria", "despacho"),
    ("adequacao_orcamentaria_valor", "Valor da adequacao orcamentaria", "despacho"),
    ("regularidade_ref", "Referencia das certidoes de regularidade", "despacho"),
    ("concordancia_ref", "Referencia da manifestacao de concordancia da contratada", "despacho"),
    ("docs_desatualizados", "Lista de documentos a desconsiderar (opcional)", "despacho"),
    ("valor_original_contrato", "Valor original do contrato", "despacho"),
]

CAMPOS_MANUAIS_TERMO = [
    ("contrato", "Numero do contrato", "termo"),
    ("empresa_contratada", "Nome/qualificacao da empresa contratada", "termo"),
    ("representante_telebras_1_nome", "Nome do 1o representante da Telebras", "termo"),
    ("representante_telebras_1_matricula", "Matricula do 1o representante", "termo"),
    ("representante_telebras_2_cargo", "Cargo do 2o representante da Telebras", "termo"),
    ("representante_telebras_2_matricula", "Matricula do 2o representante", "termo"),
    ("memoria_calculo_ref", "Referencia da memoria de calculo", "termo"),
    ("concordancia_ref", "Referencia da manifestacao de concordancia da contratada", "termo"),
    ("regularidade_ref", "Referencia das certidoes de regularidade", "termo"),
    ("adequacao_orcamentaria_ref", "Referencia da adequacao orcamentaria", "termo"),
    ("processo_ref", "Numero do processo de instrucao", "termo"),
    ("valor_pago_efetivo", "Valor pago efetivo (quando nao apurado automaticamente)", "termo"),
    ("valor_teorico", "Valor teorico calculado (quando nao apurado automaticamente)", "termo"),
    ("valor_original_contrato", "Valor original do contrato", "termo"),
    ("local_data", "Data (ex.: 20/07/2026)", "termo"),
]

TODOS_CAMPOS_MANUAIS = list(
    {c[0]: c for c in CAMPOS_MANUAIS_DESPACHO + CAMPOS_MANUAIS_TERMO}.values()
)

# Campos que sao opcionais (nao entram como pendencia critica no diagnostico).
_CAMPOS_OPCIONAIS = {
    "docs_desatualizados", "valor_pago_efetivo", "valor_teorico",
}


# ---------------------------------------------------------------------------
# Helpers XML / DOCX
# ---------------------------------------------------------------------------

def _set_highlight(run, cor: str = "yellow") -> None:
    rPr = run._r.get_or_add_rPr()
    highlight = OxmlElement("w:highlight")
    highlight.set(qn("w:val"), cor)
    rPr.append(highlight)


def _repetir_cabecalho(tabela) -> None:
    tr = tabela.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    trPr.append(tblHeader)


def _adicionar_run(p, texto: str, negrito: bool = False, tamanho: int = 11,
                   cor: RGBColor | None = None, italico: bool = False) -> Any:
    run = p.add_run(remover_emojis_leve(texto))
    run.bold = negrito
    run.italic = italico
    run.font.name = "Calibri"
    run.font.size = Pt(tamanho)
    if cor:
        run.font.color.rgb = cor
    return run


def _titulo_secao(doc: Document, texto: str, tamanho: int = 11,
                  alinhamento=WD_ALIGN_PARAGRAPH.LEFT) -> Any:
    p = doc.add_paragraph()
    p.alignment = alinhamento
    _adicionar_run(p, texto, negrito=True, tamanho=tamanho)
    return p


def _titulo_quadro(doc: Document, texto: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _adicionar_run(p, texto, negrito=True, tamanho=10)


def _run_campo_manual(p, descricao: str, tamanho: int = 11) -> Any:
    run = p.add_run(PREENCHER_TAG.format(descricao))
    run.font.name = "Calibri"
    run.font.size = Pt(tamanho)
    _set_highlight(run, "yellow")
    return run


def _texto_ou_marcador(p, valor: Any, descricao: str, tamanho: int = 11,
                        negrito: bool = False, cor: RGBColor | None = None) -> None:
    if valor is not None and str(valor).strip():
        run = p.add_run(remover_emojis_leve(valor))
        run.bold = negrito
        run.font.name = "Calibri"
        run.font.size = Pt(tamanho)
        if cor:
            run.font.color.rgb = cor
    else:
        _run_campo_manual(p, descricao, tamanho)


def _campo(campos_manuais: dict, chave: str) -> Any:
    if not campos_manuais:
        return None
    v = campos_manuais.get(chave)
    if v is None:
        return None
    if isinstance(v, str) and not v.strip():
        return None
    return v


def _valor_moeda_ou_marcador(p, valor: Any, descricao: str, tamanho: int = 11) -> None:
    numero = _num_ou_none(valor)
    if numero is not None:
        run = p.add_run(formatar_moeda(numero))
        run.font.name = "Calibri"
        run.font.size = Pt(tamanho)
        if numero < 0:
            run.font.color.rgb = COR_NEGATIVO
    else:
        _run_campo_manual(p, descricao, tamanho)


def _configurar_documento() -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)
    return doc


def _adicionar_tabela(doc: Document, cabecalho: list[str],
                       linhas: list[list[str]]) -> Any:
    n_cols = len(cabecalho)
    tabela = doc.add_table(rows=1, cols=n_cols)
    tabela.style = "Table Grid"
    celulas_cab = tabela.rows[0].cells
    for i, texto in enumerate(cabecalho):
        celulas_cab[i].text = ""
        run = celulas_cab[i].paragraphs[0].add_run(remover_emojis_leve(texto))
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(10)
    _repetir_cabecalho(tabela)
    for linha in linhas:
        row = tabela.add_row()
        for i, celula_texto in enumerate(linha):
            row.cells[i].text = ""
            texto = remover_emojis_leve(celula_texto)
            negativo = False
            try:
                val_num = float(
                    str(celula_texto).replace("R$ ", "").replace(".", "").replace(",", ".")
                )
                negativo = val_num < 0
            except (ValueError, AttributeError):
                pass
            run = row.cells[i].paragraphs[0].add_run(texto)
            run.font.name = "Calibri"
            run.font.size = Pt(10)
            if negativo and "R$" in str(celula_texto):
                run.font.color.rgb = COR_NEGATIVO
    return tabela


# ---------------------------------------------------------------------------
# Extracao de dados canonicos
# ---------------------------------------------------------------------------

def _extrair_dados(leitura_ou_objeto: dict, identificacao: dict | None) -> dict:
    dados = montar_dados_sumario_executivo(leitura_ou_objeto, identificacao)
    if not dados.get("disponivel"):
        return {"disponivel": False}

    ciclos = dados.get("ciclos") or []
    ciclos_reajuste = [c for c in ciclos if not c.get("eh_base")]
    ciclos_computados = [c for c in ciclos_reajuste if c.get("computar") == "Sim"]

    financeiro = dados.get("financeiro") or {}
    sintese = dados.get("sintese") or {}
    aditivos_raw = (dados.get("aditivos") or {}).get("itens") or []

    fin_por_ciclo = {r["ciclo"]: r for r in financeiro.get("financeiro_por_ciclo") or []}
    pc_por_ciclo = {r["ciclo"]: r for r in financeiro.get("pc_por_ciclo") or []}

    objeto_proc = obter_objeto_processo_reajuste(leitura_ou_objeto) or {}
    dados_op = objeto_proc.get("dados_operacionais") or {}
    vta_sombra = dados_op.get("vta_sombra") or {}
    parcelas_vta = vta_sombra.get("parcelas_computadas") or []

    aditivos = []
    for ad in aditivos_raw:
        aditivos.append({
            "identificador": ad.get("identificador") or ad.get("ciclo"),
            "ciclo": ad.get("ciclo"),
            "valor_atualizado": ad.get("valor_atualizado"),
            "anterior_formalizacao": ad.get("anterior_formalizacao"),
        })

    return {
        "disponivel": True,
        "ciclos": ciclos,
        "ciclos_reajuste": ciclos_reajuste,
        "ciclos_computados": ciclos_computados,
        "var_acumulada": sintese.get("variacao_acumulada"),
        "vta": sintese.get("vta"),
        "fin_por_ciclo": fin_por_ciclo,
        "pc_por_ciclo": pc_por_ciclo,
        "parcelas_vta": parcelas_vta,
        "aditivos": aditivos,
        "financeiro": financeiro,
        "sintese": sintese,
        "identificacao": dados.get("identificacao") or {},
        "historico_vu": dados.get("historico_vu") or {},
    }


def _retroativo_total(dados: dict) -> float | None:
    fin = dados.get("financeiro") or {}
    t_fin = fin.get("delta_total_financeiro")
    if t_fin is not None:
        return t_fin
    return fin.get("delta_total_pc")


def _linhas_financeiro(dados: dict) -> list[dict]:
    if dados.get("fin_por_ciclo"):
        return list(dados["fin_por_ciclo"].values())
    if dados.get("pc_por_ciclo"):
        return list(dados["pc_por_ciclo"].values())
    return []


def _valor_pago_total(dados: dict) -> float | None:
    linhas = _linhas_financeiro(dados)
    if not linhas:
        return None
    return round(sum(_num_ou_none(l.get("valor_pago")) or 0.0 for l in linhas), 2)


def _valor_atualizado_total(dados: dict) -> float | None:
    linhas = _linhas_financeiro(dados)
    if not linhas:
        return None
    return round(sum(_num_ou_none(l.get("valor_atualizado")) or 0.0 for l in linhas), 2)


def _indice_doc(dados: dict) -> str | None:
    return _indice_amigavel_doc((dados.get("identificacao") or {}).get("indice"))


def _efeito_financeiro_ciclo(c: dict) -> str:
    """Frase administrativa de efeitos financeiros de um ciclo."""
    situacao = remover_emojis_leve(c.get("situacao") or "").strip().lower()
    inicio = str(c.get("inicio_efeito_financeiro") or "").strip()
    if "preclu" in situacao:
        return "Sem efeitos financeiros"
    if inicio and inicio != NAO_INFORMADO and "/" in inicio:
        return f"A partir de {inicio}"
    return NAO_INFORMADO


# ---------------------------------------------------------------------------
# Camada de apresentacao humanizada do VTA (nunca expoe vocabulario do XLS)
# ---------------------------------------------------------------------------

def _descricao_vta_humana(parcela: dict) -> str:
    """Traduz a parcela do VTA para linguagem administrativa."""
    fonte = str(parcela.get("fonte_parcela") or "").strip().lower()
    ciclo = remover_emojis_leve(parcela.get("ciclo") or "").strip().upper()
    if "aditivo" in fonte or "supress" in fonte:
        return f"Aditivo/supressão computável ({ciclo})" if ciclo else "Aditivo/supressão computável"
    if "remanesc" in fonte or "residual" in fonte or "saldo" in fonte:
        return "Saldo remanescente atualizado"
    if ciclo:
        return f"{ciclo} - execução atualizada"
    return "Parcela de composição do Valor Total Atualizado"


def _composicao_didatica_vta(dados: dict) -> list[tuple[str, float | None]]:
    """Agrupa as parcelas em componentes didaticos (execucao por ciclo, saldo,
    aditivos), somando por rubrica. Nunca inventa; apenas soma o que existe.
    """
    parcelas = dados.get("parcelas_vta") or []
    grupos: dict[str, float | None] = {}
    ordem: list[str] = []
    for p in parcelas:
        desc = _descricao_vta_humana(p)
        valor = _num_ou_none(p.get("valor_atualizado"))
        if valor is None:
            valor = _num_ou_none(p.get("valor"))
        if desc not in grupos:
            grupos[desc] = None
            ordem.append(desc)
        if valor is not None:
            grupos[desc] = (grupos[desc] or 0.0) + valor
    return [(d, grupos[d]) for d in ordem]


def _secao_valores_unitarios_por_ciclo(doc: Document, dados: dict) -> None:
    """Tabela 'Valores Unitarios por Ciclo': C0..ultimo ciclo analisado.

    Nunca inventa zeros: celula sem valor sai vazia. C0 = valor unitario
    original. Ciclos futuros nao entram.
    """
    hvu = dados.get("historico_vu") or {}
    itens = hvu.get("itens") or []
    ciclos = hvu.get("ciclos") or []
    if not itens or not ciclos:
        return

    _titulo_quadro(doc, "Quadro de Valores Unitários por Ciclo")
    cabecalho = ["Item", "Descrição"] + [f"VU {c}" for c in ciclos]
    linhas: list[list[str]] = []
    for reg in itens:
        vus = reg.get("vus") or {}
        linha = [str(reg.get("item") or ""), str(reg.get("descricao") or "")]
        for c in ciclos:
            valor = vus.get(c)
            linha.append(formatar_moeda(valor) if valor is not None else "")
        linhas.append(linha)
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


# ---------------------------------------------------------------------------
# TERMO DE APOSTILAMENTO (modelo canonico §6)
# ---------------------------------------------------------------------------

def gerar_termo_apostila(
    leitura_ou_objeto: dict,
    identificacao: dict | None = None,
    campos_manuais: dict | None = None,
) -> bytes:
    """Gera a Minuta de Termo de Apostilamento em DOCX e retorna os bytes."""
    if campos_manuais is None:
        campos_manuais = {}
    dados = _extrair_dados(leitura_ou_objeto, identificacao)
    doc = _configurar_documento()

    _ta_titulo(doc, campos_manuais)
    _ta_qualificacao(doc, campos_manuais)
    _ta_considerandos(doc, dados, campos_manuais)
    _ta_abertura(doc)
    _ta_secao1_reajustes(doc, dados, campos_manuais)
    _ta_secao2_retroativo(doc, dados, campos_manuais)
    _ta_secao3_memoria_vta(doc, dados)
    _ta_secao4_composicao_vta(doc, dados)
    _ta_secao5_valores_unitarios(doc, dados)
    _ta_secao6_aditivos(doc, dados)
    _ta_secoes_finais(doc, campos_manuais)
    _ta_assinaturas(doc, campos_manuais)

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _ta_titulo(doc: Document, cm: dict) -> None:
    _titulo_secao(doc, "MINUTA DE TERMO DE APOSTILAMENTO", tamanho=12,
                  alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _adicionar_run(p, "Contrato nº ")
    _texto_ou_marcador(p, _campo(cm, "contrato"), "Numero do contrato")
    doc.add_paragraph()


def _ta_qualificacao(doc: Document, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p,
        "A TELECOMUNICAÇÕES BRASILEIRAS S.A. - TELEBRAS, sociedade de economia "
        "mista, vinculada ao Ministério das Comunicações, com sede no SIG, "
        "Quadra 04, Bloco A, Salas 201 a 224, Edifício Capital Financial Center, "
        "CEP nº 70.610-440, inscrita no CNPJ sob o n.º 00.336.701/0001-04, "
        "doravante denominada TELEBRAS, neste ato representada por ")
    _texto_ou_marcador(p, _campo(cm, "representante_telebras_1_nome"), "Nome do 1o representante da Telebras")
    _adicionar_run(p, ", Matrícula ")
    _texto_ou_marcador(p, _campo(cm, "representante_telebras_1_matricula"), "Matricula do 1o representante")
    _adicionar_run(p, ", e por seu ")
    _texto_ou_marcador(p, _campo(cm, "representante_telebras_2_cargo"), "Cargo do 2o representante da Telebras")
    _adicionar_run(p, ", Matrícula ")
    _texto_ou_marcador(p, _campo(cm, "representante_telebras_2_matricula"), "Matricula do 2o representante")
    _adicionar_run(p, ", nos termos da Diretriz nº 229/2018, apostila o Contrato nº ")
    _texto_ou_marcador(p, _campo(cm, "contrato"), "Numero do contrato")
    _adicionar_run(p, ", celebrado com a empresa ")
    _texto_ou_marcador(p, _campo(cm, "empresa_contratada"), "Nome/qualificacao da empresa contratada")
    _adicionar_run(p,
        ", doravante denominada CONTRATADA, com fundamento no parágrafo 7º do "
        "art. 81 da Lei nº 13.303, de 30 de junho de 2016, na legislação "
        "aplicável, no Regulamento de Licitações e Contratos da Telebras e nos "
        "documentos constantes do processo.")
    doc.add_paragraph()


def _ta_considerandos(doc: Document, dados: dict, cm: dict) -> None:
    _titulo_secao(doc, "CONSIDERANDO:")

    def item(numero: str):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _adicionar_run(p, f"{numero}. ", negrito=True)
        return p

    p1 = item("1")
    _adicionar_run(p1, "A Cláusula Oitava do Contrato nº ")
    _texto_ou_marcador(p1, _campo(cm, "contrato"), "Numero do contrato")
    _adicionar_run(p1,
        ", que disciplina o reajuste contratual, os ciclos de apuração, a "
        "admissibilidade dos pedidos e os respectivos efeitos financeiros;")

    p2 = item("2")
    _adicionar_run(p2,
        "A necessidade de distinguir o histórico já formalizado anteriormente do "
        "objeto da presente análise, evitando duplicidade de contagem ou "
        "sobreposição de efeitos financeiros;")

    p3 = item("3")
    _adicionar_run(p3, "A memória de cálculo constante em ")
    _texto_ou_marcador(p3, _campo(cm, "memoria_calculo_ref"), "Referencia da memoria de calculo")
    _adicionar_run(p3,
        ", que apurou os ciclos de reajuste, os percentuais aplicáveis, os "
        "efeitos financeiros, o saldo retroativo a pagar e a composição do Valor "
        "Total Atualizado do Contrato;")

    p4 = item("4")
    _adicionar_run(p4, "O índice contratual utilizado na análise, qual seja ")
    indice = _indice_doc(dados)
    if indice:
        _adicionar_run(p4, indice, negrito=True)
    else:
        _run_campo_manual(p4, "Indice contratual")
    _adicionar_run(p4, ", e o percentual acumulado apurado de ")
    var = dados.get("var_acumulada")
    if var is not None:
        _adicionar_run(p4, _fmt_pct_doc(var), negrito=True)
    else:
        _run_campo_manual(p4, "Percentual acumulado apurado")
    _adicionar_run(p4, ";")

    p5 = item("5")
    _adicionar_run(p5,
        "As informações encaminhadas pela área gestora/fiscal do contrato quanto "
        "à execução, ao saldo remanescente, aos itens contratuais, aos "
        "aditivos/supressões e aos documentos de suporte da apuração;")

    p6 = item("6")
    _adicionar_run(p6, "A manifestação de concordância da CONTRATADA constante em ")
    _texto_ou_marcador(p6, _campo(cm, "concordancia_ref"), "Referencia da manifestacao de concordancia da contratada")
    _adicionar_run(p6, ", quando aplicável;")

    p7 = item("7")
    _adicionar_run(p7,
        "As certidões de regularidade da CONTRATADA e a adequação orçamentária "
        "constantes dos autos, quando aplicável.")
    doc.add_paragraph()


def _ta_abertura(doc: Document) -> None:
    _titulo_secao(doc, "FORMALIZA-SE O PRESENTE TERMO DE APOSTILA:")
    doc.add_paragraph()


def _ta_secao1_reajustes(doc: Document, dados: dict, cm: dict) -> None:
    _titulo_secao(doc, "1. Dos reajustes concedidos")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "1.1. Ao Contrato nº ")
    _texto_ou_marcador(p, _campo(cm, "contrato"), "Numero do contrato")
    _adicionar_run(p,
        ", formalizam-se os reajustes contratuais apurados, conforme Quadro 1.")

    _titulo_quadro(doc, "Quadro 1 — Síntese dos reajustes concedidos")
    cabecalho = ["Ref.", "Ciclo", "Percentual aplicado", "Efeitos financeiros", "Situação"]
    linhas: list[list[str]] = []
    ciclos = dados.get("ciclos_computados") or []
    for i, c in enumerate(ciclos):
        pct = c.get("percentual_reajuste")
        linhas.append([
            _LETRAS[i] if i < len(_LETRAS) else str(i + 1),
            remover_emojis_leve(c.get("ciclo") or ""),
            _fmt_pct_doc(pct) if pct is not None else NAO_INFORMADO,
            _efeito_financeiro_ciclo(c),
            remover_emojis_leve(c.get("situacao") or NAO_INFORMADO),
        ])
    ref_acum = _LETRAS[len(ciclos)] if len(ciclos) < len(_LETRAS) else "Acum."
    var = dados.get("var_acumulada")
    linhas.append([
        ref_acum,
        "Acumulado",
        _fmt_pct_doc(var) if var is not None else NAO_INFORMADO,
        "Conforme composição dos ciclos",
        "Percentual acumulado apurado",
    ])
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ta_secao2_retroativo(doc: Document, dados: dict, cm: dict) -> None:
    _titulo_secao(doc, "2. Da apuração financeira do retroativo")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "2.1. A apuração financeira consolidada indicou valor pago efetivo de ")
    vp = _valor_pago_total(dados)
    if vp is not None:
        _adicionar_run(p, formatar_moeda(vp), negrito=True)
    else:
        _valor_moeda_ou_marcador(p, _campo(cm, "valor_pago_efetivo"), "Valor pago efetivo")
    _adicionar_run(p, ", valor teórico calculado de ")
    vat = _valor_atualizado_total(dados)
    if vat is not None:
        _adicionar_run(p, formatar_moeda(vat), negrito=True)
    else:
        _valor_moeda_ou_marcador(p, _campo(cm, "valor_teorico"), "Valor teorico calculado")
    _adicionar_run(p, ", resultando em valor retroativo a pagar de ")
    retro = _retroativo_total(dados)
    if retro is not None:
        _adicionar_run(p, formatar_moeda(retro), negrito=True)
    else:
        _run_campo_manual(p, "Valor retroativo a pagar")
    _adicionar_run(p, ", conforme Quadro 2.")

    _titulo_quadro(doc, "Quadro 2 — Apuração financeira por ciclo")
    cabecalho = ["Ciclo", "Valor pago efetivo", "Valor teórico calculado", "Diferença/retroativo"]
    linhas: list[list[str]] = []
    tot_pago = tot_teorico = tot_delta = None
    for lin in _linhas_financeiro(dados):
        vpg = _num_ou_none(lin.get("valor_pago"))
        vtc = _num_ou_none(lin.get("valor_atualizado"))
        vdl = _num_ou_none(lin.get("delta"))
        linhas.append([
            remover_emojis_leve(lin.get("ciclo") or ""),
            formatar_moeda(vpg) if vpg is not None else "",
            formatar_moeda(vtc) if vtc is not None else "",
            formatar_moeda(vdl) if vdl is not None else "",
        ])
        if vpg is not None:
            tot_pago = (tot_pago or 0.0) + vpg
        if vtc is not None:
            tot_teorico = (tot_teorico or 0.0) + vtc
        if vdl is not None:
            tot_delta = (tot_delta or 0.0) + vdl
    if not linhas:
        linhas = [["—", "", "", ""]]
    total_delta = tot_delta if tot_delta is not None else _retroativo_total(dados)
    linhas.append([
        "Total",
        formatar_moeda(tot_pago) if tot_pago is not None else "",
        formatar_moeda(tot_teorico) if tot_teorico is not None else "",
        formatar_moeda(total_delta) if total_delta is not None else "",
    ])
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ta_secao3_memoria_vta(doc: Document, dados: dict) -> None:
    _titulo_secao(doc, "3. Da memória fiscal do Valor Total Atualizado")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p,
        "3.1. Para fins de consolidação contratual, a memória fiscal do Valor "
        "Total Atualizado foi organizada de forma evolutiva, demonstrando a "
        "execução por ciclo, os remanescentes intermediários, o saldo "
        "remanescente final e os aditivos/supressões computáveis, quando "
        "aplicáveis.")

    _titulo_quadro(doc, "Quadro 3 — Memória fiscal do Valor Total Atualizado")
    cabecalho = ["Ref.", "Descrição", "Valor"]
    linhas: list[list[str]] = []
    for i, (desc, valor) in enumerate(_composicao_didatica_vta(dados)):
        linhas.append([
            _LETRAS[i] if i < len(_LETRAS) else str(i + 1),
            desc,
            formatar_moeda(valor) if valor is not None else "",
        ])
    vta = dados.get("vta")
    linhas.append([
        "Total",
        "Valor Total Atualizado do Contrato",
        formatar_moeda(vta) if vta is not None else "",
    ])
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ta_secao4_composicao_vta(doc: Document, dados: dict) -> None:
    _titulo_secao(doc, "4. Da composição sintética do Valor Total Atualizado")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p,
        "4.1. De forma sintética, o Valor Total Atualizado do Contrato pode ser "
        "compreendido pela soma das parcelas indicadas no Quadro 4.")

    componentes = _composicao_didatica_vta(dados)
    _titulo_quadro(doc, "Quadro 4 — Composição didática do Valor Total Atualizado")
    cabecalho = ["Ref.", "Parcela", "Valor"]
    linhas: list[list[str]] = []
    letras_componentes: list[str] = []
    for i, (desc, valor) in enumerate(componentes):
        ref = _LETRAS[i] if i < len(_LETRAS) else str(i + 1)
        letras_componentes.append(ref)
        linhas.append([ref, desc, formatar_moeda(valor) if valor is not None else ""])
    vta = dados.get("vta")
    ref_vta = _LETRAS[len(componentes)] if len(componentes) < len(_LETRAS) else "VTA"
    linhas.append([ref_vta, "Valor Total Atualizado do Contrato",
                   formatar_moeda(vta) if vta is not None else ""])
    linhas.append(["Total", "Valor Total Atualizado",
                   formatar_moeda(vta) if vta is not None else ""])
    _adicionar_tabela(doc, cabecalho, linhas)

    # 4.2 — leitura correta: componentes = E = VTA (nunca soma o VTA a si mesmo).
    p42 = doc.add_paragraph()
    p42.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p42, "4.2. A composição do Quadro 4 pode ser lida da seguinte forma: ")
    if letras_componentes:
        soma = " + ".join(letras_componentes)
        _adicionar_run(p42, f"{soma} = {ref_vta} = ")
    else:
        _adicionar_run(p42, f"{ref_vta} = ")
    if vta is not None:
        _adicionar_run(p42, formatar_moeda(vta), negrito=True)
    else:
        _run_campo_manual(p42, "Valor Total Atualizado")
    _adicionar_run(p42, ".")
    doc.add_paragraph()


def _ta_secao5_valores_unitarios(doc: Document, dados: dict) -> None:
    _titulo_secao(doc, "5. Dos valores unitários")
    _secao_valores_unitarios_por_ciclo(doc, dados)


def _ta_secao6_aditivos(doc: Document, dados: dict) -> None:
    _titulo_secao(doc, "6. Dos aditivos e supressões considerados")
    aditivos = dados.get("aditivos") or []
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if not aditivos:
        _adicionar_run(p1,
            "6.1. Não foram identificados aditivos ou supressões específicos na "
            "base processada, sem prejuízo da conferência dos instrumentos já "
            "formalizados no processo.")
    else:
        _adicionar_run(p1, "6.1. Foram considerados os seguintes aditivos e supressões: ")
        partes = []
        for ad in aditivos:
            ident = remover_emojis_leve(ad.get("identificador") or ad.get("ciclo") or "instrumento")
            val = _num_ou_none(ad.get("valor_atualizado"))
            if val is not None:
                partes.append(f"{ident}, com impacto atualizado de {formatar_moeda(val)}")
            else:
                partes.append(f"{ident}, com impacto a confirmar")
        _adicionar_run(p1, "; ".join(partes) + ".")

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p2,
        "6.2. Os aditivos e supressões computáveis integram o Valor Total "
        "Atualizado quando não estiverem refletidos na execução atualizada, no "
        "saldo remanescente ou no valor formalizado anterior, vedada a dupla "
        "contagem.")
    doc.add_paragraph()


def _ta_secoes_finais(doc: Document, cm: dict) -> None:
    p7 = doc.add_paragraph()
    p7.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p7, "7. ", negrito=True)
    _adicionar_run(p7,
        "Permanecem inalteradas e em pleno vigor as demais cláusulas e condições "
        "do Contrato e de seus instrumentos posteriores não modificadas por este "
        "Termo de Apostila.")

    p8 = doc.add_paragraph()
    p8.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p8, "8. ", negrito=True)
    _adicionar_run(p8,
        "A CONTRATADA deverá atualizar a garantia contratual, prevista na "
        "cláusula própria do Contrato, no prazo contratualmente estabelecido, "
        "observado o novo valor após a formalização deste Termo de Apostila.")

    p9 = doc.add_paragraph()
    p9.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p9, "9. ", negrito=True)
    _adicionar_run(p9,
        "O presente apostilamento vincula-se, para todos os fins, aos documentos "
        "instruídos no Processo ")
    _texto_ou_marcador(p9, _campo(cm, "processo_ref"), "Numero do processo de instrucao")
    _adicionar_run(p9, ".")
    doc.add_paragraph()


def _ta_assinaturas(doc: Document, cm: dict) -> None:
    p_local = doc.add_paragraph()
    p_local.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _adicionar_run(p_local, "Brasília/DF, ")
    _texto_ou_marcador(p_local, _campo(cm, "local_data"), "Data")
    _adicionar_run(p_local, ".")
    doc.add_paragraph()
    doc.add_paragraph()

    # Dois representantes da TELEBRAS (nenhuma assinatura da CONTRATADA).
    for chave_nome, desc_nome in (
        ("representante_telebras_1_nome", "Nome do 1o representante da Telebras"),
        ("representante_telebras_2_cargo", "Cargo do 2o representante da Telebras"),
    ):
        p_ent = doc.add_paragraph()
        p_ent.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _adicionar_run(p_ent, "TELECOMUNICAÇÕES BRASILEIRAS S.A. - TELEBRAS")
        p_rep = doc.add_paragraph()
        p_rep.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _texto_ou_marcador(p_rep, _campo(cm, chave_nome), desc_nome)
        doc.add_paragraph()


# ---------------------------------------------------------------------------
# DESPACHO SANEADOR (modelo canonico §7)
# ---------------------------------------------------------------------------

def gerar_despacho_saneador(
    leitura_ou_objeto: dict,
    identificacao: dict | None = None,
    campos_manuais: dict | None = None,
) -> bytes:
    """Gera o Despacho Saneador em DOCX e retorna os bytes."""
    if campos_manuais is None:
        campos_manuais = {}
    dados = _extrair_dados(leitura_ou_objeto, identificacao)
    doc = _configurar_documento()

    _ds_assunto(doc, campos_manuais)
    _ds_par1(doc)
    _ds_par2(doc, dados, campos_manuais)
    _ds_par3(doc, dados, campos_manuais)
    _ds_par4_quadro1(doc, dados, campos_manuais)
    _ds_par5_quadro2(doc, dados)
    _ds_par6_quadro3(doc, dados, campos_manuais)
    _ds_par7_composicao(doc, dados)
    _ds_par8_aditivos(doc, dados)
    _ds_par9_adequacao(doc, campos_manuais)
    _ds_par10_regularidade(doc, campos_manuais)
    _ds_par11_concordancia(doc, campos_manuais)
    _ds_par12_garantia(doc)
    _ds_par13_docs(doc, campos_manuais)
    _ds_conclusao(doc, dados, campos_manuais)
    _ds_quadro4(doc, dados, campos_manuais)

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _ds_assunto(doc: Document, cm: dict) -> None:
    _titulo_secao(doc, "DESPACHO SANEADOR", tamanho=12,
                  alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _adicionar_run(p, "Assunto: ", negrito=True)
    _adicionar_run(p, "Saneamento para formalização de Termo de Apostila de Reajuste - ")
    _texto_ou_marcador(p, _campo(cm, "contrato"), "Numero do contrato")
    _adicionar_run(p, ".")
    doc.add_paragraph()


def _ds_par(doc: Document, numero: str) -> Any:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, f"{numero}. ", negrito=True)
    return p


def _ds_par1(doc: Document) -> None:
    p = _ds_par(doc, "1")
    _adicionar_run(p,
        "Este despacho saneador consolida os elementos documentais, financeiros "
        "e formais necessários à instrução do Termo de Apostila destinado ao "
        "registro de reajuste contratual, com a finalidade de demonstrar a "
        "regularidade mínima da instrução antes da formalização.")


def _ds_par2(doc: Document, dados: dict, cm: dict) -> None:
    p = _ds_par(doc, "2")
    _adicionar_run(p, "A contratada apresentou pleito de reajuste por meio de ")
    _texto_ou_marcador(p, _campo(cm, "processo_pleito"), "Referencias do pleito da contratada")
    _adicionar_run(p,
        ". Para fins de verificação da anualidade, foram consideradas a data da "
        "proposta em ")
    _texto_ou_marcador(p, _campo(cm, "data_proposta"), "Data da proposta")
    _adicionar_run(p, ", o índice contratual ")
    indice = _indice_doc(dados)
    if indice:
        _adicionar_run(p, indice, negrito=True)
    else:
        _run_campo_manual(p, "Indice contratual")
    _adicionar_run(p, " e as datas de pedido registradas na análise: ")
    datas = []
    for c in dados.get("ciclos_reajuste") or []:
        dp = str(c.get("data_pedido") or "").strip()
        if dp and dp != NAO_INFORMADO:
            datas.append(f"{remover_emojis_leve(c.get('ciclo') or '')} em {dp}")
    if datas:
        _adicionar_run(p, "; ".join(datas) + ".")
    else:
        _run_campo_manual(p, "Datas de pedido por ciclo")
        _adicionar_run(p, ".")


def _ds_par3(doc: Document, dados: dict, cm: dict) -> None:
    p = _ds_par(doc, "3")
    ciclos = dados.get("ciclos_computados") or []
    if ciclos:
        nomes = ", ".join(remover_emojis_leve(c.get("ciclo") or "") for c in ciclos)
        _adicionar_run(p, f"Acordou-se na concessão de {nomes}, conforme exposto em ")
    else:
        _adicionar_run(p, "Acordou-se na concessão dos ciclos apurados, conforme exposto em ")
    _texto_ou_marcador(p, _campo(cm, "referencia_analise"), "Referencia onde o resultado da analise consta")
    _adicionar_run(p, ".")


def _ds_par4_quadro1(doc: Document, dados: dict, cm: dict) -> None:
    n = len(dados.get("ciclos_reajuste") or [])
    p = _ds_par(doc, "4")
    _adicionar_run(p, f"A análise de reajuste considerou {n} ciclo(s), com variação acumulada de ")
    var = dados.get("var_acumulada")
    if var is not None:
        _adicionar_run(p, _fmt_pct_doc(var), negrito=True)
    else:
        _run_campo_manual(p, "Variacao acumulada")
    _adicionar_run(p, ". O valor original do contrato informado foi de ")
    _valor_moeda_ou_marcador(p, _campo(cm, "valor_original_contrato"), "Valor original do contrato")
    _adicionar_run(p, ".")

    _titulo_quadro(doc, "Quadro 1 - Síntese dos ciclos de reajuste")
    cabecalho = ["Ciclo", "Data-base", "Data do pedido", "Início financeiro",
                 "Fim financeiro", "Situação", "Percentual aplicado"]
    linhas: list[list[str]] = []
    for c in dados.get("ciclos_reajuste") or []:
        pct = c.get("percentual_reajuste")
        linhas.append([
            remover_emojis_leve(c.get("ciclo") or ""),
            c.get("data_inicio") or NAO_INFORMADO,
            c.get("data_pedido") or NAO_INFORMADO,
            c.get("inicio_efeito_financeiro") or NAO_INFORMADO,
            c.get("data_fim") or NAO_INFORMADO,
            remover_emojis_leve(c.get("situacao") or NAO_INFORMADO),
            _fmt_pct_doc(pct) if pct is not None else NAO_INFORMADO,
        ])
    if not linhas:
        linhas = [["—"] * 7]
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ds_par5_quadro2(doc: Document, dados: dict) -> None:
    p = _ds_par(doc, "5")
    _adicionar_run(p, "A apuração financeira consolidada indicou valor pago efetivo de ")
    vp = _valor_pago_total(dados)
    if vp is not None:
        _adicionar_run(p, formatar_moeda(vp), negrito=True)
    else:
        _run_campo_manual(p, "Valor pago efetivo")
    _adicionar_run(p, " e valor teórico calculado de ")
    vat = _valor_atualizado_total(dados)
    if vat is not None:
        _adicionar_run(p, formatar_moeda(vat), negrito=True)
    else:
        _run_campo_manual(p, "Valor teorico calculado")
    _adicionar_run(p, ", resultando em valor retroativo a pagar de ")
    retro = _retroativo_total(dados)
    if retro is not None:
        _adicionar_run(p, formatar_moeda(retro), negrito=True)
    else:
        _run_campo_manual(p, "Valor retroativo a pagar")
    _adicionar_run(p, ".")

    _titulo_quadro(doc, "Quadro 2 - Apuração financeira por ciclo")
    cabecalho = ["Ciclo", "Valor pago efetivo", "Valor teórico calculado", "Diferença/retroativo"]
    linhas: list[list[str]] = []
    tot_pago = tot_teorico = tot_delta = None
    for lin in _linhas_financeiro(dados):
        vpg = _num_ou_none(lin.get("valor_pago"))
        vtc = _num_ou_none(lin.get("valor_atualizado"))
        vdl = _num_ou_none(lin.get("delta"))
        linhas.append([
            remover_emojis_leve(lin.get("ciclo") or ""),
            formatar_moeda(vpg) if vpg is not None else "",
            formatar_moeda(vtc) if vtc is not None else "",
            formatar_moeda(vdl) if vdl is not None else "",
        ])
        if vpg is not None:
            tot_pago = (tot_pago or 0.0) + vpg
        if vtc is not None:
            tot_teorico = (tot_teorico or 0.0) + vtc
        if vdl is not None:
            tot_delta = (tot_delta or 0.0) + vdl
    if not linhas:
        linhas = [["—", "", "", ""]]
    linhas.append([
        "Total",
        formatar_moeda(tot_pago) if tot_pago is not None else "",
        formatar_moeda(tot_teorico) if tot_teorico is not None else "",
        formatar_moeda(tot_delta) if tot_delta is not None else "",
    ])
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ds_par6_quadro3(doc: Document, dados: dict, cm: dict) -> None:
    p = _ds_par(doc, "6")
    _adicionar_run(p,
        "Para fins de consolidação contratual, foi adotada a premissa de "
        "considerar, para fins de cálculo do retroativo e consequente cálculo do "
        "valor remanescente do contrato, ")
    _texto_ou_marcador(p, _campo(cm, "data_corte_descricao"), "Descricao da data/posicao de corte adotada")
    _adicionar_run(p, ".")

    _titulo_quadro(doc, "Quadro 3 - Memória fiscal do Valor Total Atualizado Estimado")
    cabecalho = ["Descrição", "Valor"]
    linhas: list[list[str]] = []
    for desc, valor in _composicao_didatica_vta(dados):
        linhas.append([desc, formatar_moeda(valor) if valor is not None else ""])
    vta = dados.get("vta")
    linhas.append(["Valor total do contrato estimado",
                   formatar_moeda(vta) if vta is not None else ""])
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ds_par7_composicao(doc: Document, dados: dict) -> None:
    p = _ds_par(doc, "7")
    _adicionar_run(p,
        "De forma didática, o Valor Total Atualizado Estimado do Contrato pode "
        "ser lido pela seguinte composição:")
    componentes = _composicao_didatica_vta(dados)
    cabecalho = ["Parcela", "Valor"]
    linhas = [[desc, formatar_moeda(valor) if valor is not None else ""]
              for desc, valor in componentes]
    vta = dados.get("vta")
    linhas.append(["Valor Total Atualizado Estimado do Contrato",
                   formatar_moeda(vta) if vta is not None else ""])
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ds_par8_aditivos(doc: Document, dados: dict) -> None:
    p = _ds_par(doc, "8")
    aditivos = dados.get("aditivos") or []
    if not aditivos:
        _adicionar_run(p,
            "Quanto aos aditivos e supressões, não foram identificados eventos "
            "específicos na base processada, sem prejuízo da conferência dos "
            "instrumentos já formalizados no processo.")
        return
    _adicionar_run(p, "Quanto aos aditivos e supressões, registra-se: ")
    partes = []
    for ad in aditivos:
        ident = remover_emojis_leve(ad.get("identificador") or ad.get("ciclo") or "instrumento")
        val = _num_ou_none(ad.get("valor_atualizado"))
        if val is not None:
            partes.append(f"{ident}, com impacto atualizado de {formatar_moeda(val)}")
        else:
            partes.append(f"{ident}, com impacto a confirmar")
    _adicionar_run(p, "; ".join(partes) + ".")


def _ds_par9_adequacao(doc: Document, cm: dict) -> None:
    p = _ds_par(doc, "9")
    _adicionar_run(p,
        "Foi realizada a adequação orçamentária necessária ao prosseguimento da "
        "instrução, no valor de ")
    _valor_moeda_ou_marcador(p, _campo(cm, "adequacao_orcamentaria_valor"), "Valor da adequacao orcamentaria")
    _adicionar_run(p, ", conforme documento ")
    _texto_ou_marcador(p, _campo(cm, "adequacao_orcamentaria_ref"), "Referencia da adequacao orcamentaria")
    _adicionar_run(p, ".")


def _ds_par10_regularidade(doc: Document, cm: dict) -> None:
    p = _ds_par(doc, "10")
    _adicionar_run(p, "As certidões de regularidade estão presentes em ")
    _texto_ou_marcador(p, _campo(cm, "regularidade_ref"), "Referencia das certidoes de regularidade")
    _adicionar_run(p, ".")


def _ds_par11_concordancia(doc: Document, cm: dict) -> None:
    p = _ds_par(doc, "11")
    _adicionar_run(p, "A contratada manifestou concordância com os valores propostos conforme registrado em ")
    _texto_ou_marcador(p, _campo(cm, "concordancia_ref"), "Referencia da manifestacao de concordancia da contratada")
    _adicionar_run(p, ".")


def _ds_par12_garantia(doc: Document) -> None:
    p = _ds_par(doc, "12")
    _adicionar_run(p,
        "A contratada foi informada da necessidade de apresentação do endosso da "
        "garantia contratual, quando aplicável, observando-se o prazo e as "
        "condições previstos no contrato.")


def _ds_par13_docs(doc: Document, cm: dict) -> None:
    docs = _campo(cm, "docs_desatualizados")
    if not docs:
        return
    p = _ds_par(doc, "13")
    _adicionar_run(p,
        "Após atualizações e alinhamentos internos, alguns documentos instruídos "
        "mostram-se desatualizados, devendo ser desconsiderados: ")
    if isinstance(docs, (list, tuple)):
        _adicionar_run(p, ", ".join(str(d) for d in docs))
    else:
        _adicionar_run(p, str(docs))
    _adicionar_run(p, ".")


def _tem_pendencia_critica(dados: dict, cm: dict) -> bool:
    """Soft-block: nao afirmar 'inexiste pendencia critica' se houver pendencia."""
    if not dados.get("disponivel"):
        return True
    flag = cm.get("pendencia_critica")
    if isinstance(flag, str):
        return flag.strip().lower() in ("sim", "true", "1", "critica", "critico")
    return bool(flag)


def _ds_conclusao(doc: Document, dados: dict, cm: dict) -> None:
    # Numeracao final: o item de documentos desatualizados (13) so existe quando
    # ha docs_desatualizados. A conclusao vem logo apos — 14 nesse caso, 13 caso
    # contrario. A logica de soft-block do texto permanece inalterada.
    numero = "14" if _campo(cm, "docs_desatualizados") else "13"
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, f"{numero}. ", negrito=True)
    if _tem_pendencia_critica(dados, cm):
        _adicionar_run(p,
            "Diante do exposto, os elementos disponíveis encontram-se "
            "consolidados para análise, permanecendo pendentes as complementações "
            "ou validações indicadas antes da formalização do Termo de Apostila.")
    else:
        _adicionar_run(p,
            "Diante do exposto, estando conferidos os elementos documentais, "
            "financeiros e formais acima indicados, e inexistindo pendência "
            "crítica impeditiva, a instrução poderá prosseguir para formalização "
            "do Termo de Apostila, observadas as alçadas competentes e os "
            "procedimentos internos aplicáveis.")
    doc.add_paragraph()


def _ds_quadro4(doc: Document, dados: dict, cm: dict) -> None:
    _titulo_quadro(doc, "Quadro 4 - Síntese dos principais valores")
    cabecalho = ["Parcela", "Valor"]
    linhas: list[list[str]] = []

    val_orig = _num_ou_none(_campo(cm, "valor_original_contrato"))
    linhas.append(["Valor original do contrato",
                   formatar_moeda(val_orig) if val_orig is not None
                   else PREENCHER_TAG.format("Valor original do contrato")])

    var = dados.get("var_acumulada")
    if var is not None:
        linhas.append(["Variação acumulada do reajuste", _fmt_pct_doc(var)])

    for lin in _linhas_financeiro(dados):
        delta = _num_ou_none(lin.get("delta"))
        if delta is not None:
            ciclo = remover_emojis_leve(lin.get("ciclo") or "")
            linhas.append([f"Retroativo {ciclo}".strip(), formatar_moeda(delta)])

    retro = _retroativo_total(dados)
    if retro is not None:
        linhas.append(["Valor retroativo/represado a pagar", formatar_moeda(retro)])

    vta = dados.get("vta")
    if vta is not None:
        linhas.append(["Valor Total Atualizado Estimado do Contrato", formatar_moeda(vta)])

    adeq = _num_ou_none(_campo(cm, "adequacao_orcamentaria_valor"))
    if adeq is not None:
        linhas.append(["Adequação orçamentária registrada", formatar_moeda(adeq)])

    _adicionar_tabela(doc, cabecalho, linhas)


# ---------------------------------------------------------------------------
# Diagnostico de campos manuais
# ---------------------------------------------------------------------------

def diagnosticar_campos_manuais(
    leitura_ou_objeto: dict,
    identificacao: dict | None = None,
    campos_manuais: dict | None = None,
) -> list[dict]:
    """Retorna lista de campos manuais pendentes: {campo, descricao, documento}."""
    if campos_manuais is None:
        campos_manuais = {}
    pendentes = []
    vistos: set[str] = set()
    for chave, descricao, documento in TODOS_CAMPOS_MANUAIS:
        if chave in _CAMPOS_OPCIONAIS:
            continue
        if _campo(campos_manuais, chave) is None and chave not in vistos:
            vistos.add(chave)
            pendentes.append({"campo": chave, "descricao": descricao, "documento": documento})
    return pendentes
