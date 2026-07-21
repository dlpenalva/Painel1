"""Geradores de documentos juridicos em DOCX (Etapa 6).

Gera o Despacho Saneador e a Minuta de Termo de Apostilamento a partir dos
dados canonicos do Objeto Processo de Reajuste. Nao recalcula valores — apenas
apresenta os dados ja consolidados pelos motores oficiais.

Campos manuais ausentes recebem o marcador [PREENCHER: <descricao>] com
destaque amarelo. Ausencia de dado automatico nunca vira zero.

Interface publica:
    gerar_despacho_saneador(leitura_ou_objeto, identificacao, campos_manuais) -> bytes
    gerar_termo_apostila(leitura_ou_objeto, identificacao, campos_manuais) -> bytes
    diagnosticar_campos_manuais(leitura_ou_objeto, identificacao, campos_manuais) -> list[dict]
"""
from __future__ import annotations

from datetime import date
from io import BytesIO
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from docx.enum.style import WD_STYLE_TYPE

from _sumario_executivo import (
    NAO_INFORMADO,
    formatar_moeda,
    montar_dados_sumario_executivo,
    _fmt_data,
    _num_ou_none,
    _texto_ou_nao_informado,
)
from _objeto_processo_reajuste import obter_objeto_processo_reajuste

# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

PREENCHER_TAG = "[PREENCHER: {}]"
COR_NEGATIVO = RGBColor(0xC0, 0x00, 0x00)


def _fmt_pct_doc(valor: Any) -> str:
    """Formata percentual documental com exatamente duas casas decimais.

    Entrada no formato decimal canonico (0.0421 -> '4,21%').
    Nunca altera valores internos, fatores ou calculos.
    """
    numero = _num_ou_none(valor)
    if numero is None:
        return NAO_INFORMADO
    texto = f"{numero * 100:.2f}".replace(".", ",")
    return f"{texto}%"

CAMPOS_MANUAIS_DESPACHO = [
    ("contrato", "Numero do contrato (ex: TLB-CTR-2022/00067)", "despacho"),
    ("processo_pleito", "Processo(s) do pleito da contratada", "despacho"),
    ("processo_analise", "Processo onde resultado da analise foi exarado", "despacho"),
    ("adequacao_orcamentaria_ref", "Referencia do documento de adequacao orcamentaria", "despacho"),
    ("adequacao_orcamentaria_valor", "Valor da adequacao orcamentaria (float)", "despacho"),
    ("regularidade_ref", "Referencia das certidoes de regularidade", "despacho"),
    ("concordancia_ref", "Referencia da manifestacao de concordancia da contratada", "despacho"),
    ("docs_desatualizados", "Lista de referencias de documentos a desconsiderar", "despacho"),
    ("data_corte_descricao", "Ex: abril de 2026, ultimo mes com liquidacao", "despacho"),
    ("valor_original_contrato", "Valor original do contrato (float)", "despacho"),
]

CAMPOS_MANUAIS_TERMO = [
    ("contrato", "Numero do contrato (ex: TLB-CTR-2022/00067)", "termo"),
    ("ata_diretoria", "Referencia da Ata da reuniao da Diretoria Executiva", "termo"),
    ("clausula_reajuste", "Clausula do contrato que preve o reajuste", "termo"),
    ("clausula_garantia", "Clausula do contrato sobre garantia contratual", "termo"),
    ("despacho_saneador_ref", "Referencia do Despacho Saneador", "termo"),
    ("memoria_calculo_ref", "Referencia da memoria de calculo", "termo"),
    ("memoria_financeira_ref", "Referencia da memoria financeira", "termo"),
    ("informacoes_gestora_ref", "Referencia das informacoes da area gestora do contrato", "termo"),
    ("concordancia_ref", "Referencia da manifestacao de concordancia da contratada", "termo"),
    ("regularidade_ref", "Referencia das certidoes de regularidade", "termo"),
    ("adequacao_orcamentaria_ref", "Referencia do documento de adequacao orcamentaria", "termo"),
    ("data_corte_descricao", "Ex: abril de 2026, ultimo mes com liquidacao", "termo"),
    ("vinculacao_docs", "Documentos de vinculacao (para paragrafo 8 do Termo)", "termo"),
    ("processo_vinculacao", "Processo de vinculacao (para paragrafo 8 do Termo)", "termo"),
    ("local_data", "Ex: Brasilia/DF, dd/mm/aaaa.", "termo"),
    ("representante_telebras_nome", "Nome do representante da Telebras", "termo"),
    ("representante_telebras_titulo", "Titulo do representante (ex: Presidente)", "termo"),
    ("representante_contratada_nome", "Nome do representante da contratada", "termo"),
    ("representante_contratada_qualificacao", "Qualificacao do representante da contratada", "termo"),
    ("garantia_clausula_obs", "Observacao especifica sobre garantia (opcional)", "termo"),
]

TODOS_CAMPOS_MANUAIS = list({c[0]: c for c in CAMPOS_MANUAIS_DESPACHO + CAMPOS_MANUAIS_TERMO}.values())


# ---------------------------------------------------------------------------
# Helpers XML / DOCX
# ---------------------------------------------------------------------------

def _set_highlight(run, cor: str = "yellow") -> None:
    """Aplica destaque colorido a um run via XML direto."""
    rPr = run._r.get_or_add_rPr()
    highlight = OxmlElement("w:highlight")
    highlight.set(qn("w:val"), cor)
    rPr.append(highlight)


def _repetir_cabecalho(tabela) -> None:
    """Marca primeira linha como cabecalho repetido em paginacao."""
    tr = tabela.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    trPr.append(tblHeader)


def _nova_secao(doc: Document, texto: str) -> None:
    p = doc.add_paragraph(texto)
    p.style = doc.styles["Heading 2"]
    return p


def _adicionar_paragrafo(doc: Document, texto: str | None = None, negrito: bool = False,
                          tamanho: int = 11, alinhamento=WD_ALIGN_PARAGRAPH.JUSTIFY) -> Any:
    p = doc.add_paragraph()
    p.alignment = alinhamento
    if texto:
        run = p.add_run(texto)
        run.bold = negrito
        run.font.name = "Calibri"
        run.font.size = Pt(tamanho)
    return p


def _adicionar_run(p, texto: str, negrito: bool = False, tamanho: int = 11,
                   cor: RGBColor | None = None, italico: bool = False) -> Any:
    run = p.add_run(texto)
    run.bold = negrito
    run.italic = italico
    run.font.name = "Calibri"
    run.font.size = Pt(tamanho)
    if cor:
        run.font.color.rgb = cor
    return run


def _run_campo_manual(p, descricao: str, tamanho: int = 11) -> Any:
    """Adiciona um run com o marcador de campo manual em amarelo."""
    run = p.add_run(PREENCHER_TAG.format(descricao))
    run.font.name = "Calibri"
    run.font.size = Pt(tamanho)
    _set_highlight(run, "yellow")
    return run


def _texto_ou_marcador(p, valor: Any, descricao: str, tamanho: int = 11,
                        negrito: bool = False, cor: RGBColor | None = None) -> None:
    """Adiciona valor formatado ou marcador amarelo se ausente."""
    if valor is not None and str(valor).strip():
        run = p.add_run(str(valor))
        run.bold = negrito
        run.font.name = "Calibri"
        run.font.size = Pt(tamanho)
        if cor:
            run.font.color.rgb = cor
    else:
        _run_campo_manual(p, descricao, tamanho)


def _campo(campos_manuais: dict, chave: str) -> Any:
    """Retorna o valor do campo manual ou None."""
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
        texto = formatar_moeda(numero)
        negativo = numero < 0
        run = p.add_run(texto)
        run.font.name = "Calibri"
        run.font.size = Pt(tamanho)
        if negativo:
            run.font.color.rgb = COR_NEGATIVO
    else:
        _run_campo_manual(p, descricao, tamanho)


def _configurar_documento() -> Document:
    """Cria documento com margens e fonte padrao."""
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)
    # Estilo Normal
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)
    return doc


def _adicionar_tabela(doc: Document, cabecalho: list[str],
                       linhas: list[list[str]], larguras: list[float] | None = None) -> Any:
    """Adiciona tabela com estilo Table Grid, cabecalho em negrito."""
    n_cols = len(cabecalho)
    tabela = doc.add_table(rows=1, cols=n_cols)
    tabela.style = "Table Grid"
    # Cabecalho
    celulas_cab = tabela.rows[0].cells
    for i, texto in enumerate(cabecalho):
        celulas_cab[i].text = ""
        p = celulas_cab[i].paragraphs[0]
        run = p.add_run(texto)
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(10)
    _repetir_cabecalho(tabela)
    # Linhas
    for linha in linhas:
        row = tabela.add_row()
        for i, celula_texto in enumerate(linha):
            row.cells[i].text = ""
            p = row.cells[i].paragraphs[0]
            negativo = False
            try:
                val_num = float(str(celula_texto).replace("R$ ", "").replace(".", "").replace(",", "."))
                negativo = val_num < 0
            except (ValueError, AttributeError):
                pass
            run = p.add_run(str(celula_texto))
            run.font.name = "Calibri"
            run.font.size = Pt(10)
            if negativo and "R$" in str(celula_texto):
                run.font.color.rgb = COR_NEGATIVO
    return tabela


# ---------------------------------------------------------------------------
# Extracao de dados canonicos
# ---------------------------------------------------------------------------

def _extrair_dados(leitura_ou_objeto: dict, identificacao: dict | None) -> dict:
    """Extrai e normaliza todos os dados canonicos necessarios para os documentos."""
    dados = montar_dados_sumario_executivo(leitura_ou_objeto, identificacao)
    if not dados.get("disponivel"):
        return {"disponivel": False}

    ciclos = dados.get("ciclos") or []
    ciclos_reajuste = [c for c in ciclos if not c.get("eh_base")]
    ciclos_computados = [c for c in ciclos_reajuste if c.get("computar") == "Sim"]

    financeiro = dados.get("financeiro") or {}
    sintese = dados.get("sintese") or {}
    aditivos_raw = (dados.get("aditivos") or {}).get("itens") or []

    # Variacao acumulada
    var_acumulada = sintese.get("variacao_acumulada")

    # VTA
    vta = sintese.get("vta")

    # Retroativos por ciclo
    fin_por_ciclo = {r["ciclo"]: r for r in financeiro.get("financeiro_por_ciclo") or []}
    pc_por_ciclo = {r["ciclo"]: r for r in financeiro.get("pc_por_ciclo") or []}

    # Composicao VTA (parcelas) — navega pelo objeto_processo, nao pela raiz
    objeto_proc = obter_objeto_processo_reajuste(leitura_ou_objeto) or {}
    dados_op = objeto_proc.get("dados_operacionais") or {}
    vta_sombra = dados_op.get("vta_sombra") or {}
    parcelas_vta = vta_sombra.get("parcelas_computadas") or []

    # Aditivos
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
        "var_acumulada": var_acumulada,
        "vta": vta,
        "fin_por_ciclo": fin_por_ciclo,
        "pc_por_ciclo": pc_por_ciclo,
        "parcelas_vta": parcelas_vta,
        "aditivos": aditivos,
        "financeiro": financeiro,
        "sintese": sintese,
        # Etapa 7: estrutura canonica unica de VU por ciclo (C0..ultimo analisado).
        "historico_vu": dados.get("historico_vu") or {},
    }


def _tem_retroativo(dados: dict) -> bool:
    return bool(dados.get("fin_por_ciclo") or dados.get("pc_por_ciclo"))


def _retroativo_total(dados: dict) -> float | None:
    fin = dados.get("financeiro") or {}
    t_fin = fin.get("delta_total_financeiro")
    t_pc = fin.get("delta_total_pc")
    if t_fin is not None:
        return t_fin
    if t_pc is not None:
        return t_pc
    return None


def _linhas_financeiro(dados: dict) -> list[dict]:
    """Retorna linhas financeiras priorizando Financeiro; fallback PC."""
    if dados.get("fin_por_ciclo"):
        return list(dados["fin_por_ciclo"].values())
    if dados.get("pc_por_ciclo"):
        return list(dados["pc_por_ciclo"].values())
    return []


def _valor_pago_total(dados: dict) -> float | None:
    linhas = _linhas_financeiro(dados)
    if not linhas:
        return None
    total = sum(_num_ou_none(l.get("valor_pago")) or 0.0 for l in linhas)
    return round(total, 2) if linhas else None


def _valor_atualizado_total(dados: dict) -> float | None:
    linhas = _linhas_financeiro(dados)
    if not linhas:
        return None
    total = sum(_num_ou_none(l.get("valor_atualizado")) or 0.0 for l in linhas)
    return round(total, 2) if linhas else None


# ---------------------------------------------------------------------------
# Construcao do Despacho Saneador
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

    _ds_titulo(doc, campos_manuais)
    _ds_assunto(doc, campos_manuais)
    _ds_par1_introducao(doc)
    _ds_par2_pleito(doc, dados, campos_manuais)
    _ds_par3_acordo(doc, dados, campos_manuais)
    _ds_par4_analise(doc, dados, campos_manuais)
    _ds_quadro1_ciclos(doc, dados)
    _ds_par5_apuracao_financeira(doc, dados)
    _ds_quadro2_financeiro(doc, dados)
    _ds_par6_data_corte(doc, campos_manuais)
    _ds_quadro3_vta(doc, dados, campos_manuais)
    _ds_par7_composicao_vta(doc, dados)
    _secao_valores_unitarios_por_ciclo(doc, dados)
    _ds_par8_aditivos(doc, dados)
    _ds_par9_adequacao(doc, campos_manuais)
    _ds_par10_regularidade(doc, campos_manuais)
    _ds_par11_concordancia(doc, campos_manuais)
    _ds_par12_garantia(doc)
    _ds_par13_docs_desatualizados(doc, campos_manuais)
    _ds_par14_conclusao(doc, campos_manuais)
    _ds_quadro4_sintese(doc, dados, campos_manuais)

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _ds_titulo(doc: Document, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("DESPACHO SANEADOR PARA FORMALIZAÇÃO DE TERMO DE APOSTILA")
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(12)


def _ds_assunto(doc: Document, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _adicionar_run(p, "Assunto: ", negrito=True)
    _adicionar_run(p, "Saneamento da instrução para formalização de Termo de Apostila - Contrato ")
    _texto_ou_marcador(p, _campo(cm, "contrato"), "Numero do contrato")
    _adicionar_run(p, ".")
    doc.add_paragraph()


def _ds_par1_introducao(doc: Document) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "1. ", negrito=True)
    _adicionar_run(p,
        "Este despacho saneador consolida os elementos documentais, financeiros e "
        "formais necessários à instrução do Termo de Apostila destinado ao registro "
        "de reajuste contratual.")


def _ds_par2_pleito(doc: Document, dados: dict, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "2. ", negrito=True)
    _adicionar_run(p, "A contratada apresentou pleito de reajuste por meio de ")
    _texto_ou_marcador(p, _campo(cm, "processo_pleito"), "Referencia(s) do processo de pleito da contratada")
    _adicionar_run(p, ". Para fins de verificação da anualidade, foram consideradas as datas previstas em contrato.")
    # Verificar se algum ciclo foi admitido por negociacao
    ciclos_neg = [c for c in (dados.get("ciclos_reajuste") or [])
                  if "negociacao" in str(c.get("situacao") or "").lower()
                  or "admitido" in str(c.get("situacao") or "").lower()]
    if ciclos_neg:
        nomes = ", ".join(c["ciclo"] for c in ciclos_neg)
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _adicionar_run(p2,
            f"Registra-se que o(s) ciclo(s) {nomes} foi(ram) admitido(s) por "
            "negociação, conforme acordado entre as partes.")


def _ds_par3_acordo(doc: Document, dados: dict, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "3. ", negrito=True)
    ciclos_comp = dados.get("ciclos_computados") or []
    if ciclos_comp:
        nomes = ", ".join(c["ciclo"] for c in ciclos_comp)
        _adicionar_run(p, f"Acordou-se na concessão do(s) ciclo(s) {nomes}, conforme exposto em ")
    else:
        _adicionar_run(p, "Acordou-se na concessão dos ciclos apurados, conforme exposto em ")
    _texto_ou_marcador(p, _campo(cm, "processo_analise"), "Processo onde resultado da analise foi exarado")
    _adicionar_run(p, ".")


def _ds_par4_analise(doc: Document, dados: dict, cm: dict) -> None:
    n_ciclos = len(dados.get("ciclos_reajuste") or [])
    var = dados.get("var_acumulada")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "4. ", negrito=True)
    _adicionar_run(p, f"A análise de reajuste considerou {n_ciclos} ciclo(s), com variação acumulada de ")
    if var is not None:
        _adicionar_run(p, _fmt_pct_doc(var), negrito=True)
    else:
        _run_campo_manual(p, "Variacao percentual acumulada")
    _adicionar_run(p, ". O valor original do contrato informado foi de ")
    _valor_moeda_ou_marcador(p, _campo(cm, "valor_original_contrato"), "Valor original do contrato")
    _adicionar_run(p, ".")


def _ds_quadro1_ciclos(doc: Document, dados: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Quadro 1 – Síntese dos ciclos de reajuste")
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(10)

    cabecalho = ["Ciclo", "Data-base", "Data do pedido", "Início financeiro",
                 "Fim financeiro", "Situação", "Percentual aplicado"]
    linhas = []
    for c in dados.get("ciclos_reajuste") or []:
        pct = c.get("percentual_reajuste")
        pct_str = _fmt_pct_doc(pct) if pct is not None else NAO_INFORMADO
        linhas.append([
            c.get("ciclo", ""),
            c.get("data_inicio", NAO_INFORMADO),
            c.get("data_pedido", NAO_INFORMADO),
            c.get("data_inicio", NAO_INFORMADO),  # inicio_efeito nao disponivel diretamente aqui
            c.get("data_fim", NAO_INFORMADO),
            c.get("situacao", NAO_INFORMADO),
            pct_str,
        ])
    if not linhas:
        linhas = [["—"] * 7]
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ds_par5_apuracao_financeira(doc: Document, dados: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "5. ", negrito=True)  # renumerado internamente
    _adicionar_run(p, "A apuração financeira consolidada indicou valor pago efetivo de ")
    vp = _valor_pago_total(dados)
    if vp is not None:
        _adicionar_run(p, formatar_moeda(vp), negrito=True)
    else:
        _run_campo_manual(p, "Valor pago efetivo total")
    _adicionar_run(p, " e valor teórico calculado de ")
    vat = _valor_atualizado_total(dados)
    if vat is not None:
        _adicionar_run(p, formatar_moeda(vat), negrito=True)
    else:
        _run_campo_manual(p, "Valor teorico calculado total")
    _adicionar_run(p, ", resultando em valor retroativo a pagar de ")
    retro = _retroativo_total(dados)
    if retro is not None:
        _adicionar_run(p, formatar_moeda(retro), negrito=True)
    else:
        _run_campo_manual(p, "Valor retroativo total a pagar")
    _adicionar_run(p, ".")


def _ds_quadro2_financeiro(doc: Document, dados: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Quadro 2 – Apuração financeira por ciclo")
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(10)

    cabecalho = ["Ciclo", "Valor pago efetivo", "Valor teórico calculado", "Retroativo a pagar"]
    linhas = []
    for lin in _linhas_financeiro(dados):
        linhas.append([
            lin.get("ciclo", ""),
            formatar_moeda(lin.get("valor_pago")),
            formatar_moeda(lin.get("valor_atualizado")),
            formatar_moeda(lin.get("delta")),
        ])
    if not linhas:
        linhas = [["—", NAO_INFORMADO, NAO_INFORMADO, NAO_INFORMADO]]
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ds_par6_data_corte(doc: Document, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "6. ", negrito=True)
    _adicionar_run(p,
        "Para fins de consolidação contratual, foi adotada premissa de considerar "
        "para fins de cálculo do retroativo e, consequente cálculo do valor "
        "remanescente do contrato, a data de ")
    _texto_ou_marcador(p, _campo(cm, "data_corte_descricao"), "Ex: abril de 2026, ultimo mes com liquidacao")
    _adicionar_run(p, ".")


def _ds_quadro3_vta(doc: Document, dados: dict, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Quadro 3 – Memória fiscal do Valor Total Atualizado Estimado")
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(10)

    cabecalho = ["Descrição", "Valor"]
    linhas = _montar_linhas_vta(dados)
    if not linhas:
        linhas = [["Valor Total Atualizado Estimado", formatar_moeda(dados.get("vta"))]]
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _montar_linhas_vta(dados: dict) -> list[list[str]]:
    """Monta as linhas da composicao VTA a partir das parcelas canonicas."""
    parcelas = dados.get("parcelas_vta") or []
    linhas = []
    vta_total = dados.get("vta")
    for p in parcelas:
        fonte = str(p.get("fonte_parcela") or "")
        ciclo = str(p.get("ciclo") or "")
        descricao = str(p.get("descricao") or p.get("justificativa_vta") or f"{ciclo} - {fonte}")
        valor = _num_ou_none(p.get("valor") or p.get("valor_atualizado"))
        linhas.append([descricao, formatar_moeda(valor)])
    if vta_total is not None:
        linhas.append(["Valor Total Atualizado Estimado (VTA)", formatar_moeda(vta_total)])
    return linhas


def _secao_valores_unitarios_por_ciclo(doc: Document, dados: dict) -> None:
    """Etapa 7: tabela 'Valores Unitarios por Ciclo' (Saneador e Apostila).

    Estrutura canonica unica (dados["historico_vu"]): C0 sempre e os ciclos ate
    o ultimo efetivamente analisado. Ciclos futuros nao entram. Nao inventa zeros:
    celula sem valor sai vazia. Uma linha por item; valores em R$ com 2 casas.
    """
    hvu = dados.get("historico_vu") or {}
    itens = hvu.get("itens") or []
    ciclos = hvu.get("ciclos") or []
    if not itens or not ciclos:
        return

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Valores Unitários por Ciclo")
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(10)

    cabecalho = ["Item", "Descrição"] + [f"VU {c}" for c in ciclos]
    linhas: list[list[str]] = []
    for reg in itens:
        vus = reg.get("vus") or {}
        linha = [
            str(reg.get("item") or ""),
            str(reg.get("descricao") or ""),
        ]
        for c in ciclos:
            valor = vus.get(c)
            # Sem valor -> celula vazia (nunca zero artificial).
            linha.append(formatar_moeda(valor) if valor is not None else "")
        linhas.append(linha)
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ds_par7_composicao_vta(doc: Document, dados: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "7. ", negrito=True)
    _adicionar_run(p,
        "De forma didática, o Valor Total Atualizado Estimado do Contrato "
        "pode ser lido pela seguinte composição: soma dos valores executados "
        "atualizados em cada ciclo, acrescidos dos aditivos computáveis e "
        "deduzidas as supressões, totalizando ")
    vta = dados.get("vta")
    if vta is not None:
        _adicionar_run(p, formatar_moeda(vta), negrito=True)
    else:
        _run_campo_manual(p, "Valor Total Atualizado Estimado")
    _adicionar_run(p, ".")


def _ds_par8_aditivos(doc: Document, dados: dict) -> None:
    aditivos = dados.get("aditivos") or []
    if not aditivos:
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "8. ", negrito=True)
    _adicionar_run(p, "Quanto aos aditivos e supressões, registra-se: ")
    partes = []
    for ad in aditivos:
        ident = ad.get("identificador") or ad.get("ciclo") or "Aditivo"
        val = _num_ou_none(ad.get("valor_atualizado"))
        val_str = formatar_moeda(val) if val is not None else NAO_INFORMADO
        partes.append(f"{ident} com valor atualizado de {val_str}")
    _adicionar_run(p, "; ".join(partes) + ".")


def _ds_par9_adequacao(doc: Document, cm: dict) -> None:
    ref = _campo(cm, "adequacao_orcamentaria_ref")
    val = _campo(cm, "adequacao_orcamentaria_valor")
    # Secao condicional: so exibe se referencia for fornecida
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "9. ", negrito=True)
    _adicionar_run(p, "Foi realizada a adequação orçamentária necessária ao prosseguimento da instrução, no valor de ")
    _valor_moeda_ou_marcador(p, val, "Valor da adequacao orcamentaria")
    _adicionar_run(p, ", conforme documento ")
    _texto_ou_marcador(p, ref, "Referencia do documento de adequacao orcamentaria")
    _adicionar_run(p, ".")


def _ds_par10_regularidade(doc: Document, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "10. ", negrito=True)
    _adicionar_run(p, "As certidões de regularidade estão presentes em ")
    _texto_ou_marcador(p, _campo(cm, "regularidade_ref"), "Referencia das certidoes de regularidade")
    _adicionar_run(p, ".")


def _ds_par11_concordancia(doc: Document, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "11. ", negrito=True)
    _adicionar_run(p, "A contratada manifestou concordância com os valores propostos conforme registrado em ")
    _texto_ou_marcador(p, _campo(cm, "concordancia_ref"), "Referencia da manifestacao de concordancia da contratada")
    _adicionar_run(p, ".")


def _ds_par12_garantia(doc: Document) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "12. ", negrito=True)
    _adicionar_run(p,
        "A contratada foi informada da necessidade de apresentação do endosso "
        "da garantia contratual, quando aplicável, observando-se o prazo e as "
        "condições previstos no contrato.")


def _ds_par13_docs_desatualizados(doc: Document, cm: dict) -> None:
    docs = _campo(cm, "docs_desatualizados")
    if not docs:
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "13. ", negrito=True)
    _adicionar_run(p,
        "Após atualizações e alinhamentos internos, alguns documentos instruídos "
        "mostram-se desatualizados, devendo ser desconsiderados: ")
    if isinstance(docs, list):
        _adicionar_run(p, ", ".join(str(d) for d in docs))
    else:
        _adicionar_run(p, str(docs))
    _adicionar_run(p, ".")


def _ds_par14_conclusao(doc: Document, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "14. ", negrito=True)
    _adicionar_run(p,
        "Diante do exposto, estando conferidos os elementos documentais, "
        "financeiros e formais acima indicados, e inexistindo pendência crítica "
        "impeditiva, a instrução poderá prosseguir para formalização do "
        "Termo de Apostila relativo ao Contrato ")
    _texto_ou_marcador(p, _campo(cm, "contrato"), "Numero do contrato")
    _adicionar_run(p, ".")


def _ds_quadro4_sintese(doc: Document, dados: dict, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Quadro 4 – Síntese dos principais valores")
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(10)

    var = dados.get("var_acumulada")
    vat = _valor_atualizado_total(dados)
    retro = _retroativo_total(dados)
    vta = dados.get("vta")
    adeq_val = _campo(cm, "adequacao_orcamentaria_valor")

    cabecalho = ["Parcela", "Valor"]
    val_orig = _campo(cm, "valor_original_contrato")
    linhas = [
        ["Valor original do contrato", formatar_moeda(val_orig) if val_orig else PREENCHER_TAG.format("Valor original do contrato")],
        ["Variação percentual acumulada", _fmt_pct_doc(var) if var is not None else PREENCHER_TAG.format("Variacao acumulada")],
        ["Valor teórico calculado total", formatar_moeda(vat) if vat is not None else PREENCHER_TAG.format("Valor teorico calculado")],
        ["Retroativo total a pagar", formatar_moeda(retro) if retro is not None else PREENCHER_TAG.format("Retroativo total")],
        ["Valor Total Atualizado Estimado (VTA)", formatar_moeda(vta) if vta is not None else PREENCHER_TAG.format("VTA")],
        ["Adequação orçamentária", formatar_moeda(adeq_val) if adeq_val is not None else PREENCHER_TAG.format("Valor da adequacao orcamentaria")],
    ]
    _adicionar_tabela(doc, cabecalho, linhas)


# ---------------------------------------------------------------------------
# Construcao do Termo de Apostila
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
    _ta_clausulas(doc, dados, campos_manuais)
    _ta_assinaturas(doc, campos_manuais)

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _ta_titulo(doc: Document, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("MINUTA DE TERMO DE APOSTILAMENTO")
    run.bold = True
    run.font.name = "Calibri"
    run.font.size = Pt(12)

    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _adicionar_run(p2, "Contrato nº ")
    _texto_ou_marcador(p2, _campo(cm, "contrato"), "Numero do contrato")
    doc.add_paragraph()


def _ta_qualificacao(doc: Document, cm: dict) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p,
        "A TELECOMUNICAÇÕES BRASILEIRAS S.A. – TELEBRAS, empresa pública "
        "federal, inscrita no CNPJ sob o nº 33.200.056/0001-41, com sede no "
        "SAS Quadra 05, Bloco H, Brasília/DF, neste ato representada por seu ")
    _texto_ou_marcador(p, _campo(cm, "representante_telebras_titulo"), "Titulo do representante da Telebras")
    _adicionar_run(p, ", ")
    _texto_ou_marcador(p, _campo(cm, "representante_telebras_nome"), "Nome do representante da Telebras")
    _adicionar_run(p, ", doravante denominada CONTRATANTE, e a empresa ")
    _texto_ou_marcador(p, _campo(cm, "representante_contratada_qualificacao"), "Qualificacao da contratada")
    _adicionar_run(p, ", representada por ")
    _texto_ou_marcador(p, _campo(cm, "representante_contratada_nome"), "Nome do representante da contratada")
    _adicionar_run(p, ", doravante denominada CONTRATADA, acordam o presente Apostilamento.")
    doc.add_paragraph()


def _ta_considerandos(doc: Document, dados: dict, cm: dict) -> None:
    p_tit = doc.add_paragraph()
    _adicionar_run(p_tit, "CONSIDERANDO:", negrito=True)

    # Clausula de reajuste
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "• A Cláusula ")
    _texto_ou_marcador(p, _campo(cm, "clausula_reajuste"), "Clausula do contrato que preve o reajuste")
    _adicionar_run(p, " do Contrato nº ")
    _texto_ou_marcador(p, _campo(cm, "contrato"), "Numero do contrato")
    _adicionar_run(p, " que prevê o reajuste contratual;")

    # Ata da Diretoria
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p2, "• A deliberação da Diretoria Executiva da Telebras, consignada em ")
    _texto_ou_marcador(p2, _campo(cm, "ata_diretoria"), "Referencia da Ata da reuniao da Diretoria Executiva")
    _adicionar_run(p2, ";")

    # Despacho Saneador
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p3, "• O Despacho Saneador constante de ")
    _texto_ou_marcador(p3, _campo(cm, "despacho_saneador_ref"), "Referencia do Despacho Saneador")
    _adicionar_run(p3, ";")

    # Memoria de calculo
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p4, "• A memória de cálculo constante em ")
    _texto_ou_marcador(p4, _campo(cm, "memoria_calculo_ref"), "Referencia da memoria de calculo")
    _adicionar_run(p4, ";")

    # Memoria financeira
    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p5, "• A memória financeira disponível em ")
    _texto_ou_marcador(p5, _campo(cm, "memoria_financeira_ref"), "Referencia da memoria financeira")
    _adicionar_run(p5, ";")

    # Informacoes gestora
    p6 = doc.add_paragraph()
    p6.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p6, "• As informações prestadas pela área gestora do contrato em ")
    _texto_ou_marcador(p6, _campo(cm, "informacoes_gestora_ref"), "Referencia das informacoes da area gestora do contrato")
    _adicionar_run(p6, ";")

    # Concordancia
    p7 = doc.add_paragraph()
    p7.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p7, "• A manifestação da CONTRATADA constante de ")
    _texto_ou_marcador(p7, _campo(cm, "concordancia_ref"), "Referencia da manifestacao de concordancia da contratada")
    _adicionar_run(p7, ", pela qual anuiu com os cálculos apresentados;")

    # Certidoes e adequacao
    p8 = doc.add_paragraph()
    p8.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p8, "• As certidões de regularidade da CONTRATADA juntadas em ")
    _texto_ou_marcador(p8, _campo(cm, "regularidade_ref"), "Referencia das certidoes de regularidade")
    _adicionar_run(p8, " e a adequação orçamentária registrada em ")
    _texto_ou_marcador(p8, _campo(cm, "adequacao_orcamentaria_ref"), "Referencia do documento de adequacao orcamentaria")
    _adicionar_run(p8, ";")

    # Aditivos
    aditivos = dados.get("aditivos") or []
    for i, ad in enumerate(aditivos, 1):
        pa = doc.add_paragraph()
        pa.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        ident = ad.get("identificador") or f"{i}º Termo Aditivo"
        val = _num_ou_none(ad.get("valor_atualizado"))
        val_str = formatar_moeda(val) if val is not None else NAO_INFORMADO
        _adicionar_run(pa, f"• O {ident}, com valor atualizado de {val_str};")

    doc.add_paragraph()


def _ta_clausulas(doc: Document, dados: dict, cm: dict) -> None:
    # Clausula 1: apostilamento dos reajustes
    p1 = doc.add_paragraph()
    p1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p1, "1. ", negrito=True)
    _adicionar_run(p1, "Procede-se ao apostilamento do Contrato nº ")
    _texto_ou_marcador(p1, _campo(cm, "contrato"), "Numero do contrato")
    _adicionar_run(p1,
        " para formalizar a concessão dos reajustes contratuais apurados, "
        "nos seguintes termos:")

    # Um bullet por ciclo computado
    for c in dados.get("ciclos_computados") or []:
        pc = doc.add_paragraph()
        pc.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pct = c.get("percentual_reajuste")
        pct_str = _fmt_pct_doc(pct) if pct is not None else NAO_INFORMADO
        _adicionar_run(pc, f"• {c['ciclo']}: ")
        _adicionar_run(pc, f"percentual de {pct_str}, com efeitos financeiros a partir de ")
        _adicionar_run(pc, str(c.get("data_inicio", NAO_INFORMADO)))
        _adicionar_run(pc, ";")

    var = dados.get("var_acumulada")
    pvar = doc.add_paragraph()
    pvar.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(pvar, "• Percentual acumulado apurado: ")
    if var is not None:
        _adicionar_run(pvar, _fmt_pct_doc(var), negrito=True)
    else:
        _run_campo_manual(pvar, "Percentual acumulado")
    _adicionar_run(pvar, ".")

    # Clausula 2: apuracao financeira
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p2, "2. ", negrito=True)
    _adicionar_run(p2,
        "Para fins de apuração financeira e cálculo do retroativo a pagar, "
        "foram consolidados os seguintes valores:")

    # Tabela financeira
    _ta_tabela1_financeiro(doc, dados)

    # Clausula 3: VTA
    p3 = doc.add_paragraph()
    p3.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p3, "3. ", negrito=True)
    _adicionar_run(p3,
        "Para fins de consolidação do Valor Total Atualizado Estimado do "
        "Contrato, adotou-se como data de corte ")
    _texto_ou_marcador(p3, _campo(cm, "data_corte_descricao"), "Ex: abril de 2026, ultimo mes com liquidacao")
    _adicionar_run(p3, ".")

    # Tabela composicao VTA
    _ta_tabela2_vta(doc, dados)

    # Etapa 7: historico de Valores Unitarios por ciclo (C0..ultimo analisado).
    _secao_valores_unitarios_por_ciclo(doc, dados)

    # Clausula 4: residual
    p4 = doc.add_paragraph()
    p4.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p4, "4. ", negrito=True)
    _adicionar_run(p4,
        "Registra-se que o valor residual atualizado corresponde ao montante "
        "remanescente do contrato após a aplicação dos reajustes apurados.")

    # Clausula 5: VTA consolidado
    p5 = doc.add_paragraph()
    p5.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p5, "5. ", negrito=True)
    _adicionar_run(p5, "Assim, o Valor Total Atualizado Estimado do Contrato fica consolidado em ")
    vta = dados.get("vta")
    if vta is not None:
        _adicionar_run(p5, formatar_moeda(vta), negrito=True)
    else:
        _run_campo_manual(p5, "Valor Total Atualizado Estimado")
    _adicionar_run(p5, ".")

    # Clausula 6: demais clausulas inalteradas
    p6 = doc.add_paragraph()
    p6.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p6, "6. ", negrito=True)
    _adicionar_run(p6,
        "Permanecem inalteradas e em pleno vigor as demais cláusulas e "
        "condições do contrato ora apostilado.")

    # Clausula 7: garantia
    p7 = doc.add_paragraph()
    p7.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p7, "7. ", negrito=True)
    _adicionar_run(p7, "A CONTRATADA deverá atualizar a garantia contratual, nos termos da Cláusula ")
    _texto_ou_marcador(p7, _campo(cm, "clausula_garantia"), "Clausula do contrato sobre garantia contratual")
    obs = _campo(cm, "garantia_clausula_obs")
    if obs:
        _adicionar_run(p7, f" ({obs})")
    _adicionar_run(p7, " do contrato.")

    # Clausula 8: vinculacao
    p8 = doc.add_paragraph()
    p8.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p8, "8. ", negrito=True)
    _adicionar_run(p8, "O presente apostilamento vincula-se, para todos os fins, aos documentos ")
    _texto_ou_marcador(p8, _campo(cm, "vinculacao_docs"), "Documentos de vinculacao")
    _adicionar_run(p8, " instruídos no Processo ")
    _texto_ou_marcador(p8, _campo(cm, "processo_vinculacao"), "Processo de vinculacao")
    _adicionar_run(p8, ".")

    doc.add_paragraph()


def _ta_tabela1_financeiro(doc: Document, dados: dict) -> None:
    cabecalho = ["Ciclo", "Valor nominal pago", "Valor devido atualizado", "Retroativo a pagar"]
    linhas = []
    for lin in _linhas_financeiro(dados):
        linhas.append([
            lin.get("ciclo", ""),
            formatar_moeda(lin.get("valor_pago")),
            formatar_moeda(lin.get("valor_atualizado")),
            formatar_moeda(lin.get("delta")),
        ])
    if not linhas:
        linhas = [["—", NAO_INFORMADO, NAO_INFORMADO, NAO_INFORMADO]]
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ta_tabela2_vta(doc: Document, dados: dict) -> None:
    cabecalho = ["Parcela", "Valor"]
    linhas = _montar_linhas_vta(dados)
    if not linhas:
        vta = dados.get("vta")
        linhas = [["Valor Total Atualizado Estimado", formatar_moeda(vta)]]
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ta_assinaturas(doc: Document, cm: dict) -> None:
    p_local = doc.add_paragraph()
    p_local.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _texto_ou_marcador(p_local, _campo(cm, "local_data"), "Brasilia/DF, dd/mm/aaaa.")

    doc.add_paragraph()
    doc.add_paragraph()

    # Assinatura Telebras
    p_cont = doc.add_paragraph()
    _adicionar_run(p_cont, "TELECOMUNICAÇÕES BRASILEIRAS S.A. – TELEBRAS")
    p_rep = doc.add_paragraph()
    _texto_ou_marcador(p_rep, _campo(cm, "representante_telebras_nome"), "Nome do representante da Telebras")
    p_tit = doc.add_paragraph()
    _texto_ou_marcador(p_tit, _campo(cm, "representante_telebras_titulo"), "Titulo do representante da Telebras")

    doc.add_paragraph()

    # Assinatura contratada
    p_contratada = doc.add_paragraph()
    _adicionar_run(p_contratada, "CONTRATADA")
    p_rep2 = doc.add_paragraph()
    _texto_ou_marcador(p_rep2, _campo(cm, "representante_contratada_nome"), "Nome do representante da contratada")
    p_qual = doc.add_paragraph()
    _texto_ou_marcador(p_qual, _campo(cm, "representante_contratada_qualificacao"), "Qualificacao do representante da contratada")


# ---------------------------------------------------------------------------
# Diagnostico de campos manuais
# ---------------------------------------------------------------------------

def diagnosticar_campos_manuais(
    leitura_ou_objeto: dict,
    identificacao: dict | None = None,
    campos_manuais: dict | None = None,
) -> list[dict]:
    """Retorna lista de campos manuais pendentes de preenchimento.

    Cada item: {campo, descricao, documento}
    """
    if campos_manuais is None:
        campos_manuais = {}

    dados = _extrair_dados(leitura_ou_objeto, identificacao)
    pendentes = []

    for chave, descricao, documento in TODOS_CAMPOS_MANUAIS:
        val = _campo(campos_manuais, chave)
        # Campo opcional: garantia_clausula_obs nao e pendencia critica
        if chave == "garantia_clausula_obs":
            continue
        # docs_desatualizados: so pendente se explicitamente solicitado (opcional)
        if chave == "docs_desatualizados":
            continue
        if val is None:
            pendentes.append({"campo": chave, "descricao": descricao, "documento": documento})

    # Deduplica preservando ordem
    vistos: set[str] = set()
    resultado = []
    for item in pendentes:
        if item["campo"] not in vistos:
            vistos.add(item["campo"])
            resultado.append(item)

    return resultado
