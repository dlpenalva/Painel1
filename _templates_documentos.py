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

import re
from datetime import datetime
from io import BytesIO
from typing import Any

from docx import Document
from docx.enum.table import WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from _sumario_executivo import (
    NAO_HOUVE_PEDIDO,
    NAO_INFORMADO,
    formatar_moeda,
    montar_dados_sumario_executivo,
    _num_ou_none,
)
from _objeto_processo_reajuste import obter_objeto_processo_reajuste
from _reajuste_utils import (
    FRASE_SEM_CICLOS_COMPUTADOS,
    expressao_quantidade_ciclos,
    gerado_em_brasilia,
    tem_sem_pedido,
)
from _sanitizacao_documental import remover_emojis_leve
from _metodo_apuracao import normalizar_metodo

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
        # user-facing IST (Anatel); legado "IST (Série Local)" ainda reconhecido.
        return "IST (Anatel)"
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
    ("empresa_contratada", "Nome da empresa contratada", "despacho"),
    ("objeto_contrato", "Objeto resumido do contrato", "despacho"),
    ("vigencia_ate", "Data final da vigencia contratual", "despacho"),
    ("tipo_atualizacao", "Tipo da atualizacao contratual", "despacho"),
    ("processo_pleito", "Referencias do pleito da contratada", "despacho"),
    ("referencia_analise", "Referencia onde o resultado da analise consta", "despacho"),
    ("memoria_calculo_ref", "Referencia da memoria de calculo", "despacho"),
    ("adequacao_orcamentaria_ref", "Referencia da adequacao orcamentaria", "despacho"),
    ("adequacao_orcamentaria_valor", "Valor da adequacao orcamentaria", "despacho"),
    ("regularidade_ref", "Referencia das certidoes de regularidade", "despacho"),
    ("regularidade_situacao", "Situacao da regularidade da contratada", "despacho"),
    ("concordancia_ref", "Referencia da manifestacao de concordancia da contratada", "despacho"),
    ("concordancia_situacao", "Situacao da concordancia da contratada", "despacho"),
    ("garantia_situacao", "Situacao da garantia contratual", "despacho"),
    ("docs_desatualizados", "Lista de documentos a desconsiderar (opcional)", "despacho"),
    ("pendencias_complemento", "Complemento manual das pendencias (opcional)", "despacho"),
]

CAMPOS_MANUAIS_TERMO = [
    ("contrato", "Numero do contrato", "termo"),
    ("empresa_contratada", "Nome/qualificacao da empresa contratada", "termo"),
    ("clausula_reajuste", "Clausula contratual do reajuste", "termo"),
    ("deliberacao_institucional",
     "Deliberacao institucional aplicavel (opcional)", "termo"),
    ("instrumentos_posteriores",
     "Instrumentos posteriores considerados (opcional)", "termo"),
    ("representante_telebras_1_nome", "Nome do 1o representante da Telebras", "termo"),
    ("representante_telebras_1_matricula", "Matricula do 1o representante", "termo"),
    ("representante_telebras_2_cargo", "Cargo do 2o representante da Telebras", "termo"),
    ("representante_telebras_2_matricula", "Matricula do 2o representante", "termo"),
    ("solicitacao_data", "Data da solicitacao da contratada", "termo"),
    ("solicitacao_ref", "Referencia documental da solicitacao da contratada", "termo"),
    ("memoria_calculo_ref", "Referencia da memoria de calculo", "termo"),
    ("concordancia_ref", "Referencia da manifestacao de concordancia da contratada", "termo"),
    ("regularidade_ref", "Referencia das certidoes de regularidade", "termo"),
    ("adequacao_orcamentaria_ref", "Referencia da adequacao orcamentaria", "termo"),
    ("processo_ref", "Numero do processo de instrucao", "termo"),
    ("valor_pago_efetivo", "Valor pago efetivo (quando nao apurado automaticamente)", "termo"),
    ("valor_teorico", "Valor devido apos o reajuste (quando nao apurado automaticamente)", "termo"),
    ("valor_original_contrato", "Valor original do contrato", "termo"),
    ("local_data", "Data (ex.: 20/07/2026)", "termo"),
]

TODOS_CAMPOS_MANUAIS = list(
    {c[0]: c for c in CAMPOS_MANUAIS_DESPACHO + CAMPOS_MANUAIS_TERMO}.values()
)

# Campos que sao opcionais (nao entram como pendencia critica no diagnostico).
_CAMPOS_OPCIONAIS = {
    "docs_desatualizados", "pendencias_complemento", "valor_pago_efetivo",
    "valor_teorico", "deliberacao_institucional", "instrumentos_posteriores",
}


# ---------------------------------------------------------------------------
# Helpers XML / DOCX
# ---------------------------------------------------------------------------

# Etapa 26H — politica documental da PREVIA: numero XLS oficial sem resultado
# definitivo e exibido como "R$ x — PREVIA", com highlight verde somente na
# palavra PREVIA. Nunca declara VALIDADO nem resolve a divergencia (26C).
ROTULO_PREVIA = "PRÉVIA"
SUFIXO_PREVIA = f" — {ROTULO_PREVIA}"
COR_HIGHLIGHT_PREVIA = "green"


def _vta_texto_doc(dados: dict) -> str:
    """Texto documental do VTA: valor oficial, PREVIA do XLS, ou vazio."""
    vta = dados.get("vta")
    if vta is not None:
        return formatar_moeda(vta)
    vta_previa = dados.get("vta_previa")
    if vta_previa is not None:
        return f"{formatar_moeda(vta_previa)}{SUFIXO_PREVIA}"
    return ""


def _set_highlight(run, cor: str = "yellow") -> None:
    rPr = run._r.get_or_add_rPr()
    highlight = OxmlElement("w:highlight")
    highlight.set(qn("w:val"), cor)
    rPr.append(highlight)


def _repetir_cabecalho(tabela) -> None:
    tr = tabela.rows[0]._tr
    trPr = tr.get_or_add_trPr()
    tblHeader = OxmlElement("w:tblHeader")
    tblHeader.set(qn("w:val"), "true")
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


# VTA-POT-1: amarelo-palha muito claro, o mesmo do XLS, da web e do PDF.
# Destaca APENAS a linha da parcela POTENCIAL — nunca o total, nunca vermelho.
COR_SHADING_POTENCIAL = "FFF4CC"


def _sombrear_linha(row, cor: str) -> None:
    """Aplica sombreamento suave a uma linha de tabela (somente apresentacao)."""
    for celula in row.cells:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), cor)
        celula._tc.get_or_add_tcPr().append(shd)


def _configurar_box_discreto(paragrafo) -> None:
    ppr = paragrafo._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), "F2F2F2")
    ppr.append(shd)
    bordas = OxmlElement("w:pBdr")
    for lado in ("top", "left", "bottom", "right"):
        borda = OxmlElement(f"w:{lado}")
        borda.set(qn("w:val"), "single")
        borda.set(qn("w:sz"), "6")
        borda.set(qn("w:space"), "6")
        borda.set(qn("w:color"), "BFBFBF")
        bordas.append(borda)
    ppr.append(bordas)


def _adicionar_box_retroativos(doc: Document, dados: dict, *, saneador: bool) -> None:
    situacao = dados.get("situacao_retroativos_pc") or {}
    if not situacao:
        return
    reconhecido = _num_ou_none(situacao.get("reconhecido"))
    em_analise = _num_ou_none(situacao.get("em_analise"))
    potencial = _num_ou_none(situacao.get("potencial"))
    # PC-VTA-POT-TOTAL-1: o apurado LIQUIDO (`potencial`) e o INCORPORADO ao
    # VTA sao medidas diferentes quando ha parcela potencial negativa.
    # Publicar as duas sob o mesmo rotulo produziria dois numeros distintos
    # chamados "retroativo potencial" no mesmo documento.
    incorporado = _num_ou_none(dados.get("vta_retroativo_potencial"))
    discriminar = (
        incorporado is not None and potencial is not None
        and round(incorporado, 2) != round(potencial, 2)
    )
    valores_exibidos = (
        (reconhecido, em_analise, potencial, incorporado)
        if saneador else (reconhecido, potencial, incorporado)
    )
    if not any(
        valor is not None and abs(valor) > 0.004
        for valor in valores_exibidos
    ):
        return

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _configurar_box_discreto(p)
    _adicionar_run(p, "SITUAÇÃO DOS VALORES RETROATIVOS", negrito=True, tamanho=10)
    p.add_run().add_break()
    _adicionar_run(p, f"Retroativo reconhecido: {formatar_moeda(reconhecido)}", tamanho=10)
    if saneador:
        p.add_run().add_break()
        _adicionar_run(
            p,
            "Valor em análise pela área gestora: " + formatar_moeda(em_analise),
            tamanho=10,
        )
    p.add_run().add_break()
    if discriminar:
        _adicionar_run(
            p,
            "Retroativo potencial incorporado ao VTA: "
            + formatar_moeda(incorporado),
            tamanho=10,
        )
        p.add_run().add_break()
        _adicionar_run(
            p,
            "Retroativo potencial apurado (líquido, informativo): "
            + formatar_moeda(potencial),
            tamanho=10,
        )
    else:
        _adicionar_run(
            p, f"Retroativo potencial: {formatar_moeda(potencial)}", tamanho=10
        )
    p.add_run().add_break()
    p.add_run().add_break()

    if saneador:
        texto = (
            "O retroativo reconhecido integra a apuração corrente. "
            "O retroativo potencial está associado aos Pedidos de Compra ainda "
            "em análise pela área gestora, a quem competem a confirmação e os "
            "procedimentos relacionados ao eventual pagamento.\n\n"
            "Enquanto não houver confirmação, o retroativo potencial não integra "
            "o valor reconhecido a pagar."
        )
    else:
        texto = (
            "O retroativo reconhecido integra a apuração corrente. O retroativo "
            "potencial está associado aos Pedidos de Compra ainda em análise pela "
            "área gestora. Sua confirmação e eventual pagamento competem à área "
            "gestora.\n\nEnquanto não confirmado, o retroativo potencial não "
            "integra o valor reconhecido a pagar."
        )
    if discriminar:
        # Regra 2 (PC-VTA-POT-TOTAL-1): o negativo permanece visivel, nao
        # reduz o VTA e nao compensa parcela potencial positiva.
        texto += (
            " O valor apurado é líquido: parcelas potenciais negativas "
            "permanecem visíveis, não reduzem o Valor Total Atualizado e "
            "não compensam as parcelas potenciais positivas."
        )
    partes = texto.split("\n")
    for indice, parte in enumerate(partes):
        if indice:
            p.add_run().add_break()
        if parte:
            _adicionar_run(p, parte, tamanho=10)


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


def _adicionar_id_apuracao_rodape(doc: Document, dados: dict) -> None:
    """Registra a rastreabilidade sem interferir no corpo juridico do documento.

    So escreve quando ha id_apuracao (nunca em modelo em branco, cujo `dados`
    nao carrega essa chave). A data/hora acompanha o mesmo bloco/run — nao e
    a apuracao, o upload ou o commit: e o instante em que estes bytes
    especificos foram montados (fonte unica: gerado_em_brasilia()).
    """
    id_apuracao = str(dados.get("id_apuracao") or "").strip()
    if not id_apuracao:
        return
    paragrafo = doc.sections[0].footer.paragraphs[0]
    paragrafo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = paragrafo.add_run(
        f"ID da apuração: {id_apuracao} | Gerado em {gerado_em_brasilia()}"
    )
    run.font.name = "Calibri"
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)


def _adicionar_tabela(
    doc: Document,
    cabecalho: list[str],
    linhas: list[list[str]],
    *,
    repetir_cabecalho: bool = True,
    destacar_placeholders: bool = False,
    destacar_placeholders_embutidos: bool = False,
    linhas_destaque: set[int] | None = None,
) -> Any:
    """Monta a tabela do documento.

    ``linhas_destaque`` recebe indices de ``linhas`` (base 0, sem contar o
    cabecalho) que saem no amarelo-palha da parcela POTENCIAL (VTA-POT-1).
    """
    n_cols = len(cabecalho)
    tabela = doc.add_table(rows=1, cols=n_cols)
    tabela.style = "Table Grid"
    tabela.rows[0].height = Pt(16)
    tabela.rows[0].height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
    celulas_cab = tabela.rows[0].cells
    for i, texto in enumerate(cabecalho):
        celulas_cab[i].text = ""
        run = celulas_cab[i].paragraphs[0].add_run(remover_emojis_leve(texto))
        run.bold = True
        run.font.name = "Calibri"
        run.font.size = Pt(10)
    if repetir_cabecalho:
        _repetir_cabecalho(tabela)
    destaque = linhas_destaque or set()
    for indice_linha, linha in enumerate(linhas):
        row = tabela.add_row()
        row.height = Pt(16)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        if indice_linha in destaque:
            _sombrear_linha(row, COR_SHADING_POTENCIAL)
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
            paragrafo_celula = row.cells[i].paragraphs[0]
            if texto.endswith(SUFIXO_PREVIA):
                # Etapa 26H: highlight verde SOMENTE na palavra PREVIA.
                run = paragrafo_celula.add_run(texto[: -len(ROTULO_PREVIA)])
                run.font.name = "Calibri"
                run.font.size = Pt(10)
                run_previa = paragrafo_celula.add_run(ROTULO_PREVIA)
                run_previa.font.name = "Calibri"
                run_previa.font.size = Pt(10)
                _set_highlight(run_previa, COR_HIGHLIGHT_PREVIA)
                continue
            partes = (
                re.split(r"(\[PREENCHER:[^\]]+\])", texto)
                if destacar_placeholders_embutidos and "[PREENCHER:" in texto
                else [texto]
            )
            for parte in partes:
                if not parte:
                    continue
                run = paragrafo_celula.add_run(parte)
                run.font.name = "Calibri"
                run.font.size = Pt(10)
                if destacar_placeholders and parte.startswith("[PREENCHER:"):
                    _set_highlight(run, "yellow")
                elif negativo and "R$" in str(celula_texto):
                    run.font.color.rgb = COR_NEGATIVO
    return tabela


# ---------------------------------------------------------------------------
# Extracao de dados canonicos
# ---------------------------------------------------------------------------

def _extrair_dados(leitura_ou_objeto: dict, identificacao: dict | None) -> dict:
    dados = montar_dados_sumario_executivo(leitura_ou_objeto, identificacao)
    if not dados.get("disponivel"):
        return {
            "disponivel": False,
            "identificacao_externa": dict(identificacao or {}),
            "pendencias": {},
        }

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
    if not dados_op and isinstance(leitura_ou_objeto, dict):
        dados_op = leitura_ou_objeto
    vta_sombra = dados_op.get("vta_sombra") or {}
    parcelas_vta = vta_sombra.get("parcelas_computadas") or []
    controle_operacional = dados_op.get("controle") or {}
    # Metodo CANONICO (_metodo_apuracao): 'pc' | 'principal' | 'd' | ''.
    # Exposicao de sinal ja existente em CONTROLE!B1 — nenhuma heuristica
    # nova e nenhuma regra de negocio criada aqui.
    metodo_canonico = normalizar_metodo(controle_operacional.get("modo"))
    metodo_pc = metodo_canonico == "pc"
    situacao_pc = _situacao_retroativos_pc(dados_op) if metodo_pc else None

    aditivos = []
    for ad in aditivos_raw:
        aditivos.append({
            "identificador_interno": ad.get("identificador_interno"),
            "rotulo_documental": (
                ad.get("rotulo_documental") or ad.get("ciclo")
            ),
            "instrumento": ad.get("instrumento"),
            "item": ad.get("item"),
            "tipo_alteracao": ad.get("tipo_alteracao"),
            "ciclo": ad.get("ciclo"),
            "data_alteracao": ad.get("data_alteracao"),
            "quantidade": ad.get("quantidade"),
            "valor_original": ad.get("valor_original"),
            "valor_atualizado": ad.get("valor_atualizado"),
        })

    return {
        "disponivel": True,
        "id_apuracao": dados.get("id_apuracao"),
        "ciclos": ciclos,
        "ciclos_reajuste": ciclos_reajuste,
        "ciclos_computados": ciclos_computados,
        "var_acumulada": sintese.get("variacao_acumulada"),
        "vta": sintese.get("vta"),
        "vta_previa": sintese.get("vta_previa"),
        "vta_execucao_atualizada": sintese.get("vta_execucao_atualizada"),
        "vta_saldo_remanescente_atualizado": sintese.get(
            "vta_saldo_remanescente_atualizado"
        ),
        # VTA-POT-1: parcela prudencial lida pronta da sintese canonica.
        "vta_sem_potencial": sintese.get("vta_sem_potencial"),
        "vta_retroativo_potencial": sintese.get("vta_retroativo_potencial"),
        "vta_retroativo_potencial_apurado": sintese.get(
            "vta_retroativo_potencial_apurado"
        ),
        "vta_tem_parcela_potencial": bool(
            sintese.get("vta_tem_parcela_potencial")
        ),
        "vta_tem_potencial_apurado": bool(
            sintese.get("vta_tem_potencial_apurado")
        ),
        "fin_por_ciclo": fin_por_ciclo,
        "pc_por_ciclo": pc_por_ciclo,
        "parcelas_vta": parcelas_vta,
        "aditivos": aditivos,
        "financeiro": financeiro,
        "sintese": sintese,
        "identificacao": dados.get("identificacao") or {},
        "identificacao_externa": dict(identificacao or {}),
        "pendencias": objeto_proc.get("pendencias") or {},
        "historico_vu": dados.get("historico_vu") or {},
        "referencias_vta": (
            (leitura_ou_objeto or {}).get("referencias_vta")
            or (dados.get("referencias_vta") if isinstance(dados, dict) else None)
            or {}
        ),
        "metodo": metodo_canonico,
        "metodo_pc": metodo_pc,
        "data_corte": (
            controle_operacional.get("data_corte")
            or (situacao_pc or {}).get("data_corte")
        ),
        "situacao_retroativos_pc": situacao_pc,
    }


def _situacao_retroativos_pc(dados_operacionais: dict) -> dict[str, Any] | None:
    """Expõe nos documentos o mesmo consolidado canônico usado em RESULTADOS.

    A presença da chave é preservada: valor ausente continua ``None`` e zero
    conhecido continua ``0.0``. Os documentos não completam lacunas com zero.
    """
    totais_pc = (dados_operacionais.get("itens_pc_v10") or {}).get(
        "totais_canonicos"
    ) or dados_operacionais.get("totais_canonicos_pc") or {}
    ate_o_corte = totais_pc.get("ate_o_corte") or {}
    if not totais_pc or not ate_o_corte:
        return None

    def _valor(chave: str) -> float | None:
        if chave not in ate_o_corte:
            return None
        valor = _num_ou_none(ate_o_corte.get(chave))
        return round(valor, 2) if valor is not None else None

    return {
        "quantidade": ate_o_corte.get("quantidade"),
        "quantidade_reconhecida": ate_o_corte.get("quantidade_reconhecida"),
        "original_reconhecido": _valor("valor_original_reconhecido"),
        "atualizado_reconhecido": _valor("valor_atualizado_reconhecido"),
        "reconhecido": _valor("retroativo"),
        "em_analise": _valor("valor_atualizado_em_analise"),
        "potencial": _valor("delta_potencial"),
        "por_ciclo": totais_pc.get("por_ciclo") or {},
        "intervalo_precluso": totais_pc.get("intervalo_precluso") or {},
        "indeterminado": totais_pc.get("indeterminado") or {},
        "posterior_ao_corte": totais_pc.get("posterior_ao_corte") or {},
        "posterior_ao_corte_por_ciclo": (
            totais_pc.get("posterior_ao_corte_por_ciclo") or {}
        ),
        "data_corte": totais_pc.get("data_corte"),
        "corte_aplicado": bool(totais_pc.get("corte_aplicado")),
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


def _formatar_competencia(valor: Any) -> str | None:
    """Converte uma data já apurada para mm/aaaa, sem alterar o valor-fonte."""
    texto = str(valor or "").strip()
    if not texto or texto == NAO_INFORMADO:
        return None
    for formato, tamanho in (("%d/%m/%Y", 10), ("%Y-%m-%d", 10), ("%m/%Y", 7)):
        try:
            return datetime.strptime(texto[:tamanho], formato).strftime("%m/%Y")
        except ValueError:
            continue
    return None


def _sem_pedido_ciclo(c: dict) -> bool:
    """Ciclo em que a CONTRATADA nao apresentou pedido.

    Fonte unica: o marcador gravado na propria SITUACAO do ciclo. Data de
    pedido ausente, sozinha, NAO caracteriza "nao houve pedido" — continua
    sendo ausencia de informacao (legado), e a redacao de ausencia atual segue
    valendo para esse caso.
    """
    return tem_sem_pedido(c.get("situacao"))


def _data_pedido_documental(c: dict) -> str:
    """Data do pedido para uso em texto corrido; vazia quando nao ha data real."""
    if _sem_pedido_ciclo(c):
        return ""
    texto = str(c.get("data_pedido") or "").strip()
    if texto in ("", NAO_INFORMADO, NAO_HOUVE_PEDIDO):
        return ""
    return texto


def _efeito_financeiro_ciclo(c: dict) -> str:
    """Frase administrativa de efeitos financeiros de um ciclo."""
    situacao = remover_emojis_leve(c.get("situacao") or "").strip().lower()
    inicio = _formatar_competencia(c.get("inicio_efeito_financeiro"))
    if "preclu" in situacao:
        return "Sem efeitos financeiros"
    if inicio:
        return f"A partir de {inicio}"
    return NAO_INFORMADO


_MESES_EXTENSO = (
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
)


def _lista_natural(itens: list[str]) -> str:
    if not itens:
        return ""
    if len(itens) == 1:
        return itens[0]
    return ", ".join(itens[:-1]) + " e " + itens[-1]


def _competencias_sem_efeito(c: dict) -> list[str]:
    """Competencias do ciclo que nao produzem efeitos financeiros.

    Fonte unica: o bloco `meses_sem_efeito` que a apuracao ja consolida a
    partir dos marcos canonicos (inicio do ciclo x INICIO_EFEITO_FINANCEIRO).
    O gerador NAO recria temporalidade: sem status "ok" nao ha o que declarar.
    """
    situacao = remover_emojis_leve(c.get("situacao") or "").strip().lower()
    if "preclu" in situacao:
        # Preclusao integral ja e declarada como "Sem efeitos financeiros";
        # nunca vira perda parcial de competencias.
        return []
    bloco = c.get("meses_sem_efeito") or {}
    if str(bloco.get("status") or "") != "ok":
        return []
    return [str(x).strip() for x in (bloco.get("competencias") or []) if str(x).strip()]


def _competencias_por_extenso(competencias: list[str]) -> str:
    """'01/2026', '02/2026' -> 'janeiro e fevereiro de 2026'."""
    grupos: list[tuple[str, list[str]]] = []
    for comp in competencias:
        mes_txt, _, ano = str(comp).partition("/")
        try:
            nome_mes = _MESES_EXTENSO[int(mes_txt) - 1]
        except (ValueError, IndexError):
            nome_mes = mes_txt
        if grupos and grupos[-1][0] == ano:
            grupos[-1][1].append(nome_mes)
        else:
            grupos.append((ano, [nome_mes]))
    return _lista_natural([f"{_lista_natural(m)} de {ano}" for ano, m in grupos])


def _frase_perda_efeitos(c: dict, *, nomear_ciclo: bool) -> str | None:
    competencias = _competencias_sem_efeito(c)
    if not competencias:
        return None
    inicio = _formatar_competencia(c.get("inicio_efeito_financeiro"))
    if not inicio:
        return None
    ciclo = remover_emojis_leve(c.get("ciclo") or "").strip()
    referencia = f"do ciclo {ciclo}" if (nomear_ciclo and ciclo) else "deste ciclo"
    rotulo = "a competência de" if len(competencias) == 1 else "as competências de"
    return (
        "Em razão da data do pedido, os efeitos financeiros do reajuste "
        f"{referencia} iniciam-se em {inicio}, não alcançando {rotulo} "
        f"{_competencias_por_extenso(competencias)}."
    )


def _paragrafos_perda_efeitos(doc: Document, dados: dict) -> None:
    """Declara, ciclo a ciclo, as competencias sem efeitos financeiros.

    Compartilhada pelo Despacho Saneador e pelo Termo de Apostila para que os
    dois documentos declarem a mesma perda a partir da mesma fonte temporal.
    """
    if dados.get("_modo_branco"):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _adicionar_run(
            p,
            "Havendo competências não alcançadas pelos efeitos financeiros do "
            "reajuste em razão da data do pedido, deverão ser expressamente "
            "indicadas neste item.",
        )
        return
    ciclos = dados.get("ciclos_computados") or []
    nomear = len(ciclos) > 1
    for c in ciclos:
        frase = _frase_perda_efeitos(c, nomear_ciclo=nomear)
        if not frase:
            continue
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _adicionar_run(p, frase)


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


ROTULO_PARCELA_POTENCIAL = "Retroativo potencial — POTENCIAL"


def _texto_parcela_potencial(dados: dict) -> str:
    """Frase unica da regra prudencial.

    Potencial incorporado (> 0) -> frase da parcela somada ao VTA.
    Potencial apurado negativo  -> frase informativa: o valor e dito, mas o
    VTA NAO foi reduzido por ele (piso prudencial).
    """
    potencial = _num_ou_none(dados.get("vta_retroativo_potencial"))
    if dados.get("vta_tem_parcela_potencial") and potencial and round(potencial, 2):
        return (
            f"O Valor Total Atualizado inclui {formatar_moeda(potencial)} de "
            "retroativo potencial, considerado por critério prudencial. A parcela "
            "permanece sujeita à confirmação pela área gestora e não representa, "
            "nesta data, retroativo reconhecido a pagar."
        )
    apurado = _num_ou_none(dados.get("vta_retroativo_potencial_apurado"))
    if apurado is not None and round(apurado, 2) < 0:
        return (
            "POTENCIAL — NÃO INCORPORADO AO VTA. Parcela potencial apurada: "
            f"{formatar_moeda(apurado)}. Por critério prudencial, valores "
            "potenciais negativos não reduzem o VTA."
        )
    return ""


def _paragrafo_parcela_potencial(doc: Document, dados: dict) -> None:
    texto = _texto_parcela_potencial(dados)
    if not texto:
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, texto)


def _composicao_didatica_vta(dados: dict) -> list[tuple[str, float | None]]:
    """Agrupa as parcelas em componentes didaticos (execucao por ciclo, saldo,
    aditivos), somando por rubrica. Nunca inventa; apenas soma o que existe.

    VTA-POT-1: no metodo PC o VTA passa a carregar o retroativo POTENCIAL. Ele
    entra como PARCELA PROPRIA e identificada — nunca diluido na execucao nem
    no saldo —, e a conferencia de fechamento passa a considerar as tres
    parcelas. O valor vem pronto da cadeia canonica; aqui nada e recalculado.
    """
    executado = _num_ou_none(dados.get("vta_execucao_atualizada"))
    saldo = _num_ou_none(dados.get("vta_saldo_remanescente_atualizado"))
    potencial = _num_ou_none(dados.get("vta_retroativo_potencial")) or 0.0
    vta = _num_ou_none(dados.get("vta"))
    if vta is None:
        vta = _num_ou_none(dados.get("vta_previa"))
    if (
        executado is not None
        and saldo is not None
        and vta is not None
        and abs(round(executado + saldo + potencial, 2) - round(vta, 2)) <= 0.01
    ):
        linhas: list[tuple[str, float | None]] = [
            ("Execução atualizada anterior ao corte", round(executado, 2)),
            ("Saldo remanescente atualizado no corte", round(saldo, 2)),
        ]
        if round(potencial, 2):
            linhas.append((ROTULO_PARCELA_POTENCIAL, round(potencial, 2)))
        return linhas

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


TITULO_HISTORICO_VU = "HISTÓRICO DOS VALORES UNITÁRIOS POR CICLO"


def montar_historico_vu_documental(dados: dict) -> dict:
    """Estrutura neutra do quadro de VUs, unica para Saneador e Apostila.

    Fonte: dados["historico_vu"] (aba historico_VU via sumario executivo),
    ja truncada em C0..ultimo ciclo da analise SEM filtrar ciclos historicos
    com COMPUTAR=Nao (o quadro e historico contratual). Nunca inventa zeros.
    """
    hvu = dados.get("historico_vu") or {}
    itens = hvu.get("itens") or []
    ciclos = hvu.get("ciclos") or []
    if not itens or not ciclos:
        return {"disponivel": False, "cabecalhos": [], "linhas": [],
                "ciclo_final": None}
    cabecalhos = ["Item"] + [f"VU_{c}" for c in ciclos]
    linhas: list[list[str]] = []
    for reg in itens:
        vus = reg.get("vus") or {}
        linha = [str(reg.get("item") or "")]
        for c in ciclos:
            valor = vus.get(c)
            linha.append(formatar_moeda(valor) if valor is not None else "")
        linhas.append(linha)
    return {"disponivel": True, "cabecalhos": cabecalhos, "linhas": linhas,
            "ciclo_final": hvu.get("ultimo_ciclo")}


def _secao_valores_unitarios_por_ciclo(
    doc: Document, dados: dict, texto_intro: str | None = None,
    *, titulo: str | None = None, repetir_cabecalho: bool = True,
) -> None:
    """Renderiza o HISTORICO DOS VALORES UNITARIOS POR CICLO (C0..ultimo).

    A paginacao em blocos e preservada; cada bloco reemite o cabecalho e o
    marca como linha de cabecalho, para que a continuacao permaneca legivel.
    """
    quadro = montar_historico_vu_documental(dados)
    if not quadro["disponivel"]:
        return
    if texto_intro:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _adicionar_run(p, texto_intro)
    rotulo = titulo or TITULO_HISTORICO_VU
    _titulo_quadro(doc, rotulo)
    linhas = quadro["linhas"]
    blocos = [linhas[:10]]
    restante = linhas[10:]
    while restante:
        blocos.append(restante[:12])
        restante = restante[12:]
    for indice, bloco in enumerate(blocos):
        if indice:
            doc.add_page_break()
            _titulo_quadro(doc, f"{rotulo} (continuação)")
        _adicionar_tabela(
            doc,
            quadro["cabecalhos"],
            bloco,
            repetir_cabecalho=repetir_cabecalho,
        )
    doc.add_paragraph()


# ---------------------------------------------------------------------------
# TERMO DE APOSTILAMENTO (modelo canonico §6)
# ---------------------------------------------------------------------------

def gerar_termo_apostila(
    leitura_ou_objeto: dict,
    identificacao: dict | None = None,
    campos_manuais: dict | None = None,
    *,
    modo_modelo_em_branco: bool = False,
) -> bytes:
    """Gera o Termo de Apostila em DOCX e retorna os bytes.

    Estrutura do modelo aprovado em 09/09/2026: qualificacao, CONSIDERANDO,
    secoes 1 a 8, assinaturas e ANEXO 1 (historico dos valores unitarios) ao
    final. A secao 2 tem redacao propria por metodo canonico (PC, Financeiro
    e Itens Consumidos). `modo_modelo_em_branco=True` produz a mesma estrutura
    sem afirmar fato algum ainda nao comprovado.

    O documento e camada de APRESENTACAO: nenhum valor e recalculado aqui.
    """
    if campos_manuais is None:
        campos_manuais = {}
    dados = _extrair_dados(leitura_ou_objeto, identificacao)
    dados["_modo_branco"] = bool(modo_modelo_em_branco)
    doc = _configurar_documento()

    _ta_titulo(doc, dados, campos_manuais)
    _ta_qualificacao(doc, campos_manuais)
    _ta_considerandos(doc, dados, campos_manuais)
    _ta_abertura(doc)
    _ta_secao1_reajustes(doc, dados, campos_manuais)
    _ta_secao2_retroativo(doc, dados, campos_manuais)
    _ta_secao3_composicao_vta(doc, dados)
    _ta_secao4_valores_unitarios(doc, dados)
    _ta_secao5_aditivos(doc, dados)
    _ta_secoes_finais(doc, dados, campos_manuais)
    _ta_assinaturas(doc, campos_manuais)
    _ta_anexo1_valores_unitarios(doc, dados)
    _adicionar_id_apuracao_rodape(doc, dados)

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Etapa 29B — modelos em branco (wrappers finos, sem duplicar estrutura).
# Nao leem session_state; produzem bytes deterministicos a partir de {} apenas.
# ---------------------------------------------------------------------------

def gerar_modelo_branco_despacho() -> bytes:
    """Modelo em branco do Despacho Saneador (sem dados de Coleta/sessao)."""
    return gerar_despacho_saneador({}, {}, {}, modo_modelo_em_branco=True)


def gerar_modelo_branco_termo() -> bytes:
    """Modelo em branco do Termo de Apostila (sem dados de Coleta/sessao)."""
    return gerar_termo_apostila({}, {}, {}, modo_modelo_em_branco=True)


# Modelo aprovado em 09/09/2026. A parcela potencial e SEMPRE nomeada em caixa
# baixa: a natureza ("nao integra o valor reconhecido a pagar") e explicada no
# texto, nunca incorporada ao nome da parcela.
ROTULO_POTENCIAL_TERMO = "retroativo potencial"
ROTULO_POTENCIAL_APURADO_TERMO = "retroativo potencial apurado (líquido, informativo)"
TITULO_TERMO = "TERMO DE APOSTILA"
TITULO_TERMO_MODELO = "TERMO DE APOSTILA - MODELO PADRÃO"
TITULO_ANEXO_VU = "ANEXO 1 - HISTÓRICO DOS VALORES UNITÁRIOS POR CICLO"


def _ta_titulo(doc: Document, dados: dict, cm: dict) -> None:
    titulo = TITULO_TERMO_MODELO if dados.get("_modo_branco") else TITULO_TERMO
    _titulo_secao(doc, titulo, tamanho=12,
                  alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _adicionar_run(p, "Contrato nº ")
    _texto_ou_marcador(p, _campo(cm, "contrato"), "Numero do contrato")
    doc.add_paragraph()


def _ta_par(doc: Document, numero: str) -> Any:
    """Paragrafo numerado do corpo do Termo (ex.: "2.1.")."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, f"{numero}. ")
    return p


def _ta_potenciais(dados: dict) -> tuple[float | None, float | None, bool]:
    """(incorporado ao VTA, apurado liquido, sao_diferentes).

    Consome valores JA PRONTOS da cadeia canonica — nada e recalculado. As
    duas grandezas so coincidem quando nao ha parcela potencial negativa; por
    isso nunca sao publicadas sob o mesmo nome (PC-VTA-POT-TOTAL-1).
    """
    incorporado = _num_ou_none(dados.get("vta_retroativo_potencial"))
    if incorporado is not None and not round(incorporado, 2):
        incorporado = None
    if incorporado is not None and not dados.get("vta_tem_parcela_potencial"):
        incorporado = None
    situacao = dados.get("situacao_retroativos_pc") or {}
    apurado = _num_ou_none(situacao.get("potencial"))
    if apurado is not None and not round(apurado, 2):
        apurado = None
    diferentes = (
        incorporado is not None and apurado is not None
        and round(incorporado, 2) != round(apurado, 2)
    )
    return incorporado, apurado, diferentes


def _ta_tem_potencial(dados: dict) -> bool:
    incorporado, apurado, _ = _ta_potenciais(dados)
    return incorporado is not None or apurado is not None


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
    """CONSIDERANDO do modelo aprovado, com numeracao sempre sequencial.

    Itens inaplicaveis (deliberacao institucional e instrumentos posteriores
    nao informados) simplesmente nao sao emitidos no documento processado —
    sem lacuna na numeracao e sem considerando vazio.
    """
    _titulo_secao(doc, "CONSIDERANDO:")
    branco = dados.get("_modo_branco")
    contador = {"n": 0}

    def item():
        contador["n"] += 1
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _adicionar_run(p, f"{contador['n']}. ", negrito=True)
        return p

    deliberacao = _campo(cm, "deliberacao_institucional")
    instrumentos = _campo(cm, "instrumentos_posteriores")
    emite_instrumentos = bool(branco or instrumentos is not None)
    fim_adequacao = ";" if emite_instrumentos else "."

    # 1 — clausula contratual do reajuste (sem hardcode: nao ha fonte canonica)
    p1 = item()
    _texto_ou_marcador(
        p1, _campo(cm, "clausula_reajuste"), "Clausula contratual do reajuste"
    )
    _adicionar_run(p1, " do Contrato nº ")
    _texto_ou_marcador(p1, _campo(cm, "contrato"), "Numero do contrato")
    _adicionar_run(p1,
        ", que disciplina o reajuste contratual, os ciclos de apuração, a "
        "admissibilidade dos pedidos e os respectivos efeitos financeiros;")

    # 2 — deliberacao institucional: condicional, sem default institucional
    if branco:
        p2 = item()
        _adicionar_run(p2, "A deliberação institucional aplicável, quando houver: ")
        _run_campo_manual(p2, "Deliberacao institucional aplicavel")
        _adicionar_run(p2, ";")
    elif deliberacao is not None:
        p2 = item()
        texto_del = remover_emojis_leve(deliberacao).strip().rstrip(".;")
        _adicionar_run(p2, texto_del + ";")

    # 3 — solicitacao da CONTRATADA (situacao e data canonicas)
    p3 = item()
    if branco:
        _adicionar_run(p3,
            "A solicitação da CONTRATADA a ser identificada pela data ")
        _run_campo_manual(p3, "Data da solicitacao da contratada")
        _adicionar_run(p3, " e pela referência documental ")
        _run_campo_manual(p3, "Referencia documental da solicitacao da contratada")
        _adicionar_run(p3, ";")
    else:
        ciclos = dados.get("ciclos_computados") or []
        situacoes = [
            remover_emojis_leve(c.get("situacao") or "").strip().upper()
            for c in ciclos
        ]
        tempestiva = bool(situacoes) and all(
            situacao.startswith("TEMPESTIVO") for situacao in situacoes
        )
        # Nenhum ciclo computado teve pedido: nao se pode afirmar solicitacao
        # que nao existiu. Caso misto mantem a redacao ordinaria — a
        # solicitacao existiu para parte dos ciclos, e o Quadro 1 discrimina.
        if ciclos and all(_sem_pedido_ciclo(c) for c in ciclos):
            _adicionar_run(
                p3,
                "A inexistência de pedido da CONTRATADA para os ciclos "
                "analisados, que permanecem preclusos, sem efeitos "
                "financeiros;",
            )
        else:
            datas = {
                _data_pedido_documental(c)
                for c in ciclos
                if _data_pedido_documental(c)
            }
            data_canonica = next(iter(datas)) if len(datas) == 1 else None
            _adicionar_run(
                p3,
                "A solicitação tempestiva da CONTRATADA, de " if tempestiva
                else "A solicitação da CONTRATADA, de ",
            )
            _texto_ou_marcador(
                p3,
                data_canonica or _campo(cm, "solicitacao_data"),
                "Data da solicitacao da contratada",
            )
            _adicionar_run(p3, ", instruída em ")
            _texto_ou_marcador(
                p3,
                _campo(cm, "solicitacao_ref"),
                "Referencia documental da solicitacao da contratada",
            )
            _adicionar_run(p3, ";")

    # 4 — separacao entre historico formalizado e objeto da analise
    p4 = item()
    _adicionar_run(p4,
        "A necessidade de distinguir o histórico já formalizado anteriormente do "
        "objeto da presente análise, evitando duplicidade de contagem ou "
        "sobreposição de efeitos financeiros;")

    # 5 — informacoes da area gestora (nunca afirmadas como "aprovadas")
    p5 = item()
    if branco:
        _adicionar_run(p5,
            "As informações da área gestora que vierem a fundamentar a "
            "formalização deverão abranger, quando aplicável, a execução, o "
            "saldo remanescente, os itens contratuais, os aditivos/supressões e "
            "os documentos de suporte da apuração;")
    else:
        _adicionar_run(p5,
            "As informações encaminhadas pela área gestora do contrato quanto à "
            "execução, ao saldo remanescente, aos itens contratuais, aos "
            "aditivos/supressões e aos documentos de suporte da apuração")
        processo = _campo(cm, "processo_ref")
        if processo is not None:
            _adicionar_run(p5, ", instruídas no Processo ")
            _adicionar_run(p5, remover_emojis_leve(processo).strip())
        _adicionar_run(p5, ";")

    # 6 — memoria de calculo (potencial so quando existir de fato)
    p6 = item()
    if branco:
        _adicionar_run(p6, "A memória de cálculo a ser indicada em ")
        _run_campo_manual(p6, "Referencia da memoria de calculo")
        _adicionar_run(p6,
            " deverá apresentar, quando aplicável, os ciclos de reajuste, os "
            "percentuais aplicáveis, os efeitos financeiros, o retroativo "
            "reconhecido, o retroativo potencial e a composição do Valor Total "
            "Atualizado do Contrato;")
    else:
        _adicionar_run(p6, "A memória de cálculo constante em ")
        _texto_ou_marcador(p6, _campo(cm, "memoria_calculo_ref"),
                           "Referencia da memoria de calculo")
        _adicionar_run(p6,
            ", que apurou os ciclos de reajuste, os percentuais aplicáveis, os "
            "efeitos financeiros, o retroativo reconhecido")
        if _ta_tem_potencial(dados):
            _adicionar_run(p6, ", o retroativo potencial")
        _adicionar_run(p6,
            " e a composição do Valor Total Atualizado do Contrato;")

    # 7 — indice e percentual acumulado
    p7 = item()
    if branco:
        _adicionar_run(p7,
            "O índice contratual e o percentual aplicável deverão ser informados "
            "nos campos a seguir: ")
        _run_campo_manual(p7, "Indice contratual")
        _adicionar_run(p7, " e ")
        _run_campo_manual(p7, "Percentual aplicavel")
        _adicionar_run(p7, ";")
    else:
        _adicionar_run(p7, "O índice contratual utilizado na análise, qual seja ")
        indice = _indice_doc(dados)
        if indice:
            _adicionar_run(p7, indice, negrito=True)
        else:
            _run_campo_manual(p7, "Indice contratual")
        _adicionar_run(p7, ", e o percentual acumulado apurado de ")
        var = dados.get("var_acumulada")
        if var is not None:
            _adicionar_run(p7, _fmt_pct_doc(var), negrito=True)
        else:
            _run_campo_manual(p7, "Percentual acumulado apurado")
        _adicionar_run(p7, ";")

    # 8 — concordancia da CONTRATADA
    p8 = item()
    if branco:
        _adicionar_run(p8,
            "A concordância da CONTRATADA, se aplicável, deverá ser registrada em ")
        _run_campo_manual(p8,
                          "Referencia da manifestacao de concordancia da contratada")
        _adicionar_run(p8, ";")
    else:
        _adicionar_run(p8, "As manifestações de concordância da CONTRATADA "
                           "constantes em ")
        _texto_ou_marcador(p8, _campo(cm, "concordancia_ref"),
                           "Referencia da manifestacao de concordancia da contratada")
        _adicionar_run(p8, ";")

    # 9 — certidoes de regularidade
    p9 = item()
    if branco:
        _adicionar_run(p9,
            "As certidões de regularidade da CONTRATADA a referenciar em ")
        _run_campo_manual(p9, "Referencia das certidoes de regularidade")
        _adicionar_run(p9, ";")
    else:
        _adicionar_run(p9, "As certidões de regularidade da CONTRATADA "
                           "constantes em ")
        _texto_ou_marcador(p9, _campo(cm, "regularidade_ref"),
                           "Referencia das certidoes de regularidade")
        _adicionar_run(p9, ";")

    # 10 — adequacao orcamentaria (fail-safe: sem referencia, nao afirma o ato)
    p10 = item()
    ref_adequacao = _campo(cm, "adequacao_orcamentaria_ref")
    if branco or ref_adequacao is None:
        _adicionar_run(p10,
            "As manifestações relativas à adequação orçamentária a referenciar em ")
        _run_campo_manual(p10, "Referencia da adequacao orcamentaria")
        _adicionar_run(p10, fim_adequacao)
    else:
        _adicionar_run(p10,
            "As manifestações relativas à adequação orçamentária constantes em ")
        _texto_ou_marcador(p10, ref_adequacao,
                           "Referencia da adequacao orcamentaria")
        if _ta_tem_potencial(dados):
            _adicionar_run(p10,
                ", considerada a composição integral do Valor Total Atualizado "
                "do Contrato, inclusive o retroativo potencial")
        _adicionar_run(p10, fim_adequacao)

    # 11 — instrumentos posteriores: condicional
    if branco:
        p11 = item()
        _adicionar_run(p11,
            "Os instrumentos posteriores considerados, quando houver: ")
        _run_campo_manual(p11, "Instrumentos posteriores considerados")
        _adicionar_run(p11, ".")
    elif instrumentos is not None:
        p11 = item()
        texto_inst = remover_emojis_leve(instrumentos).strip().rstrip(".;")
        _adicionar_run(p11, "Os instrumentos posteriores considerados: "
                            + texto_inst + ".")
    doc.add_paragraph()


def _ta_abertura(doc: Document) -> None:
    _titulo_secao(doc, "FORMALIZA-SE O PRESENTE TERMO DE APOSTILA:")
    doc.add_paragraph()


def _ta_secao1_reajustes(doc: Document, dados: dict, cm: dict) -> None:
    _titulo_secao(doc, "1. Dos reajustes concedidos")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    branco = dados.get("_modo_branco")
    _adicionar_run(p, "1.1. Ao Contrato nº ")
    _texto_ou_marcador(p, _campo(cm, "contrato"), "Numero do contrato")
    if branco:
        _adicionar_run(p,
            ", os reajustes a serem formalizados deverão ser indicados no "
            "Quadro 1.")
    else:
        _adicionar_run(p,
            ", formalizam-se os reajustes contratuais apurados, conforme Quadro 1.")

    _titulo_quadro(doc, "Quadro 1 — Síntese dos reajustes concedidos")
    cabecalho = ["Ref.", "Ciclo", "Percentual aplicado", "Efeitos financeiros", "Situação"]
    if branco:
        _adicionar_tabela(doc, cabecalho, [[
            "[PREENCHER: Ref.]", "[PREENCHER: Ciclo]",
            "[PREENCHER: Percentual aplicável]", "[PREENCHER: Efeitos financeiros]",
            "[PREENCHER: Situação]",
        ]], destacar_placeholders=True)
        _paragrafos_perda_efeitos(doc, dados)
        doc.add_paragraph()
        return
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
    _paragrafos_perda_efeitos(doc, dados)
    doc.add_paragraph()


def _data_documental(valor: Any) -> str:
    if valor is None or not str(valor).strip():
        return NAO_INFORMADO
    if hasattr(valor, "strftime"):
        try:
            return valor.strftime("%d/%m/%Y")
        except (TypeError, ValueError):
            pass
    texto = str(valor).strip()
    for formato in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto[:10], formato).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return remover_emojis_leve(texto)


def _texto_metodo_pc(dados: dict, *, objetivo: bool) -> str:
    corte = _data_documental(dados.get("data_corte"))
    if objetivo:
        return (
            "Metodologia utilizada: Pedidos de Compra (PC). Data de corte: "
            f"{corte}. Integram a execução considerada os PCs pagos com data até "
            "o corte. PCs não pagos e com data do PC até o corte permanecem como "
            "valor em análise "
            "pela área gestora. C0 não recebe reajuste; de C1 em diante, o "
            "reajuste observa os efeitos financeiros. PCs posteriores ao corte "
            "não integram esta apuração."
        )
    return (
        "No método Pedidos de Compra (PC), a data de corte adotada é "
        f"{corte}. Os PCs pagos com data até o corte integram a execução "
        "considerada; os PCs não pagos e com data do PC até o corte permanecem "
        "como valor em análise pela "
        "área gestora. C0 não recebe reajuste e não gera retroativo. De C1 em "
        "diante, havendo efeito financeiro, a diferença dos PCs pagos constitui "
        "retroativo reconhecido e a diferença dos PCs não pagos constitui "
        "retroativo potencial. PCs posteriores à data de corte não integram "
        "esta apuração."
    )


def _linhas_pc_documentais(dados: dict) -> list[list[str]]:
    situacao = dados.get("situacao_retroativos_pc") or {}
    por_ciclo = situacao.get("por_ciclo") or {}
    linhas = []
    for ciclo in ("C0", "C1", "C2", "C3", "C4"):
        bloco = por_ciclo.get(ciclo)
        if not bloco:
            continue
        original = _num_ou_none(bloco.get("valor_original_reconhecido"))
        atualizado = _num_ou_none(bloco.get("valor_atualizado_reconhecido"))
        retro = _num_ou_none(bloco.get("retroativo"))
        linhas.append([
            ciclo,
            formatar_moeda(original),
            formatar_moeda(atualizado),
            formatar_moeda(retro),
        ])
    totais = {
        "valor_original_reconhecido": _num_ou_none(
            situacao.get("original_reconhecido")
        ),
        "valor_atualizado_reconhecido": _num_ou_none(
            situacao.get("atualizado_reconhecido")
        ),
        "retroativo": _num_ou_none(situacao.get("reconhecido")),
    }
    residuais = {}
    for chave, total in totais.items():
        if total is None:
            residuais[chave] = None
            continue
        exibido = sum(
            _num_ou_none((por_ciclo.get(ciclo) or {}).get(chave)) or 0.0
            for ciclo in ("C0", "C1", "C2", "C3", "C4")
        )
        residuais[chave] = round(total - exibido, 2)
    quantidade_residual = sum(
        int((situacao.get(chave) or {}).get("quantidade") or 0)
        for chave in ("intervalo_precluso", "indeterminado")
    )
    if quantidade_residual or any(
        valor is not None and abs(valor) > 0.004 for valor in residuais.values()
    ):
        linhas.append([
            "Outras situações até a data de corte",
            formatar_moeda(residuais["valor_original_reconhecido"]),
            formatar_moeda(residuais["valor_atualizado_reconhecido"]),
            formatar_moeda(residuais["retroativo"]),
        ])
    return linhas


def _ta_secao2_retroativo(doc: Document, dados: dict, cm: dict) -> None:
    """Secao 2, com redacao propria por METODO CANONICO.

    O metodo vem de `dados["metodo"]` (CONTROLE!B1 normalizado): nenhuma
    inferencia por presenca de tabela. Cada metodo publica apenas o que os
    seus dados canonicos sustentam.
    """
    _titulo_secao(doc, "2. Da apuração financeira dos valores retroativos")
    if dados.get("_modo_branco"):
        _ta_secao2_branco(doc)
        return
    metodo = dados.get("metodo")
    if metodo == "pc" and dados.get("situacao_retroativos_pc"):
        _ta_secao2_pc(doc, dados)
    elif metodo == "d":
        _ta_secao2_consumidos(doc, dados)
    else:
        _ta_secao2_financeiro(doc, dados, cm)


def _ta_secao2_branco(doc: Document) -> None:
    """Modelo em branco: instrui o preenchimento, sem afirmar apuracao."""
    p = _ta_par(doc, "2.1")
    _adicionar_run(p,
        "Os valores financeiros que eventualmente integrem a formalização "
        "deverão ser informados nos campos e quadros desta seção, incluindo o "
        "valor pago efetivo ")
    _run_campo_manual(p, "Valor pago efetivo")
    _adicionar_run(p, ", o valor devido após o reajuste ")
    _run_campo_manual(p, "Valor devido apos o reajuste")
    _adicionar_run(p, " e a diferença ou retroativo ")
    _run_campo_manual(p, "Valor retroativo a pagar")
    _adicionar_run(p, ", quando aplicável, conforme Quadro 2.")
    _titulo_quadro(doc, "Quadro 2 — Apuração financeira por ciclo")
    _adicionar_tabela(
        doc,
        ["Ciclo", "Valor pago efetivo", "Valor devido após o reajuste",
         "Diferença/retroativo"],
        [[
            "[PREENCHER: Ciclo]",
            "[PREENCHER: Valor pago efetivo]",
            "[PREENCHER: Valor devido apos o reajuste]",
            "[PREENCHER: Valor retroativo a pagar]",
        ]],
        destacar_placeholders=True,
    )
    doc.add_paragraph()


def _ta_secao2_pc(doc: Document, dados: dict) -> None:
    """Metodo Pedidos de Compra: execucao reconhecida e retroativo potencial."""
    situacao = dados["situacao_retroativos_pc"]
    corte = _data_documental(dados.get("data_corte"))

    p = _ta_par(doc, "2.1")
    _adicionar_run(p,
        "Para a apuração pelo método de Pedidos de Compra, são considerados "
        "como execução reconhecida os Pedidos de Compra marcados como pagos à "
        "CONTRATADA e com data até a data de corte da apuração, fixada em "
        f"{corte}.")

    p = _ta_par(doc, "2.2")
    _adicionar_run(p,
        "No ciclo inicial (C0), por se tratar do período anterior ao primeiro "
        "reajuste, os valores dos Pedidos de Compra são considerados pelos "
        "seus valores originais, sem geração de retroativo. A partir de C1, "
        "quando o Pedido de Compra estiver alcançado pelos efeitos financeiros "
        "do reajuste, a diferença entre o valor originalmente considerado e o "
        "valor reajustado corresponde ao retroativo reconhecido.")

    p = _ta_par(doc, "2.3")
    _adicionar_run(p,
        "Com esses critérios, a execução reconhecida até a data de corte "
        "corresponde a ")
    _adicionar_run(p, formatar_moeda(situacao.get("original_reconhecido")),
                   negrito=True)
    _adicionar_run(p, " em valores originais e a ")
    _adicionar_run(p, formatar_moeda(situacao.get("atualizado_reconhecido")),
                   negrito=True)
    _adicionar_run(p,
        " após a aplicação dos reajustes cabíveis, resultando em retroativo "
        "reconhecido de ")
    _adicionar_run(p, formatar_moeda(situacao.get("reconhecido")), negrito=True)
    _adicionar_run(p, ", conforme Quadro 2.")

    _titulo_quadro(doc, "Quadro 2 — Execução reconhecida e retroativo por ciclo")
    linhas = _linhas_pc_documentais(dados)
    linhas.append([
        "Total",
        formatar_moeda(situacao.get("original_reconhecido")),
        formatar_moeda(situacao.get("atualizado_reconhecido")),
        formatar_moeda(situacao.get("reconhecido")),
    ])
    _adicionar_tabela(
        doc,
        ["Ciclo", "Pedidos de Compra reconhecidos / valor original",
         "Pedidos de Compra reconhecidos / valor atualizado",
         "Retroativo reconhecido"],
        linhas,
    )
    doc.add_paragraph()

    incorporado, apurado, diferentes = _ta_potenciais(dados)
    if incorporado is None and apurado is None:
        # Sem parcela potencial material: nao ha 2.4-2.7 nem Quadro 3.
        return

    # A frase do piso prudencial so cabe quando a parcela apurada e de fato
    # negativa. Potencial positivo sem VTA disponivel nao e "parcela negativa".
    negativo = apurado is not None and round(apurado, 2) < 0
    p = _ta_par(doc, "2.4")
    _adicionar_run(p,
        "Além do retroativo reconhecido, foram identificados Pedidos de Compra "
        "ainda sujeitos à validação pela área gestora.")
    if incorporado is not None:
        _adicionar_run(p,
            " O retroativo potencial incorporado ao Valor Total Atualizado do "
            f"Contrato corresponde a {formatar_moeda(incorporado)}.")
        if diferentes:
            _adicionar_run(p,
                " O retroativo potencial apurado, líquido e informativo, é de "
                f"{formatar_moeda(apurado)}.")
    elif apurado is not None:
        _adicionar_run(p,
            " O retroativo potencial apurado corresponde a "
            f"{formatar_moeda(apurado)}.")
    if negativo:
        _adicionar_run(p,
            " Por critério prudencial, parcela potencial negativa permanece "
            "visível, não reduz o Valor Total Atualizado do Contrato e não "
            "compensa parcela potencial positiva.")

    p = _ta_par(doc, "2.5")
    _adicionar_run(p,
        "Para fins de transparência, a situação dos valores retroativos "
        "apurados fica consolidada no Quadro 3.")

    _titulo_quadro(doc, "Quadro 3 — Situação dos valores retroativos")
    linhas_q3 = [[
        "Retroativo reconhecido",
        formatar_moeda(situacao.get("reconhecido")),
        "Sim, já refletido na execução atualizada",
        "Sim",
        "Reconhecido",
    ]]
    destaque: set[int] = set()
    if incorporado is not None:
        destaque.add(len(linhas_q3))
        linhas_q3.append([
            ROTULO_POTENCIAL_TERMO,
            formatar_moeda(incorporado),
            "Sim",
            "Não",
            "Aguardando validação pela área gestora",
        ])
        if diferentes:
            destaque.add(len(linhas_q3))
            linhas_q3.append([
                ROTULO_POTENCIAL_APURADO_TERMO,
                formatar_moeda(apurado),
                "Não",
                "Não",
                "Parcela potencial negativa não reduz o Valor Total Atualizado"
                if negativo else
                "Valor apurado, distinto da parcela incorporada",
            ])
    elif apurado is not None:
        destaque.add(len(linhas_q3))
        linhas_q3.append([
            ROTULO_POTENCIAL_APURADO_TERMO if negativo else ROTULO_POTENCIAL_TERMO,
            formatar_moeda(apurado),
            "Não" if negativo else NAO_INFORMADO,
            "Não",
            "Parcela potencial negativa não reduz o Valor Total Atualizado"
            if negativo else "Aguardando validação pela área gestora",
        ])
    _adicionar_tabela(
        doc,
        ["Natureza", "Valor", "Integra o Valor Total Atualizado",
         "Integra o valor reconhecido a pagar nesta data", "Situação"],
        linhas_q3,
        linhas_destaque=destaque,
    )
    doc.add_paragraph()

    if incorporado is not None:
        p = _ta_par(doc, "2.6")
        _adicionar_run(p,
            f"O retroativo potencial, no valor de {formatar_moeda(incorporado)}, "
            "integra o Valor Total Atualizado do Contrato e é considerado para "
            "fins de adequação orçamentária, por critério prudencial.")

    p = _ta_par(doc, "2.7" if incorporado is not None else "2.6")
    _adicionar_run(p,
        "A inclusão dessa parcela no Valor Total Atualizado não representa "
        "reconhecimento definitivo da obrigação nem autoriza, por si só, o "
        "pagamento. A conversão do retroativo potencial em retroativo "
        "reconhecido dependerá da validação dos respectivos Pedidos de Compra "
        "pela área gestora, a quem competem a confirmação desses eventos e os "
        "procedimentos relacionados ao eventual pagamento.")
    doc.add_paragraph()


def _ta_secao2_financeiro(doc: Document, dados: dict, cm: dict) -> None:
    """Metodo Financeiro (Mensalidade): execucao por competencia.

    Nenhuma mencao a Pedido de Compra e nenhuma mencao a retroativo potencial,
    que nao e grandeza deste metodo.
    """
    p = _ta_par(doc, "2.1")
    _adicionar_run(p,
        "A apuração financeira consolidada indicou, na execução realizada por "
        "competência, valor pago efetivo de ")
    vp = _valor_pago_total(dados)
    if vp is not None:
        _adicionar_run(p, formatar_moeda(vp), negrito=True)
    else:
        _valor_moeda_ou_marcador(p, _campo(cm, "valor_pago_efetivo"),
                                 "Valor pago efetivo")
    _adicionar_run(p, " e valor devido após o reajuste de ")
    vat = _valor_atualizado_total(dados)
    if vat is not None:
        _adicionar_run(p, formatar_moeda(vat), negrito=True)
    else:
        _valor_moeda_ou_marcador(p, _campo(cm, "valor_teorico"),
                                 "Valor devido apos o reajuste")
    _adicionar_run(p, ", resultando em valor retroativo a pagar de ")
    retro = _retroativo_total(dados)
    if retro is not None:
        _adicionar_run(p, formatar_moeda(retro), negrito=True)
    else:
        _run_campo_manual(p, "Valor retroativo a pagar")
    _adicionar_run(p, ", conforme Quadro 2.")

    _titulo_quadro(doc, "Quadro 2 — Apuração financeira por ciclo")
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
    _adicionar_tabela(
        doc,
        ["Ciclo", "Valor pago efetivo", "Valor devido após o reajuste",
         "Diferença/retroativo"],
        linhas,
    )
    doc.add_paragraph()


def _ta_secao2_consumidos(doc: Document, dados: dict) -> None:
    """Metodo Itens Consumidos: consumo itemizado, sem base mensal nem PC.

    Este metodo nao produz apuracao por competencia nem Pedidos de Compra; o
    documento nao afirma pagamento e nao pede valor financeiro manual. O
    detalhamento por item fica no ANEXO 1.
    """
    p = _ta_par(doc, "2.1")
    _adicionar_run(p,
        "Para a apuração pelo método de Itens Consumidos, a base considerada é "
        "o consumo itemizado declarado para o contrato, atualizado pelos ciclos "
        "de reajuste aplicáveis, sem apuração de execução por competência "
        "mensal.")

    p = _ta_par(doc, "2.2")
    _adicionar_run(p,
        "As quantidades consumidas informadas e os respectivos valores "
        "unitários, por ciclo de reajuste, constam do ANEXO 1 deste Termo de "
        "Apostila.")

    p = _ta_par(doc, "2.3")
    retro = _retroativo_total(dados)
    if retro is not None:
        _adicionar_run(p,
            "O retroativo resultante da atualização do consumo declarado "
            "corresponde a ")
        _adicionar_run(p, formatar_moeda(retro), negrito=True)
        _adicionar_run(p, ".")
    else:
        _adicionar_run(p,
            "Não consta desta apuração valor de retroativo decorrente da "
            "atualização do consumo declarado.")
    doc.add_paragraph()


def _ta_secao3_composicao_vta(doc: Document, dados: dict) -> None:
    """Composicao do VTA (Quadro 4), com as parcelas REAIS da cadeia canonica.

    Nada e recalculado: as parcelas vem de `_composicao_didatica_vta`. A
    parcela potencial, quando existe, aparece como parcela propria e nomeada
    em caixa baixa; a explicacao da natureza fica nos paragrafos seguintes.
    """
    _titulo_secao(doc, "3. Da composição do Valor Total Atualizado do Contrato")
    contador = {"n": 1}

    def par():
        contador["n"] += 1
        return _ta_par(doc, f"3.{contador['n']}")

    p = _ta_par(doc, "3.1")
    if dados.get("_modo_branco"):
        _adicionar_run(p,
            "O Valor Total Atualizado do Contrato deverá considerar a execução "
            "já realizada em valor atualizado, os saldos ainda a executar em "
            "valor atualizado, inclusive intermediários quando existirem, e os "
            "ajustes contratuais aplicáveis, quando houver.")
    else:
        _adicionar_run(p,
            "O Valor Total Atualizado do Contrato considera a execução já "
            "realizada em valor atualizado, os saldos ainda a executar em valor "
            "atualizado, inclusive intermediários quando existirem, e os "
            "ajustes contratuais aplicáveis, quando houver.")

    _titulo_quadro(doc, "Quadro 4 — Composição do Valor Total Atualizado do Contrato")
    linhas: list[list[str]] = []
    destaque_potencial: set[int] = set()
    for i, (desc, valor) in enumerate(_composicao_didatica_vta(dados)):
        rotulo = desc
        if desc == ROTULO_PARCELA_POTENCIAL:
            # Nome documental do Termo: caixa baixa, sem sufixo em caixa alta.
            rotulo = ROTULO_POTENCIAL_TERMO
            destaque_potencial.add(i)
        linhas.append([
            _LETRAS[i] if i < len(_LETRAS) else str(i + 1),
            rotulo,
            formatar_moeda(valor) if valor is not None else "",
        ])
    linhas.append([
        "Total",
        "Valor Total Atualizado do Contrato",
        _vta_texto_doc(dados),
    ])
    _adicionar_tabela(doc, ["Ref.", "Descrição", "Valor"], linhas,
                      linhas_destaque=destaque_potencial)

    if not dados.get("_modo_branco"):
        incorporado, _apurado, _dif = _ta_potenciais(dados)
        vta_txt = _vta_texto_doc(dados)
        if incorporado is not None:
            p = par()
            if vta_txt:
                _adicionar_run(p,
                    "O Valor Total Atualizado do Contrato corresponde a "
                    f"{vta_txt} e inclui expressamente "
                    f"{formatar_moeda(incorporado)} de retroativo potencial.")
            else:
                _adicionar_run(p,
                    "O Valor Total Atualizado do Contrato inclui expressamente "
                    f"{formatar_moeda(incorporado)} de retroativo potencial.")
            p = par()
            _adicionar_run(p,
                f"A parcela de {formatar_moeda(incorporado)} integra o Valor "
                "Total Atualizado do Contrato e a adequação orçamentária, mas "
                "permanece destacada em razão de sua natureza potencial e não "
                "integra, nesta data, o valor reconhecido a pagar à CONTRATADA.")
        retro = _retroativo_total(dados)
        if retro is not None and round(retro, 2):
            p = par()
            _adicionar_run(p,
                f"O retroativo reconhecido de {formatar_moeda(retro)} não é "
                "somado como parcela autônoma no Quadro 4, pois seus efeitos já "
                "estão incorporados à execução atualizada considerada na "
                "composição. Sua inclusão adicional representaria dupla "
                "contagem.")
    doc.add_paragraph()


def _ta_secao4_valores_unitarios(doc: Document, dados: dict) -> None:
    """Somente a remissao: o quadro (potencialmente enorme) vai ao ANEXO 1."""
    _titulo_secao(doc, "4. Dos valores unitários")
    p = _ta_par(doc, "4.1")
    if dados.get("_modo_branco"):
        _adicionar_run(p,
            "Os valores unitários dos itens, considerados os ciclos de reajuste "
            "aplicáveis até a presente atualização, deverão ser consolidados no "
            "quadro do ANEXO 1 deste Termo de Apostila.")
    else:
        _adicionar_run(p,
            "Os valores unitários dos itens, considerados os ciclos de reajuste "
            "aplicáveis até a presente atualização, ficam consolidados conforme "
            "quadro constante do ANEXO 1 deste Termo de Apostila.")
    doc.add_paragraph()


def _ta_grupos_aditivos(aditivos: list[dict]) -> list[tuple[str, str, str, str]]:
    """(ciclo, alteracoes consideradas, frase de impacto, valor do impacto).

    Fonte unica do agrupamento por ciclo: alimenta tanto a sintese em texto
    quanto o Quadro 5. Nao recalcula nada — apenas soma o que ja veio pronto.
    """
    grupos: dict[str, dict[str, Any]] = {}
    for ad in aditivos:
        ciclo = remover_emojis_leve(ad.get("ciclo") or "Sem ciclo").strip()
        grupo = grupos.setdefault(
            ciclo, {"total": 0.0, "tem_valor": False, "tipos": {}}
        )
        tipo = remover_emojis_leve(
            ad.get("tipo_alteracao") or "Alteração"
        ).strip()
        chave_tipo = "supressao" if "supr" in tipo.lower() else (
            "acrescimo" if "acresc" in tipo.lower()
            or "acrésc" in tipo.lower() else "alteracao"
        )
        grupo["tipos"][chave_tipo] = grupo["tipos"].get(chave_tipo, 0) + 1
        valor = _num_ou_none(ad.get("valor_atualizado"))
        if valor is not None:
            grupo["total"] += valor
            grupo["tem_valor"] = True

    def _ordem(ciclo: str) -> tuple[int, str]:
        texto = ciclo.upper()
        return (
            int(texto[1]) if len(texto) == 2 and texto[0] == "C"
            and texto[1].isdigit() else 99,
            texto,
        )

    saida: list[tuple[str, str, str, str]] = []
    for ciclo in sorted(grupos, key=_ordem):
        grupo = grupos[ciclo]
        tipos = grupo["tipos"]
        partes_tipo = []
        if tipos.get("acrescimo"):
            n = tipos["acrescimo"]
            partes_tipo.append(f"{n} acréscimo" + ("" if n == 1 else "s"))
        if tipos.get("supressao"):
            n = tipos["supressao"]
            partes_tipo.append(f"{n} supressão" if n == 1 else f"{n} supressões")
        if tipos.get("alteracao"):
            n = tipos["alteracao"]
            partes_tipo.append(f"{n} alteração" if n == 1 else f"{n} alterações")
        tipos_txt = " e ".join(partes_tipo)
        if grupo["tem_valor"]:
            valor_txt = formatar_moeda(grupo["total"])
            impacto = f"impacto atualizado total {valor_txt}"
        else:
            valor_txt = "A confirmar"
            impacto = "impacto a confirmar"
        saida.append((ciclo, tipos_txt, impacto, valor_txt))
    return saida


def _sintese_aditivos_por_ciclo(aditivos: list[dict]) -> list[str]:
    """Rotulos executivos por ciclo, sem expor chaves tecnicas internas."""
    return [
        f"{ciclo} — {tipos} — {impacto}"
        for ciclo, tipos, impacto, _valor in _ta_grupos_aditivos(aditivos)
    ]


def _ta_secao5_aditivos(doc: Document, dados: dict) -> None:
    _titulo_secao(doc, "5. Dos aditivos e supressões considerados")
    aditivos = dados.get("aditivos") or []
    p1 = _ta_par(doc, "5.1")
    if dados.get("_modo_branco"):
        _adicionar_run(p1, "Registrar os aditivos e supressões considerados: ")
        _run_campo_manual(p1, "Aditivos e supressões considerados")
        _adicionar_run(p1, ".")
    elif not aditivos:
        # Sem alteracoes: frase objetiva, nunca uma tabela vazia.
        _adicionar_run(p1,
            "Não foram identificados aditivos ou supressões específicos na "
            "base processada, sem prejuízo da conferência dos instrumentos já "
            "formalizados no processo.")
    else:
        _adicionar_run(p1,
            "Foram consideradas na apuração as alterações contratuais "
            "registradas nos respectivos ciclos, conforme Quadro 5.")
        _titulo_quadro(doc, "Quadro 5 — Aditivos e supressões considerados")
        _adicionar_tabela(
            doc,
            ["Ciclo", "Alterações consideradas", "Impacto atualizado total"],
            [[ciclo, tipos, valor]
             for ciclo, tipos, _impacto, valor in _ta_grupos_aditivos(aditivos)],
        )
        doc.add_paragraph()

    p2 = _ta_par(doc, "5.2")
    _adicionar_run(p2,
        "Os aditivos e supressões computáveis integram o Valor Total "
        "Atualizado quando não estiverem refletidos na execução atualizada, no "
        "saldo remanescente ou no valor formalizado anterior, vedada a dupla "
        "contagem.")
    espaco_apos_capitulo = doc.add_paragraph()
    espaco_apos_capitulo.paragraph_format.space_after = Pt(6)


def _ta_secoes_finais(doc: Document, dados: dict, cm: dict) -> None:
    _titulo_secao(doc, "6. Das demais condições contratuais")
    p6 = _ta_par(doc, "6.1")
    _adicionar_run(p6,
        "Permanecem inalteradas e em pleno vigor as demais cláusulas e condições "
        "do Contrato e de seus instrumentos posteriores não modificadas por este "
        "Termo de Apostila.")
    doc.add_paragraph()

    _titulo_secao(doc, "7. Da garantia contratual")
    p7 = _ta_par(doc, "7.1")
    _adicionar_run(p7,
        "A CONTRATADA deverá atualizar a garantia contratual, prevista na "
        "cláusula própria do Contrato, no prazo contratualmente estabelecido, "
        "observado o Valor Total Atualizado do Contrato")
    vta_txt = "" if dados.get("_modo_branco") else _vta_texto_doc(dados)
    if vta_txt:
        _adicionar_run(p7, " de ")
        _adicionar_run(p7, vta_txt, negrito=True)
        incorporado, _apurado, _dif = _ta_potenciais(dados)
        if incorporado is not None:
            _adicionar_run(p7,
                f", que inclui {formatar_moeda(incorporado)} de retroativo "
                "potencial")
    else:
        # VTA indisponivel/nao confiavel: nao se afirma valor (fail-closed).
        _adicionar_run(p7, " apurado nesta atualização")
    _adicionar_run(p7, ".")
    doc.add_paragraph()

    _titulo_secao(doc, "8. Da vinculação processual")
    p8 = _ta_par(doc, "8.1")
    _adicionar_run(p8,
        "O presente apostilamento vincula-se, para todos os fins, aos documentos "
        "instruídos no Processo ")
    _texto_ou_marcador(p8, _campo(cm, "processo_ref"),
                       "Numero do processo de instrucao")
    _adicionar_run(p8, ".")
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


def _ta_anexo1_valores_unitarios(doc: Document, dados: dict) -> None:
    """ANEXO 1, ao final do documento: historico dos valores unitarios.

    O quadro pode ser extenso; a paginacao em blocos e o cabecalho repetido
    ficam a cargo de `_secao_valores_unitarios_por_ciclo`. Sem limite
    artificial de itens.
    """
    quadro = montar_historico_vu_documental(dados)
    if not quadro["disponivel"]:
        return
    doc.add_page_break()
    _titulo_secao(doc, TITULO_ANEXO_VU, tamanho=12,
                  alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    _secao_valores_unitarios_por_ciclo(
        doc, dados, titulo="Tabela 1 - Valores unitários por ciclo",
    )


# ---------------------------------------------------------------------------
# DESPACHO SANEADOR (modelo canonico §7)
# ---------------------------------------------------------------------------

def gerar_despacho_saneador(
    leitura_ou_objeto: dict,
    identificacao: dict | None = None,
    campos_manuais: dict | None = None,
    *,
    modo_modelo_em_branco: bool = False,
) -> bytes:
    """Gera o Despacho Saneador em DOCX e retorna os bytes.

    `modo_modelo_em_branco=True` reutiliza a mesma estrutura enxuta sem afirmar
    fatos não comprovados. Nos dois modos, o gerador apenas apresenta dados já
    consolidados; não recalcula valores nem cria classificação processual.
    """
    if campos_manuais is None:
        campos_manuais = {}
    dados = _extrair_dados(leitura_ou_objeto, identificacao)
    dados["_modo_branco"] = bool(modo_modelo_em_branco)
    doc = _configurar_documento()

    _ds_assunto_enxuto(doc, dados, campos_manuais)
    _ds_secao1_identificacao(doc, dados, campos_manuais)
    _ds_secao2_pedido_parametros(doc, dados, campos_manuais)
    _ds_secao3_resultado(doc, dados, campos_manuais)
    _ds_secao4_documentos(doc, dados, campos_manuais)
    _ds_secao5_pendencias(doc, dados, campos_manuais)
    _ds_secao6_conclusao(doc, dados, campos_manuais)
    _adicionar_id_apuracao_rodape(doc, dados)

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def _ds_valor_identificacao(dados: dict, cm: dict, chave_manual: str,
                            *aliases: str) -> Any:
    externa = dados.get("identificacao_externa") or {}
    for alias in aliases:
        valor = externa.get(alias)
        if valor is not None and str(valor).strip():
            return valor
    return _campo(cm, chave_manual)


def _ds_texto_ou_tag(valor: Any, descricao: str) -> str:
    if valor is None or not str(valor).strip() or str(valor).strip() == NAO_INFORMADO:
        return PREENCHER_TAG.format(descricao)
    if hasattr(valor, "strftime"):
        try:
            return valor.strftime("%d/%m/%Y")
        except (TypeError, ValueError):
            pass
    return remover_emojis_leve(valor).strip()


def _ds_tipo_atualizacao(dados: dict, cm: dict) -> str | None:
    valor = _ds_valor_identificacao(
        dados, cm, "tipo_atualizacao",
        "tipo_atualizacao", "tipo_instrumento", "tipo_analise",
    )
    if valor is not None and str(valor).strip():
        return remover_emojis_leve(valor).strip()
    if dados.get("_modo_branco"):
        return None
    return "atualização contratual"


def _ds_assunto_enxuto(doc: Document, dados: dict, cm: dict) -> None:
    _titulo_secao(doc, "DESPACHO SANEADOR", tamanho=12,
                  alinhamento=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _adicionar_run(p, "Assunto: ", negrito=True)
    _adicionar_run(p, "Saneamento para formalização de ")
    tipo = _ds_tipo_atualizacao(dados, cm)
    if tipo:
        _adicionar_run(p, tipo)
    else:
        _run_campo_manual(p, "Tipo ou instrumento")
    _adicionar_run(p, " — ")
    contrato = _ds_valor_identificacao(
        dados, cm, "contrato", "contrato", "numero_contrato"
    )
    _texto_ou_marcador(p, contrato, "Numero do contrato")
    _adicionar_run(p, ".")

    p_ref = doc.add_paragraph()
    p_ref.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _adicionar_run(p_ref, "Referência(s): ", negrito=True)
    _texto_ou_marcador(
        p_ref, _campo(cm, "processo_pleito"), "Referencia do pedido"
    )
    doc.add_paragraph()


def _ds_titulo(doc: Document, numero: int, texto: str) -> Any:
    p = _titulo_secao(doc, f"{numero}. {texto.upper()}", tamanho=11)
    p.paragraph_format.keep_with_next = True
    return p


def _ds_titulo_quadro(doc: Document, texto: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    _adicionar_run(p, texto, negrito=True, tamanho=10)


def _ds_secao1_identificacao(doc: Document, dados: dict, cm: dict) -> None:
    _ds_titulo(doc, 1, "Identificação")
    contrato = _ds_valor_identificacao(
        dados, cm, "contrato", "contrato", "numero_contrato"
    )
    contratada = _ds_valor_identificacao(
        dados, cm, "empresa_contratada",
        "empresa_contratada", "contratada",
    )
    objeto = _ds_valor_identificacao(
        dados, cm, "objeto_contrato", "objeto_contrato", "objeto"
    )
    vigencia = _ds_valor_identificacao(
        dados, cm, "vigencia_ate",
        "vigencia_ate", "fim_vigencia", "data_fim_vigencia",
    )

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "Realiza-se o saneamento processual do Contrato nº ")
    _texto_ou_marcador(p, contrato, "Numero do contrato")
    _adicionar_run(p, ", celebrado com ")
    _texto_ou_marcador(p, contratada, "Nome da empresa contratada")
    _adicionar_run(p, ", cujo objeto é ")
    _texto_ou_marcador(p, objeto, "Objeto resumido do contrato")
    _adicionar_run(p, ", com vigência até ")
    if vigencia is not None:
        _adicionar_run(p, _ds_texto_ou_tag(vigencia, "Data final da vigencia contratual"))
    else:
        _run_campo_manual(p, "Data final da vigencia contratual")
    _adicionar_run(p, ".")


def _ds_ciclos_relevantes(dados: dict) -> list[dict]:
    return list(dados.get("ciclos_computados") or [])


def _ds_secao2_pedido_parametros(doc: Document, dados: dict, cm: dict) -> None:
    _ds_titulo(doc, 2, "Pedido e parâmetros da análise")
    branco = bool(dados.get("_modo_branco"))
    ciclos = _ds_ciclos_relevantes(dados)
    tipo = _ds_tipo_atualizacao(dados, cm)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if branco:
        _adicionar_run(p, "Deverão ser registrados o pedido de ")
        _run_campo_manual(p, "Tipo da atualizacao contratual")
        _adicionar_run(
            p,
            ", a data correspondente e a respectiva referência documental. "
            "Os parâmetros da análise deverão constar do Quadro 1.",
        )
    else:
        # Nenhum ciclo relevante teve pedido: o documento NAO pode afirmar
        # pedido inexistente nem apontar referencia documental de um pleito
        # que nao houve. Caso misto segue a redacao ordinaria, e o Quadro 1
        # discrimina ciclo a ciclo.
        sem_pedido_todos = bool(ciclos) and all(_sem_pedido_ciclo(c) for c in ciclos)
        if sem_pedido_todos:
            rotulos = [
                remover_emojis_leve(c.get("ciclo") or "").strip()
                for c in ciclos
            ]
            rotulos = [r for r in rotulos if r]
            _adicionar_run(p, "Não houve pedido da CONTRATADA para ")
            if len(ciclos) == 1:
                _adicionar_run(
                    p,
                    f"o ciclo {rotulos[0]}" if rotulos else "o ciclo analisado",
                )
                _adicionar_run(p, ", que permanece precluso, sem efeitos financeiros")
            else:
                _adicionar_run(
                    p,
                    f"os ciclos {', '.join(rotulos)}" if rotulos
                    else "os ciclos analisados",
                )
                _adicionar_run(p, ", que permanecem preclusos, sem efeitos financeiros")
        else:
            _adicionar_run(p, "A CONTRATADA apresentou pedido de ")
            _adicionar_run(p, tipo or "atualização contratual")
            if len(ciclos) == 1:
                _adicionar_run(p, " em ")
                pedido = _data_pedido_documental(ciclos[0])
                if pedido:
                    _adicionar_run(p, pedido)
                else:
                    _run_campo_manual(p, "Data do pedido")
            elif ciclos:
                _adicionar_run(p, " nas datas indicadas no Quadro 1")
            else:
                _adicionar_run(p, " em ")
                _run_campo_manual(p, "Data do pedido")
            _adicionar_run(p, ", conforme ")
            _texto_ou_marcador(p, _campo(cm, "processo_pleito"), "Referencia do pedido")
        if not ciclos:
            _adicionar_run(
                p,
                ". Os parâmetros de ciclo não foram disponibilizados pela fonte "
                "canônica; o Quadro 1 deve ser complementado.",
            )
        elif len(ciclos) == 1:
            _adicionar_run(p, ". A análise considerou ")
            ciclo = ciclos[0]
            _adicionar_run(p, "a data-base ")
            _adicionar_run(p, _ds_texto_ou_tag(ciclo.get("data_inicio"), "Data-base"))
            _adicionar_run(p, ", a referência econômica ")
            indice = _indice_doc(dados)
            if indice:
                _adicionar_run(p, indice)
            else:
                _run_campo_manual(p, "Indice ou referencia economica")
            _adicionar_run(
                p,
                " e classificou o ciclo como " if _sem_pedido_ciclo(ciclo)
                else " e classificou o pedido como ",
            )
            situacao = ciclo.get("situacao")
            if situacao and situacao != NAO_INFORMADO:
                _adicionar_run(p, remover_emojis_leve(situacao))
            else:
                _run_campo_manual(p, "Situacao do pedido")
            _adicionar_run(p, ", com efeitos financeiros ")
            efeito = _efeito_financeiro_ciclo(ciclo)
            if efeito != NAO_INFORMADO:
                _adicionar_run(p, efeito.lower())
            else:
                _run_campo_manual(p, "Data ou competencia do efeito financeiro")
            _adicionar_run(p, ".")
        else:
            _adicionar_run(p, ". A análise considerou ")
            _adicionar_run(
                p,
                "as datas-base, situações e efeitos financeiros indicados no "
                "Quadro 1, com referência econômica ",
            )
            indice = _indice_doc(dados)
            if indice:
                _adicionar_run(p, indice)
            else:
                _run_campo_manual(p, "Indice ou referencia economica")
            _adicionar_run(p, ".")

    _paragrafos_perda_efeitos(doc, dados)

    if dados.get("metodo_pc") and dados.get("situacao_retroativos_pc"):
        p_metodo = doc.add_paragraph()
        p_metodo.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _adicionar_run(p_metodo, _texto_metodo_pc(dados, objetivo=True))

    _ds_titulo_quadro(doc, "Quadro 1 - Síntese da análise")
    cabecalho = [
        "Ciclo", "Data-base", "Data do pedido", "Situação",
        "Efeito financeiro", "Percentual",
    ]
    if branco:
        linhas = [[PREENCHER_TAG.format(rotulo) for rotulo in cabecalho]]
    else:
        linhas = []
        for ciclo in ciclos:
            pct = ciclo.get("percentual_reajuste")
            linhas.append([
                _ds_texto_ou_tag(ciclo.get("ciclo"), "Ciclo"),
                _ds_texto_ou_tag(ciclo.get("data_inicio"), "Data-base"),
                _ds_texto_ou_tag(ciclo.get("data_pedido"), "Data do pedido"),
                _ds_texto_ou_tag(
                    remover_emojis_leve(ciclo.get("situacao") or ""), "Situacao"
                ),
                _ds_texto_ou_tag(
                    _efeito_financeiro_ciclo(ciclo), "Efeito financeiro"
                ),
                _fmt_pct_doc(pct) if pct is not None
                else PREENCHER_TAG.format("Percentual"),
            ])
        if not linhas:
            linhas = [[PREENCHER_TAG.format(rotulo) for rotulo in cabecalho]]
    _adicionar_tabela(
        doc, cabecalho, linhas,
        destacar_placeholders=True,
        destacar_placeholders_embutidos=True,
    )
    doc.add_paragraph()


def _ds_total_presente(dados: dict, chave: str) -> float | None:
    linhas = _linhas_financeiro(dados)
    if not linhas:
        return None
    valores = []
    for linha in linhas:
        valor = _num_ou_none(linha.get(chave))
        if valor is None:
            return None
        valores.append(valor)
    return round(sum(valores), 2)


def _ds_secao3_resultado(doc: Document, dados: dict, cm: dict) -> None:
    _ds_titulo(doc, 3, "Resultado essencial")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _adicionar_run(p, "Conforme memória de cálculo ")
    memoria_ref = _campo(cm, "memoria_calculo_ref") or _campo(cm, "referencia_analise")
    if dados.get("_modo_branco"):
        _run_campo_manual(p, "Referencia da memoria de calculo")
        _adicionar_run(
            p, ", os resultados essenciais deverão ser preenchidos no Quadro 2."
        )
    else:
        _texto_ou_marcador(p, memoria_ref, "Referencia da memoria de calculo")
        _adicionar_run(p, ", a apuração apresentou os resultados abaixo.")

    tipo = (_ds_tipo_atualizacao(dados, cm) or "").strip().lower()
    rotulo_devido = (
        "Valor devido após o reajuste"
        if tipo == "reajuste"
        else "Valor devido após a atualização contratual"
    )
    situacao_pc = dados.get("situacao_retroativos_pc") or {}
    if dados.get("metodo_pc") and situacao_pc:
        pago = _num_ou_none(situacao_pc.get("original_reconhecido"))
        devido = _num_ou_none(situacao_pc.get("atualizado_reconhecido"))
        retro = _num_ou_none(situacao_pc.get("reconhecido"))
        rotulo_pago = "Valor original"
        rotulo_devido = "Valor atualizado"
        rotulo_retro = "Retroativo reconhecido"
    else:
        pago = _ds_total_presente(dados, "valor_pago")
        devido = _ds_total_presente(dados, "valor_atualizado")
        retro = _retroativo_total(dados)
        rotulo_pago = "Valor pago no período analisado"
        rotulo_retro = "Retroativo a pagar"
    vta = _vta_texto_doc(dados)
    linhas = [
        [
            rotulo_pago,
            formatar_moeda(pago) if pago is not None
            else PREENCHER_TAG.format("Valor pago no periodo analisado"),
        ],
        [
            rotulo_devido,
            formatar_moeda(devido) if devido is not None
            else PREENCHER_TAG.format("Valor devido apos a atualizacao contratual"),
        ],
        [
            rotulo_retro,
            formatar_moeda(retro) if retro is not None
            else PREENCHER_TAG.format("Valor retroativo a pagar"),
        ],
    ]
    # VTA-POT-1: a parcela potencial entra ANTES do total, identificada como
    # POTENCIAL e com shading suave. O VTA continua aparecendo UMA unica vez.
    destaque_sintese: set[int] = set()
    potencial_vta = _num_ou_none(dados.get("vta_retroativo_potencial"))
    if dados.get("vta_tem_parcela_potencial") and potencial_vta:
        destaque_sintese.add(len(linhas))
        linhas.append([
            ROTULO_PARCELA_POTENCIAL, formatar_moeda(potencial_vta),
        ])
    linhas.append([
        "Valor Total Atualizado do Contrato",
        vta or PREENCHER_TAG.format("Valor Total Atualizado do Contrato"),
    ])
    _ds_titulo_quadro(doc, "Quadro 2 - Síntese financeira")
    if dados.get("metodo_pc") and situacao_pc:
        linhas_pc = _linhas_pc_documentais(dados)
        linhas_pc.append([
            "Total",
            formatar_moeda(situacao_pc.get("original_reconhecido")),
            formatar_moeda(situacao_pc.get("atualizado_reconhecido")),
            formatar_moeda(situacao_pc.get("reconhecido")),
        ])
        _adicionar_tabela(
            doc,
            ["Ciclo", "Valor original", "Valor atualizado", "Retroativo reconhecido"],
            linhas_pc,
        )
        # VTA-POT-1: no metodo PC o quadro de resultado mostra a parcela
        # POTENCIAL separada, com shading suave, e so depois o VTA — que
        # continua aparecendo UMA unica vez.
        linhas_resultado: list[list[str]] = []
        destaque_resultado: set[int] = set()
        if dados.get("vta_tem_parcela_potencial") and potencial_vta:
            destaque_resultado.add(len(linhas_resultado))
            linhas_resultado.append([
                ROTULO_PARCELA_POTENCIAL, formatar_moeda(potencial_vta),
            ])
        linhas_resultado.append([
            "Valor Total Atualizado do Contrato",
            vta or PREENCHER_TAG.format("Valor Total Atualizado do Contrato"),
        ])
        _adicionar_tabela(
            doc,
            ["Resultado", "Valor"],
            linhas_resultado,
            destacar_placeholders=True,
            destacar_placeholders_embutidos=True,
            linhas_destaque=destaque_resultado,
        )
    else:
        _adicionar_tabela(
            doc, ["Resultado", "Valor"], linhas,
            destacar_placeholders=True,
            destacar_placeholders_embutidos=True,
            linhas_destaque=destaque_sintese,
        )
    _paragrafo_parcela_potencial(doc, dados)
    doc.add_paragraph()
    _adicionar_box_retroativos(doc, dados, saneador=True)


def _ds_juntar_campos(*valores: tuple[Any, str]) -> str:
    partes = []
    for valor, descricao in valores:
        if valor is None or not str(valor).strip():
            partes.append(PREENCHER_TAG.format(descricao))
        elif isinstance(valor, (int, float)) and not isinstance(valor, bool):
            partes.append(formatar_moeda(valor))
        else:
            partes.append(remover_emojis_leve(valor).strip())
    return " / ".join(partes)


def _ds_secao4_documentos(doc: Document, dados: dict, cm: dict) -> None:
    _ds_titulo(doc, 4, "Documentos e verificações")
    memoria_ref = _campo(cm, "memoria_calculo_ref") or _campo(cm, "referencia_analise")
    linhas = [
        ["Memória de cálculo", _ds_juntar_campos(
            (memoria_ref, "Referencia da memoria de calculo"),
        )],
        ["Adequação orçamentária", _ds_juntar_campos(
            (_campo(cm, "adequacao_orcamentaria_ref"),
             "Referencia da adequacao orcamentaria"),
            (_campo(cm, "adequacao_orcamentaria_valor"),
             "Valor ou situacao da adequacao orcamentaria"),
        )],
        ["Regularidade da contratada", _ds_juntar_campos(
            (_campo(cm, "regularidade_ref"),
             "Referencia da regularidade da contratada"),
            (_campo(cm, "regularidade_situacao"),
             "Situacao da regularidade da contratada"),
        )],
        ["Concordância da contratada", _ds_juntar_campos(
            (_campo(cm, "concordancia_ref"),
             "Referencia da concordancia da contratada"),
            (_campo(cm, "concordancia_situacao"),
             "Situacao da concordancia da contratada"),
        )],
        ["Garantia contratual", _ds_juntar_campos(
            (_campo(cm, "garantia_situacao"),
             "Situacao da garantia contratual"),
        )],
    ]
    _ds_titulo_quadro(doc, "Quadro 3 - Documentos e verificações")
    _adicionar_tabela(
        doc, ["Documento ou verificação", "Referência ou situação"], linhas,
        destacar_placeholders=True,
        destacar_placeholders_embutidos=True,
    )
    doc.add_paragraph()


def _ds_pendencias_tecnicas(dados: dict, cm: dict) -> list[str]:
    resultado: list[str] = []
    pendencias = dados.get("pendencias") or {}
    for chave in ("bloqueantes", "advertencias"):
        for item in pendencias.get(chave) or []:
            texto = remover_emojis_leve(item).strip()
            if texto and texto not in resultado:
                resultado.append(texto)
    docs = _campo(cm, "docs_desatualizados")
    if docs:
        itens = docs if isinstance(docs, (list, tuple)) else [docs]
        texto = "Documentos desatualizados: " + ", ".join(str(item) for item in itens)
        resultado.append(remover_emojis_leve(texto))
    complemento = _campo(cm, "pendencias_complemento")
    if complemento:
        resultado.append(remover_emojis_leve(complemento).strip())
    if cm.get("pendencia_critica") and not resultado:
        resultado.append("Pendência impeditiva indicada para complementação.")
    return resultado


def _ds_secao5_pendencias(doc: Document, dados: dict, cm: dict) -> None:
    _ds_titulo(doc, 5, "Pendências e providências")
    pendencias = _ds_pendencias_tecnicas(dados, cm)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if dados.get("_modo_branco"):
        _adicionar_run(p, "PENDÊNCIA TÉCNICA: ", negrito=True)
        _adicionar_run(p, "Registrar as pendências relevantes para o prosseguimento: ")
        _run_campo_manual(p, "Pendencias relevantes")
        _adicionar_run(p, ".")
    elif pendencias:
        _adicionar_run(p, "PENDÊNCIA TÉCNICA: ", negrito=True)
        _adicionar_run(p, "; ".join(pendencias) + ".")
    else:
        _adicionar_run(p, "PENDÊNCIA TÉCNICA: ", negrito=True)
        _adicionar_run(p, "Não foram identificadas pendências técnicas na apuração.")
    if _ds_ha_pendencia_documental(dados, cm):
        _adicionar_run(
            p,
            " Os campos documentais destacados permanecem sujeitos a "
            "preenchimento e conferência.",
        )

    situacao = dados.get("situacao_retroativos_pc") or {}
    em_analise = _num_ou_none(situacao.get("em_analise"))
    potencial = _num_ou_none(situacao.get("potencial"))
    if dados.get("metodo_pc") and situacao and (
        (em_analise is not None and abs(em_analise) > 0.004)
        or (potencial is not None and abs(potencial) > 0.004)
    ):
        p_gestora = doc.add_paragraph()
        p_gestora.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        _adicionar_run(p_gestora, "PROVIDÊNCIA DA ÁREA GESTORA: ", negrito=True)
        _adicionar_run(
            p_gestora,
            "Há Pedidos de Compra em análise pela área gestora.",
        )
        if potencial is not None:
            _adicionar_run(
                p_gestora,
                " O retroativo potencial associado é de "
                f"{formatar_moeda(potencial)} e não integra o retroativo "
                "reconhecido nesta apuração.",
            )


def _ds_ha_pendencia_documental(dados: dict, cm: dict) -> bool:
    if dados.get("_modo_branco"):
        return True
    memoria_ref = _campo(cm, "memoria_calculo_ref") or _campo(cm, "referencia_analise")
    obrigatorios = (
        memoria_ref,
        _campo(cm, "adequacao_orcamentaria_ref"),
        _campo(cm, "adequacao_orcamentaria_valor"),
        _campo(cm, "regularidade_ref"),
        _campo(cm, "regularidade_situacao"),
        _campo(cm, "concordancia_ref"),
        _campo(cm, "concordancia_situacao"),
        _campo(cm, "garantia_situacao"),
    )
    return any(
        valor is None or (isinstance(valor, str) and not valor.strip())
        for valor in obrigatorios
    )


def _ds_tem_pendencia_impeditiva(dados: dict, cm: dict) -> bool:
    pendencias = dados.get("pendencias") or {}
    if pendencias.get("bloqueantes"):
        return True
    flag = cm.get("pendencia_critica")
    if isinstance(flag, str):
        return flag.strip().lower() in ("sim", "true", "1", "critica", "critico")
    return bool(flag)


def _ds_secao6_conclusao(doc: Document, dados: dict, cm: dict) -> None:
    _ds_titulo(doc, 6, "Conclusão")
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if dados.get("_modo_branco"):
        _adicionar_run(
            p,
            "Após o preenchimento e a conferência das informações, deverá ser "
            "avaliado se a instrução reúne condições para prosseguir à "
            "formalização.",
        )
    elif _ds_tem_pendencia_impeditiva(dados, cm):
        _adicionar_run(
            p,
            "A instrução deverá ser complementada quanto às pendências acima "
            "antes do prosseguimento para formalização.",
        )
    else:
        _adicionar_run(
            p,
            "Após a complementação e conferência das informações documentais "
            "indicadas, deverá ser avaliado o prosseguimento da instrução para "
            "formalização.",
        )


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


def _ds_par1(doc: Document, dados: dict) -> None:
    p = _ds_par(doc, "1")
    if dados.get("_modo_branco"):
        # Etapa 29C.1.2: sentido de finalidade, sem afirmar consolidacao feita.
        _adicionar_run(p,
            "Este modelo de despacho saneador destina-se à consolidação dos "
            "elementos documentais, financeiros e formais necessários à "
            "instrução de eventual Termo de Apostila de reajuste contratual. "
            "Os campos destacados devem ser revisados e preenchidos.")
        return
    _adicionar_run(p,
        "Este despacho saneador consolida os elementos documentais, financeiros "
        "e formais necessários à instrução do Termo de Apostila destinado ao "
        "registro de reajuste contratual, com a finalidade de demonstrar a "
        "regularidade mínima da instrução antes da formalização.")


def _ds_par2(doc: Document, dados: dict, cm: dict) -> None:
    p = _ds_par(doc, "2")
    if dados.get("_modo_branco"):
        # Etapa 29C.1.2: nao afirma pleito apresentado nem datas consideradas.
        _adicionar_run(p, "Registrar a referência do eventual pleito da contratada: ")
        _run_campo_manual(p, "Referencias do pleito da contratada")
        _adicionar_run(p,
            ". Para a verificação da anualidade, deverão ser informados a data "
            "da proposta em ")
        _run_campo_manual(p, "Data da proposta")
        _adicionar_run(p, ", o índice contratual ")
        _run_campo_manual(p, "Indice contratual")
        _adicionar_run(p, " e as datas de pedido aplicáveis: ")
        _run_campo_manual(p, "Datas de pedido por ciclo")
        _adicionar_run(p, ".")
        return
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
    if dados.get("_modo_branco"):
        # Etapa 29C.1.2: nao afirma acordo, concessao nem apuracao.
        _adicionar_run(p,
            "Os ciclos que poderão integrar a formalização deverão ser "
            "conferidos nos quadros deste modelo e na seguinte referência: ")
        _run_campo_manual(p, "Referencia onde o resultado da analise consta")
        _adicionar_run(p, ".")
        return
    ciclos = dados.get("ciclos_computados") or []
    if ciclos:
        nomes = ", ".join(remover_emojis_leve(c.get("ciclo") or "") for c in ciclos)
        _adicionar_run(p, f"Acordou-se na concessão de {nomes}, conforme exposto em ")
    else:
        _adicionar_run(p, "Acordou-se na concessão dos ciclos apurados, conforme exposto em ")
    _texto_ou_marcador(p, _campo(cm, "referencia_analise"), "Referencia onde o resultado da analise consta")
    _adicionar_run(p, ".")


def _ds_par4_quadro1(doc: Document, dados: dict, cm: dict) -> None:
    branco = dados.get("_modo_branco")
    p = _ds_par(doc, "4")
    if branco:
        # Etapa 29C.1.2: valor original como instrucao ("deverá ser informado"),
        # sem afirmar que a informacao ja foi prestada.
        _adicionar_run(p, "A análise deverá registrar os ciclos indicados no Quadro 1 abaixo. Quantidade de ciclos: ")
        _run_campo_manual(p, "Quantidade de ciclos objeto da análise")
        _adicionar_run(p, ", com variação acumulada de ")
        _run_campo_manual(p, "Variacao acumulada")
        _adicionar_run(p, ". O valor original do contrato deverá ser informado: ")
        _run_campo_manual(p, "Valor original do contrato")
        _adicionar_run(p, ".")
    else:
        # Conta os ciclos efetivamente considerados, nao as linhas do Quadro 1:
        # ciclos "Fora da apuracao" continuam no quadro para rastreabilidade,
        # mas afirmar que foram considerados contradiz o proprio quadro.
        # Mesma fonte canonica ja usada no item 3 e no Quadro 1 do Termo.
        expressao = expressao_quantidade_ciclos(len(dados.get("ciclos_computados") or []))
        if expressao is None:
            _adicionar_run(p, FRASE_SEM_CICLOS_COMPUTADOS)
        else:
            _adicionar_run(p, f"A análise de reajuste considerou {expressao}, com variação acumulada de ")
            var = dados.get("var_acumulada")
            if var is not None:
                _adicionar_run(p, _fmt_pct_doc(var), negrito=True)
            else:
                _run_campo_manual(p, "Variacao acumulada")
            _adicionar_run(p, ".")
        _adicionar_run(p, " O valor original do contrato informado foi de ")
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
    if branco:
        _adicionar_tabela(doc, cabecalho, [[
            "[PREENCHER: Ciclo]", "[PREENCHER: Data-base]",
            "[PREENCHER: Data do pedido]", "[PREENCHER: Início financeiro]",
            "[PREENCHER: Fim financeiro]", "[PREENCHER: Situação]",
            "[PREENCHER: Percentual aplicável]",
        ]], destacar_placeholders=True)
        doc.add_paragraph()
        return
    if not linhas:
        linhas = [["—"] * 7]
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ds_par5_quadro2(doc: Document, dados: dict) -> None:
    p = _ds_par(doc, "5")
    if dados.get("_modo_branco"):
        # Etapa 29C.1.1: o item 5 em branco nao afirma apuracao/analise/
        # consolidacao realizadas — apenas instrui o preenchimento, com os
        # mesmos tres placeholders financeiros destacados.
        _adicionar_run(p,
            "Os valores da apuração financeira deverão ser preenchidos no "
            "quadro abaixo, incluindo, quando aplicável, o valor pago efetivo, "
            "o valor teórico calculado e a diferença ou retroativo "
            "correspondente: ")
        _run_campo_manual(p, "Valor pago efetivo")
        _adicionar_run(p, ", ")
        _run_campo_manual(p, "Valor teorico calculado")
        _adicionar_run(p, " e ")
        _run_campo_manual(p, "Valor retroativo a pagar")
        _adicionar_run(p, ".")
    else:
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
    if dados.get("_modo_branco"):
        _adicionar_tabela(doc, cabecalho, [[
            "[PREENCHER: Ciclo]", "[PREENCHER: Valor pago efetivo]",
            "[PREENCHER: Valor teórico calculado]", "[PREENCHER: Valor retroativo]",
        ]], destacar_placeholders=True)
        doc.add_paragraph()
        return
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
    if dados.get("_modo_branco"):
        # Etapa 29C.1.2: a premissa de corte como obrigacao de informar, sem
        # afirmar que ja foi definida ou utilizada.
        _adicionar_run(p,
            "Deverá ser informada a premissa de corte a ser adotada para o "
            "cálculo do eventual retroativo e do valor remanescente do "
            "contrato: ")
        _run_campo_manual(p, "Descricao da data/posicao de corte adotada")
        _adicionar_run(p, ".")
    else:
        _adicionar_run(p,
            "Para fins de consolidação contratual, foi adotada a premissa de "
            "considerar, para fins de cálculo do retroativo e consequente cálculo do "
            "valor remanescente do contrato, ")
        _texto_ou_marcador(p, _campo(cm, "data_corte_descricao"), "Descricao da data/posicao de corte adotada")
        _adicionar_run(p, ".")

    _titulo_quadro(doc, "Quadro 3 - Memória fiscal do Valor Total Atualizado Estimado")
    cabecalho = ["Descrição", "Valor"]
    if dados.get("_modo_branco"):
        _adicionar_tabela(doc, cabecalho, [
            ["[PREENCHER: Descrição da parcela]", "[PREENCHER: Valor]"],
            ["Valor total do contrato estimado", "[PREENCHER: Valor Total Atualizado]"],
        ], destacar_placeholders=True)
        doc.add_paragraph()
        return
    linhas: list[list[str]] = []
    for desc, valor in _composicao_didatica_vta(dados):
        linhas.append([desc, formatar_moeda(valor) if valor is not None else ""])
    linhas.append(["Valor total do contrato estimado", _vta_texto_doc(dados)])
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ds_par7_composicao(doc: Document, dados: dict) -> None:
    p = _ds_par(doc, "7")
    _adicionar_run(p,
        "De forma didática, o Valor Total Atualizado Estimado do Contrato pode "
        "ser lido pela seguinte composição:")
    cabecalho = ["Parcela", "Valor"]
    if dados.get("_modo_branco"):
        _adicionar_tabela(doc, cabecalho, [
            ["[PREENCHER: Parcela]", "[PREENCHER: Valor]"],
            ["Valor Total Atualizado Estimado do Contrato",
             "[PREENCHER: Valor Total Atualizado]"],
        ], destacar_placeholders=True)
        doc.add_paragraph()
        return
    componentes = _composicao_didatica_vta(dados)
    linhas = [[desc, formatar_moeda(valor) if valor is not None else ""]
              for desc, valor in componentes]
    linhas.append(["Valor Total Atualizado Estimado do Contrato",
                   _vta_texto_doc(dados)])
    _adicionar_tabela(doc, cabecalho, linhas)
    doc.add_paragraph()


def _ds_bloco_historico_vu(doc: Document, dados: dict) -> None:
    """Bloco sem numeracao propria (entre os paragrafos 7 e 8): historico de
    VUs vinculado a consolidacao dos valores apurados. Preserva a numeracao
    juridica existente do despacho."""
    _secao_valores_unitarios_por_ciclo(
        doc, dados,
        texto_intro=(
            "Para fins de consolidação da evolução dos preços contratuais, "
            "apresenta-se abaixo o histórico dos valores unitários dos itens "
            "até o último ciclo considerado nesta análise."
        ),
    )


def _ds_par8_aditivos(doc: Document, dados: dict) -> None:
    p = _ds_par(doc, "8")
    aditivos = dados.get("aditivos") or []
    if dados.get("_modo_branco"):
        _adicionar_run(p, "Registrar os aditivos ou supressões considerados: ")
        _run_campo_manual(p, "Aditivos e supressões aplicáveis")
        _adicionar_run(p, ".")
        return
    if not aditivos:
        _adicionar_run(p,
            "Quanto às alterações contratuais consideradas, não foram "
            "identificados eventos específicos na base processada, sem prejuízo "
            "da conferência dos instrumentos já formalizados no processo.")
        return
    _adicionar_run(
        p,
        "Quanto às alterações contratuais consideradas, registra-se: "
        + "; ".join(_sintese_aditivos_por_ciclo(aditivos)) + ".",
    )


def _ds_par9_adequacao(doc: Document, dados: dict, cm: dict) -> None:
    p = _ds_par(doc, "9")
    if dados.get("_modo_branco"):
        _adicionar_run(p, "Registrar a adequação orçamentária, quando aplicável: ")
        _run_campo_manual(p, "Referência e valor da adequação orçamentária")
        _adicionar_run(p, ".")
        return
    _adicionar_run(p,
        "Foi realizada a adequação orçamentária necessária ao prosseguimento da "
        "instrução, no valor de ")
    _valor_moeda_ou_marcador(p, _campo(cm, "adequacao_orcamentaria_valor"), "Valor da adequacao orcamentaria")
    _adicionar_run(p, ", conforme documento ")
    _texto_ou_marcador(p, _campo(cm, "adequacao_orcamentaria_ref"), "Referencia da adequacao orcamentaria")
    _adicionar_run(p, ".")


def _ds_par10_regularidade(doc: Document, dados: dict, cm: dict) -> None:
    p = _ds_par(doc, "10")
    if dados.get("_modo_branco"):
        _adicionar_run(p, "Registrar as certidões de regularidade, quando aplicável: ")
        _run_campo_manual(p, "Referência das certidões de regularidade")
        _adicionar_run(p, ".")
        return
    _adicionar_run(p, "As certidões de regularidade estão presentes em ")
    _texto_ou_marcador(p, _campo(cm, "regularidade_ref"), "Referencia das certidoes de regularidade")
    _adicionar_run(p, ".")


def _ds_par11_concordancia(doc: Document, dados: dict, cm: dict) -> None:
    p = _ds_par(doc, "11")
    if dados.get("_modo_branco"):
        _adicionar_run(p, "Registrar a manifestação da contratada, quando aplicável: ")
        _run_campo_manual(p, "Referência da manifestação da contratada")
        _adicionar_run(p, ".")
        return
    _adicionar_run(p, "A contratada manifestou concordância com os valores propostos conforme registrado em ")
    _texto_ou_marcador(p, _campo(cm, "concordancia_ref"), "Referencia da manifestacao de concordancia da contratada")
    _adicionar_run(p, ".")


def _ds_par12_garantia(doc: Document, dados: dict) -> None:
    p = _ds_par(doc, "12")
    if dados.get("_modo_branco"):
        # Etapa 29C.1.2: nao afirma comunicacao ja realizada a contratada.
        _adicionar_run(p,
            "Registrar, quando aplicável, a comunicação à contratada sobre a "
            "necessidade de atualização ou endosso da garantia contratual, "
            "observados o prazo e as condições previstos no contrato.")
        return
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
    if dados.get("_modo_branco"):
        # Etapa 29C.1: a conclusao do modelo em branco nao afirma consolidacao,
        # saneamento nem aptidao — apenas orienta a avaliacao apos o preenchimento.
        _adicionar_run(p,
            "Após o preenchimento e a conferência dos campos deste modelo, "
            "deverá ser avaliado se a instrução reúne condições para prosseguir "
            "à formalização do Termo de Apostila, observadas as alçadas "
            "competentes e os procedimentos internos aplicáveis.")
    elif _tem_pendencia_critica(dados, cm):
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

    vta_texto = _vta_texto_doc(dados)
    if vta_texto:
        linhas.append(["Valor Total Atualizado Estimado do Contrato", vta_texto])

    adeq = _num_ou_none(_campo(cm, "adequacao_orcamentaria_valor"))
    if adeq is not None:
        linhas.append(["Adequação orçamentária registrada", formatar_moeda(adeq)])

    _adicionar_tabela(doc, cabecalho, linhas,
                      destacar_placeholders=bool(dados.get("_modo_branco")))


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
