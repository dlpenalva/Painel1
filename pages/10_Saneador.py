from datetime import datetime
from io import BytesIO
from zoneinfo import ZoneInfo
import re

import pandas as pd
import streamlit as st

st.set_page_config(page_title="TLB · cl8us - Saneador", layout="wide")

from _ui_utils import render_marca_topo


COLUNAS_INFOS = [
    "Grupo",
    "Informação necessária",
    "Documento/fonte mínima",
    "Valor / informação levantada",
    "Link SIGA (opcional)",
    "Status",
    "Observação",
]


def moeda(valor):
    try:
        valor = round(float(valor or 0), 2)
    except Exception:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def pct(valor):
    try:
        valor = float(valor or 0)
    except Exception:
        valor = 0.0
    return f"{valor * 100:.2f}%".replace(".", ",")


def texto_seguro(valor, padrao="[campo a preencher]"):
    if valor is None:
        return padrao
    try:
        if pd.isna(valor):
            return padrao
    except Exception:
        pass
    texto = str(valor).strip()
    if not texto or texto.lower() in ["nan", "none", "null"]:
        return padrao
    return texto


def normalizar_df_infos(df):
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame(columns=COLUNAS_INFOS)

    df = df.copy()
    for col in COLUNAS_INFOS:
        if col not in df.columns:
            df[col] = ""

    return df[COLUNAS_INFOS].fillna("")


def valor_info(df, informacao, padrao="[campo a preencher]"):
    df = normalizar_df_infos(df)
    if df.empty:
        return padrao

    mask = df["Informação necessária"].astype(str).str.strip().str.lower() == informacao.strip().lower()
    if not mask.any():
        return padrao

    return texto_seguro(df.loc[mask, "Valor / informação levantada"].iloc[0], padrao)


def status_info(df, informacao):
    df = normalizar_df_infos(df)
    if df.empty:
        return ""
    mask = df["Informação necessária"].astype(str).str.strip().str.lower() == informacao.strip().lower()
    if not mask.any():
        return ""
    return texto_seguro(df.loc[mask, "Status"].iloc[0], "")


def resumo_status_df(df, nome="itens"):
    if not isinstance(df, pd.DataFrame) or df.empty or "Status" not in df.columns:
        return f"não há {nome} com status disponível."

    total = len(df)
    validado = int((df["Status"] == "Validado").sum())
    conferencia = int((df["Status"] == "Em conferência").sum())
    pendente = int((df["Status"] == "Pendente").sum())
    nao_aplica = int((df["Status"] == "Não se aplica").sum())

    return (
        f"{total} {nome}: {validado} validados, {conferencia} em conferência, "
        f"{pendente} pendentes e {nao_aplica} não aplicáveis"
    )


def extrair_delta_por_ciclo(resultado):
    df = resultado.get("df_delta_por_ciclo") if isinstance(resultado, dict) else None
    if not isinstance(df, pd.DataFrame) or df.empty:
        df = resultado.get("df_financeiro_por_ciclo") if isinstance(resultado, dict) else None
    if not isinstance(df, pd.DataFrame) or df.empty:
        return ""

    partes = []
    for _, row in df.head(8).iterrows():
        ciclo = texto_seguro(row.get("Ciclo", ""), "")
        situacao = texto_seguro(row.get("Situação", ""), "")
        pago = row.get("Valor pago efetivo", row.get("Valor pago", None))
        teorico = row.get("Valor teórico calculado", row.get("Valor devido", None))
        delta = row.get("Delta do ciclo", row.get("Valor represado", None))

        trecho = []
        if ciclo:
            trecho.append(str(ciclo))
        if situacao:
            trecho.append(f"situação {situacao}")
        if pago is not None:
            trecho.append(f"valor pago efetivo {moeda(pago)}")
        if teorico is not None:
            trecho.append(f"valor teórico calculado {moeda(teorico)}")
        if delta is not None:
            trecho.append(f"diferença/retroativo {moeda(delta)}")
        if trecho:
            partes.append(", ".join(trecho))

    if not partes:
        return ""

    return " Por ciclo, a apuração registrou: " + "; ".join(partes) + "."


def resumo_fontes(df_infos, resultado_vg, resultado_garantia, df_checklist, eventos_aditivos):
    linhas = []
    linhas.append({"Fonte": "Infos Prévias", "Status": "Disponível" if isinstance(df_infos, pd.DataFrame) and not df_infos.empty else "Não disponível"})
    linhas.append({"Fonte": "Valor Global", "Status": "Disponível" if isinstance(resultado_vg, dict) and bool(resultado_vg) else "Não disponível"})
    linhas.append({"Fonte": "Garantia", "Status": "Disponível" if isinstance(resultado_garantia, dict) and bool(resultado_garantia) else "Não disponível"})
    linhas.append({"Fonte": "Checklist Processual", "Status": "Disponível" if isinstance(df_checklist, pd.DataFrame) and not df_checklist.empty else "Não disponível"})
    linhas.append({"Fonte": "Avaliação de Aditivos", "Status": "Disponível" if isinstance(eventos_aditivos, pd.DataFrame) and not eventos_aditivos.empty else "Não disponível"})
    return pd.DataFrame(linhas)


def montar_saneador_integrado(df_infos, resultado_vg, resultado_garantia, df_checklist, eventos_aditivos, complemento_manual=""):
    df_infos = normalizar_df_infos(df_infos)

    data_final = valor_info(df_infos, "Data final vigente")
    data_proposta = valor_info(df_infos, "Data da proposta")
    indice_info = valor_info(df_infos, "Índice contratual")
    data_pedido = valor_info(df_infos, "Data do pedido da empresa")
    houve_reajuste = valor_info(df_infos, "Já houve reajuste anterior?")
    percentual_anterior = valor_info(df_infos, "Percentual já concedido")
    data_efeitos_anterior = valor_info(df_infos, "Data de efeitos financeiros anterior")
    valor_aditivo = valor_info(df_infos, "Valor do aditivo na assinatura")
    data_aditivo = valor_info(df_infos, "Data do aditivo")
    percentual_garantia_info = valor_info(df_infos, "Percentual de garantia previsto")
    valor_garantia_info = valor_info(df_infos, "Valor da garantia constituída")
    historico_endossos = valor_info(df_infos, "Histórico de endossos")
    ultimo_termo = valor_info(df_infos, "Último Termo de Apostila")
    oficio_pleito = valor_info(df_infos, "Ofício de Pleito da Empresa")

    tipo_analise = st.session_state.get("calculadora_tipo_analise", "[campo a preencher]")

    vg_disponivel = isinstance(resultado_vg, dict) and bool(resultado_vg)
    gar_disponivel = isinstance(resultado_garantia, dict) and bool(resultado_garantia)

    indice_resultado = texto_seguro(resultado_vg.get("indice") if vg_disponivel else None, indice_info)
    qtd_ciclos = texto_seguro(resultado_vg.get("quantidade_ciclos") if vg_disponivel else None)
    data_processamento_vg = texto_seguro(resultado_vg.get("data_processamento") if vg_disponivel else None)
    valor_original = moeda(resultado_vg.get("valor_original_contrato", 0)) if vg_disponivel else "[campo a preencher]"
    valor_formalizado = moeda(resultado_vg.get("valor_formalizado_anterior", 0)) if vg_disponivel else "[campo a preencher]"
    valor_pago = moeda(resultado_vg.get("valor_pago_efetivo", 0)) if vg_disponivel else "[campo a preencher]"
    valor_teorico = moeda(resultado_vg.get("valor_teorico_calculado", 0)) if vg_disponivel else "[campo a preencher]"
    valor_represado = moeda(resultado_vg.get("valor_represado_a_pagar", 0)) if vg_disponivel else "[campo a preencher]"
    remanescente = moeda(resultado_vg.get("remanescente_reajustado", 0)) if vg_disponivel else "[campo a preencher]"
    valor_executado = moeda(resultado_vg.get("valor_executado_atualizado", 0)) if vg_disponivel else "[campo a preencher]"
    valor_total = moeda(resultado_vg.get("valor_atualizado_contrato", 0)) if vg_disponivel else "[campo a preencher]"
    ciclo_rem = texto_seguro(resultado_vg.get("ciclo_ultimo_remanescente") if vg_disponivel else None)
    variacao = pct(resultado_vg.get("variacao_acumulada", 0)) if vg_disponivel else "[campo a preencher]"
    qtd_aditivos = texto_seguro(resultado_vg.get("quantidade_aditivos_total") if vg_disponivel else None)
    aditivos_informativos = moeda(resultado_vg.get("total_aditivos_informativos", 0)) if vg_disponivel else "[campo a preencher]"
    delta_ciclos = extrair_delta_por_ciclo(resultado_vg if vg_disponivel else {})

    garantia_base = moeda(resultado_garantia.get("valor_total_atualizado_contrato", resultado_garantia.get("valor_atualizado_base", 0))) if gar_disponivel else "[campo a preencher]"
    garantia_pct = pct(resultado_garantia.get("percentual_garantia", 0)) if gar_disponivel else percentual_garantia_info
    garantia_exigida = moeda(resultado_garantia.get("garantia_exigida", 0)) if gar_disponivel else "[campo a preencher]"
    garantia_constituida = moeda(resultado_garantia.get("garantia_constituida", 0)) if gar_disponivel else valor_garantia_info
    endosso = moeda(resultado_garantia.get("endosso_necessario", 0)) if gar_disponivel else "[campo a preencher]"
    validade_minima = texto_seguro(resultado_garantia.get("data_validade_minima") if gar_disponivel else None)

    checklist_resumo = resumo_status_df(df_checklist, "itens do checklist")
    infos_resumo = resumo_status_df(df_infos, "itens de informações prévias")

    aditivos_resumo = ""
    if isinstance(eventos_aditivos, pd.DataFrame) and not eventos_aditivos.empty:
        aditivos_resumo = (
            f" A página de Avaliação de Aditivos registra {len(eventos_aditivos)} evento(s) informado(s), "
            "os quais devem ser utilizados apenas para governança do limite de acréscimos e supressões, sem alterar automaticamente o Valor Global."
        )

    paragrafos = [
        "Despacho Saneador para Formalização de Termo de Apostila",
        (
            "Trata-se de saneamento da instrução processual destinada à formalização de Termo de Apostila para registro "
            "do reajuste contratual. Este documento consolida a sequência de atos, informações e resultados apurados "
            "antes da assinatura, com a finalidade de demonstrar a regularidade mínima da instrução."
        ),
        (
            f"A contratada apresentou pleito de reajuste por meio de {oficio_pleito}, com data de pedido registrada em "
            f"{data_pedido}. Para verificação da anualidade, foi considerada a data da proposta em {data_proposta}, bem como "
            f"o índice contratual {indice_resultado}, extraído da documentação contratual e/ou da análise realizada no sistema."
        ),
        (
            f"Quanto à situação contratual prévia, foi informada vigência final em {data_final}. O levantamento também registrou "
            f"que {houve_reajuste} quanto à existência de reajuste anterior, com percentual já concedido de {percentual_anterior}, "
            f"efeitos financeiros anteriores em {data_efeitos_anterior} e último Termo de Apostila identificado como {ultimo_termo}."
        ),
        (
            f"A Calculadora de Reajustes foi utilizada no modo {tipo_analise}. A análise processada em {data_processamento_vg} "
            f"considerou {qtd_ciclos} ciclo(s), com variação acumulada de {variacao}. O valor original do contrato informado foi "
            f"{valor_original}, e o valor formalizado antes desta análise foi {valor_formalizado}."
        ),
        (
            f"A apuração financeira consolidada indicou valor pago efetivo de {valor_pago} e valor teórico calculado de "
            f"{valor_teorico}, resultando em valor represado a pagar de {valor_represado}.{delta_ciclos}"
        ),
        (
            f"Para fins de consolidação contratual, o Valor Total Atualizado do Contrato foi apurado em {valor_total}, composto "
            f"pela execução atualizada por ciclo, no montante de {valor_executado}, somada ao saldo remanescente atualizado de "
            f"{remanescente}, correspondente ao ciclo/referência {ciclo_rem}. Aditivos e supressões não são somados como parcela "
            "autônoma quando seus efeitos já estiverem incorporados à execução, ao estoque ou ao saldo remanescente, evitando dupla contagem."
        ),
        (
            f"No tocante aos aditivos e supressões, as informações prévias registraram valor de aditivo na assinatura de "
            f"{valor_aditivo}, com data de referência {data_aditivo}. No Valor Global constam {qtd_aditivos} aditivo(s) ou "
            f"supressão(ões) registrados para fins informativos, com valor informativo consolidado de {aditivos_informativos}."
            f"{aditivos_resumo}"
        ),
        (
            f"Quanto à garantia contratual, o levantamento inicial indicou percentual previsto de {percentual_garantia_info}, "
            f"valor constituído de {valor_garantia_info} e histórico de endossos registrado como {historico_endossos}. A análise "
            f"de garantia calculada no sistema considerou base de {garantia_base}, percentual de {garantia_pct}, garantia exigida "
            f"de {garantia_exigida}, garantia constituída de {garantia_constituida} e endosso necessário de {endosso}. A validade mínima "
            f"indicada para a garantia é {validade_minima}."
        ),
        (
            f"Quanto à completude da instrução, as Infos Prévias registram {infos_resumo}. O Checklist Processual registra "
            f"{checklist_resumo}. Itens pendentes ou em conferência devem ser saneados antes da assinatura, especialmente quando "
            "relacionados à memória de cálculo, adequação orçamentária, garantia, certidões de regularidade, validade das certidões "
            "na véspera da assinatura e conformidade da minuta."
        ),
    ]

    complemento_manual = (complemento_manual or "").strip()
    if complemento_manual:
        paragrafos.append(complemento_manual)

    paragrafos.append(
        "Diante do exposto, estando conferidos os elementos documentais, financeiros e formais necessários, e inexistindo "
        "pendência crítica impeditiva, a instrução poderá prosseguir para formalização do Termo de Apostila, observadas as "
        "alçadas competentes e os procedimentos internos aplicáveis."
    )

    return "\n\n".join(paragrafos)


def gerar_docx_texto(texto):
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.text import WD_COLOR_INDEX
    from docx.shared import Inches, Pt

    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)

    styles = document.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10.5)

    def add_texto_com_placeholder(paragraph, conteudo):
        partes = re.split(r"(\[campo a preencher\])", conteudo)
        for parte in partes:
            if not parte:
                continue
            run = paragraph.add_run(parte)
            if parte == "[campo a preencher]":
                run.bold = True
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    paragrafos = [p.strip() for p in str(texto).split("\n\n") if p.strip()]
    if paragrafos:
        title = document.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run(paragrafos[0])
        run.bold = True
        run.font.size = Pt(14)

        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(f"Gerado em: {datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')}").italic = True
        corpo = paragrafos[1:]
    else:
        corpo = []

    for par in corpo:
        p = document.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        add_texto_com_placeholder(p, par)

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def gerar_txt(texto):
    return str(texto).encode("utf-8")


render_marca_topo()

st.title("Saneador")
st.caption("Minuta narrativa integrada que consolida Infos Prévias, análises e resultados antes da assinatura do Termo de Apostila.")

df_infos = normalizar_df_infos(st.session_state.get("infos_previas_df"))
resultado_vg = st.session_state.get("resultado_valor_global")
resultado_garantia = st.session_state.get("resultado_garantia")
df_checklist = st.session_state.get("checklist_processual")
eventos_aditivos = st.session_state.get("avaliacao_aditivos_eventos")

fontes = resumo_fontes(df_infos, resultado_vg, resultado_garantia, df_checklist, eventos_aditivos)

st.info(
    "O Saneador é gerado automaticamente a partir dos dados disponíveis no sistema. "
    "Ele conta a história processual: pleito, premissas, cálculo, Valor Global, garantia, aditivos e conferência da instrução."
)

st.dataframe(fontes, use_container_width=True, hide_index=True)

with st.expander("Complemento manual opcional", expanded=False):
    complemento_manual = st.text_area(
        "Usar apenas para ressalvas ou informações que ainda não estejam nos módulos da ferramenta.",
        value=st.session_state.get("saneador_complemento_manual", ""),
        height=100,
    )
    st.session_state["saneador_complemento_manual"] = complemento_manual

texto_gerado = montar_saneador_integrado(
    df_infos,
    resultado_vg if isinstance(resultado_vg, dict) else {},
    resultado_garantia if isinstance(resultado_garantia, dict) else {},
    df_checklist if isinstance(df_checklist, pd.DataFrame) else pd.DataFrame(),
    eventos_aditivos if isinstance(eventos_aditivos, pd.DataFrame) else pd.DataFrame(),
    complemento_manual=complemento_manual,
)

if st.button("Gerar/atualizar texto integrado", type="primary"):
    st.session_state["saneador_texto"] = texto_gerado
    st.rerun()

if "saneador_texto" not in st.session_state:
    st.session_state["saneador_texto"] = texto_gerado

texto_editado = st.text_area(
    "Texto do Saneador",
    value=st.session_state["saneador_texto"],
    height=680,
    help="Revise o texto antes de baixar. Campos sem informação aparecerão como [campo a preencher].",
)

st.session_state["saneador_texto"] = texto_editado

docx_bytes = gerar_docx_texto(texto_editado)
txt_bytes = gerar_txt(texto_editado)

st.session_state["arquivo_saneador_docx"] = docx_bytes
st.session_state["arquivo_saneador_txt"] = txt_bytes

col1, col2 = st.columns(2)
with col1:
    st.download_button(
        "Baixar Saneador em DOCX",
        data=docx_bytes,
        file_name="Saneador_Instrucao_Processual.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )

with col2:
    st.download_button(
        "Baixar Saneador em TXT",
        data=txt_bytes,
        file_name="Saneador_Instrucao_Processual.txt",
        mime="text/plain",
        use_container_width=True,
    )

with st.expander("Base de Infos Prévias utilizada", expanded=False):
    st.dataframe(df_infos, use_container_width=True, hide_index=True)
