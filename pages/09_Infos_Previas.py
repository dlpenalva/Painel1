from datetime import datetime
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

st.set_page_config(page_icon="assets/cl8us_favicon_512.png", page_title="TLB · cl8us - Infos Prévias", layout="wide")

from _ui_utils import render_marca_topo, render_aviso_privacidade


STATUS_OPCOES = ["Pendente", "Em conferência", "Validado", "Não se aplica"]

ITENS_INFOS_PREVIAS = [
    ("Vigência", "Data final vigente", "Contrato/aditivo"),
    ("Reajuste", "Data da proposta", "Proposta comercial"),
    ("Reajuste", "Índice contratual", "Contrato"),
    ("Reajuste", "Data do pedido da empresa", "Ofício/e-mail da contratada"),
    ("Reajuste", "Já houve reajuste anterior?", "Termo de apostila anterior"),
    ("Reajuste", "Percentual já concedido", "Apostila/relatório anterior"),
    ("Reajuste", "Data de efeitos financeiros anterior", "Apostila/relatório anterior"),
    ("Aditivos", "Valor do aditivo na assinatura", "Termo aditivo"),
    ("Aditivos", "Data do aditivo", "Termo aditivo"),
    ("Garantia", "Percentual de garantia previsto", "Contrato/edital"),
    ("Garantia", "Valor da garantia constituída", "Apólice/endosso/caução"),
    ("Garantia", "Histórico de endossos", "Apólices/endossos"),
    ("Documentos-chave", "Último Termo de Apostila", "Termo de apostila anterior/último vigente"),
    ("Documentos-chave", "Ofício de Pleito da Empresa", "Ofício/e-mail da contratada"),
]

DATE_FIELDS = {
    "Data final vigente", "Data da proposta", "Data do pedido da empresa",
    "Data de efeitos financeiros anterior", "Data do aditivo",
}
MONEY_FIELDS = {"Valor do aditivo na assinatura", "Valor da garantia constituída"}
PERCENT_FIELDS = {"Percentual já concedido", "Percentual de garantia previsto"}


def dataframe_inicial():
    return pd.DataFrame([
        {
            "Grupo": grupo, "Informação necessária": info,
            "Documento/fonte mínima": fonte,
            "Valor / informação levantada": "",
            "Link SIGA (opcional)": "",
            "Status": "Pendente",
            "Observação": "",
        }
        for grupo, info, fonte in ITENS_INFOS_PREVIAS
    ])


def _texto(valor):
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except Exception:
        pass
    texto = str(valor).strip()
    if texto.lower() in ["nan", "none", "null"]:
        return ""
    return texto


def _parse_numero_br(valor):
    texto = _texto(valor)
    if not texto:
        return None
    texto = texto.replace("R$", "").replace("%", "").replace("\xa0", "").replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    else:
        if texto.count(".") > 1:
            partes = texto.split(".")
            texto = "".join(partes[:-1]) + "." + partes[-1]
    try:
        return float(texto)
    except Exception:
        return None


def _formatar_moeda_br(valor):
    numero = _parse_numero_br(valor)
    if numero is None:
        return _texto(valor)
    return f"R$ {numero:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _formatar_percentual_br(valor):
    numero = _parse_numero_br(valor)
    if numero is None:
        return _texto(valor)
    return f"{numero:.2f}%".replace(".", ",")


def _formatar_data_br(valor):
    texto = _texto(valor)
    if not texto:
        return ""
    try:
        data = pd.to_datetime(texto, dayfirst=True, errors="coerce")
        if pd.isna(data):
            return texto
        return data.strftime("%d/%m/%Y")
    except Exception:
        return texto


def _normalizar_sim_nao(valor):
    texto = _texto(valor).lower()
    if texto in ["sim", "s", "yes", "y"]:
        return "Sim"
    if texto in ["não", "nao", "n", "no"]:
        return "Não"
    return _texto(valor)


def normalizar_valor_por_informacao(informacao, valor):
    informacao = _texto(informacao)
    if informacao in DATE_FIELDS:
        return _formatar_data_br(valor)
    if informacao in MONEY_FIELDS:
        return _formatar_moeda_br(valor)
    if informacao in PERCENT_FIELDS:
        return _formatar_percentual_br(valor)
    if informacao == "Já houve reajuste anterior?":
        return _normalizar_sim_nao(valor)
    if informacao == "Índice contratual":
        return _texto(valor).upper() if _texto(valor).lower() in ["ist", "ipca", "igp-m", "igpm"] else _texto(valor)
    return _texto(valor)


def normalizar_dataframe_infos(df):
    if not isinstance(df, pd.DataFrame) or df.empty:
        return dataframe_inicial()
    df = df.copy()
    base_cols = list(dataframe_inicial().columns)
    for col in base_cols:
        if col not in df.columns:
            df[col] = ""
    df = df[base_cols].fillna("")
    for idx, row in df.iterrows():
        info = row.get("Informação necessária", "")
        valor = row.get("Valor / informação levantada", "")
        df.at[idx, "Valor / informação levantada"] = normalizar_valor_por_informacao(info, valor)
    return df


def gerar_excel_infos(df):
    df = normalizar_dataframe_infos(df)
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="INFOS_PREVIAS")
        workbook = writer.book
        ws = writer.sheets["INFOS_PREVIAS"]
        fmt_header = workbook.add_format({
            "bold": True, "bg_color": "#1F4E79", "font_color": "#FFFFFF",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        fmt_wrap = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})
        fmt_pendente = workbook.add_format({"bg_color": "#FCE4D6", "border": 1})
        fmt_conferencia = workbook.add_format({"bg_color": "#FFF2CC", "border": 1})
        fmt_validado = workbook.add_format({"bg_color": "#E2F0D9", "border": 1})
        fmt_na = workbook.add_format({"bg_color": "#E7E6E6", "border": 1})
        widths = [18, 34, 30, 36, 36, 18, 34]
        for col_num, col_name in enumerate(df.columns):
            ws.write(0, col_num, col_name, fmt_header)
            ws.set_column(col_num, col_num, widths[col_num] if col_num < len(widths) else 24, fmt_wrap)
        ws.freeze_panes(1, 0)
        if "Status" in df.columns:
            idx = df.columns.get_loc("Status")
            last = max(len(df), 1)
            ws.conditional_format(1, idx, last, idx, {"type": "text", "criteria": "containing", "value": "Pendente", "format": fmt_pendente})
            ws.conditional_format(1, idx, last, idx, {"type": "text", "criteria": "containing", "value": "Em conferência", "format": fmt_conferencia})
            ws.conditional_format(1, idx, last, idx, {"type": "text", "criteria": "containing", "value": "Validado", "format": fmt_validado})
            ws.conditional_format(1, idx, last, idx, {"type": "text", "criteria": "containing", "value": "Não se aplica", "format": fmt_na})
        resumo = pd.DataFrame([
            {"Indicador": "Data de geração", "Valor": datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")},
            {"Indicador": "Total de itens", "Valor": len(df)},
            {"Indicador": "Pendentes", "Valor": int((df["Status"] == "Pendente").sum()) if "Status" in df.columns else ""},
            {"Indicador": "Validados", "Valor": int((df["Status"] == "Validado").sum()) if "Status" in df.columns else ""},
        ])
        resumo.to_excel(writer, index=False, sheet_name="RESUMO")
        ws2 = writer.sheets["RESUMO"]
        for col_num, col_name in enumerate(resumo.columns):
            ws2.write(0, col_num, col_name, fmt_header)
            ws2.set_column(col_num, col_num, 28, fmt_wrap)
    buffer.seek(0)
    return buffer.getvalue()


def importar_infos(upload):
    try:
        df = pd.read_excel(upload, sheet_name="INFOS_PREVIAS")
    except Exception:
        df = pd.read_excel(upload)
    base_cols = list(dataframe_inicial().columns)
    for col in base_cols:
        if col not in df.columns:
            df[col] = ""
    return normalizar_dataframe_infos(df[base_cols].fillna(""))


render_marca_topo()
st.title("Infos Prévias")
st.caption("Levantamento mínimo antes de iniciar a análise. A finalidade é reduzir retrabalho e orientar a instrução processual.")
render_aviso_privacidade(tem_upload=True, tem_download=True)

st.info("Preencha a tabela abaixo. Ao normalizar, o sistema padroniza datas, moedas, percentuais e respostas Sim/Não.")

if "infos_previas_df" not in st.session_state or not isinstance(st.session_state.get("infos_previas_df"), pd.DataFrame):
    st.session_state["infos_previas_df"] = dataframe_inicial()

with st.expander("Restaurar a partir de XLSX", expanded=False):
    arquivo = st.file_uploader("Enviar XLSX de Infos Prévias", type=["xlsx"], key="upload_infos_previas")
    if arquivo and st.button("Importar Infos Prévias"):
        st.session_state["infos_previas_df"] = importar_infos(arquivo)
        st.success("Infos Prévias importadas.")
        st.rerun()

col1, col2, col3 = st.columns([1, 1.35, 3])
with col1:
    if st.button("Restaurar modelo padrão", type="secondary"):
        st.session_state["infos_previas_df"] = dataframe_inicial()
        st.rerun()
with col2:
    if st.button("Normalizar dados", type="primary"):
        st.session_state["infos_previas_df"] = normalizar_dataframe_infos(st.session_state.get("infos_previas_df"))
        st.success("Dados normalizados.")
        st.rerun()
with col3:
    st.caption("Exemplos aceitos: 10/05/2026; 1234567,89; 1.234.567,89; 4,88; 5; sim/não.")

df_atual = st.session_state["infos_previas_df"].copy()

edited = st.data_editor(
    df_atual, use_container_width=True, hide_index=True, num_rows="dynamic",
    column_config={
        "Grupo": st.column_config.TextColumn("Grupo"),
        "Informação necessária": st.column_config.TextColumn("Informação necessária"),
        "Documento/fonte mínima": st.column_config.TextColumn("Documento/fonte mínima"),
        "Valor / informação levantada": st.column_config.TextColumn("Valor / informação levantada"),
        "Link SIGA (opcional)": st.column_config.LinkColumn("Link SIGA (opcional)", validate=r"^https?://.*|^$"),
        "Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPCOES, required=True),
        "Observação": st.column_config.TextColumn("Observação"),
    },
    key="infos_previas_editor",
)

st.session_state["infos_previas_df"] = edited.copy()

total = len(edited)
pendentes = int((edited["Status"] == "Pendente").sum()) if "Status" in edited.columns else 0
validado = int((edited["Status"] == "Validado").sum()) if "Status" in edited.columns else 0
conferencia = int((edited["Status"] == "Em conferência").sum()) if "Status" in edited.columns else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("Itens", total)
c2.metric("Validados", validado)
c3.metric("Pendentes", pendentes)
c4.metric("Em conferência", conferencia)

excel = gerar_excel_infos(edited)
st.session_state["arquivo_infos_previas_xlsx"] = excel

st.download_button(
    "Baixar Infos Prévias em XLSX", data=excel,
    file_name="Infos_Previas_Instrucao_Processual.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
