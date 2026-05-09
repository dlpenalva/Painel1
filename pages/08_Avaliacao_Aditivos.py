import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

st.set_page_config(page_title="TLB · cl8us - Avaliação de Aditivos", layout="wide")


from _ui_utils import render_marca_topo

def moeda(valor):
    try:
        valor = round(float(valor or 0), 2)
    except Exception:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def numero_br(valor):
    """Aceita números digitados no padrão brasileiro, como 1.234,56, ou no padrão técnico, como 1234.56."""
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        try:
            if pd.isna(valor):
                return 0.0
        except Exception:
            pass
        return float(valor)
    texto = str(valor).strip()
    if not texto or texto.lower() in ["nan", "none", "null"]:
        return 0.0
    texto = texto.replace("R$", "").replace("%", "").replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except Exception:
        return 0.0


def pct(valor):
    try:
        valor = float(valor or 0)
    except Exception:
        valor = 0.0
    return f"{valor * 100:.2f}%".replace(".", ",")


def data_br(valor):
    try:
        dt = pd.to_datetime(valor, errors="coerce")
        if pd.isna(dt):
            return ""
        return dt.strftime("%d/%m/%Y")
    except Exception:
        return ""


def normalizar_percentual(valor):
    """Interpreta o campo como percentual digitado, não como fração.
    Exemplos:
    - 0,99 = 0,99%
    - 10 = 10%
    - 10% = 10%
    """
    texto = "" if valor is None else str(valor).strip()
    if not texto or texto.lower() in ["nan", "none", "null"]:
        return 0.0
    v = numero_br(texto)
    return v / 100.0


def tabela_inicial():
    df = pd.DataFrame(
        [
            {
                "Seq.": 1,
                "Data": pd.NaT,
                "Tipo do evento": "",
                "Descrição / instrumento": "",
                "% sobre base legal": "",
                "Valor informado (R$)": "",
                "Observação": "",
            }
        ]
    )
    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    return df


def normalizar_linhas_eventos(df):
    df = df.copy()
    if "Seq." not in df.columns:
        df.insert(0, "Seq.", range(1, len(df) + 1))
    if "Data" in df.columns:
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")

    colunas_texto = [
        "Tipo do evento",
        "Descrição / instrumento",
        "% sobre base legal",
        "Valor informado (R$)",
        "Observação",
    ]
    for col in colunas_texto:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].apply(lambda x: "" if pd.isna(x) or str(x).strip().lower() in ["none", "nan", "null"] else str(x).strip())

    df["Seq."] = range(1, len(df) + 1)
    return df


def calcular_eventos(df, base_legal, limite_acrescimos):
    linhas = []
    acumulado_acrescimos = 0.0
    acumulado_supressoes = 0.0

    for _, row in df.iterrows():
        tipo = str(row.get("Tipo do evento", "") or "").strip()
        descricao = str(row.get("Descrição / instrumento", "") or "").strip()
        percentual = normalizar_percentual(row.get("% sobre base legal"))
        valor_informado = row.get("Valor informado (R$)")

        valor_num = numero_br(valor_informado)

        if not tipo and not descricao and percentual == 0 and valor_num == 0:
            continue

        valor_evento = valor_num if abs(valor_num) > 0 else base_legal * percentual
        tipo_norm = tipo.lower()

        conta_acrescimos = 0.0
        conta_supressoes = 0.0
        tratamento = "Informativo"

        if "acrésc" in tipo_norm or "acresc" in tipo_norm:
            conta_acrescimos = abs(valor_evento)
            tratamento = "Conta como acréscimo"
        elif "supress" in tipo_norm:
            conta_supressoes = abs(valor_evento)
            tratamento = "Conta como supressão"
        elif "restabelecimento" in tipo_norm:
            tratamento = "Revisão manual: restabelecimento do mesmo item"
        elif "reajuste" in tipo_norm or "revisão" in tipo_norm or "revisao" in tipo_norm or "reequil" in tipo_norm:
            tratamento = "Ajuste monetário/formal — não consome limite quantitativo"
        elif "informativo" in tipo_norm:
            tratamento = "Informativo"

        acumulado_acrescimos += conta_acrescimos
        acumulado_supressoes += conta_supressoes
        saldo_acrescimo = limite_acrescimos - acumulado_acrescimos
        percentual_acrescimos = acumulado_acrescimos / base_legal if base_legal else 0.0
        status = "OK" if saldo_acrescimo >= -0.004 else "EXTRAPOLA"

        linhas.append({
            "Seq.": row.get("Seq.", ""),
            "Data": row.get("Data", ""),
            "Tipo do evento": tipo,
            "Descrição / instrumento": descricao,
            "% sobre base legal": percentual,
            "Valor informado (R$)": valor_num if abs(valor_num) > 0 else "",
            "Valor do evento (R$)": round(valor_evento, 2),
            "Tratamento": tratamento,
            "Conta acréscimos (R$)": round(conta_acrescimos, 2),
            "Conta supressões (R$)": round(conta_supressoes, 2),
            "Acréscimos acumulados (R$)": round(acumulado_acrescimos, 2),
            "Supressões acumuladas (R$)": round(acumulado_supressoes, 2),
            "Saldo para novos acréscimos (R$)": round(saldo_acrescimo, 2),
            "% acréscimos acumulados": percentual_acrescimos,
            "Status": status,
            "Observação": row.get("Observação", ""),
        })

    return pd.DataFrame(linhas)


def gerar_excel(resultado, eventos, parametros):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        resultado.to_excel(writer, index=False, sheet_name="RESUMO")
        eventos.to_excel(writer, index=False, sheet_name="EVENTOS_ADITIVOS")
        parametros.to_excel(writer, index=False, sheet_name="PARAMETROS")

        workbook = writer.book
        fmt_header = workbook.add_format({
            "bold": True, "bg_color": "#1F4E79", "font_color": "#FFFFFF",
            "border": 1, "align": "center", "valign": "vcenter"
        })
        fmt_wrap = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})
        fmt_money = workbook.add_format({"num_format": 'R$ #,##0.00', "border": 1})
        fmt_pct = workbook.add_format({"num_format": '0.00%', "border": 1})
        fmt_ok = workbook.add_format({"bg_color": "#E2F0D9", "border": 1})
        fmt_bad = workbook.add_format({"bg_color": "#FCE4D6", "border": 1})

        for sheet_name in ["RESUMO", "EVENTOS_ADITIVOS", "PARAMETROS"]:
            ws = writer.sheets[sheet_name]
            df = {"RESUMO": resultado, "EVENTOS_ADITIVOS": eventos, "PARAMETROS": parametros}[sheet_name]
            for col_num, col_name in enumerate(df.columns):
                ws.write(0, col_num, col_name, fmt_header)
                ws.set_column(col_num, col_num, 24 if col_num != 3 else 38, fmt_wrap)
            ws.freeze_panes(1, 0)

        ws_e = writer.sheets["EVENTOS_ADITIVOS"]
        colunas_moeda = [c for c in eventos.columns if "(R$)" in c]
        colunas_pct = [c for c in eventos.columns if "%" in c]
        for c in colunas_moeda:
            idx = eventos.columns.get_loc(c)
            ws_e.set_column(idx, idx, 22, fmt_money)
        for c in colunas_pct:
            idx = eventos.columns.get_loc(c)
            ws_e.set_column(idx, idx, 18, fmt_pct)
        if "Status" in eventos.columns:
            idx = eventos.columns.get_loc("Status")
            ws_e.conditional_format(1, idx, max(len(eventos), 1), idx, {
                "type": "text", "criteria": "containing", "value": "OK", "format": fmt_ok
            })
            ws_e.conditional_format(1, idx, max(len(eventos), 1), idx, {
                "type": "text", "criteria": "containing", "value": "EXTRAPOLA", "format": fmt_bad
            })

        ws_r = writer.sheets["RESUMO"]
        ws_r.set_column(1, 1, 26, fmt_money)

    buffer.seek(0)
    return buffer.getvalue()


render_marca_topo()

st.title("Avaliação de Aditivos")
st.caption("Módulo estanque para controle de acréscimos, supressões e limite de 25%. Não altera o Valor Global.")

st.info(
    "Este módulo avalia a aderência de aditivos e supressões ao limite percentual aplicável. "
    "Ele é independente do Valor Global e não recalcula execução, saldo remanescente ou reajustes."
)

with st.expander("Parâmetros do contrato", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        valor_original = st.number_input("Valor original do contrato (R$)", min_value=0.0, value=0.0, step=1000.0, format="%.2f")
        reajustes_pct_input = st.number_input("Reajustes/revisões/reequilíbrios acumulados sobre o original (%)", value=0.0, step=0.01, format="%.2f")
    with col2:
        ajustes_monetarios = st.number_input("Ajustes monetários adicionais sobre o original (R$)", value=0.0, step=1000.0, format="%.2f")
        limite_percentual_input = st.number_input("Limite de acréscimos aplicável (%)", min_value=0.0, value=25.0, step=0.5, format="%.2f")

reajustes_pct = reajustes_pct_input / 100.0
limite_percentual = limite_percentual_input / 100.0
base_legal = round(valor_original * (1 + reajustes_pct) + ajustes_monetarios, 2)
limite_acrescimos = round(base_legal * limite_percentual, 2)
referencia_supressoes = round(base_legal * 0.25, 2)
limite_reforma = round(base_legal * 0.50, 2)

st.subheader("Eventos de alteração contratual")
st.caption("Nos campos de valor, digite no padrão brasileiro, por exemplo: 1.234,56. No percentual, informe como percentual: 0,99 = 0,99%; 10 = 10%.")

if "avaliacao_aditivos_eventos" not in st.session_state:
    st.session_state["avaliacao_aditivos_eventos"] = tabela_inicial()
else:
    st.session_state["avaliacao_aditivos_eventos"] = normalizar_linhas_eventos(st.session_state["avaliacao_aditivos_eventos"])

eventos_input = st.data_editor(
    st.session_state["avaliacao_aditivos_eventos"],
    use_container_width=True,
    hide_index=True,
    column_config={
        "Seq.": st.column_config.NumberColumn("Seq.", disabled=True),
        "Data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
        "Tipo do evento": st.column_config.SelectboxColumn(
            "Tipo do evento",
            options=["", "Acréscimo", "Supressão", "Restabelecimento do mesmo item", "Reajuste/revisão/reequilíbrio", "Informativo", "Outro"],
        ),
        "% sobre base legal": st.column_config.TextColumn("% sobre base legal", help="Informe como percentual: 0,99 = 0,99%; 10 = 10%."),
        "Valor informado (R$)": st.column_config.TextColumn("Valor informado (R$)", help="Aceita 1234,56, 1.234,56 ou 1234.56."),
        "Observação": st.column_config.TextColumn("Observação"),
    },
    num_rows="dynamic",
    key="avaliacao_aditivos_editor",
)
eventos_input = normalizar_linhas_eventos(eventos_input)
st.session_state["avaliacao_aditivos_eventos"] = eventos_input.copy()

eventos_calculados = calcular_eventos(eventos_input, base_legal, limite_acrescimos)

acrescimos_acumulados = float(eventos_calculados["Conta acréscimos (R$)"].sum()) if not eventos_calculados.empty else 0.0
supressoes_acumuladas = float(eventos_calculados["Conta supressões (R$)"].sum()) if not eventos_calculados.empty else 0.0
saldo = limite_acrescimos - acrescimos_acumulados
pct_usado = acrescimos_acumulados / base_legal if base_legal else 0.0
status_geral = "OK" if saldo >= -0.004 else "EXTRAPOLA"

st.subheader("Resultado executivo")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Base legal atualizada", moeda(base_legal))
c2.metric("Limite de acréscimos", moeda(limite_acrescimos))
c3.metric("Acréscimos acumulados", moeda(acrescimos_acumulados))
c4.metric("Saldo disponível", moeda(saldo))

c5, c6, c7 = st.columns(3)
c5.metric("% utilizado", pct(pct_usado))
c6.metric("Supressões acumuladas", moeda(supressoes_acumuladas))

if status_geral == "EXTRAPOLA":
    status_bg = "#FCE4D6"
    status_fg = "#7F1D1D"
    status_border = "#F4B183"
    status_msg = "EXTRAPOLA"
else:
    status_bg = "#DBEAFE"
    status_fg = "#1E3A8A"
    status_border = "#93C5FD"
    status_msg = "OK"

c7.markdown(
    f"""
    <div style="border:1px solid {status_border}; background:{status_bg}; color:{status_fg};
                border-radius:12px; padding:0.70rem 0.85rem; min-height:74px;">
        <div style="font-size:0.82rem; font-weight:600; margin-bottom:0.20rem;">Status</div>
        <div style="font-size:1.45rem; font-weight:800;">{status_msg}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if status_geral == "EXTRAPOLA":
    st.error("Os acréscimos acumulados ultrapassam o limite informado. Revisar a instrução antes de prosseguir.")
else:
    st.success("A avaliação não indica extrapolação do limite de acréscimos informado.")

with st.expander("Ver eventos calculados", expanded=True):
    if eventos_calculados.empty:
        st.info("Preencha ao menos um evento para visualizar a memória de cálculo.")
    else:
        visual = eventos_calculados.copy()
        if "Data" in visual.columns:
            visual["Data"] = visual["Data"].apply(data_br)
        for col in [c for c in visual.columns if "(R$)" in c]:
            visual[col] = visual[col].apply(moeda)
        for col in [c for c in visual.columns if "%" in c]:
            visual[col] = visual[col].apply(pct)
        st.dataframe(visual, use_container_width=True, hide_index=True)

parametros = pd.DataFrame([
    {"Parâmetro": "Data de geração", "Valor": datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")},
    {"Parâmetro": "Valor original do contrato", "Valor": valor_original},
    {"Parâmetro": "Reajustes/revisões/reequilíbrios acumulados", "Valor": reajustes_pct},
    {"Parâmetro": "Ajustes monetários adicionais", "Valor": ajustes_monetarios},
    {"Parâmetro": "Base legal atualizada", "Valor": base_legal},
    {"Parâmetro": "Limite de acréscimos", "Valor": limite_acrescimos},
    {"Parâmetro": "Referência de supressões 25%", "Valor": referencia_supressoes},
    {"Parâmetro": "Referência especial de reforma 50%", "Valor": limite_reforma},
])

resumo = pd.DataFrame([
    {"Indicador": "Base legal atualizada", "Valor": base_legal},
    {"Indicador": "Limite de acréscimos", "Valor": limite_acrescimos},
    {"Indicador": "Acréscimos acumulados", "Valor": acrescimos_acumulados},
    {"Indicador": "Supressões acumuladas", "Valor": supressoes_acumuladas},
    {"Indicador": "Saldo para novos acréscimos", "Valor": saldo},
    {"Indicador": "% acréscimos acumulados", "Valor": pct_usado},
    {"Indicador": "Status geral", "Valor": status_geral},
])

excel_bytes = gerar_excel(resumo, eventos_calculados, parametros)
st.session_state["arquivo_avaliacao_aditivos_xlsx"] = excel_bytes
st.download_button(
    "Baixar Avaliação de Aditivos em XLSX",
    data=excel_bytes,
    file_name="Avaliacao_Aditivos_Limite_25.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.caption(
    "Observação: o módulo é apoio de governança. A aplicação do limite legal/regulamentar deve ser validada conforme o caso concreto, "
    "o contrato, a natureza do objeto e a manifestação jurídica aplicável."
)
