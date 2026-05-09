import html
import pandas as pd
import streamlit as st
from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

st.set_page_config(page_title="TLB · cl8us - Checklist Processual", layout="wide")


def render_marca_topo():
    """Identidade visual própria do sistema, sem uso de logomarca institucional."""
    st.markdown(
        """
        <style>
        .tlb-cl8us-brand {
            display: inline-flex;
            flex-direction: column;
            gap: 1px;
            margin: 0 0 0.70rem 0;
            padding: 0;
        }
        .tlb-cl8us-brand-main {
            display: flex;
            align-items: baseline;
            gap: 0.45rem;
            line-height: 1.05;
            letter-spacing: -0.02em;
        }
        .tlb-cl8us-tlb {
            color: #123B63;
            font-size: 1.38rem;
            font-weight: 750;
            font-family: "Inter", "Segoe UI", Arial, sans-serif;
        }
        .tlb-cl8us-dot {
            color: #C0842B;
            font-size: 1.18rem;
            font-weight: 700;
        }
        .tlb-cl8us-name {
            color: #0F172A;
            font-size: 1.42rem;
            font-weight: 800;
            font-family: "Consolas", "SFMono-Regular", "Cascadia Mono", "Courier New", monospace;
            letter-spacing: -0.04em;
        }
        .tlb-cl8us-subtitle {
            color: #64748B;
            font-size: 0.74rem;
            font-weight: 500;
            font-family: "Inter", "Segoe UI", Arial, sans-serif;
            margin-top: 0.12rem;
            letter-spacing: 0.01em;
        }
        </style>
        <div class="tlb-cl8us-brand" aria-label="TLB cl8us - apoio à gestão de contratos">
            <div class="tlb-cl8us-brand-main">
                <span class="tlb-cl8us-tlb">TLB</span>
                <span class="tlb-cl8us-dot">·</span>
                <span class="tlb-cl8us-name">cl8us</span>
            </div>
            <div class="tlb-cl8us-subtitle">apoio à gestão de contratos</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def agora_brasilia():
    return datetime.now(ZoneInfo("America/Sao_Paulo"))


STATUS_OPCOES = ["Pendente", "Em conferência", "Validado", "Não se aplica"]

def sinal_status(valor):
    mapa = {
        "Validado": "🟢",
        "Pendente": "🔴",
        "Em conferência": "🟡",
        "Não se aplica": "⚪",
    }
    return mapa.get(str(valor), "⚪")

ITENS_PADRAO = [
    ("Parâmetros da análise", "Índice contratual confirmado", "Crítica"),
    ("Parâmetros da análise", "Data-base conferida", "Crítica"),
    ("Parâmetros da análise", "Pedido dentro da janela ou tratamento justificado", "Crítica"),
    ("Parâmetros da análise", "Ciclos revisados", "Crítica"),
    ("Parâmetros da análise", "Ciclo negativo, se houver, tratado como 0,00%", "Crítica"),
    ("Arquivo de coleta", "Financeiro mensal preenchido", "Crítica"),
    ("Arquivo de coleta", "Itens e saldos remanescentes preenchidos", "Crítica"),
    ("Arquivo de coleta", "Aditivos e supressões conferidos", "Alta"),
    ("Arquivo de coleta", "Contexto anterior registrado, se houver", "Média"),
    ("Valor Global", "Valor Total Atualizado conferido", "Crítica"),
    ("Valor Global", "Aditivos não somados autonomamente ao Valor Total Atualizado", "Crítica"),
    ("Valor Global", "Saldo remanescente atualizado validado", "Crítica"),
    ("Valor Global", "Auditoria sem pendência crítica", "Crítica"),
    ("Relatórios e documentos", "Planilha Executiva baixada e revisada", "Alta"),
    ("Relatórios e documentos", "Relatório Executivo revisado", "Alta"),
    ("Relatórios e documentos", "Mapa dos Marcos Contratuais gerado, se aplicável", "Média"),
    ("Relatórios e documentos", "Minuta DOCX gerada e campos pendentes revisados", "Alta"),
    ("Garantia", "Valor-base da garantia conferido", "Alta"),
    ("Garantia", "Percentual da garantia conferido", "Alta"),
    ("Garantia", "Garantia constituída informada", "Alta"),
    ("Garantia", "Endosso necessário calculado", "Alta"),
    ("Instrução processual", "Adequação orçamentária verificada", "Crítica"),
    ("Instrução processual", "Certidões emitidas e juntadas", "Crítica"),
    ("Instrução processual", "Certidões válidas na véspera da assinatura", "Crítica"),
    ("Instrução processual", "Manifestação da área gestora juntada ou solicitada", "Alta"),
    ("Instrução processual", "Minuta ou instrumento formal revisado", "Alta"),
    ("Instrução processual", "Assinaturas e encaminhamento final conferidos", "Média"),
]


def checklist_inicial():
    return pd.DataFrame(
        [
            {
                "Grupo": grupo,
                "Item": item,
                "Status": "Pendente",
                "Criticidade": criticidade,
                "Observação": "",
            }
            for grupo, item, criticidade in ITENS_PADRAO
        ]
    )


def gerar_excel_checklist(df, resumo):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="CHECKLIST")
        workbook = writer.book
        ws = writer.sheets["CHECKLIST"]
        fmt_header = workbook.add_format({
            "bold": True, "bg_color": "#1F4E79", "font_color": "#FFFFFF",
            "border": 1, "align": "center", "valign": "vcenter"
        })
        fmt_wrap = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})
        fmt_validado = workbook.add_format({"bg_color": "#E2F0D9", "border": 1})
        fmt_pendente = workbook.add_format({"bg_color": "#FCE4D6", "border": 1})
        fmt_conferencia = workbook.add_format({"bg_color": "#FFF2CC", "border": 1})
        fmt_na = workbook.add_format({"bg_color": "#E7E6E6", "border": 1})

        for col_num, value in enumerate(df.columns):
            ws.write(0, col_num, value, fmt_header)
        widths = [26, 54, 18, 14, 48]
        for col_num, width in enumerate(widths):
            ws.set_column(col_num, col_num, width, fmt_wrap)
        ws.freeze_panes(1, 0)

        status_col = df.columns.get_loc("Status")
        last_row = len(df)
        ws.conditional_format(1, status_col, last_row, status_col, {
            "type": "text", "criteria": "containing", "value": "Validado", "format": fmt_validado
        })
        ws.conditional_format(1, status_col, last_row, status_col, {
            "type": "text", "criteria": "containing", "value": "Pendente", "format": fmt_pendente
        })
        ws.conditional_format(1, status_col, last_row, status_col, {
            "type": "text", "criteria": "containing", "value": "Em conferência", "format": fmt_conferencia
        })
        ws.conditional_format(1, status_col, last_row, status_col, {
            "type": "text", "criteria": "containing", "value": "Não se aplica", "format": fmt_na
        })

        resumo.to_excel(writer, index=False, sheet_name="RESUMO")
        ws2 = writer.sheets["RESUMO"]
        for col_num, value in enumerate(resumo.columns):
            ws2.write(0, col_num, value, fmt_header)
            ws2.set_column(col_num, col_num, 28, fmt_wrap)
    buffer.seek(0)
    return buffer.getvalue()


def carregar_checklist_xlsx(arquivo):
    """Restaura o checklist a partir do XLSX exportado pelo próprio módulo."""
    try:
        try:
            df = pd.read_excel(arquivo, sheet_name="CHECKLIST")
        except Exception:
            arquivo.seek(0)
            df = pd.read_excel(arquivo)
    except Exception as exc:
        raise ValueError(f"Não foi possível ler o arquivo enviado: {exc}")

    colunas_obrigatorias = ["Grupo", "Item", "Status", "Criticidade", "Observação"]
    for coluna in colunas_obrigatorias:
        if coluna not in df.columns:
            if coluna == "Observação":
                df[coluna] = ""
            else:
                raise ValueError(f"O arquivo não possui a coluna obrigatória: {coluna}")

    df = df[colunas_obrigatorias].copy()
    df["Status"] = df["Status"].where(df["Status"].isin(STATUS_OPCOES), "Pendente")
    df["Observação"] = df["Observação"].fillna("")
    return df


def render_checklist_status_html(df):
    """Renderiza tabela visual com cores estáveis para os status."""
    cores = {
        "Validado": ("#E2F0D9", "#274E13", "#A9D18E"),
        "Pendente": ("#FCE4D6", "#7F1D1D", "#F4B183"),
        "Em conferência": ("#FFF2CC", "#7C5700", "#FFD966"),
        "Não se aplica": ("#E7E6E6", "#374151", "#BFBFBF"),
    }
    colunas = ["Grupo", "Item", "Status", "Criticidade", "Observação"]
    linhas = []
    for _, row in df[colunas].iterrows():
        status = str(row.get("Status", "") or "")
        bg, fg, border = cores.get(status, ("#FFFFFF", "#111827", "#E5E7EB"))
        linhas.append(
            "<tr>"
            f"<td>{html.escape(str(row.get('Grupo', '') or ''))}</td>"
            f"<td>{html.escape(str(row.get('Item', '') or ''))}</td>"
            f"<td><span class='status-pill' style='background:{bg}; color:{fg}; border-color:{border};'>{html.escape(status)}</span></td>"
            f"<td>{html.escape(str(row.get('Criticidade', '') or ''))}</td>"
            f"<td>{html.escape(str(row.get('Observação', '') or ''))}</td>"
            "</tr>"
        )
    return f"""
    <style>
    .checklist-status-table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 0.86rem;
        margin-top: 0.25rem;
    }}
    .checklist-status-table th {{
        background: #1F4E79;
        color: white;
        padding: 0.48rem;
        border: 1px solid #D9E2F3;
        text-align: left;
    }}
    .checklist-status-table td {{
        padding: 0.44rem;
        border: 1px solid #E5EAF0;
        vertical-align: top;
    }}
    .checklist-status-table tr:nth-child(even) td {{
        background: #F8FAFC;
    }}
    .status-pill {{
        display: inline-block;
        min-width: 112px;
        padding: 0.18rem 0.52rem;
        border: 1px solid;
        border-radius: 999px;
        font-weight: 700;
        text-align: center;
    }}
    </style>
    <table class="checklist-status-table">
        <thead>
            <tr>
                <th>Grupo</th>
                <th>Item</th>
                <th>Status</th>
                <th>Criticidade</th>
                <th>Observação</th>
            </tr>
        </thead>
        <tbody>
            {''.join(linhas)}
        </tbody>
    </table>
    """


render_marca_topo()

st.title("Checklist Processual")
st.caption("Controle de prontidão da instrução. O checklist orienta o fluxo, mas não bloqueia o uso do sistema.")

if "checklist_processual" not in st.session_state or not isinstance(st.session_state.get("checklist_processual"), pd.DataFrame):
    st.session_state["checklist_processual"] = checklist_inicial()

col_top1, col_top2 = st.columns([1, 2])
with col_top1:
    if st.button("Restaurar checklist padrão", type="secondary"):
        st.session_state["checklist_processual"] = checklist_inicial()
        st.rerun()
with col_top2:
    st.info("Use os status para diferenciar o que está pendente, em conferência, validado ou não aplicável.")

with st.expander("Restaurar checklist a partir de XLSX", expanded=False):
    arquivo_checklist = st.file_uploader(
        "Enviar Checklist_Processual.xlsx salvo anteriormente",
        type=["xlsx"],
        key="upload_checklist_processual",
    )
    if arquivo_checklist is not None:
        if st.button("Importar status do checklist enviado", type="primary"):
            try:
                st.session_state["checklist_processual"] = carregar_checklist_xlsx(arquivo_checklist)
                st.success("Checklist importado com sucesso. Os status e observações foram restaurados.")
                st.rerun()
            except Exception as exc:
                st.error(f"Não foi possível importar o checklist: {exc}")

df_atual = st.session_state["checklist_processual"].copy()
if "Sinal" in df_atual.columns:
    df_atual = df_atual.drop(columns=["Sinal"])

df_editor = df_atual.copy()
df_editor.insert(0, "Sinal", df_editor["Status"].apply(sinal_status))

edited = st.data_editor(
    df_editor,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Sinal": st.column_config.TextColumn("Sinal", disabled=True, width="small"),
        "Grupo": st.column_config.TextColumn("Grupo", disabled=True),
        "Item": st.column_config.TextColumn("Item", disabled=True),
        "Status": st.column_config.SelectboxColumn("Status", options=STATUS_OPCOES, required=True),
        "Criticidade": st.column_config.TextColumn("Criticidade", disabled=True),
        "Observação": st.column_config.TextColumn("Observação"),
    },
    key="checklist_processual_editor",
)

if "Sinal" in edited.columns:
    edited = edited.drop(columns=["Sinal"])
st.session_state["checklist_processual"] = edited.copy()

total = len(edited)
validados = int((edited["Status"] == "Validado").sum())
pendentes = int((edited["Status"] == "Pendente").sum())
em_conferencia = int((edited["Status"] == "Em conferência").sum())
nao_aplica = int((edited["Status"] == "Não se aplica").sum())
criticos_pendentes = int(((edited["Criticidade"] == "Crítica") & (edited["Status"].isin(["Pendente", "Em conferência"]))).sum())

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Itens", total)
c2.metric("Validados", validados)
c3.metric("Pendentes", pendentes)
c4.metric("Em conferência", em_conferencia)
c5.metric("Críticos pendentes", criticos_pendentes)

if criticos_pendentes:
    st.warning(f"Há {criticos_pendentes} item(ns) crítico(s) pendente(s) ou em conferência.")
else:
    st.success("Não há itens críticos pendentes ou em conferência.")

resumo = pd.DataFrame([
    {"Indicador": "Data de geração", "Valor": agora_brasilia().strftime("%d/%m/%Y %H:%M")},
    {"Indicador": "Total de itens", "Valor": total},
    {"Indicador": "Validados", "Valor": validados},
    {"Indicador": "Pendentes", "Valor": pendentes},
    {"Indicador": "Em conferência", "Valor": em_conferencia},
    {"Indicador": "Não se aplica", "Valor": nao_aplica},
    {"Indicador": "Críticos pendentes", "Valor": criticos_pendentes},
])

st.divider()
st.subheader("Exportar checklist")
excel = gerar_excel_checklist(edited, resumo)
st.session_state["arquivo_checklist_processual_xlsx"] = excel
st.download_button(
    "Baixar Checklist Processual em XLSX",
    data=excel,
    file_name="Checklist_Processual.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.caption("Recomendação: anexar ou consultar este checklist antes da geração final de relatórios e documentos.")
