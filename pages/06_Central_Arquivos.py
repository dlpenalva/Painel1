import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="TLB · cl8us - Central de Arquivos", layout="wide")


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
        .tlb-cl8us-tlb { color: #123B63; font-size: 1.38rem; font-weight: 750; font-family: "Inter", "Segoe UI", Arial, sans-serif; }
        .tlb-cl8us-dot { color: #C0842B; font-size: 1.18rem; font-weight: 700; }
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
        .central-row {
            border: 1px solid #E5EAF0;
            border-radius: 12px;
            padding: 0.70rem 0.85rem;
            margin-bottom: 0.55rem;
            background: #FFFFFF;
        }
        .central-header {
            background: #F8FAFC;
            border: 1px solid #E5EAF0;
            border-radius: 12px;
            padding: 0.62rem 0.85rem;
            margin-bottom: 0.55rem;
            font-weight: 700;
            color: #0F172A;
        }
        .central-muted {
            color: #64748B;
            font-size: 0.86rem;
        }
        .central-status-ok {
            display: inline-block;
            padding: 0.18rem 0.50rem;
            border-radius: 999px;
            background: #DBEAFE;
            color: #1E3A8A;
            font-size: 0.80rem;
            font-weight: 700;
        }
        .central-status-pendente {
            display: inline-block;
            padding: 0.18rem 0.50rem;
            border-radius: 999px;
            background: #FFF2CC;
            color: #7C5700;
            font-size: 0.80rem;
            font-weight: 700;
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


def gerar_excel_catalogo(df_catalogo, df_checklist=None):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_catalogo.to_excel(writer, index=False, sheet_name="CENTRAL_ARQUIVOS")
        workbook = writer.book
        ws = writer.sheets["CENTRAL_ARQUIVOS"]
        fmt_header = workbook.add_format({
            "bold": True, "bg_color": "#1F4E79", "font_color": "#FFFFFF",
            "border": 1, "align": "center", "valign": "vcenter"
        })
        fmt_wrap = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})
        for col_num, value in enumerate(df_catalogo.columns):
            ws.write(0, col_num, value, fmt_header)
            ws.set_column(col_num, col_num, 24 if col_num not in [1, 4] else 40, fmt_wrap)
        ws.freeze_panes(1, 0)

        if df_checklist is not None and not df_checklist.empty:
            df_checklist.to_excel(writer, index=False, sheet_name="CHECKLIST_PROCESSUAL")
            ws2 = writer.sheets["CHECKLIST_PROCESSUAL"]
            for col_num, value in enumerate(df_checklist.columns):
                ws2.write(0, col_num, value, fmt_header)
                ws2.set_column(col_num, col_num, 26 if col_num != 4 else 44, fmt_wrap)
            ws2.freeze_panes(1, 0)
    buffer.seek(0)
    return buffer.getvalue()


def item_arquivo(nome, finalidade, formato, session_key, file_name, mime, modulo=None, pagina=None):
    disponivel = bool(st.session_state.get(session_key))
    return {
        "Arquivo": nome,
        "Finalidade": finalidade,
        "Formato": formato,
        "Status": "Disponível" if disponivel else "Gerar no módulo",
        "session_key": session_key,
        "file_name": file_name,
        "mime": mime,
        "modulo": modulo or "",
        "pagina": pagina,
    }


def render_linha_arquivo(item):
    arquivo_bytes = st.session_state.get(item["session_key"])
    disponivel = bool(arquivo_bytes)
    status_html = (
        '<span class="central-status-ok">Disponível</span>'
        if disponivel
        else '<span class="central-status-pendente">Gerar no módulo</span>'
    )

    st.markdown('<div class="central-row">', unsafe_allow_html=True)
    col1, col2, col3, col4, col5 = st.columns([1.45, 2.60, 0.70, 1.05, 1.10], vertical_alignment="center")

    with col1:
        st.markdown(f"**{item['Arquivo']}**")
    with col2:
        st.markdown(f"<span class='central-muted'>{item['Finalidade']}</span>", unsafe_allow_html=True)
    with col3:
        st.write(item["Formato"])
    with col4:
        st.markdown(status_html, unsafe_allow_html=True)
    with col5:
        if disponivel:
            st.download_button(
                "Baixar",
                data=arquivo_bytes,
                file_name=item["file_name"],
                mime=item["mime"],
                key=f"download_{item['session_key']}",
                use_container_width=True,
            )
        elif item.get("pagina"):
            st.page_link(item["pagina"], label="Abrir módulo", use_container_width=True)
        else:
            st.caption("Indisponível")

    st.markdown("</div>", unsafe_allow_html=True)


render_marca_topo()

st.title("Central de Arquivos")
st.caption("Área única para localizar os principais arquivos gerados ou utilizados no fluxo de análise.")

itens = [
    item_arquivo(
        "Planilha Executiva",
        "Conferência financeira e memória consolidada da análise.",
        "XLSX",
        "arquivo_planilha_executiva_xlsx",
        "Planilha_Executiva_Analise_Reajuste.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Valor Global",
        "pages/03_Valor_Global.py",
    ),
    item_arquivo(
        "Relatório Executivo",
        "Instrução processual e síntese da análise de reajuste.",
        "PDF",
        "arquivo_relatorio_executivo_pdf",
        "Relatorio_Executivo_Analise_Reajuste.pdf",
        "application/pdf",
        "Relatórios",
        "pages/04_Relatorio_Global.py",
    ),
    item_arquivo(
        "Minuta de Apostilamento",
        "Formalização em DOCX editável.",
        "DOCX",
        "arquivo_minuta_apostilamento_docx",
        "minuta_termo_apostilamento.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Relatórios",
        "pages/04_Relatorio_Global.py",
    ),
    item_arquivo(
        "Mapa dos Marcos",
        "Linha do tempo do contrato e marcos relevantes.",
        "PDF",
        "arquivo_mapa_marcos_pdf",
        "Mapa_Marcos_Contratuais_Linha_do_Tempo.pdf",
        "application/pdf",
        "Valor Global",
        "pages/03_Valor_Global.py",
    ),
    item_arquivo(
        "Checklist Processual",
        "Controle interno de prontidão da instrução.",
        "XLSX",
        "arquivo_checklist_processual_xlsx",
        "Checklist_Processual.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Checklist Processual",
        "pages/07_Checklist_Processual.py",
    ),
    item_arquivo(
        "Garantia Contratual",
        "Endosso, controle e monitoramento da garantia.",
        "PDF",
        "arquivo_garantia_pdf",
        "relatorio_garantia_contratual.pdf",
        "application/pdf",
        "Gestão da Garantia",
        "pages/05_Garantia.py",
    ),
    item_arquivo(
        "Avaliação de Aditivos",
        "Controle de acréscimos, supressões e limite de 25%.",
        "XLSX",
        "arquivo_avaliacao_aditivos_xlsx",
        "Avaliacao_Aditivos_Limite_25.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Avaliação de Aditivos",
        "pages/08_Avaliacao_Aditivos.py",
    ),
]

st.markdown(
    """
    <div class="central-header">
      <div style="display:grid; grid-template-columns: 1.45fr 2.60fr 0.70fr 1.05fr 1.10fr; gap: 1rem; align-items:center;">
        <div>Arquivo</div>
        <div>Finalidade</div>
        <div>Formato</div>
        <div>Status</div>
        <div>Ação</div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

for item in itens:
    render_linha_arquivo(item)

st.divider()
st.subheader("Exportação da Central")

df_catalogo = pd.DataFrame(
    [
        {
            "Arquivo": item["Arquivo"],
            "Finalidade": item["Finalidade"],
            "Formato": item["Formato"],
            "Status": "Disponível" if st.session_state.get(item["session_key"]) else "Gerar no módulo",
            "Módulo": item.get("modulo", ""),
        }
        for item in itens
    ]
)

checklist = st.session_state.get("checklist_processual")
excel_bytes = gerar_excel_catalogo(df_catalogo, checklist if isinstance(checklist, pd.DataFrame) else None)
st.session_state["arquivo_central_arquivos_xlsx"] = excel_bytes
st.download_button(
    "Baixar Central de Arquivos em XLSX",
    data=excel_bytes,
    file_name="Central_de_Arquivos.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.caption("A Central mostra o botão Baixar quando o arquivo já foi gerado na sessão. Caso contrário, oferece atalho para o módulo correspondente.")
