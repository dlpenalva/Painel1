import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="TLB · cl8us - Central de Arquivos", layout="wide")


from _ui_utils import render_marca_topo

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
        "Valores Unitários e Totais por Ciclo",
        "Conferência dos valores unitários e totais remanescentes por ciclo.",
        "XLSX",
        "arquivo_valores_unitarios_xlsx",
        "Valores_Unitarios_e_Totais_por_Ciclo.xlsx",
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
    item_arquivo(
        "Infos Prévias",
        "Levantamento mínimo de dados e documentos antes da análise.",
        "XLSX",
        "arquivo_infos_previas_xlsx",
        "Infos_Previas_Instrucao_Processual.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Infos Prévias",
        "pages/09_Infos_Previas.py",
    ),
    item_arquivo(
        "Saneador",
        "Minuta narrativa integrada para conferência antes da assinatura.",
        "DOCX",
        "arquivo_saneador_docx",
        "Saneador_Instrucao_Processual.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Saneador",
        "pages/10_Saneador.py",
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

# A exportação consolidada da Central foi removida para reduzir redundância.
# Os arquivos devem ser baixados individualmente nos módulos/itens correspondentes.
