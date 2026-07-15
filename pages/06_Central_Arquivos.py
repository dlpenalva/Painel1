from __future__ import annotations

import html
from typing import Any

import streamlit as st

from _capacidades_apuracao import avaliar_capacidades_apuracao
from _ui_utils import render_marca_topo


st.set_page_config(page_title="TLB · cl8us - Central de Arquivos", layout="wide")


DOCUMENTOS = (
    {
        "nome": "Relatório Executivo",
        "descricao": "Síntese da apuração e dos fundamentos disponíveis para conferência.",
        "formato": "PDF",
        "session_key": "arquivo_relatorio_executivo_pdf",
        "file_name": "Relatorio_Executivo_Analise_Reajuste.pdf",
        "mime": "application/pdf",
        "pagina": "pages/04_Relatorio_Global.py",
    },
    {
        "nome": "Minuta de Apostilamento",
        "descricao": "Minuta editável com os dados disponíveis e campos pendentes destacados.",
        "formato": "DOCX",
        "session_key": "arquivo_minuta_apostilamento_docx",
        "file_name": "minuta_termo_apostilamento.docx",
        "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pagina": "pages/04_Relatorio_Global.py",
    },
    {
        "nome": "Despacho Saneador",
        "descricao": "Minuta narrativa para revisão da instrução e das pendências identificadas.",
        "formato": "DOCX",
        "session_key": "arquivo_saneador_docx",
        "file_name": "Saneador_Instrucao_Processual.docx",
        "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pagina": "pages/10_Saneador.py",
    },
    {
        "nome": "Previsão Orçamentária",
        "descricao": "Memorando editável com a projeção e a complementação disponíveis.",
        "formato": "DOCX",
        "session_key": "arquivo_previsao_orcamentaria_docx",
        "file_name": "memorando_adequacao_orcamentaria.docx",
        "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pagina": "pages/12_Adequacao_Orcamentaria.py",
    },
    {
        "nome": "Extrato para Publicação (DOU)",
        "descricao": "Minuta do extrato para publicação com lacunas explicitamente sinalizadas.",
        "formato": "DOCX",
        "session_key": "arquivo_dou_docx",
        "file_name": "DOU_Extrato_Termo_de_Apostila.docx",
        "mime": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "pagina": "pages/13_DOU.py",
    },
    {
        "nome": "Sumário do Reajuste",
        "descricao": "Planilha executiva para conferência financeira e memória consolidada.",
        "formato": "XLSX",
        "session_key": "arquivo_planilha_executiva_xlsx",
        "file_name": "Sumario_Executivo_Reajuste.xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "pagina": "pages/03_Valor_Global.py",
    },
    {
        "nome": "Itens por Ciclo",
        "descricao": "Evolução dos valores unitários e totais remanescentes por ciclo.",
        "formato": "XLSX",
        "session_key": "arquivo_valores_unitarios_xlsx",
        "file_name": "Itens_por_Ciclo.xlsx",
        "mime": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "pagina": "pages/03_Valor_Global.py",
    },
    {
        "nome": "Garantia Contratual",
        "descricao": "Relatório de conferência, endosso e monitoramento da garantia.",
        "formato": "PDF",
        "session_key": "arquivo_garantia_pdf",
        "file_name": "relatorio_garantia_contratual.pdf",
        "mime": "application/pdf",
        "pagina": "pages/05_Garantia.py",
    },
)


def aplicar_css_central() -> None:
    st.markdown(
        """
        <style>
        .central-intro {
            color:#52667A;
            font-size:.94rem;
            line-height:1.45;
            margin:-.25rem 0 1rem;
            max-width:760px;
        }
        .central-status {
            background:#F7FBFE;
            border:1px solid #D7E4EE;
            border-radius:14px;
            padding:14px 16px 15px;
            margin:0 0 18px;
            box-shadow:0 4px 14px rgba(31,78,121,.045);
        }
        .central-status-head {
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:12px;
            margin-bottom:10px;
        }
        .central-status-title {
            color:#173F63;
            font-size:.96rem;
            font-weight:800;
        }
        .central-status-count {
            color:#315C7C;
            background:#EAF3F9;
            border:1px solid #C9DCE9;
            border-radius:999px;
            font-size:.72rem;
            font-weight:750;
            padding:4px 9px;
            white-space:nowrap;
        }
        .central-status-grid {
            display:grid;
            grid-template-columns:repeat(5,minmax(0,1fr));
            gap:8px;
        }
        .central-status-item {
            display:flex;
            align-items:center;
            gap:7px;
            min-width:0;
            background:#FFFFFF;
            border:1px solid #DFE8EF;
            border-radius:9px;
            padding:8px 9px;
            color:#334E68;
            font-size:.79rem;
            font-weight:700;
        }
        .central-status-dot {
            width:8px;
            height:8px;
            border-radius:50%;
            background:#9BAAB7;
            flex:0 0 auto;
        }
        .central-status-item.completo .central-status-dot { background:#3D9660; }
        .central-status-item.parcial .central-status-dot { background:#D49A23; }
        .central-status-item.bloqueado .central-status-dot { background:#D27045; }
        div[data-testid="stVerticalBlockBorderWrapper"] {
            background:#FFFFFF;
            border-color:#DCE6EF !important;
            border-radius:13px !important;
            box-shadow:0 5px 16px rgba(31,78,121,.05);
        }
        div[data-testid="stVerticalBlockBorderWrapper"] h3 {
            color:#173F63;
            font-size:1rem;
            line-height:1.25;
            margin-bottom:.15rem;
        }
        div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] .central-card-description) {
            box-sizing:border-box;
            height:18rem;
            min-height:18rem;
            max-height:18rem;
            display:flex;
            flex-direction:column;
            position:relative;
        }
        div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] .central-card-description) > div[data-testid="stElementContainer"]:has(h3) {
            height:8.05rem;
            min-height:8.05rem;
            max-height:8.05rem;
            overflow:hidden;
        }
        div[data-testid="stColumn"] div[data-testid="stVerticalBlock"]:has(> div[data-testid="stElementContainer"] .central-card-description) > div[data-testid="stElementContainer"]:last-child {
            position:absolute;
            right:15px;
            bottom:15px;
            left:15px;
            margin:0 !important;
        }
        div[data-testid="stColumn"] div[data-testid="stPageLink"] a[data-testid="stPageLink-NavLink"] {
            box-sizing:border-box;
            display:flex;
            align-items:center;
            justify-content:center;
            width:100%;
            height:2.48rem;
            min-height:2.48rem;
            color:#173F63 !important;
            background:#FFFFFF;
            border:1px solid #9EB6C9;
            border-radius:8px;
            padding:.42rem .72rem;
            margin:0 !important;
            transform:translateY(-.04rem);
            font-weight:650;
            text-decoration:none !important;
            box-shadow:0 2px 6px rgba(31,78,121,.05);
        }
        div[data-testid="stColumn"] div[data-testid="stPageLink"] a[data-testid="stPageLink-NavLink"]:hover {
            color:#0F3555 !important;
            background:#EEF6FB;
            border-color:#6E96B3;
        }
        .central-card-description {
            color:#5B6F82;
            font-size:.82rem;
            line-height:1.4;
            height:3.5rem;
            min-height:3.5rem;
            max-height:3.5rem;
            display:-webkit-box;
            -webkit-box-orient:vertical;
            -webkit-line-clamp:3;
            overflow:hidden;
            margin-bottom:.45rem;
        }
        @media (max-width:900px) {
            .central-status-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
        }
        @media (max-width:560px) {
            .central-status-head { align-items:flex-start; flex-direction:column; }
            .central-status-grid { grid-template-columns:1fr; }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_status_entradas(capacidades: dict[str, Any]) -> None:
    blocos = capacidades.get("blocos") or {}
    ordem = ("financeiro", "itens", "pcs", "consumidos", "remanescentes")
    cards = []
    completos = 0
    for chave in ordem:
        item = blocos.get(chave) or {"nome": chave.title(), "estado": "nao_informado"}
        estado = str(item.get("estado") or "nao_informado")
        if estado == "completo":
            completos += 1
        nome = html.escape(str(item.get("nome") or chave.title()))
        cards.append(
            f'<div class="central-status-item {html.escape(estado)}">'
            '<span class="central-status-dot" aria-hidden="true"></span>'
            f'<span>{nome}</span></div>'
        )
    st.markdown(
        '<section class="central-status" aria-label="Status da Apuração">'
        '<div class="central-status-head">'
        '<div class="central-status-title">Status da Apuração</div>'
        f'<div class="central-status-count">{completos} de {len(ordem)} blocos preenchidos</div>'
        '</div>'
        f'<div class="central-status-grid">{"".join(cards)}</div>'
        '</section>',
        unsafe_allow_html=True,
    )


def render_documento(documento: dict[str, str], estrutura_valida: bool) -> None:
    arquivo = st.session_state.get(documento["session_key"])
    with st.container(border=True):
        st.markdown(f"### {documento['nome']} · {documento['formato']}")
        st.markdown(
            f'<div class="central-card-description">{html.escape(documento["descricao"])}</div>',
            unsafe_allow_html=True,
        )
        if not estrutura_valida:
            st.button(
                "Corrigir XLS para continuar",
                key=f"central_bloqueado_{documento['session_key']}",
                disabled=True,
                use_container_width=True,
            )
        elif arquivo:
            st.download_button(
                "Baixar documento",
                data=arquivo,
                file_name=documento["file_name"],
                mime=documento["mime"],
                key=f"central_download_{documento['session_key']}",
                use_container_width=True,
            )
        else:
            st.page_link(
                documento["pagina"],
                label="Gerar e baixar",
                use_container_width=True,
            )


render_marca_topo()
aplicar_css_central()
st.title("Central de Arquivos")
st.markdown(
    '<div class="central-intro">Os oito documentos oficiais em um único lugar. '
    'Cada arquivo utiliza as informações disponíveis e mantém eventuais pendências no próprio conteúdo.</div>',
    unsafe_allow_html=True,
)

diagnostico = st.session_state.get("diagnostico_coleta_v2") or {}
CAPACIDADES = diagnostico.get("capacidades") or avaliar_capacidades_apuracao({}, {})
render_status_entradas(CAPACIDADES)

estrutura_valida = bool(CAPACIDADES.get("estruturalmente_valido", True))
for inicio in range(0, len(DOCUMENTOS), 4):
    colunas = st.columns(4)
    for coluna, documento in zip(colunas, DOCUMENTOS[inicio : inicio + 4]):
        with coluna:
            render_documento(documento, estrutura_valida)
