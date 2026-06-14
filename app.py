import streamlit as st
from pathlib import Path
from _ui_utils import render_versao_sidebar

st.set_page_config(
    page_title="cl8us · Gestão de Contratos",
    page_icon="📋",
    layout="wide",
)

APP_SEM_SENHA_VERSAO  = "20260516_RESTORE_SEM_SENHA"
APP_DOU_13_VERSAO     = "20260516_DOU_13_GESTAO"
APP_ORIENTA_11_VERSAO = "20260516_CL8US_ORIENTA_SOLO_APOIO"
APP_IST_ALERTA_VERSAO = "20260516_ALERTA_IST_ULTIMA_COMPETENCIA"

st.markdown("""
<style>
/* ═══════════════════════════════════════════════════
   SIDEBAR — dark slate, funciona com nav + formulário
═══════════════════════════════════════════════════ */
section[data-testid="stSidebar"] {
    background: #1E293B !important;
}
section[data-testid="stSidebar"] > div,
section[data-testid="stSidebar"] > div > div {
    background: #1E293B !important;
}

/* ── Texto geral ── */
section[data-testid="stSidebar"],
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] small,
section[data-testid="stSidebar"] div {
    color: #CBD5E1 !important;
}

/* ── Grupo de nav "Fluxo principal" ── */
section[data-testid="stSidebar"] [role="heading"],
section[data-testid="stSidebar"] [role="separator"] {
    color: #64748B !important;
    font-size: 0.66rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
    background: transparent !important;
}

/* ── Links de nav ── */
section[data-testid="stSidebar"] nav a,
section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a {
    background: transparent !important;
    border: none !important;
    box-shadow: none !important;
    color: #CBD5E1 !important;
    font-size: 0.88rem !important;
    font-weight: 400 !important;
    border-radius: 6px !important;
}
section[data-testid="stSidebar"] nav a *,
section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a * {
    color: #CBD5E1 !important;
}

/* ── Hover ── */
section[data-testid="stSidebar"] nav a:hover,
section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a:hover {
    background: rgba(255,255,255,0.06) !important;
    color: #F1F5F9 !important;
}

/* ── Item ativo ── */
section[data-testid="stSidebar"] nav a[aria-current="page"],
section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] {
    background: #1F4E78 !important;
    border-radius: 6px !important;
    box-shadow: none !important;
    border: none !important;
}
section[data-testid="stSidebar"] nav a[aria-current="page"] *,
section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] * {
    color: #ffffff !important;
    font-weight: 500 !important;
    background: transparent !important;
}

/* ── Neutralizar pseudo-elementos e li ── */
section[data-testid="stSidebar"] nav a::before,
section[data-testid="stSidebar"] nav a::after,
section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a::before,
section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a::after {
    display: none !important;
    background: none !important;
}
section[data-testid="stSidebar"] ul,
section[data-testid="stSidebar"] li,
section[data-testid="stSidebar"] nav,
section[data-testid="stSidebar"] [data-testid="stSidebarNav"] {
    background: transparent !important;
    border: none !important;
}

/* ── Inputs e campos de formulário ── */
section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] textarea,
section[data-testid="stSidebar"] select {
    background: #0F172A !important;
    color: #F1F5F9 !important;
    -webkit-text-fill-color: #F1F5F9 !important;
    border: 1px solid #334155 !important;
    border-radius: 6px !important;
}

/* ── Botões +/- do number_input ── */
section[data-testid="stSidebar"] button[kind="secondary"],
section[data-testid="stSidebar"] [data-testid="stNumberInput"] button,
section[data-testid="stSidebar"] button {
    background: #334155 !important;
    color: #F1F5F9 !important;
    border: 1px solid #475569 !important;
    border-radius: 4px !important;
}
section[data-testid="stSidebar"] button:hover {
    background: #475569 !important;
}

/* ── Selectbox ── */
section[data-testid="stSidebar"] [data-baseweb="select"] > div,
section[data-testid="stSidebar"] [data-testid="stSelectbox"] > div > div {
    background: #0F172A !important;
    border-color: #334155 !important;
    color: #F1F5F9 !important;
}

/* ── Scrollbar ── */
section[data-testid="stSidebar"] ::-webkit-scrollbar { width: 3px; }
section[data-testid="stSidebar"] ::-webkit-scrollbar-track { background: transparent; }
section[data-testid="stSidebar"] ::-webkit-scrollbar-thumb {
    background: #334155;
    border-radius: 4px;
}

/* ════════════════════════════════
   BOXES COLORIDOS DA SIDEBAR
════════════════════════════════ */

/* Neutralizar qualquer fundo claro em divs da sidebar */
section[data-testid="stSidebar"] div[style*="background"],
section[data-testid="stSidebar"] div[style*="background-color"] {
    background: rgba(255,255,255,0.06) !important;
    border-color: rgba(255,255,255,0.12) !important;
}

/* Box laranja (lembrete C0) */
section[data-testid="stSidebar"] div[style*="#FFF7E6"],
section[data-testid="stSidebar"] div[style*="FFF7E6"],
section[data-testid="stSidebar"] div[style*="#FFEAA7"],
section[data-testid="stSidebar"] div[style*="rgba(255, 152, 0"] {
    background: rgba(251,146,60,0.15) !important;
    border-color: rgba(251,146,60,0.35) !important;
}

/* Box roxo (índice do contrato) */
section[data-testid="stSidebar"] div[style*="#F5F0FF"],
section[data-testid="stSidebar"] div[style*="#EDE9FE"],
section[data-testid="stSidebar"] div[style*="#F5F3FF"],
section[data-testid="stSidebar"] div[style*="4C1D95"],
section[data-testid="stSidebar"] div[style*="C4B5FD"] {
    background: rgba(139,92,246,0.15) !important;
    border-color: rgba(139,92,246,0.3) !important;
}

/* Box azul claro (alertas IST/ICTI) */
section[data-testid="stSidebar"] div[style*="#F8FAFC"],
section[data-testid="stSidebar"] div[style*="#F0F9FF"],
section[data-testid="stSidebar"] div[style*="#EFF6FF"],
section[data-testid="stSidebar"] div[style*="#DBEAFE"],
section[data-testid="stSidebar"] div[style*="E5EAF0"],
section[data-testid="stSidebar"] div[style*="BAE6FD"] {
    background: rgba(59,130,246,0.12) !important;
    border-color: rgba(59,130,246,0.25) !important;
}

/* Texto dentro dos boxes — claro */
section[data-testid="stSidebar"] div[style*="background"] p,
section[data-testid="stSidebar"] div[style*="background"] span,
section[data-testid="stSidebar"] div[style*="background"] div,
section[data-testid="stSidebar"] div[style*="background"] strong {
    color: #E2E8F0 !important;
}

/* ════════════════════════════════
   DOWNLOAD BUTTON — verde escuro
════════════════════════════════ */
div[data-testid="stDownloadButton"] > button {
    background-color: #1a5c38 !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 14px !important;
    font-weight: 600 !important;
}
div[data-testid="stDownloadButton"] > button:hover {
    background-color: #14472b !important;
}
</style>
""", unsafe_allow_html=True)

pages_dir = Path("pages")

p_calc    = st.Page("pages/00_Calculadora_Reajustes.py", title="Calculadora de Reajustes", icon="🧮", default=True)
p_analise = st.Page("pages/15_Analise_Contratual.py",    title="Análise e Documentos",     icon="📄")
nav = {"Fluxo principal": [p_calc, p_analise]}

pg = st.navigation(nav)
render_versao_sidebar()
pg.run()
