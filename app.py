import streamlit as st
from pathlib import Path
from _ui_utils import render_versao_sidebar

st.set_page_config(page_title="TLB · cl8us - Apoio à Gestão de Contratos", layout="wide")

# App sem barreira de senha interna.
APP_SEM_SENHA_VERSAO = "20260516_RESTORE_SEM_SENHA"
APP_DOU_13_VERSAO = "20260516_DOU_13_GESTAO"
APP_ORIENTA_11_VERSAO = "20260516_CL8US_ORIENTA_SOLO_APOIO"
APP_IST_ALERTA_VERSAO = "20260516_ALERTA_IST_ULTIMA_COMPETENCIA"

st.markdown(
    """
    <style>
    :root {
        --cl8us-sidebar: #C6D9E8;
        --cl8us-bg-start: #D4E3EF;
        --cl8us-bg-end: #BED4E5;
        --cl8us-navy: #123B63;
        --cl8us-action: #7A1733;
        --cl8us-action-hover: #641229;
    }
    .stAppViewContainer {
        background: linear-gradient(135deg, var(--cl8us-bg-start) 0%, var(--cl8us-bg-end) 100%);
        color: var(--cl8us-navy);
    }
    [data-testid="stHeader"] {
        background: rgba(212, 227, 239, 0.82);
    }
    section[data-testid="stSidebar"],
    section[data-testid="stSidebar"] > div {
        background: var(--cl8us-sidebar) !important;
    }
    section[data-testid="stSidebar"] * {
        color: var(--cl8us-navy);
    }
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a[aria-current="page"] {
        background: rgba(18, 59, 99, 0.12);
        border-left: 4px solid var(--cl8us-navy);
        border-radius: 0 10px 10px 0;
    }
    .stMainBlockContainer h1,
    .stMainBlockContainer h2,
    .stMainBlockContainer h3 {
        color: var(--cl8us-navy);
    }
    .stButton > button,
    .stDownloadButton > button {
        border-radius: 999px;
        font-weight: 700;
        min-height: 2.55rem;
        box-shadow: 0 5px 14px rgba(18, 59, 99, 0.13);
        transition: transform 120ms ease, box-shadow 120ms ease, background 120ms ease;
    }
    .stButton > button[kind="primary"],
    .stDownloadButton > button[kind="primary"] {
        background: var(--cl8us-action);
        border-color: var(--cl8us-action);
        color: #FFFFFF;
    }
    .stButton > button[kind="primary"]:hover,
    .stDownloadButton > button[kind="primary"]:hover {
        background: var(--cl8us-action-hover);
        border-color: var(--cl8us-action-hover);
        color: #FFFFFF;
        transform: translateY(-1px);
        box-shadow: 0 8px 18px rgba(100, 18, 41, 0.22);
    }
    .stButton > button[kind="secondary"],
    .stDownloadButton > button[kind="secondary"] {
        background: rgba(255, 255, 255, 0.72);
        border-color: rgba(18, 59, 99, 0.38);
        color: var(--cl8us-navy);
    }
    [data-testid="stMetric"],
    [data-testid="stFileUploaderDropzone"],
    [data-testid="stExpander"] {
        background: rgba(255, 255, 255, 0.68);
        border-color: rgba(18, 59, 99, 0.15);
        border-radius: 14px;
    }
    section[data-testid="stSidebar"] nav a,
    section[data-testid="stSidebar"] nav a *,
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a,
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] a * {
        font-weight: 400 !important;
    }
    section[data-testid="stSidebar"] nav [role="heading"],
    section[data-testid="stSidebar"] nav [role="heading"] *,
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] [role="heading"],
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] [role="heading"] *,
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] div:not(:has(a)) > span,
    section[data-testid="stSidebar"] [data-testid="stSidebarNav"] div:not(:has(a)) > p {
        font-weight: 700 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

pages_dir = Path("pages")
p0 = st.Page("pages/00_Calculadora_Reajustes.py", title="Calculadora de Reajustes", default=True)

nav = {"📌 Admissibilidade e Cálculo": [p0]}

grupo_valor_global = []
if (pages_dir / "03_Valor_Global.py").exists():
    grupo_valor_global.append(st.Page("pages/03_Valor_Global.py", title="Valores"))
if (pages_dir / "04_Relatorio_Global.py").exists():
    grupo_valor_global.append(st.Page("pages/04_Relatorio_Global.py", title="Relatórios"))
if (pages_dir / "06_Central_Arquivos.py").exists():
    grupo_valor_global.append(st.Page("pages/06_Central_Arquivos.py", title="Central de Arquivos"))
if (pages_dir / "07_Checklist_Processual.py").exists():
    grupo_valor_global.append(st.Page("pages/07_Checklist_Processual.py", title="Checklist Processual"))
if grupo_valor_global:
    nav["🌐 Visão Global e Relatórios"] = grupo_valor_global

grupo_gestao = []
if (pages_dir / "05_Garantia.py").exists():
    grupo_gestao.append(st.Page("pages/05_Garantia.py", title="Gestão da Garantia"))
if (pages_dir / "12_Adequacao_Orcamentaria.py").exists():
    grupo_gestao.append(st.Page("pages/12_Adequacao_Orcamentaria.py", title="Adequação Orçamentária"))
if (pages_dir / "08_Avaliacao_Aditivos.py").exists():
    grupo_gestao.append(st.Page("pages/08_Avaliacao_Aditivos.py", title="Aditivos: 25%"))
if (pages_dir / "13_DOU.py").exists():
    grupo_gestao.append(st.Page("pages/13_DOU.py", title="DOU"))
if grupo_gestao:
    nav["🛡️ Gestão Contratual"] = grupo_gestao

grupo_instrucao = []
if (pages_dir / "09_Infos_Previas.py").exists():
    grupo_instrucao.append(st.Page("pages/09_Infos_Previas.py", title="Infos Prévias"))
if (pages_dir / "10_Saneador.py").exists():
    grupo_instrucao.append(st.Page("pages/10_Saneador.py", title="Saneador"))
if grupo_instrucao:
    nav["🧾 Instrução Processual"] = grupo_instrucao

# Último bloco do menu: apoio, com Cl8us Orienta isolado.
grupo_apoio = []
if (pages_dir / "11_Cl8us_Orienta.py").exists():
    grupo_apoio.append(st.Page("pages/11_Cl8us_Orienta.py", title="Cl8us Orienta"))
if grupo_apoio:
    nav["🧭 Apoio"] = grupo_apoio

pg = st.navigation(nav)
render_versao_sidebar()
pg.run()
