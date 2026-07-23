"""Entrada do Master 2.0 com a casca operacional enxuta do Cl8us 3.0.

O XLS continua sendo o produto principal. A web concentra apenas quatro
movimentos: início, cálculo de um ciclo, cálculo multiciclo e retorno do XLS.
Os módulos legados permanecem registrados, porém fora do menu principal.
"""

from pathlib import Path

import streamlit as st

from _theme import render_cl8us_light_theme
from _ui_utils import render_versao_sidebar


st.set_page_config(
    page_title="TLB · cl8us — Reajustes contratuais",
    page_icon=str(Path(__file__).resolve().parent / "assets" / "cl8us_favicon_512.png"),
    layout="wide",
)

ROOT = Path(__file__).resolve().parent
PAGES_DIR = ROOT / "pages"


def _page(caminho: str, titulo: str, *, default: bool = False) -> st.Page:
    return st.Page(str(PAGES_DIR / caminho), title=titulo, default=default)


PAGINA_INICIO = _page("00_Calculadora_Reajustes.py", "Início", default=True)
PAGINA_UM_CICLO = _page("01_Calculo_Simples.py", "Calculadora 1 ciclo")
PAGINA_MULTICICLO = _page("02_Calculo_Represados.py", "Calculadora multiciclo")
PAGINA_UPLOAD = _page("03_Valor_Global.py", "Upload e docs")

PAGINAS_PRINCIPAIS = (
    PAGINA_INICIO,
    PAGINA_UM_CICLO,
    PAGINA_MULTICICLO,
    PAGINA_UPLOAD,
)

_AUXILIARES = (
    ("04_Relatorio_Global.py", "Relatórios"),
    ("06_Central_Arquivos.py", "Central de arquivos"),
    ("07_Checklist_Processual.py", "Checklist processual"),
    ("05_Garantia.py", "Gestão da garantia"),
    ("12_Adequacao_Orcamentaria.py", "Adequação orçamentária"),
    ("08_Avaliacao_Aditivos.py", "Aditivos: 25%"),
    ("13_DOU.py", "DOU"),
    ("09_Infos_Previas.py", "Informações prévias"),
    ("10_Saneador.py", "Saneador"),
    ("11_Cl8us_Orienta.py", "Cl8us Orienta"),
)
PAGINAS_AUXILIARES = tuple(
    _page(caminho, titulo)
    for caminho, titulo in _AUXILIARES
    if (PAGES_DIR / caminho).exists()
)


def _render_css() -> None:
    st.markdown(
        """
        <style>
        :root {
            --cl8us-sidebar: #C6D9E8;
            --cl8us-main-start: #FBF8F1;
            --cl8us-main-end: #F2ECE1;
            --cl8us-surface: #FFFCF7;
            --cl8us-surface-soft: #F5F0E7;
            --cl8us-input: #FFF9E8;
            --cl8us-step: rgba(255, 255, 255, .42);
            --cl8us-index: #FBE8AD;
            --cl8us-index-line: #D5A11E;
            --cl8us-navy: #123B63;
            --cl8us-muted: #50687D;
            --cl8us-action: #7A1733;
            --cl8us-action-hover: #641229;
            --cl8us-line: rgba(18, 59, 99, .18);
        }
        html, body, [class*="css"] {
            font-family: "Aptos", "Segoe UI", sans-serif;
        }
        .stAppViewContainer {
            color: var(--cl8us-navy);
            background:
                radial-gradient(circle at 92% 2%, rgba(255, 255, 255, .42) 0, rgba(255, 255, 255, 0) 32rem),
                linear-gradient(135deg, var(--cl8us-main-start) 0%, var(--cl8us-main-end) 100%);
        }
        [data-testid="stHeader"] { background: transparent; }
        .block-container { max-width: 1180px; padding-top: 1.45rem; padding-bottom: 3rem; }
        h1, h2, h3 { color: var(--cl8us-navy); letter-spacing: -.015em; }

        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #D4E3EF 0%, var(--cl8us-sidebar) 72%, #BED4E5 100%);
            border-right: 1px solid rgba(18, 59, 99, .14);
            box-shadow: 9px 0 26px rgba(18, 59, 99, .08);
        }
        [data-testid="stSidebarNav"] { display: none; }
        [data-testid="stSidebar"] * { color: var(--cl8us-navy); }
        [data-testid="stSidebar"] [data-testid="stPageLink"] a {
            border-left: 4px solid transparent;
            border-radius: 0 9px 9px 0;
            padding: .18rem .42rem;
            font-weight: 760;
            line-height: 1.16;
            transition: background-color .16s ease, border-color .16s ease, transform .16s ease;
        }
        [data-testid="stSidebar"] [data-testid="stPageLink"] + [data-testid="stPageLink"] {
            margin-top: -.58rem;
        }
        [data-testid="stSidebar"] [data-testid="stElementContainer"]:has(> [data-testid="stPageLink"]) {
            margin-block: -.6rem;
        }
        [data-testid="stSidebar"] [data-testid="stPageLink"] a::before {
            content: "";
            display: inline-block;
            width: .72rem;
            height: .72rem;
            margin-right: .55rem;
            border: 1.5px solid rgba(18, 59, 99, .36);
            border-radius: 999px;
            background: rgba(255, 255, 255, .84);
            vertical-align: -.06rem;
        }
        [data-testid="stSidebar"] [data-testid="stPageLink"] a:hover {
            background: rgba(255, 255, 255, .28);
            transform: translateX(1px);
        }
        [data-testid="stSidebar"] [data-testid="stPageLink"] a[aria-current="page"] {
            background: rgba(255, 255, 255, .48);
            border-left-color: var(--cl8us-navy);
        }
        [data-testid="stSidebar"] [data-testid="stPageLink"] a[aria-current="page"]::before {
            border: 4px solid var(--cl8us-action);
            background: #FFFFFF;
        }
        .cl8us-side-brand { margin: .15rem 0 .1rem 0; line-height: 1; }
        .cl8us-side-brand strong { font-size: .96rem; font-weight: 850; letter-spacing: .01em; }
        .cl8us-side-brand span { color: var(--cl8us-action); padding: 0 .15rem; }
        .cl8us-side-caption { color: var(--cl8us-muted); font-size: .72rem; margin-bottom: .72rem; }
        .cl8us-side-group {
            color: var(--cl8us-navy);
            font-size: .82rem;
            font-weight: 850;
            letter-spacing: .01em;
            margin: .58rem 0 .2rem 0;
        }
        .cl8us-cycle-step {
            background: var(--cl8us-step);
            border-left: 4px solid var(--cl8us-navy);
            border-radius: 0 8px 8px 0;
            color: var(--cl8us-navy);
            font-size: .87rem;
            font-weight: 750;
            letter-spacing: .01em;
            margin: .65rem 0 .28rem;
            padding: .48rem .72rem;
        }
        .cl8us-cycle-step.final { border-left-color: var(--cl8us-action); }
        .cl8us-interval-box {
            background: rgba(255, 253, 247, .82);
            border: 1px solid rgba(18, 59, 99, .20);
            border-radius: 9px;
            color: var(--cl8us-navy);
            font-size: .88rem;
            margin: .55rem 0 .85rem;
            padding: .62rem .74rem;
        }
        .cl8us-interval-box span { color: var(--cl8us-muted); font-size: .78rem; }
        .cl8us-version-rule {
            border-top: 1px solid rgba(18, 59, 99, .19);
            margin: 2.15rem 0 1.05rem;
        }
        .cl8us-page-header {
            background: rgba(255, 252, 247, .96);
            border: 1px solid rgba(156, 132, 92, .28);
            border-radius: 0 0 13px 13px;
            box-shadow: 0 10px 24px rgba(83, 68, 42, .08);
            margin: -1.45rem 0 1.3rem;
            padding: .82rem 1.08rem 1rem;
        }
        .cl8us-page-brand {
            color: var(--cl8us-navy);
            display: flex;
            flex-wrap: wrap;
            font-family: Georgia, "Times New Roman", serif;
            font-size: .95rem;
            font-weight: 800;
            line-height: 1;
            margin-bottom: .9rem;
        }
        .cl8us-page-brand span { font-family: "Aptos", "Segoe UI", sans-serif; font-weight: 800; margin-left: .22rem; }
        .cl8us-page-brand small {
            color: #6B6256;
            flex-basis: 100%;
            font-family: "Aptos", "Segoe UI", sans-serif;
            font-size: .59rem;
            font-weight: 500;
            margin-top: .18rem;
        }
        .cl8us-page-header h1 { font-size: 1.47rem; margin: 0 0 .48rem; }
        .cl8us-page-header p { color: #625B51; font-size: .84rem; margin: 0 0 .62rem; }
        .cl8us-page-privacy {
            background: #FFF5D8;
            border: 1px solid #D7A83B;
            border-radius: 999px;
            color: #8A5B0A;
            display: inline-block;
            font-size: .7rem;
            font-weight: 700;
            padding: .27rem .58rem;
        }
        .cl8us-docs-note {
            background: rgba(255, 252, 247, .9);
            border: 1px solid rgba(156, 132, 92, .28);
            border-left: 4px solid var(--cl8us-navy);
            border-radius: 0 8px 8px 0;
            color: var(--cl8us-navy);
            font-size: .8rem;
            margin: 0 0 .72rem;
            padding: .58rem .72rem;
        }
        .cl8us-docs-card-marker { display: none; }
        .cl8us-docs-card-title {
            color: var(--cl8us-navy);
            font-size: 1.12rem;
            font-weight: 800;
            margin: 0 0 .32rem;
        }
        [data-testid="stVerticalBlockBorderWrapper"]:has(.cl8us-docs-card-marker) {
            box-shadow: none;
            margin-bottom: .72rem;
        }
        .cl8us-hero {
            border: 1px solid var(--cl8us-line);
            border-top: 4px solid var(--cl8us-navy);
            border-radius: 15px;
            background: linear-gradient(135deg, rgba(249,252,254,.96) 0%, rgba(237,244,248,.93) 100%);
            padding: 1.12rem 1.3rem 1.08rem;
            margin-bottom: 1.05rem;
            box-shadow: 0 13px 30px rgba(18, 59, 99, .10);
        }
        .cl8us-hero h1 { margin: 0 0 .28rem 0; font-size: 1.78rem; }
        .cl8us-hero p { margin: 0; color: var(--cl8us-muted); font-size: .95rem; }
        .cl8us-kicker {
            color: var(--cl8us-action);
            font-size: .72rem;
            font-weight: 850;
            letter-spacing: .07em;
            margin-bottom: .38rem;
            text-transform: uppercase;
        }
        [data-testid="stVerticalBlockBorderWrapper"] {
            background: rgba(255, 252, 247, .94);
            border-color: var(--cl8us-line) !important;
            border-radius: 13px;
            box-shadow: 0 10px 24px rgba(18, 59, 99, .075);
        }
        [data-testid="stVerticalBlockBorderWrapper"]:has(.cl8us-index-marker) {
            background: var(--cl8us-index) !important;
            border-color: var(--cl8us-index-line) !important;
            box-shadow: 0 8px 18px rgba(122, 96, 25, .08);
        }
        .cl8us-index-marker { display: none; }
        .cl8us-index-title {
            color: var(--cl8us-navy);
            font-size: .98rem;
            font-weight: 800;
            margin: .08rem 0 .55rem;
        }
        /* Bolinha a DIREITA do titulo "Indice do contrato": circulo PRETO e
           PREENCHIDO, no mesmo peso visual dos marcadores solidos dos campos do
           formulario (nao o circulo contornado do menu lateral).
           O marker tecnico .cl8us-index-marker segue invisivel para o :has(). */
        .cl8us-index-title::after {
            content: "";
            display: inline-block;
            width: .72rem;
            height: .72rem;
            margin-left: .55rem;
            border-radius: 50%;
            background: #1A1A1A;
            vertical-align: middle;
        }
        [data-baseweb="select"] > div,
        [data-testid="stDateInput"] input,
        [data-testid="stTextInput"] input {
            background: var(--cl8us-surface) !important;
            border-color: rgba(18, 59, 99, .20) !important;
            border-radius: 9px !important;
            color: var(--cl8us-navy) !important;
        }
        [data-testid="stSidebar"] [data-baseweb="select"] > div,
        [data-testid="stSidebar"] [data-testid="stDateInput"] input,
        [data-testid="stSidebar"] [data-testid="stTextInput"] input {
            background: var(--cl8us-input) !important;
        }
        [data-baseweb="select"] > div:focus-within,
        [data-testid="stDateInput"] input:focus,
        [data-testid="stTextInput"] input:focus {
            border-color: var(--cl8us-action) !important;
            box-shadow: 0 0 0 1px var(--cl8us-action) !important;
        }
        [data-testid="stSidebar"] [data-testid="stRadio"] label {
            border-left: 3px solid transparent;
            border-radius: 0 8px 8px 0;
            padding: .24rem .38rem;
        }
        [data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked) {
            background: rgba(255, 255, 255, .40);
            border-left-color: var(--cl8us-navy);
            font-weight: 700;
        }
        .home-card {
            min-height: 132px;
            padding: .18rem .12rem .45rem;
        }
        .home-card-tag {
            color: var(--cl8us-action);
            font-size: .71rem;
            font-weight: 850;
            letter-spacing: .055em;
            margin-bottom: .42rem;
            text-transform: uppercase;
        }
        .home-card strong { color: var(--cl8us-navy); font-size: 1.06rem; }
        .home-card p { color: var(--cl8us-muted); font-size: .89rem; line-height: 1.35; margin: .38rem 0 0; }

        .stButton > button, .stDownloadButton > button {
            border-radius: 8px;
            border-color: rgba(18, 59, 99, .34);
            background: rgba(249, 252, 254, .94);
            color: var(--cl8us-navy);
            font-weight: 700;
            min-height: 2.52rem;
            box-shadow: 0 4px 10px rgba(18, 59, 99, .07);
            transition: transform .14s ease, box-shadow .14s ease, background-color .14s ease;
        }
        .stButton > button:not([kind="primary"]):hover,
        .stDownloadButton > button:not([kind="primary"]):hover {
            border-color: var(--cl8us-navy);
            color: var(--cl8us-navy);
            background: #FFFFFF;
            transform: translateY(-1px);
        }
        .stButton > button[kind="primary"],
        .stDownloadButton > button[kind="primary"] {
            color: #FFFFFF !important;
            background: var(--cl8us-action) !important;
            border-color: var(--cl8us-action) !important;
            box-shadow: 0 6px 14px rgba(122, 23, 51, .19);
        }
        .stButton > button[kind="primary"]:hover,
        .stDownloadButton > button[kind="primary"]:hover {
            color: #FFFFFF !important;
            background: var(--cl8us-action-hover) !important;
            border-color: var(--cl8us-action-hover) !important;
            transform: translateY(-1px);
        }
        [data-testid="stMetric"], [data-testid="stFileUploaderDropzone"], [data-testid="stExpander"] {
            background: rgba(249, 252, 254, .78);
            border-color: var(--cl8us-line);
            border-radius: 12px;
        }
        @media (max-width: 800px) {
            .block-container { padding-top: .9rem; }
            .cl8us-hero h1 { font-size: 1.48rem; }
            .home-card { min-height: auto; }
        }
        /* Dia selecionado no calendário: o BaseWeb desenha o círculo no ::after do
           [role="gridcell"] com aria-label contendo "Selected". Clareamos esse
           círculo (bordô a 20%) e mantemos o número legível na cor da marca. */
        [data-baseweb="calendar"] [role="gridcell"][aria-label*="Selected"]::after {
            background-color: rgba(122, 23, 51, 0.20) !important;
        }
        [data-baseweb="calendar"] [role="gridcell"][aria-label*="Selected"],
        [data-baseweb="calendar"] [role="gridcell"][aria-label*="Selected"] * {
            color: var(--cl8us-action) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _render_sidebar() -> None:
    with st.sidebar:
        st.markdown(
            '<div class="cl8us-side-brand"><strong><span>·</span>cl8us<span>·</span></strong></div>'
            '<div class="cl8us-side-caption">apoio à gestão de contratos</div>',
            unsafe_allow_html=True,
        )
        st.page_link(PAGINA_INICIO, label="Início")
        st.page_link(PAGINA_UM_CICLO, label="Calculadora 1 ciclo")
        st.page_link(PAGINA_MULTICICLO, label="Calculadora multiciclo")
        st.markdown('<div class="cl8us-side-group">Documentos</div>', unsafe_allow_html=True)
        st.page_link(PAGINA_UPLOAD, label="Upload e docs")

        render_versao_sidebar()


_render_css()
render_cl8us_light_theme()
_render_sidebar()

# Todas as páginas continuam registradas para manter links e funcionalidades.
# position="hidden" retira apenas o menu automático e deixa a navegação própria
# acima como a interface operacional principal.
pagina_atual = st.navigation(
    [*PAGINAS_PRINCIPAIS, *PAGINAS_AUXILIARES],
    position="hidden",
)
pagina_atual.run()
