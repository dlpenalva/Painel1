"""Entrada do Master 2.0 com a casca operacional enxuta do Cl8us 3.0.

O XLS continua sendo o produto principal. A web concentra apenas quatro
movimentos: início, cálculo de um ciclo, cálculo multiciclo e retorno do XLS.
Os módulos legados permanecem registrados, porém fora do menu principal.
"""

from pathlib import Path

import streamlit as st

from _ui_utils import render_versao_sidebar


st.set_page_config(
    page_title="TLB · cl8us — Reajustes contratuais",
    page_icon="CL",
    layout="wide",
)

ROOT = Path(__file__).resolve().parent
PAGES_DIR = ROOT / "pages"


def _page(caminho: str, titulo: str, *, default: bool = False) -> st.Page:
    return st.Page(str(PAGES_DIR / caminho), title=titulo, default=default)


PAGINA_INICIO = _page("00_Calculadora_Reajustes.py", "Início", default=True)
PAGINA_UM_CICLO = _page("01_Calculo_Simples.py", "Calculadora 1 ciclo")
PAGINA_MULTICICLO = _page("02_Calculo_Represados.py", "Calculadora multiciclo")
PAGINA_UPLOAD = _page("03_Valor_Global.py", "Upload e resultados")

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
            --cl8us-bg-start: #D4E3EF;
            --cl8us-bg-end: #BED4E5;
            --cl8us-surface: #F9FCFE;
            --cl8us-surface-soft: #EDF4F8;
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
                linear-gradient(135deg, var(--cl8us-bg-start) 0%, var(--cl8us-bg-end) 100%);
        }
        [data-testid="stHeader"] { background: rgba(212, 227, 239, .78); }
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
            padding: .38rem .52rem;
            font-weight: 620;
            transition: background-color .16s ease, border-color .16s ease, transform .16s ease;
        }
        [data-testid="stSidebar"] [data-testid="stPageLink"] a:hover {
            background: rgba(255, 255, 255, .28);
            transform: translateX(1px);
        }
        [data-testid="stSidebar"] [data-testid="stPageLink"] a[aria-current="page"] {
            background: rgba(255, 255, 255, .48);
            border-left-color: var(--cl8us-navy);
        }
        .cl8us-side-brand { margin: .15rem 0 .12rem 0; line-height: 1; }
        .cl8us-side-brand strong { font-size: 1.18rem; letter-spacing: .01em; }
        .cl8us-side-brand span { color: var(--cl8us-action); padding: 0 .15rem; }
        .cl8us-side-caption { color: var(--cl8us-muted); font-size: .78rem; margin-bottom: 1.15rem; }
        .cl8us-side-group {
            color: var(--cl8us-muted);
            font-size: .72rem;
            font-weight: 800;
            letter-spacing: .06em;
            margin: .8rem 0 .28rem 0;
            text-transform: uppercase;
        }
        [data-testid="stSidebar"] [data-testid="stExpander"] {
            background: rgba(255, 255, 255, .20);
            border: 1px solid rgba(18, 59, 99, .12);
            box-shadow: none;
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
            background: rgba(249, 252, 254, .94);
            border-color: var(--cl8us-line) !important;
            border-radius: 13px;
            box-shadow: 0 10px 24px rgba(18, 59, 99, .075);
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
        </style>
        """,
        unsafe_allow_html=True,
    )


def _render_sidebar() -> None:
    with st.sidebar:
        st.markdown(
            '<div class="cl8us-side-brand"><strong>TLB<span>·</span>cl8us</strong></div>'
            '<div class="cl8us-side-caption">apoio à gestão de contratos</div>',
            unsafe_allow_html=True,
        )
        st.markdown('<div class="cl8us-side-group">Operação</div>', unsafe_allow_html=True)
        st.page_link(PAGINA_INICIO, label="Início")
        st.page_link(PAGINA_UM_CICLO, label="Calculadora 1 ciclo")
        st.page_link(PAGINA_MULTICICLO, label="Calculadora multiciclo")
        st.markdown('<div class="cl8us-side-group">XLS preenchido</div>', unsafe_allow_html=True)
        st.page_link(PAGINA_UPLOAD, label="Upload e resultados")

        if PAGINAS_AUXILIARES:
            st.divider()
            with st.expander("Ferramentas complementares", expanded=False):
                for pagina in PAGINAS_AUXILIARES:
                    st.page_link(pagina, label=pagina.title)
        render_versao_sidebar()


_render_css()
_render_sidebar()

# Todas as páginas continuam registradas para manter links e funcionalidades.
# position="hidden" retira apenas o menu automático e deixa a navegação própria
# acima como a interface operacional principal.
pagina_atual = st.navigation(
    [*PAGINAS_PRINCIPAIS, *PAGINAS_AUXILIARES],
    position="hidden",
)
pagina_atual.run()
