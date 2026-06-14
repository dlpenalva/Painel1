import runpy
from pathlib import Path

import streamlit as st

st.set_page_config(page_title="TLB · cl8us - Calculadora de Reajustes", layout="wide")

from _ui_utils import render_marca_topo


def executar_motor(nome_arquivo: str):
    caminho = Path(__file__).resolve().parent / nome_arquivo
    if not caminho.exists():
        st.error(f"Arquivo do motor não localizado: {nome_arquivo}")
        st.stop()

    st.session_state["_calculadora_reajustes_embedded"] = True
    try:
        runpy.run_path(str(caminho), run_name="__main__")
    finally:
        st.session_state["_calculadora_reajustes_embedded"] = False


def _obter_fluxo_query():
    try:
        fluxo = st.query_params.get("fluxo", "")
        if isinstance(fluxo, list):
            fluxo = fluxo[0] if fluxo else ""
        return str(fluxo).lower()
    except Exception:
        try:
            fluxo = st.experimental_get_query_params().get("fluxo", [""])
            return str(fluxo[0]).lower() if fluxo else ""
        except Exception:
            return ""


def selecionar_fluxo(tipo):
    st.session_state["calculadora_tipo_analise"] = tipo
    st.session_state["calculadora_mais_de_um_ciclo"] = tipo == "Múltiplos ciclos"


render_marca_topo(titulo_pagina="Calculadora de Reajustes", subtitulo_pagina="Cálculo e admissibilidade")
st.markdown(
    """
    <style>
    .calc-note {
        background: #F8FAFC;
        border: 1px solid #E5EAF0;
        border-radius: 12px;
        padding: .75rem .95rem;
        color: #334155;
        margin: .55rem 0 1rem 0;
    }
    .calc-card-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 1.15rem;
        margin: .35rem 0 1.05rem 0;
        align-items: stretch;
    }
    .calc-card-link {
        display: block;
        height: 100%;
        min-height: 182px;
        border-radius: 14px;
        padding: 1rem 1.05rem;
        text-decoration: none !important;
        box-sizing: border-box;
        transition: all .12s ease-in-out;
    }
    .calc-card-link:hover {
        transform: translateY(-1px);
        text-decoration: none !important;
    }
    .calc-card-unico {
        border: 1.6px solid #93C5FD;
        background: #EFF6FF;
        color: #1E3A8A;
    }
    .calc-card-multiplos {
        border: 1.6px solid #F6C35B;
        background: #FFF7E6;
        color: #7A4A00;
    }
    .calc-card-selected-unico {
        border: 2.2px solid #2563EB;
        box-shadow: 0 0 0 2px rgba(37, 99, 235, .10);
    }
    .calc-card-selected-multiplos {
        border: 2.2px solid #D97706;
        box-shadow: 0 0 0 2px rgba(217, 119, 6, .10);
    }
    .calc-card-title {
        font-weight: 800;
        font-size: 1.16rem;
        margin-bottom: .40rem;
    }
    .calc-card-text {
        font-size: .93rem;
        line-height: 1.42rem;
    }
    .calc-selected {
        font-weight: 800;
        font-size: .82rem;
        margin-top: .70rem;
    }

    .modo-guia-wrap {
        background: #FFFFFF;
        border: 1px solid #E5EAF0;
        border-radius: 16px;
        padding: 1rem 1.05rem 1.1rem 1.05rem;
        margin: .65rem 0 1.25rem 0;
        box-shadow: 0 1px 2px rgba(15, 23, 42, .04);
    }
    .modo-guia-title {
        font-size: 1.02rem;
        font-weight: 850;
        color: #0F172A;
        margin-bottom: .20rem;
    }
    .modo-guia-subtitle {
        font-size: .90rem;
        color: #475569;
        margin-bottom: .85rem;
        line-height: 1.38rem;
    }
    .modo-guia-grid {
        display: grid;
        grid-template-columns: repeat(3, minmax(0, 1fr));
        gap: .95rem;
        align-items: stretch;
    }
    .modo-guia-card {
        border-radius: 14px;
        padding: .92rem .95rem;
        min-height: 305px;
        box-sizing: border-box;
        border: 1.4px solid #E2E8F0;
    }
    .modo-card-padrao {
        background: #F8FAFC;
        border-color: #CBD5E1;
        color: #1E293B;
    }
    .modo-card-reduzido {
        background: #F5F0FF;
        border-color: #C4B5FD;
        color: #4C1D95;
    }
    .modo-card-consumo {
        background: #F6F3EE;
        border-color: #B8A58A;
        color: #3F4F35;
    }
    .modo-chip {
        display: inline-block;
        font-size: .72rem;
        font-weight: 800;
        letter-spacing: .02em;
        text-transform: uppercase;
        border-radius: 999px;
        padding: .18rem .55rem;
        margin-bottom: .45rem;
    }
    .modo-chip-padrao { background: #E2E8F0; color: #334155; }
    .modo-chip-reduzido { background: #EDE9FE; color: #5B21B6; }
    .modo-chip-consumo { background: #E7E0D4; color: #4E3B24; }
    .modo-card-title {
        font-size: .98rem;
        font-weight: 850;
        margin-bottom: .35rem;
        line-height: 1.25rem;
    }
    .modo-card-leio {
        font-size: .84rem;
        font-weight: 800;
        margin: .55rem 0 .25rem 0;
        color: inherit;
    }
    .modo-card-text {
        font-size: .84rem;
        line-height: 1.28rem;
        color: inherit;
        opacity: .95;
    }
    .mini-sheet {
        width: 100%;
        border-collapse: collapse;
        margin: .45rem 0 .55rem 0;
        font-size: .76rem;
        background: rgba(255,255,255,.78);
    }
    .mini-sheet th {
        text-align: left;
        padding: .28rem .35rem;
        font-weight: 850;
        border-bottom: 1px solid rgba(100, 116, 139, .28);
        white-space: nowrap;
    }
    .mini-sheet td {
        padding: .25rem .35rem;
        border-bottom: 1px solid rgba(100, 116, 139, .16);
        white-space: nowrap;
    }
    .mini-sheet td.num { text-align: right; }
    .modo-result {
        margin-top: .55rem;
        padding-top: .50rem;
        border-top: 1px solid rgba(100, 116, 139, .22);
        font-size: .80rem;
        line-height: 1.22rem;
        font-weight: 650;
    }
    @media (max-width: 1100px) {
        .modo-guia-grid { grid-template-columns: 1fr; }
        .modo-guia-card { min-height: auto; }
    }

    @media (max-width: 900px) {
        .calc-card-grid {
            grid-template-columns: 1fr;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

fluxo_query = _obter_fluxo_query()
if fluxo_query in ["unico", "único"]:
    selecionar_fluxo("Ciclo único")
elif fluxo_query in ["multiplos", "múltiplos", "multiplo", "múltiplo"]:
    selecionar_fluxo("Múltiplos ciclos")

tipo = st.session_state.get("calculadora_tipo_analise")

st.markdown(
    """
    <style>
    .cl8us-tipo-wrap {
        margin: 1.2rem 0 1.4rem 0;
        padding: 12px 16px 12px 18px;
        border-left: 4px solid #2563EB;
        background: #F8FAFC;
        border-radius: 0 8px 8px 0;
        display: flex;
        align-items: center;
        gap: 6px;
        font-size: 1.02rem;
        color: #1E293B;
        font-weight: 500;
    }
    .cl8us-tipo-wrap a {
        color: #1E3A8A;
        font-weight: 600;
        text-decoration: none;
        padding: 3px 10px;
        border-radius: 5px;
        border: 1.5px solid transparent;
        transition: all 0.15s;
    }
    .cl8us-tipo-wrap a:hover {
        background: #EFF6FF;
        border-color: #93C5FD;
    }
    .cl8us-tipo-sel {
        background: #1E3A8A !important;
        color: #FFFFFF !important;
        border-color: #1E3A8A !important;
        font-weight: 700 !important;
    }
    .cl8us-tipo-sep { color: #94A3B8; margin: 0 2px; font-weight: 400; }
    .cl8us-tipo-hint { color: #2563EB; font-size: 0.85rem; margin-left: 4px; }
    </style>
    """,
    unsafe_allow_html=True,
)

_unico_cls = "cl8us-tipo-sel" if tipo == "Ciclo único"      else ""
_mult_cls  = "cl8us-tipo-sel" if tipo == "Múltiplos ciclos" else ""
_hint      = '<span class="cl8us-tipo-hint">← selecione</span>' if not tipo else ""

st.markdown(
    f"""
    <div class="cl8us-tipo-wrap">
        A análise envolve
        <a href="?fluxo=unico"     target="_self" class="{_unico_cls}">ciclo único</a>
        <span class="cl8us-tipo-sep">ou</span>
        <a href="?fluxo=multiplos" target="_self" class="{_mult_cls}">múltiplos ciclos</a>?
        {_hint}
    </div>
    """,
    unsafe_allow_html=True,
)

if not tipo:
    st.stop()


st.divider()

if tipo == "Ciclo único":
    executar_motor("01_Calculo_Simples.py")
else:
    executar_motor("02_Calculo_Represados.py")


