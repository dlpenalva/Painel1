"""Indicador exclusivamente visual dos dias de preclusao."""

from html import escape
from textwrap import dedent

import streamlit as st

from _reajuste_utils import dias_de_preclusao


def html_indicador_preclusao(dias, ciclo=None):
    """HTML compacto e acessivel; a barra e semantica, nao proporcional."""
    if not isinstance(dias, int) or isinstance(dias, bool) or dias <= 0:
        return ""
    unidade = "dia" if dias == 1 else "dias"
    prefixo = f"{escape(str(ciclo))} · " if ciclo else ""
    texto = f"{prefixo}{dias} {unidade} de preclusão"
    return dedent(f"""
    <style>
      .cl8us-preclusao {{
        width: min(15rem, 100%);
        margin: .28rem 0 .48rem;
        color: #7f1d1d;
        font-size: .78rem;
        font-weight: 650;
        line-height: 1.2;
      }}
      .cl8us-preclusao-barra {{
        display: grid;
        grid-template-columns: minmax(5.5rem, 3fr) 1.5rem;
        align-items: center;
        width: 100%;
        height: .32rem;
        margin-top: .28rem;
      }}
      .cl8us-preclusao-regular {{
        height: .22rem;
        border-radius: 999px 0 0 999px;
        background: #3f8f68;
      }}
      .cl8us-preclusao-atraso {{
        position: relative;
        height: .22rem;
        border-left: 2px solid #334155;
        border-radius: 0 999px 999px 0;
        background: #c94b4b;
      }}
      .cl8us-preclusao-atraso::after {{
        content: "";
        position: absolute;
        right: -.02rem;
        top: 50%;
        width: .48rem;
        height: .48rem;
        border: 2px solid #fff;
        border-radius: 50%;
        background: #a61b1b;
        box-shadow: 0 0 0 1px #7f1d1d;
        transform: translate(35%, -50%);
      }}
    </style>
    <div class="cl8us-preclusao" role="note" aria-label="{texto}">
      <span>{texto}</span>
      <span class="cl8us-preclusao-barra" aria-hidden="true">
        <span class="cl8us-preclusao-regular"></span>
        <span class="cl8us-preclusao-atraso"></span>
      </span>
    </div>
    """).strip().replace("</style>\n<div", "</style><div")


def render_indicador_preclusao(situacao, data_pedido, data_limite, ciclo=None):
    """Calcula em runtime e renderiza somente a classificacao PRECLUSO."""
    dias = dias_de_preclusao(situacao, data_pedido, data_limite)
    html = html_indicador_preclusao(dias, ciclo=ciclo)
    if html:
        st.markdown(html, unsafe_allow_html=True)
