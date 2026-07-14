"""Página inicial do fluxo XLS-first do Master 2.0."""

from pathlib import Path

import streamlit as st

from _coleta_reajuste import CAMINHO_MODELO_COLETA, NOME_ARQUIVO_COLETA
from _ui_utils import render_cabecalho_pagina


st.set_page_config(
    page_title="TLB · cl8us — Início",
    page_icon="CL",
    layout="wide",
)

MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@st.cache_data(show_spinner=False)
def _ler_modelo(caminho: str, alterado_em_ns: int) -> bytes:
    del alterado_em_ns  # participa da chave para invalidar o cache após nova versão
    return Path(caminho).read_bytes()


def _conteudo_card(tag: str, titulo: str, texto: str) -> None:
    st.markdown(
        f"""
        <div class="home-card">
            <div class="home-card-tag">{tag}</div>
            <strong>{titulo}</strong>
            <p>{texto}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


render_cabecalho_pagina(
    "Reajustes contratuais",
    "Baixe o XLS Coleta, registre as informações do processo e envie o arquivo preenchido para validação e geração progressiva dos documentos.",
)

st.subheader("Fluxo operacional")
st.caption(
    "O Coleta_Reajuste.xlsx é o produto principal. A web prepara a coleta e, no retorno, "
    "apresenta somente resultados sustentados pelos dados e pelas fórmulas do arquivo."
)

linha_1 = st.columns(2, gap="large")
with linha_1[0]:
    with st.container(border=True):
        _conteudo_card(
            "1 · Arquivo de trabalho",
            "Coleta_Reajuste.xlsx",
            "Baixe o modelo único com fórmulas. Ele atende tanto um ciclo quanto múltiplos ciclos.",
        )
        if CAMINHO_MODELO_COLETA.exists():
            st.download_button(
                "Baixar Coleta_Reajuste.xlsx",
                data=_ler_modelo(
                    str(CAMINHO_MODELO_COLETA),
                    CAMINHO_MODELO_COLETA.stat().st_mtime_ns,
                ),
                file_name=NOME_ARQUIVO_COLETA,
                mime=MIME_XLSX,
                type="primary",
                use_container_width=True,
                key="download_coleta_inicio",
            )
        else:
            st.error("O modelo Coleta_Reajuste.xlsx não foi localizado.")

with linha_1[1]:
    with st.container(border=True):
        _conteudo_card(
            "2 · Retorno do fiscal",
            "Upload e docs",
            "Envie o XLS preenchido. A web confere a estrutura e sinaliza lacunas antes de liberar totais ou documentos.",
        )
        if st.button("Abrir Upload e docs", use_container_width=True, key="abrir_upload_inicio"):
            st.switch_page("pages/03_Valor_Global.py")

st.subheader("Preparar os marcos da coleta")
linha_2 = st.columns(2, gap="large")
with linha_2[0]:
    with st.container(border=True):
        _conteudo_card(
            "3 · Análise simples",
            "Calculadora 1 ciclo",
            "Use quando apenas um ciclo é objeto da apuração. Ao final, baixe o mesmo Coleta_Reajuste.xlsx já parametrizado.",
        )
        if st.button("Abrir Calculadora 1 ciclo", use_container_width=True, key="abrir_um_ciclo_inicio"):
            st.switch_page("pages/01_Calculo_Simples.py")

with linha_2[1]:
    with st.container(border=True):
        _conteudo_card(
            "4 · Análise acumulada",
            "Calculadora multiciclo",
            "Use para dois ou mais ciclos e inicie a apuração em qualquer ciclo entre C1 e C4.",
        )
        if st.button("Abrir Calculadora multiciclo", use_container_width=True, key="abrir_multiciclo_inicio"):
            st.switch_page("pages/02_Calculo_Represados.py")
