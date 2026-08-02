"""Central de Modelos e Ferramentas — acesso sem Coleta (Etapa 29B).

Área acessível sem upload de Planilha de Coleta. Oferece:
  - modelos EM BRANCO do Despacho Saneador e do Termo de Apostila
    (gerados pelos wrappers canônicos ``gerar_modelo_branco_*``, sem qualquer
    dado de session_state, Coleta ou apuração anterior);
  - ferramentas independentes comprovadas (Garantia, Aditivos 25%, Checklist,
    Informações Prévias, Cl8us Orienta);
  - direcionamento para "Upload e docs" quando o usuário quiser documentos
    preenchidos automaticamente com os dados da apuração.

Nenhum modelo em branco usa as chaves dos documentos automáticos
(``arquivo_despacho_saneador_docx`` / ``arquivo_termo_apostila_docx``): os
bytes são produzidos localmente e nunca gravados em session_state.
"""
import streamlit as st

from _ui_utils import render_cabecalho_pagina
from _templates_documentos import (
    gerar_modelo_branco_despacho,
    gerar_modelo_branco_termo,
)

st.set_page_config(
    page_icon="assets/cl8us_favicon_512.png",
    page_title="Análises de Reajustes - Modelos e Ferramentas",
    layout="wide",
)

MIME_DOCX = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)

AVISO_MODELOS = (
    "Os modelos desta área não utilizam dados de uma Coleta e não são "
    "preenchidos automaticamente. Revise e complete todos os campos destacados."
)


@st.cache_data(show_spinner=False)
def _bytes_modelo_despacho() -> bytes:
    return gerar_modelo_branco_despacho()


@st.cache_data(show_spinner=False)
def _bytes_modelo_termo() -> bytes:
    return gerar_modelo_branco_termo()


def _card(titulo: str, descricao: str) -> None:
    st.markdown(f"**{titulo}**")
    st.caption(descricao)


def _abrir(destino: str, *, origem_garantia: bool = False) -> None:
    if origem_garantia:
        st.session_state["origem_navegacao_garantia"] = "modelos_ferramentas"
    st.switch_page(destino)


if not st.session_state.get("_calculadora_reajustes_embedded", False):
    render_cabecalho_pagina(
        "Central de Modelos e Ferramentas",
        "Modelos em branco e ferramentas que podem ser usados sem carregar uma "
        "Planilha de Coleta.",
    )

# ---------------------------------------------------------------- Seção 1
st.subheader("Modelos em branco")
st.warning(AVISO_MODELOS)

col_a, col_b = st.columns(2, gap="large")
with col_a:
    with st.container(border=True):
        _card(
            "Modelo em branco — Despacho Saneador",
            "Estrutura do Despacho Saneador com campos destacados para "
            "preenchimento manual. Não afirma fatos nem valores.",
        )
        st.download_button(
            "Baixar modelo em branco — Despacho Saneador",
            data=_bytes_modelo_despacho(),
            file_name="Modelo_Em_Branco_Despacho_Saneador.docx",
            mime=MIME_DOCX,
            use_container_width=True,
            key="dl_modelo_branco_despacho",
        )
with col_b:
    with st.container(border=True):
        _card(
            "Modelo em branco — Termo de Apostila",
            "Estrutura do Termo de Apostila com campos destacados para "
            "preenchimento manual. Não afirma reajuste nem concordância.",
        )
        st.download_button(
            "Baixar modelo em branco — Termo de Apostila",
            data=_bytes_modelo_termo(),
            file_name="Modelo_Em_Branco_Termo_de_Apostila.docx",
            mime=MIME_DOCX,
            use_container_width=True,
            key="dl_modelo_branco_termo",
        )

# ---------------------------------------------------------------- Seção 2
st.subheader("Ferramentas independentes")

linha1 = st.columns(2, gap="large")
with linha1[0]:
    with st.container(border=True):
        _card("Garantia Contratual",
              "Cálculo da garantia (5%) e do endosso a partir de valores "
              "informados manualmente.")
        if st.button("Abrir Garantia Contratual", use_container_width=True,
                     key="abrir_garantia_modelos"):
            _abrir("pages/05_Garantia.py", origem_garantia=True)
with linha1[1]:
    with st.container(border=True):
        _card("Aditivos: 25%",
              "Avaliação do limite de 25% para aditivos e supressões, com "
              "eventos informados manualmente.")
        if st.button("Abrir Aditivos: 25%", use_container_width=True,
                     key="abrir_aditivos_modelos"):
            _abrir("pages/08_Avaliacao_Aditivos.py")

linha2 = st.columns(2, gap="large")
with linha2[0]:
    with st.container(border=True):
        _card("Checklist Processual",
              "Checklist manual de conferência do processo, sem dependência de "
              "apuração.")
        if st.button("Abrir Checklist Processual", use_container_width=True,
                     key="abrir_checklist_modelos"):
            _abrir("pages/07_Checklist_Processual.py")
with linha2[1]:
    with st.container(border=True):
        _card("Informações Prévias",
              "Registro manual das informações prévias do contrato.")
        if st.button("Abrir Informações Prévias", use_container_width=True,
                     key="abrir_infos_modelos"):
            _abrir("pages/09_Infos_Previas.py")

linha3 = st.columns(2, gap="large")
with linha3[0]:
    with st.container(border=True):
        _card("Cl8us Orienta",
              "Orientações sobre os modos de análise, validações e índices.")
        if st.button("Abrir Cl8us Orienta", use_container_width=True,
                     key="abrir_orienta_modelos"):
            _abrir("pages/11_Cl8us_Orienta.py")

# ---------------------------------------------------------------- Seção 3
st.subheader("Documentos com dados da apuração")
with st.container(border=True):
    _card("Gerar documentos com dados da apuração",
          "Use a Planilha de Coleta para gerar documentos preenchidos "
          "automaticamente com os resultados da análise.")
    if st.button("Ir para Upload e docs", use_container_width=True,
                 key="ir_upload_modelos"):
        _abrir("pages/03_Valor_Global.py")
