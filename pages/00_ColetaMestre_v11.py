import streamlit as st

try:
    from _ui_utils import render_marca_topo
    render_marca_topo(titulo_pagina="Matriz 2.0", subtitulo_pagina="Coleta Analitica Experimental")
except Exception:
    st.title("Matriz 2.0 - Coleta Analitica")

adm = st.session_state.get("dados_admissibilidade")
if not adm:
    st.warning("Faca um calculo na Calculadora de Reajustes primeiro para pre-preencher a planilha.")
    st.stop()

st.markdown(
    "Nova estrutura de coleta com 5 abas: **BASE** (automatica) · "
    "**EXECUCAO_FINANCEIRA** · **HISTORICO_CICLOS** · "
    "**ITENS_CONTRATADOS** · **ADITIVOS**."
)

try:
    from _matriz_2_0_gerador import gerar_matriz_2_0
    _bytes = gerar_matriz_2_0()
    st.download_button(
        label="Baixar Matriz 2.0 (experimental)",
        data=_bytes,
        file_name="Matriz_2_0_cl8us.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    st.caption(
        "Fiscal preenche: EXECUCAO_FINANCEIRA (A:C) e ITENS_CONTRATADOS (B,C,E:J). "
        "GCC preenche: HISTORICO_CICLOS (C:F) e ADITIVOS. "
        "BASE e colunas cinzas sao automaticas."
    )
except Exception as e:
    st.error(f"Erro ao gerar Matriz 2.0: {e}")
