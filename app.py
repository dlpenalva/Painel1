import streamlit as st

st.set_page_config(page_title="FARC - Telebras", layout="wide")

p1 = st.Page(
    "pages/01_Calculo_Simples.py",
    title="Reajuste Simples",
    icon="⚖️",
    default=True
)

p2 = st.Page(
    "pages/02_Calculo_Represados.py",
    title="Reajustes Múltiplos",
    icon="🔄"
)

pg = st.navigation({
    "Admissibilidade e Cálculo": [p1, p2]
})

pg.run()