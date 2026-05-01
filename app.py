import streamlit as st

st.set_page_config(page_title="GCC - Telebras", layout="wide")

# Estilo Global Telebras
st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #003366; }
    .subtitle-gcc { font-size: 14px; color: #666; margin-top: -20px; margin-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Gestão de Contratos - Reajustes")
st.markdown('<p class="subtitle-gcc">GCC - Painel de Controle</p>', unsafe_allow_html=True)

st.info("Utilize o menu lateral à esquerda para navegar entre as ferramentas de cálculo.")

st.markdown("""
### Bem-vindo ao GCC
Esta ferramenta foi desenvolvida para automatizar a análise de admissibilidade e o cálculo de reajustes contratuais da Telebras.

*   **Cálculo Individual:** Para processos com um único ciclo de reajuste.
*   **Cálculo de Passivos (Múltiplo):** Em desenvolvimento - focado em processos acumulados de vários anos.
""")