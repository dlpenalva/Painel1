import streamlit as st

st.set_page_config(page_title="Facilitador - Telebras", layout="wide")

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Facilitador de Reajustes Contratuais")

st.info("Utilize o menu lateral à esquerda para navegar entre as ferramentas.")

st.markdown("""
### Painel de Gestão
Este ambiente centraliza ferramentas de apoio à gestão de contratos.

*   **📊 Cálculo Único:** Análise de admissibilidade e cálculo de reajuste para um ciclo específico (IPCA, IGP-M ou IST).
*   **🧮 Cálculo Múltiplo:** Ferramenta para processos acumulados (múltiplos anos) com geração de memória do retroativo (IPCA, IGP-M ou IST).
""")