import streamlit as st

st.set_page_config(page_title="Facilitador Telebras", layout="wide")

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Sistema de Gestão de Contratos")
st.markdown("---")

st.subheader("Selecione o Módulo no Menu Lateral")
st.write("1. **Cálculo Simples:** Para reajustes de ciclo único.")
st.write("2. **Cálculo de Represados:** Para processos com múltiplos anos acumulados.")

# No Streamlit Cloud, o nome no menu é definido pelo nome do arquivo ou st.set_page_config