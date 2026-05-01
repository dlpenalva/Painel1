import streamlit as st

# Configuração da página para definir o título da aba e layout
st.set_page_config(page_title="Gestão de Contratos - Telebras", layout="wide")

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Sistema de Gestão de Contratos")
st.markdown("---")

st.subheader("Bem-vindo ao Facilitador de Cálculos")
st.write("Utilize o menu lateral para acessar os módulos de **Cálculo Simples** ou **Cálculo de Represados**.")