import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="GCC - Cálculo de Passivos", layout="wide")

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("🧮 Cálculo de Passivos (Múltiplos Ciclos)")

st.warning("Módulo em desenvolvimento: Focado em processos parados com múltiplos anos acumulados.")

# Interface inicial para debate
with st.expander("Configuração de Ciclos", expanded=True):
    dt_base_original = st.date_input("Data-Base Original (Proposta):", format="DD/MM/YYYY")
    qtd_ciclos = st.number_input("Quantidade de ciclos acumulados:", min_value=2, max_value=5, value=2)
    valor_inicial = st.number_input("Valor Inicial do Contrato (R$):", min_value=0.0)

st.write("---")
st.info("A próxima etapa será a criação da tabela de efeitos financeiros retroativos.")
