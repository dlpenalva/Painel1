import streamlit as st
import pandas as pd

st.set_page_config(page_title="Reajuste Simples", layout="wide")

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("⚖️ Admissibilidade: Reajuste Simples")

# --- ENTRADA DE DADOS ---
st.header("1. Parâmetros do Contrato")
with st.container(border=True):
    col1, col2 = st.columns(2)
    with col1:
        indice_nome = st.selectbox("Selecione o Índice:", ["IST", "IPCA", "IGP-M"])
        data_base = st.date_input("Data-Base Original:")
    with col2:
        fator_reajuste = st.number_input("Fator de Reajuste Calculado (ex: 1.0468):", format="%.4f", value=1.0000)
        ciclo = st.text_input("Ciclo de Reajuste (ex: C1, C2):", value="C1")

# --- BOTÃO DE HOMOLOGAÇÃO ---
if st.button("✅ Homologar Admissibilidade"):
    if fator_reajuste > 1:
        # SALVAMENTO NA MEMÓRIA (Session State) para o Arquivo 03 ler
        st.session_state['dados_admissibilidade'] = {
            'indice': indice_nome,
            'data_base': data_base.strftime('%d/%m/%Y'),
            'fator': fator_reajuste,
            'ciclo_atual': ciclo
        }
        
        st.success(f"Admissibilidade do {ciclo} homologada com sucesso!")
        st.balloons()
        
        st.info("Os dados já estão disponíveis na aba 'Gestão de Valor Global'.")
    else:
        st.warning("Verifique o fator de reajuste antes de homologar.")

# Exibição dos dados salvos para conferência
if 'dados_admissibilidade' in st.session_state:
    st.divider()
    st.subheader("Dados em Memória:")
    st.write(st.session_state['dados_admissibilidade'])