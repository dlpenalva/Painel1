import streamlit as st
import pandas as pd

st.title("⚖️ Reajuste Simples")

with st.expander("1. Parâmetros do Contrato", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        indice = st.selectbox("Selecione o Índice:", ["IST", "IPCA", "IGP-M"])
        data_base = st.date_input("Data-Base Anterior:")
    with col2:
        fator_manual = st.number_input("Fator de Reajuste Calculado (ex: 1.0401):", format="%.4f", step=0.0001)
        data_pedido = st.date_input("Data do Pedido:")

if st.button("Homologar Admissibilidade"):
    # Salva no session_state para o Bloco B ler depois
    st.session_state['dados_admissibilidade'] = {
        'tipo': 'Simples',
        'indice': indice,
        'fator': fator_manual,
        'data_base': data_base,
        'data_pedido': data_pedido,
        'ciclo_atual': 'C1'
    }
    st.success("Admissibilidade homologada com sucesso!")

# Exibição da Memória de Cálculo (Exemplo simplificado do que você restaurou)
if 'dados_admissibilidade' in st.session_state:
    st.divider()
    st.subheader("Memória de Cálculo")
    st.info(f"Variação Apurada: {(fator_manual - 1)*100:.2f}%")