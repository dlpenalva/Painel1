import streamlit as st
import pandas as pd

st.title("💰 Valor Global do Contrato")

# O Bloco B lê o que foi feito no Bloco A
if 'dados_admissibilidade' in st.session_state:
    adm = st.session_state['dados_admissibilidade']
    st.write(f"**Parâmetros herdados:** Reajuste {adm['tipo']} | Fator: {adm['fator']:.4f}")
else:
    st.warning("⚠️ Admissibilidade não realizada no Bloco A.")

upload = st.file_uploader("Upload da Planilha de Itens (Excel)", type=["xlsx"])

if upload:
    df = pd.read_excel(upload)
    st.write("Visualização dos Itens:")
    st.dataframe(df.head())
    
    # Aqui entra sua lógica de cálculo global que você já possui
    st.success("Cálculo do Valor Global processado.")