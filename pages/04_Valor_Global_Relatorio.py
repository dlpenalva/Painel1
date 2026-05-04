import streamlit as st

st.title("📊 Relatório Global")

if 'dados_admissibilidade' not in st.session_state:
    st.info("Aguardando dados da admissibilidade...")
else:
    dados = st.session_state['dados_admissibilidade']
    st.subheader(f"Resumo da Análise - {dados['tipo']}")
    
    st.write(f"**Índice:** {dados['indice']}")
    st.write(f"**Fator Final:** {dados['fator']:.4f}")
    
    if dados['tipo'] == 'Múltiplo':
        st.write("**Ciclos Calculados:**")
        st.table(dados['detalhamento_ciclos'])