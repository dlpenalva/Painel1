import streamlit as st

st.title("📊 Relatório Global")

if 'dados_admissibilidade' not in st.session_state:
    st.error("Não há dados para gerar o relatório. Por favor, complete as etapas anteriores.")
else:
    adm = st.session_state['dados_admissibilidade']
    st.write("### Resumo do Reajuste")
    
    col1, col2, col3 = st.columns(3)
    col1.metric("Tipo de Análise", adm['tipo'])
    col2.metric("Índice Aplicado", adm['indice'])
    col3.metric("Fator de Reajuste", f"{adm['fator']:.4f}")
    
    st.divider()
    st.subheader("Minuta para Nota Técnica")
    texto_sei = f"Conforme análise de admissibilidade, o contrato fará jus ao reajuste pelo índice {adm['indice']}..."
    st.text_area("Copie para o SEI:", texto_sei, height=150)