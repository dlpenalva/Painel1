import streamlit as st

# Configuração da página deve ser a primeira instrução Streamlit
st.set_page_config(page_title="FARC - Telebras", layout="wide", initial_sidebar_state="expanded")

# Mapeamento Seguro: Arquivo Físico -> Nome na Interface
p1 = st.Page("pages/01_Calculo_Simples.py", title="Reajuste Simples", icon="⚖️", default=True)
p2 = st.Page("pages/02_Calculo_Represados.py", title="Reajustes Múltiplos", icon="🔄")
p3 = st.Page("pages/03_Valor_Global.py", title="Valor Global", icon="💰")
p4 = st.Page("pages/04_Relatorio_Global.py", title="Relatório Global", icon="📊")

# Configuração da Navegação
pg = st.navigation({
    "Admissibilidade e Cálculo": [p1, p2],
    "Execução Financeira": [p3],
    "Resultados": [p4]
})

pg.run()