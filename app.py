import streamlit as st

# Mapeamento: Arquivo Restaurado -> Nome na Interface
p1 = st.Page("pages/01_Calculo_Simples.py", title="Reajuste Simples", icon="⚖️", default=True)
p2 = st.Page("pages/02_Calculo_Represados.py", title="Reajustes Múltiplos", icon="🔄")
p3 = st.Page("pages/03_Valor_Global.py", title="Valor Global", icon="💰")
p4 = st.Page("pages/04_Relatorio_Global.py", title="Relatório Global", icon="📊")

# Configuração da Navegação
pg = st.navigation([p1, p2, p3, p4])
pg.run()