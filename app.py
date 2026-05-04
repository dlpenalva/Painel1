import streamlit as st

# Exemplo de mapeamento que deve estar no seu app.py
p2 = st.Page("pages/02_Calculo_Represados.py", title="Reajustes Múltiplos", icon="🔄")
p3 = st.Page("pages/03_Valor_Global.py", title="Valor Global", icon="💰")

pg = st.navigation([p2, p3])
pg.run()