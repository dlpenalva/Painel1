import streamlit as st

# Configuração da página
st.set_page_config(page_title="Sistema de Gestão de Contratos", layout="wide")

# Mapeamento para a pasta oculta .pages
p1 = st.Page(".pages/01_Calculo_Simples.py", title="Reajuste Simples", icon="⚖️", default=True)
p2 = st.Page(".pages/02_Calculo_Represados.py", title="Reajuste Múltiplo", icon="🔄")
p3 = st.Page(".pages/03_Valor_Global.py", title="Gestão de Valor Global", icon="💰")

# Inicializa a navegação
pg = st.navigation([p1, p2, p3])

# Executa a página
pg.run()