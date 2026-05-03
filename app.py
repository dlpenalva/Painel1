import streamlit as st

# Configuração de Navegação Segura
# Mapeia o arquivo físico para o rótulo exibido no menu
p1 = st.Page("pages/01_Calculo_Simples.py", title="Reajuste Simples", icon="📊")
p2 = st.Page("pages/02_Calculo_Represados.py", title="Reajuste Múltiplo", icon="🔄")
p3 = st.Page("pages/03_Valor_Global.py", title="Gestão de Valor Global", icon="💰")

# Inicializa a navegação
pg = st.navigation([p1, p2, p3])

# Executa a página selecionada
pg.run()