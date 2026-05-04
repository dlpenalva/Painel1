import streamlit as st

# Configuração da página
st.set_page_config(page_title="FARC - Telebras", layout="wide")

# Mapeamento apenas dos arquivos que você tem na pasta
p1 = st.Page("pages/01_Calculo_Simples.py", title="Cálculo Simples", icon="⚖️", default=True)
p2 = st.Page("pages/02_Calculo_Represados.py", title="Cálculo Represados", icon="🔄")

# Navegação sem o arquivo 03
pg = st.navigation({
    "Admissibilidade e Cálculo": [p1, p2]
})

pg.run()