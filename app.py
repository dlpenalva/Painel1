import streamlit as st

# Configuração global da página (deve ser o primeiro comando Streamlit)
st.set_page_config(
    page_title="Gestão de Contratos - Telebras",
    page_icon="https://www.telebras.com.br/wp-content/uploads/2019/06/favicon.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Definição da Navegação
# Importante: Os caminhos devem apontar para a pasta 'pages/' onde estão seus scripts
try:
    pg = st.navigation([
        st.Page(
            "pages/01_Calculo_Simples.py", 
            title="Reajuste Simples", 
            icon=":material/calculate:",
            default=True
        ),
        st.Page(
            "pages/02_Calculo_Represados_2.py", 
            title="Reajustes Múltiplos", 
            icon=":material/history:"
        )
    ])

    # Execução da navegação
    pg.run()

except Exception as e:
    st.error(f"Erro ao carregar a estrutura de navegação: {e}")
    st.info("Verifique se os arquivos 01_Calculo_Simples.py e 02_Calculo_Represados_2.py estão dentro da pasta 'pages/'.")
