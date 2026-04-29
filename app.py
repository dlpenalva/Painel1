import streamlit as st
import pandas as pd
from bcb import sgs

# Configuração de interface
st.set_page_config(page_title="FARC - Telebras", layout="wide")

st.title("📊 FARC - Analisador de Reajuste")
st.markdown("### Monitoramento de Índices (IPCA, IGP-M e IST)")

# --- FUNÇÕES DE BUSCA ---

@st.cache_data(ttl=3600)
def buscar_indices_bcb(codigo):
    """Busca automática no Banco Central (SGS)"""
    try:
        df = sgs.get({'valor': codigo}, last=12)
        fator = (1 + df['valor']/100).prod()
        variacao_anual = (fator - 1) * 100
        return round(variacao_anual, 4), df
    except Exception as e:
        return None, f"Erro na conexão BCB: {e}"

@st.cache_data(ttl=3600)
def buscar_ist_automatico():
    """Lê o arquivo ist.csv e calcula a variação acumulada"""
    try:
        # Lê o arquivo CSV que você renomeou
        df = pd.read_csv('ist.csv') 
        
        # Validação simples: verifica se as colunas necessárias existem
        if 'valor' not in df.columns:
            return None, "Coluna 'valor' não encontrada no arquivo ist.csv"
            
        # Pega os últimos 12 meses registrados no arquivo
        ultimos_12 = df.tail(12)
        fator = (1 + ultimos_12['valor']/100).prod()
        variacao_anual = (fator - 1) * 100
        return round(variacao_anual, 4), ultimos_12
    except Exception as e:
        return None, f"Erro ao ler ist.csv: {e}"

# --- INTERFACE ---

with st.sidebar:
    st.header("Configurações")
    indice = st.selectbox("Escolha o Índice:", ["Selecione...", "IPCA", "IGP-M", "IST"])
    valor_base = st.number_input("Valor Atual do Contrato (R$):", min_value=0.0, format="%.2f")
    botao_executar = st.button("Confirmar e Calcular")

if indice == "Selecione...":
    st.info("Selecione um índice no menu lateral para iniciar a busca automática.")
elif botao_executar:
    variacao = None
    dados_grafico = None

    if indice == "IPCA":
        variacao, dados_grafico = buscar_indices_bcb(433)
    elif indice == "IGP-M":
        variacao, dados_grafico = buscar_indices_bcb(189)
    elif indice == "IST":
        variacao, dados_grafico = buscar_ist_automatico()

    if variacao is not None:
        col1, col2 = st.columns(2)
        col1.metric(f"Variação {indice} (12 meses)", f"{variacao}%")
        
        if valor_base > 0:
            novo_valor = valor_base * (1 + (variacao/100))
            col2.metric("Novo Valor Reajustado", f"R$ {novo_valor:,.2f}", delta=f"R$ {novo_valor-valor_base:,.2f}")
        
        st.divider()
        st.subheader("Memória de Cálculo (Últimos 12 meses)")
        st.dataframe(dados_grafico, use_container_width=True)
    else:
        # Exibe o erro específico que aconteceu (seja no BCB ou no CSV)
        st.error(f"Erro no processamento: {dados_grafico}")

st.caption("GCC - Gerência de Compras e Contratos | Telebras")