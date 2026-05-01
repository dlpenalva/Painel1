import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo Simples", layout="wide")

def get_index_data(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=15)
        df = pd.DataFrame(response.json())
        if df.empty: return None
        df['valor'] = df['valor'].astype(float) / 100
        return df
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo Simples")

col1, col2 = st.columns(2)
with col1:
    dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 5, 1), format="DD/MM/YYYY")
    dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
with col2:
    tipo_idx = st.selectbox("Índice:", ["IPCA (Série 433)", "IGP-M (Série 189)"])

dt_aniv = dt_base + relativedelta(years=1)
limite_90 = dt_aniv + relativedelta(days=90)
status = "❌ PRECLUSO" if dt_solic > limite_90 else "✅ ADMISSÍVEL"

st.subheader("Resultado da Análise")
cod = "433" if "IPCA" in tipo_idx else "189"
df_dados = get_index_data(cod, dt_base.strftime('%d/%m/%Y'), dt_aniv.strftime('%d/%m/%Y'))

if df_dados is not None:
    variacao = (1 + df_dados['valor']).prod() - 1
    
    c1, c2 = st.columns(2)
    c1.metric("Variação do Período", f"{variacao*100:,.2f}%".replace('.', ','))
    c2.metric("Status da Solicitação", status)

    # MEMÓRIA DE CÁLCULO - CONDIÇÃO SINE QUA NON
    st.subheader("Memória de Cálculo (Fator de Prova)")
    memoria = {
        "Parâmetro": ["Data-Base Anterior", "Aniversário Contratual", "Data Limite (90 dias)", "Data do Pedido", "Índice Utilizado"],
        "Valor": [dt_base.strftime('%d/%m/%Y'), dt_aniv.strftime('%d/%m/%Y'), limite_90.strftime('%d/%m/%Y'), dt_solic.strftime('%d/%m/%Y'), tipo_idx]
    }
    st.table(pd.DataFrame(memoria))
    
    with st.expander("Ver detalhamento mensal do índice"):
        st.dataframe(df_dados)