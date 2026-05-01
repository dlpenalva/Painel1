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
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
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

# Lógica Corrigida: Mês 0 + 11 meses (Total 12 meses)
dt_fim_calculo = dt_base + relativedelta(months=11)
dt_aniv_contratual = dt_base + relativedelta(years=1)
limite_90 = dt_aniv_contratual + relativedelta(days=90)
status = "✅ ADMISSÍVEL" if dt_solic <= limite_90 else "❌ PRECLUSO"

st.subheader("Resultado da Análise")
cod = "433" if "IPCA" in tipo_idx else "189"
# Busca os 12 meses exatos
df_dados = get_index_data(cod, dt_base.strftime('%d/%m/%Y'), dt_fim_calculo.strftime('%d/%m/%Y'))

if df_dados is not None:
    variacao = (1 + df_dados['valor']).prod() - 1
    
    c1, c2 = st.columns(2)
    c1.metric("Variação do Período (12 meses)", f"{variacao*100:,.2f}%".replace('.', ','))
    c2.metric("Status da Solicitação", status)

    st.subheader("Memória de Cálculo (Fator de Prova)")
    
    # Tabela de Resumo Executivo
    resumo_prova = {
        "Descrição": ["Início do Ciclo", "Fim do Ciclo (12º mês)", "Aniversário Contratual", "Limite Admissibilidade", "Data do Pedido", "Status"],
        "Data / Valor": [
            dt_base.strftime('%m/%Y'), 
            dt_fim_calculo.strftime('%m/%Y'), 
            dt_aniv_contratual.strftime('%d/%m/%Y'), 
            limite_90.strftime('%d/%m/%Y'), 
            dt_solic.strftime('%d/%m/%Y'),
            status
        ]
    }
    st.table(pd.DataFrame(resumo_prova))

    with st.expander("Visualizar Detalhamento Mensal (Índices Utilizados)"):
        df_display = df_dados.copy()
        df_display['Variação Mensal'] = df_display['valor'].map(lambda x: f"{x*100:.4f}%")
        st.dataframe(df_display[['data', 'Variação Mensal']], use_container_width=True)