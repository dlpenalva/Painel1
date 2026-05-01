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

def get_ist_mock_api(data_inicio, data_fim):
    try:
        # Lógica para o novo formato do CSV (DD/MM/YYYY)
        df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
        df['data'] = pd.to_datetime(df['data'], format='%d/%m/%Y')
        df = df.sort_values('data')
        
        # O IST usa o índice do mês fechado anterior à data-base
        d_ini_ref = (data_inicio - relativedelta(months=1)).replace(day=1)
        d_fim_ref = data_fim.replace(day=1) 

        v_ini = df[df['data'] == d_ini_ref]['indice'].values[0]
        v_fim = df[df['data'] == d_fim_ref]['indice'].values[0]

        # Auditoria: gera variação para o período
        return pd.DataFrame({'data': [data_fim], 'valor': [(v_fim / v_ini) - 1], 'indice_ini': [v_ini], 'indice_fim': [v_fim]})
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo Simples")

col1, col2 = st.columns(2)
with col1:
    dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 10, 10), format="DD/MM/YYYY")
    dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
with col2:
    tipo_idx = st.selectbox("Índice:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Série Local)"])

dt_fim = dt_base + relativedelta(months=11)
dt_aniv = dt_base + relativedelta(years=1)
limite = dt_aniv + relativedelta(days=90)
status = "✅ ADMISSÍVEL" if dt_solic <= limite else "❌ PRECLUSO"

df_dados = get_ist_mock_api(dt_base, dt_fim) if "IST" in tipo_idx else get_index_data("433" if "IPCA" in tipo_idx else "189", dt_base.strftime('%d/%m/%Y'), dt_fim.strftime('%d/%m/%Y'))

if df_dados is not None:
    var = (1 + df_dados['valor']).prod() - 1
    st.metric("Variação do Período (12 meses)", f"{var*100:,.2f}%".replace('.', ','))
    st.info(f"**Janela de Admissibilidade:** {dt_aniv.strftime('%d/%m/%Y')} até {limite.strftime('%d/%m/%Y')}")

    st.subheader("Memória de Cálculo (Fator de Prova)")
    if "IST" in tipo_idx:
        st.write(f"**Cálculo IST:** ({df_dados['indice_fim'].iloc[0]} / {df_dados['indice_ini'].iloc[0]}) - 1")
    
    st.table(pd.DataFrame({
        "Etapa": ["Data-Base Início", "Mês de Referência Final", "Aniversário", "Limite (90 dias)", "Data do Pedido", "Status"],
        "Valor": [dt_base.strftime('%m/%Y'), dt_fim.strftime('%m/%Y'), dt_aniv.strftime('%d/%m/%Y'), limite.strftime('%d/%m/%Y'), dt_solic.strftime('%d/%m/%Y'), status]
    }))