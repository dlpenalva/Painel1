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
        df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
        df.columns = ['mes_ano', 'indice']
        df['mes_ano_clean'] = df['mes_ano'].astype(str).str.strip().str.lower()
        
        d_ref_inicio = (data_inicio - relativedelta(months=1))
        d_ref_fim = data_fim 
        
        meses_br = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 
                    7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
        
        str_inicio = f"{meses_br[d_ref_inicio.month]}/{str(d_ref_inicio.year)[2:]}"
        str_fim = f"{meses_br[d_ref_fim.month]}/{str(d_ref_fim.year)[2:]}"

        v_ini = df[df['mes_ano_clean'] == str_inicio]['indice'].values[0]
        v_fim = df[df['mes_ano_clean'] == str_fim]['indice'].values[0]

        variacao_total = (v_fim / v_ini) - 1
        return pd.DataFrame({'data': [data_fim], 'valor': [variacao_total]})
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo Simples")

col1, col2 = st.columns(2)
with col1:
    dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 5, 1), format="DD/MM/YYYY")
    dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
with col2:
    tipo_idx = st.selectbox("Índice:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Série Local)"])

dt_fim_calculo = dt_base + relativedelta(months=11)
dt_aniv_contratual = dt_base + relativedelta(years=1)
limite_90 = dt_aniv_contratual + relativedelta(days=90)
status = "✅ ADMISSÍVEL" if dt_solic <= limite_90 else "❌ PRECLUSO"

if "IST" in tipo_idx:
    df_dados = get_ist_mock_api(dt_base, dt_fim_calculo)
else:
    df_dados = get_index_data("433" if "IPCA" in tipo_idx else "189", dt_base.strftime('%d/%m/%Y'), dt_fim_calculo.strftime('%d/%m/%Y'))

if df_dados is not None:
    variacao = (1 + df_dados['valor']).prod() - 1
    c1, c2 = st.columns(2)
    c1.metric("Variação do Período (12 meses)", f"{variacao*100:,.2f}%".replace('.', ','))
    c2.metric("Status da Solicitação", status)
    st.info(f"**Janela de Admissibilidade:** {dt_aniv_contratual.strftime('%d/%m/%Y')} até {limite_90.strftime('%d/%m/%Y')}")