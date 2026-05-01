import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Facilitador - Cálculo Simples", layout="wide")

# Estilo Telebras
st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #003366; }
    </style>
    """, unsafe_allow_html=True)

def get_index_data(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=15)
        df = pd.DataFrame(response.json())
        if df.empty: return None, "Sem dados"
        df['valor'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        return df, None
    except:
        return None, "Erro API"

def calc_ist_csv(dt_base, dt_aniv):
    try:
        df = pd.read_csv('ist.csv', sep=None, engine='python', decimal=',')
        df.columns = df.columns.str.replace('^\ufeff', '', regex=True)
        meses_map = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
        ref_base = f"{meses_map[dt_base.month]}/{str(dt_base.year)[2:]}"
        ref_aniv = f"{meses_map[dt_aniv.month]}/{str(dt_aniv.year)[2:]}"
        v_base = float(df[df['MES_ANO'] == ref_base]['INDICE_NIVEL'].values[0])
        v_aniv = float(df[df['MES_ANO'] == ref_aniv]['INDICE_NIVEL'].values[0])
        return (v_aniv / v_base) - 1, ref_base, ref_aniv, None
    except:
        return None, None, None, "Erro IST"

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo Simples")

col1, col2 = st.columns(2)
with col1:
    dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 5, 1), format="DD/MM/YYYY")
    dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
with col2:
    tipo_idx = st.selectbox("Índice:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Planilha CSV)"])

dt_aniv = dt_base + relativedelta(years=1)
dias_janela = (dt_solic - dt_aniv).days
status = "⚠️ Precluso" if dias_janela > 90 else "✅ Admissível"

st.subheader("Resultado da Análise")
if "IST" in tipo_idx:
    var, rb, ra, erro = calc_ist_csv(dt_base, dt_aniv)
else:
    cod = "433" if "IPCA" in tipo_idx else "189"
    df, erro = get_index_data(cod, dt_base.strftime('%d/%m/%Y'), dt_aniv.strftime('%d/%m/%Y'))
    var = (1 + df['valor']).prod() - 1 if df is not None else None

if var is not None:
    c1, c2 = st.columns(2)
    c1.metric("Variação do Período", f"{var*100:,.4f}%".replace('.', ','))
    c2.metric("Status Admissibilidade", status)