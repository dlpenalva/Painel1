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
        df['valor'] = df['valor'].astype(float)
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        # Para IPCA/IGPM: índice inicial é o da data-base (mês 0)
        # índice final é o do mês 11 (fechando 12 meses)
        v_ini = df.iloc[0]['valor']
        v_fim = df.iloc[-1]['valor']
        return {
            'variacao': (v_fim / v_ini) - 1,
            'i_ini': v_ini, 'i_fim': v_fim,
            'd_ini': df.iloc[0]['data'], 'd_fim': df.iloc[-1]['data']
        }
    except: return None

def get_ist_local(data_inicio, data_fim):
    try:
        df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
        df.columns = [str(col).strip().lower() for col in df.columns]
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        ref_ini = (data_inicio - relativedelta(months=1)).replace(day=1)
        ref_fim = data_fim.replace(day=1)
        v_ini = df[df['data'].dt.to_period('M') == ref_ini.strftime('%Y-%m')]['indice'].values[0]
        v_fim = df[df['data'].dt.to_period('M') == ref_fim.strftime('%Y-%m')]['indice'].values[0]
        return {
            'variacao': (v_fim / v_ini) - 1,
            'i_ini': v_ini, 'i_fim': v_fim,
            'd_ini': ref_ini, 'd_fim': ref_fim
        }
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo Simples")

col1, col2 = st.columns(2)
with col1:
    dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 10, 10), format="DD/MM/YYYY")
    dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
with col2:
    tipo_idx = st.selectbox("Índice:", ["IPCA (433)", "IGP-M (189)", "IST (Série Local)"])

dt_fim_ap = dt_base + relativedelta(months=11)
dt_aniv = dt_base + relativedelta(years=1)
dt_limite = dt_aniv + relativedelta(days=90)
status_ped = "✅ ADMISSÍVEL" if dt_solic <= dt_limite else "❌ PRECLUSO"

res = get_ist_local(dt_base, dt_fim_ap) if "IST" in tipo_idx else get_index_data("433" if "IPCA" in tipo_idx else "189", dt_base.strftime('%d/%m/%Y'), dt_fim_ap.strftime('%d/%m/%Y'))

if res:
    st.metric("Variação Apurada", f"{res['variacao']*100:,.4f}%".replace('.', ','))
    
    with st.expander("🔍 Memória de Cálculo do Índice"):
        c_a, c_b = st.columns(2)
        with c_a:
            st.markdown(f"**Índice:** {tipo_idx}")
            st.write(f"**Data-Base:** {dt_base.strftime('%m/%Y')}")
            st.write(f"**Período Apuração:** {res['d_ini'].strftime('%m/%Y')} a {res['d_fim'].strftime('%m/%Y')}")
        with c_b:
            st.write(f"**Índice Inicial:** {res['i_ini']}")
            st.write(f"**Índice Final:** {res['i_fim']}")
            st.write(f"**Fórmula:** (Final / Inicial) - 1")
        
        st.code(f"({res['i_fim']} / {res['i_ini']}) - 1 = {res['variacao']:.6f}")
        st.info(f"**Resultado Final:** {res['variacao']*100:,.2f}%")
else:
    st.error("Erro na apuração dos dados.")