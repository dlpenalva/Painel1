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

def get_ist_local(data_inicio, data_fim):
    try:
        # Leitura técnica: decimal com vírgula e separador ponto e vírgula
        df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
        df.columns = [str(col).strip().lower() for col in df.columns]
        
        # Converte data do CSV (DD/MM/YYYY)
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        
        # Regra de negócio: Mês anterior à base e Mês do aniversário (mês 12)
        ref_ini = (data_inicio - relativedelta(months=1)).replace(day=1)
        ref_fim = data_fim.replace(day=1)
        
        # Busca por período mensal para evitar erros de dia
        v_ini = df[df['data'].dt.to_period('M') == ref_ini.strftime('%Y-%m')]['indice'].values[0]
        v_fim = df[df['data'].dt.to_period('M') == ref_fim.strftime('%Y-%m')]['indice'].values[0]
        
        return pd.DataFrame({
            'data': [data_fim], 
            'valor': [(v_fim / v_ini) - 1],
            'audit_ini': [v_ini],
            'audit_fim': [v_fim]
        })
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo Simples")

col1, col2 = st.columns(2)
with col1:
    dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 10, 10), format="DD/MM/YYYY")
    dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
with col2:
    tipo_idx = st.selectbox("Índice:", ["IPCA (433)", "IGP-M (189)", "IST (Série Local)"])

# Definições de prazos
dt_fim_ciclo = dt_base + relativedelta(months=11)
dt_aniv = dt_base + relativedelta(years=1)
dt_limite = dt_aniv + relativedelta(days=90)
status_ped = "✅ ADMISSÍVEL" if dt_solic <= dt_limite else "❌ PRECLUSO"

# Execução do Cálculo
if "IST" in tipo_idx:
    df_res = get_ist_local(dt_base, dt_fim_ciclo)
else:
    df_res = get_index_data("433" if "IPCA" in tipo_idx else "189", dt_base.strftime('%d/%m/%Y'), dt_fim_ciclo.strftime('%d/%m/%Y'))

if df_res is not None:
    var_perc = (1 + df_res['valor']).prod() - 1
    st.metric("Variação do Ciclo", f"{var_perc*100:,.4f}%".replace('.', ','))
    
    with st.expander("🔍 Memória de Cálculo e Auditoria"):
        st.write(f"**Janela Legal:** {dt_aniv.strftime('%d/%m/%Y')} a {dt_limite.strftime('%d/%m/%Y')}")
        st.write(f"**Status:** {status_ped}")
        if "IST" in tipo_idx:
            st.write(f"Índice Inicial ({dt_base.month-1 if dt_base.month>1 else 12}/{dt_base.year if dt_base.month>1 else dt_base.year-1}): **{df_res['audit_ini'].iloc[0]}**")
            st.write(f"Índice Final ({dt_fim_ciclo.month}/{dt_fim_ciclo.year}): **{df_res['audit_fim'].iloc[0]}**")
        st.dataframe(df_res)
else:
    st.error("Erro ao processar índice. Verifique a conexão ou o arquivo ist.csv.")