import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo de Represados", layout="wide")

def get_index_data_rep(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=10)
        df = pd.DataFrame(response.json())
        if df.empty: return None
        df['valor'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        return df
    except: return None

def get_ist_rep(data_inicio, data_fim):
    try:
        df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
        df.columns = [str(col).strip().lower() for col in df.columns]
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        ref_ini = (data_inicio - relativedelta(months=1)).replace(day=1)
        ref_fim = data_fim.replace(day=1)
        v_ini = df[df['data'].dt.to_period('M') == ref_ini.strftime('%Y-%m')]['indice'].values[0]
        v_fim = df[df['data'].dt.to_period('M') == ref_fim.strftime('%Y-%m')]['indice'].values[0]
        return pd.DataFrame({'data': [data_fim], 'valor': [(v_fim / v_ini) - 1]})
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo de Represados")

with st.sidebar:
    dt_base_org = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    n_ciclos = st.number_input("Quantidade de Ciclos:", min_value=1, max_value=5, value=2)
    sel_idx = st.selectbox("Índice:", ["IPCA (433)", "IGP-M (189)", "IST (Série Local)"])

fator_acum = 1.0
data_corrente = dt_base_org
resumo = []

for i in range(1, int(n_ciclos) + 1):
    with st.container():
        st.subheader(f"Ciclo {i}")
        dt_fim_c = data_corrente + relativedelta(months=11)
        dt_aniv_c = data_corrente + relativedelta(years=1)
        dt_lim_c = dt_aniv_c + relativedelta(days=90)
        
        c1, c2 = st.columns(2)
        with c1:
            st.info(f"**Base:** {data_corrente.strftime('%d/%m/%Y')} | **Janela:** {dt_aniv_c.strftime('%d/%m/%Y')} a {dt_lim_c.strftime('%d/%m/%Y')}")
        with c2:
            dt_ped_c = st.date_input(f"Data do Pedido C{i}:", value=dt_aniv_c, key=f"p{i}", format="DD/MM/YYYY")
        
        df_c = get_ist_rep(data_corrente, dt_fim_c) if "IST" in sel_idx else get_index_data_rep("433" if "IPCA" in sel_idx else "189", data_corrente.strftime('%d/%m/%Y'), dt_fim_c.strftime('%d/%m/%Y'))
        
        if df_c is not None:
            v_ciclo = (1 + df_c['valor']).prod() - 1
            fator_acum *= (1 + v_ciclo)
            st.success(f"Variação Ciclo {i}: **{v_ciclo*100:,.4f}%**".replace('.', ','))
            
            status = "✅ No Prazo" if dt_ped_c <= dt_lim_c else "❌ PRECLUSO"
            resumo.append({"Ciclo": i, "Variação": f"{v_ciclo*100:,.2f}%", "Status": status})
            
            # Lógica de Arrasto: Se precluso, a base do próximo ciclo é a data do pedido[cite: 2]
            data_corrente = dt_ped_c if dt_ped_c > dt_lim_c else dt_aniv_c
        else:
            st.error(f"Erro no Ciclo {i}.")

if resumo:
    st.divider()
    st.metric("Total Acumulado (Represado)", f"{(fator_acum - 1)*100:,.4f}%".replace('.', ','))
    st.table(pd.DataFrame(resumo))