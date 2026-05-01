import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo de Represados", layout="wide")

def get_ist_data(dt_i, dt_f):
    try:
        df = pd.read_csv('ist.csv', sep=None, engine='python')
        df.columns = [c.strip().upper() for c in df.columns]
        meses_map = {'jan':'Jan','fev':'Feb','mar':'Mar','abr':'Apr','mai':'May','jun':'Jun','jul':'Jul','ago':'Aug','set':'Sep','out':'Oct','nov':'Nov','dez':'Dec'}
        for pt, en in meses_map.items(): df['MES_ANO'] = df['MES_ANO'].str.lower().str.replace(pt, en)
        df['DATA_DT'] = pd.to_datetime(df['MES_ANO'], format='%b/%y')
        df['INDICE_NIVEL'] = df['INDICE_NIVEL'].astype(str).str.replace('.', '').str.replace(',', '.').astype(float)
        
        v_i = df[df['DATA_DT'] == pd.to_datetime(dt_i.strftime('%Y-%m-01'))]['INDICE_NIVEL'].values[0]
        v_f = df[df['DATA_DT'] == pd.to_datetime(dt_f.strftime('%Y-%m-01'))]['INDICE_NIVEL'].values[0]
        return (v_f / v_i) - 1, v_i, v_f
    except: return None, None, None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo de Represados")

with st.sidebar:
    st.header("Configuração")
    dt_base_orig = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_anos = st.number_input("Ciclos:", min_value=1, value=2)
    idx_ref = st.selectbox("Índice:", ["IST (Série Local)", "IPCA (433)", "IGP-M (189)"])

resumo_final = []
fator_acum = 1.0
dt_atual = dt_base_orig

for i in range(1, int(qtd_anos) + 1):
    with st.container():
        dt_fim = dt_atual + relativedelta(months=11)
        dt_aniv = dt_atual + relativedelta(years=1)
        st.subheader(f"Ciclo {i}")
        
        col_a, col_b = st.columns(2)
        with col_a: st.write(f"Período: {dt_atual.strftime('%m/%Y')} a {dt_fim.strftime('%m/%Y')}")
        with col_b: dt_ped = st.date_input(f"Pedido Ciclo {i}", value=dt_aniv, key=f"p{i}")
        
        if "IST" in idx_ref:
            var, vi, vf = get_ist_data(dt_atual, dt_fim)
            if var is not None:
                st.latex(r"IST_{C" + str(i) + r"} = \frac{" + f"{vf:,.3f}" + r"}{" + f"{vi:,.3f}" + r"} - 1 = " + f"{var*100:.2f}\%")
                fator_acum *= (1 + var)
                resumo_final.append({"Ciclo": i, "Variação": f"{var*100:.2f}%", "Status": "Processado"})
                dt_atual = dt_ped
            else: st.error("Dados não encontrados no CSV para este período.")
        else:
            st.warning("Cálculo para IPCA/IGPM em implementação nesta versão resiliente.")
    st.divider()

if resumo_final:
    st.metric("Variação Total Acumulada", f"{(fator_acum-1)*100:,.2f}%".replace('.', ','))
    st.table(pd.DataFrame(resumo_final))