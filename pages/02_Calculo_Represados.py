import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo de Represados", layout="wide")

def get_ist_data_safe(dt_i, dt_f):
    try:
        df = pd.read_csv('ist.csv', sep=None, engine='python')
        df.columns = ['M', 'V'] + list(df.columns[2:])
        meses_map = {'jan':'Jan','fev':'Feb','mar':'Mar','abr':'Apr','mai':'May','jun':'Jun','jul':'Jul','ago':'Aug','set':'Sep','out':'Oct','nov':'Nov','dez':'Dec'}
        df['M'] = df['M'].astype(str).str.strip().str.lower()
        for pt, en in meses_map.items(): df['M'] = df['M'].str.replace(pt, en)
        df['D'] = pd.to_datetime(df['M'], format='%b/%y', errors='coerce')
        df['V'] = df['V'].astype(str).str.replace('.', '').str.replace(',', '.').astype(float)
        
        v_i = df[df['D'] == pd.to_datetime(dt_i.strftime('%Y-%m-01'))]['V'].values[0]
        v_f = df[df['D'] == pd.to_datetime(dt_f.strftime('%Y-%m-01'))]['V'].values[0]
        return (v_f / v_i) - 1, v_i, v_f
    except: return None, None, None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo de Represados")

with st.sidebar:
    dt_base_orig = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_anos = st.number_input("Ciclos:", min_value=1, value=2)
    idx_ref = st.selectbox("Índice:", ["IST (Série Local)", "IPCA", "IGP-M"])

resumo = []
f_acum = 1.0
dt_at = dt_base_orig

try:
    for i in range(1, int(qtd_anos) + 1):
        dt_f = dt_at + relativedelta(months=11)
        dt_aniv = dt_at + relativedelta(years=1)
        
        st.subheader(f"Ciclo {i}")
        col_ped = st.columns(1)[0]
        dt_p = col_ped.date_input(f"Data do Pedido Ciclo {i}", value=dt_aniv, key=f"k{i}")
        
        if "IST" in idx_ref:
            v, vi, vf = get_ist_data_safe(dt_at, dt_f)
            if v is not None:
                st.latex(r"IST_{C" + str(i) + r"} = \frac{" + f"{vf:,.3f}" + r"}{" + f"{vi:,.3f}" + r"} - 1 = " + f"{v*100:.2f}\%")
                f_acum *= (1 + v)
                resumo.append({"Ciclo": i, "Variação": f"{v*100:.2f}%"})
                dt_at = dt_p
            else:
                st.warning(f"Dados do Ciclo {i} não encontrados no CSV.")
                break
    
    if resumo:
        st.divider()
        st.metric("Acumulado Final", f"{(f_acum-1)*100:,.2f}%".replace('.', ','))
except Exception as e:
    st.error(f"Erro ao processar ciclos: {e}")