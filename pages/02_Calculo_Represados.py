import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo de Represados", layout="wide")

def get_data_rep(serie, d_ini, d_fim, is_ist):
    try:
        if is_ist:
            df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
            df.columns = [str(col).strip().lower() for col in df.columns]
            df['data'] = pd.to_datetime(df['data'], dayfirst=True)
            r_ini = (d_ini - relativedelta(months=1)).replace(day=1)
            r_fim = d_fim.replace(day=1)
            v_ini = df[df['data'].dt.to_period('M') == r_ini.strftime('%Y-%m')]['indice'].values[0]
            v_fim = df[df['data'].dt.to_period('M') == r_fim.strftime('%Y-%m')]['indice'].values[0]
            return {'var': (v_fim/v_ini)-1, 'i_ini': v_ini, 'i_fim': v_fim, 'p_ini': r_ini, 'p_fim': r_fim}
        else:
            url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie}/dados?formato=json&dataInicial={d_ini.strftime('%d/%m/%Y')}&dataFinal={d_fim.strftime('%d/%m/%Y')}"
            r = requests.get(url, timeout=10).json()
            v_ini, v_fim = float(r[0]['valor']), float(r[-1]['valor'])
            return {'var': (v_fim/v_ini)-1, 'i_ini': v_ini, 'i_fim': v_fim, 'p_ini': d_ini, 'p_fim': d_fim}
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo de Represados")

with st.sidebar:
    dt_base_original = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_ciclos = st.number_input("Ciclos:", min_value=1, max_value=5, value=2)
    idx_sel = st.selectbox("Índice:", ["IPCA (433)", "IGP-M (189)", "IST (Série Local)"])

data_atual = dt_base_original
fator_total = 1.0
historico = []

for i in range(1, int(qtd_ciclos) + 1):
    st.subheader(f"Ciclo {i}")
    d_fim = data_atual + relativedelta(months=11)
    d_aniv = data_atual + relativedelta(years=1)
    d_lim = d_aniv + relativedelta(days=90)
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.info(f"**Base:** {data_atual.strftime('%d/%m/%Y')} | **Janela:** {d_aniv.strftime('%d/%m/%Y')} a {d_lim.strftime('%d/%m/%Y')}")
    with col_b:
        dt_ped = st.date_input(f"Pedido C{i}:", value=d_aniv, key=f"ped{i}", format="DD/MM/YYYY")

    res_c = get_data_rep("433" if "IPCA" in idx_sel else "189", data_atual, d_fim, "IST" in idx_sel)
    
    if res_c:
        fator_total *= (1 + res_c['var'])
        with st.expander(f"🔍 Memória de Cálculo - Ciclo {i}"):
            st.write(f"Período: {res_c['p_ini'].strftime('%m/%Y')} a {res_c['p_fim'].strftime('%m/%Y')}")
            st.write(f"Cálculo: ({res_c['i_fim']} / {res_c['i_ini']}) - 1 = **{res_c['var']*100:.4f}%**")
        
        status = "✅ No Prazo" if dt_ped <= d_lim else "❌ PRECLUSO"
        historico.append({"Ciclo": i, "Variação": f"{res_c['var']*100:.2f}%", "Status": status})
        data_atual = dt_ped if dt_ped > d_lim else d_aniv
    else:
        st.error(f"Falha no Ciclo {i}")

if historico:
    st.divider()
    st.metric("Variação Acumulada Final", f"{(fator_total - 1)*100:,.4f}%".replace('.', ','))
    st.table(pd.DataFrame(historico))