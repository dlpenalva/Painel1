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
            return {'var': (v_fim/v_ini)-1, 'detalhe': f"IST: {v_fim} / {v_ini}"}
        else:
            url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie}/dados?formato=json&dataInicial={d_ini.strftime('%d/%m/%Y')}&dataFinal={d_fim.strftime('%d/%m/%Y')}"
            r = requests.get(url, timeout=10).json()
            df_temp = pd.DataFrame(r)
            df_temp['v'] = df_temp['valor'].astype(float) / 100
            # Metodologia de produtório para IPCA/IGPM
            var_acum = (1 + df_temp['v']).prod() - 1
            return {'var': var_acum, 'detalhe': "Produtório de taxas mensais (BC)"}
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo de Represados")

with st.sidebar:
    dt_base_org = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_ciclos = st.number_input("Ciclos:", min_value=1, max_value=5, value=2)
    idx_sel = st.selectbox("Índice:", ["IPCA (433)", "IGP-M (189)", "IST (Série Local)"])

data_atual = dt_base_org
fator_total = 1.0
historico = []

for i in range(1, int(qtd_ciclos) + 1):
    st.subheader(f"Ciclo {i}")
    d_fim = data_atual + relativedelta(months=11)
    res_c = get_data_rep("433" if "IPCA" in idx_sel else "189", data_atual, d_fim, "IST" in idx_sel)
    
    if res_c:
        fator_total *= (1 + res_c['var'])
        valor_ciclo_fmt = f"{res_c['var']*100:,.2f}%".replace('.', ',')
        
        with st.expander(f"🔍 Memória de Cálculo - Ciclo {i}"):
            st.write(f"**Método:** {res_c['detalhe']}")
            st.write(f"Variação do Ciclo: {valor_ciclo_fmt}")
        
        historico.append({"Ciclo": i, "Variação": valor_ciclo_fmt})
        data_atual = data_atual + relativedelta(years=1)
    else:
        st.error(f"Falha no Ciclo {i}")

if historico:
    st.divider()
    total_fmt = f"{(fator_total - 1)*100:,.2f}%".replace('.', ',')
    st.metric("Variação Acumulada Final", total_fmt)
    st.table(pd.DataFrame(historico))