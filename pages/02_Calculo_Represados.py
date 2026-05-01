import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo de Represados", layout="wide")

def get_index_data_detailed(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=10)
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
        meses_br = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
        str_inicio = f"{meses_br[d_ref_inicio.month]}/{str(d_ref_inicio.year)[2:]}"
        str_fim = f"{meses_br[d_ref_fim.month]}/{str(d_ref_fim.year)[2:]}"
        v_ini = df[df['mes_ano_clean'] == str_inicio]['indice'].values[0]
        v_fim = df[df['mes_ano_clean'] == str_fim]['indice'].values[0]
        return pd.DataFrame({'data': [data_fim], 'valor': [(v_fim / v_ini) - 1]})
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo de Represados")

with st.sidebar:
    dt_base_original = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_anos = st.number_input("Quantidade de Ciclos:", min_value=1, max_value=10, value=2)
    indice_ref = st.selectbox("Índice:", ["IPCA (433)", "IGP-M (189)", "IST (Série Local)"])

resumo_final = []
fator_acumulado = 1.0
data_base_atual = dt_base_original

for i in range(1, int(qtd_anos) + 1):
    with st.container():
        fim_periodo = data_base_atual + relativedelta(months=11)
        aniv = data_base_atual + relativedelta(years=1)
        limite = aniv + relativedelta(days=90)
        
        st.subheader(f"Ciclo {i}")
        col_inf, col_ped = st.columns(2)
        with col_inf:
            st.markdown(f"**Data-Base:** `{data_base_atual.strftime('%d/%m/%Y')}`")
            st.info(f"**Janela:** {aniv.strftime('%d/%m/%Y')} até {limite.strftime('%d/%m/%Y')}")
        with col_ped:
            dt_pedido = st.date_input(f"Pedido Ciclo {i}:", value=aniv, key=f"p_{i}", format="DD/MM/YYYY")
        
        df_c = get_ist_mock_api(data_base_atual, fim_periodo) if "IST" in indice_ref else get_index_data_detailed("433" if "IPCA" in indice_ref else "189", data_base_atual.strftime('%d/%m/%Y'), fim_periodo.strftime('%d/%m/%Y'))
        
        if df_c is not None:
            var = (1 + df_c['valor']).prod() - 1
            fator_acumulado *= (1 + var)
            status = "✅ No Prazo" if dt_pedido <= limite else "❌ PRECLUSO"
            resumo_final.append({"Ciclo": i, "Variação": f"{var*100:,.2f}%".replace('.', ','), "Status": status})
            data_base_atual = dt_pedido # Regra de arrasto de base para o próximo ciclo
        else: st.error(f"Erro no Ciclo {i}.")

if resumo_final:
    st.metric("Acumulado Total", f"{(fator_acumulado - 1)*100:,.2f}%".replace('.', ','))
    st.table(pd.DataFrame(resumo_final))