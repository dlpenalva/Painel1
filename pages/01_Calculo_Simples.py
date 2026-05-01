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
        # Lê o CSV tratando o formato jan/22
        df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
        df.columns = ['mes_ano', 'indice']
        
        meses_map = {'jan':1,'fev':2,'mar':3,'abr':4,'mai':5,'jun':6,'jul':7,'ago':8,'set':9,'out':10,'nov':11,'dez':12}
        df['mes_str'] = df['mes_ano'].str.split('/').str[0].str.lower().map(meses_map)
        df['ano_str'] = "20" + df['mes_ano'].str.split('/').str[1]
        df['data'] = pd.to_datetime(df[['ano_str', 'mes_str']].assign(day=1).rename(columns={'ano_str':'year','mes_str':'month'}))
        df = df.sort_values('data')

        # Ajuste IST: Mês anterior ao início e mês anterior ao aniversário
        d_inicio_ref = (data_inicio - relativedelta(months=1)).replace(day=1)
        d_fim_ref = (data_fim).replace(day=1) 

        idx_ini = df[df['data'] == d_inicio_ref]['indice'].values[0]
        idx_fim = df[df['data'] == d_fim_ref]['indice'].values[0]
        
        df_periodo = df[(df['data'] > d_inicio_ref) & (df['data'] <= d_fim_ref)].copy()
        df_periodo['valor'] = df_periodo['indice'].pct_change()
        df_periodo.iloc[0, df_periodo.columns.get_loc('valor')] = (df_periodo.iloc[0]['indice'] / idx_ini) - 1
        
        return df_periodo[['data', 'valor']]
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

st.subheader("Resultado da Análise")

if "IST" in tipo_idx:
    df_dados = get_ist_mock_api(dt_base, dt_fim_calculo)
else:
    cod = "433" if "IPCA" in tipo_idx else "189"
    df_dados = get_index_data(cod, dt_base.strftime('%d/%m/%Y'), dt_fim_calculo.strftime('%d/%m/%Y'))

if df_dados is not None:
    variacao = (1 + df_dados['valor']).prod() - 1
    c1, c2 = st.columns(2)
    c1.metric("Variação do Período (12 meses)", f"{variacao*100:,.2f}%".replace('.', ','))
    c2.metric("Status da Solicitação", status)

    st.info(f"**Janela de Admissibilidade:** {dt_aniv_contratual.strftime('%d/%m/%Y')} até {limite_90.strftime('%d/%m/%Y')}")

    st.subheader("Memória de Cálculo (Fator de Prova)")
    resumo_prova = {
        "Descrição": ["Início do Ciclo", "Fim do Ciclo (Mês ref.)", "Aniversário Contratual", "Limite Admissibilidade", "Data do Pedido", "Status"],
        "Data / Valor": [dt_base.strftime('%m/%Y'), dt_fim_calculo.strftime('%m/%Y'), dt_aniv_contratual.strftime('%d/%m/%Y'), limite_90.strftime('%d/%m/%Y'), dt_solic.strftime('%d/%m/%Y'), status]
    }
    st.table(pd.DataFrame(resumo_prova))

    with st.expander("Visualizar Detalhamento Mensal"):
        df_display = df_dados.copy()
        df_display['Variação Mensal'] = df_display['valor'].map(lambda x: f"{x*100:.4f}%")
        st.dataframe(df_display[['data', 'Variação Mensal']], use_container_width=True)