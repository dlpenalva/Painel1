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
        meses_map = {'jan':1,'fev':2,'mar':3,'abr':4,'mai':5,'jun':6,'jul':7,'ago':8,'set':9,'out':10,'nov':11,'dez':12}
        df['mes_str'] = df['mes_ano'].str.split('/').str[0].str.lower().map(meses_map)
        df['ano_str'] = "20" + df['mes_ano'].str.split('/').str[1]
        df['data'] = pd.to_datetime(df[['ano_str', 'mes_str']].assign(day=1).rename(columns={'ano_str':'year','mes_str':'month'}))
        df = df.sort_values('data')

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
st.title("Cálculo de Represados")

with st.sidebar:
    st.header("Configuração")
    dt_base_original = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_anos = st.number_input("Quantidade de Ciclos:", min_value=1, max_value=10, value=2)
    indice_ref = st.selectbox("Índice:", ["IPCA (433)", "IGP-M (189)", "IST (Série Local)"])

resumo_final = []
fator_acumulado = 1.0
data_base_atual = dt_base_original

for i in range(1, int(qtd_anos) + 1):
    with st.container():
        fim_periodo_calculo = data_base_atual + relativedelta(months=11)
        aniv_teorico = data_base_atual + relativedelta(years=1)
        limite_90 = aniv_teorico + relativedelta(days=90)
        
        st.subheader(f"Ciclo {i}")
        col_inf, col_ped = st.columns(2)
        with col_inf:
            st.markdown(f"**Data-Base Atual:** `{data_base_atual.strftime('%d/%m/%Y')}`")
            st.markdown(f"**Período de Índice:** `{data_base_atual.strftime('%m/%Y')}` a `{fim_periodo_calculo.strftime('%m/%Y')}`")
            st.info(f"**Janela de Admissibilidade:** {aniv_teorico.strftime('%d/%m/%Y')} até {limite_90.strftime('%d/%m/%Y')}")
        
        with col_ped:
            dt_pedido = st.date_input(f"Data do Pedido - Ciclo {i}:", value=aniv_teorico, key=f"ped_{i}", format="DD/MM/YYYY")
        
        if "IST" in indice_ref:
            df_ciclo = get_ist_mock_api(data_base_atual, fim_periodo_calculo)
        else:
            cod_serie = "433" if "IPCA" in indice_ref else "189"
            df_ciclo = get_index_data_detailed(cod_serie, data_base_atual.strftime('%d/%m/%Y'), fim_periodo_calculo.strftime('%d/%m/%Y'))
        
        if df_ciclo is not None:
            var_ciclo = (1 + df_ciclo['valor']).prod() - 1
            fator_acumulado *= (1 + var_ciclo)
            status = "✅ No Prazo" if dt_pedido <= limite_90 else "❌ PRECLUSO (Arrasta Base)"
            
            with st.expander(f"🔍 Auditoria Mensal - Ciclo {i}"):
                df_view = df_ciclo.copy()
                df_view['Variação Mensal'] = df_view['valor'].map(lambda x: f"{x*100:.4f}%")
                st.table(df_view[['data', 'Variação Mensal']].rename(columns={'data': 'Mês/Ano'}))
            
            resumo_final.append({"Ciclo": i, "Referência": f"{data_base_atual.strftime('%m/%y')} - {fim_periodo_calculo.strftime('%m/%y')}", "Variação": f"{var_ciclo*100:,.2f}%".replace('.', ','), "Status": status})
            data_base_atual = dt_pedido
        else: st.error(f"Erro no Ciclo {i}.")
    st.markdown("---")

if resumo_final:
    st.header("Resultado Consolidado")
    st.metric("Variação Total Acumulada", f"{(fator_acumulado - 1)*100:,.2f}%".replace('.', ','))
    st.table(pd.DataFrame(resumo_final))