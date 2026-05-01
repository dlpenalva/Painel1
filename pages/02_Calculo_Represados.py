import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo de Represados", layout="wide")

def get_index_data(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=10)
        df = pd.DataFrame(response.json())
        if df.empty: return None
        df['valor'] = df['valor'].astype(float) / 100
        return (1 + df['valor']).prod() - 1
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo de Represados")

with st.sidebar:
    st.header("Configuração")
    dt_base_original = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_anos = st.number_input("Quantidade de Ciclos:", min_value=1, max_value=10, value=2)
    indice_ref = st.selectbox("Índice:", ["IPCA (433)", "IGP-M (189)"])

resumo = []
fator_acumulado = 1.0
data_base_atual = dt_base_original

for i in range(1, int(qtd_anos) + 1):
    with st.expander(f"Detalhamento Ciclo {i}", expanded=True):
        # Intervalo de 12 meses (Mês 0 + 11)
        fim_periodo_calculo = data_base_atual + relativedelta(months=11)
        aniv_teorico = data_base_atual + relativedelta(years=1)
        limite_90 = aniv_teorico + relativedelta(days=90)
        
        col_inf, col_ped = st.columns(2)
        with col_inf:
            st.markdown(f"**Data-Base:** `{data_base_atual.strftime('%d/%m/%Y')}`")
            st.markdown(f"**Intervalo Índice:** `{data_base_atual.strftime('%m/%Y')}` a `{fim_periodo_calculo.strftime('%m/%Y')}`")
        
        with col_ped:
            dt_pedido = st.date_input(f"Data do Pedido - Ciclo {i}:", value=aniv_teorico, key=f"ped_{i}", format="DD/MM/YYYY")
        
        cod_serie = "433" if "IPCA" in indice_ref else "189"
        var_ciclo = get_index_data(cod_serie, data_base_atual.strftime('%d/%m/%Y'), fim_periodo_calculo.strftime('%d/%m/%Y'))
        
        if var_ciclo is not None:
            fator_ciclo = 1 + var_ciclo
            fator_acumulado *= fator_ciclo
            status = "✅ No Prazo" if dt_pedido <= limite_90 else "❌ PRECLUSO (Arrasta Base)"
            
            resumo.append({
                "Ciclo": i,
                "Referência (12 meses)": f"{data_base_atual.strftime('%m/%y')} - {fim_periodo_calculo.strftime('%m/%y')}",
                "Variação": f"{var_ciclo*100:,.2f}%".replace('.', ','),
                "Data do Pedido": dt_pedido.strftime('%d/%m/%Y'),
                "Status": status
            })
            data_base_atual = dt_pedido
        else:
            st.error(f"Erro ao obter dados para o Ciclo {i}")

if resumo:
    st.divider()
    perc_total = (fator_acumulado - 1) * 100
    st.metric("Variação Total Acumulada", f"{perc_total:,.2f}%".replace('.', ','))

    st.subheader("Memória de Cálculo Consolidada (Fator de Prova)")
    st.table(pd.DataFrame(resumo))