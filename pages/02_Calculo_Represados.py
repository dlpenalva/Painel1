import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Facilitador - Cálculo de Represados", layout="wide")

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

with st.expander("Configuração de Ciclos Acumulados", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        dt_original = st.date_input("Data-Base Original:", value=datetime(2021, 5, 1), format="DD/MM/YYYY")
    with col2:
        qtd_anos = st.number_input("Quantidade de anos represados:", min_value=2, max_value=10, value=2)
    with col3:
        indice_ref = st.selectbox("Índice de Referência:", ["IPCA (433)", "IGP-M (189)"])

if st.button("Processar Ciclos Automáticos"):
    fator_acumulado = 1.0
    resumo = []
    cod_serie = "433" if "IPCA" in indice_ref else "189"

    for i in range(1, qtd_anos + 1):
        d_ini = dt_original + relativedelta(years=i-1)
        d_fim = dt_original + relativedelta(years=i)
        
        var_ciclo = get_index_data(cod_serie, d_ini.strftime('%d/%m/%Y'), d_fim.strftime('%d/%m/%Y'))
        
        if var_ciclo is not None:
            fator_ciclo = 1 + var_ciclo
            fator_acumulado *= fator_ciclo
            resumo.append({
                "Ciclo": f"Ano {i}",
                "Período": f"{d_ini.strftime('%d/%m/%Y')} a {d_fim.strftime('%d/%m/%Y')}",
                "Variação": f"{var_ciclo*100:,.4f}%".replace('.', ','),
                "Efeito Financeiro": d_fim.strftime('%d/%m/%Y')
            })
        else:
            st.error(f"Não foi possível obter dados para o Ciclo {i}")

    perc_total = (fator_acumulado - 1) * 100
    
    st.divider()
    c1, c2 = st.columns(2)
    c1.metric("Variação Total Acumulada", f"{perc_total:,.4f}%".replace('.', ','))
    c2.metric("Multiplicador Final (Cascata)", f"{fator_acumulado:.6f}")

    st.subheader("Memória de Cálculo dos Períodos Represados")
    st.table(pd.DataFrame(resumo))