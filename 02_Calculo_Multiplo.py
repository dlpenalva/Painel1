import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="GCC - Cálculo de Passivos", layout="wide")

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("🧮 Cálculo de Passivos")

with st.expander("⚙️ Configuração dos Períodos Acumulados", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        dt_base_original = st.date_input("Data-Base Inicial:", value=datetime(2022, 5, 4), format="DD/MM/YYYY")
    with col2:
        qtd_ciclos = st.number_input("Anos Acumulados:", min_value=2, max_value=5, value=2)

st.info("Insira a variação % de cada ciclo para calcular o efeito cascata.")

fator_acumulado = 1.0
tabela_dados = []

for i in range(1, qtd_ciclos + 1):
    ini = dt_base_original + relativedelta(years=i-1)
    fim = dt_base_original + relativedelta(years=i)
    
    col_a, col_b = st.columns([3, 2])
    with col_a:
        st.write(f"**Ciclo {i}:** {ini.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}")
    with col_b:
        perc = st.number_input(f"Variação %", key=f"p_{i}", format="%.4f")
    
    fator_ciclo = 1 + (perc / 100)
    fator_acumulado *= fator_ciclo
    
    tabela_dados.append({
        "Ciclo": i,
        "Período": f"{ini.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}",
        "Variação": f"{perc:,.2f}%".replace('.', ','),
        "Início do Efeito": fim.strftime('%d/%m/%Y')
    })

perc_final = (fator_acumulado - 1) * 100

st.divider()
c1, c2 = st.columns(2)
c1.metric("Variação Acumulada (Cascata)", f"{perc_final:,.4f}%".replace('.', ','))
c2.metric("Multiplicador Final", f"{fator_acumulado:.6f}")

st.subheader("📋 Quadro de Retroativos")
st.table(pd.DataFrame(tabela_dados))