import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

def add_years(d, years):
    try:
        return d.replace(year=d.year + years)
    except ValueError:
        return d + (datetime(d.year + years, 1, 1) - datetime(d.year, 1, 1))

st.set_page_config(page_title="Cálculos Represados", layout="wide")

st.title("🔄 Reajustes Múltiplos (Cálculo Represado)")

# Recupera data_base_original do módulo 01 ou define fallback
db_original = st.session_state.get('data_base_anterior', datetime(2023, 1, 1))

with st.expander("1. Configuração dos Ciclos", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        qtd_ciclos = st.number_input("Quantidade de Ciclos em Atraso:", min_value=1, max_value=10, value=2)
    with col2:
        indice_tipo = st.selectbox("Selecione o Índice de Reajuste:", ["IST", "IPCA", "IGP-M"])

st.subheader("2. Parâmetros por Período")

ciclos_data = []
fator_acumulado = 1.0

for i in range(qtd_ciclos):
    st.markdown(f"### Ciclo {i+1}")
    c1, c2, c3 = st.columns(3)
    
    # Lógica de Data-Base Encadeada
    if i == 0:
        # Ciclo 1: 12 meses após a data-base original
        dt_base_sugerida = add_years(db_original, 1)
    else:
        # Ciclos seguintes: 12 meses após a DATA DO PEDIDO do ciclo anterior
        dt_base_sugerida = add_years(ciclos_data[i-1]['data_pedido'], 1)
    
    with c1:
        dt_base = st.date_input(f"Data-Base Ciclo {i+1}:", value=dt_base_sugerida, key=f"db_{i}")
    with c2:
        dt_pedido = st.date_input(f"Data do Pedido Ciclo {i+1}:", value=dt_base, key=f"dp_{i}")
    with c3:
        fator = st.number_input(f"Fator do Ciclo {i+1} (ex: 1.0450):", format="%.4f", step=0.0001, key=f"f_{i}")
    
    fator_acumulado *= fator
    ciclos_data.append({
        'ciclo': i+1,
        'data_base': dt_base,
        'data_pedido': dt_pedido,
        'fator': fator
    })
    st.divider()

# Armazenamento para os módulos seguintes
st.session_state['fator_homologado'] = fator_acumulado
st.session_state['detalhamento_ciclos'] = ciclos_data
st.session_state['tipo_reajuste'] = "Múltiplo"

st.success(f"Fator Acumulado Total: {fator_acumulado:.4f}")

# Memória de Cálculo
if st.checkbox("Exibir Memória de Cálculo"):
    df_memoria = pd.DataFrame(ciclos_data)
    st.table(df_memoria)