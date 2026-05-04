import streamlit as st
import pandas as pd
from datetime import datetime

st.title("🔄 Reajustes Múltiplos (Cálculo Represado)")

# 1. Entrada de Dados - Configuração dos Ciclos
with st.expander("1. Configuração dos Ciclos", expanded=True):
    col_q, col_i = st.columns(2)
    with col_q:
        qtd_ciclos = st.number_input("Quantidade de Ciclos em Atraso:", min_value=2, max_value=10, value=2)
    with col_i:
        indice_selecionado = st.selectbox("Selecione o Índice de Reajuste:", ["IST", "IPCA", "IGP-M"], key="idx_multi")

# 2. Processamento por Ciclo
dados_ciclos = []
fator_acumulado = 1.0

st.subheader("2. Parâmetros por Período")

for i in range(qtd_ciclos):
    st.markdown(f"#### Ciclo {i+1}")
    c1, c2, c3 = st.columns([2, 2, 2])
    
    with c1:
        dt_base = st.date_input(f"Data-Base Ciclo {i+1}:", key=f"base_{i}")
    with c2:
        dt_ped = st.date_input(f"Data do Pedido Ciclo {i+1}:", key=f"ped_{i}")
    with c3:
        fator_ciclo = st.number_input(f"Fator do Ciclo {i+1} (ex: 1.0450):", 
                                      format="%.4f", step=0.0001, key=f"fator_{i}")
    
    fator_acumulado *= fator_ciclo
    dados_ciclos.append({
        "Ciclo": f"C{i+1}",
        "Data-Base": dt_base,
        "Data Pedido": dt_ped,
        "Fator": fator_ciclo
    })
    st.divider()

# 3. Resumo e Homologação
st.subheader("3. Resultado Consolidado")
col_res1, col_res2 = st.columns(2)

with col_res1:
    st.metric("Fator Acumulado Total", f"{fator_acumulado:.4f}")
with col_res2:
    variacao_total = (fator_acumulado - 1) * 100
    st.metric("Variação Percentual Total", f"{variacao_total:.2f}%")

if st.button("Homologar Reajustes Múltiplos"):
    # Salva no session_state para o Bloco B (Página 03) e Relatório (Página 04)
    st.session_state['dados_admissibilidade'] = {
        'tipo': 'Múltiplo',
        'indice': indice_selecionado,
        'fator': fator_acumulado, # Fator final acumulado
        'detalhamento_ciclos': dados_ciclos,
        'qtd_ciclos': qtd_ciclos,
        'data_homologacao': datetime.now().strftime("%d/%m/%Y")
    }
    st.success(f"Admissibilidade de {qtd_ciclos} ciclos homologada com sucesso!")

# 4. Memória de Cálculo Visual
if dados_ciclos:
    with st.expander("Visualizar Quadro Resumo dos Ciclos"):
        df_resumo = pd.DataFrame(dados_ciclos)
        st.table(df_resumo)