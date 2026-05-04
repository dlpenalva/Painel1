import streamlit as st
import pandas as pd
from datetime import datetime

st.title("🔄 Reajustes Múltiplos (Cálculo Represado)")

# 1. Configuração Inicial dos Ciclos
with st.expander("1. Configuração dos Ciclos", expanded=True):
    col_q, col_i = st.columns(2)
    with col_q:
        qtd_ciclos = st.number_input("Quantidade de Ciclos em Atraso:", min_value=2, max_value=10, value=2)
    with col_i:
        indice_selecionado = st.selectbox("Selecione o Índice de Reajuste:", ["IST", "IPCA", "IGP-M"], key="idx_multi")

# 2. Entrada de Dados por Ciclo
dados_ciclos = []
fator_acumulado = 1.0

st.subheader("2. Parâmetros por Período")

for i in range(int(qtd_ciclos)):
    st.markdown(f"#### Ciclo {i+1}")
    c1, c2, c3 = st.columns([2, 2, 2])
    
    with c1:
        # Tenta sugerir a data base com base no ciclo anterior (opcional)
        dt_base = st.date_input(f"Data-Base Ciclo {i+1}:", key=f"base_{i}")
    with c2:
        dt_ped = st.date_input(f"Data do Pedido Ciclo {i+1}:", key=f"ped_{i}")
    with c3:
        fator_ciclo = st.number_input(f"Fator do Ciclo {i+1} (ex: 1.0450):", 
                                      format="%.4f", 
                                      step=0.0001, 
                                      key=f"fator_{i}",
                                      help="Insira o índice calculado para este período específico.")
    
    # Cálculo em cascata: Fator1 * Fator2 * ...
    fator_acumulado *= fator_ciclo
    
    dados_ciclos.append({
        "Ciclo": f"C{i+1}",
        "Data-Base": dt_base.strftime("%d/%m/%Y"),
        "Data Pedido": dt_ped.strftime("%d/%m/%Y"),
        "Fator": fator_ciclo,
        "Percentual (%)": round((fator_ciclo - 1) * 100, 4)
    })
    st.divider()

# 3. Resultado e Homologação
st.subheader("3. Resultado Consolidado")
col_res1, col_res2 = st.columns(2)

with col_res1:
    st.metric("Fator Acumulado Total", f"{fator_acumulado:.4f}")
with col_res2:
    variacao_total = (fator_acumulado - 1) * 100
    st.metric("Variação Percentual Total", f"{variacao_total:.2f}%")

# Botão de Homologação com Persistência
if st.button("Homologar Reajustes Múltiplos"):
    # Gravação Crítica no Session State para os Blocos B e Relatório
    st.session_state['dados_admissibilidade'] = {
        'tipo': 'Múltiplo',
        'indice': indice_selecionado,
        'fator': fator_acumulado,
        'detalhamento_ciclos': dados_ciclos,
        'qtd_ciclos': qtd_ciclos,
        'timestamp': datetime.now().strftime("%H:%M:%S")
    }
    
    st.success("✅ Admissibilidade homologada! Os dados foram enviados para o 'Valor Global'.")
    st.balloons()

# 4. Tabela de Memória de Cálculo
if dados_ciclos:
    with st.expander("Visualizar Detalhamento dos Ciclos", expanded=False):
        df_resumo = pd.DataFrame(dados_ciclos)
        st.table(df_resumo)