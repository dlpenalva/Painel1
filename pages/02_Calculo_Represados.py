import streamlit as st
import pandas as pd
from datetime import datetime

# Título e Identidade Visual
st.title("🔄 Admissibilidade: Reajustes Múltiplos")

# 1. Recuperação das Regras de Negócio e Seleção de Índice
with st.container(border=True):
    col_q, col_i = st.columns(2)
    with col_q:
        qtd_ciclos = st.number_input("Quantidade de Ciclos em Atraso:", min_value=2, max_value=10, value=2, step=1)
    with col_i:
        # Restaurando a seleção de índice que havia sumido
        indice_selecionado = st.selectbox("Selecione o Índice de Reajuste:", ["IST", "IPCA", "IGP-M"], key="idx_multi_restore")

# 2. Processamento Dinâmico dos Ciclos (Lógica Original)
dados_ciclos = []
fator_acumulado = 1.0

st.subheader("Parâmetros por Período")

for i in range(int(qtd_ciclos)):
    with st.expander(f"Configuração do Ciclo {i+1}", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            dt_base = st.date_input(f"Data-Base C{i+1}:", key=f"base_r_{i}")
        with c2:
            dt_ped = st.date_input(f"Data do Pedido C{i+1}:", key=f"ped_r_{i}")
        with c3:
            # Campo para inserção do fator calculado (ex: 1.0500 para 5%)
            fator_ciclo = st.number_input(f"Fator C{i+1}:", format="%.4f", step=0.0001, value=1.0000, key=f"fator_r_{i}")
        
        fator_acumulado *= fator_ciclo
        dados_ciclos.append({
            "Ciclo": f"C{i+1}",
            "Data-Base": dt_base,
            "Data Pedido": dt_ped,
            "Fator": fator_ciclo,
            "Percentual (%)": round((fator_ciclo - 1) * 100, 4)
        })

# 3. Resultado Consolidado e Memória de Cálculo
st.divider()
col_res1, col_res2 = st.columns(2)
with col_res1:
    st.metric("Fator Acumulado Final", f"{fator_acumulado:.4f}")
with col_res2:
    st.metric("Variação Total (%)", f"{(fator_acumulado - 1) * 100:.2f}%")

# 4. BOTÃO DE HOMOLOGAÇÃO (A Única Injeção Nova)
if st.button("🚀 Homologar e Enviar para Valor Global"):
    # Salva no session_state para ser lido pela Página 03 e 04
    st.session_state['dados_admissibilidade'] = {
        'tipo': 'Múltiplo',
        'indice': indice_selecionado,
        'fator': fator_acumulado,
        'detalhamento_ciclos': dados_ciclos,
        'qtd_ciclos': qtd_ciclos
    }
    st.success("✅ Admissibilidade Homologada! Dados integrados ao Bloco B.")
    st.balloons()

# Exibição da Tabela de Apoio
st.table(pd.DataFrame(dados_ciclos))