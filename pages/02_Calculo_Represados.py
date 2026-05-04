import streamlit as st
import pandas as pd
from datetime import datetime

# Título robusto para o Bloco A
st.title("🔄 Admissibilidade: Reajustes Múltiplos")

# Recuperação da Inteligência do Bloco A (Regras Gerais)
with st.expander("📖 Regras de Admissibilidade (Lei 13.303/2016)", expanded=False):
    st.markdown("""
    * **Periodicidade:** Mínimo de 12 meses.
    * **Data-Base:** Conforme edital ou última repactuação.
    * **Índice:** Verificação de disponibilidade do IST, IPCA ou IGP-M.
    """)

# 1. Parâmetros de Entrada
with st.container(border=True):
    col1, col2 = st.columns(2)
    with col1:
        qtd_ciclos = st.number_input("Ciclos em atraso:", min_value=2, max_value=10, step=1)
        indice_fixo = st.selectbox("Índice de Reajuste:", ["IST", "IPCA", "IGP-M"], key="sel_idx")
    with col2:
        st.info("A inteligência de busca automática de índices deve ser mantida aqui.")

# 2. Processamento dos Ciclos (Memória de Cálculo)
dados_dos_ciclos = []
fator_total = 1.0

for i in range(int(qtd_ciclos)):
    st.subheader(f"Cálculo Ciclo {i+1}")
    c1, c2, c3 = st.columns(3)
    with c1:
        dt_base = st.date_input(f"Data-Base C{i+1}:", key=f"d_base_{i}")
    with c2:
        dt_pedido = st.date_input(f"Data Pedido C{i+1}:", key=f"d_ped_{i}")
    with c3:
        var_manual = st.number_input(f"Variação C{i+1} (ex: 1.0450):", format="%.4f", step=0.0001, key=f"f_man_{i}")
    
    fator_total *= var_manual
    dados_dos_ciclos.append({
        "Ciclo": f"C{i+1}",
        "Data-Base": dt_base,
        "Data Pedido": dt_pedido,
        "Fator": var_manual,
        "Percentual (%)": (var_manual - 1) * 100
    })

# 3. Homologação para Bloco B
st.divider()
if st.button("🚀 Homologar Admissibilidade e Enviar para Valor Global"):
    # Salva toda a inteligência para o Bloco B e Relatório
    st.session_state['dados_admissibilidade'] = {
        'tipo': 'Múltiplo',
        'indice': indice_fixo,
        'fator': fator_total,
        'detalhamento_ciclos': dados_dos_ciclos,
        'qtd_ciclos': qtd_ciclos
    }
    st.success("Dados enviados! Prossiga para o menu 'Valor Global'.")
    st.balloons()

# Memória de Cálculo Visual
if dados_dos_ciclos:
    st.table(pd.DataFrame(dados_dos_ciclos))