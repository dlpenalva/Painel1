import streamlit as st
import pandas as pd

# --- CONFIGURAÇÃO DA PÁGINA (ISOLAMENTO) ---
st.set_page_config(page_title="Valor Global do Contrato", layout="wide")

# Interface Profissional Telebras
st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Gestão de Valor Global e Execução")

st.info("""
**Objetivo:** Consolidar o valor global do contrato sob duas perspectivas: 
1. **Financeira:** Focada em pagamentos realizados e projeção do saldo remanescente.
2. **Estoque/Itens:** Focada no consumo físico e aplicação de valores unitários reajustados.
""")

# --- BLOCO 1: PARÂMETROS DO REAJUSTE (ENTRADA MANUAL) ---
st.header("1. Parâmetros do Reajuste")

col1, col2, col3 = st.columns([1, 1, 1.5])

with col1:
    indice_nome = st.selectbox("Índice de Reajuste:", ["IST", "IPCA", "IGP-M"], key="vg_indice")
    dt_base_orig = st.date_input("Data-Base Original:", format="DD/MM/YYYY", key="vg_dt_base")

with col2:
    qtd_ciclos = st.number_input("Quantidade de Ciclos Reajustados:", min_value=1, max_value=10, value=1, key="vg_qtd_ciclos")
    marco_reajuste = st.date_input("Marco do Último Reajuste:", format="DD/MM/YYYY", key="vg_marco")

with col3:
    st.markdown("**Fatores de Reajuste por Ciclo**")
    # Tabela dinâmica para entrada dos fatores calculados nos outros módulos
    dados_fatores = {
        "Ciclo": [f"C{i}" for i in range(qtd_ciclos + 1)],
        "Fator Acumulado": [1.0000] * (qtd_ciclos + 1)
    }
    df_fatores = pd.DataFrame(dados_fatores)
    fatores_editados = st.data_editor(
        df_fatores, 
        hide_index=True, 
        use_container_width=True, 
        key="vg_editor_fatores"
    )

# Captura o fator do último ciclo para atualizar o saldo
fator_vigente = fatores_editados["Fator Acumulado"].iloc[-1]

st.divider()

# --- NAVEGAÇÃO POR ABAS ---
tab_financeira, tab_estoque, tab_comparativo = st.tabs([
    "📊 Apuração Financeira", 
    "📦 Controle de Estoque / Itens", 
    "⚖️ Comparativo e Relatório"
])

# --- ABA 1: VISÃO FINANCEIRA (FLUXO DE CAIXA) ---
with tab_financeira:
    st.subheader("Fluxo de Execução Financeira")
    st.write("Insira os valores totais efetivamente pagos em cada ciclo de execução (considerando glosas e descontos).")
    
    # Grid para entrada dos valores financeiros executados
    dados_fin = {
        "Ciclo": [f"C{i}" for i in range(qtd_ciclos + 1)],
        "Executado Financeiro (R$)": [0.0] * (qtd_ciclos + 1)
    }
    df_fin = pd.DataFrame(dados_fin)
    fin_editado = st.data_editor(
        df_fin, 
        hide_index=True, 
        key="vg_editor_financeiro", 
        use_container_width=True
    )
    
    st.markdown("### Atualização do Saldo Remanescente")
    col_s1, col_s2 = st.columns(2)
    
    with col_s1:
        saldo_rem_fin = st.number_input(
            "Saldo Remanescente (Valor de Face / Preço Original):", 
            min_value=0.0, 
            format="%.2f", 
            step=1000.0,
            help="Valor que ainda não foi executado, posicionado no marco do último reajuste."
        )
    
    with col_s2:
        st.write(f"**Fator Vigente Aplicado:** `{fator_vigente:.4f}`")
        saldo_atualizado = saldo_rem_fin * fator_vigente
        st.info(f"**Saldo Atualizado Projetado:** R$ {saldo_atualizado:,.2f}")
    
    # Cálculos Consolidados
    total_exec_fin = fin_editado["Executado Financeiro (R$)"].sum()
    valor_global_fin = total_exec_fin + saldo_atualizado
    
    st.divider()
    m1, m2 = st.columns(2)
    m1.metric("Total Executado (Acumulado)", f"R$ {total_exec_fin:,.2f}")
    m2.metric("VALOR GLOBAL FINANCEIRO", f"R$ {valor_global_fin:,.2f}", delta_color="normal")

# --- ABA 2: VISÃO DE ESTOQUE (PROXIMA ETAPA) ---
with tab_estoque:
    st.subheader("Controle Físico de Itens e Quantidades")
    st.write("Esta seção calculará o valor global com base na estrutura contratual de itens.")
    
    # Placeholder para o Passo 4
    st.warning("Aguardando definição: você prefere digitar os itens ou fazer upload de um Excel?")
    
    if st.button("Simular Estrutura de Itens"):
        st.info("A estrutura será composta por: Item | Qtd Contratada | Qtd Consumida por Ciclo | Valor Unitário.")

# --- ABA 3: COMPARATIVO ---
with tab_comparativo:
    st.subheader("Comparativo: Financeiro vs. Estoque")
    
    col_c1, col_c2 = st.columns(2)
    with col_c1:
        st.metric("Visão Financeira", f"R$ {valor_global_fin:,.2f}")
    with col_c2:
        st.metric("Visão por Estoque", "R$ 0,00 (Pendente)")
        
    st.write("---")
    st.markdown("""
    **Nota Explicativa:**  
    A diferença entre as visões ocorre porque a **Financeira** reflete o desembolso real (caixa), enquanto a 
    **Visão por Estoque** reflete a obrigação contratual física valorizada pelos índices vigentes.
    """)