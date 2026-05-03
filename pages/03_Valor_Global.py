import streamlit as st
import pandas as pd

st.set_page_config(page_title="Valor Global do Contrato", layout="wide")

# Cabeçalho padrão Telebras
st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Gestão de Valor Global e Execução")

st.info("""
**Objetivo:** Consolidar o valor global do contrato sob duas perspectivas: 
1. **Financeira:** Focada em pagamentos e saldo remanescente.
2. **Estoque/Itens:** Focada em consumo físico e valores unitários reajustados.
""")

# --- BLOCO 1: PARÂMETROS DO REAJUSTE ---
st.header("1. Parâmetros do Reajuste")

col1, col2, col3 = st.columns(3)
with col1:
    indice_nome = st.selectbox("Índice de Reajuste:", ["IST", "IPCA", "IGP-M"])
    dt_base_orig = st.date_input("Data-Base Original:", format="DD/MM/YYYY")
with col2:
    qtd_ciclos = st.number_input("Quantidade de Ciclos Reajustados:", min_value=1, max_value=10, value=1)
    marco_reajuste = st.date_input("Marco do Último Reajuste:", format="DD/MM/YYYY")
with col3:
    st.write("**Fatores de Reajuste por Ciclo**")
    # Tabela para entrada manual dos fatores (Passo 2)
    dados_fatores = {
        "Ciclo": [f"C{i}" for i in range(qtd_ciclos + 1)],
        "Fator Acumulado": [1.0000] * (qtd_ciclos + 1)
    }
    df_fatores = pd.DataFrame(dados_fatores)
    # O st.data_editor permite que o usuário digite os fatores apurados nos outros módulos
    fatores_editados = st.data_editor(df_fatores, hide_index=True, use_container_width=True)

# Armazenar o fator acumulado vigente (último da lista)
fator_vigente = fatores_editados["Fator Acumulado"].iloc[-1]

# --- TABS PARA AS VISÕES (PASSO 3 E 4) ---
tab_financeira, tab_estoque, tab_comparativo = st.tabs([
    "📊 Apuração Financeira", 
    "📦 Controle de Estoque / Itens", 
    "⚖️ Comparativo e Relatório"
])

with tab_financeira:
    st.subheader("Fluxo de Execução Financeira")
    st.write("Insira os valores totais pagos em cada ciclo de execução.")
    
    # Grid para entrada financeira
    dados_fin = {
        "Ciclo": [f"C{i}" for i in range(qtd_ciclos + 1)],
        "Executado Financeiro (R$)": [0.0] * (qtd_ciclos + 1)
    }
    df_fin = pd.DataFrame(dados_fin)
    fin_editado = st.data_editor(df_fin, hide_index=True, key="fin_editor")
    
    saldo_rem_fin = st.number_input("Saldo Remanescente (Preço Inicial):", min_value=0.0, format="%.2f")
    
    # Lógica de cálculo financeiro preliminar
    total_exec_fin = fin_editado["Executado Financeiro (R$)"].sum()
    saldo_atualizado = saldo_rem_fin * fator_vigente
    valor_global_fin = total_exec_fin + saldo_atualizado
    
    st.divider()
    st.metric("Valor Global Financeiro Atualizado", f"R$ {valor_global_fin:,.2f}")

with tab_estoque:
    st.subheader("Controle Físico de Itens")
    st.warning("Aguardando definição da estrutura de itens para processamento em massa.")
    # Aqui entrará o Passo 4 (Grid dinâmico de itens)

with tab_comparativo:
    st.subheader("Análise de Desvio e Relatório")
    st.write("Os dados serão consolidados após o preenchimento da aba de Estoque.")