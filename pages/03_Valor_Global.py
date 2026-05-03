import streamlit as st
import pandas as pd

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Valor Global - Homologação", layout="wide")

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Gestão de Valor Global e Execução")

# --- BLOCO 1: PARÂMETROS ---
st.header("1. Parâmetros do Reajuste")

col1, col2, col3 = st.columns([1, 1, 1.5])

with col1:
    indice_nome = st.selectbox("Índice:", ["IST", "IPCA", "IGP-M"], key="vg_indice")
    dt_base_orig = st.date_input("Data-Base Original:", format="DD/MM/YYYY", key="vg_dt_base")

with col2:
    qtd_ciclos = st.number_input("Quantidade de Ciclos:", min_value=1, max_value=10, value=1, key="vg_qtd_ciclos")
    marco_reajuste = st.date_input("Marco do Último Reajuste:", format="DD/MM/YYYY", key="vg_marco")

with col3:
    st.markdown("**Fatores de Reajuste (Referência)**")
    df_fatores_base = pd.DataFrame({
        "Ciclo": [f"C{i}" for i in range(qtd_ciclos + 1)],
        "Fator Acumulado": [1.0000] + [1.0468] * qtd_ciclos 
    })
    fatores_editados = st.data_editor(df_fatores_base, hide_index=True, use_container_width=True, key="vg_edit_fat")

# Definição do fator antes do uso nas abas
fator_vigente = fatores_editados["Fator Acumulado"].iloc[-1]

st.divider()

# --- NAVEGAÇÃO (CRIAÇÃO DAS ABAS) ---
# A variável tab_estoque é criada aqui!
tab_financeira, tab_estoque, tab_comparativo = st.tabs(["📊 Apuração Financeira", "📦 Controle de Estoque", "⚖️ Comparativo"])

with tab_estoque:
    st.subheader("📦 Processamento da Planilha do Fiscal")
    uploaded_file = st.file_uploader("Suba a planilha preenchida (Abas: RETROATIVO e ITENS_CICLOS)", type="xlsx")

    if uploaded_file:
        try:
            # 1. LEITURA DA ABA RETROATIVO
            df_retro = pd.read_excel(uploaded_file, sheet_name="RETROATIVO", skiprows=2)
            total_faturado_real = df_retro["Valor bruto faturado após descontos (R$)"].sum()

            # 2. LEITURA DA ABA ITENS_CICLOS
            df_itens = pd.read_excel(uploaded_file, sheet_name="ITENS_CICLOS", skiprows=2)
            col_rem = [c for c in df_itens.columns if "PREENCHER" in str(c)][-1]
            df_itens['Saldo_Valorizado'] = df_itens[col_rem] * df_itens["VU C0 (R$)"] * fator_vigente
            total_saldo_fisico = df_itens['Saldo_Valorizado'].sum()

            # 3. CONSOLIDAÇÃO
            valor_global_final = total_faturado_real + total_saldo_fisico

            st.success("Planilha processada com sucesso!")
            m1, m2, m3 = st.columns(3)
            m1.metric("Realizado (Retroativo)", f"R$ {total_faturado_real:,.2f}")
            m2.metric("Saldo (Físico)", f"R$ {total_saldo_fisico:,.2f}")
            m3.metric("VALOR GLOBAL REAL", f"R$ {valor_global_final:,.2f}")

            st.session_state['vg_real'] = valor_global_final

        except Exception as e:
            st.error(f"Erro no processamento: {e}")

with tab_financeira:
    st.info("Esta aba é alimentada pelo upload na aba 'Controle de Estoque'.")
    if 'vg_real' in st.session_state:
        st.write(f"Valor Global detectado: **R$ {st.session_state['vg_real']:,.2f}**")

with tab_comparativo:
    st.subheader("Análise Comparativa")
    if 'vg_real' in st.session_state:
        st.write("Dados prontos para comparação.")
    else:
        st.warning("Aguardando upload da planilha na aba de Estoque.")