import streamlit as st
import pandas as pd
import io

# --- CONFIGURAÇÃO (ISOLAMENTO) ---
st.set_page_config(page_title="Valor Global - Homologação", layout="wide")

# Interface Telebras
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
    st.markdown("**Fatores de Reajuste (Extraídos da sua Aba PARAMETROS)**")
    df_fatores_base = pd.DataFrame({
        "Ciclo": [f"C{i}" for i in range(qtd_ciclos + 1)],
        "Fator Acumulado": [1.0000] + [1.0468] * qtd_ciclos  # Exemplo baseado no seu print de 4,68%
    })
    
    fatores_editados = st.data_editor(
        df_fatores_base, 
        hide_index=True, 
        use_container_width=True, 
        key="vg_editor_fatores"
    )

# Fator para atualizar o saldo remanescente
fator_vigente = fatores_editados["Fator Acumulado"].iloc[-1]

st.divider()

# --- NAVEGAÇÃO ---
tab_financeira, tab_estoque, tab_comparativo = st.tabs([
    "📊 Apuração Financeira", 
    "📦 Controle de Estoque / Itens", 
    "⚖️ Comparativo e Relatório"
])

# --- ABA FINANCEIRA ---
with tab_financeira:
    st.subheader("Fluxo de Execução Financeira")
    
    df_fin_base = pd.DataFrame({
        "Ciclo": [f"C{i}" for i in range(qtd_ciclos + 1)],
        "Executado Financeiro (R$)": [0.0] * (qtd_ciclos + 1)
    })
    
    fin_editado = st.data_editor(df_fin_base, hide_index=True, key="vg_editor_fin", use_container_width=True)
    
    c_s1, c_s2 = st.columns(2)
    with c_s1:
        saldo_orig = st.number_input("Saldo Remanescente (Valor de Face):", min_value=0.0, format="%.2f", step=1000.0)
        total_pago = fin_editado["Executado Financeiro (R$)"].sum()
    with c_s2:
        saldo_at = saldo_orig * fator_vigente
        st.metric("Saldo Atualizado (Projeção)", f"R$ {saldo_at:,.2f}")

    valor_global_fin = total_pago + saldo_at
    st.divider()
    st.metric("VALOR GLOBAL FINANCEIRO", f"R$ {valor_global_fin:,.2f}")

# --- ABA ESTOQUE (INTEGRAÇÃO COM PLANILHA DO FISCAL) ---
with tab_estoque:
    st.subheader("📦 Processamento da Planilha ITENS_CICLOS")
    st.write("Suba aqui a planilha que você recebe dos fiscais para processar o valor global físico.")
    
    uploaded_file = st.file_uploader("Selecione o arquivo Excel (.xlsx)", type="xlsx")

    if uploaded_file:
        try:
            # Lendo a aba específica conforme seu print
            df_itens = pd.read_excel(uploaded_file, sheet_name="ITENS_CICLOS", skiprows=2) # Ajustar skiprows se necessário
            
            # Identificação das colunas (Baseado no seu print da Imagem 2)
            # Vamos assumir nomes genéricos para processamento
            st.success("Planilha carregada com sucesso!")
            
            # Cálculo de VU Atualizado por Ciclo
            # Aqui o código mapeia o VU C0 e aplica os fatores da Aba 1
            if 'VU C0 (R$)' in df_itens.columns:
                df_itens['VU Atualizado (Projetado)'] = df_itens['VU C0 (R$)'] * fator_vigente
                
                # Exibição de colunas não sensíveis para auditoria
                colunas_seguras = ['Item/Bloco', 'Qtd original C0', 'VU C0 (R$)', 'VU Atualizado (Projetado)']
                st.dataframe(df_itens[colunas_seguras].head(10), use_container_width=True)
                
                # Cálculo do Valor Global por Estoque
                # Soma (Consumido nos Ciclos) + (Remanescente * VU Atualizado)
                valor_global_estoque = (df_itens['VU C0 (R$)'] * df_itens['Qtd original C0']).sum() * fator_vigente # Simplificação para teste
                
                st.divider()
                st.metric("VALOR GLOBAL POR ESTOQUE", f"R$ {valor_global_estoque:,.2f}")
                st.session_state['vg_estoque'] = valor_global_estoque
            else:
                st.error("Coluna 'VU C0 (R$)' não encontrada. Verifique o cabeçalho da planilha.")
        except Exception as e:
            st.error(f"Erro ao ler a aba ITENS_CICLOS: {e}")
    else:
        st.info("Aguardando upload da planilha padronizada.")

# --- ABA COMPARATIVO ---
with tab_comparativo:
    st.subheader("Relatório de Diferenças")
    
    vg_estoque = st.session_state.get('vg_estoque', 0.0)
    diferenca = vg_estoque - valor_global_fin
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Visão Financeira", f"R$ {valor_global_fin:,.2f}")
    c2.metric("Visão Estoque", f"R$ {vg_estoque:,.2f}")
    c3.metric("Diferença (Glosas/Ajustes)", f"R$ {diferenca:,.2f}", delta_color="inverse")
    
    st.markdown(f"""
    ### Resumo Executivo
    O valor global atualizado do contrato, considerando o índice **{indice_nome}** e o fator acumulado de **{fator_vigente:.4f}**, 
    apresenta uma execução financeira de **R$ {total_pago:,.2f}**. 
    
    O saldo remanescente, quando valorizado ao preço atualizado no marco de **{marco_reajuste.strftime('%d/%m/%Y')}**, 
    representa um montante de **R$ {saldo_at:,.2f}**.
    """)