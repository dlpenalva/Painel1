import streamlit as st
import pandas as pd

# ... (Mantenha o topo do código com os parâmetros e abas) ...

with tab_estoque:
    st.subheader("📦 Processamento Integrado (Físico + Financeiro)")
    uploaded_file = st.file_uploader("Suba a planilha devolvida pelo fiscal", type="xlsx")

    if uploaded_file:
        try:
            # 1. LEITURA DA ABA RETROATIVO (VISÃO FINANCEIRA REAL)
            # Pulamos as linhas iniciais para chegar no cabeçalho real (ajuste o skiprows se necessário)
            df_retro = pd.read_excel(uploaded_file, sheet_name="RETROATIVO", skiprows=2)
            
            # Limpeza: remove linhas totalmente vazias que o Excel costuma gerar
            df_retro = df_retro.dropna(subset=["Competência", "Valor bruto faturado após descontos (R$)"], how='all')
            
            total_pago_fiscal = df_retro["Valor bruto faturado após descontos (R$)"].sum()

            # 2. LEITURA DA ABA ITENS_CICLOS (VISÃO FÍSICA/ESTOQUE)
            df_itens = pd.read_excel(uploaded_file, sheet_name="ITENS_CICLOS", skiprows=2)
            
            # Identifica dinamicamente a coluna de saldo remanescente (que contém "PREENCHER")
            col_saldo = [c for c in df_itens.columns if "PREENCHER" in str(c)][-1]
            
            # Cálculo: Quantidade Remanescente * Valor Unitário C0 * Fator de Reajuste (da Aba 1)
            df_itens['Subtotal_Atualizado'] = df_itens[col_saldo] * df_itens["VU C0 (R$)"] * fator_vigente
            saldo_fisico_total = df_itens['Subtotal_Atualizado'].sum()

            # 3. CONSOLIDAÇÃO FINAL
            valor_global_total = total_pago_fiscal + saldo_fisico_total

            # --- EXIBIÇÃO DOS RESULTADOS ---
            st.success("Planilha processada com sucesso!")
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Realizado (Fiscal)", f"R$ {total_pago_fiscal:,.2f}")
            c2.metric("Saldo Remanescente", f"R$ {saldo_fisico_total:,.2f}")
            c3.metric("VALOR GLOBAL REAL", f"R$ {valor_global_total:,.2f}")

            # Gráfico rápido de execução (Aba Retroativo)
            st.write("### Evolução do Faturamento Real")
            st.line_chart(df_retro.set_index("Competência")["Valor bruto faturado após descontos (R$)"])

        except Exception as e:
            st.error(f"Erro ao processar: {e}. Verifique se os nomes das abas e colunas coincidem com o padrão.")