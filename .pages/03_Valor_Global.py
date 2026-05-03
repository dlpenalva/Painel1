import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Gestão de Valor Global", layout="wide")

# --- RECUPERAÇÃO DE DADOS DO BLOCO A ---
# Se o Bloco A já foi rodado, pegamos os dados reais. Caso contrário, usamos padrão.
if 'dados_admissibilidade' in st.session_state:
    dados_origem = st.session_state['dados_admissibilidade']
    indice_final = dados_origem['indice']
    fator_final = dados_origem['fator']
    data_base_final = dados_origem['data_base']
    ciclo_final = dados_origem['ciclo_atual']
    st.sidebar.success("✅ Dados importados da Admissibilidade")
else:
    indice_final = "IST (Padrão)"
    fator_final = 1.0468
    data_base_final = "01/05/2025"
    ciclo_final = "C1"
    st.sidebar.warning("⚠️ Usando valores padrão (Admissibilidade não executada)")

st.title("💰 Gestão de Valor Global do Contrato")

# --- CABEÇALHO COM MÉTRICAS ---
st.header(f"Parâmetros de Reajuste - {ciclo_final}")
c1, c2, c3 = st.columns(3)
c1.metric("Índice Aplicado", indice_final)
c2.metric("Fator de Reajuste", f"{fator_final:.4f}")
c3.metric("Data-Base", data_base_final)

st.divider()

# --- ÁREA DE CÁLCULO (BLOCO B) ---
st.subheader("Simulação de Impacto Financeiro")

with st.container(border=True):
    col_input = st.columns([1, 1])
    with col_input[0]:
        valor_original = st.number_input("Valor Global Atual (R$):", min_value=0.0, format="%.2f")
    
    if valor_original > 0:
        novo_valor = valor_original * fator_final
        impacto = novo_valor - valor_original
        
        res1, res2 = st.columns(2)
        res1.metric("Novo Valor Global", f"R$ {novo_valor:,.2/f}".replace(",", "X").replace(".", ",").replace("X", "."))
        res2.metric("Impacto Financeiro (Aumento)", f"R$ {impacto:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), delta_color="inverse")

# --- EXPORTAÇÃO ---
if st.button("📊 Gerar Memória de Cálculo (Excel)"):
    # Lógica para gerar o Excel com Coluna A (Número) e Coluna C (Moeda)
    output = BytesIO()
    df_export = pd.DataFrame({
        'Item': [1],
        'Descrição': [f"Reajuste Global {ciclo_final}"],
        'Valor Anterior': [valor_original],
        'Fator': [fator_final],
        'Novo Valor': [novo_valor]
    })
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_export.to_excel(writer, index=False, sheet_name='Global')
        # Aqui entra a formatação que combinamos (R$ nas colunas certas)
    
    st.download_button(
        label="📥 Baixar Planilha de Valor Global",
        data=output.getvalue(),
        file_name=f"Valor_Global_{ciclo_final}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )