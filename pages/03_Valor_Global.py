import streamlit as st
import pandas as pd
import io

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

fator_vigente = fatores_editados["Fator Acumulado"].iloc[-1]

# --- FUNÇÃO PARA GERAR PLANILHA PADRONIZADA ---
def gerar_planilha_fiscal():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Aba PARAMETROS
        df_params = pd.DataFrame({
            "Parametro": ["Indice", "Data-Base", "Fator Aplicado", "Data Marco"],
            "Valor": [indice_nome, dt_base_orig.strftime('%m/%Y'), fator_vigente, marco_reajuste.strftime('%m/%Y')]
        })
        df_params.to_sheet_name = "PARAMETROS"
        df_params.to_excel(writer, sheet_name="PARAMETROS", index=False)

        # Aba ITENS_CICLOS (Modelo)
        df_modelo_itens = pd.DataFrame(columns=["ID/Item", "Descrição", "Unidade", "VU C0 (R$)", f"Qtd remanescente em {marco_reajuste.strftime('%m/%Y')} (PREENCHER)"])
        df_modelo_itens.to_excel(writer, sheet_name="ITENS_CICLOS", index=False, startrow=2)

        # Aba RETROATIVO (Modelo)
        df_modelo_retro = pd.DataFrame(columns=["Competência", "Valor bruto faturado após descontos (R$)"])
        df_modelo_retro.to_excel(writer, sheet_name="RETROATIVO", index=False, startrow=2)

    return output.getvalue()

st.divider()

# --- BOTÃO DE EXPORTAÇÃO ---
st.subheader("2. Preparação de Dados")
st.write("Gere a planilha padronizada com os parâmetros acima para enviar ao fiscal.")
excel_data = gerar_planilha_fiscal()
st.download_button(
    label="📥 Baixar Planilha para o Fiscal",
    data=excel_data,
    file_name=f"Coleta_Valor_Global_{indice_nome}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.divider()

# --- NAVEGAÇÃO ---
tab_financeira, tab_estoque, tab_comparativo = st.tabs(["📊 Apuração Financeira", "📦 Controle de Estoque", "⚖️ Comparativo"])

with tab_estoque:
    st.subheader("📦 Processamento da Planilha Retornada")
    uploaded_file = st.file_uploader("Suba a planilha PREENCHIDA pelo fiscal", type="xlsx")
    # ... (O resto do código de leitura continua igual) ...