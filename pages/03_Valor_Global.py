import streamlit as st
import pandas as pd
import io

st.title("💰 Valor Global do Contrato")

# Verificação de Integração com Bloco A
if 'dados_admissibilidade' not in st.session_state:
    st.warning("⚠️ Admissibilidade não detectada. Realize o cálculo no Bloco A primeiro.")
    st.stop()

adm = st.session_state['dados_admissibilidade']
tipo = adm['tipo']

st.info(f"✅ **Modo de Execução:** Reajuste {tipo}")

# --- LOGICA DE GERAÇÃO DE EXCEL PARA DOWNLOAD ---
def gerar_excel_coleta(dados_adm):
    output = io.BytesIO()
    
    # Criando DataFrame estruturado para o Fiscal
    if dados_adm['tipo'] == 'Simples':
        df_modelo = pd.DataFrame({
            'Item': ['Exemplo: Link Dedicado'],
            'Data-Marco': [dados_adm['data_base']],
            'Percentual Reajuste (%)': [(dados_adm['fator'] - 1) * 100],
            'Valor Unitário Atual': [0.0],
            'Quantidade Remanescente': [0]
        })
    else:
        # Para Múltiplos, criamos linhas para cada ciclo
        ciclos = dados_adm['detalhamento_ciclos']
        df_modelo = pd.DataFrame([{
            'Ciclo': c['Ciclo'],
            'Data-Marco': c['Data-Base'],
            'Percentual Ciclo (%)': (c['Fator'] - 1) * 100,
            'Valor Unitário': 0.0,
            'Quantidade': 0
        } for c in ciclos])
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_modelo.to_excel(writer, index=False, sheet_name='Coleta_Dados')
    
    return output.getvalue()

# --- INTERFACE DO BLOCO B ---
col_a, col_b = st.columns(2)

with col_a:
    st.subheader("1. Download de Coleta")
    st.write("Baixe a planilha pré-preenchida com os parâmetros da Admissibilidade.")
    
    excel_data = gerar_excel_coleta(adm)
    st.download_button(
        label="📥 Baixar Planilha de Itens",
        data=excel_data,
        file_name=f"Coleta_Reajuste_{tipo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

with col_b:
    st.subheader("2. Upload de Resultados")
    upload = st.file_uploader("Upload da planilha preenchida pelo Fiscal", type=["xlsx"])

if upload:
    df_resultado = pd.read_excel(upload)
    st.success("Dados carregados! Iniciando cálculo do impacto financeiro global...")
    st.dataframe(df_resultado)
    
    # Cálculo Global (Lógica baseada nos percentuais herdados)
    if tipo == 'Múltiplo':
        st.write("### Impacto por Ciclo")
        # Aqui o sistema aplica os fatores em cascata sobre o valor total
        fator_final = adm['fator']
        st.metric("Fator Final Aplicado", f"{fator_final:.4f}")