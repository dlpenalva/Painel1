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
    # Tabela editável para o usuário conferir os fatores da admissibilidade
    df_fatores_base = pd.DataFrame({
        "Ciclo": [f"C{i}" for i in range(qtd_ciclos + 1)],
        "Fator Acumulado": [1.0000] + [1.0468] * qtd_ciclos 
    })
    fatores_editados = st.data_editor(df_fatores_base, hide_index=True, use_container_width=True)

fator_vigente = fatores_editados["Fator Acumulado"].iloc[-1]

# --- FUNÇÃO PARA GERAR PLANILHA PADRONIZADA (CONFORME 1.XLSX) ---
def gerar_planilha_fiscal():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        fmt_moeda = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
        
        # Aba PARAMETROS
        df_params = pd.DataFrame({
            "Parametro": ["Indice", "Data-Base", "Fator Aplicado", "Data Marco"],
            "Valor": [indice_nome, dt_base_orig.strftime('%m/%Y'), fator_vigente, marco_reajuste.strftime('%m/%Y')]
        })
        df_params.to_excel(writer, sheet_name="PARAMETROS", index=False)

        # Aba ITENS_CICLOS (Estrutura Sincronizada)
        # Removida coluna 'Descrição' conforme solicitado
        cols_itens = ["Item/ Bloco", "Quantidade", "VU C0 (R$)", "TOTAL C0 (R$)", f"Qtd remanescente em {marco_reajuste.strftime('%m/%Y')} (PREENCHER)"]
        df_modelo_itens = pd.DataFrame(columns=cols_itens)
        df_modelo_itens.to_excel(writer, sheet_name="ITENS_CICLOS", index=False, startrow=2)
        
        ws_itens = writer.sheets['ITENS_CICLOS']
        ws_itens.set_column('C:D', 18, fmt_moeda) # Colunas VU e TOTAL
        ws_itens.write(1, 0, "DADOS CONTRATUAIS (PREENCHER)", fmt_header)

        # Aba RETROATIVO
        df_modelo_retro = pd.DataFrame(columns=["Competência", "Valor bruto faturado após descontos (R$)"])
        df_modelo_retro.to_excel(writer, sheet_name="RETROATIVO", index=False, startrow=2)
        ws_retro = writer.sheets['RETROATIVO']
        ws_retro.set_column('B:B', 30, fmt_moeda)

    return output.getvalue()

st.divider()

# --- 2. PREPARAÇÃO E EXPORTAÇÃO ---
st.subheader("2. Preparação de Dados")
st.download_button(
    label="📥 Baixar Planilha para o Fiscal",
    data=gerar_planilha_fiscal(),
    file_name=f"Coleta_Valor_Global_{indice_nome}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.divider()

# --- 3. IMPORTAÇÃO E PROCESSAMENTO ---
tab_financeira, tab_estoque, tab_comparativo = st.tabs(["📊 Apuração Financeira", "📦 Controle de Estoque", "⚖️ Comparativo"])

with tab_estoque:
    st.subheader("📦 Processamento da Planilha Retornada")
    uploaded_file = st.file_uploader("Suba a planilha PREENCHIDA pelo fiscal", type="xlsx")

    if uploaded_file:
        try:
            # Lendo a aba de ITENS (Pula 2 linhas para pegar o cabeçalho real)
            df_itens = pd.read_excel(uploaded_file, sheet_name="ITENS_CICLOS", skiprows=2)
            
            # Cálculo do Valor Total do Contrato Original
            # Tenta ler 'TOTAL C0 (R$)' ou calcula se estiver vazio
            if "TOTAL C0 (R$)" in df_itens.columns:
                v_total_contrato = df_itens["TOTAL C0 (R$)"].sum()
            else:
                v_total_contrato = (df_itens["Quantidade"] * df_itens["VU C0 (R$)"]).sum()

            # Localiza coluna de preenchimento do fiscal (dinâmico)
            col_rem = [c for c in df_itens.columns if "PREENCHER" in str(c)][-1]
            
            # Cálculo do Saldo Remanescente Reajustado
            # Fórmula: Qtd_Remanescente * Valor_Unitário_C0 * Fator_da_Tela
            saldo_reajustado = (df_itens[col_rem] * df_itens["VU C0 (R$)"] * fator_vigente).sum()

            # Leitura da Aba Retroativo
            df_retro = pd.read_excel(uploaded_file, sheet_name="RETROATIVO", skiprows=2)
            total_faturado_real = df_retro.iloc[:, 1].sum()

            # CONSOLIDAÇÃO
            valor_global_real = total_faturado_real + saldo_reajustado

            st.success("Planilha processada com sucesso!")
            
            st.metric("Valor Total Original (C0)", f"R$ {v_total_contrato:,.2f}")
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Faturado (Real)", f"R$ {total_faturado_real:,.2f}")
            c2.metric("Saldo Remanescente (Reaj.)", f"R$ {saldo_reajustado:,.2f}")
            c3.metric("VALOR GLOBAL ESTIMADO", f"R$ {valor_global_real:,.2f}")

            # Salva no estado da sessão para as outras abas
            st.session_state['vg_data'] = {
                'original': v_total_contrato,
                'real': valor_global_real,
                'faturado': total_faturado_real
            }

        except Exception as e:
            st.error(f"Erro ao processar: {e}. Certifique-se de preencher as colunas Quantidade e VU.")

with tab_financeira:
    if 'vg_data' in st.session_state:
        st.write(f"### Detalhamento Financeiro")
        st.info(f"O valor faturado reportado pelo fiscal foi de **R$ {st.session_state['vg_data']['faturado']:,.2f}**.")
    else:
        st.warning("Aguardando upload na aba de estoque.")

with tab_comparativo:
    if 'vg_data' in st.session_state:
        st.subheader("Balanço do Contrato")
        diff = st.session_state['vg_data']['real'] - st.session_state['vg_data']['original']
        st.metric("Variação do Valor Global", f"R$ {diff:,.2f}", delta=f"{((diff/st.session_state['vg_data']['original'])*100):.2f}%")