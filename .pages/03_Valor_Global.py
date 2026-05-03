import streamlit as st
import pandas as pd
import io

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Valor Global - Homologação", layout="wide")

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Gestão de Valor Global e Execução")

# --- BLOCO 1: RESUMO DA ADMISSIBILIDADE (SOMENTE LEITURA) ---
# Aqui o sistema busca o que foi feito na Etapa A (01_Admissibilidade)
if 'dados_admissibilidade' not in st.session_state:
    # Valores padrão caso o usuário pule etapas no teste
    st.session_state['dados_admissibilidade'] = {
        'indice': "IST",
        'data_base': "01/05/2025",
        'fator': 1.0468,
        'ciclo_atual': "C1"
    }

adm = st.session_state['dados_admissibilidade']

st.header("1. Parâmetros de Reajuste")
st.info("Estes dados foram importados automaticamente da Etapa de Admissibilidade.")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Índice", adm['indice'])
c2.metric("Data-Base", adm['data_base'])
c3.metric("Fator Aplicado", f"{adm['fator']:.4f}")
c4.metric("Ciclo", adm['ciclo_atual'])

# Variáveis globais para o script
fator_vigente = adm['fator']
indice_nome = adm['indice']

# --- FUNÇÃO PARA GERAR PLANILHA COM FORMATAÇÃO CORRIGIDA ---
def gerar_planilha_fiscal():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # DEFINIÇÃO DE FORMATOS
        fmt_moeda = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'align': 'center'})
        fmt_numero = workbook.add_format({'num_format': '0', 'border': 1, 'align': 'center'})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_input = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'align': 'center'})

        # ABA ITENS_CICLOS
        ws_itens = workbook.add_worksheet("ITENS_CICLOS")
        headers = ["Item", "Quantidade", "VU C0 (R$)", "TOTAL C0 (R$)", "Qtd remanescente (PREENCHER)"]
        
        for col_num, header in enumerate(headers):
            ws_itens.write(2, col_num, header, fmt_header)
        
        ws_itens.set_column('A:A', 10) # Coluna Item
        ws_itens.set_column('B:E', 25) # Demais colunas
        
        # Preenchimento de 10 linhas de exemplo
        for row in range(3, 13):
            ws_itens.write(row, 0, row-2, fmt_numero) # COLUNA A: Agora é número puro
            ws_itens.write(row, 1, 0, fmt_numero)    # Coluna B: Qtd
            ws_itens.write(row, 2, 0, fmt_moeda)     # COLUNA C: Agora é R$
            ws_itens.write_formula(row, 3, f'=B{row+1}*C{row+1}', fmt_moeda) # Coluna D: Total R$
            ws_itens.write(row, 4, 0, fmt_input)     # Coluna E: Preenchimento
        
        # ABA RETROATIVO
        ws_retro = workbook.add_worksheet("RETROATIVO")
        ws_retro.write(0, 0, "Preencha os valores faturados mês a mês:", workbook.add_format({'bold': True}))
        ws_retro.write(2, 0, "Competência", fmt_header)
        ws_retro.write(2, 1, "Valor bruto faturado (R$)", fmt_header)
        ws_retro.set_column('A:B', 30)
        
        for row in range(3, 15):
            ws_retro.write(row, 0, "", fmt_numero)
            ws_retro.write(row, 1, 0, fmt_moeda) # COLUNA B: Agora é R$

    return output.getvalue()

st.divider()

# --- 2. PREPARAÇÃO ---
st.subheader("2. Download da Planilha de Coleta")
st.download_button(
    label="📥 Gerar Planilha para o Fiscal",
    data=gerar_planilha_fiscal(),
    file_name=f"Coleta_Global_{indice_nome}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.divider()

# --- 3. PROCESSAMENTO ---
tab1, tab2 = st.tabs(["📦 Processar Retorno", "⚖️ Comparativo"])

with tab1:
    uploaded_file = st.file_uploader("Suba a planilha preenchida pelo fiscal", type="xlsx")
    if uploaded_file:
        df_itens = pd.read_excel(uploaded_file, sheet_name="ITENS_CICLOS", skiprows=2).dropna(subset=["Item"])
        df_retro = pd.read_excel(uploaded_file, sheet_name="RETROATIVO", skiprows=2).dropna(axis=0, how='all')

        # Cálculos
        v_orig = (df_itens["Quantidade"] * df_itens["VU C0 (R$)"]).sum()
        faturado = df_retro.iloc[:, 1].sum()
        col_rem = df_itens.columns[-1]
        remanescente_reaj = (df_itens[col_rem] * df_itens["VU C0 (R$)"] * fator_vigente).sum()
        global_estimado = faturado + remanescente_reaj

        st.metric("VALOR GLOBAL ESTIMADO", f"R$ {global_estimado:,.2f}")
        st.session_state['balanco'] = {'orig': v_orig, 'final': global_estimado, 'fat': faturado}

with tab2:
    if 'balanco' in st.session_state:
        b = st.session_state['balanco']
        diff = b['final'] - b['orig']
        perc = (diff / b['orig'] * 100) if b['orig'] > 0 else 0
        
        st.header("Balanço do Contrato")
        st.metric("Variação do Valor Global", f"R$ {diff:,.2f}", delta=f"{perc:.2f}%")
        
        st.write(f"""
        **Resumo da Lógica:**
        1. **Valor de Origem (C0):** R$ {b['orig']:,.2f}
        2. **Execução Realizada:** R$ {b['fat']:,.2f}
        3. **Projeção do Remanescente:** R$ {b['final'] - b['fat']:,.2f} (aplicado fator {fator_vigente})
        """)