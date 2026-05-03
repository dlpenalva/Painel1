import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Gestão de Valor Global", layout="wide")

# --- RECUPERAÇÃO DE DADOS DO BLOCO A ---
if 'dados_admissibilidade' not in st.session_state:
    st.session_state['dados_admissibilidade'] = {
        'indice': "IST",
        'data_base': "01/05/2025",
        'fator': 1.0468,
        'ciclo_atual': "C1"
    }

adm = st.session_state['dados_admissibilidade']
fator_vigente = adm['fator']

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("💰 Gestão de Valor Global e Execução")

# --- 1. PARÂMETROS ---
st.header("1. Parâmetros do Reajuste")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Índice", adm['indice'])
c2.metric("Data-Base", adm['data_base'])
c3.metric("Fator Aplicado", f"{adm['fator']:.4f}")
c4.metric("Ciclo", adm['ciclo_atual'])

# --- FUNÇÃO DE GERAÇÃO DE PLANILHA OTIMIZADA ---
def gerar_planilha_fiscal():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Formatos
        fmt_moeda = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        fmt_numero = workbook.add_format({'num_format': '0', 'border': 1})
        fmt_data = workbook.add_format({'num_format': 'mm/yyyy', 'border': 1, 'align': 'center'})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_input = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1})

        # --- ABA 1: ITENS_CICLOS (100 LINHAS) ---
        ws_itens = workbook.add_worksheet("ITENS_CICLOS")
        headers = ["Item", "Quantidade", "VU C0 (R$)", "TOTAL C0 (R$)", "Qtd remanescente (PREENCHER)"]
        for col_num, header in enumerate(headers):
            ws_itens.write(2, col_num, header, fmt_header)
        
        ws_itens.set_column('A:A', 10)
        ws_itens.set_column('B:E', 25)
        
        # Gerar 100 linhas com fórmulas
        for row in range(3, 103):
            ws_itens.write(row, 0, row-2, fmt_numero) # ID
            ws_itens.write(row, 1, 0, fmt_numero)    # Qtd
            ws_itens.write(row, 2, 0, fmt_moeda)     # VU
            ws_itens.write_formula(row, 3, f'=B{row+1}*C{row+1}', fmt_moeda) # Total C0
            ws_itens.write(row, 4, 0, fmt_input)     # Campo Fiscal

        # --- ABA 2: RETROATIVO (36 MESES AUTOMÁTICOS) ---
        ws_retro = workbook.add_worksheet("RETROATIVO")
        ws_retro.write(2, 0, "Competência (Preencha a 1ª)", fmt_header)
        ws_retro.write(2, 1, "Valor bruto faturado (R$)", fmt_header)
        ws_retro.set_column('A:B', 35)

        # Primeira data em A4 (Row 3, Col 0)
        ws_retro.write(3, 0, datetime(2024, 5, 1), fmt_data) # Exemplo inicial
        ws_retro.write(3, 1, 0, fmt_input)

        # Fórmulas para os próximos 36 meses
        for row in range(4, 40):
            # Fórmula Excel que adiciona 1 mês à célula anterior
            ws_retro.write_formula(row, 0, f'=EDATE(A{row}, 1)', fmt_data)
            ws_retro.write(row, 1, 0, fmt_input)

    return output.getvalue()

st.divider()

# --- 2. DOWNLOAD ---
st.subheader("2. Download da Planilha de Coleta")
st.download_button(
    label="📥 Baixar Planilha Otimizada (100 itens + 36 meses)",
    data=gerar_planilha_fiscal(),
    file_name="Coleta_Global_Otimizada.xlsx"
)

st.divider()

# --- 3. PROCESSAMENTO ---
tab1, tab2 = st.tabs(["📦 Processar Retorno", "⚖️ Comparativo"])

with tab1:
    uploaded_file = st.file_uploader("Suba a planilha preenchida", type="xlsx")
    if uploaded_file:
        df_itens = pd.read_excel(uploaded_file, sheet_name="ITENS_CICLOS", skiprows=2).dropna(subset=["Item"])
        df_retro = pd.read_excel(uploaded_file, sheet_name="RETROATIVO", skiprows=2).dropna(subset=["Valor bruto faturado (R$)"])

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
        # Proteção contra divisão por zero e cálculo de variação real
        perc = (diff / b['orig'] * 100) if b['orig'] > 0 else 0
        
        st.header("Balanço do Contrato")
        st.metric("Variação do Valor Global", f"R$ {diff:,.2f}", delta=f"{perc:.2f}%")
        
        if perc > 25:
            st.warning("Atenção: A variação ultrapassa o limite legal de 25% (Art. 81 da Lei 13.303).")