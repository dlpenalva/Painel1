import streamlit as st
import pandas as pd
import io
from datetime import datetime

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
percentual_reajuste = (fator_vigente - 1) * 100

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("💰 Gestão de Valor Global e Execução")

# --- 1. PARÂMETROS ---
st.header("1. Parâmetros do Reajuste")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Índice", adm['indice'])
c2.metric("Data-Base", adm['data_base'])
c3.metric("Fator Aplicado", f"{fator_vigente:.4f}")
c4.metric("% de Reajuste", f"{percentual_reajuste:.2f}%")

# --- FUNÇÃO DE GERAÇÃO DE PLANILHA ---
def gerar_planilha_fiscal():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Formatos
        fmt_moeda = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        fmt_numero = workbook.add_format({'num_format': '0', 'border': 1, 'align': 'center'})
        fmt_data = workbook.add_format({'num_format': 'mm/yyyy', 'border': 1, 'align': 'center'})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_label_yellow = workbook.add_format({'bold': True, 'bg_color': '#FFFF00', 'border': 1, 'align': 'center'})
        fmt_input = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1})

        # --- ABA 1: ITENS_CICLOS ---
        ws_itens = workbook.add_worksheet("ITENS_CICLOS")
        
        # Informação em E2
        ws_itens.write('E2', "DATA DO REMANESCENTE", fmt_label_yellow)
        
        headers = ["Item", "Quantidade", "VU C0 (R$)", "TOTAL C0 (R$)", "Qtd remanescente (PREENCHER)", "CONSUMIDO NO CICLO (R$)"]
        for col_num, header in enumerate(headers):
            ws_itens.write(2, col_num, header, fmt_header)
        
        ws_itens.set_column('A:A', 10)
        ws_itens.set_column('B:F', 25)
        
        for row in range(3, 103):
            xl_row = row + 1
            ws_itens.write(row, 0, row-2, fmt_numero) # A: Item
            ws_itens.write(row, 1, 0, fmt_numero)    # B: Qtd
            ws_itens.write(row, 2, 0, fmt_moeda)     # C: VU
            ws_itens.write_formula(row, 3, f'=B{xl_row}*C{xl_row}', fmt_moeda) # D: Total
            ws_itens.write(row, 4, 0, fmt_input)     # E: Qtd Remanescente
            # F: Valor Consumido = (Qtd Total - Qtd Remanescente) * VU
            ws_itens.write_formula(row, 5, f'=(B{xl_row}-E{xl_row})*C{xl_row}', fmt_moeda)

        # --- ABA 2: RETROATIVO ---
        ws_retro = workbook.add_worksheet("RETROATIVO")
        ws_retro.write(2, 0, "Competência (Preencha a 1ª)", fmt_header)
        ws_retro.write(2, 1, "Valor bruto faturado (R$)", fmt_header)
        ws_retro.set_column('A:B', 35)
        
        ws_retro.write(3, 0, datetime(2024, 5, 1), fmt_data)
        ws_retro.write(3, 1, 0, fmt_moeda) # Coluna B agora como moeda
        
        for row in range(4, 40):
            ws_retro.write_formula(row, 0, f'=EDATE(A{row}, 1)', fmt_data)
            ws_retro.write(row, 1, 0, fmt_moeda)

    return output.getvalue()

st.divider()

# --- 2. DOWNLOAD ---
st.subheader("2. Download da Planilha de Coleta")
st.download_button(
    label="📥 Gerar Planilha para o Fiscal",
    data=gerar_planilha_fiscal(),
    file_name=f"Coleta_Global_{adm['ciclo_atual']}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# ... (Manter o restante do código de Processamento e Relatório igual ao anterior)