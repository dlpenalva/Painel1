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
percentual_reajuste = (fator_vigente - 1) 

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("💰 Gestão de Valor Global e Execução")

# --- 1. PARÂMETROS ---
st.header("1. Parâmetros do Reajuste")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Índice", adm['indice'])
c2.metric("Data-Base", adm['data_base'])
c3.metric("Fator Aplicado", f"{fator_vigente:.4f}")
c4.metric("% de Reajuste", f"{percentual_reajuste*100:.2f}%")

# --- FUNÇÃO DE GERAÇÃO DE PLANILHA ---
def gerar_planilha_fiscal():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        
        # Formatos
        fmt_moeda = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        fmt_pct = workbook.add_format({'num_format': '0.00%', 'border': 1, 'align': 'center'})
        fmt_numero = workbook.add_format({'num_format': '0', 'border': 1, 'align': 'center'})
        fmt_data = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1, 'align': 'center'})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_label = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})
        fmt_input = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1})

        # --- ABA 1: ITENS_CICLOS ---
        ws = workbook.add_worksheet("ITENS_CICLOS")
        
        # Cabeçalhos de Resumo (Linhas 1 e 2)
        ws.write('C2', "VALOR TOTAL C0:", fmt_label)
        ws.write_formula('D2', '=SUM(D4:D104)', fmt_moeda) # D2 Automático
        
        ws.write('F1', "DATA REMANESCENTE:", fmt_label)
        ws.write('G1', "", fmt_input) # Fiscal preenche
        
        ws.write('F2', "PERCENTUAL REAJUSTE:", fmt_label)
        ws.write('G2', percentual_reajuste, fmt_pct) # G2 Automático
        
        # Tabela de Itens (Linha 3)
        headers = [
            "Item", "Qtd C0", "VU C0 (R$)", "TOTAL C0 (R$)", "Consumido no Ciclo", 
            " ", "Qtd Remanescente", "VU Reajustado", "TOTAL Reajustado"
        ]
        for col_num, header in enumerate(headers):
            ws.write(2, col_num, header, fmt_header)
        
        ws.set_column('A:A', 8)
        ws.set_column('B:E', 18)
        ws.set_column('F:F', 5) # Espaçador
        ws.set_column('G:I', 22)
        
        for row in range(3, 103):
            xl_row = row + 1
            ws.write(row, 0, row-2, fmt_numero) # A
            ws.write(row, 1, 0, fmt_numero)    # B
            ws.write(row, 2, 0, fmt_moeda)     # C
            ws.write_formula(row, 3, f'=B{xl_row}*C{xl_row}', fmt_moeda) # D (Total C0)
            ws.write_formula(row, 4, f'=B{xl_row}-G{xl_row}', fmt_numero) # E (Consumido)
            
            ws.write(row, 5, " ") # F (Vazio)
            
            ws.write(row, 6, 0, fmt_input) # G (Remanescente - Fiscal)
            ws.write(row, 7, 0, fmt_moeda) # H (VU Reajustado - será processado no Python ou preenchido)
            ws.write_formula(row, 8, f'=G{xl_row}*H{xl_row}', fmt_moeda) # I (Total Reajustado)

        # --- ABA 2: RETROATIVO ---
        ws_retro = workbook.add_worksheet("RETROATIVO")
        ws_retro.write(2, 0, "Competência", fmt_header)
        ws_retro.write(2, 1, "Valor bruto faturado (R$)", fmt_header)
        ws_retro.set_column('A:B', 30)
        ws_retro.write(3, 0, datetime(2024, 5, 1), fmt_data)
        ws_retro.write(3, 1, 0, fmt_moeda)
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
    file_name=f"Coleta_Reajuste_{adm['ciclo_atual']}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

st.divider()

# --- 3. PROCESSAMENTO ---
st.subheader("3. Processamento de Dados")
uploaded_file = st.file_uploader("Suba a planilha preenchida", type="xlsx")

if uploaded_file:
    df_itens = pd.read_excel(uploaded_file, sheet_name="ITENS_CICLOS", skiprows=2)
    df_retro = pd.read_excel(uploaded_file, sheet_name="RETROATIVO", skiprows=2).dropna()
    
    # Validação Básica
    v_orig = df_itens["TOTAL C0 (R$)"].sum()
    faturado_retro = df_retro.iloc[:, 1].sum()
    
    # Cálculo do Remanescente Reajustado (Coluna I)
    # Se o fiscal preencheu a coluna G, calculamos o novo global
    total_remanescente_reaj = (df_itens["Qtd Remanescente"] * df_itens["VU C0 (R$)"] * fator_vigente).sum()
    global_estimado = faturado_retro + total_remanescente_reaj
    
    c1, c2 = st.columns(2)
    c1.metric("Total Retroativo", f"R$ {faturado_retro:,.2f}")
    c2.metric("Novo Global Estimado", f"R$ {global_estimado:,.2f}")
    
    st.session_state['balanco'] = {'orig': v_orig, 'final': global_estimado, 'retro': faturado_retro}

st.divider()

# --- 4. RELATÓRIO ---
st.subheader("4. Finalização e Relatório")
if 'balanco' in st.session_state:
    b = st.session_state['balanco']
    diff = b['final'] - b['orig']
    
    texto_nt = f"""
    1. RELATÓRIO: Análise de reajuste ({adm['ciclo_atual']}) com índice {adm['indice']}.
    2. CÁLCULO: Fator {adm['fator']:.4f} aplicado sobre saldo remanescente.
    - Valor Original: R$ {b['orig']:,.2f}
    - Valor Global Atualizado: R$ {b['final']:,.2f}
    - Diferença Financeira: R$ {diff:,.2f}
    """
    st.text_area("Minuta da NT:", texto_nt, height=200)
else:
    st.info("Aguardando upload para gerar relatório.")