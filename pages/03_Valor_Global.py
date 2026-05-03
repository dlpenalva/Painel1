import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- CONFIGURAÇÃO (Sempre no topo) ---
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
        fmt_moeda = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        fmt_numero = workbook.add_format({'num_format': '0', 'border': 1, 'align': 'center'})
        fmt_data = workbook.add_format({'num_format': 'mm/yyyy', 'border': 1, 'align': 'center'})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_input = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1})

        ws_itens = workbook.add_worksheet("ITENS_CICLOS")
        headers = ["Item", "Quantidade", "VU C0 (R$)", "TOTAL C0 (R$)", "Qtd remanescente (PREENCHER)"]
        for col_num, header in enumerate(headers):
            ws_itens.write(2, col_num, header, fmt_header)
        
        ws_itens.set_column('A:A', 10)
        ws_itens.set_column('B:E', 25)
        for row in range(3, 103):
            ws_itens.write(row, 0, row-2, fmt_numero)
            ws_itens.write(row, 1, 0, fmt_numero)
            ws_itens.write(row, 2, 0, fmt_moeda)
            ws_itens.write_formula(row, 3, f'=B{row+1}*C{row+1}', fmt_moeda)
            ws_itens.write(row, 4, 0, fmt_input)

        ws_retro = workbook.add_worksheet("RETROATIVO")
        ws_retro.write(2, 0, "Competência (Preencha a 1ª)", fmt_header)
        ws_retro.write(2, 1, "Valor bruto faturado (R$)", fmt_header)
        ws_retro.set_column('A:B', 35)
        ws_retro.write(3, 0, datetime(2024, 5, 1), fmt_data)
        ws_retro.write(3, 1, 0, fmt_input)
        for row in range(4, 40):
            ws_retro.write_formula(row, 0, f'=EDATE(A{row}, 1)', fmt_data)
            ws_retro.write(row, 1, 0, fmt_input)

    return output.getvalue()

st.divider()

# --- 2. DOWNLOAD ---
st.subheader("2. Download da Planilha de Coleta")
st.download_button(
    label="📥 Gerar Planilha para o Fiscal",
    data=gerar_planilha_fiscal(),
    file_name=f"Coleta_Global_{adm['ciclo_atual']}.xlsx"
)

st.divider()

# --- 3. PROCESSAMENTO ---
tab1, tab2 = st.tabs(["📦 Processar Retorno", "⚖️ Comparativo"])

with tab1:
    uploaded_file = st.file_uploader("Suba a planilha preenchida pelo fiscal", type="xlsx")
    if uploaded_file:
        st.success("Planilha processada com sucesso!")
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
        perc = (diff / b['orig'] * 100) if b['orig'] > 0 else 0
        st.header("Balanço do Contrato")
        st.metric("Variação do Valor Global", f"R$ {diff:,.2f}", delta=f"{perc:.2f}%")

# --- 4. GERAÇÃO DE RELATÓRIO ---
st.divider()
st.subheader("4. Finalização e Relatório")

if 'balanco' in st.session_state:
    b = st.session_state['balanco']
    diff = b['final'] - b['orig']
    perc = (diff / b['orig'] * 100) if b['orig'] > 0 else 0
    
    texto_nt = f"""
    NOTA TÉCNICA - REAJUSTE CONTRATUAL {adm['ciclo_atual']}
    
    1. RELATÓRIO
    Trata-se de análise de reajuste de preços para o contrato em epígrafe, 
    utilizando o índice {adm['indice']} com data-base em {adm['data_base']}.
    
    2. ANÁLISE TÉCNICA
    Após processamento da planilha de coleta enviada pela fiscalização, 
    apurou-se um fator de reajuste de {adm['fator']:.4f}. 
    O impacto financeiro foi calculado com base no faturamento retroativo 
    e no saldo remanescente reajustado.
    
    - Valor Original do Contrato: R$ {b['orig']:,.2f}
    - Valor Global Estimado Pós-Reajuste: R$ {b['final']:,.2f}
    - Impacto Financeiro Total: R$ {diff:,.2f}
    - Variação Percentual: {perc:.2f}%
    
    3. CONCLUSÃO
    Considerando que a variação de {perc:.2f}% decorre exclusivamente da 
    aplicação de índice de preços previsto contratualmente, a alteração 
    caracteriza-se como reajustamento de preços em sentido estrito, 
    não computando para fins de limite de aditamentos previsto no 
    Art. 81 da Lei 13.303/2016. 
    
    Sugere-se o prosseguimento do feito para formalização via Apostilamento.
    """
    
    with st.expander("Visualizar Minuta da Nota Técnica"):
        st.text_area("Copie o texto abaixo para o seu SEI/Documento:", texto_nt, height=400)
        
    st.download_button(
        label="📄 Baixar Relatório em TXT",
        data=texto_nt,
        file_name=f"NT_Reajuste_{adm['ciclo_atual']}.txt",
        mime="text/plain"
    )
else:
    st.info("Processe uma planilha no Passo 3 para gerar o relatório.")