import streamlit as st
import pandas as pd
import io
from datetime import datetime

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="Gestão de Valor Global", layout="wide")

if 'dados_admissibilidade' not in st.session_state:
    st.session_state['dados_admissibilidade'] = {
        'indice': "IST", 'data_base': "01/05/2025", 'fator': 1.0468, 'ciclo_atual': "C1"
    }

adm = st.session_state['dados_admissibilidade']
fator = adm['fator']
reajuste_pct = (fator - 1)

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("💰 Gestão de Valor Global e Execução")

# --- 1. PARÂMETROS ---
st.header("1. Parâmetros do Reajuste")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Índice", adm['indice'])
c2.metric("Data-Base", adm['data_base'])
c3.metric("Fator Aplicado", f"{fator:.4f}")
c4.metric("% de Reajuste", f"{reajuste_pct*100:.2f}%")

# --- 2. PLANILHA DE COLETA ---
def gerar_planilha():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        wb = writer.book
        # Formatos
        fmt_moeda = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        fmt_pct = wb.add_format({'num_format': '0.00%', 'border': 1, 'align': 'center'})
        fmt_header = wb.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_input = wb.add_format({'bg_color': '#FFF2CC', 'border': 1})
        fmt_label = wb.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})

        ws = wb.add_worksheet("ITENS_CICLOS")
        ws.write('C2', "VALOR TOTAL C0:", fmt_label)
        ws.write_formula('D2', '=SUM(D4:D104)', fmt_moeda)
        ws.write('F1', "DATA REMANESCENTE:", fmt_label)
        ws.write('G1', "", fmt_input)
        ws.write('F2', "PERCENTUAL REAJUSTE:", fmt_label)
        ws.write('G2', reajuste_pct, fmt_pct)

        headers = ["Item", "Qtd C0", "VU C0 (R$)", "TOTAL C0 (R$)", "Consumido no Ciclo (Físico)", " ", "Qtd Remanescente", "VU Reajustado", "TOTAL Reajustado"]
        for col, h in enumerate(headers): ws.write(2, col, h, fmt_header)
        
        for row in range(3, 103):
            r = row + 1
            ws.write(row, 0, row-2)
            ws.write(row, 1, 0)
            ws.write(row, 2, 0, fmt_moeda)
            ws.write_formula(row, 3, f'=B{r}*C{r}', fmt_moeda)
            ws.write_formula(row, 4, f'=B{r}-G{r}', wb.add_format({'border': 1}))
            ws.write(row, 6, 0, fmt_input)
            ws.write(row, 7, 0, fmt_moeda)
            ws.write_formula(row, 8, f'=G{r}*H{r}', fmt_moeda)

        ws_retro = wb.add_worksheet("RETROATIVO")
        ws_retro.write(2, 0, "Competência", fmt_header)
        ws_retro.write(2, 1, "Valor bruto faturado (R$)", fmt_header)
        ws_retro.write(3, 0, datetime(2024, 5, 1), wb.add_format({'num_format': 'mm/yyyy'}))
        ws_retro.write(3, 1, 0, fmt_moeda)
    return output.getvalue()

st.header("2. Download da Planilha de Coleta")
st.download_button("📥 Gerar Planilha para o Fiscal", gerar_planilha(), f"Coleta_{adm['ciclo_atual']}.xlsx")

st.divider()

# --- 3. PROCESSAMENTO E COMPARATIVO ---
st.header("3. Processamento e Comparativo")
file = st.file_uploader("Suba a planilha preenchida", type="xlsx")

if file:
    df_itens = pd.read_excel(file, sheet_name="ITENS_CICLOS", skiprows=2).dropna(subset=["Item"])
    df_retro = pd.read_excel(file, sheet_name="RETROATIVO", skiprows=2).dropna(subset=["Valor bruto faturado (R$)"])
    
    # Cálculos de Estoque vs Financeiro
    valor_fisico_consumido = (df_itens["Consumido no Ciclo (Físico)"] * df_itens["VU C0 (R$)"]).sum()
    valor_financeiro_real = df_retro.iloc[:, 1].sum()
    delta_consumo = valor_fisico_consumido - valor_financeiro_real
    
    # Cálculos de Represado (Retroativo do Reajuste)
    valor_retroativo_reajustado = valor_financeiro_real * fator
    represado_total = valor_retroativo_reajustado - valor_financeiro_real
    
    # Saldo Remanescente
    saldo_remanescente_reaj = (df_itens["Qtd Remanescente"] * df_itens["VU C0 (R$)"] * fator).sum()
    novo_global = valor_financeiro_real + represado_total + saldo_remanescente_reaj

    # Exibição do Comparativo
    t1, t2 = st.tabs(["📊 Resumo Financeiro", "🔍 Detalhes do Delta"])
    with t1:
        col_a, col_b, col_c = st.columns(3)
        col_a.metric("Consumo Físico (Estoque)", f"R$ {valor_fisico_consumido:,.2f}")
        col_b.metric("Faturamento Real (Financeiro)", f"R$ {valor_financeiro_real:,.2f}", delta=f"Delta: R$ {delta_consumo:,.2f}", delta_color="inverse")
        col_c.metric("Novo Global Estimado", f"R$ {novo_global:,.2f}")

    with t2:
        st.write("**Análise de Represados:**")
        st.info(f"O valor de R$ {represado_total:,.2f} refere-se à diferença do reajuste ({reajuste_pct*100:.2f}%) sobre o que já foi faturado.")

    st.divider()

    # --- 4. RELATÓRIO COMPLETO ---
    st.header("4. Finalização e Relatório")
    
    relatorio = f"""NOTA TÉCNICA - REAJUSTE CONTRATUAL {adm['ciclo_atual']}

1. OBJETO E PARÂMETROS
Reajuste baseado no índice {adm['indice']} com fator de {fator:.4f}.

2. CONCILIAÇÃO FÍSICO-FINANCEIRA (ESTOQUE VS FINANCEIRO)
- Valor Consumido (Físico/Estoque): R$ {valor_fisico_consumido:,.2f}
- Valor Faturado (Real/Financeiro): R$ {valor_financeiro_real:,.2f}
- Diferença (Delta de Consumo): R$ {delta_consumo:,.2f}

3. MEMÓRIA DE CÁLCULO DO REAJUSTE
- Valores Represados (Retroativo): R$ {represado_total:,.2f} (Diferença de {reajuste_pct*100:.2f}% sobre faturamento real).
- Saldo Remanescente Reajustado: R$ {saldo_remanescente_reaj:,.2f}
- VALOR GLOBAL ESTIMADO PÓS-REAJUSTE: R$ {novo_global:,.2f}

4. CONCLUSÃO
O impacto total do reajuste no contrato é de R$ {represado_total + (saldo_remanescente_reaj / fator * reajuste_pct):,.2f} considerando saldo e retroativo.
"""
    st.text_area("Minuta para o SEI (Copie abaixo):", relatorio, height=400)
    st.download_button("📄 Baixar Relatório (.txt)", relatorio, file_name=f"Relatorio_Reajuste_{adm['ciclo_atual']}.txt")

else:
    st.info("Aguardando upload da planilha para gerar o comparativo e relatório.")