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
    qtd_ciclos = st.number_input("Quantidade de Ciclos:", min_value=1, max_value=10, value=1)
    marco_reajuste = st.date_input("Marco do Último Reajuste:", format="DD/MM/YYYY")

with col3:
    st.markdown("**Fatores de Reajuste (Referência)**")
    df_fatores_base = pd.DataFrame({
        "Ciclo": [f"C{i}" for i in range(qtd_ciclos + 1)],
        "Fator Acumulado": [1.0000] + [1.0468] * qtd_ciclos 
    })
    fatores_editados = st.data_editor(df_fatores_base, hide_index=True, use_container_width=True)

fator_vigente = fatores_editados["Fator Acumulado"].iloc[-1]

# --- FUNÇÃO PARA GERAR PLANILHA COM FÓRMULAS E LAYOUT ---
def gerar_planilha_fiscal():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        # Formatos Estéticos
        fmt_moeda = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1, 'align': 'center'})
        fmt_header = workbook.add_format({'bold': True, 'bg_color': '#1F4E78', 'font_color': 'white', 'border': 1, 'align': 'center'})
        fmt_label = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
        fmt_input = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1, 'align': 'center'}) # Amarelo para preenchimento

        # 1. ABA PARAMETROS
        df_params = pd.DataFrame({"Parâmetro": ["Índice", "Data-Base", "Fator Aplicado"], "Valor": [indice_nome, dt_base_orig.strftime('%m/%Y'), fator_vigente]})
        df_params.to_excel(writer, sheet_name="PARAMETROS", index=False)

        # 2. ABA ITENS_CICLOS (COM FÓRMULAS)
        ws_itens = workbook.add_worksheet("ITENS_CICLOS")
        headers = ["Item/ Bloco", "Quantidade", "VU C0 (R$)", "TOTAL C0 (R$)", f"Qtd remanescente em {marco_reajuste.strftime('%m/%Y')} (PREENCHER)"]
        
        for col_num, header in enumerate(headers):
            ws_itens.write(2, col_num, header, fmt_header)
        
        ws_itens.set_column('A:A', 15)
        ws_itens.set_column('B:E', 25)
        
        # Inserir 10 linhas de exemplo com fórmulas de multiplicação
        for row in range(3, 13):
            ws_itens.write(row, 0, row-2, fmt_moeda) # ID
            ws_itens.write_formula(row, 3, f'=B{row+1}*C{row+1}', fmt_moeda) # Fórmulas automáticas B*C
            ws_itens.write(row, 4, 0, fmt_input) # Campo para fiscal preencher
        
        ws_itens.write(1, 0, "DADOS CONTRATUAIS (PREENCHA QUANTIDADE E VU)", fmt_label)

        # 3. ABA RETROATIVO
        ws_retro = workbook.add_worksheet("RETROATIVO")
        ws_retro.write(2, 0, "Competência", fmt_header)
        ws_retro.write(2, 1, "Valor bruto faturado após descontos (R$)", fmt_header)
        ws_retro.set_column('A:B', 35)
        for row in range(3, 15):
            ws_retro.write(row, 1, 0, fmt_input)

    return output.getvalue()

st.divider()

# --- 2. PREPARAÇÃO ---
st.subheader("2. Preparação de Dados")
st.download_button(label="📥 Baixar Planilha para o Fiscal", data=gerar_planilha_fiscal(), file_name="Coleta_Valor_Global.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()

# --- 3. PROCESSAMENTO ---
tabs = st.tabs(["📊 Apuração Financeira", "📦 Controle de Estoque", "⚖️ Comparativo"])

with tabs[1]: # Aba Estoque
    uploaded_file = st.file_uploader("Suba a planilha PREENCHIDA", type="xlsx")
    if uploaded_file:
        try:
            df_itens = pd.read_excel(uploaded_file, sheet_name="ITENS_CICLOS", skiprows=2).dropna(subset=["Item/ Bloco"])
            df_retro = pd.read_excel(uploaded_file, sheet_name="RETROATIVO", skiprows=2).dropna(subset=["Competência"])
            
            # Cálculos Logísticos
            v_total_original = (df_itens["Quantidade"] * df_itens["VU C0 (R$)"]).sum()
            total_faturado = df_retro.iloc[:, 1].sum()
            col_rem = df_itens.columns[-1]
            v_remanescente_reaj = (df_itens[col_rem] * df_itens["VU C0 (R$)"] * fator_vigente).sum()
            v_global_estimado = total_faturado + v_remanescente_reaj

            st.success("Planilha processada!")
            st.metric("Valor Contratual Original (Teto C0)", f"R$ {v_total_original:,.2f}")
            
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Faturado (Realizado)", f"R$ {total_faturado:,.2f}")
            c2.metric("Saldo Remanescente (Reajustado)", f"R$ {v_remanescente_reaj:,.2f}")
            c3.metric("VALOR GLOBAL ESTIMADO", f"R$ {v_global_estimado:,.2f}")

            st.session_state['res'] = {'orig': v_total_original, 'final': v_global_estimado, 'fat': total_faturado}
        except Exception as e: st.error(f"Erro: {e}")

with tabs[2]: # Aba Comparativo
    if 'res' in st.session_state:
        st.header("Balanço do Contrato")
        r = st.session_state['res']
        # Definição: A variação é o quanto o Valor Global Reajustado supera o Valor Original
        variacao = r['final'] - r['orig']
        percentual = (variacao / r['orig'] * 100) if r['orig'] > 0 else 0
        
        st.metric("Variação Estimada do Valor Global", f"R$ {variacao:,.2f}", delta=f"{percentual:.2f}%")
        
        st.info(f"""
        **Explicação Técnica:**
        * **Valor Original:** R$ {r['orig']:,.2f} (Soma dos itens no preço de licitação).
        * **Execução Atual:** O contrato já faturou R$ {r['fat']:,.2f}.
        * **Projeção:** Considerando o reajuste de {((fator_vigente-1)*100):.2f}%, o saldo que resta custará R$ {r['final']-r['fat']:,.2f}.
        """)