import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta
from fpdf import FPDF  # Importação permanece igual, mas usará a fpdf2 instalada

st.set_page_config(page_title="GCC - Telebras", layout="wide")

# Estilo Telebras
st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #003366; }
    .stTabs [aria-selected="true"] { background-color: #003366 !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

def get_index_data(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=15)
        df = pd.DataFrame(response.json())
        if df.empty: return None, "Vazio"
        df['valor'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        if len(df) > 12: df = df.head(12)
        return df, None
    except:
        return None, "Erro API"

def calc_ist_csv(dt_base, dt_aniv):
    try:
        df = pd.read_csv('ist.csv', sep=None, engine='python', decimal=',')
        df.columns = df.columns.str.replace('^\ufeff', '', regex=True)
        meses_map = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
        ref_base = f"{meses_map[dt_base.month]}/{str(dt_base.year)[2:]}"
        ref_aniv = f"{meses_map[dt_aniv.month]}/{str(dt_aniv.year)[2:]}"
        
        row_base = df[df['MES_ANO'] == ref_base].iloc[0]
        row_aniv = df[df['MES_ANO'] == ref_aniv].iloc[0]
        
        v_base = float(row_base['INDICE_NIVEL'])
        v_aniv = float(row_aniv['INDICE_NIVEL'])
        var = (v_aniv / v_base) - 1
        
        memoria = pd.DataFrame([
            {"Referência": f"Inicial ({ref_base})", "Índice Nível": v_base},
            {"Referência": f"Final ({ref_aniv})", "Índice Nível": v_aniv}
        ])
        return var, ref_base, ref_aniv, memoria, None
    except:
        return None, None, None, None, "Erro nas referências do IST no CSV"

def gerar_pdf(texto):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    # Substituindo caracteres que podem dar erro no PDF padrão
    texto_limpo = texto.replace('º', '.').replace('R$', 'RS')
    pdf.multi_cell(0, 10, texto_limpo)
    return pdf.output()

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Gestão de Cálculos Contratuais")

if 'farc' not in st.session_state: st.session_state.farc = {}

tab_adm, tab_calc, tab_rel = st.tabs(["Admissibilidade", "Cálculo Detalhado", "Relatório Final"])

with tab_adm:
    col1, col2 = st.columns(2)
    with col1:
        dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 5, 1), format="DD/MM/YYYY")
        dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
    with col2:
        valor_base = st.number_input("Valor Atual (R$):", min_value=0.0, step=100.0)
        tipo_idx = st.selectbox("Índice de Reajuste:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Planilha CSV)"])

    dt_aniv = dt_base + relativedelta(years=1)
    dt_fim_calc = dt_aniv - relativedelta(months=1)
    dias_janela = (dt_solic - dt_aniv).days
    intersticio_ok = dt_solic >= dt_aniv

    if dias_janela > 90: status = "Precluso"
    elif not intersticio_ok and (dt_solic.month == dt_aniv.month): status = "Admissível (Ajuste Prévio)"
    elif not intersticio_ok: status = "Antecipado"
    else: status = "Admissível"

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("Janela Temporal", f"{max(0, dias_janela)} dias")
    c2.info(f"Status: {status}")
    c3.success("Interstício Legal: Ok") if intersticio_ok else c3.warning("Interstício: Pendente")

    st.session_state.farc = {
        'dt_base': dt_base, 'dt_aniv': dt_aniv, 'dt_pedido': dt_solic,
        'dt_fim_calc': dt_fim_calc, 'valor': valor_base, 'idx': tipo_idx, 'status': status
    }

with tab_calc:
    f = st.session_state.farc
    if f.get('valor', 0) > 0:
        if "IST" in f['idx']:
            var, rb, ra, mem, erro = calc_ist_csv(f['dt_base'], f['dt_aniv'])
            if not erro:
                v_novo = f['valor'] * (1 + var)
                st.subheader(f"Memória de Cálculo - IST")
                st.write(f"Fórmula: Var = (Nível Final / Nível Inicial) - 1")
                col_a, col_b = st.columns(2)
                col_a.metric(f"Variação ({rb} a {ra})", f"{var:.6%}")
                col_b.metric("Novo Valor Contratual", f"R$ {v_novo:,.2f}")
                st.table(mem)
                st.session_state.farc.update({'var': var, 'v_novo': v_novo})
        else:
            cod = "433" if "IPCA" in f['idx'] else "189"
            df, erro = get_index_data(cod, f['dt_base'].strftime('%d/%m/%Y'), f['dt_fim_calc'].replace(day=28).strftime('%d/%m/%Y'))
            if df is not None:
                var = (1 + df['valor']).prod() - 1
                v_novo = f['valor'] * (1 + var)
                st.subheader(f"Memória de Cálculo - {f['idx']}")
                st.metric(f"Período: {df.iloc[0]['data'].strftime('%m/%Y')} a {df.iloc[-1]['data'].strftime('%m/%Y')}", f"{var:.6%}")
                st.metric("Novo Valor Contratual", f"R$ {v_novo:,.2f}")
                st.dataframe(df.assign(data=df['data'].dt.strftime('%m/%Y')), use_container_width=True)
                st.session_state.farc.update({'var': var, 'v_novo': v_novo})

with tab_rel:
    f = st.session_state.farc
    if f.get('v_novo'):
        texto_rel = f"""RELATÓRIO TÉCNICO DE REAJUSTE CONTRATUAL

1. FUNDAMENTAÇÃO LEGAL E REFERÊNCIAS
- Amparo: Lei 13.303/2016 e Decreto n. 12.500/2025.
- Empresa: Telebras (Status: Não Dependente).
- Data-Base Anterior: {f['dt_base'].strftime('%d/%m/%Y')}
- Aniversário do Direito: {f['dt_aniv'].strftime('%d/%m/%Y')}

2. ANÁLISE DE ADMISSIBILIDADE
- Data do Protocolo: {f['dt_pedido'].strftime('%d/%m/%Y')}
- Parecer: O pedido é considerado {f['status']}.

3. MEMÓRIA DE CÁLCULO
- Índice Aplicado: {f['idx']}
- Variação Acumulada (12 meses): {f['var']:.6%}
- Valor Atual: R$ {f['valor']:,.2f}
- Valor Reajustado: R$ {f['v_novo']:,.2f}

4. CONCLUSÃO
Considerando o cumprimento do interstício de 12 meses, o novo valor de R$ {f['v_novo']:,.2f} é considerado apto para processamento, retroagindo seus efeitos financeiros a {f['dt_aniv'].strftime('%d/%m/%Y')}."""
        
        st.text_area("Conteúdo do Relatório:", texto_rel, height=380)
        
        pdf_data = gerar_pdf(texto_rel)
        st.download_button(
            label="📥 Baixar Relatório em PDF",
            data=pdf_data,
            file_name=f"Relatorio_Reajuste_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )