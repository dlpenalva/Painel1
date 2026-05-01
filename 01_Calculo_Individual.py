import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="GCC - Cálculo Individual", layout="wide")

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
        v_base = float(df[df['MES_ANO'] == ref_base]['INDICE_NIVEL'].values[0])
        v_aniv = float(df[df['MES_ANO'] == ref_aniv]['INDICE_NIVEL'].values[0])
        var = (v_aniv / v_base) - 1
        return var, ref_base, ref_aniv, None
    except:
        return None, None, None, "Erro IST"

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("📊 Cálculo Único")

if 'farc' not in st.session_state: st.session_state.farc = {}

tab_adm, tab_calc, tab_rel = st.tabs(["📊 Admissibilidade", "🧮 Cálculo", "📄 Relatório"])

with tab_adm:
    col1, col2 = st.columns(2)
    with col1:
        dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 5, 1), format="DD/MM/YYYY")
        dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
    with col2:
        valor_base = st.number_input("Valor Atual (R$):", min_value=0.0, step=100.0)
        tipo_idx = st.selectbox("Índice:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Planilha CSV)"])

    dt_aniv = dt_base + relativedelta(years=1)
    dias_janela = (dt_solic - dt_aniv).days
    status = "Precluso" if dias_janela > 90 else "Admissível"
    
    st.metric("Status do Pedido", status)
    st.session_state.farc = {'dt_base': dt_base, 'dt_aniv': dt_aniv, 'dt_pedido': dt_solic, 'valor': valor_base, 'idx': tipo_idx, 'status': status}

with tab_calc:
    f = st.session_state.farc
    if "IST" in f['idx']:
        var, rb, ra, erro = calc_ist_csv(f['dt_base'], f['dt_aniv'])
        if not erro:
            st.metric(f"Variação IST ({rb} a {ra})", f"{var*100:,.2f}%".replace('.', ','))
            st.session_state.farc.update({'var': var})
    else:
        cod = "433" if "IPCA" in f['idx'] else "189"
        df, erro = get_index_data(cod, f['dt_base'].strftime('%d/%m/%Y'), f['dt_aniv'].strftime('%d/%m/%Y'))
        if df is not None:
            var = (1 + df['valor']).prod() - 1
            st.metric(f"Variação {f['idx']}", f"{var*100:,.2f}%".replace('.', ','))
            st.session_state.farc.update({'var': var})

with tab_rel:
    f = st.session_state.farc
    if 'var' in f:
        texto = f"Relatório: O pedido de reajuste é {f['status']}. Variação apurada de {f['var']*100:,.2f}%."
        st.text_area("Minuta", texto, height=200)