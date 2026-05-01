import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

# 1. Configuração e Estilo
st.set_page_config(page_title="GCC - Telebras", layout="wide")

st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #003366; }
    .stTabs [aria-selected="true"] { background-color: #003366 !important; color: white !important; }
    </style>
    """, unsafe_allow_html=True)

# 2. Motores de Cálculo
def get_index_data(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=15)
        df = pd.DataFrame(response.json())
        if df.empty: return None, "Vazio"
        df['valor'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        return df, None
    except:
        return None, "Erro na API do Banco Central"

def calc_ist_csv(dt_base, dt_aniv):
    try:
        # Resolve o erro de 'MES_ANO' limpando caracteres invisíveis (UTF-8 SIG)
        df = pd.read_csv('ist.csv', sep=None, engine='python', decimal=',')
        df.columns = df.columns.str.replace('^\ufeff', '', regex=True) # Remove o erro do seu print
        
        meses_map = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 
                     7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
        
        ref_base = f"{meses_map[dt_base.month]}/{str(dt_base.year)[2:]}"
        ref_aniv = f"{meses_map[dt_aniv.month]}/{str(dt_aniv.year)[2:]}"
        
        val_base = float(df[df['MES_ANO'] == ref_base]['INDICE_NIVEL'].values[0])
        val_aniv = float(df[df['MES_ANO'] == ref_aniv]['INDICE_NIVEL'].values[0])
        
        var = (val_aniv / val_base) - 1
        return var, ref_base, ref_aniv, None
    except Exception as e:
        return None, None, None, f"Erro: Referência {ref_base} ou {ref_aniv} não encontrada no CSV."

# 3. Interface Principal
st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Admissibilidade e Variação de Reajuste")

if 'farc' not in st.session_state: st.session_state.farc = {}

tab_adm, tab_calc, tab_rel = st.tabs(["Análise de Admissibilidade", "Cálculo de Reajuste", "Relatório"])

with tab_adm:
    col1, col2 = st.columns(2)
    with col1:
        dt_base = st.date_input("Data-Base:", value=datetime(2023, 5, 1), format="DD/MM/YYYY")
        dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
    with col2:
        valor_base = st.number_input("Valor a Reajustar (R$):", min_value=0.0, step=100.0)
        tipo_idx = st.selectbox("Índice:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Planilha CSV)"])

    # Lógica de Prazos e Status
    dt_aniv = dt_base + relativedelta(years=1)
    dt_fim_api = dt_aniv - relativedelta(days=1) # Lógica Perfeita 12 meses
    dias_atraso = (dt_solic - dt_aniv).days
    mesmo_mes = (dt_solic.month == dt_aniv.month and dt_solic.year == dt_aniv.year)
    intersticio_ok = dt_solic >= dt_aniv

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("Dias da Janela", f"{max(0, dias_atraso)} dias")

    # Definição de Status
    if dias_atraso > 90:
        status_txt, status_tipo = "Precluso", "error"
    elif not intersticio_ok and mesmo_mes:
        status_txt, status_tipo = "Admissível (Ajuste Prévio)", "warning"
    elif not intersticio_ok:
        status_txt, status_tipo = "Antecipado", "warning"
    else:
        status_txt, status_tipo = "Admissível", "success"

    if status_tipo == "error": st.error(f"Status: {status_txt}")
    elif status_tipo == "warning": st.warning(f"Status: {status_txt}")
    else: st.success(f"Status: {status_txt}")

    st.session_state.farc = {
        'dt_base': dt_base, 'dt_aniv': dt_aniv, 'dt_fim_api': dt_fim_api,
        'valor': valor_base, 'idx': tipo_idx, 'status': status_txt, 'intersticio': intersticio_ok
    }

with tab_calc:
    f = st.session_state.farc
    if f.get('valor', 0) > 0:
        st.subheader(f"Cálculo: {f['idx']}")
        st.info(f"**Status:** {f['status']} | **Período:** {f['dt_base'].strftime('%d/%m/%Y')} a {f['dt_fim_api'].strftime('%d/%m/%Y')}")

        if "IST" in f['idx']:
            var, r_base, r_aniv, erro = calc_ist_csv(f['dt_base'], f['dt_aniv'])
            if erro: st.error(erro)
            else:
                v_novo = f['valor'] * (1 + var)
                m1, m2 = st.columns(2)
                m1.metric(f"Variação ({r_base} a {r_aniv})", f"{var:.6%}")
                m2.metric("Novo Valor", f"R$ {v_novo:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        else:
            cod = "433" if "IPCA" in f['idx'] else "189"
            df, erro = get_index_data(cod, f['dt_base'].strftime('%d/%m/%Y'), f['dt_fim_api'].strftime('%d/%m/%Y'))
            if df is not None:
                var = (1 + df['valor']).prod() - 1
                v_novo = f['valor'] * (1 + var)
                m1, m2 = st.columns(2)
                m1.metric("Variação Acumulada", f"{var:.6%}")
                m2.metric("Novo Valor", f"R$ {v_novo:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                st.dataframe(df.assign(data=df['data'].dt.strftime('%m/%Y')), use_container_width=True)