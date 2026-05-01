import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="GCC - Telebras", layout="wide")

# Estilo Telebras
st.markdown("<style>.stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #003366; } .stTabs [aria-selected='true'] { background-color: #003366 !important; color: white !important; }</style>", unsafe_allow_html=True)
st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Gestão de Reajuste Contratual")

# --- MOTORES DE CAPTURA ---

def get_bcb_data(serie, inicio, fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie}/dados?formato=json&dataInicial={inicio}&dataFinal={fim}"
    try:
        res = requests.get(url, timeout=10)
        df = pd.DataFrame(res.json())
        df['valor'] = pd.to_numeric(df['valor']) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        return df, None
    except:
        return None, "Erro na API do Banco Central."

def get_ist_anatel():
    url = "https://www.gov.br/anatel/pt-br/regulado/competicao/tarifas-e-precos/valores-do-ist"
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(url, headers=headers, timeout=15)
        tabelas = pd.read_html(response.text, decimal=',', thousands='.')
        df_ist = pd.concat(tabelas, ignore_index=True)
        # Ajuste de colunas baseado na estrutura da Anatel
        df_ist = df_ist.iloc[:, :2]
        df_ist.columns = ['Ref', 'Var']
        meses_map = {'Jan':1,'Fev':2,'Mar':3,'Abr':4,'Mai':5,'Jun':6,'Jul':7,'Ago':8,'Set':9,'Out':10,'Nov':11,'Dez':12}
        df_ist['mes'] = df_ist['Ref'].str.split('/').str[0].map(meses_map)
        df_ist['ano'] = pd.to_numeric("20" + df_ist['Ref'].str.split('/').str[1])
        df_ist['data'] = pd.to_datetime(df_ist[['ano', 'mes']].assign(day=1))
        df_ist['valor'] = df_ist['Var'].astype(str).str.replace('%', '').str.replace(',', '.').astype(float) / 100
        return df_ist[['data', 'valor']].dropna(), None
    except:
        return None, "Erro ao acessar site da Anatel. Verifique o link oficial."

# --- INTERFACE ---

tab_adm, tab_calc, tab_rel = st.tabs(["Análise de Admissibilidade", "Cálculo de Reajuste", "Relatório"])

with tab_adm:
    col1, col2 = st.columns(2)
    with col1:
        dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 5, 3), format="DD/MM/YYYY")
        dt_pedido = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
    with col2:
        valor_base = st.number_input("Valor a Reajustar (R$):", min_value=0.0, step=100.0)
        tipo_idx = st.selectbox("Índice:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Anatel)"])

    dt_aniv = dt_base + relativedelta(years=1)
    dias_janela = (dt_pedido - dt_aniv).days
    mesmo_mes = (dt_pedido.month == dt_aniv.month and dt_pedido.year == dt_aniv.year)
    intersticio_ok = dt_pedido >= dt_aniv
    precluso = dias_janela > 90

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("Dias da Janela", f"{max(0, dias_janela)} dias")

    if precluso:
        status_txt, status_cor = "Precluso", "error"
    elif not intersticio_ok and mesmo_mes:
        status_txt, status_cor = "Admissível (Adiantado)", "warning"
    elif not intersticio_ok:
        status_txt, status_cor = "Antecipado", "warning"
    else:
        status_txt, status_cor = "Admissível", "success"

    if status_cor == "error": c2.error(f"Status: {status_txt}")
    elif status_cor == "warning": c2.warning(f"Status: {status_txt}")
    else: c2.success(f"Status: {status_txt}")

    c3.success("Interstício: Ok") if intersticio_ok else c3.warning("Interstício: Pendente")

    # Guardar estado
    st.session_state.farc = {
        'dt_base': dt_base, 'dt_aniv': dt_aniv, 'valor': valor_base, 'idx': tipo_idx,
        'precluso': precluso, 'status': status_txt, 'intersticio': intersticio_ok
    }

with tab_calc:
    d = st.session_state.farc
    if d.get('valor', 0) > 0:
        # Lógica de Janela (Senior Fix)
        if "IST" in d['idx']:
            inicio_calc = d['dt_base'] # Inclui o mês base (13 meses)
            fim_calc = d['dt_aniv']
            label_int = "13 meses (Regra IST)"
        else:
            inicio_calc = d['dt_base'] + relativedelta(months=1) # Pula o mês base
            fim_calc = d['dt_aniv'] # Vai até o aniversário
            label_int = "12 meses"

        st.subheader(f"Cálculo da Variação — {label_int}")
        st.info(f"**Status Atual:** {d['status']} | **Interstício:** {'Ok' if d['intersticio'] else 'Pendente'}")

        with st.spinner("Obtendo dados oficiais..."):
            if "IST" in d['idx']:
                df, erro = get_ist_anatel()
            else:
                serie = "433" if "IPCA" in d['idx'] else "189"
                df, erro = get_bcb_data(serie, inicio_calc.strftime('%d/%m/%Y'), fim_calc.strftime('%d/%m/%Y'))

        if df is not None:
            # Filtro de segurança para garantir o número exato de meses
            df = df[(df['data'] >= pd.to_datetime(inicio_calc).replace(day=1)) & 
                    (df['data'] <= pd.to_datetime(fim_calc).replace(day=1))].sort_values('data')
            
            if not df.empty:
                var_acum = (1 + df['valor']).prod() - 1
                v_novo = d['valor'] * (1 + var_acum)
                
                m1, m2 = st.columns(2)
                m1.metric("Variação Acumulada", f"{var_acum:.6%}")
                m2.metric("Valor Reajustado", f"R$ {v_novo:,.2f}")
                
                st.write(f"**Período analisado:** {df['data'].min().strftime('%m/%Y')} a {df['data'].max().strftime('%m/%Y')} ({len(df)} índices)")
                st.dataframe(df.assign(data=df['data'].dt.strftime('%m/%Y')), use_container_width=True)
                st.session_state.farc.update({'var': var_acum, 'v_novo': v_novo, 'len': len(df)})
            else:
                st.error("Dados não disponíveis para este intervalo no momento.")
        else:
            st.error(erro)