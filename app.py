import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="GCC - Telebras", layout="wide")

# Estilo Institucional
st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #003366; }
    .stTabs [aria-selected="true"] { background-color: #003366 !important; color: white !important; }
    h1 { color: #003366; }
    </style>
    """, unsafe_allow_html=True)

def get_index_data(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        df = pd.DataFrame(response.json())
        if df.empty: return None, "Vazio"
        df['valor'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        return df, None
    except:
        return None, "Erro API"

if 'farc_data' not in st.session_state:
    st.session_state.farc_data = {}

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Admissibilidade e Variação de Reajuste")

tab_adm, tab_calc, tab_rel = st.tabs(["Análise de Admissibilidade", "Cálculo de Reajuste", "Relatório"])

with tab_adm:
    st.subheader("Critérios da Cláusula Oitava")
    col1, col2 = st.columns(2)
    with col1:
        dt_base = st.date_input("Data-Base (Proposta; Último Reajuste; ou Aniversário Anterior):", value=datetime(2024, 5, 1), format="DD/MM/YYYY")
        dt_solic = st.date_input("Data do Pedido do Reajuste:", format="DD/MM/YYYY")
    with col2:
        valor_contrato = st.number_input("Valor a Reajustar (R$):", min_value=0.0, step=100.0)
        tipo_idx = st.selectbox("Índice de Reajuste:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Série 25344)"])

    # Lógica de Prazos
    data_aniversario = dt_base + relativedelta(years=1)
    
    # --- DIFERENCIAÇÃO DE REGRA DE INTERVALO ---
    if "IST" in tipo_idx:
        # IST: 13 meses (Ex: 04/23 a 04/24)
        data_fim_calc = data_aniversario
        label_intervalo = "13 meses (Regra IST)"
    else:
        # Demais: 12 meses (Ex: 05/24 a 04/25)
        data_fim_calc = data_aniversario - relativedelta(days=1)
        label_intervalo = "12 meses"

    dias_janela = (dt_solic - data_aniversario).days
    intersticio_ok = dt_solic >= data_aniversario
    precluso = dias_janela > 90

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("Dias da Janela", f"{max(0, dias_janela)} dias")
    c2.error("Status: Precluso") if precluso else (c2.warning("Status: Antecipado") if not intersticio_ok else c2.success("Status: Admissível"))
    c3.success("Interstício: Ok") if intersticio_ok else c3.warning("Interstício: Pendente")

    st.session_state.farc_data = {
        'dt_base': dt_base.strftime('%d/%m/%Y'),
        'dt_pedido': dt_solic.strftime('%d/%m/%Y'),
        'dt_aniv': data_aniversario.strftime('%d/%m/%Y'),
        'inicio_api': dt_base.strftime('%d/%m/%Y'),
        'fim_api': data_fim_calc.strftime('%d/%m/%Y'),
        'valor': valor_contrato,
        'indice': tipo_idx,
        'cod_bcb': tipo_idx.split('(')[1].replace('Série ', '').replace(')', ''),
        'admissivel': intersticio_ok and not precluso,
        'precluso': precluso,
        'dias': dias_janela,
        'label_intervalo': label_intervalo
    }

with tab_calc:
    d = st.session_state.farc_data
    if d.get('valor', 0) > 0:
        st.subheader(f"Cálculo da Variação ({d['label_intervalo']})")
        df, erro = get_index_data(d['cod_bcb'], d['inicio_api'], d['fim_api'])
        if df is not None:
            var_acum = (1 + df['valor']).prod() - 1
            v_novo = d['valor'] * (1 + var_acum)
            st.session_state.farc_data['var_acum'] = var_acum
            st.session_state.farc_data['v_novo'] = v_novo
            
            m1, m2 = st.columns(2)
            m1.metric("Variação Acumulada", f"{var_acum:.6%}")
            m2.metric("Valor Atualizado", f"R$ {v_novo:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            st.dataframe(df.assign(data=df['data'].dt.strftime('%m/%Y')), use_container_width=True)

with tab_rel:
    d = st.session_state.farc_data
    if d.get('valor', 0) > 0:
        st.subheader("Minuta de Relatório Técnico")
        data_efeitos = d['dt_aniv']
        if d['precluso']:
            data_aniv_obj = datetime.strptime(d['dt_aniv'], '%d/%m/%Y')
            data_efeitos = (data_aniv_obj + relativedelta(months=2)).strftime('%d/%m/%Y')
        
        texto = f"""RELATÓRIO GCC/TELEBRAS
        
1. ADMISSIBILIDADE
- Data-Base: {d['dt_base']} | Pedido: {d['dt_pedido']}
- Status: {"TEMPESTIVO" if not d['precluso'] else "PRECLUSO"}

2. MEMÓRIA DE CÁLCULO
- Índice: {d['indice']}
- Período: {d['inicio_api']} a {d['fim_api']} ({d['label_intervalo']})
- Variação: {d.get('var_acum', 0):.6%}

3. CONCLUSÃO
- Valor Reajustado: R$ {d.get('v_novo', 0):,.2f}
- Efeitos Financeiros: {data_efeitos}
"""
        st.text_area("Texto SEI:", value=texto, height=350)