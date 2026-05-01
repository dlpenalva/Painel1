import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Configuração de Interface
st.set_page_config(page_title="GCC - Telebras", layout="wide")

# Estilo Institucional Telebras
st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #0054a6; }
    
    /* Estilização das Abas */
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #e1e4e8;
        border-radius: 4px 4px 0px 0px;
        padding: 10px 20px;
        color: #555;
    }
    .stTabs [aria-selected="true"] {
        background-color: #0054a6 !important;
        color: white !important;
    }
    
    h1 { color: #0054a6; }
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

# Logo (Link alternativo estável)
st.image("https://logodownload.org/wp-content/uploads/2020/02/telebras-logo.png", width=200)
st.title("Admissibilidade e Variação de Reajuste")
st.caption("Ferramenta de apoio à Gestão de Contratos - Telebras (GCC)")

tab_adm, tab_calc, tab_rel = st.tabs(["Análise de Admissibilidade", "Cálculo de Reajuste", "Relatório"])

with tab_adm:
    st.subheader("Critérios da Cláusula Oitava")
    col1, col2 = st.columns(2)
    with col1:
        dt_base = st.date_input("Data-Base (Proposta; Último Reajuste; ou Aniversário Anterior):", value=datetime(2024, 5, 1), format="DD/MM/YYYY")
        dt_solic = st.date_input("Data do Pedido do Reajuste:", format="DD/MM/YYYY")
    with col2:
        valor_contrato = st.number_input("Valor a Reajustar (R$):", min_value=0.0, step=100.0)
        tipo_idx = st.selectbox("Índice de Reajuste:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Manual)"])

    # Lógica
    data_aniversario = dt_base + relativedelta(years=1)
    data_fim_calc = data_aniversario - relativedelta(days=1)
    dias_janela = (dt_solic - data_aniversario).days
    intersticio_ok = dt_solic >= data_aniversario
    precluso = dias_janela > 90

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("Dias da Janela", f"{max(0, dias_janela)} dias")
    
    if precluso: c2.error("Status: Precluso")
    elif not intersticio_ok: c2.warning("Status: Antecipado")
    else: c2.success("Status: Admissível")
    
    if intersticio_ok: c3.success("Interstício: Ok")
    else: c3.warning("Interstício: Pendente")

    # Guardar dados
    st.session_state.farc_data = {
        'dt_base': dt_base.strftime('%d/%m/%Y'),
        'dt_pedido': dt_solic.strftime('%d/%m/%Y'),
        'dt_aniv': data_aniversario.strftime('%d/%m/%Y'),
        'inicio_api': dt_base.strftime('%d/%m/%Y'),
        'fim_api': data_fim_calc.strftime('%d/%m/%Y'),
        'valor': valor_contrato,
        'indice': tipo_idx,
        'cod_bcb': tipo_idx.split('(')[1].replace('Série ', '').replace(')', '') if '(' in tipo_idx else None,
        'admissivel': intersticio_ok and not precluso,
        'dias': dias_janela
    }

with tab_calc:
    d = st.session_state.farc_data
    if not d or d.get('valor') == 0:
        st.info("Preencha os dados na aba de Admissibilidade.")
    else:
        st.subheader("Cálculo da Variação (12 meses)")
        if d['cod_bcb']:
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
    if not d.get('admissivel') and d.get('valor') != 0:
        st.error("O pedido não atende aos critérios de admissibilidade (Precluso ou Antecipado).")
    elif d.get('valor') > 0:
        st.subheader("Minuta para Relatório Técnico")
        
        texto_relatorio = f"""
        ANÁLISE DE REAJUSTE CONTRATUAL
        
        1. ADMISSIBILIDADE:
        - Data-Base Anterior: {d['dt_base']}
        - Data do Aniversário do Direito: {d['dt_aniv']}
        - Data do Pedido: {d['dt_pedido']}
        - Prazo Decorrido: {d['dias']} dias.
        - Status: {"ADMISSÍVEL" if d['admissivel'] else "NÃO ADMISSÍVEL"}
        
        2. VARIAÇÃO ECONÔMICA:
        - Índice Aplicado: {d['indice']}
        - Período de Apuração: {d['inicio_api']} a {d['fim_api']} (12 meses)
        - Variação Acumulada: {d.get('var_acum', 0):.6%}
        
        3. CONCLUSÃO:
        - Valor Original: R$ {d['valor']:,.2f}
        - Valor Reajustado: R$ {d.get('v_novo', 0):,.2f}
        """
        st.text_area("Copie o texto abaixo:", value=texto_relatorio, height=300)
        st.button("📋 Copiar para Área de Transferência (Simulado)")
    else:
        st.info("Realize o cálculo primeiro.")