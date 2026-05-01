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
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #003366; }
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [aria-selected="true"] { background-color: #003366 !important; color: white !important; }
    h1 { color: #003366; }
    </style>
    """, unsafe_allow_html=True)

def get_index_data(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=15)
        if response.status_code != 200:
            return None, "Erro na API do Banco Central."
        dados = response.json()
        if not dados:
            return None, "Nenhum dado encontrado para este período."
        df = pd.DataFrame(dados)
        df['valor'] = pd.to_numeric(df['valor']) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        return df, None
    except Exception as e:
        return None, f"Erro de conexão: {str(e)}"

if 'farc_data' not in st.session_state:
    st.session_state.farc_data = {}

# Logo Telebras Azul Profundo
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
    
    # Regra de Intervalo: IST (13 meses) vs Outros (12 meses)
    if "IST" in tipo_idx:
        data_fim_calc = data_aniversario # Inclui o mês do aniversário
        label_int = "13 meses (Regra IST)"
    else:
        data_fim_calc = data_aniversario - relativedelta(days=1) # Até o mês anterior
        label_int = "12 meses"

    dias_janela = (dt_solic - data_aniversario).days
    intersticio_ok = dt_solic >= data_aniversario
    precluso = dias_janela > 90

    st.divider()
    c1, c2, c3 = st.columns(3)
    
    # Métricas limpas para evitar erro DeltaGenerator
    c1.metric("Dias da Janela", f"{max(0, dias_janela)} dias")
    
    if precluso:
        c2.error("Status: Precluso")
    elif not intersticio_ok:
        c2.warning("Status: Antecipado")
    else:
        c2.success("Status: Admissível")

    if intersticio_ok:
        c3.success("Interstício: Ok")
    else:
        c3.warning("Interstício: Pendente")

    # Extrair código da série de forma segura
    try:
        cod_serie = tipo_idx.split('Série ')[1].replace(')', '')
    except:
        cod_serie = "433" # Default IPCA

    st.session_state.farc_data = {
        'dt_base': dt_base.strftime('%d/%m/%Y'),
        'dt_pedido': dt_solic.strftime('%d/%m/%Y'),
        'dt_aniv': data_aniversario.strftime('%d/%m/%Y'),
        'inicio_api': dt_base.strftime('%d/%m/%Y'),
        'fim_api': data_fim_calc.strftime('%d/%m/%Y'),
        'valor': valor_contrato,
        'indice_nome': tipo_idx,
        'cod_bcb': cod_serie,
        'precluso': precluso,
        'admissivel': intersticio_ok and not precluso,
        'dias': dias_janela,
        'intervalo': label_int
    }

with tab_calc:
    d = st.session_state.farc_data
    if d.get('valor', 0) > 0:
        st.subheader(f"Cálculo da Variação ({d['intervalo']})")
        with st.spinner("Buscando índices oficiais..."):
            df, erro = get_index_data(d['cod_bcb'], d['inicio_api'], d['fim_api'])
            
            if df is not None:
                var_acum = (1 + df['valor']).prod() - 1
                v_novo = d['valor'] * (1 + var_acum)
                
                st.session_state.farc_data['var_acum'] = var_acum
                st.session_state.farc_data['v_novo'] = v_novo
                
                m1, m2 = st.columns(2)
                m1.metric("Variação Acumulada", f"{var_acum:.6%}")
                m2.metric("Valor Reajustado", f"R$ {v_novo:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                st.dataframe(df.assign(data=df['data'].dt.strftime('%m/%Y')), use_container_width=True)
            else:
                st.error(f"Erro: {erro}")
    else:
        st.info("Insira o valor do contrato na primeira aba.")

with tab_rel:
    d = st.session_state.farc_data
    if d.get('valor', 0) > 0:
        st.subheader("Minuta de Relatório")
        
        efeitos = d['dt_aniv']
        txt_preclusao = ""
        if d['precluso']:
            # Penalidade de 2 meses para efeitos financeiros em caso de preclusão
            base_date = datetime.strptime(d['dt_aniv'], '%d/%m/%Y')
            efeitos = (base_date + relativedelta(months=2)).strftime('%d/%m/%Y')
            txt_preclusao = f"\n- NOTA: Pedido precluso. Efeitos financeiros postergados para {efeitos}."

        minuta = f"""RELATÓRIO GCC/TELEBRAS
        
1. ADMISSIBILIDADE
- Data-Base Anterior: {d['dt_base']}
- Aniversário do Direito: {d['dt_aniv']}
- Data do Protocolo: {d['dt_pedido']}
- Status: {"TEMPESTIVO" if not d['precluso'] else "PRECLUSO"}{txt_preclusao}

2. MEMÓRIA DE CÁLCULO
- Índice Utilizado: {d['indice_nome']}
- Período: {d['inicio_api']} a {d['fim_api']} ({d['intervalo']})
- Variação Acumulada: {d.get('var_acum', 0):.6%}

3. CONCLUSÃO
- Valor Original: R$ {d['valor']:,.2f}
- Valor Atualizado: R$ {d.get('v_novo', 0):,.2f}
- Data de Início dos Efeitos: {efeitos}
"""
        st.text_area("Texto para o SEI:", value=minuta, height=400)
    else:
        st.info("Realize a análise para gerar o relatório.")