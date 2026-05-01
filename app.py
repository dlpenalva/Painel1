import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Configuração de Interface
st.set_page_config(page_title="GCC - Telebras", layout="wide")

# Estilo Institucional Telebras (Azul Profundo #003366)
st.markdown("""
    <style>
    .main { background-color: #f0f2f6; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #003366; }
    
    /* Estilização das Abas */
    .stTabs [data-baseweb="tab-list"] { gap: 10px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #e1e4e8;
        border-radius: 4px 4px 0px 0px;
        padding: 10px 20px;
        color: #555;
    }
    .stTabs [aria-selected="true"] {
        background-color: #003366 !important;
        color: white !important;
    }
    
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

# Logo Atualizada (Azul Profundo)
st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
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

    # Lógica de Prazos
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
        'precluso': precluso,
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
    if d.get('valor', 0) > 0:
        st.subheader("Minuta de Relatório Técnico")
        
        # Definição dos Efeitos Financeiros
        data_efeitos = d['dt_aniv']
        obs_preclusao = ""
        
        if d['precluso']:
            # Se precluso, efeitos financeiros ocorrem 2 meses após o aniversário (exemplo de penalidade)
            data_aniv_obj = datetime.strptime(d['dt_aniv'], '%d/%m/%Y')
            data_efeitos = (data_aniv_obj + relativedelta(months=2)).strftime('%d/%m/%Y')
            obs_preclusao = f"""
            NOTA DE PRECLUSÃO: O pedido foi protocolado com {d['dias']} dias de atraso em relação ao marco inicial (limite de 90 dias). 
            Conforme entendimento administrativo e cláusulas de preclusão lógica, o reajuste é devido, porém os efeitos financeiros 
            são postergados para {data_efeitos}, em virtude da inércia da contratada."""
        
        texto_relatorio = f"""RELATÓRIO DE ANÁLISE ECONÔMICO-FINANCEIRA

1. OBJETO E FUNDAMENTAÇÃO LEGAL
Trata-se de análise de admissibilidade do pedido de reajuste de preços. A fundamentação baseia-se na Lei 13.303/2016 e no Regulamento Interno de Licitações e Contratos (RELIC) da Telebras, que estabelecem o direito à manutenção do equilíbrio econômico-fundo do contrato após o interstício mínimo de 12 meses.

2. ANÁLISE DE ADMISSIBILIDADE
- Data-Base Anterior/Proposta: {d['dt_base']}
- Data do Aniversário do Direito (12 meses): {d['dt_aniv']}
- Data do Protocolo do Pedido: {d['dt_pedido']}
- Prazo Decorrido após Aniversário: {d['dias']} dias.

STATUS: {"TEMPESTIVO" if not d['precluso'] else "PRECLUSO (FORA DO PRAZO DE 90 DIAS)"}
{obs_preclusao}

3. MEMÓRIA DE CÁLCULO
- Índice: {d['indice']}
- Período de Variação: {d['inicio_api']} a {d['fim_api']}
- Variação Acumulada Apurada: {d.get('var_acum', 0):.6%}

4. CONCLUSÃO E EFEITOS FINANCEIROS
Considerando a variação apurada, o valor do contrato passa de R$ {d['valor']:,.2f} para R$ {d.get('v_novo', 0):,.2f}.
Os efeitos financeiros retroagem a: {data_efeitos}.
        """
        
        st.text_area("Texto para cópia (SEI):", value=texto_relatorio, height=450)
        st.info("💡 Dica: Se o pedido for precluso, o relatório acima já inclui a justificativa da alteração dos efeitos financeiros.")
    else:
        st.info("Realize a análise de admissibilidade e o cálculo para gerar o relatório.")