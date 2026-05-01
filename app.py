import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Configuração de Interface
st.set_page_config(page_title="Admissibilidade e Variação de Reajuste", layout="wide")

st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_html=True)

def get_index_data(serie_codigo, data_inicio, data_fim):
    """Busca dados na API SGS/BCB com o intervalo exato de 12 meses."""
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        df = pd.DataFrame(response.json())
        if df.empty:
            return None, "Nenhum dado encontrado."
        df['valor'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        return df, None
    except Exception as e:
        return None, f"Erro de conexão: {str(e)}"

if 'farc_data' not in st.session_state:
    st.session_state.farc_data = {}

st.title("Admissibilidade e Variação de Reajuste")
st.caption("Ferramenta de apoio à Gestão de Contratos - Telebras (GCC)")

tab_adm, tab_calc, tab_rel = st.tabs(["Análise de Admissibilidade", "Cálculo de Reajuste", "Minuta de Relatório"])

with tab_adm:
    st.subheader("Critérios da Cláusula Oitava")
    col1, col2 = st.columns(2)
    
    with col1:
        dt_base = st.date_input("Data do Último Reajuste (ou Aniversário anterior):", value=datetime(2024, 5, 1), format="DD/MM/YYYY")
        dt_solic = st.date_input("Data do Protocolo da Solicitação:", format="DD/MM/YYYY")
    
    with col2:
        valor_contrato = st.number_input("Valor a Reajustar (R$):", min_value=0.0, step=100.0)
        tipo_idx = st.selectbox("Índice de Reajuste:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Inserção Manual)"])

    # --- LÓGICA DE DATAS ---
    # O direito nasce 1 ano após a base
    data_aniversario_direito = dt_base + relativedelta(years=1)
    
    # A variação deve ser SEMPRE de 12 meses (Ex: 05/2024 a 04/2025)
    # Para a API do BCB, pegamos do mês da base até o mês anterior ao novo aniversário
    data_fim_calculo = data_aniversario_direito - relativedelta(days=1)
    
    dias_decorridos = (dt_solic - data_aniversario_direito).days
    intersticio_ok = dt_solic >= data_aniversario_direito
    precluso = dias_decorridos > 90

    st.divider()
    c1, c2, c3 = st.columns(3)
    
    if dias_decorridos < 0:
        c1.metric("Status do Prazo", f"Faltam {abs(dias_decorridos)} dias")
    else:
        c1.metric("Dias da Janela de Reajuste", f"{dias_decorridos} dias")

    if precluso:
        c2.error("Status: Precluso")
    elif not intersticio_ok:
        c2.warning("Status: Antecipado")
    else:
        c2.success("Status: Admissível")

    c3.success("Interstício: Ok") if intersticio_ok else c3.warning("Interstício: Pendente")

    st.session_state.farc_data = {
        'inicio_api': dt_base.strftime('%d/%m/%Y'),
        'fim_api': data_fim_calculo.strftime('%d/%m/%Y'),
        'exibicao_inicio': dt_base.strftime('%m/%Y'),
        'exibicao_fim': data_fim_calculo.strftime('%m/%Y'),
        'valor': valor_contrato,
        'cod_bcb': tipo_idx.split('(')[1].replace('Série ', '').replace(')', '') if '(' in tipo_idx else None,
        'admissivel': intersticio_ok and not precluso
    }

with tab_calc:
    data = st.session_state.farc_data
    if not data or data.get('valor') == 0:
        st.info("Aguardando dados de admissibilidade.")
    else:
        st.subheader("Memória de Cálculo (Intervalo Fixo de 12 Meses)")
        st.write(f"Variação apurada de **{data['exibicao_inicio']}** a **{data['exibicao_fim']}**")
        
        if data['cod_bcb']:
            df_idx, erro = get_index_data(data['cod_bcb'], data['inicio_api'], data['fim_api'])
            if df_idx is not None:
                # Cálculo da variação acumulada composta
                variacao_total = (1 + df_idx['valor']).prod() - 1
                valor_final = data['valor'] * (1 + variacao_total)
                
                m1, m2 = st.columns(2)
                m1.metric("Variação Acumulada (12 meses)", f"{variacao_total:.6%}")
                m2.metric("Novo Valor", f"R$ {valor_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                with st.expander("Ver Detalhamento Mensal (Trava de 12 meses)"):
                    df_exibir = df_idx.copy()
                    df_exibir['data'] = df_exibir['data'].dt.strftime('%m/%Y')
                    st.table(df_exibir) # Tabela fixa para conferência rápida
        else:
            st.warning("Para IST, insira a variação manual baseada em 12 meses.")