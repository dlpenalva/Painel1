import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Configuração de Interface
st.set_page_config(page_title="Gestão de Contratos - Telebras", layout="wide")

# Estilo CSS
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    img { margin-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)

def get_index_data(serie_codigo, data_inicio, data_fim):
    """Busca dados na API SGS/BCB."""
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

# Logo Telebras e Título
st.image("https://www.telebras.com.br/wp-content/uploads/2021/04/logo-telebras-horizontal.png", width=250)
st.title("Admissibilidade e Variação de Reajuste")
st.caption("Ferramenta de apoio à Gestão de Contratos - Telebras (GCC)")

tab_adm, tab_calc, tab_rel = st.tabs(["Análise de Admissibilidade", "Cálculo de Reajuste", "Minuta de Relatório"])

with tab_adm:
    st.subheader("Critérios da Cláusula Oitava")
    col1, col2 = st.columns(2)
    
    with col1:
        # Nomes dos campos atualizados conforme solicitado
        dt_base = st.date_input("Data-Base (Proposta; Último Reajuste; ou Aniversário Anterior):", value=datetime(2024, 5, 1), format="DD/MM/YYYY")
        dt_solic = st.date_input("Data do Pedido do Reajuste:", format="DD/MM/YYYY")
    
    with col2:
        valor_contrato = st.number_input("Valor a Reajustar (R$):", min_value=0.0, step=100.0)
        tipo_idx = st.selectbox("Índice de Reajuste:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Inserção Manual)"])

    # --- LÓGICA DE DATAS E ADMISSIBILIDADE ---
    data_aniversario_direito = dt_base + relativedelta(years=1)
    
    # Intervalo de variação fixo de 12 meses (Ex: se base é 05/24, fim é 04/25)
    # Subtraímos um dia para garantir que a API pegue até o mês anterior ao aniversário
    data_fim_calc = data_aniversario_direito - relativedelta(days=1)
    
    dias_decorridos = (dt_solic - data_aniversario_direito).days
    intersticio_ok = dt_solic >= data_aniversario_direito
    precluso = dias_decorridos > 90

    st.divider()
    c1, c2, c3 = st.columns(3)
    
    # Card de Dias
    if dias_decorridos < 0:
        c1.metric("Status do Prazo", f"Faltam {abs(dias_decorridos)} dias")
    else:
        c1.metric("Dias da Janela de Reajuste", f"{dias_decorridos} dias")

    # Card de Status
    if precluso:
        c2.error("Status: Precluso")
    elif not intersticio_ok:
        c2.warning("Status: Antecipado")
    else:
        c2.success("Status: Admissível")

    # Card de Interstício (Correção do erro da imagem)
    if intersticio_ok:
        c3.success("Interstício: Ok")
    else:
        c3.warning("Interstício: Pendente")

    # Salva dados para as outras abas
    st.session_state.farc_data = {
        'inicio_api': dt_base.strftime('%d/%m/%Y'),
        'fim_api': data_fim_calc.strftime('%d/%m/%Y'),
        'valor': valor_contrato,
        'cod_bcb': tipo_idx.split('(')[1].replace('Série ', '').replace(')', '') if '(' in tipo_idx else None,
        'admissivel': intersticio_ok and not precluso
    }

with tab_calc:
    data = st.session_state.farc_data
    if not data or data.get('valor') == 0:
        st.info("Aguardando dados de admissibilidade.")
    else:
        st.subheader("Memória de Cálculo (Intervalo de 12 Meses)")
        st.write(f"Variação apurada de: **{data['inicio_api']}** até **{data['fim_api']}**")
        
        if data['cod_bcb']:
            df_idx, erro = get_index_data(data['cod_bcb'], data['inicio_api'], data['fim_api'])
            if df_idx is not None:
                # Variação composta
                variacao_total = (1 + df_idx['valor']).prod() - 1
                valor_final = data['valor'] * (1 + variacao_total)
                
                m1, m2 = st.columns(2)
                m1.metric("Variação Acumulada", f"{variacao_total:.6%}")
                m2.metric("Valor Reajustado", f"R$ {valor_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                with st.expander("Detalhamento dos Índices"):
                    df_excluir = df_idx.copy()
                    df_excluir['data'] = df_excluir['data'].dt.strftime('%m/%Y')
                    st.dataframe(df_excluir, use_container_width=True)
            else:
                st.error(f"Erro ao buscar índices: {erro}")
        else:
            st.info("Para IST, proceda com o cálculo manual na minuta.")

with tab_rel:
    if st.session_state.farc_data.get('admissivel'):
        st.success("Dados prontos para o Relatório Técnico.")
    else:
        st.warning("Verifique a admissibilidade para gerar o relatório.")