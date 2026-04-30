import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

# Configuração de Interface e Identificação
st.set_page_config(page_title="Admissibilidade e Variação de Reajuste", layout="wide")

# Estilo para formatar métricas e textos
st.markdown("""
    <style>
    .main { background-color: #f5f7f9; }
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
    </style>
    """, unsafe_allow_stdio=True)

def get_index_data(serie_codigo, data_inicio, data_fim):
    """
    Automação da Busca de Índices via API SGS/BCB.
    Tratamento de erro para índices ainda não publicados.
    """
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        df = pd.DataFrame(response.json())
        
        if df.empty:
            return None, "Nenhum dado encontrado para o período informado."
            
        df['valor'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        
        # Verificação de atualização do índice
        ultima_data_api = df['data'].max()
        data_fim_dt = pd.to_datetime(data_fim, dayfirst=True)
        
        aviso = None
        if ultima_data_api.month < data_fim_dt.month or ultima_data_api.year < data_fim_dt.year:
            aviso = f"Nota: O índice de {data_fim_dt.strftime('%m/%Y')} ainda não foi publicado oficialmente."
            
        return df, aviso
    except Exception as e:
        return None, f"Erro de conexão com a API oficial: {str(e)}"

# Inicialização do estado da sessão
if 'farc_data' not in st.session_state:
    st.session_state.farc_data = {}

st.title("⚖️ Admissibilidade e Variação de Reajuste")
st.caption("Ferramenta de apoio à Gestão de Contratos - GCC")

tab_adm, tab_calc, tab_rel = st.tabs(["Análise de Admissibilidade", "Cálculo de Reajuste", "Minuta de Relatório"])

with tab_adm:
    st.subheader("Critérios da Cláusula Oitava")
    col1, col2 = st.columns(2)
    
    with col1:
        dt_base = st.date_input("Data do Último Reajuste (ou Aniversário):", format="DD/MM/YYYY")
        dt_solic = st.date_input("Data do Protocolo da Solicitação:", format="DD/MM/YYYY")
    
    with col2:
        valor_contrato = st.number_input("Valor a Reajustar (R$):", min_value=0.0, step=100.0, help="Pode ser o valor global ou o saldo remanescente.")
        tipo_idx = st.selectbox("Índice de Reajuste:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Inserção Manual)"])

    # Lógica de Negócio: Interstício e Preclusão
    dias_atraso = (dt_solic - dt_base).days
    intersticio_ok = relativedelta(dt_solic, dt_base).years >= 1
    precluso = dias_atraso > 90 

    st.divider()
    
    c1, c2, c3 = st.columns(3)
    c1.metric("Dias desde o Aniversário", f"{dias_atraso} dias")
    
    if precluso:
        c2.error("Status: Precluso")
        st.error(f"O pedido foi realizado {dias_atraso} dias após o direito, superando o limite de 90 dias.")
    else:
        c2.success("Status: Tempestivo")

    if not intersticio_ok:
        c3.warning("Interstício: Não atingido")
    else:
        c3.success("Interstício: Ok")

    # Sincronização de dados para as próximas abas
    st.session_state.farc_data = {
        'inicio': dt_base.strftime('%d/%m/%Y'),
        'fim': dt_solic.strftime('%d/%m/%Y'),
        'valor': valor_contrato,
        'indice_nome': tipo_idx,
        'cod_bcb': tipo_idx.split('(')[1].replace('Série ', '').replace(')', '') if '(' in tipo_idx else None,
        'admissivel': not precluso and intersticio_ok
    }

with tab_calc:
    data = st.session_state.farc_data
    if not data or data.get('valor') == 0:
        st.info("Aguardando preenchimento dos dados de Admissibilidade.")
    else:
        st.subheader("Memória de Cálculo")
        st.info(f"**Período de Apuração:** {data['inicio']} a {data['fim']}")
        
        if data['cod_bcb']:
            with st.spinner("Buscando índices oficiais..."):
                df_idx, aviso = get_index_data(data['cod_bcb'], data['inicio'], data['fim'])
            
            if aviso:
                st.warning(aviso)
            
            if df_idx is not None:
                variacao = (1 + df_idx['valor']).prod() - 1
                valor_final = data['valor'] * (1 + variacao)
                impacto = valor_final - data['valor']
                
                m1, m2, m3 = st.columns(3)
                m1.metric("Variação Acumulada", f"{variacao:.4%}")
                m2.metric("Novo Valor Reajustado", f"R$ {valor_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                m3.metric("Impacto Financeiro", f"R$ {impacto:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                
                with st.expander("Ver Detalhamento Mensal"):
                    st.dataframe(df_idx.assign(data=df_idx['data'].dt.strftime('%m/%Y')), use_container_width=True)
        else:
            taxa_ist = st.number_input("Informe a variação acumulada do IST (%):", step=0.0001) / 100
            if taxa_ist > 0:
                v_final = data['valor'] * (1 + taxa_ist)
                st.metric("Novo Valor (IST)", f"R$ {v_final:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

with tab_rel:
    if not st.session_state.farc_data.get('admissivel'):
        st.warning("O relatório não pode ser gerado devido ao descumprimento dos critérios de Admissibilidade.")
    else:
        st.subheader("Minuta para o SEI")
        texto_sei = f"""
        OBJETO: Reajuste Contratual.
        PERÍODO DE APURAÇÃO: {data['inicio']} a {data['fim']}.
        
        1. DA ADMISSIBILIDADE
        Verifica-se que a solicitação foi protocolada em {data['fim']}, respeitando o interstício de 12 meses 
        em relação ao último evento contratual ({data['inicio']}) e dentro do prazo de 90 dias para evitar a preclusão.
        
        2. DO CÁLCULO
        Utilizando o índice {data['indice_nome']}, apurou-se a variação acumulada no período. 
        O valor base de R$ {data['valor']:,.2f} resulta no novo valor reajustado de R$ {valor_final:,.2f}.
        """
        st.text_area("Copie o texto abaixo:", texto_sei, height=300)