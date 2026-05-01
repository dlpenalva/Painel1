import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="GCC - Telebras", layout="wide")

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
        
        # Garantia de 12 meses partindo do INÍCIO (Data-Base)
        # Se a API trouxer mais que 12, pegamos os 12 primeiros a partir da data base informada
        if len(df) > 12:
            df = df.head(12)
        return df, None
    except:
        return None, "Erro API"

def calc_ist_csv(dt_base, dt_aniv):
    try:
        df = pd.read_csv('ist.csv', sep=None, engine='python', decimal=',')
        df.columns = df.columns.str.replace('^\ufeff', '', regex=True)
        meses_map = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
        ref_base = f"{meses_map[dt_base.month]}/{str(dt_base.year)[2:]}"
        # No IST, o aniversário é o mês de referência final (12 meses depois)
        ref_aniv = f"{meses_map[dt_aniv.month]}/{str(dt_aniv.year)[2:]}"
        v_base = float(df[df['MES_ANO'] == ref_base]['INDICE_NIVEL'].values[0])
        v_aniv = float(df[df['MES_ANO'] == ref_aniv]['INDICE_NIVEL'].values[0])
        var = (v_aniv / v_base) - 1
        return var, ref_base, ref_aniv, None
    except:
        return None, None, None, "Erro nas referências do IST"

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Admissibilidade e Variação de Reajuste")

if 'farc' not in st.session_state: st.session_state.farc = {}

tab_adm, tab_calc, tab_rel = st.tabs(["Análise de Admissibilidade", "Cálculo de Reajuste", "Relatório"])

with tab_adm:
    col1, col2 = st.columns(2)
    with col1:
        dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 5, 1), format="DD/MM/YYYY")
        dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
    with col2:
        valor_base = st.number_input("Valor Atual (R$):", min_value=0.0, step=100.0)
        tipo_idx = st.selectbox("Índice:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Planilha CSV)"])

    dt_aniv = dt_base + relativedelta(years=1)
    # Trava do cálculo: sempre um mês antes do aniversário para evitar o 13º mês da API
    dt_fim_calculo = dt_aniv - relativedelta(months=1)
    
    dias_janela = (dt_solic - dt_aniv).days
    intersticio_ok = dt_solic >= dt_aniv
    mesmo_mes = (dt_solic.month == dt_aniv.month and dt_solic.year == dt_aniv.year)

    if dias_janela > 90: status = "Precluso"
    elif not intersticio_ok and mesmo_mes: status = "Admissível (Ajuste Prévio)"
    elif not intersticio_ok: status = "Antecipado"
    else: status = "Admissível"

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("Dias da Janela", f"{max(0, dias_janela)} dias")
    
    # Correção do erro visual (removendo o código DeltaGenerator que aparecia no print)
    if "Admissível" in status: c2.success(f"Status: {status}")
    elif "Precluso" in status: c2.error(f"Status: {status}")
    else: c2.warning(f"Status: {status}")
    
    if intersticio_ok: c3.success("Interstício: Ok")
    else: c3.warning("Interstício: Pendente")

    st.session_state.farc = {
        'dt_base': dt_base, 'dt_aniv': dt_aniv, 'dt_pedido': dt_solic,
        'dt_fim_calc': dt_fim_calculo, 'valor': valor_base, 'idx': tipo_idx, 
        'status': status, 'dias': dias_janela
    }

with tab_calc:
    f = st.session_state.farc
    if f.get('valor', 0) > 0:
        if "IST" in f['idx']:
            var, rb, ra, erro = calc_ist_csv(f['dt_base'], f['dt_aniv'])
            if not erro:
                v_novo = f['valor'] * (1 + var)
                st.metric(f"Variação IST ({rb} a {ra})", f"{var:.6%}")
                st.metric("Valor Reajustado", f"R$ {v_novo:,.2f}")
                st.session_state.farc.update({'var': var, 'v_novo': v_novo})
        else:
            cod = "433" if "IPCA" in f['idx'] else "189"
            # Consulta forçada para terminar no mês anterior ao aniversário
            df, erro = get_index_data(cod, f['dt_base'].strftime('%d/%m/%Y'), f['dt_fim_calc'].replace(day=28).strftime('%d/%m/%Y'))
            if df is not None:
                var = (1 + df['valor']).prod() - 1
                v_novo = f['valor'] * (1 + var)
                st.metric(f"Período: {df.iloc[0]['data'].strftime('%m/%Y')} a {df.iloc[-1]['data'].strftime('%m/%Y')}", f"{var:.6%}")
                st.metric("Valor Reajustado", f"R$ {v_novo:,.2f}")
                st.dataframe(df.assign(data=df['data'].dt.strftime('%m/%Y')), use_container_width=True)
                st.session_state.farc.update({'var': var, 'v_novo': v_novo})

with tab_rel:
    f = st.session_state.farc
    if f.get('v_novo'):
        st.subheader("Minuta para o SEI")
        # Mantendo o relatório completo que você aprovou
        texto = f"""RELATÓRIO TÉCNICO DE REAJUSTE

1. ANÁLISE DE ADMISSIBILIDADE
- Data-Base Anterior: {f['dt_base'].strftime('%d/%m/%Y')}
- Aniversário do Direito: {f['dt_aniv'].strftime('%d/%m/%Y')}
- Data do Pedido: {f['dt_pedido'].strftime('%d/%m/%Y')}
- Status: {f['status']}

2. MEMÓRIA DE CÁLCULO
- Índice Utilizado: {f['idx']}
- Variação Apurada: {f['var']:.6%}
- Valor Anterior: R$ {f['valor']:,.2f}
- Valor Reajustado: R$ {f['v_novo']:,.2f}

3. CONCLUSÃO
O pedido encontra-se {f['status']}. O novo valor contratual de R$ {f['v_novo']:,.2f} passa a vigorar a partir de {f['dt_aniv'].strftime('%d/%m/%Y')}."""
        st.text_area("Copie o texto abaixo:", texto, height=350)