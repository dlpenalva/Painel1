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
        return df, None
    except:
        return None, "Erro API"

def calc_ist_csv(dt_base, dt_aniv):
    try:
        df = pd.read_csv('ist.csv', decimal=',', sep=None, engine='python')
        # Mapeamento para garantir match com o CSV (ex: jan/22)
        meses_map = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
        
        ref_base = f"{meses_map[dt_base.month]}/{str(dt_base.year)[2:]}"
        ref_aniv = f"{meses_map[dt_aniv.month]}/{str(dt_aniv.year)[2:]}"
        
        idx_base = float(df[df['MES_ANO'] == ref_base]['INDICE_NIVEL'].values[0])
        idx_aniv = float(df[df['MES_ANO'] == ref_aniv]['INDICE_NIVEL'].values[0])
        
        var = (idx_aniv / idx_base) - 1
        return var, f"IST ({ref_base} a {ref_aniv})", None
    except Exception as e:
        return None, None, f"Erro no CSV: {str(e)}"

if 'farc_data' not in st.session_state:
    st.session_state.farc_data = {}

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Admissibilidade e Variação de Reajuste")

tab_adm, tab_calc, tab_rel = st.tabs(["Análise de Admissibilidade", "Cálculo de Reajuste", "Relatório"])

with tab_adm:
    col1, col2 = st.columns(2)
    with col1:
        dt_base = st.date_input("Data-Base:", value=datetime(2023, 5, 1), format="DD/MM/YYYY")
        dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
    with col2:
        valor_contrato = st.number_input("Valor a Reajustar (R$):", min_value=0.0, step=100.0)
        tipo_idx = st.selectbox("Índice de Reajuste:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Planilha CSV)"])

    # Lógica de Prazos (Configuração Perfeita do Usuário)
    data_aniversario = dt_base + relativedelta(years=1)
    data_fim_calc = data_aniversario - relativedelta(days=1)
    dias_janela = (dt_solic - data_aniversario).days
    
    # Novo Status: Admissibilidade Prévia (Problema 2)
    mesmo_mes_aniv = (dt_solic.month == data_aniversario.month and dt_solic.year == data_aniversario.year)
    intersticio_ok = dt_solic >= data_aniversario
    
    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("Dias da Janela", f"{max(0, dias_janela)} dias")
    
    if dias_janela > 90:
        c2.error("Status: Precluso")
        status_final = "Precluso"
    elif not intersticio_ok and mesmo_mes_aniv:
        c2.warning("Status: Admissível (Ajuste Prévio)")
        status_final = "Admissível (Ajuste Prévio)"
    elif not intersticio_ok:
        c2.warning("Status: Antecipado")
        status_final = "Antecipado"
    else:
        c2.success("Status: Admissível")
        status_final = "Admissível"
    
    c3.success("Interstício: Ok") if intersticio_ok else c3.warning("Interstício: Pendente")

    st.session_state.farc_data = {
        'dt_base': dt_base, 'dt_aniv': data_aniversario, 'dt_pedido': dt_solic,
        'inicio_api': dt_base.strftime('%d/%m/%Y'), 'fim_api': data_fim_calc.strftime('%d/%m/%Y'),
        'valor': valor_contrato, 'indice': tipo_idx, 'status': status_final
    }

with tab_calc:
    d = st.session_state.farc_data
    if d.get('valor', 0) > 0:
        st.subheader(f"Cálculo da Variação — {d['indice']}")
        st.info(f"**Período de Referência:** {d['inicio_api']} a {d['fim_api']}")
        
        if "IST" in d['indice']:
            var, desc, erro = calc_ist_csv(d['dt_base'], d['dt_aniv'])
            if erro: st.error(erro)
            else:
                v_novo = d['valor'] * (1 + var)
                st.metric("Variação Acumulada (IST)", f"{var:.6%}")
                st.metric("Valor Atualizado", f"R$ {v_novo:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                st.session_state.farc_data.update({'var_acum': var, 'v_novo': v_novo})
        else:
            cod = "433" if "IPCA" in d['indice'] else "189"
            df, erro = get_index_data(cod, d['inicio_api'], d['fim_api'])
            if df is not None:
                var_acum = (1 + df['valor']).prod() - 1
                v_novo = d['valor'] * (1 + var_acum)
                st.metric("Variação Acumulada", f"{var_acum:.6%}")
                st.metric("Valor Atualizado", f"R$ {v_novo:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                st.dataframe(df.assign(data=df['data'].dt.strftime('%m/%Y')), use_container_width=True)
                st.session_state.farc_data.update({'var_acum': var_acum, 'v_novo': v_novo})