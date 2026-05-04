import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo Simples", layout="wide")

def get_index_data(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio.strftime('%d/%m/%Y')}&dataFinal={data_fim.strftime('%d/%m/%Y')}"
    try:
        response = requests.get(url, timeout=15)
        df = pd.DataFrame(response.json())
        if df.empty: return None
        df['valor_decimal'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        var_final = (1 + df['valor_decimal']).prod() - 1
        return {'variacao': var_final, 'metodo': "Produtório de taxas mensais (SGS/BCB)", 'dados': df[['data', 'valor']]}
    except: return None

def get_ist_local(data_inicio, data_fim):
    try:
        df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
        df.columns = [str(col).strip().lower() for col in df.columns]
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        r_ini = (data_inicio - relativedelta(months=1)).replace(day=1)
        r_fim = data_fim.replace(day=1)
        v_ini = df[df['data'].dt.to_period('M') == r_ini.strftime('%Y-%m')]['indice'].values[0]
        v_fim = df[df['data'].dt.to_period('M') == r_fim.strftime('%Y-%m')]['indice'].values[0]
        return {'variacao': (v_fim / v_ini) - 1, 'i_ini': v_ini, 'i_fim': v_fim, 'd_ini': r_ini, 'd_fim': r_fim, 'metodo': "Divisão de Número-Índice (Série Local)"}
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Reajuste Simples")

col1, col2 = st.columns(2)
with col1:
    dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 8, 2), format="DD/MM/YYYY")
    dt_solic = st.date_input("Data do Pedido:", value=datetime(2024, 4, 9), format="DD/MM/YYYY")
with col2:
    tipo_idx = st.selectbox("Índice:", ["IST (Série Local)", "IPCA (433)", "IGP-M (189)"])

# Definição de datas do ciclo
dt_fim_ap = dt_base + relativedelta(months=11)
dt_aniv = dt_base + relativedelta(years=1)
dt_limite = dt_aniv + relativedelta(days=90)

# Regra de Admissibilidade
if dt_solic < dt_aniv:
    if dt_solic.year == dt_aniv.year and dt_solic.month == dt_aniv.month:
        status_ped = "🟡 ADMISSÍVEL - RESSALVA"
    else:
        status_ped = "⚠️ ADIANTADO (ANTES DOS 12 MESES)"
elif dt_solic <= dt_limite:
    status_ped = "✅ TEMPESTIVO"
else:
    status_ped = "❌ PRECLUSO"

res = get_ist_local(dt_base, dt_fim_ap) if "IST" in tipo_idx else get_index_data("433" if "IPCA" in tipo_idx else "189", dt_base, dt_fim_ap)

if res:
    v_fmt = f"{res['variacao']*100:,.2f}%".replace('.', ',')
    st.metric("Variação Apurada", v_fmt)
    
    st.markdown("### Dados do Ciclo")
    # Ajuste solicitado de nomenclatura e nova linha de janela
    janela_str = f"{res['d_ini'].strftime('%m/%Y') if 'd_ini' in res else dt_base.strftime('%m/%Y')} a {res['d_fim'].strftime('%m/%Y') if 'd_fim' in res else dt_fim_ap.strftime('%m/%Y')}"
    st.write(f"- **Intervalo do C1:** {janela_str}")
    st.write(f"- **Janela de Admissibilidade:** {dt_aniv.strftime('%d/%m/%Y')} a {dt_limite.strftime('%d/%m/%Y')}")
    st.write(f"- **Situação:** {status_ped}")

    with st.expander("🔍 Memória de Cálculo Detalhada"):
        st.write(f"**Metodologia:** {res['metodo']}")
        if "IST" in tipo_idx:
            st.code(f"({res['i_fim']} / {res['i_ini']}) - 1 = {res['variacao']*100:.4f}%")
        else:
            st.dataframe(res['dados'])

    st.subheader("Relatório de Apuração")
    relatorio_simples = f"""
    **C1:** Pedido realizado em {dt_solic.strftime('%d/%m/%Y')}. Intervalo do C1: {janela_str}.  
    Janela de Admissibilidade: {dt_aniv.strftime('%d/%m/%Y')} a {dt_limite.strftime('%d/%m/%Y')}.  
    Resultado: {status_ped}.  
    Variação do Ciclo: {v_fmt}. Variação acumulada: {v_fmt}.  
    Índice {tipo_idx}.  
    Data de início dos efeitos financeiros: {dt_solic.strftime('%d/%m/%Y')}.
    """
    st.info(relatorio_simples)