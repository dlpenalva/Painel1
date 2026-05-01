import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo Simples", layout="wide")

def get_index_data_bc(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=15)
        df = pd.DataFrame(response.json())
        if df.empty: return None
        df['valor'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        return df
    except: return None

def get_ist_local(dt_inicio, dt_fim):
    try:
        df = pd.read_csv('ist.csv', sep=None, engine='python')
        # Limpeza básica de nomes de colunas
        df.columns = [c.strip().upper() for c in df.columns]
        
        meses_map = {'jan': 'Jan', 'fev': 'Feb', 'mar': 'Mar', 'abr': 'Apr', 'mai': 'May', 'jun': 'Jun',
                     'jul': 'Jul', 'ago': 'Aug', 'set': 'Sep', 'out': 'Oct', 'nov': 'Nov', 'dez': 'Dec'}
        
        for pt, en in meses_map.items():
            df['MES_ANO'] = df['MES_ANO'].str.lower().str.replace(pt, en)
        
        df['DATA_DT'] = pd.to_datetime(df['MES_ANO'], format='%b/%y')
        df['INDICE_NIVEL'] = df['INDICE_NIVEL'].astype(str).str.replace('.', '').str.replace(',', '.').astype(float)
        
        target_inicio = pd.to_datetime(dt_inicio.strftime('%Y-%m-01'))
        target_fim = pd.to_datetime(dt_fim.strftime('%Y-%m-01'))
        
        val_i = df[df['DATA_DT'] == target_inicio]['INDICE_NIVEL'].values[0]
        val_f = df[df['DATA_DT'] == target_fim]['INDICE_NIVEL'].values[0]
        
        var_total = (val_f / val_i) - 1
        mask = (df['DATA_DT'] >= target_inicio) & (df['DATA_DT'] <= target_fim)
        return var_total, df.loc[mask], val_i, val_f
    except Exception as e:
        return None, str(e), None, None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo Simples")

col1, col2 = st.columns(2)
with col1:
    dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 5, 1), format="DD/MM/YYYY")
    dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
with col2:
    tipo_idx = st.selectbox("Índice:", ["IST (Série Local)", "IPCA (Série 433)", "IGP-M (Série 189)"])

dt_fim_calc = dt_base + relativedelta(months=11)
dt_aniv = dt_base + relativedelta(years=1)
limite_90 = dt_aniv + relativedelta(days=90)
status = "✅ ADMISSÍVEL" if dt_solic <= limite_90 else "❌ PRECLUSO"

st.subheader("Resultado da Análise")

if "IST" in tipo_idx:
    var, df_detalhe, v_ini, v_fim = get_ist_local(dt_base, dt_fim_calc)
    if var is not None:
        st.latex(r"IST = \left( \frac{\text{Índice}_{" + dt_fim_calc.strftime('%m/%Y') + r"}}{\text{Índice}_{" + dt_base.strftime('%m/%Y') + r"}} \right) - 1")
        st.metric("Variação IST (12 meses)", f"{var*100:,.2f}%".replace('.', ','))
        
        st.subheader("Memória de Cálculo (Fator de Prova)")
        resumo = {
            "Parâmetro": ["Índice Inicial", "Índice Final", "Variação", "Limite Admissibilidade", "Status"],
            "Valor": [f"{v_ini:,.3f}", f"{v_fim:,.3f}", f"{var*100:,.2f}%", limite_90.strftime('%d/%m/%Y'), status]
        }
        st.table(pd.DataFrame(resumo))
    else:
        st.error(f"Erro ao ler IST: {df_detalhe}")
else:
    cod = "433" if "IPCA" in tipo_idx else "189"
    df_bc = get_index_data_bc(cod, dt_base.strftime('%d/%m/%Y'), dt_fim_calc.strftime('%d/%m/%Y'))
    if df_bc is not None:
        var = (1 + df_bc['valor']).prod() - 1
        st.metric(f"Variação {tipo_idx}", f"{var*100:,.2f}%".replace('.', ','))
        st.table(pd.DataFrame({"Item": ["Data do Pedido", "Limite", "Status"], "Valor": [dt_solic.strftime('%d/%m/%Y'), limite_90.strftime('%d/%m/%Y'), status]}))