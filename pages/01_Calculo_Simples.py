import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo Simples", layout="wide")

def get_index_data_bc(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=10)
        df = pd.DataFrame(response.json())
        if df.empty: return None
        df['valor'] = df['valor'].astype(float) / 100
        return df
    except: return None

def get_ist_safe(dt_inicio, dt_fim):
    try:
        # Carregamento do CSV
        df = pd.read_csv('ist.csv', sep=None, engine='python', header=0)
        df.columns = ['MES_ANO_BRUTO', 'INDICE_BRUTO'] + list(df.columns[2:])
        
        # Normalização de Meses
        meses_map = {'jan': 'Jan', 'fev': 'Feb', 'mar': 'Mar', 'abr': 'Apr', 'mai': 'May', 'jun': 'Jun',
                     'jul': 'Jul', 'ago': 'Aug', 'set': 'Sep', 'out': 'Oct', 'nov': 'Nov', 'dez': 'Dec'}
        
        df['MES_ANO_CLEAN'] = df['MES_ANO_BRUTO'].astype(str).str.strip().str.lower()
        for pt, en in meses_map.items():
            df['MES_ANO_CLEAN'] = df['MES_ANO_CLEAN'].str.replace(pt, en)
        
        df['DATA_DT'] = pd.to_datetime(df['MES_ANO_CLEAN'], format='%b/%y', errors='coerce')
        df['INDICE_VAL'] = df['INDICE_BRUTO'].astype(str).str.replace('.', '').str.replace(',', '.').astype(float)
        df = df.dropna(subset=['DATA_DT', 'INDICE_VAL'])

        # Busca de Valores
        t_ini = pd.to_datetime(dt_inicio.strftime('%Y-%m-01'))
        t_fim = pd.to_datetime(dt_fim.strftime('%Y-%m-01'))
        
        row_i = df[df['DATA_DT'] == t_ini]
        row_f = df[df['DATA_DT'] == t_fim]

        if row_i.empty or row_f.empty:
            return None, f"Datas não encontradas no CSV: {dt_inicio.strftime('%m/%y')} ou {dt_fim.strftime('%m/%y')}", 0, 0
            
        v_i = row_i['INDICE_VAL'].values[0]
        v_f = row_f['INDICE_VAL'].values[0]
        return (v_f / v_i) - 1, None, v_i, v_f
    except Exception as e:
        return None, f"Erro na leitura do arquivo: {str(e)}", 0, 0

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo Simples")

dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 5, 1), format="DD/MM/YYYY")
dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
tipo_idx = st.selectbox("Índice:", ["IST (Série Local)", "IPCA (Série 433)", "IGP-M (Série 189)"])

# Definição do Ciclo (12 meses)
dt_fim_calc = dt_base + relativedelta(months=11)
dt_aniv = dt_base + relativedelta(years=1)
limite_90 = dt_aniv + relativedelta(days=90)
status = "✅ ADMISSÍVEL" if dt_solic <= limite_90 else "❌ PRECLUSO"

st.markdown("---")
st.subheader("Resultado da Análise")

try:
    if "IST" in tipo_idx:
        var, erro, vi, vf = get_ist_safe(dt_base, dt_fim_calc)
        if var is not None:
            st.latex(r"IST = \left( \frac{" + f"{vf:,.3f}" + r"}{" + f"{vi:,.3f}" + r"} \right) - 1")
            st.metric("Variação IST", f"{var*100:,.2f}%".replace('.', ','))
            st.table(pd.DataFrame({"Item": ["Data Inicial", "Data Final", "Limite", "Status"], 
                                 "Valor": [dt_base.strftime('%m/%Y'), dt_fim_calc.strftime('%m/%Y'), limite_90.strftime('%d/%m/%Y'), status]}))
        else:
            st.warning(erro)
    else:
        cod = "433" if "IPCA" in tipo_idx else "189"
        df_bc = get_index_data_bc(cod, dt_base.strftime('%d/%m/%Y'), dt_fim_calc.strftime('%d/%m/%Y'))
        if df_bc is not None:
            var = (1 + df_bc['valor']).prod() - 1
            st.metric(f"Variação {tipo_idx}", f"{var*100:,.2f}%".replace('.', ','))
            st.info(f"Status: {status}")
except Exception as e:
    st.error(f"Erro inesperado no processamento: {e}")