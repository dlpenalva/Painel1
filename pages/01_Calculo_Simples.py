import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo Simples", layout="wide")

def get_ist_local(dt_inicio, dt_fim):
    try:
        # Lê o CSV tratando o separador decimal e nomes de colunas
        df = pd.read_csv('ist.csv', sep=None, engine='python')
        # Mapeamento de meses em português para o pandas entender
        meses_map = {
            'jan': 'Jan', 'fev': 'Feb', 'mar': 'Mar', 'abr': 'Apr', 'mai': 'May', 'jun': 'Jun',
            'jul': 'Jul', 'ago': 'Aug', 'set': 'Sep', 'out': 'Oct', 'nov': 'Nov', 'dez': 'Dec'
        }
        for pt, en in meses_map.items():
            df['MES_ANO'] = df['MES_ANO'].str.replace(pt, en)
        
        df['data'] = pd.to_datetime(df['MES_ANO'], format='%b/%y')
        df['INDICE_NIVEL'] = df['INDICE_NIVEL'].str.replace('.', '').str.replace(',', '.').astype(float)
        
        # Busca o valor do mês inicial e do mês final
        idx_inicio = df[df['data'] == pd.to_datetime(dt_inicio.strftime('%Y-%m-01'))]['INDICE_NIVEL'].values[0]
        idx_fim = df[df['data'] == pd.to_datetime(dt_fim.strftime('%Y-%m-01'))]['INDICE_NIVEL'].values[0]
        
        var_total = (idx_fim / idx_inicio) - 1
        
        # Filtra apenas o intervalo para a memória de cálculo
        mask = (df['data'] >= pd.to_datetime(dt_inicio)) & (df['data'] <= pd.to_datetime(dt_fim))
        return var_total, df.loc[mask]
    except: return None, None

# ... (restante da lógica de UI permanece igual)