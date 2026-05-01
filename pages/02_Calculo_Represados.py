import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

def get_ist_represado(dt_inicio, dt_fim):
    try:
        df = pd.read_csv('ist.csv', sep=None, engine='python')
        meses_map = {'jan':'Jan','fev':'Feb','mar':'Mar','abr':'Apr','mai':'May','jun':'Jun',
                     'jul':'Jul','ago':'Aug','set':'Sep','out':'Oct','nov':'Nov','dez':'Dec'}
        for pt, en in meses_map.items():
            df['MES_ANO'] = df['MES_ANO'].str.replace(pt, en)
        
        df['data'] = pd.to_datetime(df['MES_ANO'], format='%b/%y')
        df['INDICE_NIVEL'] = df['INDICE_NIVEL'].str.replace('.', '').str.replace(',', '.').astype(float)
        
        val_i = df[df['data'] == pd.to_datetime(dt_inicio.strftime('%Y-%m-01'))]['INDICE_NIVEL'].values[0]
        val_f = df[df['data'] == pd.to_datetime(dt_fim.strftime('%Y-%m-01'))]['INDICE_NIVEL'].values[0]
        
        var_ciclo = (val_f / val_i) - 1
        mask = (df['data'] >= pd.to_datetime(dt_inicio)) & (df['data'] <= pd.to_datetime(dt_fim))
        return var_ciclo, df.loc[mask]
    except: return None, None

# No loop dos ciclos, a exibição da memória de cálculo usará o df_detalhado retornado