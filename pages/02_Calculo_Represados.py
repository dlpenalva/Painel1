def get_ist_data(dt_i, dt_f):
    try:
        df = pd.read_csv('ist.csv', sep=None, engine='python', header=0)
        df.columns = ['MES_ANO_COL', 'INDICE_COL'] + list(df.columns[2:])
        
        meses_map = {'jan':'Jan','fev':'Feb','mar':'Mar','abr':'Apr','mai':'May','jun':'Jun',
                     'jul':'Jul','ago':'Aug','set':'Sep','out':'Oct','nov':'Nov','dez':'Dec'}
        
        df['MES_ANO_COL'] = df['MES_ANO_COL'].astype(str).str.strip().str.lower()
        for pt, en in meses_map.items(): 
            df['MES_ANO_COL'] = df['MES_ANO_COL'].str.replace(pt, en)
            
        df['DATA_DT'] = pd.to_datetime(df['MES_ANO_COL'], format='%b/%y', errors='coerce')
        df['INDICE_VAL'] = df['INDICE_COL'].astype(str).str.replace('.', '').str.replace(',', '.').astype(float)
        
        v_i = df[df['DATA_DT'] == pd.to_datetime(dt_i.strftime('%Y-%m-01'))]['INDICE_VAL'].values[0]
        v_f = df[df['DATA_DT'] == pd.to_datetime(dt_f.strftime('%Y-%m-01'))]['INDICE_VAL'].values[0]
        return (v_f / v_i) - 1, v_i, v_f
    except:
        return None, None, None