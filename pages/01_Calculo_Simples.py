def get_ist_local(dt_inicio, dt_fim):
    try:
        # Lê o CSV sem assumir cabeçalhos fixos para evitar o erro 'MES_ANO'
        df = pd.read_csv('ist.csv', sep=None, engine='python', header=0)
        
        # Forçamos a nomeação das colunas pela posição (A e B)
        df.columns = ['MES_ANO_COL', 'INDICE_COL'] + list(df.columns[2:])
        
        # Limpeza de dados
        df['MES_ANO_COL'] = df['MES_ANO_COL'].astype(str).str.strip().str.lower()
        
        meses_map = {'jan': 'Jan', 'fev': 'Feb', 'mar': 'Mar', 'abr': 'Apr', 'mai': 'May', 'jun': 'Jun',
                     'jul': 'Jul', 'ago': 'Aug', 'set': 'Sep', 'out': 'Oct', 'nov': 'Nov', 'dez': 'Dec'}
        
        for pt, en in meses_map.items():
            df['MES_ANO_COL'] = df['MES_ANO_COL'].str.replace(pt, en)
        
        # Conversão de datas e valores
        df['DATA_DT'] = pd.to_datetime(df['MES_ANO_COL'], format='%b/%y', errors='coerce')
        df['INDICE_VAL'] = df['INDICE_COL'].astype(str).str.replace('.', '').str.replace(',', '.').astype(float)
        
        # Remove linhas que falharam na conversão
        df = df.dropna(subset=['DATA_DT', 'INDICE_VAL'])

        target_inicio = pd.to_datetime(dt_inicio.strftime('%Y-%m-01'))
        target_fim = pd.to_datetime(dt_fim.strftime('%Y-%m-01'))
        
        val_i = df[df['DATA_DT'] == target_inicio]['INDICE_VAL'].values[0]
        val_f = df[df['DATA_DT'] == target_fim]['INDICE_VAL'].values[0]
        
        var_total = (val_f / val_i) - 1
        return var_total, df, val_i, val_f
    except Exception as e:
        return None, f"Erro técnico: {str(e)}", None, None