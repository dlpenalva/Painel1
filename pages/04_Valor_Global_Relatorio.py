def calcular_relatorio_global(df_estoque, df_financeiro, fatores_ciclos):
    """
    df_estoque: colunas [item, qtd_c1, vu_original]
    df_financeiro: colunas [ciclo, valor_pago]
    fatores_ciclos: dicionário { 'C1': 1.05, 'C2': 1.10 }
    """
    
    # 1. Processamento de Estoque
    df_estoque['VU_C1'] = df_estoque['vu_original'] * fatores_ciclos['C1']
    df_estoque['Consumo_C1_Reaj'] = df_estoque['qtd_c1'] * df_estoque['VU_C1']
    
    # 2. Processamento Financeiro
    pago_c1 = df_financeiro.loc[df_financeiro['ciclo'] == 'C1', 'valor_pago'].sum()
    devido_c1 = pago_c1 * fatores_ciclos['C1']
    delta_c1 = devido_c1 - pago_c1
    
    # 3. Comparativo Físico x Financeiro
    consumo_fisico_c1 = df_estoque['Consumo_C1_Reaj'].sum()
    divergencia = consumo_fisico_c1 - devido_c1
    
    return {
        "devido": devido_c1,
        "pago": pago_c1,
        "delta": delta_c1,
        "estoque": consumo_fisico_c1,
        "divergencia": divergencia
    }