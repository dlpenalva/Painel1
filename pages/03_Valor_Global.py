with tab1:
    uploaded_file = st.file_uploader("Suba a planilha preenchida pelo fiscal", type="xlsx")
    if uploaded_file:
        # Aviso de sucesso restaurado
        st.success("Planilha processada com sucesso!")
        
        df_itens = pd.read_excel(uploaded_file, sheet_name="ITENS_CICLOS", skiprows=2).dropna(subset=["Item"])
        df_retro = pd.read_excel(uploaded_file, sheet_name="RETROATIVO", skiprows=2).dropna(subset=["Valor bruto faturado (R$)"])

        v_orig = (df_itens["Quantidade"] * df_itens["VU C0 (R$)"]).sum()
        faturado = df_retro.iloc[:, 1].sum()
        col_rem = df_itens.columns[-1]
        remanescente_reaj = (df_itens[col_rem] * df_itens["VU C0 (R$)"] * fator_vigente).sum()
        global_estimado = faturado + remanescente_reaj

        st.metric("VALOR GLOBAL ESTIMADO", f"R$ {global_estimado:,.2f}")
        st.session_state['balanco'] = {'orig': v_orig, 'final': global_estimado, 'fat': faturado}