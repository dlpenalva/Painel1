# --- 4. GERAÇÃO DE RELATÓRIO (NOTA TÉCNICA) ---
st.divider()
st.subheader("4. Finalização e Relatório")

if 'balanco' in st.session_state:
    b = st.session_state['balanco']
    diff = b['final'] - b['orig']
    perc = (diff / b['orig'] * 100) if b['orig'] > 0 else 0
    
    # Texto Sugerido para a Nota Técnica
    texto_nt = f"""
    NOTA TÉCNICA - REAJUSTE CONTRATUAL {adm['ciclo_atual']}
    
    1. RELATÓRIO
    Trata-se de análise de reajuste de preços para o contrato em epígrafe, 
    utilizando o índice {adm['indice']} com data-base em {adm['data_base']}.
    
    2. ANÁLISE TÉCNICA
    Após processamento da planilha de coleta enviada pela fiscalização, 
    apurou-se um fator de reajuste de {adm['fator']:.4f}. 
    O impacto financeiro foi calculado com base no faturamento retroativo 
    e no saldo remanescente reajustado.
    
    - Valor Original do Contrato: R$ {b['orig']:,.2f}
    - Valor Global Estimado Pós-Reajuste: R$ {b['final']:,.2f}
    - Impacto Financeiro Total: R$ {diff:,.2f}
    - Variação Percentual: {perc:.2f}%
    
    3. CONCLUSÃO
    Considerando que a variação de {perc:.2f}% decorre exclusivamente da 
    aplicação de índice de preços previsto contratualmente, a alteração 
    caracteriza-se como reajustamento de preços em sentido estrito, 
    não computando para fins de limite de aditamentos previsto no 
    Art. 81 da Lei 13.303/2016. 
    
    Sugere-se o prosseguimento do feito para formalização via Apostilamento.
    """
    
    with st.expander("Visualizar Minuta da Nota Técnica"):
        st.text_area("Copie o texto abaixo para o seu SEI/Documento:", texto_nt, height=400)
        
    st.download_button(
        label="📄 Baixar Relatório em TXT",
        data=texto_nt,
        file_name=f"NT_Reajuste_{adm['ciclo_atual']}.txt",
        mime="text/plain"
    )
else:
    st.info("Processe uma planilha no Passo 3 para gerar o relatório.")