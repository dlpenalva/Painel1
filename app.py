with tab_rel:
    f = st.session_state.farc
    if f.get('status'):
        # Verificação se é precluso para mudar o corpo do relatório
        is_precluso = f['status'] == "Precluso"
        
        # Formatação do Status com cor no relatório (Markdown)
        status_formatado = f"**:red[{f['status'].upper()}]**" if is_precluso else f['status']
        
        if is_precluso:
            # Relatório enxuto para Preclusão
            texto_rel = f"""RELATÓRIO TÉCNICO DE ADMISSIBILIDADE CONTRATUAL

1. FUNDAMENTAÇÃO LEGAL E REFERÊNCIAS
- Amparo: Lei nº 13.303/2016 e Decreto nº 12.500/2025.
- Empresa: Telebras (Status: Não Dependente).
- Cláusula de Reajuste: Cláusula Sétima, Parágrafo Primeiro.
- Data-Base Anterior: {f['dt_base'].strftime('%d/%m/%Y')}
- Aniversário do Direito: {f['dt_aniv'].strftime('%d/%m/%Y')}

2. ANÁLISE DE ADMISSIBILIDADE
- Data do Pedido: {f['dt_pedido'].strftime('%d/%m/%Y')}
- Parecer: O pedido é considerado {f['status']}.

3. CONCLUSÃO
Considerando que o pedido de reajuste foi protocolado em {f['dt_pedido'].strftime('%d/%m/%Y')}, restou superado o prazo de 90 dias após o aniversário do direito ({f['dt_aniv'].strftime('%d/%m/%Y')}). Ante o exposto, opera-se a PRECLUSÃO do direito ao reajuste relativo a este ciclo, devido ao lapso temporal transcorrido, não sendo cabível a apuração de valores ou concessão do índice."""
        else:
            # Relatório padrão para casos Admissíveis
            linha_valor_ant = f"- Valor Atual: R$ {f['valor']:,.2f}" if f['valor'] > 0 else ""
            linha_valor_nov = f"- Novo Valor Reajustado: R$ {f['v_novo']:,.2f}" if f['valor'] > 0 else ""
            
            texto_rel = f"""RELATÓRIO TÉCNICO DE REAJUSTE CONTRATUAL

1. FUNDAMENTAÇÃO LEGAL E REFERÊNCIAS
- Amparo: Lei nº 13.303/2016 e Decreto nº 12.500/2025.
- Empresa: Telebras (Status: Não Dependente).
- Cláusula de Reajuste: Cláusula Sétima, Parágrafo Primeiro.
- Data-Base Anterior: {f['dt_base'].strftime('%d/%m/%Y')}
- Aniversário do Direito: {f['dt_aniv'].strftime('%d/%m/%Y')}

2. ANÁLISE DE ADMISSIBILIDADE
- Data do Pedido: {f['dt_pedido'].strftime('%d/%m/%Y')}
- Parecer: O pedido é considerado {f['status']}.

3. MEMÓRIA DE CÁLCULO
- Índice Aplicado: {f['idx']}
- Variação Acumulada (12 meses): {f['var']:.6%}
{linha_valor_ant}
{linha_valor_nov}

4. CONCLUSÃO
Considerando o cumprimento do interstício de 12 meses e a previsão contratual, a variação de {f['var']:.6%} está apta para aplicação, retroagindo seus efeitos financeiros a {f['dt_aniv'].strftime('%d/%m/%Y')}."""

        st.subheader("Informações para relatório")
        # O st.text_area não suporta cores, então mostramos um preview formatado acima para conferência
        if is_precluso:
            st.error(f"Parecer Final: {f['status']}")
        
        st.text_area("", texto_rel.replace('\n\n\n', '\n').strip(), height=450)