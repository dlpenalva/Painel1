import streamlit as st
import pandas as pd
from bcb import sgs
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

# Configuração da página
st.set_page_config(page_title="FARC - Painel de Admissibilidade", layout="wide")

# --- FUNÇÕES DE LÓGICA JURÍDICA ---

def calcular_admissibilidade(d_proposta, d_pedido, d_ultimo_reajuste=None):
    # Marco inicial: Se houve reajuste anterior, conta dele. Se não, da proposta.
    marco_inicial = d_ultimo_reajuste if d_ultimo_reajuste else d_proposta
    
    # Aniversário do Ciclo (12 meses após o marco)
    data_aniversario = marco_inicial + relativedelta(months=12)
    
    # Prazo de Preclusão (90 dias após o aniversário - Parágrafo Quinto)
    data_limite_preclusao = data_aniversario + relativedelta(days=90)
    
    # Verificação de status
    if d_pedido < data_aniversario:
        return "Antecipado", data_aniversario, "Pedido realizado antes do interstício de 12 meses."
    elif d_pedido <= data_limite_preclusao:
        return "Tempestivo", data_aniversario, "Pedido dentro do prazo de 90 dias."
    else:
        return "Precluso", data_aniversario, f"Pedido após {data_limite_preclusao.strftime('%d/%m/%Y')} (Parágrafo Quinto)."

# --- INTERFACE ---

st.title("⚖️ Painel de Admissibilidade e Cálculo")
st.caption("Gestão de Contratos (GCC) - Telebras | Conformidade com Cláusula Oitava")

# Abas do Painel
tab_adm, tab_calc, tab_relat = st.tabs(["📋 Admissibilidade", "🧮 Cálculo por Itens", "📄 Relatório Final"])

with tab_adm:
    st.subheader("Análise de Tempestividade e Interstício")
    
    c1, c2 = st.columns(2)
    with c1:
        data_base = st.date_input("Data da Proposta ou Última Data-Base:", value=date(2023, 1, 1))
        data_solicitacao = st.date_input("Data do Pedido (Protocolo):", value=date.today())
    
    with c2:
        ultimo_reajuste = st.date_input("Efeito Financeiro do Último Reajuste (se houver):", value=None, help="Deixe em branco se for o primeiro reajuste.")
        data_vigencia = st.date_input("Fim da Vigência Atual:", value=date(2025, 1, 1))

    # Execução da Lógica
    status, aniversario, motivo = calcular_admissibilidade(data_base, data_solicitacao, ultimo_reajuste)
    
    st.divider()
    
    # Exibição do Parecer
    if status == "Tempestivo":
        st.success(f"**PARECER: ADMISSÍVEL**")
        st.write(f"✅ **Interstício:** Respeitado (Aniversário em {aniversario.strftime('%d/%m/%Y')})")
        st.write(f"✅ **Tempestividade:** Pedido dentro do prazo de 90 dias.")
        st.info(f"**Efeito Financeiro:** A partir de {data_solicitacao.strftime('%d/%m/%Y')} (Conforme Parágrafo Segundo).")
    
    elif status == "Precluso":
        st.error(f"**PARECER: PRECLUSO**")
        st.write(f"❌ **Motivo:** {motivo}")
        st.warning("⚠️ Conforme Parágrafo Sétimo, a contratada só poderá solicitar novo reajuste após 12 meses deste ciclo.")
    
    else:
        st.warning(f"**PARECER: AGUARDANDO**")
        st.write(f"🕒 **Situação:** {motivo}")

with tab_calc:
    st.subheader("Liquidação sobre Itens Remanescentes")
    st.info("O cálculo deve incidir sobre o saldo de itens não consumidos.")
    
    # Exemplo de entrada de valor (será substituído por upload de planilha)
    valor_itens = st.number_input("Valor total dos itens remanescentes (R$):", min_value=0.0)
    escolha_indice = st.selectbox("Índice de Reajuste:", ["IST (Anatel)", "IPCA", "IGP-M"])
    
    if st.button("Calcular Impacto"):
        # Aqui chamaremos as funções de índice que já criamos
        st.write("Cálculo processado com base no acumulado de 12 meses...")

with tab_relat:
    st.subheader("Minuta de Apostilamento")
    if status == "Tempestivo":
        texto_relatorio = f"""
        RELATÓRIO DE ADMISSIBILIDADE
        
        Trata-se de análise de pedido de reajuste referente ao contrato com data-base em {data_base.strftime('%d/%m/%Y')}.
        
        1. DA ADMISSIBILIDADE:
        Verifica-se que o pedido foi protocolado em {data_solicitacao.strftime('%d/%m/%Y')}, respeitando o interstício 
        mínimo de 12 meses previsto no Parágrafo Primeiro da Cláusula Oitava.
        
        2. DA PRECLUSÃO:
        O aniversário do ciclo ocorreu em {aniversario.strftime('%d/%m/%Y')}. O pedido foi realizado dentro do prazo 
        de 90 dias estabelecido no Parágrafo Quinto, não havendo que se falar em preclusão.
        
        3. CONCLUSÃO:
        O reajuste é admissível, com efeitos financeiros a partir da data da solicitação.
        """
        st.text_area("Cópia para o Processo:", value=texto_relatorio, height=300)
    else:
        st.write("Relatório indisponível: O pedido não atende aos critérios de admissibilidade.")