import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Facilitador - Cálculo de Represados", layout="wide")

# Estilização para destaque da Data-Base e Janela Admissível
st.markdown("""
    <style>
    .highlight-base { color: #003366; font-weight: bold; font-size: 18px; background-color: #f0f2f6; padding: 5px; border-radius: 5px; }
    .admissible-box { border-left: 5px solid #28a745; padding-left: 10px; margin-bottom: 10px; }
    </style>
    """, unsafe_allow_html=True)

def get_index_data(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=10)
        df = pd.DataFrame(response.json())
        if df.empty: return None
        df['valor'] = df['valor'].astype(float) / 100
        return (1 + df['valor']).prod() - 1
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo de Represados")

with st.sidebar:
    st.header("Configuração Geral")
    dt_base_original = st.date_input("Data-Base Original (Proposta):", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_anos = st.number_input("Quantidade de Ciclos:", min_value=1, max_value=10, value=2)
    indice_ref = st.selectbox("Índice:", ["IPCA (433)", "IGP-M (189)"])

resumo = []
fator_acumulado = 1.0
data_base_atual = dt_base_original

for i in range(1, qtd_anos + 1):
    with st.expander(f"Ciclo {i}", expanded=True):
        aniv_teorico = data_base_atual + relativedelta(years=1)
        limite_admissibilidade = aniv_teorico + relativedelta(days=90)
        
        st.markdown(f'<div class="admissible-box"><b>Janela Admissível:</b> {aniv_teorico.strftime("%d/%m/%Y")} até {limite_admissibilidade.strftime("%d/%m/%Y")}</div>', unsafe_allow_html=True)
        
        col_inf, col_ped = st.columns(2)
        with col_inf:
            st.markdown(f'Data-Base deste Ciclo: <span class="highlight-base">{data_base_atual.strftime("%d/%m/%Y")}</span>', unsafe_allow_html=True)
            st.write(f"Aniversário: {aniv_teorico.strftime('%d/%m/%Y')}")
        
        with col_ped:
            dt_pedido = st.date_input(f"Data do Pedido - Ciclo {i}:", value=aniv_teorico, key=f"ped_{i}", format="DD/MM/YYYY")
        
        cod_serie = "433" if "IPCA" in indice_ref else "189"
        var_ciclo = get_index_data(cod_serie, data_base_atual.strftime('%d/%m/%Y'), aniv_teorico.strftime('%d/%m/%Y'))
        
        if var_ciclo is not None:
            fator_ciclo = 1 + var_ciclo
            fator_acumulado *= fator_ciclo
            
            # Lógica de Status: PRECLUSO ou No Prazo
            if dt_pedido > limite_admissibilidade:
                status = "❌ PRECLUSO (Arrasta Base)"
                fundamento = "Cláusula 8ª, §4º: Solicitação após 90 dias consome a anuidade e desloca a data-base."
            else:
                status = "✅ No Prazo"
                fundamento = "Cláusula 8ª, §1º: Reajuste anual em conformidade com o aniversário contratual."
            
            resumo.append({
                "Ciclo": i,
                "Base Anterior": data_base_atual.strftime('%d/%m/%Y'),
                "Variação": f"{var_ciclo*100:,.2f}%".replace('.', ','),
                "Pedido (Efeito)": dt_pedido.strftime('%d/%m/%Y'),
                "Status": status,
                "Fundamento": fundamento
            })
            data_base_atual = dt_pedido
        else:
            st.warning(f"Dados não encontrados para o Ciclo {i}")

if resumo:
    st.divider()
    perc_total = (fator_acumulado - 1) * 100
    
    c1, c2 = st.columns(2)
    c1.metric("Variação Total Acumulada", f"{perc_total:,.2f}%".replace('.', ','))
    c2.metric("Multiplicador Final", f"{fator_acumulado:.6f}")

    st.subheader("Relatório Consolidado de Análise Técnica")
    df_resumo = pd.DataFrame(resumo)
    st.table(df_resumo[["Ciclo", "Base Anterior", "Variação", "Pedido (Efeito)", "Status"]])
    
    st.info("**Fundamentação Legal (Cláusula Oitava):**")
    for r in resumo:
        st.write(f"**Ciclo {r['Ciclo']}:** {r['Fundamento']}")