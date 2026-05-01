import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Facilitador - Cálculo de Represados", layout="wide")

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

st.subheader("Entrada de Dados por Ciclo")
st.info("A data do pedido de um ciclo define a data-base do ciclo seguinte.")

for i in range(1, qtd_anos + 1):
    with st.expander(f"Ciclo {i}", expanded=True):
        col_inf, col_ped = st.columns(2)
        
        # Aniversário teórico (12 meses após a base atual)
        aniv_teorico = data_base_atual + relativedelta(years=1)
        
        with col_inf:
            st.write(f"**Data-Base deste Ciclo:** {data_base_atual.strftime('%d/%m/%Y')}")
            st.write(f"**Aniversário do Contrato:** {aniv_teorico.strftime('%d/%m/%Y')}")
        
        with col_ped:
            dt_pedido = st.date_input(f"Data do Pedido (Efeito Financeiro) - Ciclo {i}:", 
                                      value=aniv_teorico, key=f"ped_{i}", format="DD/MM/YYYY")
        
        # Cálculo da variação (Sempre 12 meses do índice)
        cod_serie = "433" if "IPCA" in indice_ref else "189"
        # O intervalo de variação é do início da base até o aniversário (12 meses de inflação)
        var_ciclo = get_index_data(cod_serie, data_base_atual.strftime('%d/%m/%Y'), aniv_teorico.strftime('%d/%m/%Y'))
        
        if var_ciclo is not None:
            fator_ciclo = 1 + var_ciclo
            fator_acumulado *= fator_ciclo
            
            # Checa preclusão (90 dias após o aniversário)
            atraso = (dt_pedido - aniv_teorico).days
            status = "✅ No Prazo" if atraso <= 90 else "⚠️ Atraso (Arrasta Base)"
            
            resumo.append({
                "Ciclo": i,
                "Base Anterior": data_base_atual.strftime('%d/%m/%Y'),
                "Variação (%)": f"{var_ciclo*100:,.4f}%".replace('.', ','),
                "Efeito Financeiro": dt_pedido.strftime('%d/%m/%Y'),
                "Status": status
            })
            
            # A NOVA DATA BASE é a data do pedido (Efeito Financeiro)
            data_base_atual = dt_pedido
        else:
            st.warning(f"Aguardando dados para o Ciclo {i}...")

if resumo:
    st.divider()
    perc_total = (fator_acumulado - 1) * 100
    
    c1, c2 = st.columns(2)
    c1.metric("Variação Total Acumulada", f"{perc_total:,.4f}%".replace('.', ','))
    c2.metric("Multiplicador Final", f"{fator_acumulado:.6f}")

    st.subheader("Memória de Cálculo de Represados")
    st.table(pd.DataFrame(resumo))
    
    st.markdown(f"""
    > **Nota Técnica:** O cálculo observa a sucessão de datas-base. O efeito financeiro do Ciclo {qtd_anos} 
    > passa a ser a nova data-base para o próximo período de 12 meses, conforme o princípio da anualidade.
    """)