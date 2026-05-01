import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo de Represados", layout="wide")

def get_index_data_detailed(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=10)
        df = pd.DataFrame(response.json())
        if df.empty: return None
        df['valor'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        return df
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo de Represados")

with st.sidebar:
    st.header("Configuração")
    dt_base_original = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_anos = st.number_input("Quantidade de Ciclos:", min_value=1, max_value=10, value=2)
    indice_ref = st.selectbox("Índice:", ["IPCA (433)", "IGP-M (189)"])

resumo_final = []
fator_acumulado = 1.0
data_base_atual = dt_base_original

for i in range(1, int(qtd_anos) + 1):
    with st.container():
        # Lógica de 12 meses (Mês 0 + 11)
        fim_periodo_calculo = data_base_atual + relativedelta(months=11)
        aniv_teorico = data_base_atual + relativedelta(years=1)
        limite_90 = aniv_teorico + relativedelta(days=90)
        
        st.subheader(f"Ciclo {i}")
        col_inf, col_ped = st.columns(2)
        
        with col_inf:
            st.markdown(f"**Data-Base Atual:** `{data_base_atual.strftime('%d/%m/%Y')}`")
            st.markdown(f"**Período de Índice:** `{data_base_atual.strftime('%m/%Y')}` a `{fim_periodo_calculo.strftime('%m/%Y')}`")
        
        with col_ped:
            dt_pedido = st.date_input(f"Data do Pedido - Ciclo {i}:", value=aniv_teorico, key=f"ped_{i}", format="DD/MM/YYYY")
        
        cod_serie = "433" if "IPCA" in indice_ref else "189"
        df_ciclo = get_index_data_detailed(cod_serie, data_base_atual.strftime('%d/%m/%Y'), fim_periodo_calculo.strftime('%d/%m/%Y'))
        
        if df_ciclo is not None:
            var_ciclo = (1 + df_ciclo['valor']).prod() - 1
            fator_ciclo = 1 + var_ciclo
            fator_acumulado *= fator_ciclo
            status = "✅ No Prazo" if dt_pedido <= limite_90 else "❌ PRECLUSO (Arrasta Base)"
            
            # MEMÓRIA DE CÁLCULO DETALHADA POR CICLO
            with st.expander(f"🔍 Auditoria Mensal - Ciclo {i} (Fator de Prova)"):
                df_view = df_ciclo.copy()
                df_view['Variação Mensal'] = df_view['valor'].map(lambda x: f"{x*100:.4f}%")
                st.write(f"**Índice Acumulado no Ciclo:** {var_ciclo*100:,.4f}%".replace('.', ','))
                st.table(df_view[['data', 'Variação Mensal']].rename(columns={'data': 'Mês/Ano'}))
            
            resumo_final.append({
                "Ciclo": i,
                "Referência": f"{data_base_atual.strftime('%m/%y')} - {fim_periodo_calculo.strftime('%m/%y')}",
                "Variação": f"{var_ciclo*100:,.2f}%".replace('.', ','),
                "Data do Pedido": dt_pedido.strftime('%d/%m/%Y'),
                "Status": status
            })
            data_base_atual = dt_pedido # Regra de arrasto de base para o próximo ciclo
        else:
            st.error(f"Erro ao obter dados para o Ciclo {i}. Verifique a conexão com o Banco Central.")
    st.markdown("---")

if resumo_final:
    perc_total = (fator_acumulado - 1) * 100
    st.header("Resultado Consolidado")
    
    col_res1, col_res2 = st.columns(2)
    col_res1.metric("Variação Total Acumulada", f"{perc_total:,.2f}%".replace('.', ','))
    col_res2.metric("Multiplicador Final", f"{fator_acumulado:.6f}")

    st.subheader("Memória de Cálculo Resumida")
    st.table(pd.DataFrame(resumo_final))