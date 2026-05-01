import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo de Represados", layout="wide")

def get_data_rep(serie, d_ini, d_fim, is_ist):
    try:
        if is_ist:
            df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
            df.columns = [str(col).strip().lower() for col in df.columns]
            df['data'] = pd.to_datetime(df['data'], dayfirst=True)
            r_ini = (d_ini - relativedelta(months=1)).replace(day=1)
            r_fim = d_fim.replace(day=1)
            v_ini = df[df['data'].dt.to_period('M') == r_ini.strftime('%Y-%m')]['indice'].values[0]
            v_fim = df[df['data'].dt.to_period('M') == r_fim.strftime('%Y-%m')]['indice'].values[0]
            return {'var': (v_fim/v_ini)-1, 'i_ini': v_ini, 'i_fim': v_fim, 'p_ini': r_ini, 'p_fim': r_fim, 'metodo': "Divisão de Número-Índice"}
        else:
            url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie}/dados?formato=json&dataInicial={d_ini.strftime('%d/%m/%Y')}&dataFinal={d_fim.strftime('%d/%m/%Y')}"
            r = requests.get(url, timeout=10).json()
            df_t = pd.DataFrame(r)
            df_t['v'] = df_t['valor'].astype(float) / 100
            return {'var': (1 + df_t['v']).prod() - 1, 'metodo': "Produtório de taxas mensais", 'p_ini': d_ini, 'p_fim': d_fim}
    except: return None

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Cálculo de Represados")

with st.sidebar:
    dt_base_original = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_ciclos = st.number_input("Ciclos:", min_value=1, max_value=5, value=2)
    idx_sel = st.selectbox("Índice:", ["IPCA (433)", "IGP-M (189)", "IST (Série Local)"])

data_atual = dt_base_original
fator_acum = 1.0
historico = []

for i in range(1, int(qtd_ciclos) + 1):
    st.markdown(f"### Ciclo {i}")
    d_fim = data_atual + relativedelta(months=11)
    d_aniv = data_atual + relativedelta(years=1)
    d_lim = d_aniv + relativedelta(days=90)
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.write(f"**Data-Base do Ciclo:** {data_atual.strftime('%d/%m/%Y')}")
    with col_b:
        dt_ped = st.date_input(f"Data do Pedido C{i}:", value=d_aniv, key=f"p{i}", format="DD/MM/YYYY")

    res_c = get_data_rep("433" if "IPCA" in idx_sel else "189", data_atual, d_fim, "IST" in idx_sel)
    
    if res_c:
        fator_acum *= (1 + res_c['var'])
        v_fmt = f"{res_c['var']*100:,.2f}%".replace('.', ',')
        sit = "✅ TEMPESTIVO" if dt_ped <= d_lim else "❌ PRECLUSO"
        
        # Bloco de Resumo de Rastreabilidade
        st.markdown(f"""
        **Dados do Ciclo {i}:**
        - Janela de apuração: {res_c['p_ini'].strftime('%m/%Y')} a {res_c['p_fim'].strftime('%m/%Y')}
        - Situação: {sit}
        - Metodologia: {res_c['metodo']}
        - Variação do Ciclo: **{v_fmt}**
        """)
        
        with st.expander(f"🔍 Memória de Cálculo Detalhada - Ciclo {i}"):
            if "IST" in idx_sel:
                st.write(f"Cálculo: ({res_c['i_fim']} / {res_c['i_ini']}) - 1")
            else:
                st.write("Baseado em série histórica do Banco Central.")
            st.write(f"Resultado bruto: {res_c['var']}")

        historico.append({"Ciclo": i, "Base": data_atual.strftime('%d/%m/%Y'), "Variação": v_fmt, "Situação": sit})
        # Lógica de arrasto
        data_atual = dt_ped if dt_ped > d_lim else d_aniv
    else:
        st.error(f"Erro no Ciclo {i}")

if historico:
    st.divider()
    res_final = f"{(fator_acum - 1)*100:,.2f}%".replace('.', ',')
    st.metric("Variação Acumulada Final (Represado)", res_final)
    st.table(pd.DataFrame(historico))
    
    # Reintrodução dos Relatórios
    relatorio_md = f"### Relatório de Reajuste\n\n**Índice:** {idx_sel}\n**Variação Total:** {res_final}\n\n"
    for h in historico:
        relatorio_md += f"- Ciclo {h['Ciclo']}: Base {h['Base']} | Var: {h['Variação']} | {h['Situação']}\n"
    
    st.download_button("Baixar Relatório (TXT)", relatorio_md, file_name="relatorio_reajuste.txt")