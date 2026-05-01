import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo de Represados", layout="wide")

def get_data_rep(serie, d_ini, d_fim, is_ist):
    try:
        if is_ist:
            # Leitura e normalização da base IST
            df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
            df.columns = [str(col).strip().lower() for col in df.columns]
            
            # 1. Padronização: Converter coluna para datetime e normalizar (zerar horas)
            df['data'] = pd.to_datetime(df['data'], dayfirst=True).dt.normalize()
            
            # 2. Padronização: Converter datas de controle (date) para pd.Timestamp normalizado
            r_ini = pd.to_datetime((d_ini - relativedelta(months=1)).replace(day=1)).normalize()
            r_fim = pd.to_datetime(d_fim.replace(day=1)).normalize()
            
            # 3. Filtros agora funcionam com tipos idênticos (Timestamp vs Timestamp)
            df_detalhado = df[(df['data'] >= r_ini) & (df['data'] <= r_fim)].copy()
            
            # Extração de valores para a fórmula do IST
            v_ini_rows = df[df['data'] == r_ini]
            v_fim_rows = df[df['data'] == r_fim]
            
            if v_ini_rows.empty or v_fim_rows.empty:
                st.error(f"Dados faltantes no IST: {r_ini.strftime('%m/%Y')} ou {r_fim.strftime('%m/%Y')}")
                return None
                
            v_ini = v_ini_rows['indice'].values[0]
            v_fim = v_fim_rows['indice'].values[0]
            
            return {
                'var': (v_fim/v_ini)-1, 'i_ini': v_ini, 'i_fim': v_fim, 
                'p_ini': r_ini, 'p_fim': r_fim, 'metodo': "Divisão de Número-Índice (IST)",
                'dados': df_detalhado[['data', 'indice']]
            }
        else:
            # IPCA e IGP-M via SGS/BCB permanecem inalterados
            url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie}/dados?formato=json&dataInicial={d_ini.strftime('%d/%m/%Y')}&dataFinal={d_fim.strftime('%d/%m/%Y')}"
            response = requests.get(url, timeout=10)
            df_t = pd.DataFrame(response.json())
            df_t['v'] = df_t['valor'].astype(float) / 100
            df_t['data'] = pd.to_datetime(df_t['data'], dayfirst=True)
            return {
                'var': (1 + df_t['v']).prod() - 1, 
                'metodo': "Produtório de taxas mensais (SGS/BCB)", 
                'p_ini': d_ini, 'p_fim': d_fim,
                'dados': df_t[['data', 'valor']]
            }
    except Exception as e:
        st.error(f"Erro técnico na coleta de dados: {str(e)}")
        return None

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
        v_acum_parcial = f"{(fator_acum - 1)*100:,.2f}%".replace('.', ',')
        sit_emoji = "✅ TEMPESTIVO" if dt_ped <= d_lim else "❌ PRECLUSO"
        sit_limpo = "TEMPESTIVO" if dt_ped <= d_lim else "PRECLUSO"
        
        st.markdown(f"""
        **Dados do Ciclo {i}:**
        - Janela de apuração: {res_c['p_ini'].strftime('%m/%Y')} a {res_c['p_fim'].strftime('%m/%Y')}
        - Situação: {sit_emoji}
        - Metodologia: {res_c['metodo']}
        - Variação do Ciclo: **{v_fmt}**
        """)
        
        with st.expander(f"🔍 Memória de Cálculo Detalhada - Ciclo {i}"):
            st.write(f"**Metodologia:** {res_c['metodo']}")
            st.write(f"**Janela de Apuração:** {res_c['p_ini'].strftime('%m/%Y')} a {res_c['p_fim'].strftime('%m/%Y')}")
            
            if "IST" in idx_sel:
                # Memória específica para o IST restaurada
                st.dataframe(res_c['dados'], use_container_width=True)
                st.write(f"**Data Inicial:** {res_c['p_ini'].strftime('%m/%Y')} | **Valor:** {res_c['i_ini']}")
                st.write(f"**Data Final:** {res_c['p_fim'].strftime('%m/%Y')} | **Valor:** {res_c['i_fim']}")
                st.code(f"Fórmula aplicada: ({res_c['i_fim']} / {res_c['i_ini']}) - 1")
            else:
                st.dataframe(res_c['dados'], use_container_width=True)
                st.write("Fórmula: Produtório de (1 + taxa_mensal)")
            
            st.write(f"**Resultado bruto:** {res_c['var']}")
            st.write(f"**Percentual apurado:** {res_c['var']*100:.4f}%")
            st.write(f"**Percentual utilizado:** {v_fmt}")

        historico.append({
            "Ciclo": i, "Base": data_atual.strftime('%d/%m/%Y'), 
            "Variação": v_fmt, "Acumulada": v_acum_parcial,
            "Situação": sit_limpo, "Pedido": dt_ped.strftime('%d/%m/%Y'),
            "Janela": f"{res_c['p_ini'].strftime('%m/%Y')} a {res_c['p_fim'].strftime('%m/%Y')}"
        })
        data_atual = dt_ped if dt_ped > d_lim else d_aniv
    else:
        st.warning(f"Não foi possível processar o Ciclo {i}. Verifique a base de dados.")

if historico:
    st.divider()
    res_final = f"{(fator_acum - 1)*100:,.2f}%".replace('.', ',')
    st.metric("Variação Acumulada Final (Represado)", res_final)
    
    df_hist = pd.DataFrame(historico)[["Ciclo", "Base", "Variação", "Situação"]]
    st.dataframe(df_hist, hide_index=True, use_container_width=True)
    
    st.subheader("Relatório de Apuração")
    corpo_relatorio = ""
    for h in historico:
        corpo_relatorio += f"""
        **C{h['Ciclo']}:** Pedido em {h['Pedido']}. Janela {h['Janela']}.  
        Resultado: {h['Situação']}. Variação: {h['Variação']}.  
        Índice {idx_sel}. Início dos efeitos financeiros: {h['Pedido']}.
        \n\n"""
    st.info(corpo_relatorio)