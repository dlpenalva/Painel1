import streamlit as st
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Facilitador - Passivos", layout="wide")

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("🧮 Cálculo de Passivos (Múltiplos Ciclos)")
st.markdown('<p style="color: #666; margin-top: -20px;">Foco em Percentuais Acumulados e Admissibilidade temporal</p>', unsafe_allow_html=True)

with st.expander("⚙️ Configuração dos Períodos Acumulados", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        dt_base_original = st.date_input("Data-Base Inicial (Ciclo 0):", value=datetime(2022, 5, 4), format="DD/MM/YYYY")
    with col2:
        qtd_ciclos = st.number_input("Qtd. de Ciclos Pendentes:", min_value=2, max_value=5, value=2)
    with col3:
        dt_pedido = st.date_input("Data do Pedido Atual:", format="DD/MM/YYYY")

st.info("💡 Informe abaixo os índices (em %) apurados para cada período de 12 meses.")

# Gerando a grade de entrada e cálculos
dados_ciclos = []
fator_acumulado = 1.0

for i in range(1, qtd_ciclos + 1):
    st.subheader(f"Ciclo {i}")
    c1, c2, c3 = st.columns([2, 2, 2])
    
    # Cálculo das Datas do Ciclo
    inicio_ciclo = dt_base_original + relativedelta(years=i-1)
    fim_ciclo = dt_base_original + relativedelta(years=i)
    
    with c1:
        st.write(f"**Período:** {inicio_ciclo.strftime('%d/%m/%Y')} a {fim_ciclo.strftime('%d/%m/%Y')}")
    
    with c2:
        # Input do percentual apurado (ex: 5.50)
        perc_apurado = st.number_input(f"Variação do Ciclo {i} (%):", key=f"perc_{i}", format="%.4f")
        fator_ciclo = 1 + (perc_apurado / 100)
        fator_acumulado *= fator_ciclo
    
    with c3:
        # Checagem de Admissibilidade (Preclusão apenas no ciclo atual, os anteriores estão parados na empresa)
        dias_atraso = (dt_pedido - fim_ciclo).days
        if i == qtd_ciclos: # Apenas o último ciclo sofre o rigor dos 90 dias do pedido
            status = "⚠️ Analisar Preclusão" if dias_atraso > 90 else "✅ Admissível"
        else:
            status = "📦 Passivo Interno (Aguardando)"
        st.write(f"**Status:** {status}")

    dados_ciclos.append({
        "Ciclo": i,
        "Início": inicio_ciclo.strftime('%d/%m/%Y'),
        "Fim (Aniversário)": fim_ciclo.strftime('%d/%m/%Y'),
        "Variação Ciclo": f"{perc_apurado:,.2f}%".replace('.', ','),
        "Efeito Financeiro": fim_ciclo.strftime('%d/%m/%Y')
    })

st.divider()

# Resultado Final do Acúmulo
total_perc_acumulado = (fator_acumulado - 1) * 100

res_col1, res_col2 = st.columns(2)
with res_col1:
    st.metric("Variação Acumulada Total (Cascata)", f"{total_perc_acumulado:,.4f}%".replace('.', ','))
    st.caption("Nota: Este é o percentual final a ser aplicado sobre o valor original do contrato.")

with res_col2:
    st.metric("Multiplicador Final", f"{fator_acumulado:.6f}")
    st.caption("Utilize este fator para multiplicar o valor unitário dos itens.")

st.subheader("📋 Quadro Resumo para Relatório")
df_resumo = pd.DataFrame(dados_ciclos)
st.table(df_resumo)

st.markdown(f"""
> **Parecer Sugerido:**  
> Considerando os {qtd_ciclos} ciclos de reajuste acumulados, o índice final resultante da aplicação em cascata é de **{total_perc_acumulado:,.4f}%**.  
> O efeito financeiro de cada parcela deve retroagir às respectivas datas de aniversário listadas acima, observando-se que o passivo decorre de trâmite interno desta Telebras.
""")