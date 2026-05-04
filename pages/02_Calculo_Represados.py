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
            df['data'] = pd.to_datetime(df['data'], dayfirst=True, errors='coerce').dt.normalize()
            df = df.dropna(subset=['data'])

            r_ini = pd.to_datetime((d_ini - relativedelta(months=1)).replace(day=1)).normalize()
            r_fim = pd.to_datetime(d_fim.replace(day=1)).normalize()

            df_detalhado = df[(df['data'] >= r_ini) & (df['data'] <= r_fim)].copy()
            v_ini_rows = df[df['data'] == r_ini]
            v_fim_rows = df[df['data'] == r_fim]

            if v_ini_rows.empty or v_fim_rows.empty:
                return None

            v_ini = v_ini_rows['indice'].values[0]
            v_fim = v_fim_rows['indice'].values[0]

            return {
                'var': (v_fim / v_ini) - 1,
                'i_ini': v_ini,
                'i_fim': v_fim,
                'p_ini': r_ini,
                'p_fim': r_fim,
                'metodo': "Divisão de Número-Índice (IST)",
                'dados': df_detalhado[['data', 'indice']]
            }
        else:
            url = (
                f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie}/dados"
                f"?formato=json&dataInicial={d_ini.strftime('%d/%m/%Y')}"
                f"&dataFinal={d_fim.strftime('%d/%m/%Y')}"
            )
            response = requests.get(url, timeout=10)
            df_t = pd.DataFrame(response.json())

            if df_t.empty:
                return None

            df_t['v'] = df_t['valor'].astype(float) / 100
            df_t['data'] = pd.to_datetime(df_t['data'], dayfirst=True)

            return {
                'var': (1 + df_t['v']).prod() - 1,
                'metodo': "Produtório de taxas mensais (SGS/BCB)",
                'p_ini': d_ini,
                'p_fim': d_fim,
                'dados': df_t[['data', 'valor']]
            }
    except Exception:
        return None


st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Reajustes Múltiplos")

with st.sidebar:
    dt_base_original = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_ciclos = st.number_input("Ciclos:", min_value=1, max_value=5, value=2)
    idx_sel = st.selectbox("Índice:", ["IST (Série Local)", "IPCA (433)", "IGP-M (189)"])

# data_atual representa a Data-Base do ciclo em análise.
# Para C1, é a Data-Base Original.
# Para ciclos seguintes, será atualizada conforme a regra contratual:
# - ciclo concedido/tempestivo: próxima Data-Base = data do pedido anterior;
# - pedido no mesmo mês, antes do aniversário: próxima Data-Base = aniversário contratual;
# - pedido adiantado/precluso: próxima Data-Base = aniversário contratual.
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

    # Lógica de Admissibilidade
    if dt_ped < d_aniv:
        if dt_ped.year == d_aniv.year and dt_ped.month == d_aniv.month:
            sit_emoji = "🟡 ADMISSÍVEL - RESSALVA"
            situacao_limpa = "ADMISSÍVEL - RESSALVA"
        else:
            sit_emoji = "⚠️ ADIANTADO"
            situacao_limpa = "ADIANTADO"
    elif dt_ped <= d_lim:
        sit_emoji = "✅ TEMPESTIVO"
        situacao_limpa = "TEMPESTIVO"
    else:
        sit_emoji = "❌ PRECLUSO"
        situacao_limpa = "PRECLUSO"

    # Regra de ancoragem do próximo ciclo.
    # Ponto central do ajuste:
    # Se o C1 foi tempestivo com pedido em 21/10/2023, o C2 passa a ter Data-Base em 21/10/2023.
    # Logo, o pedido do C2 somente será plenamente tempestivo a partir de 21/10/2024.
    if situacao_limpa == "TEMPESTIVO":
        data_base_proximo_ciclo = dt_ped
    else:
        data_base_proximo_ciclo = d_aniv

    res_c = get_data_rep("433" if "IPCA" in idx_sel else "189", data_atual, d_fim, "IST" in idx_sel)

    # Intervalo exibido independentemente de haver dados disponíveis para o índice.
    if res_c:
        janela_ciclo = f"{res_c['p_ini'].strftime('%m/%Y')} a {res_c['p_fim'].strftime('%m/%Y')}"
    else:
        if "IST" in idx_sel:
            p_ini_preview = (data_atual - relativedelta(months=1)).replace(day=1)
            p_fim_preview = d_fim.replace(day=1)
        else:
            p_ini_preview = data_atual
            p_fim_preview = d_fim
        janela_ciclo = f"{p_ini_preview.strftime('%m/%Y')} a {p_fim_preview.strftime('%m/%Y')}"

    janela_adm = f"{d_aniv.strftime('%d/%m/%Y')} a {d_lim.strftime('%d/%m/%Y')}"

    st.markdown(f"""
    **Dados do Ciclo {i}:**
    - Intervalo do C{i}: {janela_ciclo}
    - Janela de Admissibilidade: {janela_adm}
    - Situação: {sit_emoji}
    """)

    if res_c:
        fator_acum *= (1 + res_c['var'])
        v_fmt = f"{res_c['var'] * 100:,.2f}%".replace('.', ',')
        v_acum_parcial = f"{(fator_acum - 1) * 100:,.2f}%".replace('.', ',')

        st.markdown(f"- Variação do Ciclo: **{v_fmt}**")

        with st.expander(f"🔍 Memória de Cálculo Detalhada - Ciclo {i}"):
            st.write(f"**Metodologia:** {res_c['metodo']}")
            st.dataframe(res_c['dados'], use_container_width=True)

        historico.append({
            "Ciclo": i,
            "Variação": v_fmt,
            "Acumulada": v_acum_parcial,
            "Situação": sit_emoji,
            "Pedido": dt_ped.strftime('%d/%m/%Y'),
            "Janela": janela_ciclo,
            "JanelaAdm": janela_adm
        })
    else:
        st.warning(
            "Não há dados disponíveis para o índice selecionado no intervalo de apuração deste ciclo. "
            "O ciclo foi exibido para controle, mas não foi incluído no cálculo acumulado."
        )

    # A progressão do ciclo deve ocorrer sempre, mesmo quando não houver índice disponível.
    data_atual = data_base_proximo_ciclo

if historico:
    st.divider()
    res_final = f"{(fator_acum - 1) * 100:,.2f}%".replace('.', ',')
    st.metric("Variação Acumulada Final", res_final)

    st.subheader("Relatório de Apuração")
    corpo_relatorio = ""
    for h in historico:
        corpo_relatorio += f"""
        **C{h['Ciclo']}:** Pedido em {h['Pedido']}. Intervalo do C{h['Ciclo']}: {h['Janela']}.  
        Janela de Admissibilidade: {h['JanelaAdm']}.  
        Resultado: {h['Situação']}. Variação: {h['Variação']}.  
        Índice {idx_sel}.
        \n\n"""
    st.info(corpo_relatorio)
