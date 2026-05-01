import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="GCC - Cálculo Individual", layout="wide")

# Estilo Telebras
st.markdown("""
    <style>
    .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; border-left: 5px solid #003366; }
    .stTabs [aria-selected="true"] { background-color: #003366 !important; color: white !important; }
    .subtitle-gcc { font-size: 14px; color: #666; margin-top: -20px; margin-bottom: 20px; }
    </style>
    """, unsafe_allow_html=True)

def get_index_data(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio}&dataFinal={data_fim}"
    try:
        response = requests.get(url, timeout=15)
        df = pd.DataFrame(response.json())
        if df.empty: return None, "Vazio"
        df['valor'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        if len(df) > 12: df = df.head(12)
        return df, None
    except:
        return None, "Erro API"

def calc_ist_csv(dt_base, dt_aniv):
    try:
        df = pd.read_csv('ist.csv', sep=None, engine='python', decimal=',')
        df.columns = df.columns.str.replace('^\ufeff', '', regex=True)
        meses_map = {1:'jan', 2:'fev', 3:'mar', 4:'abr', 5:'mai', 6:'jun', 7:'jul', 8:'ago', 9:'set', 10:'out', 11:'nov', 12:'dez'}
        ref_base = f"{meses_map[dt_base.month]}/{str(dt_base.year)[2:]}"
        ref_aniv = f"{meses_map[dt_aniv.month]}/{str(dt_aniv.year)[2:]}"
        v_base = float(df[df['MES_ANO'] == ref_base]['INDICE_NIVEL'].values[0])
        v_aniv = float(df[df['MES_ANO'] == ref_aniv]['INDICE_NIVEL'].values[0])
        var = (v_aniv / v_base) - 1
        memoria_df = pd.DataFrame([
            {"Referência": f"Inicial ({ref_base})", "Índice Nível": f"{v_base:.4f}"},
            {"Referência": f"Final ({ref_aniv})", "Índice Nível": f"{v_aniv:.4f}"}
        ])
        return var, ref_base, ref_aniv, memoria_df, None
    except:
        return None, None, None, None, "Erro IST"

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Gestão de Contratos - Reajustes")
st.markdown('<p class="subtitle-gcc">Cálculo Individual</p>', unsafe_allow_html=True)

if 'farc' not in st.session_state: st.session_state.farc = {}

tab_adm, tab_calc, tab_rel = st.tabs(["📊 Admissibilidade", "🧮 Cálculo Detalhado", "📄 Relatório Final"])

with tab_adm:
    col1, col2 = st.columns(2)
    with col1:
        dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 5, 1), format="DD/MM/YYYY")
        dt_solic = st.date_input("Data do Pedido:", format="DD/MM/YYYY")
    with col2:
        valor_base = st.number_input("Valor Atual (R$) - Opcional:", min_value=0.0, step=100.0)
        tipo_idx = st.selectbox("Índice:", ["IPCA (Série 433)", "IGP-M (Série 189)", "IST (Planilha CSV)"])

    dt_aniv = dt_base + relativedelta(years=1)
    dt_fim_calc = dt_aniv - relativedelta(months=1)
    dias_janela = (dt_solic - dt_aniv).days
    intersticio_ok = dt_solic >= dt_aniv
    mesmo_mes = (dt_solic.month == dt_aniv.month and dt_solic.year == dt_aniv.year)

    if dias_janela > 90: status = "Precluso"
    elif not intersticio_ok and mesmo_mes: status = "Admissível (Ajuste Prévio)"
    elif not intersticio_ok: status = "Antecipado"
    else: status = "Admissível"

    st.divider()
    c1, c2, c3 = st.columns(3)
    c1.metric("Janela Temporal", f"{max(0, dias_janela)} dias")
    if "Admissível" in status: c2.success(f"Status: {status}")
    elif status == "Precluso": c2.error(f"Status: {status}")
    else: c2.warning(f"Status: {status}")
    
    if intersticio_ok: c3.success("Interstício Legal: Ok")
    else: c3.warning("Interstício: Pendente")

    st.session_state.farc = {
        'dt_base': dt_base, 'dt_aniv': dt_aniv, 'dt_pedido': dt_solic,
        'dt_fim_calc': dt_fim_calc, 'valor': valor_base, 'idx': tipo_idx, 
        'status': status, 'var': 0.0, 'v_novo': 0.0
    }

with tab_calc:
    f = st.session_state.farc
    if "IST" in f['idx']:
        var, rb, ra, mem_df, erro = calc_ist_csv(f['dt_base'], f['dt_aniv'])
        if not erro:
            st.subheader("Memória de Cálculo - IST")
            st.metric(f"Variação ({rb} a {ra})", f"{var*100:,.2f}%".replace('.', ','))
            if f['valor'] > 0:
                v_novo = f['valor'] * (1 + var)
                st.metric("Novo Valor", f"R$ {v_novo:,.2f}")
                st.session_state.farc.update({'v_novo': v_novo})
            st.table(mem_df)
            st.session_state.farc.update({'var': var})
    else:
        cod = "433" if "IPCA" in f['idx'] else "189"
        df, erro = get_index_data(cod, f['dt_base'].strftime('%d/%m/%Y'), f['dt_fim_calc'].replace(day=28).strftime('%d/%m/%Y'))
        if df is not None:
            var = (1 + df['valor']).prod() - 1
            st.subheader(f"Memória de Cálculo - {f['idx']}")
            st.metric(f"Período: {df.iloc[0]['data'].strftime('%m/%Y')} a {df.iloc[-1]['data'].strftime('%m/%Y')}", f"{var*100:,.2f}%".replace('.', ','))
            if f['valor'] > 0:
                v_novo = f['valor'] * (1 + var)
                st.metric("Novo Valor", f"R$ {v_novo:,.2f}")
                st.session_state.farc.update({'v_novo': v_novo})
            df_display = df.copy()
            df_display['data'] = df_display['data'].dt.strftime('%m/%Y')
            df_display['valor (%)'] = (df_display['valor'] * 100).apply(lambda x: f"{x:,.2f}%".replace('.', ','))
            st.dataframe(df_display[['data', 'valor (%)']], use_container_width=True)
            st.session_state.farc.update({'var': var})

with tab_rel:
    f = st.session_state.farc
    if f.get('status'):
        is_precluso = f['status'] == "Precluso"
        var_formatada = f"{f['var']*100:,.2f}%".replace('.', ',')
        
        if is_precluso:
            texto_rel = f"""RELATÓRIO TÉCNICO DE ADMISSIBILIDADE CONTRATUAL

1. FUNDAMENTAÇÃO LEGAL E REFERÊNCIAS
- Amparo: Lei nº 13.303/2016 e Decreto nº 12.500/2025.
- Empresa: Telebras (Status: Não Dependente).
- Cláusula de Reajuste: Cláusula Sétima, Parágrafo Primeiro.
- Data-Base Anterior: {f['dt_base'].strftime('%d/%m/%Y')}
- Aniversário do Direito: {f['dt_aniv'].strftime('%d/%m/%Y')}

2. ANÁLISE DE ADMISSIBILIDADE
- Data do Pedido: {f['dt_pedido'].strftime('%d/%m/%Y')}
- Parecer: O pedido é considerado PRECLUSO.

3. CONCLUSÃO
Considerando que o pedido de reajuste foi protocolado em {f['dt_pedido'].strftime('%d/%m/%Y')}, restou superado o prazo de 90 dias após o aniversário do direito ({f['dt_aniv'].strftime('%d/%m/%Y')}). Ante o exposto, opera-se a PRECLUSÃO do direito ao reajuste relativo a este ciclo, devido ao lapso temporal transcorrido, não sendo cabível a apuração de valores ou concessão do índice."""
        else:
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
- Variação Acumulada (12 meses): {var_formatada}
{linha_valor_ant}
{linha_valor_nov}

4. CONCLUSÃO
Considerando o cumprimento do interstício de 12 meses e a previsão contratual, a variação de {var_formatada} está apta para aplicação, retroagindo seus efeitos financeiros a {f['dt_aniv'].strftime('%d/%m/%Y')}."""

        st.text_area("", texto_rel.replace('\n\n\n', '\n').strip(), height=450)
