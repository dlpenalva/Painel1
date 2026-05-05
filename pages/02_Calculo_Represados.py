import streamlit as st
import pandas as pd
import requests
import io
from datetime import datetime
from zoneinfo import ZoneInfo
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Cálculo de Represados", layout="wide")

def get_data_rep(serie, d_ini, d_fim, is_ist):
    try:
        if is_ist:
            # IST é número-índice: usa a competência da data-base do ciclo e a
            # mesma competência 12 meses depois. Ex.: 10/2022 => out/2022 a out/2023.
            # A leitura aceita tanto o CSV data;indice quanto o CSV MES_ANO;INDICE_NIVEL.
            df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
            df.columns = [str(col).strip().lower() for col in df.columns]

            def _to_numero(serie):
                if pd.api.types.is_numeric_dtype(serie):
                    return pd.to_numeric(serie, errors='coerce')
                return pd.to_numeric(
                    serie.astype(str).str.strip().str.replace('.', '', regex=False).str.replace(',', '.', regex=False),
                    errors='coerce'
                )

            if {'data', 'indice'}.issubset(df.columns):
                df['data'] = pd.to_datetime(df['data'], dayfirst=True, errors='coerce')
                df['indice'] = _to_numero(df['indice'])
            elif {'mes_ano', 'indice_nivel'}.issubset(df.columns):
                meses = {
                    'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
                    'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12,
                }

                def _parse_mes_ano(valor):
                    try:
                        mes_txt, ano_txt = str(valor).strip().lower().split('/')
                        mes = meses[mes_txt[:3]]
                        ano = int(ano_txt)
                        ano = 2000 + ano if ano < 100 else ano
                        return pd.Timestamp(year=ano, month=mes, day=1)
                    except Exception:
                        return pd.NaT

                df['data'] = df['mes_ano'].apply(_parse_mes_ano)
                df['indice'] = _to_numero(df['indice_nivel'])
            else:
                st.error("A base IST precisa conter as colunas 'data'/'indice' ou 'MES_ANO'/'INDICE_NIVEL'.")
                return None

            df = df.dropna(subset=['data', 'indice']).copy()

            r_ini = d_ini.replace(day=1)
            r_fim = (d_ini + relativedelta(years=1)).replace(day=1)

            periodo_ini = r_ini.strftime('%Y-%m')
            periodo_fim = r_fim.strftime('%Y-%m')

            v_ini_rows = df[df['data'].dt.to_period('M') == periodo_ini]
            v_fim_rows = df[df['data'].dt.to_period('M') == periodo_fim]

            if v_ini_rows.empty or v_fim_rows.empty:
                st.error(f"Dados do IST não encontrados para o período {r_ini.strftime('%m/%Y')} ou {r_fim.strftime('%m/%Y')}.")
                return None

            v_ini = float(v_ini_rows['indice'].iloc[0])
            v_fim = float(v_fim_rows['indice'].iloc[0])

            return {
                'var': (v_fim / v_ini) - 1,
                'i_ini': v_ini,
                'i_fim': v_fim,
                'p_ini': r_ini,
                'p_fim': r_fim,
                'metodo': "Divisão de Número-Índice (IST)",
            }
        else:
            # Manutenção da lógica do IPCA/IGP-M via SGS/BCB.
            url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie}/dados?formato=json&dataInicial={d_ini.strftime('%d/%m/%Y')}&dataFinal={d_fim.strftime('%d/%m/%Y')}"
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
    except Exception as e:
        st.error(f"Erro técnico na coleta de dados: {str(e)}")
        return None

def _formatar_data(valor):
    """Formata datas para DD/MM/AAAA sem quebrar quando o valor estiver vazio."""
    if valor is None or valor == "":
        return ""
    try:
        return pd.to_datetime(valor, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return str(valor)


def _data_para_datetime(valor):
    if valor is None or valor == "":
        return None
    try:
        return pd.to_datetime(valor, dayfirst=True).to_pydatetime()
    except Exception:
        return None


def _competencias_mensais(data_inicio, data_fim):
    """Gera competências mensais inclusivas no formato MM/AAAA."""
    inicio = _data_para_datetime(data_inicio)
    fim = _data_para_datetime(data_fim)
    if inicio is None or fim is None:
        return []
    atual = inicio.replace(day=1)
    limite = fim.replace(day=1)
    competencias = []
    while atual <= limite:
        competencias.append(atual.strftime("%m/%Y"))
        atual = atual + relativedelta(months=1)
    return competencias


def _percentual_formatado(valor):
    try:
        return f"{float(valor) * 100:,.2f}%".replace('.', ',')
    except Exception:
        return ""


def gerar_arquivo_coleta_excel(dados_admissibilidade):
    """Gera o Arquivo de Coleta para a fase de Valor Global.

    Mantém a lógica de coleta separada da apuração: FINANCEIRO_MENSAL para valores pagos
    e ITENS_REMANESCENTES para o saldo físico/contratual no início de cada ciclo.
    """
    output = io.BytesIO()
    ciclos = dados_admissibilidade.get('ciclos', [])
    data_geracao = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        fmt_header = workbook.add_format({
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F4E79",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        fmt_subheader = workbook.add_format({
            "bold": True,
            "bg_color": "#D9EAD3",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
        })
        fmt_input = workbook.add_format({"bg_color": "#FFF2CC", "border": 1})
        fmt_money = workbook.add_format({"num_format": 'R$ #,##0.00', "border": 1})
        fmt_money_input = workbook.add_format({"num_format": 'R$ #,##0.00', "bg_color": "#FFF2CC", "border": 1})
        fmt_number_input = workbook.add_format({"num_format": '#,##0.00', "bg_color": "#FFF2CC", "border": 1})
        fmt_text = workbook.add_format({"border": 1})
        fmt_total = workbook.add_format({"bold": True, "bg_color": "#E2F0D9", "border": 1})
        fmt_percent = workbook.add_format({"num_format": "0.00%", "border": 1})
        fmt_factor = workbook.add_format({"num_format": "0.0000", "border": 1})

        # Aba PARAMETROS_REAJUSTE
        parametros = pd.DataFrame([
            ["Origem da análise", dados_admissibilidade.get('origem', dados_admissibilidade.get('tipo', ''))],
            ["Índice utilizado", dados_admissibilidade.get('indice', '')],
            ["Data-base original", dados_admissibilidade.get('data_base_original', '')],
            ["Quantidade de ciclos", len(ciclos)],
            ["Fator acumulado final", dados_admissibilidade.get('fator_acumulado', dados_admissibilidade.get('fator', 1.0))],
            ["Variação acumulada final", dados_admissibilidade.get('variacao_acumulada', 0.0)],
            ["Data de geração do arquivo", data_geracao],
        ], columns=["Campo", "Valor"])
        parametros.to_excel(writer, sheet_name="PARAMETROS_REAJUSTE", index=False)
        ws = writer.sheets["PARAMETROS_REAJUSTE"]
        ws.set_column("A:A", 32)
        ws.set_column("B:B", 35)
        for col, title in enumerate(parametros.columns):
            ws.write(0, col, title, fmt_header)
        ws.write_number(5, 1, float(dados_admissibilidade.get('variacao_acumulada', 0.0)), fmt_percent)
        try:
            ws.write_number(4, 1, float(dados_admissibilidade.get('fator_acumulado', dados_admissibilidade.get('fator', 1.0))), fmt_factor)
        except Exception:
            pass

        # Aba CICLOS
        ciclos_rows = []
        for ciclo in ciclos:
            situacao = str(ciclo.get('situacao', ''))
            tratamento = "Precluso" if "PRECLUSO" in situacao.upper() else "A apurar"
            ciclos_rows.append({
                "Ciclo": ciclo.get('ciclo', ''),
                "Data-base": ciclo.get('data_base', ''),
                "Intervalo do índice": ciclo.get('intervalo_indice', ciclo.get('Janela', '')),
                "Janela de admissibilidade": ciclo.get('janela_admissibilidade', ciclo.get('JanelaAdm', '')),
                "Data do pedido": ciclo.get('data_pedido', ciclo.get('Pedido', '')),
                "Início financeiro": ciclo.get('financeiro_inicio', ''),
                "Fim financeiro": ciclo.get('financeiro_fim', ''),
                "Situação": situacao,
                "Variação": ciclo.get('variacao', 0.0),
                "Fator": ciclo.get('fator', 1.0),
                "Fator acumulado": ciclo.get('fator_acumulado', 1.0),
                "Tratamento financeiro do ciclo": tratamento,
            })
        df_ciclos = pd.DataFrame(ciclos_rows)
        df_ciclos.to_excel(writer, sheet_name="CICLOS", index=False)
        ws = writer.sheets["CICLOS"]
        ws.set_column("A:A", 12)
        ws.set_column("B:H", 24)
        ws.set_column("I:I", 12)
        ws.set_column("J:K", 14)
        ws.set_column("L:L", 30)
        for col, title in enumerate(df_ciclos.columns):
            ws.write(0, col, title, fmt_header)
        for row in range(1, len(df_ciclos) + 1):
            ws.write(row, 11, df_ciclos.iloc[row-1]["Tratamento financeiro do ciclo"], fmt_input)
            ws.write(row, 8, df_ciclos.iloc[row-1]["Variação"], fmt_percent)
            ws.write(row, 9, df_ciclos.iloc[row-1]["Fator"], fmt_factor)
            ws.write(row, 10, df_ciclos.iloc[row-1]["Fator acumulado"], fmt_factor)

        # Aba FINANCEIRO_MENSAL
        financeiro_rows = []
        for ciclo in ciclos:
            ciclo_nome = ciclo.get('ciclo', '')
            for competencia in _competencias_mensais(ciclo.get('financeiro_inicio', ''), ciclo.get('financeiro_fim', '')):
                financeiro_rows.append({
                    "Ciclo": ciclo_nome,
                    "Competência": competencia,
                    "Valor pago/faturado": "",
                })
        if not financeiro_rows:
            financeiro_rows.append({"Ciclo": "", "Competência": "", "Valor pago/faturado": ""})
        df_fin = pd.DataFrame(financeiro_rows)
        df_fin.to_excel(writer, sheet_name="FINANCEIRO_MENSAL", index=False)
        ws = writer.sheets["FINANCEIRO_MENSAL"]
        ws.set_column("A:A", 12)
        ws.set_column("B:B", 18)
        ws.set_column("C:C", 22)
        for col, title in enumerate(df_fin.columns):
            ws.write(0, col, title, fmt_header)
        for row in range(1, len(df_fin) + 1):
            ws.write(row, 2, "", fmt_money_input)
        total_row = len(df_fin) + 1
        ws.write(total_row, 0, "TOTAL", fmt_total)
        ws.write(total_row, 1, "", fmt_total)
        ws.write_formula(total_row, 2, f"=SUM(C2:C{len(df_fin)+1})", fmt_money)

        # Aba ITENS_REMANESCENTES
        wb_sheet = workbook.add_worksheet("ITENS_REMANESCENTES")
        writer.sheets["ITENS_REMANESCENTES"] = wb_sheet
        rem_cols = []
        for ciclo in ciclos:
            nome = ciclo.get('ciclo', '')
            data_ref = ciclo.get('data_base', '')
            rem_cols.append((f"Remanescente início {nome}", data_ref))
        base_headers = ["Item", "Quantidade contratada", "Valor unitário original", "Valor total"]
        headers = base_headers + [c[0] for c in rem_cols]
        # Linha 1: datas de referência dos remanescentes
        wb_sheet.merge_range(0, 0, 0, 3, "Dados do item e valor original", fmt_subheader)
        for idx, (_, data_ref) in enumerate(rem_cols, start=4):
            wb_sheet.write(0, idx, data_ref, fmt_subheader)
        # Linha 2: cabeçalhos
        for col, title in enumerate(headers):
            wb_sheet.write(1, col, title, fmt_header)
        wb_sheet.set_column(0, 0, 12)
        wb_sheet.set_column(1, 1, 24)
        wb_sheet.set_column(2, 3, 22)
        if rem_cols:
            wb_sheet.set_column(4, 4 + len(rem_cols) - 1, 24)
        start_row = 2
        end_row = 101
        for row in range(start_row, end_row + 1):
            wb_sheet.write(row, 0, "", fmt_text)
            wb_sheet.write(row, 1, "", fmt_number_input)
            wb_sheet.write(row, 2, "", fmt_money_input)
            wb_sheet.write_formula(row, 3, f"=IF(OR(B{row+1}=\"\",C{row+1}=\"\"),\"\",B{row+1}*C{row+1})", fmt_money)
            for col in range(4, 4 + len(rem_cols)):
                wb_sheet.write(row, col, "", fmt_number_input)
        total_row = end_row + 1
        wb_sheet.write(total_row, 0, "TOTAL", fmt_total)
        wb_sheet.write(total_row, 1, "", fmt_total)
        wb_sheet.write(total_row, 2, "", fmt_total)
        wb_sheet.write_formula(total_row, 3, f"=SUM(D3:D{end_row+1})", fmt_money)
        for col in range(4, 4 + len(rem_cols)):
            col_letter = chr(ord('A') + col)
            # Total quantitativo simples para conferência.
            wb_sheet.write_formula(total_row, col, f"=SUM({col_letter}3:{col_letter}{end_row+1})", fmt_number_input)

    output.seek(0)
    return output.getvalue()


st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Reajustes Múltiplos")

with st.sidebar:
    dt_base_original = st.date_input("Data-Base Original:", value=datetime(2022, 10, 10), format="DD/MM/YYYY")
    qtd_ciclos = st.number_input("Ciclos:", min_value=1, max_value=5, value=2)
    idx_sel = st.selectbox("Índice:", ["IST (Série Local)", "IPCA (433)", "IGP-M (189)"])

# Primeira etapa: coleta dos dados de cada ciclo, sem processar automaticamente os índices.
# Isso evita que a página abra com um cenário fictício já calculado.
input_ciclos = []
containers_ciclos = []
data_atual = dt_base_original

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

    # Lógica de Admissibilidade preservada.
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

    # Regra de ancoragem do próximo ciclo preservada.
    if situacao_limpa == "TEMPESTIVO":
        data_base_proximo_ciclo = dt_ped
    else:
        data_base_proximo_ciclo = d_aniv

    inicio_efeito_financeiro = None if situacao_limpa == "PRECLUSO" else (dt_ped if dt_ped >= d_aniv else d_aniv)

    input_ciclos.append({
        'numero': i,
        'data_atual': data_atual,
        'd_fim': d_fim,
        'd_aniv': d_aniv,
        'd_lim': d_lim,
        'dt_ped': dt_ped,
        'sit_emoji': sit_emoji,
        'situacao_limpa': situacao_limpa,
        'inicio_efeito_financeiro': inicio_efeito_financeiro,
        'data_base_proximo_ciclo': data_base_proximo_ciclo,
    })

    containers_ciclos.append(st.container())
    data_atual = data_base_proximo_ciclo

chave_analise_multiplos = (
    dt_base_original.isoformat(),
    int(qtd_ciclos),
    idx_sel,
    tuple(c['dt_ped'].isoformat() for c in input_ciclos),
)

processar_multiplos = st.button(
    "Processar Análise",
    type="primary",
    use_container_width=True,
)

if processar_multiplos:
    st.session_state["processar_reajustes_multiplos_key"] = chave_analise_multiplos

if st.session_state.get("processar_reajustes_multiplos_key") != chave_analise_multiplos:
    st.info("Informe os dados dos ciclos e clique em **Processar Análise** para iniciar a apuração.")
    st.stop()

# Segunda etapa: processamento dos ciclos, somente após o comando do usuário.
fator_acum = 1.0
historico = []
historico_coleta = []

for idx_ciclo, dados_ciclo in enumerate(input_ciclos):
    i = dados_ciclo['numero']
    data_atual = dados_ciclo['data_atual']
    d_fim = dados_ciclo['d_fim']
    d_aniv = dados_ciclo['d_aniv']
    d_lim = dados_ciclo['d_lim']
    dt_ped = dados_ciclo['dt_ped']
    sit_emoji = dados_ciclo['sit_emoji']
    inicio_efeito_financeiro = dados_ciclo['inicio_efeito_financeiro']

    with containers_ciclos[idx_ciclo]:
        res_c = get_data_rep("433" if "IPCA" in idx_sel else "189", data_atual, d_fim, "IST" in idx_sel)

        # Intervalo exibido independentemente de haver dados disponíveis para o índice.
        if res_c:
            periodo_inicio = res_c['p_ini']
            periodo_fim = res_c['p_fim']
            janela_ciclo = f"{res_c['p_ini'].strftime('%m/%Y')} a {res_c['p_fim'].strftime('%m/%Y')}"
        else:
            if "IST" in idx_sel:
                periodo_inicio = data_atual.replace(day=1)
                periodo_fim = (data_atual + relativedelta(years=1)).replace(day=1)
            else:
                periodo_inicio = data_atual
                periodo_fim = d_fim
            janela_ciclo = f"{periodo_inicio.strftime('%m/%Y')} a {periodo_fim.strftime('%m/%Y')}"

        janela_adm = f"{d_aniv.strftime('%d/%m/%Y')} a {d_lim.strftime('%d/%m/%Y')}"

        st.markdown(f"""
        **Dados do Ciclo {i}:**
        - Intervalo do C{i}: {janela_ciclo}
        - Janela de Admissibilidade: {janela_adm}
        - Situação: {sit_emoji}
        """)

        v_fmt = "Indisponível"
        v_acum_parcial = f"{(fator_acum - 1) * 100:,.2f}%".replace('.', ',')
        fator_ciclo = 1.0
        ciclo_calculado = False

        if res_c:
            fator_ciclo_calculado = 1 + res_c['var']
            v_fmt = f"{res_c['var'] * 100:,.2f}%".replace('.', ',')

            # Ciclos preclusos podem ter a variação apurada para fins de memória,
            # mas não compõem o fator acumulado final nem geram efeito financeiro.
            if dados_ciclo['situacao_limpa'] == "PRECLUSO":
                fator_ciclo = 1.0
                v_acum_parcial = f"{(fator_acum - 1) * 100:,.2f}%".replace('.', ',')
                ciclo_calculado = False
            else:
                fator_ciclo = fator_ciclo_calculado
                fator_acum *= fator_ciclo
                v_acum_parcial = f"{(fator_acum - 1) * 100:,.2f}%".replace('.', ',')
                ciclo_calculado = True

            st.markdown(f"- Variação do Ciclo: **{v_fmt}**")
            if dados_ciclo['situacao_limpa'] == "PRECLUSO":
                st.caption("Ciclo precluso: variação apurada apenas para registro, sem composição no acumulado final.")

            with st.expander(f"🔍 Memória de Cálculo Detalhada - Ciclo {i}"):
                st.write(f"**Metodologia:** {res_c['metodo']}")
                st.write(f"**Janela de Apuração:** {res_c['p_ini'].strftime('%m/%Y')} a {res_c['p_fim'].strftime('%m/%Y')}")

                if "IST" in idx_sel:
                    st.write(
                        f"**Competência inicial:** {res_c['p_ini'].strftime('%m/%Y')} | "
                        f"**Índice inicial:** {res_c['i_ini']}"
                    )
                    st.write(
                        f"**Competência final:** {res_c['p_fim'].strftime('%m/%Y')} | "
                        f"**Índice final:** {res_c['i_fim']}"
                    )
                    st.code(
                        f"({res_c['i_fim']} / {res_c['i_ini']}) - 1 = {res_c['var'] * 100:.4f}%"
                    )
                else:
                    st.dataframe(res_c['dados'], use_container_width=True)
                    st.write("Fórmula: Produtório de (1 + taxa_mensal/100) - 1")

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

        historico_coleta.append({
            'ciclo': f'C{i}',
            'data_base': data_atual.strftime('%d/%m/%Y'),
            'intervalo_indice': janela_ciclo,
            'janela_admissibilidade': janela_adm,
            'data_pedido': dt_ped.strftime('%d/%m/%Y'),
            'situacao': sit_emoji,
            'variacao': float(res_c['var']) if res_c else 0.0,
            'variacao_formatada': v_fmt,
            'fator': float(fator_ciclo),
            'fator_acumulado': float(fator_acum),
            'ciclo_calculado': ciclo_calculado,
            'financeiro_inicio': _formatar_data(inicio_efeito_financeiro),
            'financeiro_fim': '',
            'periodo_inicio': _formatar_data(periodo_inicio),
            'periodo_fim': _formatar_data(periodo_fim),
        })

# Finaliza os períodos financeiros com base no início financeiro do ciclo seguinte.
# Isso evita gerar 13 competências e separa o período de índice do período financeiro.
for idx, ciclo in enumerate(historico_coleta):
    inicio_txt = ciclo.get('financeiro_inicio', '')
    if not inicio_txt:
        ciclo['financeiro_fim'] = ''
        continue
    inicio_dt = pd.to_datetime(inicio_txt, dayfirst=True).to_pydatetime()
    proximo_inicio = None
    for prox in historico_coleta[idx + 1:]:
        prox_txt = prox.get('financeiro_inicio', '')
        if prox_txt:
            proximo_inicio = pd.to_datetime(prox_txt, dayfirst=True).to_pydatetime()
            break
    if proximo_inicio:
        fim_dt = proximo_inicio - relativedelta(months=1)
    else:
        fim_dt = inicio_dt + relativedelta(months=11)
    ciclo['financeiro_fim'] = fim_dt.strftime('%d/%m/%Y')

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

if historico_coleta:
    variacao_acumulada = fator_acum - 1
    st.session_state['dados_admissibilidade'] = {
        'origem': 'Reajustes Múltiplos',
        'tipo': 'Múltiplo',
        'indice': idx_sel,
        'data_base_original': dt_base_original.strftime('%d/%m/%Y'),
        'fator': float(fator_acum),
        'fator_acumulado': float(fator_acum),
        'variacao_acumulada': float(variacao_acumulada),
        'variacao_acumulada_formatada': f"{variacao_acumulada * 100:,.2f}%".replace('.', ','),
        'ciclos': historico_coleta,
    }

    st.download_button(
        label="📥 Gerar Arquivo de Coleta",
        data=gerar_arquivo_coleta_excel(st.session_state['dados_admissibilidade']),
        file_name="Coleta_Reajustes_Multiplos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )
