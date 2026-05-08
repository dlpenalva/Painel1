import streamlit as st
import pandas as pd
import requests
import io
from datetime import datetime
from zoneinfo import ZoneInfo
from dateutil.relativedelta import relativedelta

st.set_page_config(page_title="Análises de Reajustes - Reajustes Múltiplos", layout="wide")

def _carregar_ist_local():
    """Carrega o IST local aceitando os dois layouts usados no projeto.

    Layout novo/atual: MES_ANO;INDICE_NIVEL, com competências como jan/22.
    Layout antigo: data;indice.
    Retorna DataFrame padronizado com colunas data e indice.
    """
    df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
    df.columns = [str(col).strip().lower() for col in df.columns]

    if 'data' in df.columns and 'indice' in df.columns:
        df['data'] = pd.to_datetime(df['data'], dayfirst=True, errors='coerce').dt.normalize()
        df['indice'] = pd.to_numeric(df['indice'], errors='coerce')
    elif 'mes_ano' in df.columns and 'indice_nivel' in df.columns:
        meses = {
            'jan': 1, 'fev': 2, 'mar': 3, 'abr': 4, 'mai': 5, 'jun': 6,
            'jul': 7, 'ago': 8, 'set': 9, 'out': 10, 'nov': 11, 'dez': 12,
        }

        def converter_mes_ano(valor):
            texto = str(valor).strip().lower()
            if '/' not in texto:
                return pd.NaT
            mes_txt, ano_txt = texto.split('/', 1)
            mes = meses.get(mes_txt[:3])
            if mes is None:
                return pd.NaT
            ano = int(ano_txt)
            if ano < 100:
                ano += 2000
            return pd.Timestamp(ano, mes, 1)

        df['data'] = df['mes_ano'].apply(converter_mes_ano)
        df['indice'] = pd.to_numeric(df['indice_nivel'], errors='coerce')
    else:
        raise KeyError("O arquivo ist.csv deve conter as colunas 'data'/'indice' ou 'MES_ANO'/'INDICE_NIVEL'.")

    df = df.dropna(subset=['data', 'indice']).copy()
    return df[['data', 'indice']]


def get_data_rep(serie, d_ini, d_fim, is_ist):
    try:
        if is_ist:
            df = _carregar_ist_local()

            # IST por número-índice: mês-base do ciclo versus o mesmo mês 12 meses depois.
            # Exemplo: data-base 10/2023 => out/2023 a out/2024.
            r_ini = pd.Timestamp(d_ini.year, d_ini.month, 1).normalize()
            marco_final = d_ini + relativedelta(years=1)
            r_fim = pd.Timestamp(marco_final.year, marco_final.month, 1).normalize()

            v_ini_rows = df[df['data'].dt.to_period('M') == r_ini.to_period('M')]
            v_fim_rows = df[df['data'].dt.to_period('M') == r_fim.to_period('M')]

            if v_ini_rows.empty or v_fim_rows.empty:
                st.error(f"Dados do IST não encontrados para o período {r_ini.strftime('%m/%Y')} ou {r_fim.strftime('%m/%Y')}")
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
                'dados': pd.DataFrame({'data': [r_ini, r_fim], 'indice': [v_ini, v_fim]})
            }
        else:
            # Manutenção da lógica do IPCA/IGP-M via SGS/BCB
            url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie}/dados?formato=json&dataInicial={d_ini.strftime('%d/%m/%Y')}&dataFinal={d_fim.strftime('%d/%m/%Y')}"
            response = requests.get(url, timeout=10)
            df_t = pd.DataFrame(response.json())
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


def _render_equacao_ist(res_c):
    equacao_html = f"""
    <div style=\"background:#F4F6F8;border:1px solid #E1E6EB;border-radius:10px;padding:14px 18px;margin-top:10px;\">
        <div style=\"font-family:Consolas, Monaco, monospace;font-size:1.15rem;line-height:1.8;color:#334155;\">
            <span style=\"color:#0F766E;\">({res_c['i_fim']:.3f}</span>
            <span style=\"color:#94A3B8;\"> / </span>
            <span style=\"color:#0F766E;\">{res_c['i_ini']:.3f}</span>
            <span style=\"color:#94A3B8;\">) - 1 = </span>
            <span style=\"color:#B45309;font-weight:600;\">{res_c['var']*100:.4f}%</span>
        </div>
    </div>
    """
    st.markdown(equacao_html, unsafe_allow_html=True)


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




def _parse_moeda_br(valor):
    """Converte texto monetário brasileiro em float, preservando campos vazios como 0."""
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if not texto:
        return 0.0
    texto = texto.replace('R$', '').replace('\xa0', '').replace(' ', '')
    if ',' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    else:
        if texto.count('.') > 1:
            partes = texto.split('.')
            texto = ''.join(partes[:-1]) + '.' + partes[-1]
    try:
        return float(texto)
    except Exception:
        return 0.0


def _formatar_moeda_br(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')


def _formatar_moeda_br_md(valor):
    """Formata moeda escapando o $ para exibição correta no Markdown do Streamlit."""
    return _formatar_moeda_br(valor).replace('$', '\\$')



def _render_card_contexto_contrato():
    """Exibe cabeçalho executivo do bloco de contexto do contrato."""
    st.markdown(
        """
        <div style="background:#F3F6FA;border:1px solid #D9E2EC;border-left:5px solid #1F4E78;border-radius:12px;padding:14px 16px;margin:10px 0 8px 0;">
            <div style="color:#123B63;font-weight:800;font-size:1.02rem;margin-bottom:4px;">Contexto do Contrato</div>
            <div style="color:#475569;font-size:0.92rem;line-height:1.45;">
                Use este bloco quando o contrato já possuir reajustes, repactuações, aditivos ou supressões formalizados antes desta análise.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def _render_contexto_contratual_anterior():
    """Coleta contexto do contrato para uso pelo Valor Global.

    Este bloco é opcional. Quando preenchido, permite que o módulo Valor Global
    parta da fotografia contratual já formalizada antes da análise atual.
    """
    contexto_salvo = st.session_state.get('contexto_contratual_anterior', {}) or {}
    _render_card_contexto_contrato()
    with st.expander('Preencher/editar Contexto do Contrato', expanded=False):
        st.caption(
            'Preencha apenas quando o contrato já possuir reajustes, repactuações, aditivos ou supressões formalizados antes desta análise. '
            'O valor formalizado deve representar a fotografia consolidada do contrato antes dos ciclos agora analisados.'
        )
        col_ctx1, col_ctx2 = st.columns(2)
        with col_ctx1:
            valor_original_txt = st.text_input(
                'Valor original do contrato',
                value=contexto_salvo.get('valor_original_contrato_texto', ''),
                placeholder='Ex.: 20.000.000,00',
                key='ctx_valor_original_contrato',
            )
            opcoes_ciclo_concedido = ['Nenhum / C0', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'Outro / informar em observação']
            ciclo_salvo = str(contexto_salvo.get('ultimo_ciclo_concedido', '') or '').strip()
            if ciclo_salvo in opcoes_ciclo_concedido:
                indice_ciclo_salvo = opcoes_ciclo_concedido.index(ciclo_salvo)
            elif ciclo_salvo == '' or ciclo_salvo.upper() == 'C0':
                indice_ciclo_salvo = 0
            elif ciclo_salvo.upper() in opcoes_ciclo_concedido:
                indice_ciclo_salvo = opcoes_ciclo_concedido.index(ciclo_salvo.upper())
            else:
                indice_ciclo_salvo = len(opcoes_ciclo_concedido) - 1
            ultimo_ciclo = st.selectbox(
                'Último ciclo já concedido/formalizado',
                options=opcoes_ciclo_concedido,
                index=indice_ciclo_salvo,
                key='ctx_ultimo_ciclo_concedido',
            )
        with col_ctx2:
            valor_formalizado_txt = st.text_input(
                'Valor contratual formalizado antes desta análise',
                value=contexto_salvo.get('valor_formalizado_anterior_texto', ''),
                placeholder='Ex.: 22.800.000,00',
                key='ctx_valor_formalizado_anterior',
            )
            observacao = st.text_area(
                'Observação sobre o histórico anterior',
                value=contexto_salvo.get('observacao_historico', ''),
                placeholder='Ex.: C1 e C2 já concedidos; valor inclui aditivo anterior formalizado.',
                key='ctx_observacao_historico',
                height=82,
            )

        valor_original = _parse_moeda_br(valor_original_txt)
        valor_formalizado = _parse_moeda_br(valor_formalizado_txt)
        contexto = {
            'valor_original_contrato': valor_original,
            'valor_original_contrato_texto': valor_original_txt,
            'valor_formalizado_anterior': valor_formalizado,
            'valor_formalizado_anterior_texto': valor_formalizado_txt,
            'ultimo_ciclo_concedido': ultimo_ciclo.strip(),
            'observacao_historico': observacao.strip(),
        }
        st.session_state['contexto_contratual_anterior'] = contexto

        if valor_original > 0 or valor_formalizado > 0 or ultimo_ciclo.strip() or observacao.strip():
            st.info(
                f"Contexto informado: valor original { _formatar_moeda_br_md(valor_original) }; "
                f"valor formalizado antes desta análise { _formatar_moeda_br_md(valor_formalizado) }."
            )
    return st.session_state.get('contexto_contratual_anterior', {})

def gerar_arquivo_coleta_excel(dados_admissibilidade):
    """Gera o Arquivo de Coleta para as fases de Valor Global e Relatório.

    Regras da planilha:
    - 200 linhas pré-formatadas nas abas de preenchimento.
    - Linha 201 reservada para somatórios.
    - Aditivos por item/lançamento, com acréscimo ou supressão.
    """
    output = io.BytesIO()
    ciclos = dados_admissibilidade.get('ciclos', [])
    data_geracao = datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        fmt_header = workbook.add_format({
            'bold': True, 'font_color': 'white', 'bg_color': '#1F4E79',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        fmt_subheader = workbook.add_format({
            'bold': True, 'bg_color': '#D9EAD3', 'border': 1,
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True,
        })
        fmt_input = workbook.add_format({'bg_color': '#FFF2CC', 'border': 1})
        fmt_input_date = workbook.add_format({'num_format': 'dd/mm/yyyy', 'bg_color': '#FFF2CC', 'border': 1})
        fmt_date = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1})
        fmt_input_num = workbook.add_format({'num_format': '#,##0.00', 'bg_color': '#FFF2CC', 'border': 1})
        fmt_input_money = workbook.add_format({'num_format': 'R$ #,##0.00', 'bg_color': '#FFF2CC', 'border': 1})
        fmt_money = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        fmt_money_auto = workbook.add_format({'num_format': 'R$ #,##0.00', 'bg_color': '#EDEDED', 'border': 1})
        fmt_number = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
        fmt_text = workbook.add_format({'border': 1})
        fmt_auto = workbook.add_format({'bg_color': '#EDEDED', 'border': 1})
        fmt_total = workbook.add_format({'bold': True, 'bg_color': '#E2F0D9', 'border': 1})
        fmt_total_money = workbook.add_format({'bold': True, 'bg_color': '#E2F0D9', 'border': 1, 'num_format': 'R$ #,##0.00'})
        fmt_total_num = workbook.add_format({'bold': True, 'bg_color': '#E2F0D9', 'border': 1, 'num_format': '#,##0.00'})
        fmt_decrease_red = workbook.add_format({'font_color': '#C00000'})
        fmt_percent = workbook.add_format({'num_format': '0.00%', 'border': 1})
        fmt_factor = workbook.add_format({'num_format': '0.0000', 'border': 1})
        fmt_factor_auto = workbook.add_format({'num_format': '0.0000', 'bg_color': '#EDEDED', 'border': 1})
        fmt_no_border = workbook.add_format({})
        fmt_int_no_border = workbook.add_format({'num_format': '0'})
        fmt_percent_no_border = workbook.add_format({'num_format': '0.00%'})
        fmt_factor_no_border = workbook.add_format({'num_format': '0.0000'})

        # PARAMETROS_REAJUSTE
        parametros = pd.DataFrame([
            ['Origem da análise', dados_admissibilidade.get('origem', dados_admissibilidade.get('tipo', ''))],
            ['Índice utilizado', dados_admissibilidade.get('indice', '')],
            ['Data-base original', dados_admissibilidade.get('data_base_original', '')],
            ['Quantidade de ciclos', len(ciclos)],
            ['Variação acumulada final', dados_admissibilidade.get('variacao_acumulada', 0.0)],
            ['Fator acumulado total', dados_admissibilidade.get('fator_acumulado', dados_admissibilidade.get('fator', 1.0))],
            ['Valor original do contrato (contexto)', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('valor_original_contrato', '')],
            ['Valor contratual formalizado antes desta análise', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('valor_formalizado_anterior', '')],
            ['Último ciclo já concedido/formalizado', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('ultimo_ciclo_concedido', '')],
            ['Observação sobre histórico anterior', dados_admissibilidade.get('contexto_contratual_anterior', {}).get('observacao_historico', '')],
            ['Data de geração do arquivo', data_geracao],
        ], columns=['Campo', 'Valor'])
        parametros.to_excel(writer, sheet_name='PARAMETROS_REAJUSTE', index=False)
        ws = writer.sheets['PARAMETROS_REAJUSTE']
        ws.set_column('A:A', 34)
        ws.set_column('B:B', 36)
        ws.write(0, 0, 'Campo', fmt_header)
        ws.write(0, 1, 'Valor', fmt_header)
        ws.write_number(4, 1, len(ciclos), fmt_int_no_border)     # B5 Quantidade de ciclos
        ws.write_number(5, 1, float(dados_admissibilidade.get('variacao_acumulada', 0.0)), fmt_percent_no_border)  # B6
        ws.write_number(6, 1, float(dados_admissibilidade.get('fator_acumulado', dados_admissibilidade.get('fator', 1.0))), fmt_factor_no_border)  # B7
        contexto_excel = dados_admissibilidade.get('contexto_contratual_anterior', {}) or {}
        ws.write_number(7, 1, float(contexto_excel.get('valor_original_contrato') or 0.0), fmt_money)
        ws.write_number(8, 1, float(contexto_excel.get('valor_formalizado_anterior') or 0.0), fmt_money)

        # CICLOS
        ciclos_rows = []
        for ciclo in ciclos:
            situacao = str(ciclo.get('situacao', ''))
            tratamento = 'Precluso' if 'PRECLUSO' in situacao.upper() else 'A apurar'
            ciclos_rows.append({
                'Ciclo': ciclo.get('ciclo', ''),
                'Data-base': ciclo.get('data_base', ''),
                'Intervalo do índice': ciclo.get('intervalo_indice', ciclo.get('Janela', '')),
                'Janela de admissibilidade': ciclo.get('janela_admissibilidade', ciclo.get('JanelaAdm', '')),
                'Data do pedido': ciclo.get('data_pedido', ciclo.get('Pedido', '')),
                'Início financeiro': ciclo.get('financeiro_inicio', ''),
                'Fim financeiro': ciclo.get('financeiro_fim', ''),
                'Situação': situacao,
                'Variação': ciclo.get('variacao', 0.0),
                'Fator': ciclo.get('fator', 1.0),
                'Fator acumulado': ciclo.get('fator_acumulado', 1.0),
                'Tratamento financeiro do ciclo': tratamento,
                'Situação automática': ciclo.get('situacao_automatica', ciclo.get('situacao', '')),
                'Acordo negocial': 'Sim' if ciclo.get('superacao_negocial', False) else 'Não',
                'Situação aplicada': ciclo.get('situacao_aplicada', ciclo.get('situacao', '')),
                'Percentual apurado pelo índice': ciclo.get('percentual_indice', ciclo.get('variacao', 0.0)),
                'Percentual aplicado': ciclo.get('percentual_aplicado', ciclo.get('variacao', 0.0)),
                'Justificativa negocial': ciclo.get('justificativa_negocial', ''),
                'Referência documental': ciclo.get('referencia_documental', ''),
            })
        df_ciclos = pd.DataFrame(ciclos_rows)
        df_ciclos.to_excel(writer, sheet_name='CICLOS', index=False)
        ws = writer.sheets['CICLOS']
        ws.set_column('A:A', 12)
        ws.set_column('B:H', 24)
        ws.set_column('I:I', 12)
        ws.set_column('J:K', 14)
        ws.set_column('L:L', 30)
        ws.set_column('M:S', 26)
        for col, title in enumerate(df_ciclos.columns):
            ws.write(0, col, title, fmt_header)
        for row in range(1, len(df_ciclos) + 1):
            data_base_excel = pd.to_datetime(df_ciclos.iloc[row-1]['Data-base'], dayfirst=True, errors='coerce')
            if pd.notna(data_base_excel):
                ws.write_datetime(row, 1, data_base_excel.to_pydatetime(), fmt_date)
            ws.write(row, 8, df_ciclos.iloc[row-1]['Variação'], workbook.add_format({'num_format': '0.00%'}))
            ws.write(row, 9, df_ciclos.iloc[row-1]['Fator'], workbook.add_format({'num_format': '0.0000'}))
            ws.write(row, 10, df_ciclos.iloc[row-1]['Fator acumulado'], workbook.add_format({'num_format': '0.0000'}))
            ws.write(row, 11, df_ciclos.iloc[row-1]['Tratamento financeiro do ciclo'], workbook.add_format({}))

        # FINANCEIRO_MENSAL
        financeiro_rows = []
        for ciclo in ciclos:
            ciclo_nome = ciclo.get('ciclo', '')
            for competencia in _competencias_mensais(ciclo.get('financeiro_inicio', ''), ciclo.get('financeiro_fim', '')):
                financeiro_rows.append({'Ciclo': ciclo_nome, 'Competência': competencia, 'Valor pago/faturado': ''})
        while len(financeiro_rows) < 199:
            financeiro_rows.append({'Ciclo': '', 'Competência': '', 'Valor pago/faturado': ''})
        financeiro_rows = financeiro_rows[:199]
        df_fin = pd.DataFrame(financeiro_rows)
        df_fin.to_excel(writer, sheet_name='FINANCEIRO_MENSAL', index=False)
        ws = writer.sheets['FINANCEIRO_MENSAL']
        ws.set_column('A:A', 12)
        ws.set_column('B:B', 18)
        ws.set_column('C:C', 24)
        for col, title in enumerate(df_fin.columns):
            ws.write(0, col, title, fmt_header)
        for row in range(1, 200):
            ws.write(row, 2, '', fmt_input_money)
        ws.write(200, 0, 'TOTAL', fmt_total)
        ws.write(200, 1, '', fmt_total)
        ws.write_formula(200, 2, '=SUM(C2:C200)', fmt_total_money)

        # ITENS_REMANESCENTES
        ws_it = workbook.add_worksheet('ITENS_REMANESCENTES')
        writer.sheets['ITENS_REMANESCENTES'] = ws_it
        rem_cols = []
        for ciclo in ciclos:
            rem_cols.append((f"Remanescente início {ciclo.get('ciclo', '')}", ciclo.get('data_base', '')))
        base_headers = ['Item', 'Quantidade contratada', 'Valor unitário original', 'Valor total']
        headers = base_headers + [c[0] for c in rem_cols]
        ws_it.merge_range(0, 0, 0, 3, 'Dados do item e valor original', fmt_subheader)
        for idx, (_, data_ref) in enumerate(rem_cols, start=4):
            ws_it.write(0, idx, data_ref, fmt_subheader)
        for col, title in enumerate(headers):
            ws_it.write(1, col, title, fmt_header)
        ws_it.set_column(0, 0, 12)
        ws_it.set_column(1, 1, 24)
        ws_it.set_column(2, 3, 22)
        if rem_cols:
            ws_it.set_column(4, 4 + len(rem_cols) - 1, 25)
        for row in range(2, 200):
            ws_it.write(row, 0, '', fmt_text)       # A sem preenchimento
            ws_it.write(row, 1, '', fmt_number)     # B sem preenchimento
            ws_it.write(row, 2, '', fmt_money)      # C sem preenchimento
            ws_it.write_formula(row, 3, f'=IF(OR(B{row+1}="",C{row+1}=""),"",B{row+1}*C{row+1})', fmt_money_auto)
            for col in range(4, 4 + len(rem_cols)):
                ws_it.write(row, col, '', fmt_input_num)
        ws_it.write(200, 0, 'TOTAL', fmt_total)
        ws_it.write(200, 1, '', fmt_total)
        ws_it.write(200, 2, '', fmt_total)
        ws_it.write_formula(200, 3, '=SUM(D3:D200)', fmt_total_money)
        for col in range(4, 4 + len(rem_cols)):
            col_letter = chr(ord('A') + col)
            ws_it.write_formula(200, col, f'=SUM({col_letter}3:{col_letter}200)', fmt_total_num)

        # ADITIVOS_QUANTITATIVOS
        ws_ad = workbook.add_worksheet('ADITIVOS_QUANTITATIVOS')
        writer.sheets['ADITIVOS_QUANTITATIVOS'] = ws_ad
        ad_headers = [
            'Item', 'Data do aditivo', 'Ciclo/Marco', 'Tipo de alteração',
            'Quantidade acrescida/suprimida', 'Valor unitário original',
            'Valor original da alteração', 'Aplicar reajuste acumulado? (Sim/Não)',
            'Fator acumulado aplicável', 'Valor atualizado da alteração',
            'Tratamento do aditivo'
        ]
        for col, title in enumerate(ad_headers):
            ws_ad.write(0, col, title, fmt_header)
        ws_ad.set_column('A:A', 12)
        ws_ad.set_column('B:B', 18)
        ws_ad.set_column('C:C', 16)
        ws_ad.set_column('D:D', 20)
        ws_ad.set_column('E:E', 28)
        ws_ad.set_column('F:G', 24)
        ws_ad.set_column('H:H', 32)
        ws_ad.set_column('I:I', 22)
        ws_ad.set_column('J:J', 26)
        ws_ad.set_column('K:K', 42)
        ws_ad.write(202, 0, "Orientação", fmt_total)
        ws_ad.write(202, 1, "Use 'Computar nesta análise' para aditivos/supressões que devem impactar o Valor Global atual. Use 'Informativo - já incluído no valor formalizado' quando o lançamento já estiver contemplado no campo Valor contratual formalizado antes desta análise.", fmt_text)
        ciclo_range = f'CICLOS!$A$2:$K${len(ciclos)+1}' if ciclos else 'CICLOS!$A$2:$K$2'
        data_range = f'CICLOS!$B$2:$B${len(ciclos)+1}' if ciclos else 'CICLOS!$B$2:$B$2'
        ciclo_nome_range = f'CICLOS!$A$2:$A${len(ciclos)+1}' if ciclos else 'CICLOS!$A$2:$A$2'
        for row in range(1, 200):
            excel_row = row + 1
            ws_ad.write(row, 0, '', fmt_text)
            ws_ad.write(row, 1, '', fmt_input_date)
            # Coluna C: identifica automaticamente o ciclo pela data do aditivo.
            # Regra: último ciclo cuja Data-base seja menor ou igual à Data do Aditivo.
            if len(ciclos) > 0:
                formula_ciclo = (
                    f'=IF(B{excel_row}="","",'
                    f'IF(B{excel_row}<MIN(CICLOS!$B$2:$B${len(ciclos)+1}),"C0",'
                    f'IFERROR(LOOKUP(2,1/(CICLOS!$B$2:$B${len(ciclos)+1}<=B{excel_row}),'
                    f'CICLOS!$A$2:$A${len(ciclos)+1}),"Fora de Ciclo")))'
                )
                ws_ad.write_formula(row, 2, formula_ciclo, fmt_auto)
            else:
                ws_ad.write(row, 2, '', fmt_auto)
            ws_ad.write(row, 3, 'Acréscimo', fmt_input)
            ws_ad.write(row, 4, '', fmt_input_num)
            ws_ad.write(row, 5, '', fmt_input_money)
            ws_ad.write_formula(row, 6, f'=IF(OR(E{excel_row}="",F{excel_row}=""),"",E{excel_row}*F{excel_row})', fmt_money_auto)
            ws_ad.write(row, 7, 'Sim', fmt_input)
            ws_ad.write_formula(row, 8, f'=IF(C{excel_row}="","",IFERROR(VLOOKUP(C{excel_row},{ciclo_range},11,FALSE),1))', fmt_factor_auto)
            ws_ad.write_formula(row, 9, f'=IF(G{excel_row}="","",IF(OR(UPPER(D{excel_row})="DECRÉSCIMO",UPPER(D{excel_row})="DECRESCIMO",UPPER(D{excel_row})="SUPRESSÃO",UPPER(D{excel_row})="SUPRESSAO"),-ABS(G{excel_row}),ABS(G{excel_row}))*IF(OR(UPPER(H{excel_row})="NÃO",UPPER(H{excel_row})="NAO"),1,I{excel_row}))', fmt_money_auto)
            ws_ad.write(row, 10, 'Computar nesta análise', fmt_input)
        ws_ad.data_validation(1, 3, 199, 3, {'validate': 'list', 'source': ['Acréscimo', 'Decréscimo']})
        ws_ad.data_validation(1, 7, 199, 7, {'validate': 'list', 'source': ['Sim', 'Não']})
        ws_ad.data_validation(1, 10, 199, 10, {'validate': 'list', 'source': ['Computar nesta análise', 'Informativo - já incluído no valor formalizado']})
        # Se o usuário selecionar Decréscimo/Supressão, destacar toda a linha em fonte vermelha.
        ws_ad.conditional_format(1, 0, 199, 10, {
            'type': 'formula',
            # Fórmula robusta: usa início do texto para evitar falhas por acento/localidade.
            'criteria': '=OR(LEFT(UPPER($D2),4)="DECR",LEFT(UPPER($D2),6)="SUPRES")',
            'format': fmt_decrease_red,
        })
        ws_ad.write(200, 0, 'TOTAL', fmt_total)
        for col in range(1, 6):
            ws_ad.write(200, col, '', fmt_total)
        ws_ad.write_formula(200, 6, '=SUM(G2:G200)', fmt_total_money)
        ws_ad.write(200, 7, '', fmt_total)
        ws_ad.write(200, 8, '', fmt_total)
        ws_ad.write_formula(200, 9, '=SUM(J2:J200)', fmt_total_money)
        ws_ad.write(200, 10, '', fmt_total)

    output.seek(0)
    return output.getvalue()


st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Reajustes Múltiplos")

contexto_contratual = _render_contexto_contratual_anterior()

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
        # A chave considera a âncora do ciclo. Assim, se um ciclo anterior for admitido
        # por negociação entre as partes e arrastar a data-base para frente, o campo do
        # pedido do ciclo seguinte também é recalculado a partir da nova âncora.
        chave_pedido = f"p{i}_{data_atual.strftime('%Y%m%d')}"
        dt_ped = st.date_input(
            f"Data do Pedido C{i}:",
            value=d_aniv,
            key=chave_pedido,
            format="DD/MM/YYYY",
        )

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

    superacao_negocial = False
    percentual_negocial = 0.0
    justificativa_negocial = ""
    referencia_documental = ""
    data_inicio_efeito_negocial = None
    if situacao_limpa == "PRECLUSO":
        with st.expander(f"Acordo negocial de admissão de reajuste - C{i}", expanded=False):
            st.caption(
                "Use apenas quando houver decisão negocial fundamentada para conceder o ciclo, "
                "sem apagar o diagnóstico automático de preclusão."
            )
            superacao_negocial = st.checkbox(
                f"Ciclo admitido por negociação entre as partes - C{i}",
                value=False,
                key=f"superacao_negocial_c{i}",
            )
            if superacao_negocial:
                col_neg1, col_neg2 = st.columns(2)
                with col_neg1:
                    percentual_negocial = st.number_input(
                        f"Percentual aplicado por acordo C{i} (%)",
                        min_value=0.0,
                        max_value=100.0,
                        value=0.0,
                        step=0.01,
                        format="%.4f",
                        key=f"percentual_negocial_c{i}",
                    ) / 100
                    data_inicio_efeito_negocial = st.date_input(
                        f"Início dos efeitos financeiros por acordo C{i}",
                        value=dt_ped,
                        format="DD/MM/YYYY",
                        key=f"inicio_negocial_c{i}",
                    )
                with col_neg2:
                    referencia_documental = st.text_input(
                        f"Referência documental C{i}",
                        placeholder="Ex.: Despacho, Ata, Ofício ou Nota Técnica",
                        key=f"referencia_negocial_c{i}",
                    )
                justificativa_negocial = st.text_area(
                    f"Justificativa técnica/negocial C{i}",
                    placeholder="Registre a fundamentação da concessão por acordo negocial.",
                    key=f"justificativa_negocial_c{i}",
                    height=90,
                )
                if not justificativa_negocial.strip():
                    st.info("A justificativa deve ser preenchida para fins de memória processual antes da instrução final.")
                inicio_efeito_financeiro = data_inicio_efeito_negocial
                # Regra de ancoragem negocial:
                # se um ciclo precluso for admitido por negociação entre as partes, a âncora
                # do ciclo seguinte é arrastada para a data de início dos efeitos financeiros
                # pactuada para o ciclo admitido. Consequentemente, o próximo ciclo somente
                # estará apto após 12 meses desse novo marco.
                data_base_proximo_ciclo = data_inicio_efeito_negocial
                st.info(
                    f"Com a admissão negocial do C{i}, o próximo ciclo será ancorado em "
                    f"{data_base_proximo_ciclo.strftime('%d/%m/%Y')} e somente estará apto "
                    f"a partir de {(data_base_proximo_ciclo + relativedelta(years=1)).strftime('%d/%m/%Y')}."
                )

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
        'superacao_negocial': bool(superacao_negocial),
        'percentual_negocial': float(percentual_negocial),
        'justificativa_negocial': justificativa_negocial.strip(),
        'referencia_documental': referencia_documental.strip(),
        'data_base_proximo_ciclo': data_base_proximo_ciclo,
    })

    containers_ciclos.append(st.container())
    data_atual = data_base_proximo_ciclo

chave_analise_multiplos = (
    dt_base_original.isoformat(),
    int(qtd_ciclos),
    idx_sel,
    tuple(c['dt_ped'].isoformat() for c in input_ciclos),
    tuple(str(c.get('superacao_negocial', False)) for c in input_ciclos),
    tuple(str(c.get('percentual_negocial', 0.0)) for c in input_ciclos),
    tuple(str(c.get('justificativa_negocial', '')) for c in input_ciclos),
    str(contexto_contratual.get('valor_original_contrato', 0.0)),
    str(contexto_contratual.get('valor_formalizado_anterior', 0.0)),
    contexto_contratual.get('ultimo_ciclo_concedido', ''),
    contexto_contratual.get('observacao_historico', ''),
)

processar_multiplos = st.button(
    "Processar Análise",
    type="primary",
    use_container_width=False,
)

if processar_multiplos:
    st.session_state["processar_reajustes_multiplos_key"] = chave_analise_multiplos

if st.session_state.get("processar_reajustes_multiplos_key") != chave_analise_multiplos:
    st.info(f"Foram configurados **{int(qtd_ciclos)} ciclos** para análise a partir da Data-Base original de **{dt_base_original.strftime('%d/%m/%Y')}**. Confira as datas dos pedidos antes de clicar em **Processar Análise**.")
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
    situacao_limpa = dados_ciclo['situacao_limpa']
    inicio_efeito_financeiro = dados_ciclo['inicio_efeito_financeiro']
    superacao_negocial = bool(dados_ciclo.get('superacao_negocial', False))
    percentual_negocial = float(dados_ciclo.get('percentual_negocial', 0.0) or 0.0)
    justificativa_negocial = dados_ciclo.get('justificativa_negocial', '')
    referencia_documental = dados_ciclo.get('referencia_documental', '')

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
            fator_indice = 1 + res_c['var']
            percentual_indice = float(res_c['var'])
            percentual_aplicado = percentual_indice
            situacao_aplicada = sit_emoji
            if situacao_limpa == "PRECLUSO" and superacao_negocial:
                percentual_aplicado = percentual_negocial
                fator_ciclo = 1 + percentual_aplicado
                situacao_aplicada = "🟣 CICLO ADMITIDO POR NEGOCIAÇÃO ENTRE AS PARTES"
            elif situacao_limpa == "PRECLUSO":
                percentual_aplicado = 0.0
                fator_ciclo = 1.0
            else:
                fator_ciclo = fator_indice
            fator_acum *= fator_ciclo
            v_fmt = f"{res_c['var'] * 100:,.2f}%".replace('.', ',')
            v_aplicado_fmt = f"{percentual_aplicado * 100:,.2f}%".replace('.', ',')
            v_acum_parcial = f"{(fator_acum - 1) * 100:,.2f}%".replace('.', ',')
            ciclo_calculado = True

            st.markdown(f"- Variação do Ciclo: **{v_fmt}**")
            if situacao_limpa == "PRECLUSO" and superacao_negocial:
                st.caption(f"Ciclo admitido por negociação entre as partes. Percentual aplicado: {v_aplicado_fmt}.")
            elif situacao_limpa == "PRECLUSO":
                st.caption("Variação apurada apenas para registro, sem composição no acumulado final.")

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
                    _render_equacao_ist(res_c)
                else:
                    st.dataframe(res_c['dados'], use_container_width=True)
                    st.write("Fórmula: Produtório de (1 + taxa_mensal/100) - 1")

            historico.append({
                "Ciclo": i,
                "Variação": v_fmt,
                "Percentual aplicado": v_aplicado_fmt,
                "Acumulada": v_acum_parcial,
                "Situação": situacao_aplicada,
                "Situação automática": sit_emoji,
                "Acordo negocial": bool(superacao_negocial),
                "Pedido": dt_ped.strftime('%d/%m/%Y'),
                "Janela": janela_ciclo,
                "JanelaAdm": janela_adm,
                "Início financeiro": inicio_efeito_financeiro.strftime('%d/%m/%Y') if inicio_efeito_financeiro else "",
            })
        else:
            percentual_indice = 0.0
            percentual_aplicado = 0.0
            situacao_aplicada = sit_emoji
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
            'situacao': situacao_aplicada,
            'situacao_automatica': sit_emoji,
            'situacao_aplicada': situacao_aplicada,
            'superacao_negocial': bool(superacao_negocial),
            'percentual_indice': float(percentual_indice),
            'percentual_aplicado': float(percentual_aplicado),
            'justificativa_negocial': justificativa_negocial.strip(),
            'referencia_documental': referencia_documental.strip(),
            'variacao': float(percentual_aplicado),
            'variacao_formatada': f"{percentual_aplicado * 100:,.2f}%".replace('.', ','),
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
        if h.get("Acordo negocial"):
            corpo_relatorio += f"""
            **C{h['Ciclo']}:** Pedido em {h['Pedido']}. Intervalo do C{h['Ciclo']}: {h['Janela']}.  
            Janela de Admissibilidade: {h['JanelaAdm']}.  
            **Resultado automático:** {h.get('Situação automática', h.get('Situação', ''))}.  
            **Tratamento aplicado:** (*) ciclo admitido por negociação entre as partes.  
            Variação apurada pelo índice: {h['Variação']}.  
            Percentual aplicado por acordo: {h.get('Percentual aplicado', h['Variação'])}.  
            Índice {idx_sel}.  
            Data de início dos efeitos financeiros por acordo: {h.get('Início financeiro', '')}.

            (*) O diagnóstico automático de preclusão foi preservado. O ciclo foi considerado aplicável por decisão negocial registrada pelo usuário.
            \n\n"""
        else:
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
        'contexto_contratual_anterior': contexto_contratual,
        'fator': float(fator_acum),
        'fator_acumulado': float(fator_acum),
        'variacao_acumulada': float(variacao_acumulada),
        'variacao_acumulada_formatada': f"{variacao_acumulada * 100:,.2f}%".replace('.', ','),
        'ciclos': historico_coleta,
    }

    st.download_button(
        label="📥 Gerar Arquivo de Coleta",
        type="primary",
        data=gerar_arquivo_coleta_excel(st.session_state['dados_admissibilidade']),
        file_name="Coleta_Reajustes_Multiplos.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )
