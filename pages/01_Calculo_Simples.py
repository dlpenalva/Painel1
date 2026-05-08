import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from zoneinfo import ZoneInfo
from dateutil.relativedelta import relativedelta
from io import BytesIO

st.set_page_config(page_title="Análises de Reajustes - Reajuste Simples", layout="wide")


def get_index_data(serie_codigo, data_inicio, data_fim):
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{serie_codigo}/dados?formato=json&dataInicial={data_inicio.strftime('%d/%m/%Y')}&dataFinal={data_fim.strftime('%d/%m/%Y')}"
    try:
        response = requests.get(url, timeout=15)
        df = pd.DataFrame(response.json())
        if df.empty:
            return None
        df['valor_decimal'] = df['valor'].astype(float) / 100
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        var_final = (1 + df['valor_decimal']).prod() - 1
        return {'variacao': var_final, 'metodo': "Produtório de taxas mensais (SGS/BCB)", 'dados': df[['data', 'valor']]}
    except Exception:
        return None


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


def get_ist_local(data_inicio, data_fim):
    try:
        df = _carregar_ist_local()

        # IST por número-índice: mês-base do ciclo versus o mesmo mês 12 meses depois.
        # Exemplo: data-base 10/2023 => out/2023 a out/2024.
        r_ini = pd.Timestamp(data_inicio.year, data_inicio.month, 1).normalize()
        marco_final = data_inicio + relativedelta(years=1)
        r_fim = pd.Timestamp(marco_final.year, marco_final.month, 1).normalize()

        v_ini_rows = df[df['data'].dt.to_period('M') == r_ini.to_period('M')]
        v_fim_rows = df[df['data'].dt.to_period('M') == r_fim.to_period('M')]

        if v_ini_rows.empty or v_fim_rows.empty:
            return None

        v_ini = float(v_ini_rows['indice'].iloc[0])
        v_fim = float(v_fim_rows['indice'].iloc[0])

        return {
            'variacao': (v_fim / v_ini) - 1,
            'i_ini': v_ini,
            'i_fim': v_fim,
            'd_ini': r_ini,
            'd_fim': r_fim,
            'metodo': "Divisão de Número-Índice (Série Local)"
        }
    except Exception:
        return None


def _formatar_data(valor):
    try:
        return pd.to_datetime(valor).strftime('%d/%m/%Y')
    except Exception:
        return ""


def _formatar_mes_ano(valor):
    try:
        return pd.to_datetime(valor).strftime('%m/%Y')
    except Exception:
        return ""


def _competencias_mensais(data_inicio, data_fim):
    if not data_inicio or not data_fim:
        return []
    inicio = pd.to_datetime(data_inicio, dayfirst=True).replace(day=1)
    fim = pd.to_datetime(data_fim, dayfirst=True).replace(day=1)
    if fim < inicio:
        return []
    return [d.strftime('%m/%Y') for d in pd.date_range(inicio, fim, freq='MS')]




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

def _render_equacao_ist(i_ini, i_fim, variacao):
    equacao_html = f"""
    <div style="background:#F4F6F8;border:1px solid #E1E6EB;border-radius:10px;padding:14px 18px;margin-top:10px;">
        <div style="font-family:Consolas, Monaco, monospace;font-size:1.15rem;line-height:1.8;color:#334155;">
            <span style="color:#0F766E;">({i_fim:.3f}</span>
            <span style="color:#94A3B8;"> / </span>
            <span style="color:#0F766E;">{i_ini:.3f}</span>
            <span style="color:#94A3B8;">) - 1 = </span>
            <span style="color:#B45309;font-weight:600;">{variacao*100:.4f}%</span>
        </div>
    </div>
    """
    st.markdown(equacao_html, unsafe_allow_html=True)


def _ajustar_larguras(writer, nomes_abas):
    from openpyxl.utils import get_column_letter

    for nome_aba in nomes_abas:
        ws = writer.book[nome_aba]
        # ITENS_REMANESCENTES usa duas linhas de cabeçalho: linha 1 para datas de referência e linha 2 para títulos.
        ws.freeze_panes = "A3" if nome_aba == 'ITENS_REMANESCENTES' else "A2"
        for column_cells in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column_cells[0].column)
            for cell in column_cells:
                try:
                    max_length = max(max_length, len(str(cell.value)) if cell.value is not None else 0)
                except Exception:
                    pass
            ws.column_dimensions[column_letter].width = min(max(max_length + 2, 14), 45)


def _aplicar_estilos_coleta(writer, ciclos):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(color='FFFFFF', bold=True)
    input_fill = PatternFill('solid', fgColor='FFF2CC')
    date_fill = PatternFill('solid', fgColor='D9EAD3')
    light_fill = PatternFill('solid', fgColor='EAF2F8')
    thin = Side(style='thin', color='D9E2F3')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    money_fmt = 'R$ #,##0.00'
    number_fmt = '#,##0.0000'

    for ws in writer.book.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical='center')

        header_row = 2 if ws.title == 'ITENS_REMANESCENTES' else 1
        for cell in ws[header_row]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    if 'FINANCEIRO_MENSAL' in writer.book.sheetnames:
        ws = writer.book['FINANCEIRO_MENSAL']
        # Coluna C é a coluna de preenchimento pelo fiscal.
        for row in range(1, ws.max_row + 1):
            ws.cell(row=row, column=3).fill = input_fill
            ws.cell(row=row, column=3).alignment = Alignment(horizontal='right' if row > 1 else 'center', vertical='center')
            if row > 1:
                ws.cell(row=row, column=3).number_format = money_fmt
        ws['C1'].value = 'Valor pago/faturado (preencher)'
        ws['C1'].font = header_font
        ws['C1'].fill = header_fill

    if 'ITENS_REMANESCENTES' in writer.book.sheetnames:
        ws = writer.book['ITENS_REMANESCENTES']
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
        ws.cell(row=1, column=1).value = 'Dados do item e valor original'
        ws.cell(row=1, column=1).fill = light_fill
        ws.cell(row=1, column=1).font = Font(bold=True)
        ws.cell(row=1, column=1).alignment = Alignment(horizontal='center', vertical='center')

        # Colunas: A Item | B Quantidade contratada | C Valor unitário original | D Valor total | E... remanescentes
        for row in range(3, ws.max_row + 1):
            ws.cell(row=row, column=3).number_format = money_fmt
            ws.cell(row=row, column=4).number_format = money_fmt
            ws.cell(row=row, column=4).value = f'=IF(OR(B{row}="",C{row}=""),"",B{row}*C{row})'

        # Destacar colunas de remanescente para preenchimento pelo fiscal e colocar a data de referência na linha 1.
        first_rem_col = 5
        for idx, ciclo in enumerate(ciclos):
            col = first_rem_col + idx
            letra = get_column_letter(col)
            ws.cell(row=1, column=col).value = ciclo.get('data_base', '')
            ws.cell(row=1, column=col).fill = date_fill
            ws.cell(row=1, column=col).font = Font(bold=True)
            ws.cell(row=1, column=col).alignment = Alignment(horizontal='center', vertical='center')
            for row in range(2, ws.max_row + 1):
                ws.cell(row=row, column=col).fill = input_fill if row > 2 else header_fill
                if row == 2:
                    ws.cell(row=row, column=col).font = header_font
                    ws.cell(row=row, column=col).alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws.column_dimensions[letra].width = 24

    if 'CICLOS' in writer.book.sheetnames:
        ws = writer.book['CICLOS']
        for row in range(2, ws.max_row + 1):
            # Fator e Fator acumulado.
            if ws.max_column >= 8:
                ws.cell(row=row, column=8).number_format = number_fmt
            if ws.max_column >= 9:
                ws.cell(row=row, column=9).number_format = number_fmt

    if 'ADITIVOS_QUANTITATIVOS' in writer.book.sheetnames:
        ws = writer.book['ADITIVOS_QUANTITATIVOS']
        for row in range(2, ws.max_row + 1):
            ws.cell(row=row, column=4).number_format = money_fmt
            ws.cell(row=row, column=4).fill = input_fill


def gerar_arquivo_coleta_excel(dados_admissibilidade):
    """Gera o Arquivo de Coleta para as fases de Valor Global e Relatório.

    Regras da planilha:
    - 200 linhas pré-formatadas nas abas de preenchimento.
    - Linha 201 reservada para somatórios.
    - Aditivos por item/lançamento, com acréscimo ou supressão.
    """
    output = BytesIO()
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
st.title("Reajuste Simples")

contexto_contratual = _render_contexto_contratual_anterior()

col1, col2 = st.columns(2)
with col1:
    dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 8, 2), format="DD/MM/YYYY")
    dt_solic = st.date_input("Data do Pedido:", value=datetime(2024, 4, 9), format="DD/MM/YYYY")
with col2:
    tipo_idx = st.selectbox("Índice:", ["IST (Série Local)", "IPCA (433)", "IGP-M (189)"])

chave_analise_simples = (
    dt_base.isoformat(),
    dt_solic.isoformat(),
    tipo_idx,
    str(contexto_contratual.get('valor_original_contrato', 0.0)),
    str(contexto_contratual.get('valor_formalizado_anterior', 0.0)),
    contexto_contratual.get('ultimo_ciclo_concedido', ''),
    contexto_contratual.get('observacao_historico', ''),
)

st.info("Informe os dados e clique em Processar Análise para iniciar a apuração.")

if st.button("Processar Análise", type="primary", use_container_width=False):
    st.session_state["chave_analise_simples_processada"] = chave_analise_simples

if st.session_state.get("chave_analise_simples_processada") != chave_analise_simples:
    st.stop()

# Definição de datas do ciclo
dt_fim_ap = dt_base + relativedelta(months=11)
dt_aniv = dt_base + relativedelta(years=1)
dt_limite = dt_aniv + relativedelta(days=90)

# Regra de Admissibilidade
if dt_solic < dt_aniv:
    if dt_solic.year == dt_aniv.year and dt_solic.month == dt_aniv.month:
        status_ped = "🟡 ADMISSÍVEL - RESSALVA"
    else:
        status_ped = "⚠️ ADIANTADO (ANTES DOS 12 MESES)"
elif dt_solic <= dt_limite:
    status_ped = "✅ TEMPESTIVO"
else:
    status_ped = "❌ PRECLUSO"

res = get_ist_local(dt_base, dt_fim_ap) if "IST" in tipo_idx else get_index_data("433" if "IPCA" in tipo_idx else "189", dt_base, dt_fim_ap)

if res:
    v_fmt = f"{res['variacao']*100:,.2f}%".replace('.', ',')
    fator_ciclo = 1 + res['variacao']
    st.metric("Variação Apurada", v_fmt)

    st.markdown("### Dados do Ciclo")
    # Ajuste solicitado de nomenclatura e nova linha de janela
    periodo_inicio = res['d_ini'] if 'd_ini' in res else dt_base
    periodo_fim = res['d_fim'] if 'd_fim' in res else dt_fim_ap
    janela_str = f"{pd.to_datetime(periodo_inicio).strftime('%m/%Y')} a {pd.to_datetime(periodo_fim).strftime('%m/%Y')}"
    janela_adm_str = f"{dt_aniv.strftime('%d/%m/%Y')} a {dt_limite.strftime('%d/%m/%Y')}"

    st.write(f"- **Intervalo do C1:** {janela_str}")
    st.write(f"- **Janela de Admissibilidade:** {janela_adm_str}")
    st.write(f"- **Situação:** {status_ped}")

    # Acordo negocial: preserva o diagnóstico automático, mas permite aplicar percentual manual.
    superacao_negocial = False
    percentual_indice = float(res['variacao'])
    percentual_aplicado = percentual_indice
    fator_ciclo_efetivo = float(fator_ciclo)
    situacao_automatica = status_ped
    situacao_aplicada = status_ped
    justificativa_negocial = ""
    referencia_documental = ""
    data_inicio_efeito_negocial = None

    if "PRECLUSO" in status_ped.upper():
        st.warning(
            "O ciclo foi classificado automaticamente como precluso. "
            "Caso exista decisão negocial fundamentada na cláusula contratual aplicável, "
            "é possível registrar a admissão do reajuste por acordo entre as partes, sem apagar o diagnóstico automático."
        )
        with st.expander("Acordo negocial de admissão de reajuste", expanded=False):
            superacao_negocial = st.checkbox(
                "Ciclo admitido por negociação entre as partes",
                value=False,
                key="superacao_negocial_simples_c1",
            )
            if superacao_negocial:
                col_neg1, col_neg2 = st.columns(2)
                with col_neg1:
                    percentual_manual_pct = st.number_input(
                        "Percentual aplicado por acordo (%)",
                        min_value=0.0,
                        max_value=100.0,
                        value=round(percentual_indice * 100, 4),
                        step=0.01,
                        format="%.4f",
                        key="percentual_negocial_simples_c1",
                    )
                    data_inicio_efeito_negocial = st.date_input(
                        "Início dos efeitos financeiros por acordo",
                        value=dt_solic,
                        format="DD/MM/YYYY",
                        key="inicio_negocial_simples_c1",
                    )
                with col_neg2:
                    referencia_documental = st.text_input(
                        "Referência documental, se houver",
                        placeholder="Ex.: Despacho, Ata, Ofício ou Nota Técnica",
                        key="referencia_negocial_simples_c1",
                    )
                justificativa_negocial = st.text_area(
                    "Justificativa técnica/negocial",
                    placeholder="Registre a fundamentação da concessão por acordo negocial.",
                    key="justificativa_negocial_simples_c1",
                    height=90,
                )
                if not justificativa_negocial.strip():
                    st.info("A justificativa deve ser preenchida para fins de memória processual antes da instrução final.")
                percentual_aplicado = float(percentual_manual_pct) / 100
                fator_ciclo_efetivo = 1 + percentual_aplicado
                situacao_aplicada = "🟣 CICLO ADMITIDO POR NEGOCIAÇÃO ENTRE AS PARTES"
        if not superacao_negocial:
            percentual_aplicado = 0.0
            fator_ciclo_efetivo = 1.0

    with st.expander("🔍 Memória de Cálculo Detalhada"):
        st.write(f"**Metodologia:** {res['metodo']}")
        if "IST" in tipo_idx:
            _render_equacao_ist(float(res['i_ini']), float(res['i_fim']), float(res['variacao']))
        else:
            st.dataframe(res['dados'])

    st.subheader("Relatório de Apuração")
    if superacao_negocial:
        percentual_aplicado_fmt = f"{percentual_aplicado * 100:,.4f}%".replace('.', ',')
        inicio_negocial_txt = data_inicio_efeito_negocial.strftime('%d/%m/%Y') if data_inicio_efeito_negocial else dt_solic.strftime('%d/%m/%Y')
        relatorio_simples = f"""
        **C1:** Pedido realizado em {dt_solic.strftime('%d/%m/%Y')}. Intervalo do C1: {janela_str}.  
        Janela de Admissibilidade: {janela_adm_str}.  
        **Resultado automático:** {situacao_automatica}.  
        **Tratamento aplicado:** (*) ciclo admitido por negociação entre as partes.  
        Variação apurada pelo índice: {v_fmt}.  
        Percentual aplicado por acordo: {percentual_aplicado_fmt}.  
        Índice {tipo_idx}.  
        Data de início dos efeitos financeiros por acordo: {inicio_negocial_txt}.

        (*) O diagnóstico automático de preclusão foi preservado. O ciclo foi considerado aplicável por decisão negocial registrada pelo usuário.
        """
    else:
        relatorio_simples = f"""
        **C1:** Pedido realizado em {dt_solic.strftime('%d/%m/%Y')}. Intervalo do C1: {janela_str}.  
        Janela de Admissibilidade: {janela_adm_str}.  
        Resultado: {status_ped}.  
        Variação do Ciclo: {v_fmt}. Variação acumulada: {v_fmt}.  
        Índice {tipo_idx}.  
        Data de início dos efeitos financeiros: {dt_solic.strftime('%d/%m/%Y')}.
        """
    st.info(relatorio_simples)

    inicio_efeito_financeiro = data_inicio_efeito_negocial if superacao_negocial and data_inicio_efeito_negocial else (dt_solic if dt_solic >= dt_aniv else dt_aniv)
    fim_efeito_financeiro = inicio_efeito_financeiro + relativedelta(months=11)

    ciclo_unico = {
        'ciclo': 'C1',
        'data_base': dt_base.strftime('%d/%m/%Y'),
        'intervalo_indice': janela_str,
        'janela_admissibilidade': janela_adm_str,
        'data_pedido': dt_solic.strftime('%d/%m/%Y'),
        'situacao': situacao_aplicada,
        'situacao_automatica': situacao_automatica,
        'situacao_aplicada': situacao_aplicada,
        'superacao_negocial': bool(superacao_negocial),
        'percentual_indice': float(percentual_indice),
        'percentual_aplicado': float(percentual_aplicado),
        'justificativa_negocial': justificativa_negocial.strip(),
        'referencia_documental': referencia_documental.strip(),
        'variacao': float(percentual_aplicado),
        'variacao_formatada': f"{percentual_aplicado*100:,.2f}%".replace('.', ','),
        'fator': float(fator_ciclo_efetivo),
        'fator_acumulado': float(fator_ciclo_efetivo),
        'periodo_inicio': _formatar_data(periodo_inicio),
        'periodo_fim': _formatar_data(periodo_fim),
        'financeiro_inicio': _formatar_data(inicio_efeito_financeiro),
        'financeiro_fim': _formatar_data(fim_efeito_financeiro),
    }

    st.session_state['dados_admissibilidade'] = {
        'origem': 'Reajuste Simples',
        'tipo': 'Simples',
        'indice': tipo_idx,
        'data_base_original': dt_base.strftime('%d/%m/%Y'),
        'contexto_contratual_anterior': contexto_contratual,
        'fator': float(fator_ciclo_efetivo),
        'fator_acumulado': float(fator_ciclo_efetivo),
        'variacao': float(percentual_aplicado),
        'variacao_acumulada': float(percentual_aplicado),
        'variacao_acumulada_formatada': f"{percentual_aplicado*100:,.2f}%".replace('.', ','),
        'ciclos': [ciclo_unico],
    }

    arquivo_coleta = gerar_arquivo_coleta_excel(st.session_state['dados_admissibilidade'])
    st.download_button(
        label="📥 Gerar Arquivo de Coleta",
        type="primary",
        data=arquivo_coleta,
        file_name="Coleta_Reajuste_Simples.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )
