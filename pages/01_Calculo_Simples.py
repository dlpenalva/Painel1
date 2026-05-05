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


def get_ist_local(data_inicio, data_fim):
    try:
        df = pd.read_csv('ist.csv', sep=';', decimal=',', encoding='utf-8-sig')
        df.columns = [str(col).strip().lower() for col in df.columns]
        df['data'] = pd.to_datetime(df['data'], dayfirst=True)
        # IST por número-índice: mês-base do ciclo versus o mesmo mês 12 meses depois.
        # Ex.: data-base 10/2023 => out/2023 a out/2024.
        r_ini = pd.Timestamp(data_inicio.year, data_inicio.month, 1).normalize()
        marco_final = data_inicio + relativedelta(years=1)
        r_fim = pd.Timestamp(marco_final.year, marco_final.month, 1).normalize()
        v_ini = df[df['data'].dt.to_period('M') == r_ini.strftime('%Y-%m')]['indice'].values[0]
        v_fim = df[df['data'].dt.to_period('M') == r_fim.strftime('%Y-%m')]['indice'].values[0]
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


def _render_equacao_ist(i_ini, i_fim, variacao):
    equacao_html = f"""
    <div style=\"background:#F4F6F8;border:1px solid #E1E6EB;border-radius:10px;padding:14px 18px;margin-top:10px;\">
        <div style=\"font-family:Consolas, Monaco, monospace;font-size:1.15rem;line-height:1.8;color:#334155;\">
            <span style=\"color:#0F766E;\">({i_fim:.3f}</span>
            <span style=\"color:#94A3B8;\"> / </span>
            <span style=\"color:#0F766E;\">{i_ini:.3f}</span>
            <span style=\"color:#94A3B8;\">) - 1 = </span>
            <span style=\"color:#B45309;font-weight:600;\">{variacao*100:.4f}%</span>
        </div>
    </div>
    """
    st.markdown(equacao_html, unsafe_allow_html=True)


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
    from openpyxl.worksheet.datavalidation import DataValidation

    header_fill = PatternFill('solid', fgColor='1F4E78')
    header_font = Font(color='FFFFFF', bold=True)
    input_fill = PatternFill('solid', fgColor='FFF2CC')
    calc_fill = PatternFill('solid', fgColor='EDEDED')
    date_fill = PatternFill('solid', fgColor='D9EAD3')
    light_fill = PatternFill('solid', fgColor='EAF2F8')
    no_fill = PatternFill(fill_type=None)
    thin = Side(style='thin', color='D9E2F3')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    no_border = Border()
    money_fmt = 'R$ #,##0.00'
    number_fmt = '#,##0.0000'
    date_fmt = 'DD/MM/YYYY'

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

    if 'PARAMETROS_REAJUSTE' in writer.book.sheetnames:
        ws = writer.book['PARAMETROS_REAJUSTE']
        # Refinamento visual pontual: células de valor em B5/B6 sem bordas pesadas.
        for ref in ('B5', 'B6'):
            ws[ref].border = no_border
        ws.column_dimensions['A'].width = 34
        ws.column_dimensions['B'].width = 36

    if 'FINANCEIRO_MENSAL' in writer.book.sheetnames:
        ws = writer.book['FINANCEIRO_MENSAL']
        # Coluna C é o campo de preenchimento pelo fiscal.
        for row in range(2, 201):
            ws.cell(row=row, column=3).fill = input_fill
            ws.cell(row=row, column=3).alignment = Alignment(horizontal='right', vertical='center')
            ws.cell(row=row, column=3).number_format = money_fmt
        ws['C1'].value = 'Valor pago/faturado (preencher)'
        ws['C1'].font = header_font
        ws['C1'].fill = header_fill
        ws.cell(row=201, column=2).value = 'TOTAL'
        ws.cell(row=201, column=2).font = Font(bold=True)
        ws.cell(row=201, column=2).fill = light_fill
        ws.cell(row=201, column=3).value = '=SUM(C2:C200)'
        ws.cell(row=201, column=3).number_format = money_fmt
        ws.cell(row=201, column=3).font = Font(bold=True)
        ws.cell(row=201, column=3).fill = light_fill

    if 'ITENS_REMANESCENTES' in writer.book.sheetnames:
        ws = writer.book['ITENS_REMANESCENTES']
        # Linha 1: datas de referência. Linha 2: cabeçalhos.
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
        ws.cell(row=1, column=1).value = 'Dados do item e valor original'
        ws.cell(row=1, column=1).fill = light_fill
        ws.cell(row=1, column=1).font = Font(bold=True)
        ws.cell(row=1, column=1).alignment = Alignment(horizontal='center', vertical='center')

        # A, B e C sem cor de preenchimento nos campos de dados. D é cálculo automático em cinza.
        for row in range(3, 201):
            for col in (1, 2, 3):
                ws.cell(row=row, column=col).fill = no_fill
            ws.cell(row=row, column=3).number_format = money_fmt
            ws.cell(row=row, column=4).fill = calc_fill
            ws.cell(row=row, column=4).number_format = money_fmt
            ws.cell(row=row, column=4).value = f'=IF(OR(B{row}="",C{row}=""),"",B{row}*C{row})'

        first_rem_col = 5
        for idx, ciclo in enumerate(ciclos):
            col = first_rem_col + idx
            letra = get_column_letter(col)
            ws.cell(row=1, column=col).value = ciclo.get('data_base', '')
            ws.cell(row=1, column=col).fill = date_fill
            ws.cell(row=1, column=col).font = Font(bold=True)
            ws.cell(row=1, column=col).alignment = Alignment(horizontal='center', vertical='center')
            for row in range(3, 201):
                ws.cell(row=row, column=col).fill = input_fill
            ws.column_dimensions[letra].width = 24

        ws.cell(row=201, column=1).value = 'TOTAL'
        ws.cell(row=201, column=1).font = Font(bold=True)
        ws.cell(row=201, column=1).fill = light_fill
        ws.cell(row=201, column=4).value = '=SUM(D3:D200)'
        ws.cell(row=201, column=4).number_format = money_fmt
        ws.cell(row=201, column=4).font = Font(bold=True)
        ws.cell(row=201, column=4).fill = light_fill
        for col in range(5, 5 + len(ciclos)):
            letra = get_column_letter(col)
            ws.cell(row=201, column=col).value = f'=SUM({letra}3:{letra}200)'
            ws.cell(row=201, column=col).font = Font(bold=True)
            ws.cell(row=201, column=col).fill = light_fill

    if 'CICLOS' in writer.book.sheetnames:
        from openpyxl.utils.datetime import from_excel
        ws = writer.book['CICLOS']
        headers = [str(cell.value).strip() if cell.value is not None else '' for cell in ws[1]]
        for row in range(2, ws.max_row + 1):
            # Converter Data-base para data Excel, quando possível, para fórmulas de ADITIVOS.
            if 'Data-base' in headers:
                col_data = headers.index('Data-base') + 1
                valor = ws.cell(row=row, column=col_data).value
                try:
                    data = pd.to_datetime(valor, dayfirst=True).to_pydatetime()
                    ws.cell(row=row, column=col_data).value = data
                    ws.cell(row=row, column=col_data).number_format = date_fmt
                except Exception:
                    pass
            for idx_header, nome_header in enumerate(headers, start=1):
                if nome_header in ('Fator', 'Fator acumulado', 'Fator acumulado efetivo', 'Fator ciclo efetivo'):
                    ws.cell(row=row, column=idx_header).number_format = number_fmt
            if 'Tratamento financeiro do ciclo' in headers:
                col_trat = headers.index('Tratamento financeiro do ciclo') + 1
                ws.cell(row=row, column=col_trat).fill = input_fill
                ws.cell(row=row, column=col_trat).alignment = Alignment(horizontal='center', vertical='center')
        if 'Tratamento financeiro do ciclo' in headers:
            col_trat = headers.index('Tratamento financeiro do ciclo') + 1
            letra_trat = get_column_letter(col_trat)
            dv = DataValidation(
                type='list',
                formula1='"A apurar,Já concedido,Precluso,Sem efeito financeiro"',
                allow_blank=False,
            )
            ws.add_data_validation(dv)
            dv.add(f'{letra_trat}2:{letra_trat}{ws.max_row}')
            ws.column_dimensions[letra_trat].width = 28
        # Remover bordas do bloco dinâmico I:L nas linhas de ciclos.
        for row in range(2, ws.max_row + 1):
            for col in range(9, min(12, ws.max_column) + 1):
                ws.cell(row=row, column=col).border = no_border

    if 'ADITIVOS_QUANTITATIVOS' in writer.book.sheetnames:
        ws = writer.book['ADITIVOS_QUANTITATIVOS']
        # Estrutura: A Item | B Data | C Ciclo | D Tipo | E Qtd | F VU | G Valor Original | H Aplicar? | I Fator | J Valor Atualizado
        for row in range(2, 201):
            # Gerais em branco; C, G, I e J como calculadas em cinza claro.
            for col in range(1, 11):
                ws.cell(row=row, column=col).fill = no_fill
            ws.cell(row=row, column=2).number_format = date_fmt
            ws.cell(row=row, column=6).number_format = money_fmt
            ws.cell(row=row, column=7).number_format = money_fmt
            ws.cell(row=row, column=9).number_format = number_fmt
            ws.cell(row=row, column=10).number_format = money_fmt
            ws.cell(row=row, column=3).fill = calc_fill
            ws.cell(row=row, column=7).fill = calc_fill
            ws.cell(row=row, column=9).fill = calc_fill
            ws.cell(row=row, column=10).fill = calc_fill
            ws.cell(row=row, column=3).value = f'=IF(B{row}="","",LOOKUP(B{row},CICLOS!$B$2:$B$200,CICLOS!$A$2:$A$200))'
            ws.cell(row=row, column=7).value = f'=IF(OR(E{row}="",F{row}=""),"",E{row}*F{row})'
            ws.cell(row=row, column=9).value = f'=IF(C{row}="","",IFERROR(VLOOKUP(C{row},CICLOS!$A$2:$K$200,11,FALSE),1))'
            ws.cell(row=row, column=10).value = f'=IF(G{row}="","",IF(D{row}="Supressão",-1,1)*IF(H{row}="Sim",G{row}*I{row},G{row}))'
        ws.cell(row=201, column=1).value = 'TOTAL'
        ws.cell(row=201, column=1).font = Font(bold=True)
        ws.cell(row=201, column=10).value = '=SUM(J2:J200)'
        ws.cell(row=201, column=10).number_format = money_fmt
        ws.cell(row=201, column=10).font = Font(bold=True)
        ws.cell(row=201, column=10).fill = light_fill

        dv_tipo = DataValidation(type='list', formula1='"Acréscimo,Supressão"', allow_blank=True)
        dv_sim_nao = DataValidation(type='list', formula1='"Sim,Não"', allow_blank=True)
        ws.add_data_validation(dv_tipo)
        ws.add_data_validation(dv_sim_nao)
        dv_tipo.add('D2:D200')
        dv_sim_nao.add('H2:H200')
        ws.column_dimensions['A'].width = 16
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 16
        ws.column_dimensions['D'].width = 18
        ws.column_dimensions['E'].width = 24
        ws.column_dimensions['F'].width = 22
        ws.column_dimensions['G'].width = 24
        ws.column_dimensions['H'].width = 28
        ws.column_dimensions['I'].width = 22
        ws.column_dimensions['J'].width = 26


def gerar_arquivo_coleta_excel(dados_admissibilidade):
    ciclos = dados_admissibilidade.get('ciclos', [])

    parametros = pd.DataFrame([
        {'Campo': 'Origem da análise', 'Valor': dados_admissibilidade.get('origem', dados_admissibilidade.get('tipo', ''))},
        {'Campo': 'Índice utilizado', 'Valor': dados_admissibilidade.get('indice', '')},
        {'Campo': 'Data-base original', 'Valor': dados_admissibilidade.get('data_base_original', '')},
        {'Campo': 'Quantidade de ciclos', 'Valor': len(ciclos)},
        {'Campo': 'Variação acumulada final', 'Valor': dados_admissibilidade.get('variacao_acumulada_formatada', '')},
        {'Campo': 'Fator acumulado total', 'Valor': round(float(dados_admissibilidade.get('fator_acumulado', dados_admissibilidade.get('fator', 1.0))), 4)},
        {'Campo': 'Data de geração do arquivo', 'Valor': datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')},
    ])

    df_ciclos = pd.DataFrame([
        {
            'Ciclo': c.get('ciclo', ''),
            'Data-base': c.get('data_base', ''),
            'Intervalo do índice': c.get('intervalo_indice', c.get('Janela', '')),
            'Janela de admissibilidade': c.get('janela_admissibilidade', c.get('JanelaAdm', '')),
            'Data do pedido': c.get('data_pedido', c.get('Pedido', '')),
            'Início financeiro': c.get('financeiro_inicio', ''),
            'Fim financeiro': c.get('financeiro_fim', ''),
            'Situação': c.get('situacao', c.get('Situação', '')),
            'Variação': c.get('variacao_formatada', c.get('Variação', '')),
            'Fator': round(float(c.get('fator', 1.0)), 4),
            'Fator acumulado': round(float(c.get('fator_acumulado', 1.0)), 4),
            'Tratamento financeiro do ciclo': c.get('tratamento_financeiro', 'Precluso' if 'PRECLUSO' in str(c.get('situacao', '')).upper() else 'A apurar'),
        }
        for c in ciclos
    ])

    # FINANCEIRO_MENSAL: cabeçalho na linha 1, linhas 2:200 formatadas, total na linha 201.
    linhas_financeiro = []
    for c in ciclos:
        inicio_financeiro = c.get('periodo_inicio') or c.get('financeiro_inicio')
        fim_financeiro = c.get('periodo_fim') or c.get('financeiro_fim')
        for competencia in _competencias_mensais(inicio_financeiro, fim_financeiro):
            linhas_financeiro.append({
                'Ciclo': c.get('ciclo', ''),
                'Competência': competencia,
                'Valor pago/faturado': None,
            })
    while len(linhas_financeiro) < 199:
        linhas_financeiro.append({'Ciclo': None, 'Competência': None, 'Valor pago/faturado': None})
    linhas_financeiro = linhas_financeiro[:199]
    df_financeiro = pd.DataFrame(linhas_financeiro, columns=['Ciclo', 'Competência', 'Valor pago/faturado'])

    # ITENS_REMANESCENTES: linha 1 datas, linha 2 cabeçalhos, linhas 3:200 formatadas, total na linha 201.
    colunas_remanescentes = [f"Remanescente início {c.get('ciclo', '')}" for c in ciclos]
    colunas_itens = ['Item', 'Quantidade contratada', 'Valor unitário original', 'Valor total', *colunas_remanescentes]
    linhas_itens = [{col: None for col in colunas_itens} for _ in range(198)]
    df_itens = pd.DataFrame(linhas_itens, columns=colunas_itens)

    df_aditivos = pd.DataFrame([
        {
            'Item': None,
            'Data do aditivo': None,
            'Ciclo/Marco': None,
            'Tipo de alteração': None,
            'Quantidade acrescida/suprimida': None,
            'Valor unitário original': None,
            'Valor original da alteração': None,
            'Aplicar reajuste acumulado? (Sim/Não)': 'Sim',
            'Fator acumulado aplicável': None,
            'Valor atualizado da alteração': None,
        }
        for _ in range(199)
    ])

    try:
        output = BytesIO()
    except NameError:
        output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        parametros.to_excel(writer, sheet_name='PARAMETROS_REAJUSTE', index=False)
        df_ciclos.to_excel(writer, sheet_name='CICLOS', index=False)
        df_financeiro.to_excel(writer, sheet_name='FINANCEIRO_MENSAL', index=False)
        df_itens.to_excel(writer, sheet_name='ITENS_REMANESCENTES', index=False, startrow=1)
        df_aditivos.to_excel(writer, sheet_name='ADITIVOS_QUANTITATIVOS', index=False)
        _aplicar_estilos_coleta(writer, ciclos)
        _ajustar_larguras(writer, [
            'PARAMETROS_REAJUSTE', 'CICLOS', 'FINANCEIRO_MENSAL', 'ITENS_REMANESCENTES', 'ADITIVOS_QUANTITATIVOS'
        ])
    output.seek(0)
    return output

st.image("https://www.telebras.com.br/wp-content/uploads/2019/06/Telebras_Logo_AzulProfundo.png", width=250)
st.title("Reajuste Simples")

col1, col2 = st.columns(2)
with col1:
    dt_base = st.date_input("Data-Base Anterior:", value=datetime(2023, 8, 2), format="DD/MM/YYYY")
    dt_solic = st.date_input("Data do Pedido:", value=datetime(2024, 4, 9), format="DD/MM/YYYY")
with col2:
    tipo_idx = st.selectbox("Índice:", ["IST (Série Local)", "IPCA (433)", "IGP-M (189)"])

chave_analise_simples = (
    dt_base.isoformat() if hasattr(dt_base, "isoformat") else str(dt_base),
    dt_solic.isoformat() if hasattr(dt_solic, "isoformat") else str(dt_solic),
    tipo_idx,
)

st.info("Informe a data-base, a data do pedido e o índice. Em seguida, clique em **Processar Análise**.")
if st.button("Processar Análise", type="primary", use_container_width=False):
    st.session_state["processar_reajuste_simples_key"] = chave_analise_simples

if st.session_state.get("processar_reajuste_simples_key") != chave_analise_simples:
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

    with st.expander("🔍 Memória de Cálculo Detalhada"):
        st.write(f"**Metodologia:** {res['metodo']}")
        if "IST" in tipo_idx:
            _render_equacao_ist(float(res['i_ini']), float(res['i_fim']), float(res['variacao']))
        else:
            st.dataframe(res['dados'])

    st.subheader("Relatório de Apuração")
    relatorio_simples = f"""
    **C1:** Pedido realizado em {dt_solic.strftime('%d/%m/%Y')}. Intervalo do C1: {janela_str}.  
    Janela de Admissibilidade: {janela_adm_str}.  
    Resultado: {status_ped}.  
    Variação do Ciclo: {v_fmt}. Variação acumulada: {v_fmt}.  
    Índice {tipo_idx}.  
    Data de início dos efeitos financeiros: {dt_solic.strftime('%d/%m/%Y')}.
    """
    st.info(relatorio_simples)

    inicio_efeito_financeiro = dt_solic if dt_solic >= dt_aniv else dt_aniv
    fim_efeito_financeiro = inicio_efeito_financeiro + relativedelta(months=11)

    ciclo_unico = {
        'ciclo': 'C1',
        'data_base': dt_base.strftime('%d/%m/%Y'),
        'intervalo_indice': janela_str,
        'janela_admissibilidade': janela_adm_str,
        'data_pedido': dt_solic.strftime('%d/%m/%Y'),
        'situacao': status_ped,
        'variacao': float(res['variacao']),
        'variacao_formatada': v_fmt,
        'fator': float(fator_ciclo),
        'fator_acumulado': float(fator_ciclo),
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
        'fator': float(fator_ciclo),
        'fator_acumulado': float(fator_ciclo),
        'variacao': float(res['variacao']),
        'variacao_acumulada': float(res['variacao']),
        'variacao_acumulada_formatada': v_fmt,
        'ciclos': [ciclo_unico],
    }

    arquivo_coleta = gerar_arquivo_coleta_excel(st.session_state['dados_admissibilidade'])
    st.download_button(
        label="📥 Gerar Arquivo de Coleta",
        data=arquivo_coleta,
        file_name="Coleta_Reajuste_Simples.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )
