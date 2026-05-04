import streamlit as st
import pandas as pd
import requests
from datetime import datetime
from zoneinfo import ZoneInfo
from dateutil.relativedelta import relativedelta
from io import BytesIO

st.set_page_config(page_title="Cálculo Simples", layout="wide")


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
        r_ini = (data_inicio - relativedelta(months=1)).replace(day=1)
        r_fim = data_fim.replace(day=1)
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
    ciclos = dados_admissibilidade.get('ciclos', [])

    parametros = pd.DataFrame([
        {'Campo': 'Origem da análise', 'Valor': dados_admissibilidade.get('origem', '')},
        {'Campo': 'Índice utilizado', 'Valor': dados_admissibilidade.get('indice', '')},
        {'Campo': 'Data-base original', 'Valor': dados_admissibilidade.get('data_base_original', '')},
        {'Campo': 'Quantidade de ciclos', 'Valor': len(ciclos)},
        {'Campo': 'Fator acumulado final', 'Valor': round(float(dados_admissibilidade.get('fator_acumulado', 1.0)), 4)},
        {'Campo': 'Variação acumulada final', 'Valor': dados_admissibilidade.get('variacao_acumulada_formatada', '')},
        {'Campo': 'Data de geração do arquivo', 'Valor': datetime.now(ZoneInfo('America/Sao_Paulo')).strftime('%d/%m/%Y %H:%M')},
    ])

    df_ciclos = pd.DataFrame([
        {
            'Ciclo': c.get('ciclo', ''),
            'Data-base': c.get('data_base', ''),
            'Intervalo do índice': c.get('intervalo_indice', ''),
            'Janela de admissibilidade': c.get('janela_admissibilidade', ''),
            'Data do pedido': c.get('data_pedido', ''),
            'Início financeiro': c.get('financeiro_inicio', ''),
            'Fim financeiro': c.get('financeiro_fim', ''),
            'Situação': c.get('situacao', ''),
            'Variação': c.get('variacao_formatada', ''),
            'Fator': round(float(c.get('fator', 1.0)), 4),
            'Fator acumulado': round(float(c.get('fator_acumulado', 1.0)), 4),
        }
        for c in ciclos
    ])

    # Aba financeira simplificada: somente competências com potencial efeito financeiro.
    # Não há linha TOTAL C0: o executado em C0 é inferido pela diferença entre
    # o valor original do contrato e o remanescente informado no início de C1.
    linhas_financeiro = []
    for c in ciclos:
        inicio_financeiro = c.get('financeiro_inicio') or c.get('periodo_inicio')
        fim_financeiro = c.get('financeiro_fim') or c.get('periodo_fim')
        for competencia in _competencias_mensais(inicio_financeiro, fim_financeiro):
            linhas_financeiro.append({
                'Ciclo': c.get('ciclo', ''),
                'Competência': competencia,
                'Valor pago/faturado': None,
            })
    df_financeiro = pd.DataFrame(linhas_financeiro, columns=[
        'Ciclo', 'Competência', 'Valor pago/faturado'
    ])

    # Aba de remanescentes simplificada. A linha 1 recebe a data-base de referência de cada ciclo.
    colunas_remanescentes = [f"Remanescente início {c.get('ciclo', '')}" for c in ciclos]
    colunas_itens = [
        'Item', 'Quantidade contratada', 'Valor unitário original', 'Valor total',
        *colunas_remanescentes
    ]
    linhas_itens = [{col: None for col in colunas_itens} for _ in range(30)]
    df_itens = pd.DataFrame(linhas_itens, columns=colunas_itens)

    # Aba opcional para acréscimos quantitativos/aditivos posteriores ao reajuste.
    # Usar apenas quando houver aditivo quantitativo a compor o valor para publicação.
    df_aditivos = pd.DataFrame([
        {
            'Descrição do aditivo': None,
            'Data do aditivo': None,
            'Ciclo/Marco': None,
            'Valor original do acréscimo': None,
            'Aplicar reajuste acumulado? (Sim/Não)': 'Sim',
            'Observação': None,
        }
        for _ in range(10)
    ])

    output = BytesIO()
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
            st.code(f"({res['i_fim']} / {res['i_ini']}) - 1 = {res['variacao']*100:.4f}%")
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
