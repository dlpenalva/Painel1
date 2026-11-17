from datetime import date
from io import BytesIO

import pytest
from openpyxl import load_workbook

from _conferencias_execucao import aplicar_conferencias, classificar_conferencia
from _coleta_oficial import TEMPLATE_COLETA_OFICIAL, obter_coleta_oficial_bytes
from _motor_reconciliacao import reconciliar_execucoes

P = (date(2025, 9, 1), date(2026, 9, 1))

@pytest.mark.parametrize('financeiro,fisico,completo,pf,pp,just,expected', [
    (100,100,True,P,P,None,'VALIDADO'),
    (0,0,True,P,P,None,'VALIDADO'),
    (.15,.20,True,P,P,None,'VALIDADO'),
    (.15,.21,True,P,P,None,'CONFERIR'),
    (100,120,True,P,P,'Glosa documentada','JUSTIFICADO'),
    (100,120,True,P,P,None,'CONFERIR'),
    (None,0,True,P,P,None,'BASE INSUFICIENTE'),
    (100,100,False,P,P,None,'BASE INSUFICIENTE'),
    (100,100,True,None,P,None,'BASE INSUFICIENTE'),
    (100,120,True,P,(P[0],date(2026,10,15)),None,'COMPARAÇÃO NÃO APLICÁVEL'),
])
def test_periodos(financeiro,fisico,completo,pf,pp,just,expected):
    assert classificar_conferencia(financeiro,fisico,pf,pp,completo,
                                    justificativa=just) == expected

@pytest.mark.parametrize('ciclo',[f'C{n}' for n in range(5)])
def test_motor_sem_cobertura_nao_bloqueia(ciclo):
    leitura = {'vta_sombra':{'parcelas_computadas':[
        {'fonte_parcela':'Financeiro','ciclo':ciclo,'valor':100}]},
        'itens_consumidos_v10':{'itens':[{'ciclo_inferido':ciclo,'valor_total':120}]}}
    r = reconciliar_execucoes(leitura)
    assert not r['bloqueia_formalizacao']
    assert r['registros'][0]['status_reconciliacao'] == 'INCONCLUSIVO'
    assert r['total_computado_prevalencia'] == 100
    leitura['periodos_reconciliacao']={ciclo:{'financeiro':P,'consumo':P,'completo':True}}
    r = reconciliar_execucoes(leitura)
    assert r['registros'][0]['conferencias_periodos'] == ['CONFERIR']
    assert not r['bloqueia_formalizacao']


def test_c0_financeiro_permanece_fonte_prevalente():
    leitura = {
        'vta_sombra': {'parcelas_computadas': [
            {'fonte_parcela': 'Financeiro', 'ciclo': 'C0', 'valor': 783824},
        ]},
        'itens_consumidos_v10': {'itens': [
            {'ciclo_inferido': 'C0', 'valor_total': 790000},
        ]},
        'periodos_reconciliacao': {
            'C0': {'financeiro': P, 'consumo': P, 'completo': True},
        },
    }

    resultado = reconciliar_execucoes(leitura)
    registro = resultado['registros'][0]
    assert registro['fonte_principal'] == 'financeiro'
    assert registro['valor_computado'] == 783824
    assert resultado['total_computado_prevalencia'] == 783824
    assert registro['status_reconciliacao'] == 'RECONCILIADO_COM_ALERTA'
    assert not resultado['bloqueia_formalizacao']


def test_conferencia_informativa_preserva_bloqueios_economicos_oficiais():
    before = load_workbook(TEMPLATE_COLETA_OFICIAL, data_only=False)
    after = load_workbook(BytesIO(obter_coleta_oficial_bytes()), data_only=False)
    formula_vta = before['MEMORIA_RESULTADOS']['B26'].value
    formula_alertas = before['MEMORIA_RESULTADOS']['T48'].value

    assert '$T$48>0,""' in formula_vta
    assert 'COUNTIF(aditivos!$O$2:$O$200,"ALERTA:*")' in formula_alertas
    assert after['MEMORIA_RESULTADOS']['B26'].value == formula_vta
    assert after['MEMORIA_RESULTADOS']['T48'].value == formula_alertas

    aplicar_conferencias(after)
    assert after['MEMORIA_RESULTADOS']['B26'].value == formula_vta
    assert after['MEMORIA_RESULTADOS']['T48'].value == formula_alertas


def test_template_runtime_e_layout():
    before = load_workbook(TEMPLATE_COLETA_OFICIAL)
    after = load_workbook(BytesIO(obter_coleta_oficial_bytes()))
    # Runtime tambem cria abas operacionais existentes; correcao e idempotente.
    shapes = [(s.title,str(s.merged_cells),s.sheet_state) for s in after]
    formulas = {s.title:{c.coordinate:c.value for row in s for c in row
                        if c.data_type=='f'} for s in after}
    aplicar_conferencias(after)
    assert shapes == [(s.title,str(s.merged_cells),s.sheet_state) for s in after]
    assert formulas == {s.title:{c.coordinate:c.value for row in s for c in row
                        if c.data_type=='f'} for s in after}
    for n,col in enumerate(('AC','N','P','R','T')):
        f = after['RESULTADOS_DETALHE'][f'C{73+n}'].value
        assert 'SUM(N(' not in f
        if n == 4:
            assert 'Sem historico quantitativo suficiente' in f
            continue
        assert f'SUMIF(itens_Remanesc!$A$2:$A$201,"<>",itens_Remanesc!${col}$2:${col}$201)' in f
    assert 'VTA_FINAL<MEMORIA_RESULTADOS!B20' in after['RESULTADOS']['D27'].value
    assert 'MEMORIA_RESULTADOS!$W$52' in after['RESULTADOS_DETALHE']['B3'].value
    assert after['MEMORIA_RESULTADOS']['B26'].value == before['MEMORIA_RESULTADOS']['B26'].value
    for s in after:
        for row in s:
            for c in row:
                if c.data_type=='f':
                    assert c.value.count('(')==c.value.count(')'),(s.title,c.coordinate,c.value)


def test_c73_excel_real_ignora_totais_e_itens_vazios(tmp_path):
    import os
    if os.environ.get('RUN_EXCEL_INTEGRATION') != '1':
        pytest.skip('Requer Excel COM para avaliar a formula corrigida')
    client = pytest.importorskip('win32com.client')
    pythoncom = pytest.importorskip('pythoncom')
    path = tmp_path/'conferencia_sintetica.xlsx'
    wb = load_workbook(TEMPLATE_COLETA_OFICIAL)
    for r in range(2,202):
        wb['itens_Remanesc'][f'A{r}'] = None
        wb['itens_Remanesc'][f'AC{r}'] = None
    for r,v in enumerate((600000,118000,65824),2):
        wb['itens_Remanesc'][f'A{r}'] = str(r)
        wb['itens_Remanesc'][f'AC{r}'] = v
    wb['itens_Remanesc']['AC5'] = 783824  # linha de total, sem ITEM
    wb['itens_Remanesc']['AC20'] = 999 # linha vazia com residuo
    wb['MEMORIA_RESULTADOS']['B4'] = 'Financeiro'
    wb.save(path)
    pythoncom.CoInitialize()
    excel = client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    book = None
    try:
        book = excel.Workbooks.Open(str(path))
        excel.CalculateFullRebuild()
        assert book.Worksheets('RESULTADOS_DETALHE').Range('C73').Value == 783824
    finally:
        if book is not None: book.Close(False)
        excel.Quit()
        pythoncom.CoUninitialize()


@pytest.fixture
def caso_sintetico():
    from openpyxl import Workbook
    from datetime import timedelta
    wb = Workbook()
    wb.remove(wb.active)
    for name in ('CONTROLE', 'parametros', 'financeiro', 'itens_Remanesc', 'MEMORIA_RESULTADOS'):
        wb.create_sheet(name)
    wb['CONTROLE']['B2'] = 'C4'
    wb['MEMORIA_RESULTADOS']['B4'] = 'Financeiro'
    for n in range(5):
        p = n+2
        wb['parametros'][f'C{p}'] = date(2020+n, 1, 1)
        wb['parametros'][f'D{p}'] = date(2021+n, 1, 1)-timedelta(days=1)
        wb['parametros'][f'I{p}'] = date(2020+n, 1, 1)
    wb['itens_Remanesc']['A2'] = '1'
    wb['itens_Remanesc']['A3'] = '2'
    return wb


def preencher_ciclo(wb,n,valor=10):
    for k in range(12):
        wb['financeiro'].cell(k+2,1,date(2020+n,k+1,1))
        wb['financeiro'].cell(k+2,2,f'c{n}')
        wb['financeiro'].cell(k+2,3,valor)
    col = ('AC','N','P','R','T')[n]
    wb['itens_Remanesc'][f'{col}2'] = valor*2
    wb['itens_Remanesc'][f'{col}3'] = valor*10
    wb['itens_Remanesc'][f'{col}4'] = 9999  # total sem identificador


@pytest.mark.parametrize('n',range(4))
def test_leitura_ciclos_completos_e_diferencas(caso_sintetico,n):
    from _conferencias_execucao import conferir_workbook
    wb = caso_sintetico
    preencher_ciclo(wb,n)
    antes = conferir_workbook(wb)[n]
    assert antes['status']=='VALIDADO'
    assert antes['financeiro']==antes['fisico']==120
    col = ('AC','N','P','R')[n]
    wb['itens_Remanesc'][f'{col}2'] = 30
    depois = conferir_workbook(wb)[n]
    assert depois['status']=='CONFERIR'
    assert depois['diferenca']==10
    justificado = conferir_workbook(wb,{f'C{n}':'Glosa documentada'})[n]
    assert justificado['status']=='JUSTIFICADO'
    assert not justificado['bloqueia_formalizacao']
    assert justificado['financeiro']==120


def test_zero_e_vazio_nao_sao_equivalentes(caso_sintetico):
    from _conferencias_execucao import conferir_workbook
    wb = caso_sintetico
    preencher_ciclo(wb,0,0)
    assert conferir_workbook(wb)[0]['status']=='VALIDADO'
    wb['financeiro']['C2'] = None
    assert conferir_workbook(wb)[0]['status']=='BASE INSUFICIENTE'


def test_mes_duplicado_nao_comprova_cobertura(caso_sintetico):
    from _conferencias_execucao import conferir_workbook
    wb = caso_sintetico
    preencher_ciclo(wb,0)
    wb['financeiro']['A3'] = wb['financeiro']['A2'].value
    assert conferir_workbook(wb)[0]['status']=='BASE INSUFICIENTE'


def test_periodos_distintos_e_ciclo_sem_fechamento(caso_sintetico):
    from _conferencias_execucao import conferir_workbook
    wb = caso_sintetico
    preencher_ciclo(wb,0)
    wb['parametros']['I3'] = date(2021,1,15)
    assert conferir_workbook(wb)[0]['status']=='COMPARAÇÃO NÃO APLICÁVEL'
    preencher_ciclo(wb,4)
    assert conferir_workbook(wb)[4]['status']=='BASE INSUFICIENTE'
