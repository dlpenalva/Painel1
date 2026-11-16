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
    for r,v in enumerate((600,118,22),2):
        wb['itens_Remanesc'][f'A{r}'] = str(r)
        wb['itens_Remanesc'][f'AC{r}'] = v
    wb['itens_Remanesc']['AC5'] = 740  # linha de total, sem ITEM
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
        assert book.Worksheets('RESULTADOS_DETALHE').Range('C73').Value == 740
    finally:
        if book is not None: book.Close(False)
        excel.Quit()
        pythoncom.CoUninitialize()
