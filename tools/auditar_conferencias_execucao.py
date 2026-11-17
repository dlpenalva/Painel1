"""Uso: python tools/auditar_conferencias_execucao.py /caminho/caso.xlsx.

Emite JSON sem modificar a entrada. Nunca adicionar a entrada real ao Git.
"""
from pathlib import Path
import sys
import json
import warnings
from decimal import Decimal
from openpyxl import load_workbook
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from _conferencias_execucao import conferir_workbook

if __name__=='__main__':
    warnings.simplefilter('ignore',UserWarning)
    wb=load_workbook(sys.argv[1],data_only=True)
    result={'conferencias':conferir_workbook(wb)}
    cells={'contrato':('MEMORIA_RESULTADOS','B20'),
           'vta_oficial':('MEMORIA_RESULTADOS','B26'),
           'financeiro_total':('financeiro','C74'),
           'fisico_c73_cache':('RESULTADOS_DETALHE','C73'),
           'decomposicao_abertura':('MEMORIA_RESULTADOS','W48'),
           'decomposicao_atual':('MEMORIA_RESULTADOS','W50'),
           'diferenca_decomposicoes':('MEMORIA_RESULTADOS','W51')}
    result['valores_cache']={k:wb[s][a].value for k,(s,a) in cells.items()}
    result['nota']='Valores oficiais sao caches do arquivo; requer Excel para homologar o recalculo.'
    print(json.dumps(result,ensure_ascii=False,indent=2,default=str))
