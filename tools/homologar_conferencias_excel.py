"""Homologa uma entrada em copias temporarias, com Excel COM instalado.

python tools/homologar_conferencias_excel.py /caminho/entrada.xlsx
Nao altera a entrada, nao cria um segundo calculo do VTA e nao faz publicacao.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys
import tempfile

sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from openpyxl import load_workbook
from tools.aplicar_conferencias_execucao import aplicar


def homologar(entrada: Path):
    try:
        import pythoncom
        from win32com.client import DispatchEx
    except ImportError as exc:
        raise RuntimeError('Homologacao requer Windows, Microsoft Excel e pywin32. '
                           'Nenhuma entrada foi alterada.') from exc
    digest = hashlib.sha256(entrada.read_bytes()).hexdigest()
    caches = load_workbook(entrada,data_only=True)
    vta_cache = caches['MEMORIA_RESULTADOS']['B26'].value
    caches.close()
    with tempfile.TemporaryDirectory(prefix='cl8us-homologacao-') as directory:
        antes = Path(directory)/'antes.xlsx'
        depois = Path(directory)/'depois.xlsx'
        shutil.copyfile(entrada,antes)
        shutil.copyfile(entrada,depois)
        aplicar(depois)
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = DispatchEx('Excel.Application')
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.AskToUpdateLinks = False
            excel.AutomationSecurity = 3  # desabilitar macros das copias
            valores = []
            for path in (antes,depois):
                book = excel.Workbooks.Open(str(path),UpdateLinks=0,ReadOnly=False)
                try:
                    excel.CalculateFullRebuild()
                    values = {
                        'vta':book.Worksheets('MEMORIA_RESULTADOS').Range('B26').Value,
                        'fisico_c0':book.Worksheets('RESULTADOS_DETALHE').Range('C73').Value,
                        'status':book.Worksheets('RESULTADOS').Range('G6').Value,
                        'conferencias':list(book.Worksheets('RESULTADOS_DETALHE').Range('E73:E77').Value),
                    }
                    book.Save()
                finally:
                    book.Close(False)
                # Reabrir a copia calculada: verifica persistencia dos caches.
                reopen = load_workbook(path,data_only=True)
                assert reopen['MEMORIA_RESULTADOS']['B26'].value == values['vta']
                reopen.close()
                valores.append(values)
        finally:
            if excel is not None: excel.Quit()
            pythoncom.CoUninitialize()
    if hashlib.sha256(entrada.read_bytes()).hexdigest()!=digest:
        raise RuntimeError('Hash da entrada mudou durante a homologacao.')
    before,after = valores
    if not all(isinstance(v,(int,float)) for v in (vta_cache,before['vta'],after['vta'])):
        raise RuntimeError('VTA indisponivel em uma das etapas: investigar antes de homologar.')
    result = {'cache_entrada':vta_cache,'antes_recalculado':before,
              'depois_recalculado':after,'entrada_preservada':True,
              'vta_preservado':abs(before['vta']-after['vta'])<.005,
              'cache_confere':abs(vta_cache-before['vta'])<.005}
    print(json.dumps(result,ensure_ascii=False,indent=2,default=str))
    if not result['vta_preservado'] or not result['cache_confere']:
        raise RuntimeError('Valores oficiais diferem: homologacao reprovada; investigar o relatorio.')
    return result


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('entrada',type=Path)
    args=parser.parse_args()
    try:
        homologar(args.entrada.resolve(strict=True))
    except (RuntimeError,OSError) as exc:
        parser.exit(1,str(exc)+'\n')
