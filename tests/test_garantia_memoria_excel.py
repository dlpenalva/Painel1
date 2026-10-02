"""Validação opt-in da memória da Garantia no Microsoft Excel real.

Executar no Windows com Excel instalado:

    RUN_EXCEL_INTEGRATION=1 python -m pytest -q tests/test_garantia_memoria_excel.py

Opcionalmente, ``GARANTIA_MEMORIA_EVIDENCIA_DIR`` preserva o XLSX e a imagem
renderizada pelo próprio Excel para inspeção visual.
"""
from __future__ import annotations

import gc
import os
import time
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pytest

from _garantia_calculo import (
    COLUNA_EVENTO_DATA,
    COLUNA_EVENTO_GARANTIA,
    COLUNA_EVENTO_TIPO,
    COLUNA_EVENTO_VALIDADE,
    COLUNA_EVENTO_VALOR,
    COLUNA_EVENTO_VIGENCIA,
    DIAS_VALIDADE_MINIMA,
    TIPO_REAJUSTE,
    analisar_garantia,
    calcular_situacao_atual,
    normalizar_eventos,
)
from _garantia_memoria import NOME_ABA, gerar_memoria_garantia_xlsx
from _versao import CL8US_VERSION


pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para executar o Excel COM",
)


def test_xlsx_abre_sem_reparo_e_renderiza_no_excel(tmp_path: Path) -> None:
    import pythoncom
    import win32com.client

    registros = [
        {
            COLUNA_EVENTO_TIPO: TIPO_REAJUSTE,
            COLUNA_EVENTO_DATA: date(2027 + indice // 12, indice % 12 + 1, 15),
            COLUNA_EVENTO_VALOR: (
                f"{5_130_000 + indice * 10_000:,}".replace(",", ".") + ",00"
            ),
            COLUNA_EVENTO_VIGENCIA: date(2030, 12, 23),
            COLUNA_EVENTO_GARANTIA: None,
            COLUNA_EVENTO_VALIDADE: None,
        }
        for indice in range(20)
    ]
    eventos, avisos, pendencias = normalizar_eventos(registros)
    assert not avisos and not pendencias
    situacao = calcular_situacao_atual(
        "5.117.331,91", "5", date(2029, 12, 31), eventos,
        "255.866,60", date(2031, 3, 23),
    )
    analise = analisar_garantia(
        situacao["valor_atual"], situacao["percentual"], situacao["vigencia_atual"],
        situacao["garantia_apresentada"], situacao["validade_apresentada"],
    )

    evidencia = Path(os.environ.get("GARANTIA_MEMORIA_EVIDENCIA_DIR", tmp_path))
    evidencia.mkdir(parents=True, exist_ok=True)
    caminho = evidencia / "memoria_calculo_garantia_excel.xlsx"
    imagem = evidencia / "memoria_calculo_garantia_excel.png"
    impressao = evidencia / "memoria_calculo_garantia_impressao.pdf"
    caminho.write_bytes(gerar_memoria_garantia_xlsx(
        situacao,
        analise,
        versao_cl8us=CL8US_VERSION,
        gerado_em=datetime(2026, 10, 1, 14, 30, tzinfo=ZoneInfo("America/Sao_Paulo")),
        dias_validade_minima=DIAS_VALIDADE_MINIMA,
    ))

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    pasta = None
    grafico = None
    try:
        pasta = excel.Workbooks.Open(
            str(caminho.resolve()), UpdateLinks=0, ReadOnly=True, CorruptLoad=0
        )
        assert pasta.Worksheets.Count == 1
        planilha = pasta.Worksheets(NOME_ABA)
        planilha.Activate()
        excel.ActiveWindow.Zoom = 70
        assert planilha.UsedRange.Columns.Count == 9
        assert planilha.Range("A2").Value == "MEMÓRIA DE CÁLCULO DA GARANTIA CONTRATUAL"
        assert planilha.PageSetup.Orientation == 2  # xlLandscape
        assert planilha.PageSetup.FitToPagesWide == 1
        assert not planilha.PageSetup.PrintTitleRows
        assert planilha.Application.WorksheetFunction.CountA(planilha.UsedRange) > 30

        planilha.ExportAsFixedFormat(
            0, str(impressao.resolve()), Quality=0, IncludeDocProperties=True,
            IgnorePrintAreas=False, OpenAfterPublish=False,
        )
        assert impressao.exists() and impressao.stat().st_size > 10_000
        assert planilha.HPageBreaks.Count >= 1

        # A própria renderização do Excel vira imagem para inspeção humana.
        area = planilha.UsedRange
        area.CopyPicture(Appearance=1, Format=2)
        time.sleep(0.8)
        grafico = planilha.ChartObjects().Add(0, 0, area.Width, area.Height)
        grafico.Chart.Paste()
        time.sleep(0.8)
        assert grafico.Chart.Shapes.Count > 0
        assert grafico.Chart.Export(str(imagem.resolve()), "PNG")
        assert imagem.exists() and imagem.stat().st_size > 10_000

        grafico.Delete()
        grafico = None
        area = None
        planilha = None
        pasta.Close(SaveChanges=False)
        pasta = None
    finally:
        if grafico is not None:
            try:
                grafico.Delete()
            except Exception:
                pass
        if pasta is not None:
            pasta.Close(SaveChanges=False)
        excel.Quit()
        excel = None
        gc.collect()
        pythoncom.CoUninitialize()
