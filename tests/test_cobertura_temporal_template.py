"""Etapa — aba de diagnostico "cobertura_temporal" no template oficial.

Testes estaticos (openpyxl) validam a estrutura da aba (marcos / cobertura /
decisao), o reuso do painel homologado de posicao_referencia, o esquema de
cores reutilizado, a categoria nova de PROJECAO e a INVARIANCIA do VTA oficial
(RESULTADOS!B23/B25/B26). O teste de integracao (RUN_EXCEL_INTEGRATION=1)
recalcula no Excel real um cenario e confere o modo temporal e as datas.
"""
from __future__ import annotations

import gc
import os
import shutil
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
LEGADO = ROOT / "templates" / "Coleta_Reajuste.xlsx"
ABA = "cobertura_temporal"

pytest_com = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para executar o Excel COM",
)


def _wb():
    return load_workbook(TEMPLATE, data_only=False)


# ---------------------------------------------------------------- estaticos

def test_aba_existe_e_visivel():
    wb = _wb()
    assert ABA in wb.sheetnames
    assert wb[ABA].sheet_state == "visible"


def test_titulo_e_banners():
    ws = _wb()[ABA]
    assert "COBERTURA TEMPORAL" in str(ws["A1"].value)
    assert "NAO altera o VTA" in str(ws["A1"].value)
    assert "BLOCO A" in str(ws["A3"].value)
    assert "BLOCO B" in str(ws["A10"].value)
    assert "BLOCO C" in str(ws["A17"].value)
    # cabecalho reutilizado (azul FF1F4E79).
    assert ws["A1"].fill.fgColor.rgb == "FF1F4E79"


def test_reusa_painel_posicao_referencia():
    ws = _wb()[ABA]
    assert ws["B7"].value == "=posicao_referencia!$I$6"      # abertura do corte
    assert ws["B8"].value == '=IF(posicao_referencia!$I$2,posicao_referencia!$I$5,"")'
    assert ws["B11"].value == "=posicao_referencia!$I$5"     # posicao fisica ate
    assert ws["B21"].value == "=posicao_referencia!$I$5"     # observada (data)
    assert ws["B22"].value == "=posicao_referencia!$I$8"     # observada (origem)


def test_fronteiras_financeiro_e_pc():
    ws = _wb()[ABA]
    assert "MAX(financeiro!$A$2:$A$200)" in ws["B12"].value   # financeiro ate
    assert "MAX(itens_PC!$B$2:$B$200)" in ws["B13"].value     # PC ate (DATA_PC)


def test_modo_temporal_seis_estados():
    f = _wb()[ABA]["B18"].value
    for estado in ("POSICAO_ATUAL", "HIBRIDO_TEMPORAL", "FINANCEIRO_POSTERIOR",
                   "PC_POSTERIOR", "POSICAO_DE_CORTE"):
        assert estado in f
    assert "posicao_referencia!$I$2" in f   # completa reutiliza o painel


def test_entrada_gcc_unica_amarela():
    ws = _wb()[ABA]
    assert ws["B4"].fill.fgColor.rgb == "FFFEF9C3"           # amarelo de entrada
    # nao ha entrada fiscal nova: apenas B4 (GCC) e amarela na coluna B.
    amarelas = [r for r in range(4, 24) if ws.cell(r, 2).fill.fgColor.rgb == "FFFEF9C3"]
    assert amarelas == [4]


def test_projecao_categoria_nova_laranja():
    ws = _wb()[ABA]
    assert ws["B15"].fill.fgColor.rgb == "FFFCE4D6"          # projecao a partir de
    assert ws["B23"].fill.fgColor.rgb == "FFFCE4D6"          # posicao projetada
    assert "nao cria retroativo" in ws["B23"].value


def test_legenda_quatro_categorias():
    ws = _wb()[ABA]
    assert "LEGENDA" in str(ws["A25"].value)
    rotulos = [str(ws.cell(r, 1).value) for r in range(26, 30)]
    assert rotulos == ["FISCAL", "GCC", "AUTOMATICO", "PROJECAO"]


def test_sem_today_now_hoje():
    ws = _wb()[ABA]
    textos = [str(c.value) for row in ws.iter_rows() for c in row if isinstance(c.value, str)]
    for f in textos:
        for proib in ("TODAY(", "NOW(", "HOJE(", "AGORA("):
            assert proib not in f.upper()


def test_vta_oficial_invariante():
    wb = _wb()
    r = wb["RESULTADOS"]
    assert r["B23"].value == '=IF(OR(B20="",B21="",B22=""),"",ROUND(B20+B21+B22,2))'
    assert "$N$263" in r["B26"].value and ABA not in str(r["B26"].value)
    assert r["B25"].value in (None, "")


def test_template_legado_intacto():
    if not LEGADO.exists():
        pytest.skip("legado ausente")
    assert ABA not in load_workbook(LEGADO, data_only=False).sheetnames


# ---------------------------------------------------------------- COM

def _recalc(caminho: Path, inspecionar):
    import time
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    for a, v in (("Visible", False), ("DisplayAlerts", False)):
        try:
            setattr(excel, a, v)
        except AttributeError:
            pass
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho.resolve()), UpdateLinks=0)
        excel.CalculateFullRebuild()
        time.sleep(0.3)
        return inspecionar(wb)
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        gc.collect()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


@pytest_com
def test_com_pc_posterior_nao_move_fisica(tmp_path):
    """Foto marco (fallback) + PC maio -> PC_POSTERIOR; fisica continua marco."""
    dest = tmp_path / "COLETA_REAJUSTE_OFICIAL.xlsx"
    shutil.copy2(TEMPLATE, dest)
    wb = load_workbook(dest, data_only=False)
    p = wb["parametros"]
    for num, ini, fim in ((0, date(2023, 1, 1), date(2023, 12, 31)),
                          (1, date(2024, 1, 1), date(2024, 12, 31)),
                          (2, date(2025, 1, 1), date(2025, 12, 31)),
                          (3, date(2026, 1, 1), date(2026, 12, 31))):
        p.cell(num + 2, 3).value = ini
        p.cell(num + 2, 4).value = fim
    ir = wb["itens_Remanesc"]
    ir["A2"], ir["B2"], ir["C2"] = "ITEM-1", 100.0, 10.0
    ir["E2"] = ir["G2"] = ir["I2"] = 100.0    # fotografia C1,C2,C3
    wb["CONTROLE"]["B2"] = "C3"
    pc = wb["itens_PC"]
    pc["A2"], pc["B2"], pc["D2"] = "PC-1", date(2026, 5, 20), 2000.0
    wb.save(dest)

    def ler(w):
        cv = w.Worksheets(ABA)
        r = w.Worksheets("RESULTADOS")
        return {
            "modo": cv.Range("B18").Value,
            "fisica": cv.Range("B11").Value,
            "pc_ate": cv.Range("B13").Value,
            "b23": r.Range("B23").Value,
            "b26": r.Range("B26").Value,
        }

    d = _recalc(dest, ler)
    assert d["modo"] == "PC_POSTERIOR"
    # posicao fisica = abertura C3 (jan/2026), NAO maio.
    assert (d["fisica"].year, d["fisica"].month) == (2026, 1)
    assert (d["pc_ate"].year, d["pc_ate"].month) == (2026, 5)
