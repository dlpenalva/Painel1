"""Coleta 12.0 no Excel real — grafias de PC_PAGO e coerencia metodo x fonte.

Mesmo caso de uso da 11.8 (C1 com pedido, C2 PRECLUSO e em execucao, posicao
atual 800 sem fotografia de C2):

* "Não", "NÃO", " nao ", "SIM"... dao exatamente o mesmo VTA, retroativo e
  potencial que "Sim"/"Nao"; o aviso conta as grafias diferentes;
* texto que nao e Sim nem Nao (ex.: "S") e invalido, sem significado inventado;
* metodo com a propria fonte vazia bloqueia VTA e retroativo e explica a causa,
  sem usar a outra fonte; com as duas fontes, o metodo selecionado prevalece.
"""
from __future__ import annotations

import gc
import os
import zipfile
from io import BytesIO
from pathlib import Path

import pytest

from _coleta_oficial import gerar_coleta_oficial_preenchida
from test_coleta_118_excel_real import (
    CORTE, EXEC_C0, EXEC_C1, EXEC_C2, FINANCEIRO, PC, PCS, POTENCIAL, VU_C2,
    _dados, _tentar,
)

pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para executar o Excel COM",
)

VTA_PCS = EXEC_C0 + EXEC_C1 + EXEC_C2 + 800 * VU_C2 + POTENCIAL
MEMORIA = ("B26", "T40", "T39", "T49", "D44", "T54", "T55", "T56", "T58", "T61", "T62",
           "T63", "T64", "T65")


def _preencher(livro, metodo: str, *, pcs: bool, financeiro: bool, grafias=None) -> None:
    from pywintypes import Time

    livro.Worksheets("CONTROLE").Range("B1").Value = metodo
    livro.Worksheets("CONTROLE").Range("B3").Value = Time(CORTE)
    itens = livro.Worksheets("itens_Remanesc")
    itens.Range("A2").Value = "1.1"
    itens.Range("B2").Value = 1000
    itens.Range("C2").Value = 10
    itens.Range("E2").Value = 950
    if pcs:
        aba = livro.Worksheets("itens_PC")
        for linha, (numero, data_pc, valor, pago) in enumerate(PCS, start=2):
            aba.Range(f"A{linha}").Value = numero
            aba.Range(f"B{linha}").Value = Time(data_pc)
            aba.Range(f"D{linha}").Value = valor
            aba.Range(f"G{linha}").Value = (grafias or {}).get(linha, pago)
    if financeiro:
        aba = livro.Worksheets("financeiro")
        for linha in range(2, 74):
            competencia = aba.Range(f"A{linha}").Value
            if hasattr(competencia, "year") and competencia.replace(tzinfo=None) <= CORTE:
                aba.Range(f"C{linha}").Value = 50
    ciclo = livro.Worksheets("CICLO_EM_EXECUCAO")
    ciclo.Range("D5").Value2 = livro.Worksheets("CONTROLE").Range("B3").Value2
    livro.Application.CalculateFullRebuild()
    ciclo.Range("C13").Value = 800


def _ler(livro) -> dict:
    mem = livro.Worksheets("MEMORIA_RESULTADOS")
    rd = livro.Worksheets("RESULTADOS_DETALHE")
    pcs = livro.Worksheets("itens_PC")
    return {
        "mem": {c: mem.Range(c).Value2 for c in MEMORIA},
        "D22": rd.Range("D22").Value2,
        "B3": rd.Range("B3").Value2,
        "J5": rd.Range("J5").Value2,
        "B7": rd.Range("B7").Value2,
        "H14": rd.Range("H14").Value2,
        "C1": livro.Worksheets("CONTROLE").Range("C1").Value2,
        "K": [pcs.Range(f"K{r}").Value2 for r in range(2, 2 + len(PCS))],
        "AD": [pcs.Range(f"AD{r}").Value2 for r in range(2, 2 + len(PCS))],
        "H": [pcs.Range(f"H{r}").Value2 for r in range(2, 2 + len(PCS))],
        "J": [pcs.Range(f"J{r}").Value2 for r in range(2, 2 + len(PCS))],
        "M31": pcs.Range("M31").Value2,
        "M32": pcs.Range("M32").Value2,
    }


def _rodar(tmp_path: Path, nome: str, metodo: str, **opcoes) -> dict:
    import pythoncom
    import win32com.client

    caminho = tmp_path / f"Coleta_12_0_{nome}.xlsx"
    caminho.write_bytes(gerar_coleta_oficial_preenchida(_dados()))
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    livro = None
    try:
        livro = _tentar(pythoncom, lambda: excel.Workbooks.Open(
            str(caminho.resolve()), UpdateLinks=0, ReadOnly=False, CorruptLoad=0
        ))
        _tentar(pythoncom, lambda: _preencher(livro, metodo, **opcoes))
        _tentar(pythoncom, excel.CalculateFullRebuild)
        resultado = _tentar(pythoncom, lambda: _ler(livro))
        _tentar(pythoncom, livro.Save)
        _tentar(pythoncom, lambda: livro.Close(SaveChanges=False))
        livro = None
    finally:
        if livro is not None:
            try:
                livro.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        excel = None
        gc.collect()
        pythoncom.CoUninitialize()
    with zipfile.ZipFile(BytesIO(caminho.read_bytes())) as pacote:
        assert not any(b"repairLoad" in pacote.read(n) for n in pacote.namelist())
    return resultado


def test_grafias_diferentes_de_nao_e_sim_dao_o_mesmo_resultado(tmp_path: Path):
    oficial = _rodar(tmp_path, "oficial", PC, pcs=True, financeiro=False)
    # Linhas: 2 PC-C0 Sim, 3 PC-C1 Sim, 4 PC-C1-NAO "Não", 5 PC-C2 Sim.
    variado = _rodar(tmp_path, "grafias", PC, pcs=True, financeiro=False,
                     grafias={2: "SIM", 3: " sim ", 4: " NÃO ", 5: "Sim"})
    assert oficial["mem"]["B26"] == pytest.approx(VTA_PCS)
    for chave in ("B26", "T40", "T39", "D44", "T58"):
        assert variado["mem"][chave] == pytest.approx(oficial["mem"][chave]), chave
    assert variado["D22"] == pytest.approx(oficial["D22"])
    assert variado["H"] == oficial["H"] and variado["J"] == oficial["J"]
    assert variado["AD"] == ["Sim", "Sim", "Nao", "Sim"]
    assert variado["K"] == ["OK", "OK", "OK", "OK"]
    assert (variado["mem"]["T63"], variado["mem"]["T64"], variado["mem"]["T65"]) == (1, 2, 0)
    assert variado["M31"].startswith("1 valor foi reconhecido como 'Nao'")
    assert "2 valores foram reconhecidos como 'Sim'" in variado["M31"]
    assert oficial["mem"]["T63"] == 1  # "Não" do fixture: grafia diferente de "Nao"


def test_valor_invalido_e_detectado_sem_significado_inventado(tmp_path: Path):
    r = _rodar(tmp_path, "invalido", PC, pcs=True, financeiro=False, grafias={2: "S"})
    assert r["AD"][0] == "INVALIDO"
    assert r["K"][0] == "PC_PAGO_A_CONTRATADA invalido"
    assert r["mem"]["T65"] == 1 and r["mem"]["D44"] >= 1
    assert r["M32"].startswith("1 valor de PC_PAGO_A_CONTRATADA não pôde ser reconhecido")
    assert r["B3"] == "REVISE" and r["J5"] >= 1


def test_financeiro_sem_valores_com_pcs_bloqueia_e_explica(tmp_path: Path):
    r = _rodar(tmp_path, "fin_vazio", FINANCEIRO, pcs=True, financeiro=False)
    m = r["mem"]
    assert (m["T54"], m["T55"], m["T58"]) == (0, len(PCS), 1)
    assert m["B26"] in (None, "") and m["T40"] in (None, "") and r["D22"] in (None, "")
    assert r["B3"] == "REVISE" and r["H14"] == "REVISE"
    assert r["B7"].startswith(
        "REVISE O MÉTODO DE APURAÇÃO. Método selecionado: Financeiro (Mensalidade). "
        "Fonte esperada: aba financeiro, sem valores informados. "
        "Foram encontrados dados em: itens_PC."
    )
    assert r["C1"].startswith("REVISE O MÉTODO: aba financeiro sem valores informados")


def test_pcs_sem_pcs_com_financeiro_bloqueia_sem_usar_o_financeiro(tmp_path: Path):
    r = _rodar(tmp_path, "pc_vazio", PC, pcs=False, financeiro=True)
    m = r["mem"]
    assert m["T55"] == 0 and m["T54"] > 0 and m["T58"] == 1
    assert m["B26"] in (None, "") and r["D22"] in (None, "")
    assert "Fonte esperada: aba itens_PC" in r["B7"] and "financeiro" in r["B7"]


def test_duas_fontes_o_metodo_selecionado_prevalece(tmp_path: Path):
    so_pcs = _rodar(tmp_path, "so_pcs", PC, pcs=True, financeiro=False)
    ambas = _rodar(tmp_path, "ambas", PC, pcs=True, financeiro=True)
    assert ambas["mem"]["T58"] == 2  # aviso, nao bloqueio
    assert ambas["mem"]["B26"] == pytest.approx(so_pcs["mem"]["B26"])
    assert ambas["D22"] == pytest.approx(so_pcs["D22"])
    # A conferencia cruzada entre metodos (MEMORIA!F16, anterior a 12.0) pode
    # gerar pendencia com valores divergentes; o bloqueio de coerencia nao.
    assert ambas["C1"].startswith("CONFERIR: também há dados em financeiro")
    assert "REVISE O MÉTODO" not in ambas["B7"]
