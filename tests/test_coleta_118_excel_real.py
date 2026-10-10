"""Coleta 11.8 no Excel real — VTA com a posicao atual e pendencias.

Caso de uso: so C1 tem pedido (3,81%); C2 e PRECLUSO | SEM PEDIDO e e o ciclo
em execucao. Item 1.1 (base 1000, VU 10) tem a fotografia de C1 (950) e
NENHUMA de C2; a posicao atual (CICLO_EM_EXECUCAO) e 800 em 31/08/2025.

* PCs sem fotografia do ciclo vigente: o VTA sai da execucao historica pelos
  PCs (C0, C1 e C2 ate a posicao) + remanescente atual (800 x 10,38), sem
  copiar a posicao para o residual de C2 e sem pendencias;
* com a fotografia de C2 o calculo e o da 11.7 (T49 = 0, T50 = 0);
* PC_PAGO "Não" (com acento) e PC de ciclo precluso nao geram pendencia;
* Financeiro sem fotografia: remanescente oficial = posicao atual;
* ALERTA em aditivos continua esvaziando o VTA.
"""
from __future__ import annotations

import gc
import os
import time
import zipfile
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pytest

from _coleta_oficial import gerar_coleta_oficial_preenchida


pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para executar o Excel COM",
)

CORTE = datetime(2025, 8, 31)
PRECLUSO = "❌ PRECLUSO | SEM PEDIDO NESTE CICLO"
PC = "PC (Pedidos de Compra)"
FINANCEIRO = "Financeiro (Mensalidade)"
PCS = (
    ("PC-C0", datetime(2023, 6, 10), 500.0, "Sim"),
    ("PC-C1", datetime(2024, 6, 10), 300.0, "Sim"),
    ("PC-C1-NAO", datetime(2024, 7, 10), 100.0, "Não"),
    ("PC-C2", datetime(2025, 3, 10), 200.0, "Sim"),
)
MEMORIA = ("T21", "T22", "T23", "T25", "T26", "T39", "T40", "T48", "T49", "T50",
           "T51", "D35", "E26", "F16", "D44", "W49", "W52", "B26")

VU_C2 = 10.38  # 10 x 1,0381 (C1 carregado; C2 precluso = 0%)
EXEC_C0 = 500.0
EXEC_C1 = 400.0 + 11.43  # PC-C1 pago com efeito (300 x 3,81%) + PC-C1-NAO nominal
EXEC_C2 = 200.0  # PC-C2: ciclo precluso, sem efeito proprio
POTENCIAL = 3.81  # PC-C1-NAO: 100 x 3,81%, nao pago


def _dados() -> dict:
    return {
        "origem": "Validacao Coleta 11.8",
        "indice": "IST",
        "data_base_original": "01/01/2023",
        "data_corte": CORTE.date(),
        "ciclos": [
            {
                "ciclo": "C1",
                "data_inicio": date(2024, 1, 1),
                "data_fim": date(2024, 12, 31),
                "data_pedido": date(2024, 1, 1),
                "financeiro_inicio": date(2024, 1, 1),
                "percentual_aplicado": 0.0381,
                "situacao": "✅ TEMPESTIVO",
                "objeto_analise_atual": True,
            },
            {
                "ciclo": "C2",
                "data_inicio": date(2025, 1, 1),
                "data_fim": date(2025, 12, 31),
                "percentual_aplicado": 0.0,
                "situacao": PRECLUSO,
                "objeto_analise_atual": True,
            },
        ],
    }


def _tentar(pythoncom, acao):
    ultimo = None
    for _ in range(30):
        try:
            return acao()
        except Exception as exc:  # pragma: no cover - somente Excel COM
            ultimo = exc
            codigo = getattr(exc, "hresult", None)
            if codigo is None and getattr(exc, "args", ()):
                codigo = exc.args[0]
            if codigo != -2147418111:
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(0.2)
    raise ultimo


def _preencher(livro, metodo: str, *, fotografia_c2, alerta_aditivo: bool,
               posicao_incompleta: bool = False) -> None:
    from pywintypes import Time

    livro.Worksheets("CONTROLE").Range("B1").Value = metodo
    livro.Worksheets("CONTROLE").Range("B3").Value = Time(CORTE)
    itens = livro.Worksheets("itens_Remanesc")
    itens.Range("A2").Value = "1.1"
    itens.Range("B2").Value = 1000
    itens.Range("C2").Value = 10
    itens.Range("E2").Value = 950
    if fotografia_c2 is not None:
        itens.Range("G2").Value = fotografia_c2
    if posicao_incompleta:
        # Segundo item com fotografia de C1, mas SEM quantidade na posicao atual.
        itens.Range("A3").Value = "1.2"
        itens.Range("B3").Value = 10
        itens.Range("C3").Value = 10
        itens.Range("E3").Value = 10
    if metodo == PC:
        pcs = livro.Worksheets("itens_PC")
        for linha, (numero, data_pc, valor, pago) in enumerate(PCS, start=2):
            pcs.Range(f"A{linha}").Value = numero
            pcs.Range(f"B{linha}").Value = Time(data_pc)
            pcs.Range(f"D{linha}").Value = valor
            pcs.Range(f"G{linha}").Value = pago
    else:
        financeiro = livro.Worksheets("financeiro")
        for linha in range(2, 74):
            competencia = financeiro.Range(f"A{linha}").Value
            if hasattr(competencia, "year") and competencia.replace(tzinfo=None) <= CORTE:
                financeiro.Range(f"C{linha}").Value = 50
    if alerta_aditivo:
        # Base economica em linha que nao e "novo item": ALERTA em aditivos!M.
        # Data posterior a posicao: nao mexe na posicao atual nem no residual.
        aditivos = livro.Worksheets("aditivos")
        aditivos.Range("A2").Value = "1.1"
        aditivos.Range("B2").Value = Time(datetime(2025, 9, 15))
        aditivos.Range("D2").Value = "Acrescimo"
        aditivos.Range("E2").Value = 10
        aditivos.Range("H2").Value = "Sim"
        aditivos.Range("K2").Value = "Nao"
        aditivos.Range("N2").Value = "C0"
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
        "D20": mem.Range("D20").Value2,
        "B21": mem.Range("B21").Value2,
        "N263": mem.Range("N263").Value2,
        "K": [pcs.Range(f"K{r}").Value2 for r in range(2, 2 + len(PCS))],
        "L": [pcs.Range(f"L{r}").Value2 for r in range(2, 2 + len(PCS))],
        "cee_A9": livro.Worksheets("CICLO_EM_EXECUCAO").Range("A9").Value2,
        "cee_P13": livro.Worksheets("CICLO_EM_EXECUCAO").Range("P13").Value2,
        "H8": rd.Range("H8").Value2,
        "H14": rd.Range("H14").Value2,
        "H33": rd.Range("H33").Value2,
        "J5": rd.Range("J5").Value2,
        "B7": rd.Range("B7").Value2,
        "B36_B38": [rd.Range(f"B{r}").Value2 for r in (36, 37, 38)],
        "B87": rd.Range("B87").Value2,
        "cobertura_B8": livro.Worksheets("cobertura_temporal").Range("B8").Value2,
        "aditivo_M": livro.Worksheets("aditivos").Range("M2").Value2,
    }


def _rodar(tmp_path: Path, nome: str, metodo: str, *, fotografia_c2=None,
           alerta_aditivo: bool = False, posicao_incompleta: bool = False) -> dict:
    import pythoncom
    import win32com.client

    caminho = tmp_path / f"Coleta_11_8_{nome}.xlsx"
    caminho.write_bytes(gerar_coleta_oficial_preenchida(_dados()))
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    livro = None
    try:
        abrir = lambda: excel.Workbooks.Open(  # noqa: E731
            str(caminho.resolve()), UpdateLinks=0, ReadOnly=False, CorruptLoad=0
        )
        livro = _tentar(pythoncom, abrir)
        time.sleep(1.0)
        _tentar(pythoncom, lambda: _preencher(
            livro, metodo, fotografia_c2=fotografia_c2, alerta_aditivo=alerta_aditivo,
            posicao_incompleta=posicao_incompleta,
        ))
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


def test_pcs_sem_fotografia_do_ciclo_vigente_usa_a_posicao_atual(tmp_path: Path):
    r = _rodar(tmp_path, "pc_sem_c2", PC)
    m = r["mem"]
    assert r["K"] == ["OK", "OK", "OK", "OK"]  # "Não" e precluso sem pendencia
    assert r["L"] == ["Nao", "Sim", "Sim", "Nao"]
    assert m["D44"] == 0
    assert r["cee_P13"] == "C1"  # referencia fisica anterior ao ciclo
    assert r["cee_A9"] == pytest.approx(800 * VU_C2)
    assert m["T26"] >= 1 and m["T49"] == 1
    assert m["T21"] == pytest.approx(EXEC_C0)
    assert m["T22"] == pytest.approx(EXEC_C1)
    assert m["T50"] == pytest.approx(EXEC_C2)
    assert m["T23"] == pytest.approx(800 * VU_C2)
    assert m["T39"] == pytest.approx(POTENCIAL)
    vta = EXEC_C0 + EXEC_C1 + EXEC_C2 + 800 * VU_C2 + POTENCIAL
    assert m["B26"] == pytest.approx(vta)
    assert m["T40"] == pytest.approx(vta - POTENCIAL)
    assert m["W52"] == "POSICAO ATUAL SEM FOTOGRAFIA DO CICLO VIGENTE - CONFERIR"
    # Pendencias: nenhuma (VTA e retroativo validados, ciclo atual validado).
    assert (r["H8"], r["H14"], r["H33"], r["J5"]) == ("VALIDADO", "VALIDADO", "VALIDADO", 0)
    assert r["B7"].startswith("Nenhuma")
    assert r["B36_B38"] == [
        pytest.approx(EXEC_C2), pytest.approx(EXEC_C2 + 800 * VU_C2), pytest.approx(800 * VU_C2),
    ]
    assert r["B87"] == pytest.approx(0.0)
    assert int(r["cobertura_B8"]) == (CORTE - datetime(1899, 12, 30)).days  # = CICLO_EM_EXECUCAO!D5


def test_pcs_com_fotografia_do_ciclo_vigente_mantem_o_calculo_11_7(tmp_path: Path):
    r = _rodar(tmp_path, "pc_com_c2", PC, fotografia_c2=850)
    m = r["mem"]
    assert m["T26"] == 0 and m["T49"] == 0 and m["T50"] == 0
    assert m["T51"] in (None, "")
    assert m["T23"] == pytest.approx(850 * VU_C2)  # SUM(Y): residual x VU, como na 11.7
    vta = EXEC_C0 + EXEC_C1 + 850 * VU_C2 + POTENCIAL
    assert m["T25"] == pytest.approx(vta)
    assert m["B26"] == pytest.approx(vta)


def test_financeiro_sem_fotografia_usa_a_posicao_atual_como_remanescente(tmp_path: Path):
    r = _rodar(tmp_path, "fin_sem_c2", FINANCEIRO)
    m = r["mem"]
    assert m["T49"] == 1
    assert m["D35"] == pytest.approx(800 * VU_C2)
    extra = r["N263"] if isinstance(r["N263"], float) else 0.0
    assert m["B26"] == pytest.approx(r["D20"] + r["B21"] + 800 * VU_C2 + extra)
    assert "INCOMPLETO" not in str(m["E26"])


def test_alerta_em_aditivos_continua_esvaziando_o_vta(tmp_path: Path):
    r = _rodar(tmp_path, "pc_alerta", PC, alerta_aditivo=True)
    m = r["mem"]
    assert str(r["aditivo_M"]).startswith("ALERTA:")
    assert m["T48"] >= 1 and m["T49"] == 1
    assert m["B26"] in (None, "") and m["T40"] in (None, "")


def test_negativo_posicao_atual_incompleta_nao_usa_o_caminho_alternativo(tmp_path: Path):
    r = _rodar(tmp_path, "pc_incompleta", PC, posicao_incompleta=True)
    m = r["mem"]
    assert r["cee_A9"] in (None, "")  # um item sem quantidade atual
    assert m["W49"] == 0 and m["T49"] == 0 and m["T50"] == 0
    assert m["T51"] in (None, "")
    assert m["T25"] == "CALCULO MANUAL REQUERIDO"
    assert m["B26"] in (None, "") and m["T40"] in (None, "")
    assert r["H8"] == "REVISE" and r["J5"] >= 1
    assert "composição do VTA" in r["B7"]
