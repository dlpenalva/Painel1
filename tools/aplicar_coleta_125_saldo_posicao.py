# -*- coding: utf-8 -*-
"""Coleta 12.5 — posicao da tabela SALDO REMANESCENTE — EVOLUCAO na RESULTADOS.

Somente POSICIONAMENTO. Nenhum calculo, valor, texto, estilo ou fonte muda.

A tabela da 12.4 (RESULTADOS!B76:G83) ficava depois do bloco DETALHES DO
METODO — PEDIDOS DE COMPRA (B65:G74), que so aparece no metodo PCs; nos demais
metodos ele fica em branco e abria um vao entre a observacao final (B63) e a
tabela. A ordem dos dois blocos e invertida:

    SALDO REMANESCENTE — EVOLUCAO      B76:G83 -> B65:G72
    DETALHES DO METODO — PCs           B65:G74 -> B74:G83

O movimento e feito pelo proprio Excel (Range.Cut), que leva formulas, estilos,
bordas, mesclas e formatacao condicional e reajusta as referencias internas
(C78 -> D78 vira C67 -> D67). As alturas de linha acompanham cada bloco; a
linha 64 segue como respiro depois da observacao e a 73 separa os blocos. A
area de impressao continua $B$1:$G$83 (o bloco PCs termina na 83).

Unico texto acompanhado: o rotulo MEMORIA_RESULTADOS!V68, que cita a celula da
ultima posicao informada (RESULTADOS!B83 -> RESULTADOS!B72).

uso: python tools/aplicar_coleta_125_saldo_posicao.py [caminho.xlsx]
     (exige pywin32 + Excel real; sem argumento, altera o template oficial)
"""
from __future__ import annotations

import gc
import re
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
for _caminho in (RAIZ, RAIZ / "tools"):
    if str(_caminho) not in sys.path:
        sys.path.insert(0, str(_caminho))

import aplicar_coleta_124_precisao_saldo as c124  # noqa: E402
import aplicar_resultados_executivo_v3 as v3  # noqa: E402
import aplicar_resultados_retro_reconhecido_potencial as c121  # noqa: E402

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
EXE, MEM = c124.EXE, c124.MEM

# ------------------------------------------------------------- posicoes
SALDO_ANTES = (c124.L_SALDO_TITULO, c124.L_SALDO_ULTIMA)   # 76..83
PCS_ANTES = (65, 74)
L_SALDO_TITULO = 65
L_SALDO_CAB = 66
L_SALDO_C0 = 67
L_SALDO_ULTIMA = 72
L_RESPIRO = 73
PCS_DEPOIS = (74, 83)
L_TEMP = 90                       # area livre usada so durante a troca
AREA_IMPRESSAO = "$B$1:$G$83"
ROTULO_APOIO_68 = (
    "APOIO - SALDO REMANESCENTE: ULTIMA POSICAO INFORMADA (RESULTADOS!B72)")

_REF = re.compile(r"(?<![A-Za-z0-9_!$.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\d(!A-Za-z_])")


def _deslocar(texto: str, faixa: tuple[int, int], delta: int) -> str:
    """Reajusta as referencias da propria aba (fora de aspas) que caem na faixa."""
    def troca(m: re.Match) -> str:
        linha = int(m.group(4))
        if faixa[0] <= linha <= faixa[1]:
            linha += delta
        return f"{m.group(1)}{m.group(2)}{m.group(3)}{linha}"

    partes = texto.split('"')
    partes[::2] = [_REF.sub(troca, p) for p in partes[::2]]
    return '"'.join(partes)


def _mover(itens: dict[str, str], faixa: tuple[int, int], delta: int) -> dict[str, str]:
    return {_deslocar(k, faixa, delta): (_deslocar(v, faixa, delta) if v.startswith("=") else v)
            for k, v in itens.items()}


DELTA_SALDO = L_SALDO_TITULO - SALDO_ANTES[0]   # -11
DELTA_PCS = PCS_DEPOIS[0] - PCS_ANTES[0]        # +9


def formulas_saldo() -> dict[str, str]:
    """Formulas da tabela de saldo na posicao 12.5 (B65:G72)."""
    return _mover(c124.formulas_saldo(), SALDO_ANTES, DELTA_SALDO)


def rotulos_saldo() -> dict[str, str]:
    return _mover(c124.rotulos_saldo(), SALDO_ANTES, DELTA_SALDO)


def formulas_executivo() -> dict[str, str]:
    """Formulas da RESULTADOS executiva da 12.1 com o bloco PCs em B74:G83."""
    return _mover(c121.formulas_executivo(), PCS_ANTES, DELTA_PCS)


APOIO_ULTIMA = dict(c124.APOIO_ULTIMA)
APOIO_ULTIMA[68] = (ROTULO_APOIO_68, None)


# ------------------------------------------------------------- conferencias
def conferir(wb) -> None:
    """Forma 12.4 exigida antes da troca (fail-closed)."""
    ws = wb.Worksheets(EXE)
    for celula, formula in c124.formulas_saldo().items():
        if ws.Range(celula).Formula != formula:
            raise RuntimeError(f"{EXE}!{celula} fora da forma 12.4; nada aplicado")
    for celula, formula in c121.formulas_executivo().items():
        if ws.Range(celula).Formula != formula:
            raise RuntimeError(f"{EXE}!{celula} fora da forma 12.4; nada aplicado")
    if ws.PageSetup.PrintArea.replace("$", "") != AREA_IMPRESSAO.replace("$", ""):
        raise RuntimeError("area de impressao da RESULTADOS fora da forma 12.4")
    for linha in range(SALDO_ANTES[1] + 1, L_TEMP + 10):
        for col in "ABCDEFGH":
            if ws.Range(f"{col}{linha}").Formula not in ("", None):
                raise RuntimeError(f"{EXE}!{col}{linha} ocupada; nada aplicado")
    if wb.Worksheets(MEM).Range("V68").Value != c124.APOIO_ULTIMA[68][0]:
        raise RuntimeError(f"{MEM}!V68 fora da forma 12.4; nada aplicado")
    if ws.ProtectContents:
        raise RuntimeError("RESULTADOS protegida inesperada; nada aplicado")


def conferir_resultado(wb) -> None:
    ws = wb.Worksheets(EXE)
    esperado = {**formulas_executivo(), **formulas_saldo()}
    for celula, formula in esperado.items():
        if ws.Range(celula).Formula != formula:
            raise RuntimeError(f"{EXE}!{celula} fora da forma 12.5")
    for celula, texto in rotulos_saldo().items():
        if ws.Range(celula).Value != texto:
            raise RuntimeError(f"{EXE}!{celula} sem o rotulo esperado")
    for linha in range(PCS_DEPOIS[1] + 1, L_TEMP + 10):
        for col in "BCDEFG":
            if ws.Range(f"{col}{linha}").Formula not in ("", None):
                raise RuntimeError(f"{EXE}!{col}{linha} nao ficou limpa")
    if ws.PageSetup.PrintArea.replace("$", "") != AREA_IMPRESSAO.replace("$", ""):
        raise RuntimeError("area de impressao alterada")


# ------------------------------------------------------------- COM
@c124._com_seguro
def _recortar(ws, origem: str, destino: str) -> None:
    ws.Range(origem).Cut(ws.Range(destino))


def trocar_blocos(wb) -> None:
    ws = wb.Worksheets(EXE)
    alturas = {r: ws.Rows(r).RowHeight for r in range(PCS_ANTES[0], SALDO_ANTES[1] + 1)}
    n_saldo = SALDO_ANTES[1] - SALDO_ANTES[0]
    n_pcs = PCS_ANTES[1] - PCS_ANTES[0]
    # Origem e destino nunca se sobrepoem: com sobreposicao o Excel perde a
    # mescla e a formatacao condicional da ultima linha do bloco PCs.
    _recortar(ws, f"B{PCS_ANTES[0]}:G{PCS_ANTES[1]}", f"B{L_TEMP}")
    _recortar(ws, f"B{SALDO_ANTES[0]}:G{SALDO_ANTES[1]}", f"B{L_SALDO_TITULO}")
    _recortar(ws, f"B{L_TEMP}:G{L_TEMP + n_pcs}", f"B{PCS_DEPOIS[0]}")
    for i in range(n_saldo + 1):
        ws.Rows(L_SALDO_TITULO + i).RowHeight = alturas[SALDO_ANTES[0] + i]
    ws.Rows(L_RESPIRO).RowHeight = alturas[SALDO_ANTES[0] - 1]
    for i in range(PCS_ANTES[1] - PCS_ANTES[0] + 1):
        ws.Rows(PCS_DEPOIS[0] + i).RowHeight = alturas[PCS_ANTES[0] + i]
    ws.PageSetup.PrintArea = AREA_IMPRESSAO
    wb.Worksheets(MEM).Range("V68").Value = ROTULO_APOIO_68


def aplicar(caminho: Path) -> None:
    import pythoncom
    import win32com.client as com

    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        excel.Calculation = v3.XL_CALCULO_MANUAL
        conferir(wb)
        trocar_blocos(wb)
        conferir_resultado(wb)
        excel.Calculation = v3.XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()
        wb.Worksheets("CONTROLE").Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        wb.Save()
        v3._fechar(wb)
        wb = None
    finally:
        if wb is not None:
            v3._fechar(wb)
        for _ in range(10):
            try:
                excel.Quit()
                break
            except Exception:
                time.sleep(1.0)
        wb = excel = None
        gc.collect()
        pythoncom.CoUninitialize()


def main() -> int:
    caminho = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else TEMPLATE
    if not caminho.exists():
        print("ERRO: arquivo nao encontrado: %s" % caminho)
        return 1
    aplicar(caminho)
    print("OK: Coleta 12.5 (posicao do saldo remanescente) aplicada em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
