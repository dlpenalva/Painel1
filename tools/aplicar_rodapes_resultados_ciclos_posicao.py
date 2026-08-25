# -*- coding: utf-8 -*-
"""RESULTADOS — rodapes dos cards RETROATIVO e REMANESCENTE (apresentacao).

Duas correcoes de LEITURA na aba RESULTADOS. Nenhum valor, formula financeira
ou regra de negocio muda: apenas dois rodapes de card passam a dizer o que o
usuario precisa saber.

1) Card "RETROATIVO TOTAL A PAGAR" (rodape E6)

   Antes: "Ciclos apurados" (nao dizia QUAIS ciclos).
   Depois: "Ciclos apurados — C1 e C2" (somente os efetivamente computados).

   Fonte canonica: `parametros!A3:A6` (COMPUTAR_NESTA_APURACAO) com o nome do
   ciclo em `parametros!B3:B6`. E a MESMA fonte que a propria aba ja usa para
   validar o card do retroativo (RESULTADOS!H14 exige B16..B20 numerico
   exatamente para os ciclos com A="Sim") e que o gerador usa para montar
   `ciclos_computados`. C0 e a base contratual (percentual 0, situacao "Base")
   e nunca e um ciclo apurado — por isso a leitura comeca na linha 3 (C1).

   Nenhuma regra nova decide o que foi apurado: as celulas auxiliares J9/J11
   apenas LISTAM o que a fonte canonica ja declara. O valor do retroativo
   (D5 = D22 = RETRO_OFICIAL) permanece intocado.

2) Card "REMANESCENTE ATUALIZADO" (rodape F6)

   Antes: "Saldo do ciclo em execução" (anunciava um dado que nao vinha).
   Depois: "Posição em DD/MM/AAAA".

   Fonte canonica: `CONTROLE!B3` (DATA DE CORTE DA APURACAO) — a mesma data
   que o proprio card ja ancora em F5 e que o cabecalho exibe em E3. Nunca a
   data corrente da maquina. Sem data de corte informada, o rodape declara a
   ausencia (fail-closed), sem inventar data.

   A data e composta por RIGHT/DAY/MONTH/YEAR (mesmo padrao ja usado em G2 e
   H10 desta aba) em vez de TEXT(...,"dd/mm/yyyy"): codigo de formato dentro
   de TEXT e dependente de locale e quebraria o arquivo fora do pt-BR.

Celulas auxiliares criadas: J9, J10 e J11 — coluna J ja e a coluna oculta de
apoio da aba (J2/J3/J5/J6/J8). Nenhuma celula existente muda de endereco.

Uso:
    python tools/aplicar_rodapes_resultados_ciclos_posicao.py <origem.xlsx> <destino.xlsx>
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

ABA_RESULTADOS = "RESULTADOS"
ABA_PARAMETROS = "parametros"
ABA_CONTROLE = "CONTROLE"
ABA_MEMORIA = "MEMORIA_RESULTADOS"

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

CINZA_AUXILIAR = 0x595959

# ---------------------------------------------------------------- formulas

# Quantidade de ciclos efetivamente computados nesta apuracao (C1..C4).
FORMULA_J9 = '=COUNTIF(parametros!$A$3:$A$6,"Sim")'

# Lista bruta, separada por ", ", dos ciclos computados — nomes lidos de
# parametros!B (nunca digitados aqui).
FORMULA_J10 = (
    '=TEXTJOIN(", ",TRUE,'
    'IF(parametros!$A$3="Sim",parametros!$B$3,""),'
    'IF(parametros!$A$4="Sim",parametros!$B$4,""),'
    'IF(parametros!$A$5="Sim",parametros!$B$5,""),'
    'IF(parametros!$A$6="Sim",parametros!$B$6,""))'
)

# Mesma lista com " e " antes do ultimo item ("C1, C2 e C3").
FORMULA_J11 = (
    '=IF($J$9=0,"",IF($J$9=1,$J$10,'
    'LEFT($J$10,FIND("|",SUBSTITUTE($J$10,", ","|",$J$9-1))-1)&" e "&'
    'MID($J$10,FIND("|",SUBSTITUTE($J$10,", ","|",$J$9-1))+2,200)))'
)

FORMULA_E6 = (
    '=IF($D$22<>"",'
    'IF($J$11="","Ciclos apurados","Ciclos apurados — "&$J$11),'
    'IF($J$6=0,"SELECIONE O MÉTODO NA ABA CONTROLE",'
    '"INFORME A BASE DO MÉTODO"))'
)

FORMULA_F6 = (
    '=IF(CONTROLE!$B$3="","Data de corte não informada",'
    '"Posição em "&RIGHT("0"&DAY(CONTROLE!$B$3),2)&"/"&'
    'RIGHT("0"&MONTH(CONTROLE!$B$3),2)&"/"&YEAR(CONTROLE!$B$3))'
)

TEXTO_REMOVIDO_F6 = "Saldo do ciclo em execução"

# Tudo que esta etapa de apresentacao NAO pode tocar.
TRAVAS_RESULTADOS = (
    "C5", "D5", "G5", "D22", "B38", "B37", "B36", "H8", "H14", "H24", "H33",
    "B83", "B84", "B85", "B86", "B87", "B63", "J5", "J6", "J8",
)
TRAVAS_MEMORIA = ("B16", "B21", "B26", "D20", "D35", "F20", "T21")


def _nomes_abas(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


def _capturar_protecao(ws):
    estado = bool(ws.ProtectContents)
    selecao = None
    if estado:
        selecao = ws.EnableSelection
        ws.Unprotect()
    return estado, selecao


def _restaurar_protecao(ws, estado, selecao) -> None:
    if not estado:
        return
    ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
    if selecao is not None:
        ws.EnableSelection = selecao


def _validar_origem(wb) -> None:
    abas = _nomes_abas(wb)
    for obrig in (ABA_RESULTADOS, ABA_PARAMETROS, ABA_CONTROLE, ABA_MEMORIA):
        if obrig not in abas:
            raise ValueError(f"Aba {obrig} ausente na origem.")
    res = wb.Worksheets(ABA_RESULTADOS)
    if str(res.Range("D4").Value or "").strip() != "RETROATIVO TOTAL A PAGAR":
        raise ValueError("RESULTADOS!D4 nao e o card do retroativo.")
    if str(res.Range("F4").Value or "").strip() != "REMANESCENTE ATUALIZADO":
        raise ValueError("RESULTADOS!F4 nao e o card do remanescente.")
    for endereco in ("J9", "J10", "J11"):
        if str(res.Range(endereco).Formula or "").strip() not in ("", "0"):
            raise ValueError(
                f"RESULTADOS!{endereco} ja esta ocupada; a coluna auxiliar "
                "precisa estar livre para receber a lista de ciclos."
            )


def _snapshot_travas(wb) -> dict[str, str]:
    res = wb.Worksheets(ABA_RESULTADOS)
    mem = wb.Worksheets(ABA_MEMORIA)
    travas = {f"RESULTADOS!{e}": str(res.Range(e).Formula) for e in TRAVAS_RESULTADOS}
    travas.update(
        {f"MEMORIA!{e}": str(mem.Range(e).Formula) for e in TRAVAS_MEMORIA}
    )
    return travas


def _aplicar(wb) -> None:
    res = wb.Worksheets(ABA_RESULTADOS)
    estado, selecao = _capturar_protecao(res)
    try:
        # Auxiliares na coluna oculta de apoio (J), sem alterar nenhum endereco
        # ja existente. Formato/fonte iguais aos auxiliares vizinhos (J5/J6).
        for endereco, formula in (
            ("J9", FORMULA_J9), ("J10", FORMULA_J10), ("J11", FORMULA_J11),
        ):
            celula = res.Range(endereco)
            celula.Formula = formula
            # NumberFormat nao e tocado: as celulas nascem em "Geral" e o
            # Excel pt-BR rejeita o codigo "General" nesta propriedade
            # (armadilha de locale ja conhecida no repositorio).
            celula.Font.Name = "Calibri"
            celula.Font.Size = 9
            celula.Font.Bold = False
            celula.Font.Color = CINZA_AUXILIAR

        res.Range("E6").Formula = FORMULA_E6
        res.Range("F6").Formula = FORMULA_F6
    finally:
        _restaurar_protecao(res, estado, selecao)


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_rodapes_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    salvo = False
    try:
        wb = excel.Workbooks.Open(
            str(tmp_xlsx), UpdateLinks=0, ReadOnly=False, CorruptLoad=0
        )
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALC_MANUAL
        aba_ativa = wb.ActiveSheet.Name

        _validar_origem(wb)
        travas_antes = _snapshot_travas(wb)

        _aplicar(wb)

        travas_depois = _snapshot_travas(wb)
        if travas_antes != travas_depois:
            difs = {
                k: (travas_antes[k], travas_depois[k])
                for k in travas_antes if travas_antes[k] != travas_depois[k]
            }
            raise RuntimeError(f"TRAVA VIOLADA pelos rodapes: {difs}")

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()

        if aba_ativa in _nomes_abas(wb):
            wb.Worksheets(aba_ativa).Activate()
        wb.Save()
        salvo = True
        wb.Close(SaveChanges=False)
        wb = None

        # Reabre sem reparo para provar zero-corrupcao.
        wb = excel.Workbooks.Open(
            str(tmp_xlsx), UpdateLinks=0, ReadOnly=True, CorruptLoad=0
        )
        res = wb.Worksheets(ABA_RESULTADOS)
        if "Ciclos apurados" not in str(res.Range("E6").Formula):
            raise RuntimeError("E6 perdeu o rodape de ciclos apurados.")
        if "Posição em " not in str(res.Range("F6").Formula):
            raise RuntimeError("F6 perdeu o rodape de posicao.")
        if TEXTO_REMOVIDO_F6 in str(res.Range("F6").Formula):
            raise RuntimeError("F6 ainda anuncia o saldo do ciclo em execucao.")
        for endereco in ("J9", "J10", "J11"):
            if not str(res.Range(endereco).Formula).startswith("="):
                raise RuntimeError(f"{endereco} nao recebeu formula.")
        wb.Close(SaveChanges=False)
        wb = None
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        excel.Quit()
        del wb
        del excel
        pythoncom.CoUninitialize()

    if not salvo:
        shutil.rmtree(tmp_dir, ignore_errors=True)
        raise RuntimeError("Excel nao salvou; destino preservado.")
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(tmp_xlsx, destino)
    shutil.rmtree(tmp_dir, ignore_errors=True)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("origem", type=Path)
    parser.add_argument("destino", type=Path)
    args = parser.parse_args()
    aplicar(args.origem, args.destino)
    print(f"OK: {args.destino}")


if __name__ == "__main__":
    main()
