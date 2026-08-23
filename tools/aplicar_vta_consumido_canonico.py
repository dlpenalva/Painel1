# -*- coding: utf-8 -*-
"""VTA-C2 — VTA do metodo Consumido/Itens alinhado a identidade canonica no XLS.

Aplica, via Excel COM em copia temporaria (padrao zero-corrupcao ja usado por
`aplicar_vta_financeiro_canonico.py`):

Identidade para o metodo Consumido (CONTROLE!B1 -> MEMORIA_RESULTADOS!B4 =
"Itens"):

    VTA = execucao consumida atualizada (Soma VALOR_CONS_C0..C4)
        + futuro remanescente atualizado (D35, quando B4="Itens" ja
          resolve para D33 — corrigido nesta etapa)
        + ajustes manuais/aditivos ja legitimamente previstos em B26

    NAO soma B21 novamente — B21 e apenas a decomposicao do reajuste ja
    contida na execucao atualizada (mesma identidade provada em VTA-C1.1:
    execucao_base(280) + ajuste(4) = execucao_atualizada(284)).

* MEMORIA_RESULTADOS:
    - E20/F20 (novas): "Execucao Consumida Atualizada" = SOMA de
      itens_Consumidos!(F+H+J+L+N) (VALOR_CONS_C0..C4, ja calculados pelo
      proprio template com VU_original x fator_acumulado do ciclo — nunca
      VU_C0 universal), com gate de presenca (COUNT das colunas QTD_CONS)
      para distinguir "nenhum consumo informado" (vazio) de "consumo
      explicitamente zero" (zero valido).
    - B26 (VTA FINAL): ganha um ramo NOVO exclusivamente para
      $B$4="Itens" (F20+D35+ajustes). Os ramos PC, Financeiro e o ramo
      generico residual (agora so para metodo indeterminado) NAO sao
      alterados — permanecem byte-a-byte identicos.
    - B28 (referencia comparativa): passa a incluir "Itens" na condicao que
      ja preserva a formula generica antiga (B20+B21+B22) para Financeiro,
      exatamente como o Financeiro ja faz. PC continua mostrando $B$26.
* itens_Consumidos:
    - V2:V200 (nova coluna auxiliar): "Remanescente atualizado (base) do
      item" = MAX(QTD_CONTRATADA-CONS_QTD_TOTAL,0) x VU_ORIGINAL, por
      LINHA (formula escalar, sem SUMPRODUCT/array) — evita um bug real de
      coercao booleana do Excel confirmado empiricamente (SUMPRODUCT com
      mascara booleana como array retornava 20 em vez do correto 120 no
      caso controlado; a mesma formula em escala por linha + SUM simples
      calcula corretamente).
* MEMORIA_RESULTADOS!C33/D33 (remanescente futuro do metodo Itens):
    - C33 passa a ser SUM(itens_Consumidos!V2:V200) (antes: SUMPRODUCT
      buggy). D33 passa a multiplicar o C33 corrigido pelo fator do ciclo
      vigente (mesma logica de selecao de fator, inalterada).

Trava anti-regressao: T21/T22/T23/T25 (PC), o ramo PC dentro de B26, o ramo
Financeiro dentro de B26 (a subexpressao literal do IF($B$4="Financeiro",
...)) e a formula generica residual (ELSE final) precisam permanecer
identicas antes/depois, verificado por assinatura de texto.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

ABA_MEMORIA = "MEMORIA_RESULTADOS"
ABA_CONSUMIDOS = "itens_Consumidos"

_PROT_FLAGS = (
    "AllowFormattingCells",
    "AllowFormattingColumns",
    "AllowFormattingRows",
    "AllowInsertingColumns",
    "AllowInsertingRows",
    "AllowInsertingHyperlinks",
    "AllowDeletingColumns",
    "AllowDeletingRows",
    "AllowSorting",
    "AllowFiltering",
    "AllowUsingPivotTables",
)

# Ramos preservados byte-a-byte dentro de B26 — travas anti-regressao.
_RAMO_PC_B26 = (
    'IF($T$25="CALCULO MANUAL REQUERIDO","",ROUND($T$25+IF(ISNUMBER(B24),B24,0),2))'
)
_RAMO_FINANCEIRO_B26 = (
    'IF(OR($D$20="",B21="",D35="",AND(B24<>"",NOT(ISNUMBER(B24)))),"",'
    'ROUND($D$20+B21+D35+IF(ISNUMBER(B24),B24,0)+IF(ISNUMBER($N$263),$N$263,0),2))'
)
_RAMO_RESIDUAL_B26 = (
    'IF(OR(B23="",AND(B24<>"",NOT(ISNUMBER(B24)))),"",'
    'ROUND(B23+IF(ISNUMBER(B24),B24,0)+IF(ISNUMBER($N$263),$N$263,0),2))'
)
_B26_ANTES = (
    '=IF(AND(B24<>"",B25<>""),"",IF(ISNUMBER(B25),B25,IF($B$4="PCs",'
    + _RAMO_PC_B26
    + ',IF($B$4="Financeiro",'
    + _RAMO_FINANCEIRO_B26
    + ','
    + _RAMO_RESIDUAL_B26
    + '))))'
)
_RAMO_ITENS_B26 = (
    'IF(OR($F$20="",D35="",AND(B24<>"",NOT(ISNUMBER(B24)))),"",'
    'ROUND($F$20+D35+IF(ISNUMBER(B24),B24,0)+IF(ISNUMBER($N$263),$N$263,0),2))'
)
_B26_NOVO = (
    '=IF(AND(B24<>"",B25<>""),"",IF(ISNUMBER(B25),B25,IF($B$4="PCs",'
    + _RAMO_PC_B26
    + ',IF($B$4="Financeiro",'
    + _RAMO_FINANCEIRO_B26
    + ',IF($B$4="Itens",'
    + _RAMO_ITENS_B26
    + ','
    + _RAMO_RESIDUAL_B26
    + ')))))'
)

# VTA-C2.1 (item 4): F20 fica vazio quando QUALQUER par QTD_CONS_Cn/
# VALOR_CONS_Cn tiver contagens numericas diferentes — sinal de que uma
# quantidade foi informada mas nao pode ser valorada (ex.: VU_original ou
# fator ausente), o que a formula VALOR_CONS_Cn ja expressa devolvendo ""
# nesse caso. COUNT(par qtd)<>COUNT(par valor) detecta essa divergencia
# sem somar apenas as parcelas validas. Zero explicito continua valido
# (COUNT conta zeros).
_F20_FORMULA = (
    '=IF(OR('
    'COUNT(itens_Consumidos!$E$2:$E$200)<>COUNT(itens_Consumidos!$F$2:$F$200),'
    'COUNT(itens_Consumidos!$G$2:$G$200)<>COUNT(itens_Consumidos!$H$2:$H$200),'
    'COUNT(itens_Consumidos!$I$2:$I$200)<>COUNT(itens_Consumidos!$J$2:$J$200),'
    'COUNT(itens_Consumidos!$K$2:$K$200)<>COUNT(itens_Consumidos!$L$2:$L$200),'
    'COUNT(itens_Consumidos!$M$2:$M$200)<>COUNT(itens_Consumidos!$N$2:$N$200)'
    '),"",'
    'IF(COUNT(itens_Consumidos!$E$2:$E$200,itens_Consumidos!$G$2:$G$200,'
    'itens_Consumidos!$I$2:$I$200,itens_Consumidos!$K$2:$K$200,'
    'itens_Consumidos!$M$2:$M$200)=0,"",'
    'ROUND(SUM(itens_Consumidos!$F$2:$F$200,itens_Consumidos!$H$2:$H$200,'
    'itens_Consumidos!$J$2:$J$200,itens_Consumidos!$L$2:$L$200,'
    'itens_Consumidos!$N$2:$N$200),2)))'
)

_B28_ANTES = '=IF($B$4="Financeiro",$B$23,$B$26)'
_B28_NOVO = '=IF(OR($B$4="Financeiro",$B$4="Itens"),$B$23,$B$26)'

_FATOR_VIGENTE_EXPR = (
    'IF($F$4="C0",parametros!$F$2,IF($F$4="C1",parametros!$F$3,'
    'IF($F$4="C2",parametros!$F$4,IF($F$4="C3",parametros!$F$5,parametros!$F$6))))'
)
# VTA-C2.1 (item 6): C33 fica vazio quando QUALQUER item real (A<>"") tiver
# V vazio (dado incompleto ou sobreconsumo sinalizado por V, ver
# _linha_v_consumidos) — nunca soma so os itens validos. COUNTIFS evita o
# bug real de SUMPRODUCT com mascara booleana ja identificado nesta tarefa.
_C33_NOVO = (
    '=IF(COUNTIF(itens_Consumidos!$A$2:$A$200,"<>")=0,"",'
    'IF(COUNTIFS(itens_Consumidos!$A$2:$A$200,"<>",itens_Consumidos!$V$2:$V$200,"")>0,"",'
    'ROUND(SUM(itens_Consumidos!$V$2:$V$200),2)))'
)
_D33_NOVO = (
    f'=IF(OR(C33="",NOT(ISNUMBER({_FATOR_VIGENTE_EXPR}))),"",'
    f'ROUND(C33*{_FATOR_VIGENTE_EXPR},2))'
)


def _linha_v_consumidos(linha: int) -> str:
    # VTA-C2.1 (itens 1, 5, 7): fail-closed por linha — QTD_CONTRATADA (B),
    # CONS_QTD_TOTAL (O) e VU_ORIGINAL (C) precisam ser numericos, e o
    # proprio CHECK do template (Q, ja existente: "DIVERGENCIA: CONSUMO
    # MAIOR QUE CONTRATADO" quando O>B) precisa estar "OK" — caso
    # contrario, V fica vazio (nunca mascara sobreconsumo/incompletude
    # como saldo zero). Quando tudo valido, zero real (contrato 100%
    # consumido) e um resultado numerico legitimo (MAX(...,0)=0), nao
    # vazio.
    return (
        f'=IF(A{linha}="","",'
        f'IF(OR(NOT(ISNUMBER(B{linha})),NOT(ISNUMBER(O{linha})),'
        f'NOT(ISNUMBER(C{linha})),Q{linha}<>"OK"),"",'
        f'ROUND(MAX(B{linha}-O{linha},0)*C{linha},2)))'
    )


def _nomes_abas(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


def _nomes_definidos(wb) -> set[str]:
    return {str(nome.Name).split("!")[-1] for nome in wb.Names}


def _capturar_protecao(ws):
    if not bool(ws.ProtectContents):
        return None, None
    protecao = ws.Protection
    estado = {}
    for flag in _PROT_FLAGS:
        try:
            estado[flag] = getattr(protecao, flag)
        except Exception:
            estado[flag] = True
    try:
        selecao = ws.EnableSelection
    except Exception:
        selecao = None
    ws.Unprotect()
    return estado, selecao


def _restaurar_protecao(ws, estado, selecao) -> None:
    if estado is None:
        return
    ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True, **estado)
    if selecao is not None:
        try:
            ws.EnableSelection = selecao
        except Exception:
            pass


def _validar_origem(wb) -> None:
    abas = _nomes_abas(wb)
    obrigatorias = {ABA_MEMORIA, ABA_CONSUMIDOS, "parametros", "CONTROLE"}
    ausentes = sorted(obrigatorias.difference(abas))
    if ausentes:
        raise ValueError(f"Abas obrigatorias ausentes: {', '.join(ausentes)}")
    mem = wb.Worksheets(ABA_MEMORIA)
    for endereco in ("B4", "B20", "B21", "B22", "B23", "B26", "B28", "C33", "D33", "T25"):
        if not str(mem.Range(endereco).Formula or ""):
            raise ValueError(f"MEMORIA_RESULTADOS!{endereco} ausente; layout inesperado.")
    atual_b26 = str(mem.Range("B26").Formula)
    if atual_b26 != _B26_ANTES:
        raise ValueError(
            "MEMORIA_RESULTADOS!B26 nao corresponde ao formato esperado "
            "(template ja mudou desde a auditoria VTA-C1/C2). Formula atual:\n"
            f"{atual_b26}"
        )
    atual_b28 = str(mem.Range("B28").Formula)
    if atual_b28 != _B28_ANTES:
        raise ValueError(
            "MEMORIA_RESULTADOS!B28 nao corresponde ao formato esperado. "
            f"Formula atual:\n{atual_b28}"
        )
    if "VTA_FINAL" not in _nomes_definidos(wb):
        raise ValueError("Nome definido VTA_FINAL ausente.")
    if mem.Range("E20").Value not in (None, ""):
        raise ValueError("MEMORIA_RESULTADOS!E20 ja ocupada; escolher outra celula.")
    if mem.Range("F20").Value not in (None, ""):
        raise ValueError("MEMORIA_RESULTADOS!F20 ja ocupada; escolher outra celula.")
    consumidos = wb.Worksheets(ABA_CONSUMIDOS)
    if consumidos.Range("V1").Value not in (None, ""):
        raise ValueError("itens_Consumidos!V1 ja ocupada; escolher outra coluna.")


def _snapshot_travas(wb) -> dict[str, str]:
    mem = wb.Worksheets(ABA_MEMORIA)
    return {e: str(mem.Range(e).Formula) for e in ("T21", "T22", "T23", "T25")}


def _aplicar_itens_consumidos(wb) -> None:
    ws = wb.Worksheets(ABA_CONSUMIDOS)
    estado, selecao = _capturar_protecao(ws)
    try:
        ws.Range("V1").Value = "Remanescente atualizado (base) do item"
        for linha in range(2, 201):
            ws.Range(f"V{linha}").Formula = _linha_v_consumidos(linha)
    finally:
        _restaurar_protecao(ws, estado, selecao)


def _aplicar_memoria(wb) -> None:
    mem = wb.Worksheets(ABA_MEMORIA)
    estado, selecao = _capturar_protecao(mem)
    try:
        mem.Range("E20").Value = "Execucao Consumida Atualizada"
        mem.Range("F20").Formula = _F20_FORMULA
        mem.Range("B26").Formula = _B26_NOVO
        mem.Range("B28").Formula = _B28_NOVO
        mem.Range("C33").Formula = _C33_NOVO
        mem.Range("D33").Formula = _D33_NOVO
    finally:
        _restaurar_protecao(mem, estado, selecao)


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_vta_cons_"))
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

        _aplicar_itens_consumidos(wb)
        _aplicar_memoria(wb)

        travas_depois = _snapshot_travas(wb)
        if travas_antes != travas_depois:
            difs = {
                k: (travas_antes[k], travas_depois[k])
                for k in travas_antes if travas_antes[k] != travas_depois[k]
            }
            raise RuntimeError(f"TRAVA VIOLADA: T21/T22/T23/T25 alterados: {difs}")
        nova_b26 = str(wb.Worksheets(ABA_MEMORIA).Range("B26").Formula)
        if _RAMO_PC_B26 not in nova_b26:
            raise RuntimeError("TRAVA VIOLADA: ramo PC ausente/alterado dentro de B26.")
        if _RAMO_FINANCEIRO_B26 not in nova_b26:
            raise RuntimeError("TRAVA VIOLADA: ramo Financeiro ausente/alterado dentro de B26.")
        if _RAMO_RESIDUAL_B26 not in nova_b26:
            raise RuntimeError("TRAVA VIOLADA: ramo residual ausente/alterado dentro de B26.")

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
        abas = _nomes_abas(wb)
        for obrig in (ABA_MEMORIA, ABA_CONSUMIDOS):
            if obrig not in abas:
                raise RuntimeError(f"Aba {obrig} ausente apos reabertura.")
        if "VTA_FINAL" not in _nomes_definidos(wb):
            raise RuntimeError("Nome VTA_FINAL ausente apos reabertura.")
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
    parser = argparse.ArgumentParser()
    parser.add_argument("origem", type=Path)
    parser.add_argument("destino", type=Path)
    args = parser.parse_args()
    aplicar(args.origem, args.destino)
    print("VTA Consumido canonico aplicado:", args.destino)


if __name__ == "__main__":
    main()
