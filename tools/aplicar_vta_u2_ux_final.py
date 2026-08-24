# -*- coding: utf-8 -*-
"""VTA-U2 (UX final) — a aba RESULTADOS destaca exclusivamente o VTA Oficial.

Complemento de APRESENTACAO sobre o estado ja aplicado pela VTA-U2. Nenhuma
formula de negocio muda: MEMORIA_RESULTADOS, C5 (VTA_FINAL) e as celulas
tecnicas B10/B11/B12/B13 + H10/H11/H13 sao travadas e conferidas ao final.

Restricao que dita o desenho: `_leitor_masterfile_v10._ler_referencias_vta` le
`RESULTADOS!B10/B11/B12/B13` e `H10/H11/H13` por ENDERECO FIXO. Mover essas
celulas alteraria dependencias (Sumario/Apostila consomem `referencias_vta`).
Por isso elas ficam onde estao e apenas deixam de ser APRESENTADAS.

O que muda:

1) Card principal: o subtitulo "Posição física atual" (A6) sai. O card passa a
   falar de uma unica grandeza — VTA. O valor auxiliar de B6 continua no
   arquivo, com exibicao suprimida pelo formato ";;;" (padrao ja usado na aba
   para ancoras que nao devem aparecer).

2) Linhas 10-13 (as tres referencias + a reconciliacao fisica) ficam OCULTAS.
   Formulas e valores intactos nos mesmos enderecos. Os rotulos passam a
   declarar-se como auditoria interna, para o caso de alguem reexibi-las.

3) A9 deixa de ser so um titulo de tabela de referencias e passa a enunciar a
   identidade canonica, com ponteiro para o quadro da secao 9. Os cabecalhos
   B9/C9/F9/H9 (que titulavam a tabela agora oculta) saem.

4) Bloco 8 (conferencia): "historico fisico" vira "historico quantitativo" /
   "dados quantitativos" — a UX so fala em posicao fisica onde ela e o proprio
   assunto (secao 4, ciclo em execucao), nunca perto do VTA.

Uso:
    python tools/aplicar_vta_u2_ux_final.py <origem.xlsx> <destino.xlsx>
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

ABA_MEMORIA = "MEMORIA_RESULTADOS"
ABA_RESULTADOS = "RESULTADOS"

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
MOEDA_BR = "R$ #.##0,00"
OCULTO = ";;;"

LINHA_REFS_INICIO = 10
LINHA_REFS_FIM = 13
LINHA_CONF_C0 = 73

TITULO_COMPOSICAO = (
    "1. COMPOSIÇÃO DO VTA — Executado apurado + Ajustes ainda devidos "
    "+ Remanescente atualizado = VTA Oficial (quadro detalhado na seção 9)"
)

ROTULOS_INTERNOS = {
    10: "[AUDITORIA INTERNA] Posicao fisica atual - nao e VTA",
    11: "[AUDITORIA INTERNA] Ultima posicao de abertura - nao e VTA",
    12: "[AUDITORIA INTERNA] Contrato original integralmente reajustado - comparativo",
    13: "[AUDITORIA INTERNA] Reconciliacao (posicao atual - ultima abertura)",
}

SEM_HISTORICO_C = "Sem historico quantitativo suficiente"
SEM_HISTORICO_E = "Sem dados quantitativos para comparar"
NAO_APLICAVEL = "Nao aplicavel ao metodo selecionado"

COLUNA_EXECUCAO = {"C0": "AC", "C1": "N", "C2": "P", "C3": "R"}


def _nomes_abas(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


def _nomes_definidos(wb) -> set[str]:
    return {n.Name for n in wb.Names}


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


# ------------------------------------------------------------- validacoes

def _validar_origem(wb) -> None:
    abas = _nomes_abas(wb)
    for obrig in (ABA_MEMORIA, ABA_RESULTADOS):
        if obrig not in abas:
            raise ValueError(f"Aba {obrig} ausente na origem.")
    if "VTA_FINAL" not in _nomes_definidos(wb):
        raise ValueError("Nome definido VTA_FINAL ausente.")
    res = wb.Worksheets(ABA_RESULTADOS)
    if "VTA_FINAL" not in str(res.Range("C5").Formula):
        raise ValueError(
            "RESULTADOS!C5 nao aponta para VTA_FINAL; aplicar antes o "
            "tools/aplicar_vta_uniformizacao_u2.py."
        )
    if str(res.Range("A79").Value or "").strip() != "9. COMO E FORMADO O VTA?":
        raise ValueError("Bloco 9 ausente; origem fora do estado da VTA-U2.")


def _snapshot_travas(wb) -> dict[str, str]:
    """Tudo que esta etapa de UX NAO pode tocar."""
    mem = wb.Worksheets(ABA_MEMORIA)
    res = wb.Worksheets(ABA_RESULTADOS)
    travas = {
        f"MEMORIA!{e}": str(mem.Range(e).Formula)
        for e in ("B16", "B20", "B21", "B22", "B23", "B26", "B28",
                  "D20", "F20", "D35", "T21", "T22", "T23", "T25")
    }
    travas.update({
        f"RESULTADOS!{e}": str(res.Range(e).Formula)
        for e in ("C5", "B10", "B11", "B12", "B13",
                  "H10", "H11", "H13", "B63", "B83", "B84", "B85", "B86", "B87")
    })
    return travas


# ------------------------------------------------------------- aplicacoes

def _aplicar_card(res) -> None:
    # O card do VTA fala de uma unica grandeza: VTA.
    res.Range("A6").ClearContents()
    # B6 permanece no arquivo (ancora tecnica), sem exibicao.
    res.Range("B6").NumberFormat = OCULTO


def _aplicar_bloco_composicao(res) -> None:
    res.Range("A9").Value = TITULO_COMPOSICAO
    res.Range("A9").Font.Bold = True
    # Cabecalhos da tabela de referencias, que deixa de ser apresentada.
    # C9 e F9 sao ancoras de intervalos mesclados (C9:E9 e F9:G9) — limpar a
    # celula isolada levanta "Nao podemos fazer isto em uma celula mesclada".
    for area in ("B9", "C9:E9", "F9:G9", "H9"):
        res.Range(area).ClearContents()

    for linha, rotulo in ROTULOS_INTERNOS.items():
        res.Range(f"A{linha}").Value = rotulo

    res.Range(f"{LINHA_REFS_INICIO}:{LINHA_REFS_FIM}").EntireRow.Hidden = True

    # Medida 10 do bloco 6 apontava para "Linhas 10 a 12 desta aba", que agora
    # estao ocultas. O ponteiro passa a dizer o que sao de fato.
    res.Range("B64").Formula = '="Linhas 10 a 13 (ocultas - auditoria interna)"'


def _formula_execucao_teorica(ciclo: str) -> str:
    if ciclo == "C4":
        corpo = f'"{SEM_HISTORICO_C}"'
    else:
        coluna = COLUNA_EXECUCAO[ciclo]
        alvo = f"itens_Remanesc!${coluna}$2:${coluna}$201"
        corpo = (
            'IF(SUMPRODUCT((itens_Remanesc!$A$2:$A$201<>"")*'
            f"(1-ISNUMBER({alvo})))>0,"
            f'"{SEM_HISTORICO_C}",ROUND(SUM(N({alvo})),2))'
        )
    return (
        f'=IF(MEMORIA_RESULTADOS!$B$4<>"Financeiro","{NAO_APLICAVEL}",{corpo})'
    )


def _formula_diferenca(linha: int) -> str:
    return (
        f'=IF(MEMORIA_RESULTADOS!$B$4<>"Financeiro","{NAO_APLICAVEL}",'
        f'IF(OR(B{linha}="",C{linha}="",NOT(ISNUMBER(C{linha}))),"",'
        f"ROUND(B{linha}-C{linha},2)))"
    )


def _formula_conferencia(linha: int) -> str:
    return (
        '=IF(MEMORIA_RESULTADOS!$B$4<>"Financeiro","NAO APLICAVEL",'
        f'IF(OR(B{linha}="",C{linha}="",NOT(ISNUMBER(C{linha}))),'
        f'"{SEM_HISTORICO_E}",'
        f'IF(ROUND(D{linha},2)=0,"OK","REVISAR")))'
    )


def _aplicar_conferencia(res) -> None:
    for indice, ciclo in enumerate(("C0", "C1", "C2", "C3", "C4")):
        linha = LINHA_CONF_C0 + indice
        res.Range(f"C{linha}").Formula = _formula_execucao_teorica(ciclo)
        res.Range(f"D{linha}").Formula = _formula_diferenca(linha)
        res.Range(f"E{linha}").Formula = _formula_conferencia(linha)
        res.Range(f"B{linha}:D{linha}").NumberFormat = MOEDA_BR


def _aplicar(wb) -> None:
    res = wb.Worksheets(ABA_RESULTADOS)
    estado, selecao = _capturar_protecao(res)
    try:
        _aplicar_card(res)
        _aplicar_bloco_composicao(res)
        _aplicar_conferencia(res)
    finally:
        _restaurar_protecao(res, estado, selecao)


# ------------------------------------------------------------------ driver

def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_vta_u2ux_"))
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
            raise RuntimeError(f"TRAVA VIOLADA pela UX final: {difs}")

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
        if "VTA_FINAL" not in str(res.Range("C5").Formula):
            raise RuntimeError("C5 deixou de apontar para VTA_FINAL.")
        for linha in range(LINHA_REFS_INICIO, LINHA_REFS_FIM + 1):
            if not res.Rows(linha).Hidden:
                raise RuntimeError(f"Linha {linha} nao ficou oculta.")
            if str(res.Range(f"B{linha}").Formula).strip() in ("", "0"):
                raise RuntimeError(f"B{linha} perdeu a formula tecnica.")
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
    print("VTA-U2 UX final aplicada:", args.destino)


if __name__ == "__main__":
    main()
