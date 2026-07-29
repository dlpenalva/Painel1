# -*- coding: utf-8 -*-
"""Etapa 26H.1: base zero VISUAL do novo item em itens_Remanesc!B.

Fecha o requisito de UX do Gate 26H: ao digitar N001 em itens_Remanesc!A,
o fiscal passa a VER 0 em QTD_BASE_ORIGINAL na propria aba (a matematica
ja coagia a zero nas abas derivadas, mas a celula B ficava vazia).

Gravador unico: Microsoft Excel real (padrao dos aplicadores anteriores).
Trabalha em copia temporaria e so promove ao destino apos recalculo completo
sem erros e reabertura verificada. ASCII puro em formulas e mensagens.

O que este owner aplica ao template oficial:

1. itens_Remanesc!B2:B200 — formula pre-semeada: novo item (padrao canonico
   Nxxx, o MESMO espelho Excel usado em posicao_contratual!Z) -> 0; qualquer
   outro caso -> "" (visual identico ao vazio). O fiscal digita/cola por
   cima normalmente (a formula e destruida pela digitacao — comportamento
   esperado de coluna de input); se sobrescrever o Nxxx com quantidade
   diferente de 0, o CHECK posicao_contratual!X acusa (ramo 26H existente).
   A geracao oficial preserva a formula (_limpar_residuos nunca toca
   formulas) e o leitor le o valor em cache (data_only=True), com
   qtd_base_efetiva cobrindo "" no lado Python.
2. posicao_contratual!C2:C200 — coercao explicita ""->0 na referencia a
   itens_Remanesc!B (formula que devolve "" NAO e celula vazia na aritmetica
   do Excel; sem a coercao, E2=ROUND(C2+D2,2) viraria #VALUE!). Preserva
   exatamente o numero exibido hoje (B vazio -> 0).
3. posicao_referencia!C2:C200 — mesma coercao ""->0 na soma com SUMIFS.

Consumidores auditados que ja toleram "" por construcao (sem mudanca):
posicao_contratual!F (B<>""), posicao_contratual!X (ISNUMBER),
itens_Remanesc!D (B2="" guard), MEMORIA_RESULTADOS!J44 (COUNTIFS ""),
MEMORIA_RESULTADOS!K54+ (quadro com guard K=""; item sem base informada
passa a "INCOMPLETO" em vez de qtd 0 — fail-closed mais correto).
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

_RPC_E_CALL_REJECTED = -2147418111


def _com_ocupado_retry(funcao, tentativas: int = 15, espera: float = 2.0):
    """Reexecuta chamadas COM rejeitadas enquanto o Excel recalcula (RPC busy)."""
    for tentativa in range(tentativas):
        try:
            return funcao()
        except pywintypes.com_error as exc:
            if exc.hresult == _RPC_E_CALL_REJECTED and tentativa < tentativas - 1:
                time.sleep(espera)
                continue
            raise


XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

# Espelho Excel do regex ^[Nn]\d+$ — IDENTICO ao teste de posicao_contratual!Z
# (26H), aplicado a $A da propria linha de itens_Remanesc. A2 vazio cai no
# ramo "" (LEN=0), entao a grade vazia continua visualmente vazia.
_F_B_BASE_ZERO = (
    '=IF(AND(LEN($A2)>1,LEFT($A2,1)="N",'
    'ISNUMBER(--MID($A2,2,255)),'
    'ISERROR(FIND(".",$A2)),ISERROR(FIND(",",$A2)),'
    'ISERROR(FIND("-",$A2)),ISERROR(FIND("+",$A2)),'
    'ISERROR(SEARCH("E",$A2,2)),ISERROR(FIND(" ",TRIM($A2)))),0,"")'
)

_F_C_PC_ANTIGA = '=IF(A2="","",itens_Remanesc!B2)'
_F_C_PC_NOVA = '=IF(A2="","",IF(itens_Remanesc!B2="",0,itens_Remanesc!B2))'

_MARCADOR_C_REF_ANTIGO = 'ROUND(itens_Remanesc!B2+SUMIFS('
_MARCADOR_C_REF_NOVO = (
    'ROUND(IF(itens_Remanesc!B2="",0,itens_Remanesc!B2)+SUMIFS('
)

_PROT_FLAGS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
    "AllowFiltering", "AllowUsingPivotTables",
)


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


def _ultima_linha_formula(ws, coluna: str, teto: int) -> int:
    r = teto
    while r > 1 and not str(ws.Range(f"{coluna}{r}").Formula or "").startswith("="):
        r -= 1
    return r


def _aplicar_itens_remanesc_b(ws) -> int:
    """Pre-semeia a formula de base zero visual em B2:B{ultima}."""
    estado, selecao = _capturar_protecao(ws)

    b1 = str(ws.Range("B1").Value or "")
    if "QTD_BASE_ORIGINAL" not in b1:
        raise ValueError(f"itens_Remanesc!B1 fora do layout oficial: {b1!r}")
    # 26H.2 (auditoria): capacidade canonica de cadastro = linhas 2:200
    # (A2:A200 em toda a cadeia). NAO inferir a extensao da coluna D — ela
    # alcanca 201 por causa da linha extra do total dinamico, que NAO e
    # linha valida de cadastro (off-by-one B201).
    ultima = 200
    if not str(ws.Range("D200").Formula or "").startswith("="):
        raise RuntimeError("itens_Remanesc!D200 sem formula; layout inesperado.")
    # Fail-closed: coluna de input deve estar 100% livre (sem valores demo e
    # sem formula previa) antes da pre-semeadura.
    for linha in range(2, ultima + 1):
        v = ws.Range(f"B{linha}").Formula
        if v not in (None, ""):
            raise RuntimeError(
                f"itens_Remanesc!B{linha} ja possui conteudo ({v!r}); "
                "reaplicacao 26H.1 recusada (fail-closed)."
            )
    ws.Range(f"B2:B{ultima}").Formula = _F_B_BASE_ZERO

    _restaurar_protecao(ws, estado, selecao)
    return ultima


def _aplicar_posicao_contratual_c(ws) -> None:
    estado, selecao = _capturar_protecao(ws)

    c2 = str(ws.Range("C2").Formula or "")
    if c2 != _F_C_PC_ANTIGA:
        raise RuntimeError(f"posicao_contratual!C2 fora do padrao esperado: {c2!r}")
    # Extensao da PROPRIA grade (C termina em 200; itens_Remanesc vai a 201
    # por causa da linha extra do total dinamico — nao herdar aquela extensao).
    ultima = _ultima_linha_formula(ws, "C", 260)
    ws.Range(f"C2:C{ultima}").Formula = _F_C_PC_NOVA

    _restaurar_protecao(ws, estado, selecao)


def _aplicar_posicao_referencia_c(ws) -> None:
    estado, selecao = _capturar_protecao(ws)

    c2 = str(ws.Range("C2").Formula or "")
    if _MARCADOR_C_REF_ANTIGO not in c2 or 'B2="",0' in c2:
        raise RuntimeError(f"posicao_referencia!C2 fora do padrao esperado: {c2!r}")
    ultima = _ultima_linha_formula(ws, "C", 260)
    nova = c2.replace(_MARCADOR_C_REF_ANTIGO, _MARCADOR_C_REF_NOVO, 1)
    ws.Range(f"C2:C{ultima}").Formula = nova

    _restaurar_protecao(ws, estado, selecao)


def _verificar_sem_erros(wb) -> None:
    problemas = []
    for ws in wb.Worksheets:
        for tipo in (XL_CELLTYPE_FORMULAS, XL_CELLTYPE_CONSTANTS):
            try:
                celulas = ws.UsedRange.SpecialCells(tipo, XL_ERRORS)
                problemas.append(f"{ws.Name}!{celulas.Address}")
            except pywintypes.com_error:
                continue
    if problemas:
        raise RuntimeError(f"Erros de formula apos recalculo: {problemas}")


def _verificar_destino(excel, caminho: Path) -> None:
    wb = excel.Workbooks.Open(str(caminho), UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
    try:
        ir = wb.Worksheets("itens_Remanesc")
        b2 = str(ir.Range("B2").Formula or "")
        if not b2.startswith("=") or "MID($A2,2,255)" not in b2:
            raise RuntimeError(f"itens_Remanesc!B2 sem a formula 26H.1 ({b2!r}).")
        pc = wb.Worksheets("posicao_contratual")
        if str(pc.Range("C2").Formula or "") != _F_C_PC_NOVA:
            raise RuntimeError("posicao_contratual!C2 sem a coercao 26H.1.")
        pr = wb.Worksheets("posicao_referencia")
        if 'B2="",0' not in str(pr.Range("C2").Formula or ""):
            raise RuntimeError("posicao_referencia!C2 sem a coercao 26H.1.")
    finally:
        wb.Close(SaveChanges=False)


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_26h1_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    salvo = False
    try:
        wb = excel.Workbooks.Open(str(tmp_xlsx), UpdateLinks=0, ReadOnly=False, CorruptLoad=0)

        def _preparar():
            excel.ScreenUpdating = False
            excel.Calculation = XL_CALC_MANUAL
            return wb.ActiveSheet.Name

        aba_ativa = _com_ocupado_retry(_preparar)

        _aplicar_itens_remanesc_b(wb.Worksheets("itens_Remanesc"))
        _aplicar_posicao_contratual_c(wb.Worksheets("posicao_contratual"))
        _aplicar_posicao_referencia_c(wb.Worksheets("posicao_referencia"))

        try:
            wb.Worksheets(aba_ativa).Activate()
        except Exception:
            pass

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)
        wb.Save()
        salvo = True
        try:
            wb.Close(SaveChanges=False)
        except TypeError:
            # Dispatch dinamico pode materializar Close como bool; integridade
            # arbitrada adiante por _verificar_destino (reabertura sem reparo).
            pass
        wb = None
        _verificar_destino(excel, tmp_xlsx)
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
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
    print("Etapa 26H.1 aplicada:", args.destino)


if __name__ == "__main__":
    main()
