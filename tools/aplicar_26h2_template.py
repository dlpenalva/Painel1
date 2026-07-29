# -*- coding: utf-8 -*-
"""Etapa 26H.2: correcoes dos achados da auditoria independente da 26H/26H.1.

Gravador unico: Microsoft Excel real (padrao dos aplicadores anteriores).
Trabalha em copia temporaria e so promove ao destino apos recalculo completo
sem erros e reabertura verificada. ASCII puro em formulas e mensagens.

O que este owner aplica ao template oficial:

1. BLOQUEADOR off-by-one B201: a automacao 26H.1 alcancou B2:B201 por herdar
   a extensao da coluna D (que inclui a linha 201 do total dinamico). A
   capacidade canonica de cadastro e 2:200 (A2:A200 em toda a cadeia:
   posicao_contratual/posicao_referencia/historico_VU 2:200, aditivos
   COUNTIF A2:A200, leitor range(2,201)). Este owner LIMPA B201; a linha 201
   volta ao estado do template homologado 26H (sem formula, sem valor) e nao
   oferece zero automatico enganoso. A linha 201 e neutralizada tambem nos
   totais/consumidores: mesmo que alguem force A201, o valor e ignorado e nao
   pode chegar parcialmente a itens_RC ou MEMORIA_RESULTADOS.
2. Padrao canonico Nxxx = "N" + EXATAMENTE 3 algarismos (N001..N999),
   espelho Excel do regex ^[Nn]\\d{3}$ de _reajuste_utils:
   LEN(TRIM)=4 + LEFT="N" (comparacao de texto ja case-insensitive) +
   MID(TRIM,2,3) coercivel a numero + guardas ./,/-/+/E/espaco-interno.
   Reaplicado em itens_Remanesc!B2:B200 (base zero visual) e em
   posicao_contratual!Z2:Z200 (EH_NOVO_ITEM). Consumidores de Z (F/X/
   historico_VU/MEMORIA/posicao_referencia) nao mudam.
3. UX aditivos!M2:M200: WrapText=True + coluna M mais larga + altura das
   linhas da grade ajustada para a mensagem de novo item caber integralmente
   (3 linhas visiveis), sem coluna exageradamente larga.
4. Verificacao adversarial: forca N201 em itens_Remanesc!A201, recalcula e
   exige workbook inteiro sem celulas de erro e sem propagacao parcial;
   reabre sem reparo e repete as invariantes estruturais.
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

# Capacidade canonica de cadastro (linhas de input): 2:200.
ULTIMA_LINHA_CADASTRO = 200

# Espelho Excel do regex ^[Nn]\d{3}$ (teste sobre $A da linha).
_TESTE_NXXX = (
    'AND(LEN(TRIM($A2))=4,LEFT(TRIM($A2),1)="N",'
    'ISNUMBER(--MID(TRIM($A2),2,3)),'
    'ISERROR(FIND(".",$A2)),ISERROR(FIND(",",$A2)),'
    'ISERROR(FIND("-",$A2)),ISERROR(FIND("+",$A2)),'
    'ISERROR(SEARCH("E",$A2,2)),ISERROR(FIND(" ",TRIM($A2))))'
)
_F_B_BASE_ZERO_262 = f'=IF({_TESTE_NXXX},0,"")'
_F_Z_262 = f'=IF($A2="",FALSE,{_TESTE_NXXX})'

# Linha 201 e exclusivamente a linha final de total quando A200 esta ocupada.
_COLUNAS_TOTAL_REMANESC = ("D", "F", "H", "J", "L", "N", "P", "R")
_COLUNAS_VAZIAS_RC_202 = ("B", "C", "E", "F", "H", "I", "K", "L", "N", "O")
_COLUNAS_VAZIAS_MEMORIA_253 = ("J", "K", "L", "M", "N")

# UX aditivos!M: mensagem em ate 3 linhas visiveis.
LARGURA_COLUNA_M = 45.0
ALTURA_LINHAS_ADITIVOS = 45.0

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


def _aplicar_itens_remanesc(ws) -> None:
    """Fronteira 200 + Nxxx em B2:B200; linha 201 e total, nunca input."""
    estado, selecao = _capturar_protecao(ws)

    b2 = str(ws.Range("B2").Formula or "")
    if "MID($A2,2,255)" not in b2 and "MID(TRIM($A2),2,3)" not in b2:
        raise RuntimeError(f"itens_Remanesc!B2 fora do padrao 26H.1/26H.2: {b2!r}")
    b201 = str(ws.Range("B201").Formula or "")
    if b201 and "MID($A201,2,255)" not in b201:
        raise RuntimeError(f"itens_Remanesc!B201 inesperado: {b201!r}")
    ws.Range("B201").ClearContents()
    ws.Range(f"B2:B{ULTIMA_LINHA_CADASTRO}").Formula = _F_B_BASE_ZERO_262

    for coluna in _COLUNAS_TOTAL_REMANESC:
        celula = ws.Range(f"{coluna}201")
        atual = str(celula.Formula or "")
        destino = (
            f'=IF($A$200<>"",ROUND(SUMIF($A$2:$A$200,"<>",'
            f'${coluna}$2:${coluna}$200),2),"")'
        )
        if "A201" not in atual and atual != destino:
            raise RuntimeError(
                f"itens_Remanesc!{coluna}201 fora do total esperado: {atual!r}"
            )
        celula.Formula = destino

    u201 = str(ws.Range("U201").Formula or "")
    destino_u201 = '=IF($A$200<>"","TOTAL","")'
    if "A201" not in u201 and u201 != destino_u201:
        raise RuntimeError(f"itens_Remanesc!U201 fora do total esperado: {u201!r}")
    ws.Range("U201").Formula = destino_u201

    _restaurar_protecao(ws, estado, selecao)


def _aplicar_posicao_contratual_z(ws) -> None:
    estado, selecao = _capturar_protecao(ws)

    if str(ws.Range("Z1").Value or "") != "EH_NOVO_ITEM":
        raise RuntimeError("posicao_contratual!Z1 (EH_NOVO_ITEM) ausente.")
    z2 = str(ws.Range("Z2").Formula or "")
    if "MID($A2,2,255)" not in z2 and "MID(TRIM($A2),2,3)" not in z2:
        raise RuntimeError(f"posicao_contratual!Z2 fora do padrao 26H/26H.2: {z2!r}")
    if str(ws.Range("Z201").Formula or ""):
        raise RuntimeError("posicao_contratual!Z201 nao deveria ter formula.")
    ws.Range(f"Z2:Z{ULTIMA_LINHA_CADASTRO}").Formula = _F_Z_262

    _restaurar_protecao(ws, estado, selecao)


def _neutralizar_consumidores_da_linha_201(wb) -> None:
    """Remove qualquer suporte parcial a itens_Remanesc!A201."""
    rc = wb.Worksheets("itens_RC")
    estado_rc, selecao_rc = _capturar_protecao(rc)
    a202 = str(rc.Range("A202").Formula or "")
    destino_a202 = '=IF(itens_Remanesc!$A$200<>"","TOTAL","")'
    if "itens_Remanesc!A201" not in a202 and a202 != destino_a202:
        raise RuntimeError(f"itens_RC!A202 fora da fronteira esperada: {a202!r}")
    rc.Range("A202").Formula = destino_a202
    for coluna in _COLUNAS_VAZIAS_RC_202:
        celula = rc.Range(f"{coluna}202")
        atual = str(celula.Formula or "")
        if not atual:
            raise RuntimeError(f"itens_RC!{coluna}202 perdeu a formula estrutural.")
        if atual != '=""' and "201" not in atual:
            raise RuntimeError(
                f"itens_RC!{coluna}202 fora da fronteira esperada: {atual!r}"
            )
        celula.Formula = '=""'
    _restaurar_protecao(rc, estado_rc, selecao_rc)

    memoria = wb.Worksheets("MEMORIA_RESULTADOS")
    estado_mem, selecao_mem = _capturar_protecao(memoria)
    for coluna in _COLUNAS_VAZIAS_MEMORIA_253:
        celula = memoria.Range(f"{coluna}253")
        atual = str(celula.Formula or "")
        if not atual:
            raise RuntimeError(
                f"MEMORIA_RESULTADOS!{coluna}253 perdeu a formula estrutural."
            )
        if atual != '=""' and "253" not in atual and "201" not in atual:
            raise RuntimeError(
                f"MEMORIA_RESULTADOS!{coluna}253 inesperada: {atual!r}"
            )
        celula.Formula = '=""'
    _restaurar_protecao(memoria, estado_mem, selecao_mem)


def _aplicar_ux_aditivos_m(ws) -> None:
    """Mensagem de novo item integralmente visivel (WrapText + dimensoes)."""
    estado, selecao = _capturar_protecao(ws)

    m2 = str(ws.Range("M2").Formula or "")
    if "sera 0 automaticamente" not in m2:
        raise RuntimeError("aditivos!M2 sem a mensagem 26H.")
    faixa = ws.Range(f"M2:M{ULTIMA_LINHA_CADASTRO}")
    faixa.WrapText = True
    ws.Columns("M").ColumnWidth = LARGURA_COLUNA_M
    ws.Rows(f"2:{ULTIMA_LINHA_CADASTRO}").RowHeight = ALTURA_LINHAS_ADITIVOS

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


def _verificar_cenario_n201(wb, excel) -> None:
    """Prova negativa: N201 forcado e ignorado, sem IFERROR ou erro oculto."""
    ir = wb.Worksheets("itens_Remanesc")
    pc = wb.Worksheets("posicao_contratual")
    rc = wb.Worksheets("itens_RC")
    memoria = wb.Worksheets("MEMORIA_RESULTADOS")
    original = ir.Range("A201").Value
    if original not in (None, ""):
        raise RuntimeError("itens_Remanesc!A201 deveria estar vazio antes da prova.")
    try:
        ir.Range("A201").Value = "N201"
        excel.CalculateFullRebuild()
        if str(ir.Range("B201").Formula or ""):
            raise RuntimeError("N201 recebeu automacao parcial em itens_Remanesc!B201.")
        if str(pc.Range("C201").Formula or "") or str(pc.Range("Z201").Formula or ""):
            raise RuntimeError("N201 recebeu automacao parcial em posicao_contratual!201.")
        if str(rc.Range("A202").Value or "") == "N201":
            raise RuntimeError("N201 propagou para itens_RC!A202.")
        if str(memoria.Range("J253").Value or ""):
            raise RuntimeError("N201 propagou para MEMORIA_RESULTADOS!J253.")
        _verificar_sem_erros(wb)
    finally:
        ir.Range("A201").ClearContents()
        excel.CalculateFullRebuild()


def _verificar_destino(excel, caminho: Path) -> None:
    wb = excel.Workbooks.Open(str(caminho), UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
    try:
        ir = wb.Worksheets("itens_Remanesc")
        for linha in (2, 100, 199, 200):
            f = str(ir.Range(f"B{linha}").Formula or "")
            if "MID(TRIM($A" not in f:
                raise RuntimeError(f"itens_Remanesc!B{linha} sem o padrao 26H.2 ({f!r}).")
        if str(ir.Range("B201").Formula or "") not in ("",):
            raise RuntimeError("itens_Remanesc!B201 ainda possui conteudo (off-by-one).")
        for coluna in _COLUNAS_TOTAL_REMANESC:
            if "A201" in str(ir.Range(f"{coluna}201").Formula or ""):
                raise RuntimeError(
                    f"itens_Remanesc!{coluna}201 ainda depende da linha invalida."
                )
        if "A201" in str(ir.Range("U201").Formula or ""):
            raise RuntimeError("itens_Remanesc!U201 ainda depende da linha invalida.")
        pc = wb.Worksheets("posicao_contratual")
        if "LEN(TRIM($A2))=4" not in str(pc.Range("Z2").Formula or ""):
            raise RuntimeError("posicao_contratual!Z2 sem o padrao N+3 digitos.")
        ad = wb.Worksheets("aditivos")
        if not bool(ad.Range("M2").WrapText):
            raise RuntimeError("aditivos!M2 sem WrapText.")
        rc = wb.Worksheets("itens_RC")
        if "A201" in str(rc.Range("A202").Formula or ""):
            raise RuntimeError("itens_RC!A202 ainda consome a linha invalida.")
        for coluna in _COLUNAS_VAZIAS_RC_202:
            if str(rc.Range(f"{coluna}202").Formula or "") != '=""':
                raise RuntimeError(f"itens_RC!{coluna}202 nao foi neutralizada.")
        memoria = wb.Worksheets("MEMORIA_RESULTADOS")
        for coluna in _COLUNAS_VAZIAS_MEMORIA_253:
            if str(memoria.Range(f"{coluna}253").Formula or "") != '=""':
                raise RuntimeError(
                    f"MEMORIA_RESULTADOS!{coluna}253 nao foi neutralizada."
                )
        d202 = rc.Range("D202").Value
        if isinstance(d202, int) and d202 < 0:  # codigos de erro COM sao int negativos
            raise RuntimeError(f"itens_RC!D202 com erro ({d202}).")
        _verificar_sem_erros(wb)
    finally:
        wb.Close(SaveChanges=False)


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_26h2_"))
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

        _aplicar_itens_remanesc(wb.Worksheets("itens_Remanesc"))
        _aplicar_posicao_contratual_z(wb.Worksheets("posicao_contratual"))
        _neutralizar_consumidores_da_linha_201(wb)
        _aplicar_ux_aditivos_m(wb.Worksheets("aditivos"))

        try:
            wb.Worksheets(aba_ativa).Activate()
        except Exception:
            pass

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)
        _verificar_cenario_n201(wb, excel)
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
    print("Etapa 26H.2 aplicada:", args.destino)


if __name__ == "__main__":
    main()
