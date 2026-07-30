# -*- coding: utf-8 -*-
"""Etapa 26H: novos itens (base zero automatica), temporalidade e dropdown G.

Gravador unico: Microsoft Excel real (padrao dos aplicadores anteriores).
Trabalha em copia temporaria e so promove ao destino apos recalculo completo
sem erros e reabertura verificada. ASCII puro em formulas e mensagens.

O que este owner aplica ao template oficial:

1. posicao_contratual — colunas tecnicas ocultas:
   Z (EH_NOVO_ITEM): padrao canonico Nxxx ("N"/"n" + somente digitos),
   com guardas contra separadores/sinais/notacao cientifica;
   Y (CICLO_NASCIMENTO): primeiro ciclo com QTD_CONTRATADA > 0 (0..4),
   derivado dos aditivos/posicao contratual — sem nova data manual.
2. posicao_contratual!F (QTD_REM_BASE_C0): novo item com base vazia -> 0
   (base zero automatica); item normal preserva o comportamento atual.
3. posicao_contratual!X (CHECK): novos ramos fail-closed para Nxxx —
   base informada diferente de 0 e VU_ORIGINAL ausente/nao numerico.
   Nao corrige input do fiscal; somente acusa.
4. historico_VU!C:G — temporalidade do VU: celula vazia para ciclo anterior
   ao nascimento (nunca zero/VU inventado); VU do ciclo de nascimento =
   VU_ORIGINAL; ciclos seguintes pela razao dos fatores acumulados (J:L).
5. MEMORIA_RESULTADOS — X (C0 fisico) devolve 0 para item nascido apos C0 e
   T27 deixa de exigir VU_C0 de item nascido apos C0 (sem "CALCULO MANUAL
   REQUERIDO" indevido); fail-closed preservado para item sem nascimento
   derivavel.
6. posicao_referencia!I2 (POSICAO ATUAL COMPLETA?) — regra temporal: exige
   QTD_REM_ATUAL somente dos itens que integram o contrato na data de corte
   (nascimento <= ciclo da data); item criado depois nao torna a posicao
   incompleta. Zero explicito segue valido.
7. aditivos!M — mensagem canonica 26H de novo item nao cadastrado.
8. itens_PC!G2:G{cap} — Data Validation com fonte robusta =OPCOES_SIM_NAO
   (nome definido; nao depende de separador regional e sobrevive ao
   round-trip openpyxl da geracao runtime).
9. Orientacoes de cabecalho (itens_Remanesc!A1/B1, aditivos!A1) alinhadas a
   base zero automatica.
"""
from __future__ import annotations

import argparse
import shutil
import sys
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

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from _capacidade_pcs import ULTIMA_LINHA_PCS  # noqa: E402

CAP = ULTIMA_LINHA_PCS  # 5001

XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_VALIDATE_LIST = 3
XL_VALID_ALERT_STOP = 1
XL_BETWEEN = 1
XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105

ABA_MEMORIA = "MEMORIA_RESULTADOS"

# Padrao canonico Nxxx (espelho Excel do regex ^[Nn]\d+$ de _reajuste_utils):
# LEFT="N" (comparacao de texto ja e case-insensitive), resto coercivel a
# numero e sem separadores/sinais/espacos/notacao cientifica.
_F_EH_NOVO_ITEM = (
    '=IF($A2="",FALSE,AND(LEN($A2)>1,LEFT($A2,1)="N",'
    'ISNUMBER(--MID($A2,2,255)),'
    'ISERROR(FIND(".",$A2)),ISERROR(FIND(",",$A2)),'
    'ISERROR(FIND("-",$A2)),ISERROR(FIND("+",$A2)),'
    'ISERROR(SEARCH("E",$A2,2)),ISERROR(FIND(" ",TRIM($A2)))))'
)

_F_CICLO_NASCIMENTO = (
    '=IF($A2="","",IF($E2>0,0,IF($I2>0,1,IF($M2>0,2,'
    'IF($Q2>0,3,IF($U2>0,4,""))))))'
)

_F_F2_NOVA = '=IF(A2="","",IF(itens_Remanesc!B2<>"",itens_Remanesc!B2,IF($Z2,0,"")))'
_F_F2_ANTIGA = '=IF(A2="","",IF(itens_Remanesc!B2="","",itens_Remanesc!B2))'

_MARCADOR_X = '"ALERTA: ITEM_DUPLICADO",'
_RAMOS_X_26H = (
    'IF(AND($Z2,ISNUMBER(itens_Remanesc!B2),itens_Remanesc!B2<>0),'
    '"NOVO ITEM COM BASE DIFERENTE DE 0 - AJUSTAR itens_Remanesc",'
    'IF(AND($Z2,NOT(ISNUMBER(itens_Remanesc!C2))),'
    '"NOVO ITEM SEM VU_ORIGINAL - INFORMAR EM itens_Remanesc",'
)

_MSG_ADITIVOS_ANTIGA = (
    'NOVO ITEM NAO CADASTRADO - CADASTRAR EM itens_Remanesc '
    'COMO N001, N002... (BASE 0 + VU)'
)
_MSG_ADITIVOS_NOVA = (
    'NOVO ITEM NAO CADASTRADO - CADASTRAR EM itens_Remanesc; '
    'QTD_BASE_ORIGINAL sera 0 automaticamente; informar VU_ORIGINAL.'
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


def _aplicar_posicao_contratual(ws) -> int:
    """Colunas tecnicas Y/Z, F com base zero automatica e CHECK X ampliado.

    Retorna a ultima linha de dados (extensao atual das formulas de X).
    """
    estado, selecao = _capturar_protecao(ws)

    if str(ws.Range("X1").Value or "") != "CHECK_POSICAO_CONTRATUAL":
        raise ValueError("posicao_contratual fora do layout oficial (X1).")
    for coluna in ("Y", "Z"):
        if ws.Range(f"{coluna}1").Value not in (None, ""):
            raise RuntimeError(
                f"posicao_contratual!{coluna}1 ja possui conteudo; "
                "reaplicacao 26H recusada (fail-closed)."
            )

    ultima = _ultima_linha_formula(ws, "X", 260)
    if ultima < 100:
        raise RuntimeError(f"posicao_contratual!X com extensao inesperada ({ultima}).")

    # Z (EH_NOVO_ITEM) primeiro: Y nao depende de Z; F e X dependem de Z.
    ws.Range("Z1").Value = "EH_NOVO_ITEM"
    ws.Range(f"Z2:Z{ultima}").Formula = _F_EH_NOVO_ITEM
    ws.Range("Y1").Value = "CICLO_NASCIMENTO"
    ws.Range(f"Y2:Y{ultima}").Formula = _F_CICLO_NASCIMENTO

    f2 = str(ws.Range("F2").Formula or "")
    if f2 != _F_F2_ANTIGA:
        raise RuntimeError(f"posicao_contratual!F2 fora do padrao esperado: {f2!r}")
    ws.Range(f"F2:F{ultima}").Formula = _F_F2_NOVA

    x2 = str(ws.Range("X2").Formula or "")
    if _MARCADOR_X not in x2 or "NOVO ITEM" in x2:
        raise RuntimeError("posicao_contratual!X2 fora do padrao esperado p/ 26H.")
    # Insere os ramos 26H apos o ramo de duplicidade; fecha os 2 IFs extras.
    nova_x = x2.replace(_MARCADOR_X, _MARCADOR_X + _RAMOS_X_26H, 1)
    if not nova_x.endswith("))))"):
        raise RuntimeError("posicao_contratual!X2 sem o fechamento esperado.")
    nova_x = nova_x[:-4] + "))))))"
    ws.Range(f"X2:X{ultima}").Formula = nova_x

    ws.Columns("Y:Z").Hidden = True  # oculta (nunca veryHidden)

    _restaurar_protecao(ws, estado, selecao)
    return ultima


def _aplicar_historico_vu(ws) -> None:
    """VU temporal: vazio antes do nascimento; nascimento = VU_ORIGINAL."""
    estado, selecao = _capturar_protecao(ws)

    cabecalhos = {c: str(ws.Range(f"{c}1").Value or "") for c in "CDEFG"}
    esperado = {"C": "VU_C0", "D": "VU_C1", "E": "VU_C2", "F": "VU_C3", "G": "VU_C4"}
    if cabecalhos != esperado:
        raise ValueError(f"historico_VU fora do layout oficial C:G ({cabecalhos}).")

    ultima = _ultima_linha_formula(ws, "C", 260)
    if ultima < 100:
        raise RuntimeError(f"historico_VU!C com extensao inesperada ({ultima}).")

    nasc = "posicao_contratual!$Y2"
    ws.Range(f"C2:C{ultima}").Formula = (
        f'=IF(itens_Remanesc!A2="","",IF({nasc}="","",'
        f'IF({nasc}>0,"",itens_Remanesc!C2)))'
    )
    # VU_Ck (k=1..4): fator do ciclo em $L$(k+2); fator do nascimento por INDEX.
    fator_nasc = f'INDEX($L$2:$L$6,{nasc}+1)'
    for k, coluna in ((1, "D"), (2, "E"), (3, "F"), (4, "G")):
        fator_k = f"$L${k + 2}"
        ws.Range(f"{coluna}2:{coluna}{ultima}").Formula = (
            f'=IF(OR($A2="",{nasc}=""),"",'
            f'IF({nasc}>{k},"",'
            f'IF({nasc}={k},IF(ISNUMBER(itens_Remanesc!C2),itens_Remanesc!C2,""),'
            f'IF(OR(NOT(ISNUMBER(itens_Remanesc!C2)),NOT(ISNUMBER({fator_k})),'
            f'NOT(ISNUMBER({fator_nasc}))),"",'
            f'ROUND(itens_Remanesc!C2*{fator_k}/{fator_nasc},2)))))'
        )

    _restaurar_protecao(ws, estado, selecao)


def _aplicar_memoria(ws) -> None:
    """Item nascido apos C0 nao dispara fail-closed indevido em X/T27."""
    estado, selecao = _capturar_protecao(ws)

    x2 = str(ws.Range("X2").Formula or "")
    prefixo = '=IF(posicao_contratual!$A2="",0,'
    if not x2.startswith(prefixo) or "$Y2" in x2:
        raise RuntimeError(f"MEMORIA!X2 fora do padrao esperado p/ 26H: {x2!r}")
    nova_x = (
        prefixo
        + 'IF(AND(ISNUMBER(posicao_contratual!$Y2),posicao_contratual!$Y2>0),0,'
        + x2[len(prefixo):-1]
        + '))'
    )
    ws.Range("X2:X201").Formula = nova_x

    t27 = str(ws.Range("T27").Formula or "")
    alvo = "ISNUMBER(historico_VU!$C$2:$C$201)"
    if alvo not in t27 or "$Y$" in t27:
        raise RuntimeError(f"MEMORIA!T27 fora do padrao esperado p/ 26H: {t27!r}")
    ws.Range("T27").Formula = t27.replace(
        alvo,
        "(ISNUMBER(historico_VU!$C$2:$C$201)"
        "+(posicao_contratual!$Y$2:$Y$201<>\"\")"
        "*(posicao_contratual!$Y$2:$Y$201>0))",
        1,
    )

    _restaurar_protecao(ws, estado, selecao)


_MARCADOR_F_REF = 'IF(D2="","REMANESCENTE DE REFERENCIA INDISPONIVEL"'
# Item posterior a referencia: nascido em ciclo posterior ao da referencia OU
# novo item (Nxxx) cuja contratada NA DATA da referencia (C) e zero — cobre o
# aditivo formalizado depois da data de referencia dentro do mesmo ciclo.
_RAMO_F_POSTERIOR = (
    'IF(OR(AND(posicao_contratual!$Y2<>"",'
    'posicao_contratual!$Y2>IFERROR(VALUE(MID($I$4,2,1)),-1)),'
    'AND(posicao_contratual!$Z2,ISNUMBER(C2),C2<=0)),'
    '"ITEM POSTERIOR A REFERENCIA",'
)


def _aplicar_posicao_referencia(ws) -> None:
    """POSICAO ATUAL COMPLETA? com filtragem pela vigencia na data de corte.

    Tambem qualifica o CHECK por item: item nascido depois do ciclo de
    referencia recebe "ITEM POSTERIOR A REFERENCIA" (informativo) em vez de
    "REMANESCENTE DE REFERENCIA INDISPONIVEL" (que sugere pendencia).
    """
    estado, selecao = _capturar_protecao(ws)

    i2 = str(ws.Range("I2").Formula or "")
    if "COUNT($B$2:$B$200)>=COUNTIF" not in i2:
        raise RuntimeError(f"posicao_referencia!I2 fora do padrao esperado: {i2!r}")
    ciclo_data = 'IFERROR(VALUE(MID($I$1,2,1)),-1)'
    ws.Range("I2").Formula = (
        '=AND(ISNUMBER(CONTROLE!$B$3),'
        'OR($I$1="C0",$I$1="C1",$I$1="C2",$I$1="C3",$I$1="C4"),'
        'COUNTIF(itens_Remanesc!$A$2:$A$200,"<>")>0,'
        'SUMPRODUCT((itens_Remanesc!$A$2:$A$200<>"")*'
        '(posicao_contratual!$Y$2:$Y$200<>"")*'
        f'(posicao_contratual!$Y$2:$Y$200<={ciclo_data})*'
        '(1-ISNUMBER($B$2:$B$200)))=0)'
    )

    f2 = str(ws.Range("F2").Formula or "")
    if _MARCADOR_F_REF not in f2 or "POSTERIOR" in f2:
        raise RuntimeError(f"posicao_referencia!F2 fora do padrao esperado: {f2!r}")
    ultima = _ultima_linha_formula(ws, "F", 260)
    nova_f = f2.replace(_MARCADOR_F_REF, _RAMO_F_POSTERIOR + _MARCADOR_F_REF, 1)
    if not nova_f.endswith(")"):
        raise RuntimeError("posicao_referencia!F2 sem fechamento esperado.")
    nova_f = nova_f[:-1] + "))"
    faixa_f = ws.Range(f"F2:F{ultima}")
    # A coluna F esta formatada como Texto ("@"); sem voltar a Geral a formula
    # seria armazenada como literal (celula-texto) e o CHECK morreria. Alguns
    # Excel rejeitam NumberFormat="General" (0x800A03EC); NumberFormatLocal e
    # o fallback aceito na instalacao pt-BR.
    try:
        faixa_f.NumberFormat = "General"
    except pywintypes.com_error:
        faixa_f.NumberFormatLocal = "Geral"
    faixa_f.Formula = nova_f
    if str(ws.Range("F2").Value or "").startswith("="):
        raise RuntimeError("posicao_referencia!F2 armazenou a formula como texto.")

    _restaurar_protecao(ws, estado, selecao)


def _aplicar_aditivos(ws) -> None:
    estado, selecao = _capturar_protecao(ws)

    m2 = str(ws.Range("M2").Formula or "")
    if _MSG_ADITIVOS_ANTIGA not in m2:
        raise RuntimeError(f"aditivos!M2 sem a mensagem esperada: {m2!r}")
    ultima = _ultima_linha_formula(ws, "M", 260)
    nova = m2.replace(_MSG_ADITIVOS_ANTIGA, _MSG_ADITIVOS_NOVA, 1)
    ws.Range(f"M2:M{ultima}").Formula = nova

    a1 = str(ws.Range("A1").Value or "")
    if "BASE 0 + VU" in a1:
        ws.Range("A1").Value = (
            "ITEM\n(NOVO ITEM: cadastrar em itens_Remanesc como N001, N002...; "
            "QTD_BASE_ORIGINAL fica 0 automaticamente; informar o VU_ORIGINAL)"
        )

    _restaurar_protecao(ws, estado, selecao)


def _aplicar_itens_remanesc(ws) -> None:
    estado, selecao = _capturar_protecao(ws)

    a1 = str(ws.Range("A1").Value or "")
    if "QTD_BASE_ORIGINAL=0" in a1:
        ws.Range("A1").Value = (
            "ITEM\n(NOVO ITEM: use N001, N002...; QTD_BASE_ORIGINAL fica 0 "
            "automaticamente - nao digite; informe o VU_ORIGINAL)"
        )
    b1 = str(ws.Range("B1").Value or "")
    if "0 PARA NOVO ITEM FORMALIZADO" in b1:
        ws.Range("B1").Value = (
            "QTD_BASE_ORIGINAL\n(NOVO ITEM: deixar vazio - assume 0 "
            "automaticamente)"
        )

    _restaurar_protecao(ws, estado, selecao)


def _aplicar_dropdown_itens_pc(ws) -> None:
    validacao = ws.Range(f"G2:G{CAP}").Validation
    try:
        validacao.Delete()
    except Exception:
        pass
    validacao.Add(
        Type=XL_VALIDATE_LIST,
        AlertStyle=XL_VALID_ALERT_STOP,
        Operator=XL_BETWEEN,
        Formula1="=OPCOES_SIM_NAO",
    )
    validacao.IgnoreBlank = True
    validacao.InCellDropdown = True
    validacao.ErrorMessage = "Selecione Sim ou Nao."


def _verificar_sem_erros(wb) -> None:
    import pywintypes

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
        pc = wb.Worksheets("posicao_contratual")
        if str(pc.Range("Z1").Value or "") != "EH_NOVO_ITEM":
            raise RuntimeError("posicao_contratual!Z1 (EH_NOVO_ITEM) ausente.")
        if str(pc.Range("Y1").Value or "") != "CICLO_NASCIMENTO":
            raise RuntimeError("posicao_contratual!Y1 (CICLO_NASCIMENTO) ausente.")
        if "$Z2" not in str(pc.Range("F2").Formula or ""):
            raise RuntimeError("posicao_contratual!F2 sem a base zero automatica.")
        if "NOVO ITEM COM BASE DIFERENTE DE 0" not in str(pc.Range("X2").Formula or ""):
            raise RuntimeError("posicao_contratual!X2 sem os ramos 26H.")
        hv = wb.Worksheets("historico_VU")
        if "posicao_contratual!$Y2" not in str(hv.Range("C2").Formula or ""):
            raise RuntimeError("historico_VU!C2 sem a temporalidade 26H.")
        pr = wb.Worksheets("posicao_referencia")
        if "SUMPRODUCT" not in str(pr.Range("I2").Formula or ""):
            raise RuntimeError("posicao_referencia!I2 sem a regra temporal 26H.")
        if "ITEM POSTERIOR A REFERENCIA" not in str(pr.Range("F2").Formula or ""):
            raise RuntimeError("posicao_referencia!F2 sem o ramo de item posterior.")
        if str(pr.Range("F2").Value or "").startswith("="):
            raise RuntimeError("posicao_referencia!F2 virou texto literal.")
        ad = wb.Worksheets("aditivos")
        if "sera 0 automaticamente" not in str(ad.Range("M2").Formula or ""):
            raise RuntimeError("aditivos!M2 sem a mensagem 26H.")
        mem = wb.Worksheets(ABA_MEMORIA)
        if "$Y$2:$Y$201" not in str(mem.Range("T27").Formula or ""):
            raise RuntimeError("MEMORIA!T27 sem o filtro de nascimento 26H.")
        ipc = wb.Worksheets("itens_PC")
        f1 = str(ipc.Range("G2").Validation.Formula1 or "")
        if "OPCOES_SIM_NAO" not in f1:
            raise RuntimeError(f"itens_PC!G2 sem DV =OPCOES_SIM_NAO ({f1!r}).")
    finally:
        wb.Close(SaveChanges=False)


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_26h_"))
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

        _aplicar_posicao_contratual(wb.Worksheets("posicao_contratual"))
        _aplicar_historico_vu(wb.Worksheets("historico_VU"))
        _aplicar_memoria(wb.Worksheets(ABA_MEMORIA))
        _aplicar_posicao_referencia(wb.Worksheets("posicao_referencia"))
        _aplicar_aditivos(wb.Worksheets("aditivos"))
        _aplicar_itens_remanesc(wb.Worksheets("itens_Remanesc"))
        _aplicar_dropdown_itens_pc(wb.Worksheets("itens_PC"))

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
            # Dispatch dinamico pode materializar Close como resultado bool;
            # o arquivo ja esta salvo e a integridade e arbitrada adiante
            # por _verificar_destino (reabertura real sem reparo).
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
    print("Etapa 26H aplicada:", args.destino)


if __name__ == "__main__":
    main()
