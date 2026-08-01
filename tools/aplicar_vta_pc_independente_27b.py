"""Etapa 27B — desacoplamento definitivo entre VTA e pagamento dos PCs.

Aplica sobre o template oficial, numa unica sessao Excel COM (Microsoft
Excel real; NAO usa openpyxl para salvar), as correcoes semanticas da
coluna G (PC_PAGO_A_CONTRATADA) de itens_PC:

  itens_PC!I2:I5001 : VALOR_ATUALIZADO_EM_ANALISE — G="Nao" OU G vazio
                      viram potencial (antes: somente "Nao"; vazio zerava
                      silenciosamente o valor).
  itens_PC!J2:J5001 : DELTA_POTENCIAL — mesma regra (G<>"Sim" com efeito
                      aplicavel -> ROUND(F-D,2)).
  itens_PC!K2:K5001 : CHECK — G vazio deixa de ser erro estrutural
                      ("PC_PAGO_A_CONTRATADA invalido" somente para texto
                      diferente de Sim/Nao/vazio). Duplicidade, data,
                      ciclo, valor e efeito continuam identicos.
  MEMORIA_RESULTADOS!C10:C14 : retroativo reconhecido por ciclo — a
                      presenca da base passa a ser "ha PC no ciclo"
                      (D>0), nao "ha PC PAGO no ciclo"; o valor segue
                      SUMIFS(H; G="Sim"). Com PCs e nenhum pago, o
                      reconhecido e 0 CALCULADO (nao "sem base"), o
                      potencial segue em D45, e F16/E26/RESULTADOS nao
                      rebaixam o selo do VTA por causa de G.

  NAO altera: H (reconhecido), F/P (valor atualizado), T20:T25, B26,
  X/Y auxiliares, linha de corte, posicao_contratual, historico_VU,
  itens_RC, aditivos, metodo Financeiro.

Idempotente. Preserva estilos, validacoes, protecao e named ranges.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path

import pythoncom
import win32com.client

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16

CAP = 5001  # ultima linha operacional de itens_PC (Etapa 26G)

CABECALHO_ITENS_PC = [
    "NUMERO_PC", "DATA_PC", "CICLO_PC", "VALOR_PC", "FATOR_ACUMULADO",
    "VALOR_ATUALIZADO", "PC_PAGO_A_CONTRATADA",
    "RETROATIVO_RECONHECIDO_A_PAGAR", "VALOR_ATUALIZADO_EM_ANALISE",
    "DELTA_POTENCIAL", "CHECK_PC_FINANCEIRO", "EFEITO_FINANCEIRO_PC",
]

# Formulas em referencia relativa da linha 2; o Excel ajusta a linha ao
# aplicar na faixa inteira. ASCII puro e parenteses balanceados.
FORMULA_I2 = '=IF(G2="Sim",0,IF(OR(F2="",L2=""),"",F2))'
FORMULA_J2 = (
    '=IF(G2="Sim",0,IF(OR(F2="",D2="",L2=""),"",'
    'IF(L2="Sim",ROUND(F2-D2,2),0)))'
)
FORMULA_K2 = (
    '=IF(AND(A2="",B2="",D2="",G2=""),"",'
    f'IF(AND(A2<>"",COUNTIF($A$2:$A${CAP},A2)>1),"NUMERO_PC duplicado",'
    'IF(B2="","DATA_PC vazia",'
    'IF(NOT(ISNUMBER(B2)),"DATA_PC invalida",'
    'IF(OR(C2="",C2="Fora dos ciclos"),"CICLO_PC nao identificado",'
    'IF(OR(D2="",D2=0),"VALOR_PC vazio ou zero",'
    'IF(AND(G2<>"Sim",G2<>"Nao",G2<>""),"PC_PAGO_A_CONTRATADA invalido",'
    'IF(AND(C2<>"C0",IFERROR(INDEX(parametros!$A$2:$A$6,'
    'MATCH(C2,parametros!$B$2:$B$6,0)),"")="Sim",'
    'IFERROR(INDEX(parametros!$H$2:$H$6,'
    'MATCH(C2,parametros!$B$2:$B$6,0)),"")=""),'
    '"INICIO_EFEITO ausente: PC "&A2&" - "&C2,'
    'IF(L2="","EFEITO_PC nao calculado: PC "&A2&" - "&C2,"OK")))))))))'
)


def _formula_memoria_c(ciclo: str) -> str:
    """C10:C14 — presenca por PC no ciclo; valor por PCs pagos (G="Sim")."""
    return (
        f'=IF(COUNTIFS(itens_PC!$C$2:$C${CAP},"{ciclo}",'
        f'itens_PC!$D$2:$D${CAP},">0")=0,"",'
        f'ROUND(SUMIFS(itens_PC!$H$2:$H${CAP},'
        f'itens_PC!$C$2:$C${CAP},"{ciclo}",'
        f'itens_PC!$G$2:$G${CAP},"Sim"),2))'
    )


def _validar_layout(wb) -> None:
    ipc = wb.Worksheets("itens_PC")
    encontrados = [ipc.Cells(1, c).Value for c in range(1, 13)]
    if encontrados != CABECALHO_ITENS_PC:
        raise ValueError(f"Layout A:L inesperado em itens_PC: {encontrados!r}")
    mem = wb.Worksheets("MEMORIA_RESULTADOS")
    c10 = str(mem.Range("C10").Formula or "")
    if "itens_PC!$H$2:$H$" not in c10 or "SUMIFS" not in c10:
        raise ValueError(f"MEMORIA_RESULTADOS!C10 inesperada: {c10!r}")
    t25 = str(mem.Range("T25").Formula or "")
    if "$T$21+$T$22+$T$23" not in t25:
        raise ValueError(f"Cadeia VTA-PC (T25) inesperada: {t25!r}")


def _aplicar_itens_pc(wb) -> None:
    ipc = wb.Worksheets("itens_PC")
    protegida = bool(ipc.ProtectContents)
    if protegida:
        ipc.Unprotect()
    ipc.Range(f"I2:I{CAP}").Formula = FORMULA_I2
    ipc.Range(f"J2:J{CAP}").Formula = FORMULA_J2
    ipc.Range(f"K2:K{CAP}").Formula = FORMULA_K2
    if protegida:
        ipc.Protect(DrawingObjects=True, Contents=True, Scenarios=True)


def _aplicar_memoria(wb) -> None:
    mem = wb.Worksheets("MEMORIA_RESULTADOS")
    for i, ciclo in enumerate(["C0", "C1", "C2", "C3", "C4"]):
        mem.Range(f"C{10 + i}").Formula = _formula_memoria_c(ciclo)


def _verificar_aplicado(wb) -> None:
    ipc = wb.Worksheets("itens_PC")
    mem = wb.Worksheets("MEMORIA_RESULTADOS")
    k_2 = str(ipc.Range("K2").Formula or "")
    k_fim = str(ipc.Range(f"K{CAP}").Formula or "")
    if 'G2<>""' not in k_2:
        raise RuntimeError("K2 nao recebeu a regra de G vazio.")
    if f'G{CAP}<>""' not in k_fim:
        raise RuntimeError(f"K{CAP} nao recebeu a regra de G vazio.")
    if not str(ipc.Range("I2").Formula or "").startswith('=IF(G2="Sim",0,'):
        raise RuntimeError("I2 nao recebeu a regra de potencial para vazio.")
    if not str(ipc.Range(f"J{CAP}").Formula or "").startswith(
        f'=IF(G{CAP}="Sim",0,'
    ):
        raise RuntimeError(f"J{CAP} nao recebeu a regra de potencial.")
    c14 = str(mem.Range("C14").Formula or "")
    if '"C4"' not in c14 or "itens_PC!$G$2" not in c14:
        raise RuntimeError("MEMORIA_RESULTADOS!C14 nao recebeu a nova base.")
    t25 = str(mem.Range("T25").Formula or "")
    if "$T$21+$T$22+$T$23" not in t25:
        raise RuntimeError("Cadeia VTA-PC (T25) foi alterada indevidamente.")


def _verificar_sem_erros(wb) -> None:
    import pywintypes
    problemas = []
    for ws in wb.Worksheets:
        for tipo in (XL_CELLTYPE_FORMULAS, XL_CELLTYPE_CONSTANTS):
            try:
                cel = ws.UsedRange.SpecialCells(tipo, XL_ERRORS)
                problemas.append(f"{ws.Name}!{cel.Address}")
            except pywintypes.com_error:
                continue
    if problemas:
        raise RuntimeError(f"Erros de formula apos recalculo: {problemas}")


def _validar_zip(caminho: Path) -> None:
    with zipfile.ZipFile(caminho) as zf:
        defeito = zf.testzip()
        if defeito is not None:
            raise RuntimeError(f"ZIP do XLSX corrompido em {defeito!r}")
        nomes = zf.namelist()
    if "xl/workbook.xml" not in nomes:
        raise RuntimeError("XLSX sem xl/workbook.xml; pacote invalido.")


def aplicar(origem: Path, destino: Path) -> None:
    origem, destino = Path(origem), Path(destino)
    if not origem.is_file():
        raise FileNotFoundError(origem)
    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_27b_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = True
    wb = None
    salvo = False
    try:
        wb = excel.Workbooks.Open(str(tmp_xlsx), UpdateLinks=0)
        excel.ScreenUpdating = False
        excel.Calculation = XL_CALC_MANUAL
        aba_ativa = wb.ActiveSheet.Name
        _validar_layout(wb)
        _aplicar_itens_pc(wb)
        _aplicar_memoria(wb)
        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_aplicado(wb)
        _verificar_sem_erros(wb)
        if aba_ativa in [ws.Name for ws in wb.Worksheets]:
            wb.Worksheets(aba_ativa).Activate()
        wb.Save()
        salvo = True
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
    _validar_zip(tmp_xlsx)
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(tmp_xlsx, destino)
    shutil.rmtree(tmp_dir, ignore_errors=True)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("origem", type=Path)
    parser.add_argument("destino", type=Path)
    args = parser.parse_args()
    aplicar(args.origem, args.destino)


if __name__ == "__main__":
    main()
