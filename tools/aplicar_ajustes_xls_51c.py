# -*- coding: utf-8 -*-
"""Etapa 51C: ajustes visuais/estruturais de abas da Coleta, via Excel COM.

ESCOPO: APRESENTACAO. Nenhuma formula, valor, endereco, nome definido ou
regra de negocio muda — a guarda compara as formulas de TODAS as abas antes
e depois e aborta em qualquer divergencia. Mudancas:

  1. parametros!I1:I6 (DATA_ABERTURA_FISICA_EXATA): recebe por COPIA DE
     FORMATOS o estilo do bloco (H1:H6) — cabecalho azul institucional,
     corpo com bordas/fills do quadro. Endereco, largura e conteudo
     intactos (a coluna e escrita pelo runtime da Etapa 48).
  2. financeiro linha 74 (TOTAL): linha inteira em NEGRITO com fill
     institucional. A ancora NAO se move: as linhas 62:73 sao capacidade
     estrutural viva (grade de 72 competencias referenciada por
     MEMORIA/RESULTADOS em $2:$73 e reescrita pelo runtime) — mover ou
     espelhar o TOTAL na linha 64 quebraria 6 contratos de teste e 6
     modulos. A aproximacao visual e feita pelo runtime (_gerador_
     masterfile oculta as linhas de capacidade nao usadas).
  3. aditivos linha 1: altura 95 (o texto de orientacao gravado pelo
     runtime tem 173 caracteres e nao cabia nos 58 atuais). Wrap ja era
     True; conteudo intacto.
  4. cobertura_temporal: padronizacao cirurgica — fontes residuais 11 ->
     Calibri 10, formatos-residuo ('0' em celulas de texto, 'mm-dd-yy'
     nos rotulos da legenda) -> General, alinhamento da coluna B,
     bordas/estilo do bloco de legenda (A25:C29) e das notas C8/C14/C17.
     Os fills FUNCIONAIS sao preservados e verificados: amarelo GCC
     (B13/B15), laranja projecao (B16/B23) e os swatches A26:A29.

Aplicado SOBRE o template pos-51A (cadeia oficial: aplicador 50 ->
aplicador 51A -> aplicador 51C). Excel COM preserva as regras de CF x14
que o openpyxl removeria ao salvar.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

RPC_E_CALL_REJECTED = -2147418111

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16
XL_PASTE_FORMATS = -4122
XL_THIN = 2
XL_CONTINUOUS = 1
XL_CENTER = -4108
XL_EDGES = (7, 8, 9, 10)  # xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight

ABA_MEMORIA = "MEMORIA_RESULTADOS"

ABAS_ESCOPO = ("parametros", "financeiro", "aditivos", "cobertura_temporal")

# Fills FUNCIONAIS da cobertura_temporal (travados por teste) — verificados
# apos a aplicacao para garantir que a padronizacao nao apagou semantica.
FILLS_FUNCIONAIS_COBERTURA = {
    "B13": "FEF9C3", "B15": "FEF9C3",   # entrada GCC (amarelo)
    "B16": "FCE4D6", "B23": "FCE4D6",   # projecao (laranja)
    "A26": "FEF9C3", "A27": "FEF9C3",   # swatches da legenda
    "A28": "EBF3FB", "A29": "FCE4D6",
}


def _cor(hex_rgb: str) -> int:
    texto = hex_rgb.strip().lstrip("#")
    r, g, b = (int(texto[i : i + 2], 16) for i in (0, 2, 4))
    return r + (g << 8) + (b << 16)


AZUL_INSTITUCIONAL = _cor("1F4E79")
FILL_TOTAL = _cor("D6E4F0")
BRANCO = _cor("FFFFFF")


# --------------------------------------------------------------------------- #
# Guardas
# --------------------------------------------------------------------------- #

def _snapshot_aba(ws) -> dict[str, object]:
    usado = ws.UsedRange
    linhas, colunas = usado.Rows.Count, usado.Columns.Count
    dados = usado.Formula
    primeira_linha, primeira_coluna = usado.Row, usado.Column
    if linhas == 1 and colunas == 1:
        return {f"{primeira_linha}:{primeira_coluna}": dados}
    achatado: dict[str, object] = {}
    for i in range(linhas):
        linha = dados[i] if linhas > 1 else dados
        for j in range(colunas):
            valor = linha[j] if colunas > 1 else linha
            if valor not in (None, ""):
                achatado[f"{primeira_linha + i}:{primeira_coluna + j}"] = valor
    return achatado


def _snapshot_workbook(wb) -> dict[str, dict]:
    return {ws.Name: _snapshot_aba(ws) for ws in wb.Worksheets}


def _verificar_formulas_intactas(wb, antes: dict[str, dict]) -> None:
    depois = _snapshot_workbook(wb)
    for aba in antes:
        if antes[aba] != depois.get(aba):
            divergentes = sorted(
                k for k in set(antes[aba]) | set(depois.get(aba, {}))
                if antes[aba].get(k) != depois.get(aba, {}).get(k)
            )
            raise RuntimeError(
                f"Formulas/valores da aba {aba} mudaram ({len(divergentes)}): "
                + ", ".join(divergentes[:10])
            )


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


def _verificar_fills_funcionais(res) -> None:
    for endereco, hex_rgb in FILLS_FUNCIONAIS_COBERTURA.items():
        atual = int(res.Range(endereco).Interior.Color)
        if atual != _cor(hex_rgb):
            raise RuntimeError(
                f"Fill funcional cobertura_temporal!{endereco} foi alterado "
                f"(esperado {hex_rgb}, encontrado {atual:06X})."
            )


def _validar_origem(wb) -> None:
    abas = [ws.Name for ws in wb.Worksheets]
    for aba in ABAS_ESCOPO + (ABA_MEMORIA, "RESULTADOS"):
        if aba not in abas:
            raise ValueError(f"Template sem a aba {aba}.")
    fin = wb.Worksheets("financeiro")
    if str(fin.Range("B74").Value or "") != "TOTAL":
        raise ValueError("financeiro!B74 nao e o rotulo TOTAL; leiaute inesperado.")
    if not str(fin.Range("C74").Formula or "").startswith("=SUM"):
        raise ValueError("financeiro!C74 nao e a soma canonica; abortando.")
    res = wb.Worksheets("RESULTADOS")
    if str(res.Range("E5").Formula or "") != "=$D$22":
        raise ValueError("RESULTADOS fora do leiaute 51A; abortando.")


# --------------------------------------------------------------------------- #
# Ajustes (todos de apresentacao)
# --------------------------------------------------------------------------- #

def _ajuste_parametros_coluna_i(wb) -> None:
    ws = wb.Worksheets("parametros")
    ws.Range("H1:H6").Copy()
    ws.Range("I1").PasteSpecial(Paste=XL_PASTE_FORMATS)
    wb.Application.CutCopyMode = False


def _ajuste_financeiro_total(wb) -> None:
    ws = wb.Worksheets("financeiro")
    faixa = ws.Range("A74:G74")
    faixa.Font.Bold = True
    faixa.Interior.Color = FILL_TOTAL
    for borda in XL_EDGES:
        faixa.Borders(borda).LineStyle = XL_CONTINUOUS
        faixa.Borders(borda).Weight = XL_THIN


def _ajuste_aditivos_linha1(wb) -> None:
    ws = wb.Worksheets("aditivos")
    # O runtime substitui A1 por um texto de 173 caracteres; 58 nao comporta.
    ws.Rows(1).RowHeight = 95
    ws.Range("A1").WrapText = True


def _ajuste_cobertura_temporal(wb) -> None:
    ws = wb.Worksheets("cobertura_temporal")
    # 1) fontes residuais Calibri 11 -> padrao da aba (Calibri 10).
    for endereco in ("B4", "B13", "B15", "C8", "C14", "C17",
                     "A26", "A27", "B26", "B27", "B28", "B29"):
        cel = ws.Range(endereco)
        cel.Font.Name = "Calibri"
        cel.Font.Size = 10
    # 2) alinhamento da coluna B igual aos demais valores automaticos.
    for endereco in ("B4", "B13", "B15"):
        cel = ws.Range(endereco)
        cel.HorizontalAlignment = XL_CENTER
        cel.VerticalAlignment = XL_CENTER
    # 3) formatos-residuo em celulas de TEXTO -> Geral (nao afeta exibicao).
    for endereco in ("B2", "B5", "B19", "B20", "B22", "B23", "A26", "A27",
                     "A28", "A29"):
        cel = ws.Range(endereco)
        try:
            cel.NumberFormat = "General"
        except Exception:
            cel.NumberFormatLocal = "Geral"
    # 4) bloco de legenda: banner A25 completo + molduras das descricoes.
    banner = ws.Range("A25:C25")
    banner.Interior.Color = AZUL_INSTITUCIONAL
    banner.Font.Name = "Calibri"
    banner.Font.Size = 10
    banner.Font.Bold = True
    banner.Font.Color = BRANCO
    for faixa in ("A25:C25", "A26:C29"):
        rng = ws.Range(faixa)
        for borda in XL_EDGES:
            rng.Borders(borda).LineStyle = XL_CONTINUOUS
            rng.Borders(borda).Weight = XL_THIN
    # 5) notas da coluna C alinhadas ao padrao (fonte 10 ja aplicada acima).
    for endereco in ("C8", "C14", "C17"):
        ws.Range(endereco).VerticalAlignment = XL_CENTER


def _aplicar_51c(wb) -> None:
    _ajuste_parametros_coluna_i(wb)
    _ajuste_financeiro_total(wb)
    _ajuste_aditivos_linha1(wb)
    _ajuste_cobertura_temporal(wb)


# --------------------------------------------------------------------------- #
# Orquestracao
# --------------------------------------------------------------------------- #

def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_51c_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.Interactive = False
        excel.EnableEvents = False
    except Exception:
        pass
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

        memoria = wb.Worksheets(ABA_MEMORIA)
        formulas_antes = _snapshot_workbook(wb)
        w48_antes = memoria.Range("W48").Formula
        nomes_antes = {str(n.Name) for n in wb.Names}

        _aplicar_51c(wb)

        _verificar_formulas_intactas(wb, formulas_antes)
        if memoria.Range("W48").Formula != w48_antes:
            raise RuntimeError("MEMORIA_RESULTADOS!W48 foi alterada; abortando.")
        if {str(n.Name) for n in wb.Names} != nomes_antes:
            raise RuntimeError("Nomes definidos mudaram; abortando.")
        _verificar_fills_funcionais(wb.Worksheets("cobertura_temporal"))

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)

        if aba_ativa in [ws.Name for ws in wb.Worksheets] and aba_ativa != ABA_MEMORIA:
            wb.Worksheets(aba_ativa).Activate()
        else:
            wb.Worksheets("CONTROLE").Activate()
        wb.Save()
        salvo = True
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


def aplicar_com_retentativas(origem: Path, destino: Path,
                             tentativas: int = 6) -> None:
    """Reexecuta do zero quando o Excel rejeita chamadas COM (transitorio)."""
    import time
    import pywintypes

    for tentativa in range(1, tentativas + 1):
        try:
            aplicar(origem, destino)
            return
        except pywintypes.com_error as exc:
            transitorio = RPC_E_CALL_REJECTED in (
                getattr(exc, "hresult", None), *(exc.args or ())
            )
            if not transitorio or tentativa == tentativas:
                raise
            print(f"Excel ocupado (tentativa {tentativa}/{tentativas}); "
                  "aguardando 20s e reexecutando do zero...")
            time.sleep(20)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("origem", type=Path)
    parser.add_argument("destino", type=Path)
    args = parser.parse_args()
    aplicar_com_retentativas(args.origem, args.destino)
    print("Ajustes XLS (51C) aplicados:", args.destino)


if __name__ == "__main__":
    main()
