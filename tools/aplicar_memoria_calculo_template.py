"""Aplica, de forma reproduzivel, a Etapa 4 ao template oficial da Coleta.

Cria SOMENTE os cabecalhos do bloco MEMORIA DE CALCULO (parametros!J1:R1),
larguras e formatos de coluna. Gravador unico: Microsoft Excel real via COM
(DisplayAlerts=True). Trabalha sobre copia temporaria e somente promove o
resultado ao destino depois de o proprio Excel salvar e fechar sem erros.
Preserva integralmente A:H, a coluna I (separacao visual, vazia), a aba
financeiro e a aba itens_PC. Os VALORES do bloco sao gravados pelo fluxo
runtime homologado (openpyxl) em _memoria_calculo.escrever_memoria_calculo.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

XL_PASTE_FORMATS = -4122
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16

CABECALHO_PARAMETROS = [
    "COMPUTAR_NESTA_APURACAO", "CICLO", "DATA_INICIO", "DATA_FIM",
    "PERCENTUAL_DO_CICLO", "FATOR_ACUMULADO", "SITUACAO",
    "INICIO_EFEITO_FINANCEIRO",
]
CABECALHOS_MEMORIA = [
    "CICLO", "TIPO_REGISTRO", "ORDEM", "COMPETENCIA", "VALOR_INDICE",
    "FATOR_MENSAL", "FATOR_ACUMULADO", "VARIACAO_FINAL", "METODO_FONTE",
]
LARGURAS = {
    "J": 9, "K": 15, "L": 8, "M": 14, "N": 14,
    "O": 14, "P": 16, "Q": 15, "R": 60,
}
LINHA_FIM = 80


def _validar_layout(ws_par, excel) -> None:
    encontrados = [ws_par.Cells(1, c).Value for c in range(1, 9)]
    if encontrados != CABECALHO_PARAMETROS:
        raise ValueError(f"Layout A:H inesperado em parametros: {encontrados!r}")
    coluna_i = ws_par.Range(f"I1:I{LINHA_FIM}")
    if excel.WorksheetFunction.CountA(coluna_i) != 0:
        raise ValueError("Coluna I de parametros nao esta vazia; separacao visual comprometida.")
    bloco = ws_par.Range(f"J1:R{LINHA_FIM}")
    if excel.WorksheetFunction.CountA(bloco) != 0:
        raise ValueError("Regiao J1:R80 de parametros ja possui conteudo inesperado.")


def _aplicar_cabecalhos(ws_par, excel) -> None:
    # Mesmo estilo visual do cabecalho oficial A1:H1.
    ws_par.Range("H1").Copy()
    ws_par.Range("J1:R1").PasteSpecial(XL_PASTE_FORMATS)
    excel.CutCopyMode = False
    ws_par.Range("J1:R1").Value = [CABECALHOS_MEMORIA]
    for coluna, largura in LARGURAS.items():
        ws_par.Columns(coluna).ColumnWidth = largura
    # Competencia exibida como mm/aaaa. Em Excel pt-BR, NumberFormat en-US
    # grava "y" literal (defeito comprovado na Etapa 3); usar o caminho local
    # com ";@" e validar por renderizacao real.
    try:
        ws_par.Range(f"M2:M{LINHA_FIM}").NumberFormatLocal = "mm/aaaa;@"
    except Exception:
        ws_par.Range(f"M2:M{LINHA_FIM}").NumberFormat = "mm/yyyy;@"
    ws_par.Range("M2").Value2 = 45400  # 18/04/2024
    renderizado = ws_par.Range("M2").Text
    ws_par.Range("M2").ClearContents()
    if renderizado != "04/2024":
        raise RuntimeError(
            f"Formato de competencia nao aplicado em parametros!M2:M{LINHA_FIM}: "
            f"data serial 45400 renderizou {renderizado!r}"
        )
    # Formatos numericos padrao das colunas de valores (o gerador reforca
    # por celula). Em pt-BR, NumberFormat en-US com "." decimal e armazenado
    # corrompido ("00,000%"); usar o caminho LOCAL com virgula decimal, que
    # persiste o codigo OOXML deterministico com ponto.
    ws_par.Range(f"N2:N{LINHA_FIM}").NumberFormatLocal = "0,0000%"
    ws_par.Range(f"O2:P{LINHA_FIM}").NumberFormatLocal = "0,000000"
    ws_par.Range(f"Q2:Q{LINHA_FIM}").NumberFormatLocal = "0,0000%"
    ws_par.Range("N2").Value2 = 0.0045
    if ws_par.Range("N2").Text.replace("\xa0", " ") != "0,4500%":
        raise RuntimeError(
            f"Formato percentual nao aplicado em parametros!N: 0.0045 "
            f"renderizou {ws_par.Range('N2').Text!r}"
        )
    ws_par.Range("N2").ClearContents()


def _verificar_sem_erros(wb) -> None:
    import pywintypes

    problemas = []
    for ws in wb.Worksheets:
        for tipo in (XL_CELLTYPE_FORMULAS, XL_CELLTYPE_CONSTANTS):
            try:
                celulas = ws.UsedRange.SpecialCells(tipo, XL_ERRORS)
                problemas.append(f"{ws.Name}!{celulas.Address}")
            except pywintypes.com_error:
                continue  # nenhuma celula de erro deste tipo
    if problemas:
        raise RuntimeError(f"Erros de formula apos recalculo: {problemas}")


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem)
    destino = Path(destino)
    if not origem.is_file():
        raise FileNotFoundError(origem)

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_etapa4_"))
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
        ws_par = wb.Worksheets("parametros")

        _validar_layout(ws_par, excel)
        _aplicar_cabecalhos(ws_par, excel)

        ws_par.Activate()
        ws_par.Range("A1").Select()
        wb.Worksheets(1).Activate()

        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)

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
