# -*- coding: utf-8 -*-
"""RESULTADOS-ROLLBACK-1: retorno cirurgico a apresentacao anterior da aba.

A camada humana introduzida pela frente RESULTADOS-UX2 (PRs #136/#137) vive
integralmente em RESULTADOS!90:166. Este aplicador a remove e devolve a aba a
apresentacao do checkpoint f8296f7, SEM tocar no motor tecnico das linhas 1:87
e sem substituir o template inteiro — as melhorias posteriores dos PRs #134,
#138, #139 e #140 permanecem intactas em todas as abas.

O QUE E RESTAURADO DO DOADOR f8296f7
  - linhas 90:166 removidas (valores, formulas, merges, CF, alturas, estilos);
  - visibilidade: ocultas apenas 10:13, 31, 40 e 51; colunas ocultas J:N;
  - print area A1:H50, paisagem, fitToPage com escala 68, margens do doador;
  - quebra de pagina manual da linha 116 removida;
  - sheetView sem tabSelected/topLeftCell, celula ativa D5, aba ativa CONTROLE;
  - RESULTADOS!B55:B59 volta ao formato "Geral".

O QUE E DELIBERADAMENTE PRESERVADO (nao volta ao doador)
  - RESULTADOS!H5/H8/C12: fator historico canonico do PR #139;
  - C43:G50 (ajustes manuais) e suas duas validacoes de dados;
  - MEMORIA_RESULTADOS!S38/T38 e o name RETROATIVO_POTENCIAL_PC;
  - os 14 defined names do PR #135 e todas as demais abas.

REGRA ZERO CORRUPCAO XLSX: aplicacao por Excel COM — openpyxl remove a
formatacao condicional x14 que a aba usa nas linhas 1:87.

Uso:  python tools/aplicar_rollback_resultados_apresentacao.py [caminho.xlsx]
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

PRIMEIRA_LINHA_UX2 = 90
ULTIMA_LINHA_UX2 = 166

# Visibilidade da apresentacao anterior (f8296f7).
LINHAS_OCULTAS = (10, 11, 12, 13, 31, 40, 51)
ULTIMA_LINHA_TECNICA = 89          # 88 e 89 existem e sao visiveis no doador
COLUNA_REEXIBIDA = "I"             # a UX2 ocultou I por causa dos helpers I100/I125

# Formato monetario aplicado pela UX2; o doador tem "Geral".
CELULAS_FORMATO_GERAL = "B55:B59"

# PageSetup do doador (polegadas).
PRINT_AREA = "$A$1:$H$50"
ESCALA = 68
MARGEM_LATERAL = 0.511811024
MARGEM_VERTICAL = 0.78740157499999996
MARGEM_CABECALHO = 0.31496062000000002

ABA_ATIVA_FINAL = "CONTROLE"       # workbookView activeTab="1" no doador
CELULA_ATIVA = "D5"

XL_LANDSCAPE = 2
XL_PAPER_A4 = 9
XL_OPENXML = 51
XL_SHIFT_UP = -4162


def _fechar(wb, tentativas: int = 10) -> None:
    """Excel recusa Close logo apos um Save pesado (RPC_E_CALL_REJECTED)."""
    for _ in range(tentativas):
        try:
            wb.Close(SaveChanges=False)
            return
        except Exception:
            time.sleep(1.0)


def aplicar(caminho: Path) -> None:
    import win32com.client as com

    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        ws = wb.Worksheets("RESULTADOS")

        faixa = f"A{PRIMEIRA_LINHA_UX2}:N{ULTIMA_LINHA_UX2}"
        linhas = f"{PRIMEIRA_LINHA_UX2}:{ULTIMA_LINHA_UX2}"

        # 1. camada UX2: formatacao condicional, merges e, por fim, as linhas.
        ws.Range(faixa).FormatConditions.Delete()
        ws.Range(faixa).UnMerge()
        ws.Rows(linhas).Delete(XL_SHIFT_UP)

        # 2. visibilidade da apresentacao anterior.
        ws.Rows(f"1:{ULTIMA_LINHA_TECNICA}").Hidden = False
        for linha in LINHAS_OCULTAS:
            ws.Rows(f"{linha}:{linha}").Hidden = True
        ws.Columns(COLUNA_REEXIBIDA).Hidden = False

        # 3. formato monetario da UX2 revertido (pt-BR rejeita "General").
        try:
            ws.Range(CELULAS_FORMATO_GERAL).NumberFormatLocal = "Geral"
        except Exception:
            ws.Range(CELULAS_FORMATO_GERAL).NumberFormat = "General"

        # 4. quebra manual da linha 116 e print setup do doador.
        ws.ResetAllPageBreaks()
        setup = ws.PageSetup
        setup.PrintArea = PRINT_AREA
        setup.Orientation = XL_LANDSCAPE
        setup.PaperSize = XL_PAPER_A4
        setup.Zoom = ESCALA          # grava scale="68"
        setup.Zoom = False           # liga pageSetUpPr fitToPage="1"
        setup.FitToPagesWide = 1
        setup.FitToPagesTall = False
        setup.LeftMargin = excel.InchesToPoints(MARGEM_LATERAL)
        setup.RightMargin = excel.InchesToPoints(MARGEM_LATERAL)
        setup.TopMargin = excel.InchesToPoints(MARGEM_VERTICAL)
        setup.BottomMargin = excel.InchesToPoints(MARGEM_VERTICAL)
        setup.HeaderMargin = excel.InchesToPoints(MARGEM_CABECALHO)
        setup.FooterMargin = excel.InchesToPoints(MARGEM_CABECALHO)

        # 5. sheetView: sem topLeftCell/tabSelected, celula ativa D5.
        ws.Activate()
        janela = excel.ActiveWindow
        janela.ScrollRow = 1
        janela.ScrollColumn = 1
        ws.Range(CELULA_ATIVA).Select()
        wb.Worksheets(ABA_ATIVA_FINAL).Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1

        wb.Save()
        _fechar(wb)
        wb = None
    finally:
        if wb is not None:
            _fechar(wb)
        excel.Quit()


def main() -> int:
    caminho = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else TEMPLATE
    if not caminho.exists():
        print(f"ERRO: arquivo nao encontrado: {caminho}")
        return 1
    aplicar(caminho)
    print(f"OK: apresentacao anterior da aba RESULTADOS restaurada em {caminho}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
