# -*- coding: utf-8 -*-
"""Validacao visual da aba RESULTADOS executiva no Microsoft Excel REAL.

Para cada XLSX informado:
  1. abre no Excel (sem alertas) e verifica se houve REPARO: o Excel grava um
     log errorNNNNNN_NN.xml em %TEMP% sempre que remove/repara conteudo;
  2. CalculateFullRebuild e, com a janela em Zoom 100%, le a propriedade
     `.Text` de TODAS as celulas visiveis de RESULTADOS!B1:G<ultima>:
       - falha com '###' (valor que nao coube) ou erro (#REF!, #NAME?...);
       - falha com 'yyyy'/'aaaa' literal ou percentual sem 2 casas;
  3. ENCAIXE REAL do texto: copia cada celula (mesma fonte/tamanho/negrito/
     recuo) para uma aba-rascunho NAO mesclada com a largura da area visivel
     (soma das colunas da mesclagem) e mede com AutoFit do proprio Excel:
       - sem quebra: largura necessaria <= largura disponivel;
       - com quebra: altura necessaria <= altura da linha (ou da mesclagem);
  4. --estresse: substitui os valores monetarios por 999.999.999,99 (e
     negativos) numa COPIA e repete os passos 2 e 3;
  5. exporta PNG (Range.CopyPicture) e PDF da aba para inspecao.
O workbook original nunca e salvo.

Uso: python tools/validar_visual_resultados_v3.py ARQ.xlsx [--saida DIR] [--estresse]
"""
from __future__ import annotations

import argparse
import gc
import os
import re
import shutil
import sys
import tempfile
import time
from pathlib import Path

ABA = "RESULTADOS"
COLUNAS = "BCDEFG"
ULTIMA = 71
RE_ERRO = re.compile(r"#(REF!|NAME\?|VALUE!|DIV/0!|N/A|NUM!|NULL!)")
RE_PCT = re.compile(r"-?\d{1,3}(\.\d{3})*,\d{2}%$")
XL_SCREEN, XL_BITMAP = 1, 2
XL_TYPE_PDF = 0

VALORES_ESTRESSE = (999_999_999.99, -999_999_999.99, 132_581_980.10,
                    13_480_627.05, 8_713_820.26, 605.17, 0.0)


def _logs_reparo() -> set[str]:
    pasta = Path(tempfile.gettempdir())
    return {p.name for p in pasta.glob("error*.xml")}


def _abrir_excel():
    import pythoncom
    import win32com.client as com

    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel


def _fechar_excel(excel) -> None:
    import pythoncom

    for _ in range(10):
        try:
            excel.Quit()
            break
        except Exception:
            time.sleep(1.0)
    gc.collect()
    pythoncom.CoUninitialize()


def _area_visivel(cel):
    area = cel.MergeArea
    largura = sum(area.Columns(i).ColumnWidth for i in range(1, area.Columns.Count + 1))
    altura = sum(area.Rows(i).RowHeight for i in range(1, area.Rows.Count + 1))
    return area, largura, altura


def _ancora(cel) -> bool:
    return not cel.MergeCells or cel.MergeArea.Cells(1, 1).Address == cel.Address


def _checar_textos(ws, problemas: list[str], rotulo: str) -> None:
    for linha in range(1, ULTIMA + 1):
        for col in COLUNAS:
            cel = ws.Range(f"{col}{linha}")
            if not _ancora(cel):
                continue
            texto = str(cel.Text or "")
            if not texto:
                continue
            if "###" in texto or texto.strip("#") == "":
                problemas.append(f"[{rotulo}] {col}{linha}: '{texto}' (nao coube)")
            if RE_ERRO.search(texto):
                problemas.append(f"[{rotulo}] {col}{linha}: erro {texto}")
            if "yyyy" in texto.lower() or "aaaa" in texto.lower():
                problemas.append(f"[{rotulo}] {col}{linha}: data mal formatada '{texto}'")
            if texto.endswith("%") and not RE_PCT.match(texto):
                problemas.append(f"[{rotulo}] {col}{linha}: percentual fora de xx,xx% '{texto}'")


def _checar_encaixe(wb, ws, problemas: list[str], rotulo: str) -> list[str]:
    """Mede, com AutoFit real, se cada texto cabe na area em que e exibido."""
    medidas: list[str] = []
    rascunho = wb.Worksheets.Add()
    try:
        alvo = rascunho.Range("A1")
        for linha in range(1, ULTIMA + 1):
            for col in COLUNAS:
                cel = ws.Range(f"{col}{linha}")
                if not _ancora(cel):
                    continue
                texto = str(cel.Text or "")
                if not texto:
                    continue
                _area, largura, altura = _area_visivel(cel)
                alvo.Clear()
                alvo.NumberFormat = "@"
                alvo.Value = texto
                alvo.Font.Name = cel.Font.Name
                alvo.Font.Size = cel.Font.Size
                alvo.Font.Bold = cel.Font.Bold
                alvo.Font.Italic = cel.Font.Italic
                alvo.IndentLevel = cel.IndentLevel
                alvo.WrapText = bool(cel.WrapText)
                if cel.WrapText:
                    rascunho.Columns("A:A").ColumnWidth = largura
                    rascunho.Rows("1:1").AutoFit()
                    necessaria = rascunho.Rows(1).RowHeight
                    if necessaria > altura + 0.5:
                        problemas.append(
                            f"[{rotulo}] {col}{linha}: texto com quebra precisa de "
                            f"{necessaria:.1f}pt e a linha tem {altura:.1f}pt: '{texto[:70]}'"
                        )
                    medidas.append(f"{col}{linha} quebra {necessaria:.1f}/{altura:.1f}pt")
                else:
                    # Texto sem quebra, nao centralizado/direita, transborda para
                    # vizinhas REALMENTE vazias (sem formula), como o Excel faz.
                    if (not cel.MergeCells
                            and cel.HorizontalAlignment in (1, -4131)):
                        idx = COLUNAS.index(col) + 1
                        while idx < len(COLUNAS):
                            viz = ws.Range(f"{COLUNAS[idx]}{linha}")
                            if viz.Formula not in (None, "") or viz.MergeCells:
                                break
                            largura += viz.ColumnWidth
                            idx += 1
                    rascunho.Columns("A:A").AutoFit()
                    necessaria = rascunho.Columns(1).ColumnWidth
                    if necessaria > largura + 0.01:
                        problemas.append(
                            f"[{rotulo}] {col}{linha}: texto precisa de largura "
                            f"{necessaria:.2f} e a area tem {largura:.2f}: '{texto[:70]}'"
                        )
                    medidas.append(f"{col}{linha} largura {necessaria:.2f}/{largura:.2f}")
    finally:
        wb.Application.DisplayAlerts = False
        rascunho.Delete()
    return medidas


def _exportar(ws, saida: Path, base: str) -> None:
    saida.mkdir(parents=True, exist_ok=True)
    ws.Activate()
    rng = ws.Range(f"A1:H{ULTIMA}")
    rng.CopyPicture(XL_SCREEN, XL_BITMAP)
    grafico = ws.ChartObjects().Add(0, 0, rng.Width, rng.Height)
    grafico.Activate()
    grafico.Chart.Paste()
    grafico.Chart.Export(str((saida / f"{base}.png").resolve()))
    grafico.Delete()
    ws.ExportAsFixedFormat(XL_TYPE_PDF, str((saida / f"{base}.pdf").resolve()))


def _estressar(ws) -> None:
    """Valores monetarios extremos nas celulas monetarias (copia descartavel)."""
    alvos = [f"{c}9" for c in "BCDE"] + [f"C{n}" for n in range(14, 21)] + [
        f"{c}{n}" for c in "CDE" for n in range(43, 50)] + [
        f"{c}{n}" for c in "CDE" for n in range(54, 58)] + [
        f"C{n}" for n in range(64, 72)] + ["C30"]
    for i, endereco in enumerate(alvos):
        ws.Range(endereco).Value = VALORES_ESTRESSE[i % len(VALORES_ESTRESSE)]


def validar(caminho: Path, saida: Path, estresse: bool) -> list[str]:
    problemas: list[str] = []
    saida.mkdir(parents=True, exist_ok=True)
    excel = _abrir_excel()
    wb = None
    try:
        antes = _logs_reparo()
        wb = excel.Workbooks.Open(str(caminho.resolve()))
        novos = _logs_reparo() - antes
        if novos:
            problemas.append(f"REPARO do Excel ao abrir (logs: {sorted(novos)})")
        excel.CalculateFullRebuild()
        ws = wb.Worksheets(ABA)
        ws.Activate()
        excel.ActiveWindow.Zoom = 100
        _checar_textos(ws, problemas, "normal")
        medidas = _checar_encaixe(wb, ws, problemas, "normal")
        (saida / f"{caminho.stem}.medidas.txt").write_text("\n".join(medidas), encoding="utf-8")
        _exportar(ws, saida, caminho.stem)
        wb.Close(SaveChanges=False)
        wb = None
        if estresse:
            copia = Path(tempfile.gettempdir()) / f"estresse_{caminho.name}"
            shutil.copyfile(caminho, copia)
            wb = excel.Workbooks.Open(str(copia))
            ws = wb.Worksheets(ABA)
            _estressar(ws)
            excel.Calculate()
            ws.Activate()
            excel.ActiveWindow.Zoom = 100
            _checar_textos(ws, problemas, "estresse")
            _checar_encaixe(wb, ws, problemas, "estresse")
            _exportar(ws, saida, caminho.stem + "_estresse")
            wb.Close(SaveChanges=False)
            wb = None
            os.remove(copia)
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        _fechar_excel(excel)
    return problemas


def main() -> int:
    p = argparse.ArgumentParser()
    p.add_argument("arquivos", nargs="+", type=Path)
    p.add_argument("--saida", type=Path, default=Path("validacao_visual"))
    p.add_argument("--estresse", action="store_true")
    args = p.parse_args()
    total = 0
    for arquivo in args.arquivos:
        problemas = validar(arquivo, args.saida, args.estresse)
        print(f"{arquivo.name}: {'OK' if not problemas else f'{len(problemas)} PROBLEMA(S)'}")
        for linha in problemas:
            print("   ", linha)
        total += len(problemas)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
