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
XL_CALCULO_MANUAL, XL_CALCULO_AUTOMATICO = -4135, -4105
XL_CELULAS_FORMULA, XL_ERROS = -4123, 16
RE_DATA = re.compile(r"\d{2}/\d{2}/\d{4}$")
RPC_E_CALL_REJECTED = -2147418111

VALORES_ESTRESSE = (999_999_999.99, -999_999_999.99, 132_581_980.10,
                    13_480_627.05, 8_713_820.26, 605.17, 0.0)


def _logs_reparo() -> set[str]:
    pasta = Path(tempfile.gettempdir())
    return {p.name for p in pasta.glob("error*.xml")}


class _FiltroMensagens:
    """IMessageFilter: quando o Excel esta ocupado e REJEITA a chamada
    (RPC_E_CALL_REJECTED), o COM reenvia apos 250 ms em vez de falhar."""

    _com_interfaces_ = ["IMessageFilter"]
    _public_methods_ = ["HandleInComingCall", "RetryRejectedCall", "MessagePending"]

    def HandleInComingCall(self, tipo, tarefa, tempo, info):  # noqa: N802
        return 0                                             # SERVERCALL_ISHANDLED

    def RetryRejectedCall(self, tarefa, tempo, tipo):        # noqa: N802
        return 250 if tempo < 120_000 else -1                # desiste apos 2 min

    def MessagePending(self, tarefa, tempo, tipo):           # noqa: N802
        return 2                                             # PENDINGMSG_WAITDEFPROCESS


def _registrar_filtro() -> None:
    import pythoncom

    registrar = getattr(pythoncom, "CoRegisterMessageFilter", None)
    if registrar is None:          # pywin32 sem a API: fica a retentativa do main
        return
    from win32com.server.util import wrap

    registrar(wrap(_FiltroMensagens(), pythoncom.IID_IMessageFilter))


def _abrir_excel():
    import pythoncom
    import win32com.client as com

    pythoncom.CoInitialize()
    _registrar_filtro()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel


def _pid_excel(excel) -> int | None:
    try:
        import win32process

        return win32process.GetWindowThreadProcessId(excel.Hwnd)[1]
    except Exception:
        return None


def _fechar_excel(excel) -> None:
    """Quit + espera o EXCEL.EXE desta instancia terminar de fato, para que o
    cenario seguinte nunca rode com outra instancia ainda viva."""
    import pythoncom

    pid = _pid_excel(excel)
    for _ in range(10):
        try:
            excel.Quit()
            break
        except Exception:
            time.sleep(1.0)
    excel = None
    gc.collect()
    pythoncom.CoUninitialize()
    if pid is None:
        return
    import win32api
    import win32con
    import win32event

    try:
        handle = win32api.OpenProcess(
            win32con.SYNCHRONIZE | win32con.PROCESS_TERMINATE, False, pid
        )
    except Exception:
        return                                   # processo ja terminou
    try:
        if win32event.WaitForSingleObject(handle, 60_000) != win32event.WAIT_OBJECT_0:
            win32api.TerminateProcess(handle, 1)  # so ESTA instancia
            win32event.WaitForSingleObject(handle, 10_000)
    finally:
        win32api.CloseHandle(handle)


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
            formato = str(cel.NumberFormat or "")
            if "dd/mm" in formato and texto != "—" and not RE_DATA.match(texto):
                problemas.append(f"[{rotulo}] {col}{linha}: data fora de dd/mm/aaaa '{texto}'")
            numerico = isinstance(cel.Value, (int, float)) and not isinstance(cel.Value, bool)
            if numerico and "%" in formato and not RE_PCT.match(texto):
                problemas.append(f"[{rotulo}] {col}{linha}: percentual fora de xx,xx% '{texto}'")


def _checar_encaixe(wb, ws, problemas: list[str], rotulo: str) -> list[str]:
    """Mede, com AutoFit real, se cada texto cabe na area em que e exibido."""
    medidas: list[str] = []
    # Modo manual SO durante a medicao (evita recalculos ao criar/apagar a aba
    # rascunho). Fora daqui o calculo fica automatico: no Excel 365, valores
    # "desatualizados" em modo manual aparecem TACHADOS na tela e no PNG.
    excel = wb.Application
    excel.Calculation = XL_CALCULO_MANUAL
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
        excel.DisplayAlerts = False
        rascunho.Delete()
        ws.Activate()
        excel.Calculation = XL_CALCULO_AUTOMATICO
        excel.Calculate()
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


def _erros_no_workbook(wb, problemas: list[str]) -> None:
    """#REF!/#NAME?/#VALUE!/#DIV/0!... em formulas de QUALQUER aba."""
    for i in range(1, wb.Worksheets.Count + 1):
        ws = wb.Worksheets(i)
        try:
            erros = ws.UsedRange.SpecialCells(XL_CELULAS_FORMULA, XL_ERROS)
        except Exception:
            continue                      # SpecialCells sem resultado: nenhum erro
        enderecos = str(erros.Address).split(",")
        problemas.append(
            f"{erros.Count} celula(s) com erro em {ws.Name}: {', '.join(enderecos[:8])}"
        )


def _abrir_checando_reparo(excel, caminho: Path, problemas: list[str], etapa: str):
    antes = _logs_reparo()
    wb = excel.Workbooks.Open(str(caminho.resolve()))
    novos = _logs_reparo() - antes
    if novos:
        problemas.append(f"REPARO do Excel ({etapa}) (logs: {sorted(novos)})")
    return wb


def _ciclo_salvar_reabrir(excel, caminho: Path, problemas: list[str]) -> None:
    """Abrir -> salvar -> fechar -> reabrir numa COPIA, sem reparo."""
    copia = Path(tempfile.gettempdir()) / f"reabrir_{caminho.name}"
    shutil.copyfile(caminho, copia)
    try:
        wb = _abrir_checando_reparo(excel, copia, problemas, "abertura da copia")
        wb.Save()
        wb.Close(SaveChanges=False)
        wb = _abrir_checando_reparo(excel, copia, problemas, "reabertura apos salvar")
        if wb.Worksheets(ABA).Range("B2").Text != "RESULTADO DA APURAÇÃO":
            problemas.append("reabertura: titulo da RESULTADOS executiva ausente")
        wb.Close(SaveChanges=False)
    finally:
        try:
            os.remove(copia)
        except OSError:
            pass


def validar(caminho: Path, saida: Path, estresse: bool) -> list[str]:
    problemas: list[str] = []
    saida.mkdir(parents=True, exist_ok=True)
    excel = _abrir_excel()
    wb = None
    try:
        wb = _abrir_checando_reparo(excel, caminho, problemas, "abertura")
        excel.CalculateFullRebuild()
        _erros_no_workbook(wb, problemas)
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
        _ciclo_salvar_reabrir(excel, caminho, problemas)
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
    import pywintypes

    for arquivo in args.arquivos:
        for tentativa in range(1, 7):
            try:
                problemas = validar(arquivo, args.saida, args.estresse)
                break
            except pywintypes.com_error as erro:
                # Excel ocupado: nada foi salvo; refaz o arquivo do zero, com
                # instancia nova (a anterior ja foi encerrada no finally).
                if erro.args[0] != RPC_E_CALL_REJECTED or tentativa == 6:
                    raise
                print(f"  (Excel ocupado em {arquivo.name}; tentativa {tentativa + 1})")
                time.sleep(15.0)
        print(f"{arquivo.name}: {'OK' if not problemas else f'{len(problemas)} PROBLEMA(S)'}")
        for linha in problemas:
            print("   ", linha)
        total += len(problemas)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
