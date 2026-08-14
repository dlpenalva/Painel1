# -*- coding: utf-8 -*-
"""Hotfix RESULTADOS: retroativo visivel no card + VTA method-aware + contraste.

Aplicado via Excel COM sobre o template pos-51C (openpyxl destruiria a CF x14).

A. RETROATIVO (apresentacao):
   O valor do card ficava em E5 enquanto o usuario olhava sob o rotulo (D5,
   ancora tecnica ';;;' do ciclo). Correcao: a ancora migra para J8 (coluna J,
   oculta), D5 recebe '=$D$22' com o estilo do valor e D5:E5 sao mescladas —
   o valor passa a ocupar toda a base do card, exatamente sob o rotulo.
   Unico consumidor da ancora (RESULTADOS!C3) repontado para $J$8.

B. VTA method-aware (causa raiz):
   MEMORIA_RESULTADOS!W48/W50 compunham a execucao historica SOMENTE por
   itens_PC!O+Q (ou fisico implicito T21). No metodo Financeiro sem PCs, os
   ciclos encerrados viravam zero. Camada auxiliar nova:
     W61:W65 = execucao considerada por ciclo (SUMIF financeiro!E por c0..c4;
               inclui preclusos — E preserva o valor quando nao ha efeito);
     W66     = historico pelo metodo oficial ate o ciclo vigente (exclusivo);
     W67     = historico pelo metodo oficial ate a abertura adotada (exclusivo).
   W48 = W67 + W53 + W54; W50 = W66 + CICLO_EM_EXECUCAO!F + G.
   O ramo NAO-Financeiro preserva byte a byte as expressoes homologadas
   (T21+T22 e T21+SUMPRODUCT PC ate W46) — PC e Itens nao mudam de valor.
   O financeiro do ciclo vigente NUNCA entra no historico (indice < vigente),
   logo a posicao fisica valida nao e duplicada.

C. CONTRASTE: fonte residual 8497B0 (contraste ~2,5:1 em fundo claro)
   substituida por 1F4E78 (titulos de card), FFFFFF (chip sobre fundo azul
   escuro) e 595959 (demais textos auxiliares). Fills funcionais intactos.
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


def _cor(hex_rgb: str) -> int:
    texto = hex_rgb.strip().lstrip("#")
    r, g, b = (int(texto[i:i + 2], 16) for i in (0, 2, 4))
    return r + (g << 8) + (b << 16)


COR_RESIDUAL = _cor("8497B0")
COR_TITULO_CARD = _cor("1F4E78")
COR_BRANCO = _cor("FFFFFF")
COR_CINZA = _cor("595959")
FILL_AZUL_ESCURO = _cor("1F4E78")

# Titulos de card sobre fundo claro: azul institucional.
TITULOS_CARD = {"D4", "F4"}

# ---------------------------------------------------------------- formulas ---
W48_ANTIGA = (
    '=IF($W$46="","",ROUND($T$21+SUMPRODUCT((ROW(itens_PC!$O$2:$O$6)-2>=1)'
    '*(ROW(itens_PC!$O$2:$O$6)-2<$W$46)*(itens_PC!$O$2:$O$6+itens_PC!$Q$2:$Q$6))'
    '+$W$53+$W$54,2))'
)
W50_ANTIGA = (
    '=IF(OR($W$49=0,$T$21="",NOT(ISNUMBER($T$21))),"",ROUND($T$21+$T$22'
    '+IFERROR(SUM(INDIRECT("CICLO_EM_EXECUCAO!F13:F211")),0)'
    '+IFERROR(SUM(INDIRECT("CICLO_EM_EXECUCAO!G13:G211")),0),2))'
)

ROTULOS_APOIO = {
    "V60": "APOIO - EXECUCAO HISTORICA PELO METODO OFICIAL",
    "V61": "Financeiro considerado C0 (financeiro!E)",
    "V62": "Financeiro considerado C1 (financeiro!E)",
    "V63": "Financeiro considerado C2 (financeiro!E)",
    "V64": "Financeiro considerado C3 (financeiro!E)",
    "V65": "Financeiro considerado C4 (financeiro!E)",
    "V66": "Historico pelo metodo oficial ate o ciclo vigente (exclusivo)",
    "V67": "Historico pelo metodo oficial ate a abertura adotada (exclusivo)",
}

FORMULAS_APOIO = {
    f"W{61 + n}": (
        f'=ROUND(SUMIF(financeiro!$B$2:$B$73,"c{n}",financeiro!$E$2:$E$73),2)'
    )
    for n in range(5)
}
FORMULAS_APOIO["W66"] = (
    '=IF($T$20="","",IF($B$4="Financeiro",'
    'SUMPRODUCT((ROW($W$61:$W$65)-ROW($W$61)<$T$20)*$W$61:$W$65),'
    'IF(NOT(ISNUMBER($T$21)),"",$T$21+$T$22)))'
)
FORMULAS_APOIO["W67"] = (
    '=IF($W$46="","",IF($B$4="Financeiro",'
    'SUMPRODUCT((ROW($W$61:$W$65)-ROW($W$61)<$W$46)*$W$61:$W$65),'
    'IF(NOT(ISNUMBER($T$21)),"",$T$21+SUMPRODUCT((ROW(itens_PC!$O$2:$O$6)-2>=1)'
    '*(ROW(itens_PC!$O$2:$O$6)-2<$W$46)'
    '*(itens_PC!$O$2:$O$6+itens_PC!$Q$2:$Q$6)))))'
)

W48_NOVA = '=IF(OR($W$46="",$W$67=""),"",ROUND($W$67+$W$53+$W$54,2))'
W50_NOVA = (
    '=IF(OR($W$49=0,$W$66=""),"",ROUND($W$66'
    '+IFERROR(SUM(INDIRECT("CICLO_EM_EXECUCAO!F13:F211")),0)'
    '+IFERROR(SUM(INDIRECT("CICLO_EM_EXECUCAO!G13:G211")),0),2))'
)

C10_NOVA = (
    '=IF(MEMORIA_RESULTADOS!$W$50="",'
    '"POSICAO ATUAL NAO INFORMADA (CICLO_EM_EXECUCAO ausente/incompleta)",'
    'IF(MEMORIA_RESULTADOS!$B$4="Financeiro",'
    '"execucao historica financeiro!E ciclos encerrados (MEMORIA!W66)'
    ' + execucao ate a data + remanescente atual (CICLO_EM_EXECUCAO!F/G)",'
    '"C0 exec (MEMORIA!T21) + ciclos encerrados (MEMORIA!T22)'
    ' + execucao ate a data + remanescente atual (CICLO_EM_EXECUCAO!F/G)"))'
)

C11_PREFIXO_ANTIGO = (
    '="C0 exec (MEMORIA!T21) + ciclos encerrados (MEMORIA!T22)'
)
C11_PREFIXO_NOVO = (
    '=IF(MEMORIA_RESULTADOS!$B$4="Financeiro",'
    '"execucao historica financeiro!E ate a abertura adotada (MEMORIA!W67)",'
    '"C0 exec (MEMORIA!T21) + ciclos encerrados (MEMORIA!T22)")&"'
)

CELULAS_ALTERADAS = {
    "MEMORIA_RESULTADOS": {"W48", "W50", "W61", "W62", "W63", "W64", "W65",
                           "W66", "W67", "V60", "V61", "V62", "V63", "V64",
                           "V65", "V66", "V67"},
    "RESULTADOS": {"C3", "C10", "C11", "D5", "E5", "J8"},
}


# ----------------------------------------------------------------- guardas ---
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


def _endereco(chave: str) -> str:
    linha, coluna = (int(p) for p in chave.split(":"))
    letras = ""
    while coluna:
        coluna, resto = divmod(coluna - 1, 26)
        letras = chr(65 + resto) + letras
    return f"{letras}{linha}"


def _verificar_somente_celulas_previstas(wb, antes: dict[str, dict]) -> None:
    depois = _snapshot_workbook(wb)
    for aba in sorted(set(antes) | set(depois)):
        permitidas = CELULAS_ALTERADAS.get(aba, set())
        chaves = set(antes.get(aba, {})) | set(depois.get(aba, {}))
        divergentes = sorted(
            _endereco(k) for k in chaves
            if antes.get(aba, {}).get(k) != depois.get(aba, {}).get(k)
            and _endereco(k) not in permitidas
        )
        if divergentes:
            raise RuntimeError(
                f"Celulas fora do escopo mudaram em {aba}: {divergentes[:12]}"
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


def _validar_origem(wb) -> None:
    res = wb.Worksheets("RESULTADOS")
    mem = wb.Worksheets("MEMORIA_RESULTADOS")
    if str(res.Range("E5").Formula or "") != "=$D$22":
        raise ValueError("RESULTADOS!E5 fora do leiaute 51A/51C; abortando.")
    if str(res.Range("D5").Formula or "") != "=UPPER(CONTROLE!$B$2)":
        raise ValueError("RESULTADOS!D5 nao e a ancora do ciclo; abortando.")
    if str(mem.Range("W48").Formula or "") != W48_ANTIGA:
        raise ValueError("MEMORIA!W48 diverge da homologada; abortando.")
    if str(mem.Range("W50").Formula or "") != W50_ANTIGA:
        raise ValueError("MEMORIA!W50 diverge da homologada; abortando.")
    for endereco in tuple(ROTULOS_APOIO) + tuple(FORMULAS_APOIO):
        if mem.Range(endereco).Formula not in ("", None):
            raise ValueError(f"MEMORIA!{endereco} ja ocupada; abortando.")
    fin = wb.Worksheets("financeiro")
    if str(fin.Range("E1").Value or "") != "VALOR_ATUALIZADO":
        raise ValueError("financeiro!E1 nao e VALOR_ATUALIZADO; abortando.")
    if str(fin.Range("B74").Value or "") != "TOTAL":
        raise ValueError("financeiro!B74 nao e TOTAL; abortando.")


# ------------------------------------------------------------------ edicao ---
def _aplicar_memoria(wb) -> None:
    mem = wb.Worksheets("MEMORIA_RESULTADOS")
    # Estilo do bloco de apoio: mesmo formato do bloco V/W existente.
    mem.Range("V47:W47").Copy()
    mem.Range("V60:W67").PasteSpecial(Paste=XL_PASTE_FORMATS)
    mem.Range("V40:W40").Copy()
    mem.Range("V60:W60").PasteSpecial(Paste=XL_PASTE_FORMATS)
    wb.Application.CutCopyMode = False
    for endereco, rotulo in ROTULOS_APOIO.items():
        mem.Range(endereco).Value = rotulo
    for endereco, formula in FORMULAS_APOIO.items():
        mem.Range(endereco).Formula = formula
    mem.Range("W48").Formula = W48_NOVA
    mem.Range("W50").Formula = W50_NOVA


def _aplicar_card_retroativo(wb) -> None:
    res = wb.Worksheets("RESULTADOS")
    # 1) ancora tecnica do ciclo migra de D5 para J8 (coluna J oculta).
    res.Range("J8").Formula = "=UPPER(CONTROLE!$B$2)"
    res.Range("J8").NumberFormat = ";;;"
    # 2) unico consumidor da ancora repontado (preserva o texto original).
    c3 = str(res.Range("C3").Formula)
    if "$D$5" not in c3:
        raise RuntimeError("RESULTADOS!C3 nao referencia $D$5; abortando.")
    # C3 usa formato '"CICLO ATUAL   "@' (secao de texto): uma ENTRADA nova
    # seria armazenada como literal. Grava com General e restaura o formato.
    formato_c3 = res.Range("C3").NumberFormatLocal
    res.Range("C3").NumberFormatLocal = "Geral"
    res.Range("C3").Formula = c3.replace("$D$5", "$J$8")
    res.Range("C3").NumberFormatLocal = formato_c3
    if not str(res.Range("C3").Formula).startswith("=IF($J$8"):
        raise RuntimeError("RESULTADOS!C3 nao permaneceu formula; abortando.")
    # 3) o VALOR do retroativo assume a base do card (D5:E5 mescladas).
    res.Range("E5").Copy()
    res.Range("D5").PasteSpecial(Paste=XL_PASTE_FORMATS)
    wb.Application.CutCopyMode = False
    res.Range("E5").ClearContents()
    res.Range("D5").Formula = "=$D$22"
    res.Range("D5:E5").Merge()


def _aplicar_textos_auditaveis(wb) -> None:
    res = wb.Worksheets("RESULTADOS")
    res.Range("C10").Formula = C10_NOVA
    c11 = str(res.Range("C11").Formula)
    if not c11.startswith(C11_PREFIXO_ANTIGO):
        raise RuntimeError("RESULTADOS!C11 fora do padrao homologado; abortando.")
    res.Range("C11").Formula = C11_PREFIXO_NOVO + c11[len(C11_PREFIXO_ANTIGO):]


def _aplicar_contraste(wb) -> None:
    res = wb.Worksheets("RESULTADOS")
    recoloridas = 0
    for linha in range(1, 67):
        for coluna in range(1, 11):
            cel = res.Cells(linha, coluna)
            if int(cel.Font.Color) != COR_RESIDUAL:
                continue
            if int(cel.Interior.Color) == FILL_AZUL_ESCURO:
                cel.Font.Color = COR_BRANCO
            elif _endereco(f"{linha}:{coluna}") in TITULOS_CARD:
                cel.Font.Color = COR_TITULO_CARD
            else:
                cel.Font.Color = COR_CINZA
            recoloridas += 1
    if recoloridas < 40:
        raise RuntimeError(
            f"Esperava >=40 celulas 8497B0 recoloridas; encontrei {recoloridas}."
        )


# ------------------------------------------------------------- orquestracao ---
def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_hotfix_retro_vta_"))
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

        formulas_antes = _snapshot_workbook(wb)
        nomes_antes = {str(n.Name) for n in wb.Names}
        mem = wb.Worksheets("MEMORIA_RESULTADOS")
        travas_antes = {
            endereco: mem.Range(endereco).Formula
            for endereco in ("B26", "T21", "T22", "T23", "T25", "W46", "W49",
                             "W51", "W52", "W53", "W54", "W55", "W56")
        }

        _aplicar_memoria(wb)
        _aplicar_card_retroativo(wb)
        _aplicar_textos_auditaveis(wb)
        _aplicar_contraste(wb)

        _verificar_somente_celulas_previstas(wb, formulas_antes)
        for endereco, formula in travas_antes.items():
            if mem.Range(endereco).Formula != formula:
                raise RuntimeError(f"MEMORIA!{endereco} foi alterada; abortando.")
        if {str(n.Name) for n in wb.Names} != nomes_antes:
            raise RuntimeError("Nomes definidos mudaram; abortando.")

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)

        if aba_ativa in [ws.Name for ws in wb.Worksheets]:
            wb.Worksheets(aba_ativa).Activate()
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
    print("Hotfix RESULTADOS retro/VTA aplicado:", args.destino)


if __name__ == "__main__":
    main()
