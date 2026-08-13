# -*- coding: utf-8 -*-
"""Etapa 51A: acabamento controlado da aba RESULTADOS, via Excel COM.

ESCOPO: APRESENTACAO. Nenhuma regra de negocio e alterada. O diagnostico da
51A comprovou que a cadeia do retroativo (RETRO_OFICIAL -> RESULTADOS!D22 ->
card E5) esta correta; este aplicador NAO toca calculo algum. Mudancas:

  1. E6 (rodape do card 2) vira formula de estado vazio: enquanto D22 nao
     tem retroativo apurado, orienta o usuario a selecionar o metodo (via
     $J$6, que ja existia) ou informar a base. Nao substitui o "—" do valor
     e nao inventa R$ 0,00: ausencia de base nao e retroativo zero.
  2. G3:H3 mesclada para a faixa "VARIACAO ACUMULADA" nao exibir ##### com
     dados reais (H3 comprovadamente livre: sem formula, sem consumidor,
     fora das ancoras).
  3. Formato-lixo '\\Pyyd\\ryy\\o' (regressao da Etapa 50: NumberFormatLocal
     "Padrão" nao e o token de Geral) corrigido para General nas celulas de
     apresentacao afetadas.
  4. Larguras das colunas A:H ajustadas por conteudo com piso/teto por
     coluna (largura_final = min(max(autofit, minimo), maximo)); nenhuma
     coluna encolhe abaixo do homologado na Etapa 50.
  5. ShrinkToFit nos valores dos cards (C5/E5/G5): moeda de qualquer
     magnitude aparece completa, sem #####.
  6. Contraste do Quadro 1 (1. COMPOSICAO DO VTA): cabecalho branco sobre
     azul institucional; rotulos e descricoes das linhas 10-13 escurecidos.

PRESERVADO (guardas abortam se divergirem): todas as formulas da aba exceto
E6; RESULTADOS!D5 (ancora do CICLO ATUAL — nao e o retroativo); MEMORIA_
RESULTADOS integral, incluindo W48; nomes definidos; linhas separadoras.
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
XL_SHEET_HIDDEN = 0
XL_SHEET_VISIBLE = -1

ABA_MEMORIA = "MEMORIA_RESULTADOS"
ABA_RESULTADOS = "RESULTADOS"
TITULO_OFICIAL = "RESULTADOS CONSOLIDADOS — REAJUSTE CONTRATUAL"

# Token do defeito latente da Etapa 50 (NumberFormatLocal "Padrão" gravado
# literalmente como formato de data). O XLSX armazena '\Pyyd\ryy\o'; o Excel
# COM pt-BR devolve a renderizacao localizada 'Paad\raa\o'. Deteccao por
# IGUALDADE EXATA — nunca por fragmento, para jamais tocar formatos legitimos
# como 'dd/mm/aaaa' ou ';;;'.
FORMATOS_LIXO = ("Paad\\raa\\o", "\\Pyyd\\ryy\\o")

# Estado vazio do card 2. Le exclusivamente celulas que ja existiam:
# D22 (retroativo oficial na aba) e J6 (flag "metodo selecionado").
FORMULA_ESTADO_VAZIO_E6 = (
    '=IF($D$22<>"","Ciclos apurados",'
    'IF($J$6=0,"SELECIONE O MÉTODO NA ABA CONTROLE",'
    '"INFORME A BASE DO MÉTODO"))'
)

# largura_final = min(max(autofit, minimo), maximo). Pisos = larguras
# homologadas na Etapa 50 (nada encolhe); tetos evitam deformar o dashboard.
LARGURAS = {
    "A": (32.27, 36.0),
    "B": (23.27, 26.0),
    "C": (14.0, 16.0),   # unica sem largura definida na Etapa 50 (8,43)
    "D": (35.27, 38.0),
    "E": (23.27, 26.0),
    "F": (15.27, 18.0),
    "G": (19.27, 22.0),
    "H": (25.27, 28.0),
}

LINHAS_BRANCAS = (8, 14, 23, 32, 39, 52)
LINHAS_OCULTAS = (31, 40, 51)

_PROT_FLAGS = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
    "AllowFiltering", "AllowUsingPivotTables",
)


def _cor(hex_rgb: str) -> int:
    texto = hex_rgb.strip().lstrip("#")
    r, g, b = (int(texto[i : i + 2], 16) for i in (0, 2, 4))
    return r + (g << 8) + (b << 16)


CORES = {
    "azul_escuro": _cor("1F4E78"),
    "branco": _cor("FFFFFF"),
    "cinza_texto": _cor("595959"),
    "cinza_escuro": _cor("404040"),
}


# --------------------------------------------------------------------------- #
# Protecao
# --------------------------------------------------------------------------- #

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


# --------------------------------------------------------------------------- #
# Guardas de contrato
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


def _validar_origem(wb) -> None:
    abas = [ws.Name for ws in wb.Worksheets]
    if ABA_RESULTADOS not in abas or ABA_MEMORIA not in abas:
        raise ValueError("Template sem RESULTADOS/MEMORIA_RESULTADOS.")
    res = wb.Worksheets(ABA_RESULTADOS)
    if str(res.Range("A1").Value or "") != TITULO_OFICIAL:
        raise ValueError("RESULTADOS de origem nao corresponde ao layout oficial.")
    if str(res.Range("D4").Value or "") != "RETROATIVO TOTAL A PAGAR":
        raise ValueError("Leiaute da Etapa 50 nao encontrado (D4 divergente).")
    if str(res.Range("E5").Formula or "") != "=$D$22":
        raise ValueError("Card do retroativo (E5) nao espelha D22; abortando.")
    if str(res.Range("D5").Formula or "") != "=UPPER(CONTROLE!$B$2)":
        raise ValueError(
            "RESULTADOS!D5 (ancora do CICLO ATUAL) divergente; abortando."
        )
    h3 = res.Range("H3")
    if str(h3.Formula or ""):
        raise ValueError(f"H3 deveria estar livre; encontrado {h3.Formula!r}.")
    if bool(h3.MergeCells):
        raise ValueError("H3 ja participa de mescla; leiaute inesperado.")


def _verificar_formulas(res, antes: dict[str, object]) -> None:
    depois = _snapshot_aba(res)
    permitidas = {"6:5"}  # E6 (linha 6, coluna 5): unico conteudo alterado
    divergentes = sorted(
        k for k in set(antes) | set(depois)
        if antes.get(k) != depois.get(k) and k not in permitidas
    )
    if divergentes:
        raise RuntimeError(
            "Formulas de RESULTADOS alteradas fora do escopo 51A: "
            + ", ".join(divergentes[:10])
        )
    if depois.get("6:5") != FORMULA_ESTADO_VAZIO_E6:
        raise RuntimeError(f"E6 nao recebeu o estado vazio: {depois.get('6:5')!r}")


def _verificar_memoria(memoria, antes: dict[str, object], w48_antes) -> None:
    depois = _snapshot_aba(memoria)
    if depois != antes:
        divergentes = sorted(
            k for k in set(antes) | set(depois) if antes.get(k) != depois.get(k)
        )
        raise RuntimeError(
            f"MEMORIA_RESULTADOS foi modificada ({len(divergentes)} celulas): "
            + ", ".join(divergentes[:10])
        )
    if memoria.Range("W48").Formula != w48_antes:
        raise RuntimeError("MEMORIA_RESULTADOS!W48 foi alterada; abortando.")


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


def _verificar_linhas_separadoras(res) -> None:
    for linha in LINHAS_OCULTAS:
        if not bool(res.Rows(linha).Hidden):
            raise RuntimeError(f"Linha oculta {linha} deixou de estar oculta.")
    for linha in LINHAS_BRANCAS:
        for coluna in "ABCDEFGH":
            cel = res.Range(f"{coluna}{linha}")
            if str(cel.Text or ""):
                raise RuntimeError(
                    f"Separador {coluna}{linha} passou a renderizar texto."
                )


def _verificar_destino_excel(excel, caminho: Path) -> None:
    wb = excel.Workbooks.Open(str(caminho), UpdateLinks=0, ReadOnly=True, CorruptLoad=0)
    try:
        abas = [ws.Name for ws in wb.Worksheets]
        if abas[-2:] != [ABA_MEMORIA, ABA_RESULTADOS]:
            raise RuntimeError(f"Ordem final de abas inesperada: {abas[-3:]}")
        if wb.Worksheets(ABA_MEMORIA).Visible != XL_SHEET_HIDDEN:
            raise RuntimeError("MEMORIA_RESULTADOS nao ficou oculta.")
        res = wb.Worksheets(ABA_RESULTADOS)
        if res.Visible != XL_SHEET_VISIBLE:
            raise RuntimeError("RESULTADOS nao ficou visivel.")
        if not bool(res.Range("G3").MergeCells):
            raise RuntimeError("Mescla G3:H3 nao persistiu.")
        nomes = {str(n.Name).split("!")[-1] for n in wb.Names}
        obrigatorios = {
            "VTA_FINAL", "RETRO_OFICIAL", "REM_BASE_OFICIAL",
            "REM_ATUALIZADO_OFICIAL", "STATUS_RESULTADOS",
            "VTA_ATUALIZACAO_CHEIA", "EXECUCAO_ATUALIZADA_CICLO",
            "SALDO_REMANESCENTE_ATUAL", "OPCOES_APLICAR_MANUAL",
        }
        faltantes = sorted(obrigatorios.difference(nomes))
        if faltantes:
            raise RuntimeError(f"Nomes definidos ausentes: {faltantes}")
    finally:
        wb.Close(SaveChanges=False)


# --------------------------------------------------------------------------- #
# Mudancas 51A (todas de apresentacao)
# --------------------------------------------------------------------------- #

def _definir_geral(cel) -> None:
    """Formato Geral independente de locale, com fallback localizado."""
    import pywintypes

    try:
        cel.NumberFormat = "General"
    except pywintypes.com_error:
        cel.NumberFormatLocal = "Geral"


def _corrigir_formatos_lixo(res) -> list[str]:
    corrigidas: list[str] = []
    for linha in range(1, 71):
        for coluna in "ABCDEFGHIJ":
            cel = res.Range(f"{coluna}{linha}")
            try:
                formato = str(cel.NumberFormat or "")
            except Exception:
                # Formatos mistos em area mesclada devolvem Null; celula a
                # celula isso nao ocorre, mas o guarda evita aborto cego.
                continue
            if formato in FORMATOS_LIXO:
                _definir_geral(cel)
                corrigidas.append(f"{coluna}{linha}")
    return corrigidas


def _mesclar_variacao(res) -> None:
    res.Range("G3:H3").MergeCells = True


def _estado_vazio_card(res) -> None:
    cel = res.Range("E6")
    cel.Formula = FORMULA_ESTADO_VAZIO_E6
    _definir_geral(cel)


def _ajustar_larguras(res) -> None:
    for coluna, (minimo, maximo) in LARGURAS.items():
        res.Columns(f"{coluna}:{coluna}").AutoFit()
        automatica = float(res.Range(f"{coluna}1").ColumnWidth)
        final = min(max(automatica, minimo), maximo)
        res.Columns(f"{coluna}:{coluna}").ColumnWidth = final


def _garantir_valores_completos(res) -> None:
    for endereco in ("C5", "E5", "G5"):
        cel = res.Range(endereco)
        cel.WrapText = False
        cel.ShrinkToFit = True


def _contraste_quadro1(res) -> None:
    # Cabecalho da tabela do Quadro 1: branco pleno sobre azul institucional.
    for endereco in ("B9", "C9", "F9", "H9"):
        res.Range(endereco).Font.Color = CORES["branco"]
    # Rotulos das linhas comparativas: azul institucional em vez de cinza-azulado.
    for linha in (11, 12, 13):
        res.Range(f"A{linha}").Font.Color = CORES["azul_escuro"]
        res.Range(f"B{linha}").Font.Color = CORES["cinza_escuro"]
    # Descricoes e fontes das linhas 10-13: cinza legivel em vez de cinza-azulado.
    for linha in (10, 11, 12, 13):
        for coluna in ("C", "F", "H"):
            res.Range(f"{coluna}{linha}").Font.Color = CORES["cinza_texto"]


def _aplicar_51a(res) -> list[str]:
    corrigidas = _corrigir_formatos_lixo(res)
    _mesclar_variacao(res)
    _estado_vazio_card(res)
    _ajustar_larguras(res)
    _garantir_valores_completos(res)
    _contraste_quadro1(res)
    return corrigidas


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

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_51a_"))
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

        res = wb.Worksheets(ABA_RESULTADOS)
        memoria = wb.Worksheets(ABA_MEMORIA)
        formulas_antes = _snapshot_aba(res)
        memoria_antes = _snapshot_aba(memoria)
        w48_antes = memoria.Range("W48").Formula

        estado, selecao = _capturar_protecao(res)
        corrigidas = _aplicar_51a(res)
        _verificar_formulas(res, formulas_antes)

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)
        _verificar_linhas_separadoras(res)
        _verificar_memoria(memoria, memoria_antes, w48_antes)
        _restaurar_protecao(res, estado, selecao)

        if aba_ativa in [ws.Name for ws in wb.Worksheets] and aba_ativa != ABA_MEMORIA:
            wb.Worksheets(aba_ativa).Activate()
        else:
            wb.Worksheets("CONTROLE").Activate()
        wb.Save()
        salvo = True
        wb.Close(SaveChanges=False)
        wb = None
        _verificar_destino_excel(excel, tmp_xlsx)
        print(f"Formatos-lixo corrigidos ({len(corrigidas)}): {corrigidas}")
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
    print("RESULTADOS (acabamento 51A) aplicada:", args.destino)


if __name__ == "__main__":
    main()
