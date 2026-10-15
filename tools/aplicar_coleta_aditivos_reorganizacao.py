# -*- coding: utf-8 -*-
"""Coleta 12.1 — reorganizacao controlada da aba aditivos (Pacote 3).

Somente GEOMETRIA e APRESENTACAO. Nenhuma regra muda: as colunas sao movidas
fisicamente pelo proprio Excel (que reescreve todas as referencias, nesta e
nas demais abas, inclusive formatacao condicional e validacao) e a unica
coluna nova, VALOR UNITARIO ATUALIZADO, expoe o termo que a formula do valor
atualizado da alteracao ja usava — que passa a le-la.

    antes (ate 12.0)                       depois (12.1)
    A ITEM                                 A ITEM                        bloco 1
    B Data do aditivo                      B Data do aditivo             identificacao
    C Ciclo/Marco                          C Ciclo/Marco                 (azul claro)
    D Tipo de alteracao                    D Tipo de alteracao
    E Quantidade                           E Quantidade                  bloco 2
    F VU original                          F VU original                 base economica
    G Valor original da alteracao          G Ultimo reajuste no VU  (N)  (areia)
    H Aplicar reajuste?                    H Aplicar reajuste?           bloco 3
    I Fator acumulado aplicavel            I Fator acumulado aplicavel   reajuste
    J Valor atualizado da alteracao        J VU ATUALIZADO (novo)        (verde claro)
    K Considerado no calculo?              K Valor original         (G)  bloco 4 totais
    L DELTA_QTD_CONTRATUAL                 L Valor atualizado       (J)  (teal claro)
    M CHECK_POSICAO_CONTRATUAL             M Considerado no calculo (K)  bloco 5
    N Ultimo reajuste no VU                N DELTA_QTD_CONTRATUAL   (L)  controles
    O INDICE_BASE_ECONOMICA_VU (oculta)    O CHECK_POSICAO_CONTRATUAL(M) (pessego)
                                           P INDICE_BASE_ECONOMICA_VU (O, oculta)

Movimentos (Excel COM, Cut + Insert): N->G; H(valor original)->antes de K;
coluna nova em J. Linhas NAO sao inseridas (cabecalho continua na linha 1,
dados em 2:200). A orientacao de novo item sai do cabecalho A1 e vai para a
mensagem de entrada da validacao de A2:A200 (sem caixa de texto/Shape).

PROVA: antes de gravar, toda formula do workbook e comparada com a traducao
fisica da formula original (_layout_aditivos); J e L sao conferidas com os
geradores de _coleta_oficial. Qualquer divergencia aborta sem gravar.

REGRA ZERO CORRUPCAO XLSX: formulas em ingles, ASCII, parenteses balanceados.

uso: python tools/aplicar_coleta_aditivos_reorganizacao.py [caminho.xlsx]
     (exige pywin32 + Excel real; sem argumento, altera o template oficial)
"""
from __future__ import annotations

import gc
import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
for _caminho in (RAIZ, RAIZ / "tools"):
    if str(_caminho) not in sys.path:
        sys.path.insert(0, str(_caminho))

from openpyxl.utils import get_column_letter  # noqa: E402

import _coleta_oficial as co  # noqa: E402
import aplicar_resultados_executivo_v3 as v3  # noqa: E402
from _layout_aditivos import (  # noqa: E402
    CABECALHOS_ADITIVOS_12_1,
    COLUNAS_ADITIVOS_LEGADO,
    TRADUCAO_LEGADO_PARA_12_1,
    traduzir_formula_legado_para_12_1,
)

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
ABA = "aditivos"
LINHAS = range(2, 201)
XL_TO_RIGHT = -4161

CABECALHO_LEGADO = {
    "B1": "Data do aditivo",
    "C1": "Ciclo/Marco",
    "D1": "TIPO DE ALTERAÇÃO FORMALIZADA",
    "E1": "Quantidade acrescida/suprimida",
    "F1": "Valor unitario original",
    "G1": "Valor original da alteracao",
    "H1": "Aplicar reajuste? (Sim/Nao)",
    "I1": "Fator acumulado aplicavel",
    "J1": "Valor atualizado da alteracao (QTD x VU atualizado)",
    "K1": "CONSIDERADO NO CÁLCULO FINANCEIRO? (SIM/NÃO)",
    "L1": "DELTA_QTD_CONTRATUAL",
    "M1": "CHECK_POSICAO_CONTRATUAL",
    "N1": "ÚLTIMO REAJUSTE JÁ INCORPORADO AO VU",
    "O1": "INDICE_BASE_ECONOMICA_VU",
}

# Blocos de leitura: cabecalho claro, texto escuro (sem fundo escuro).
BLOCOS = (
    ("A", "D", "DDEBF7"),  # 1 identificacao e contexto — azul claro
    ("E", "G", "F2E8D5"),  # 2 quantidade e base economica — areia
    ("H", "J", "E2EFDA"),  # 3 aplicacao do reajuste — verde claro
    ("K", "L", "D5EEEA"),  # 4 totais da alteracao — teal claro
    ("M", "O", "FCE4D6"),  # 5 controles tecnicos — pessego
)
COR_TEXTO_CABECALHO = "1F1F1F"
COR_BORDA_CABECALHO = "BFBFBF"
COR_SEPARADOR_BLOCO = "A6A6A6"
ALTURA_CABECALHO = 54
LARGURAS = {
    "A": 12, "B": 14, "C": 12, "D": 26, "E": 16, "F": 18, "G": 18,
    "H": 14, "I": 14, "J": 18, "K": 20, "L": 20, "M": 18, "N": 23,
    "O": 46,
}

# Orientacao de novo item (antes no cabecalho A1): mensagem de entrada da
# validacao ja existente em A2:A200 — celula, nao Shape. Limite Excel: 255.
TITULO_ORIENTACAO_ITEM = "ITEM DO ADITIVO"
ORIENTACAO_ITEM = (
    "Use o mesmo código de itens_Remanesc.\n"
    "Novo item: cadastrar primeiro em itens_Remanesc como N001, N002...\n"
    "Informe o VU original. Depois, registre aqui a data, a quantidade e os "
    "demais dados do aditivo."
)

# Formatacao condicional depois do movimento: faixas canonicas.
CF_SUPRESSAO = ("A2:F200", "H2:M200")  # cobertura antiga (A:K sem a base) + J
CF_NOVOS_ITENS = ("A2:N200", "P2:P200")  # tudo menos o CHECK (O)


# ------------------------------------------------------------- conferencias
def validar_ascii_e_parenteses() -> None:
    formulas = {}
    for linha in LINHAS:
        formulas[f"J{linha}"] = co._formula_vu_atualizado_aditivo(linha)
        formulas[f"L{linha}"] = co._formula_valor_aditivo(linha)
    v3.validar_ascii_e_parenteses(formulas)
    if len(ORIENTACAO_ITEM) > 255:
        raise RuntimeError("orientacao de ITEM excede 255 caracteres")


def conferir(ws) -> None:
    """Forma 12.0 exigida antes de qualquer escrita (fail-closed)."""
    if ws.ProtectContents:
        raise RuntimeError("aditivos protegida inesperadamente; nada aplicado")
    if not str(ws.Range("A1").Value or "").startswith("ITEM"):
        raise RuntimeError("aditivos!A1 fora do cabecalho esperado; nada aplicado")
    for endereco, texto in CABECALHO_LEGADO.items():
        if ws.Range(endereco).Value != texto:
            raise RuntimeError(f"aditivos!{endereco} fora do cabecalho esperado; nada aplicado")
    if ws.Application.WorksheetFunction.CountA(ws.Range("P1:Z200")):
        raise RuntimeError("aditivos!P:Z nao esta livre; nada aplicado")
    if not ws.Columns("O").Hidden:
        raise RuntimeError("aditivos!O deveria estar oculta; nada aplicado")
    legado = {
        COLUNAS_ADITIVOS_LEGADO["fator"]: co._formula_fator_base_economica,
        COLUNAS_ADITIVOS_LEGADO["check"]: co._formula_status_aditivo,
        COLUNAS_ADITIVOS_LEGADO["indice_base"]: co._formula_indice_base_economica,
    }
    for coluna, funcao in legado.items():
        for linha in LINHAS:
            atual = traduzir_formula_legado_para_12_1(
                ws.Range(f"{coluna}{linha}").Formula, aba_da_formula=ABA
            )
            if atual != funcao(linha):
                raise RuntimeError(f"aditivos!{coluna}{linha} fora da forma 12.0; nada aplicado")


def _tentar(acao, tentativas: int = 60):
    """Repete a chamada COM enquanto o Excel responder 'ocupado'."""
    import pythoncom

    for _ in range(tentativas):
        try:
            return acao()
        except Exception as exc:
            codigo = getattr(exc, "hresult", None)
            if codigo is None and getattr(exc, "args", ()):
                codigo = exc.args[0]
            if codigo != -2147418111:
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(0.5)
    return acao()


def fotografar_formulas(wb) -> dict[tuple[str, str], str]:
    """Todas as formulas do workbook (UsedRange.Formula, um acesso por aba)."""
    foto = {}
    for indice in range(1, _tentar(lambda: wb.Worksheets.Count) + 1):
        ws = _tentar(lambda: wb.Worksheets(indice))
        nome = _tentar(lambda: ws.Name)
        rng = _tentar(lambda: ws.UsedRange)
        valores = _tentar(lambda: rng.Formula)
        if not isinstance(valores, tuple):
            valores = ((valores,),)
        linha0, coluna0 = _tentar(lambda: rng.Row), _tentar(lambda: rng.Column)
        for i, linha in enumerate(valores):
            for j, formula in enumerate(linha):
                if isinstance(formula, str) and formula.startswith("="):
                    foto[(nome, f"{get_column_letter(coluna0 + j)}{linha0 + i}")] = formula
    return foto


def conferir_formulas(antes, depois) -> None:
    """Toda formula = traducao fisica da original; J e L = geradores."""
    esperado = {}
    for (aba, coord), formula in antes.items():
        if aba == ABA:
            coluna = coord.rstrip("0123456789")
            coord = TRADUCAO_LEGADO_PARA_12_1.get(coluna, coluna) + coord[len(coluna):]
        esperado[(aba, coord)] = traduzir_formula_legado_para_12_1(formula, aba_da_formula=aba)
    for linha in LINHAS:
        esperado[(ABA, f"J{linha}")] = co._formula_vu_atualizado_aditivo(linha)
        esperado[(ABA, f"L{linha}")] = co._formula_valor_aditivo(linha)
    faltam = sorted(set(esperado) - set(depois))
    sobram = sorted(set(depois) - set(esperado))
    divergem = sorted(k for k in esperado.keys() & depois.keys() if esperado[k] != depois[k])
    if faltam or sobram or divergem:
        raise RuntimeError(
            f"formulas fora da traducao: faltam={faltam[:3]} sobram={sobram[:3]} "
            f"divergem={divergem[:3]}; nada gravado"
        )


# ------------------------------------------------------------- etapas COM
def mover_colunas(ws) -> None:
    ws.Columns("N:N").Cut()
    ws.Columns("G:G").Insert(Shift=XL_TO_RIGHT)   # base do VU -> G
    ws.Columns("H:H").Cut()
    ws.Columns("K:K").Insert(Shift=XL_TO_RIGHT)   # valor original -> J (K apos a nova)
    ws.Columns("J:J").Insert(Shift=XL_TO_RIGHT)   # nova coluna J
    ws.Application.CutCopyMode = False


def escrever_formulas(ws) -> None:
    for linha in LINHAS:
        ws.Range(f"J{linha}").Formula = co._formula_vu_atualizado_aditivo(linha)
        ws.Range(f"L{linha}").Formula = co._formula_valor_aditivo(linha)


def formatar_coluna_nova(ws) -> None:
    """J recebe o padrao das colunas calculadas monetarias (F)."""
    modelo = ws.Range("F2")
    alvo = ws.Range("J2:J200")
    alvo.NumberFormat = modelo.NumberFormat
    alvo.Interior.Color = modelo.Interior.Color
    alvo.Font.Name = modelo.Font.Name
    alvo.Font.Size = modelo.Font.Size
    alvo.Font.Bold = modelo.Font.Bold
    alvo.Font.Color = modelo.Font.Color
    alvo.HorizontalAlignment = modelo.HorizontalAlignment
    alvo.VerticalAlignment = modelo.VerticalAlignment
    alvo.WrapText = modelo.WrapText
    alvo.Locked = modelo.Locked
    for indice in (7, 8, 9, 10, 12):  # bordas esquerda/topo/base/direita/interna
        borda = alvo.Borders(indice)
        borda.LineStyle = v3.XL_CONTINUOUS
        borda.Weight = v3.XL_THIN
        borda.Color = v3._bgr("D9D9D9")


def cabecalho(ws) -> None:
    for coluna, texto in CABECALHOS_ADITIVOS_12_1.items():
        if coluna == "P":
            continue  # tecnica oculta: rotulo e estilo preservados
        ws.Range(f"{coluna}1").Value = texto
    ws.Rows(1).RowHeight = ALTURA_CABECALHO
    for inicio, fim, cor in BLOCOS:
        rng = ws.Range(f"{inicio}1:{fim}1")
        rng.Interior.Color = v3._bgr(cor)
        rng.Font.Bold = True
        rng.Font.Size = 10
        rng.Font.Color = v3._bgr(COR_TEXTO_CABECALHO)
        rng.HorizontalAlignment = v3.XL_CENTER
        rng.VerticalAlignment = v3.XL_VCENTER
        rng.WrapText = True
        for indice in (7, 8, 9, 10, 11):
            borda = rng.Borders(indice)
            borda.LineStyle = v3.XL_CONTINUOUS
            borda.Weight = v3.XL_THIN
            borda.Color = v3._bgr(COR_BORDA_CABECALHO)
    for inicio, _fim, _cor in BLOCOS[1:]:
        borda = ws.Range(f"{inicio}1:{inicio}200").Borders(7)  # separador de bloco
        borda.LineStyle = v3.XL_CONTINUOUS
        borda.Weight = v3.XL_THIN
        borda.Color = v3._bgr(COR_SEPARADOR_BLOCO)
    for coluna, largura in LARGURAS.items():
        ws.Columns(coluna).ColumnWidth = largura
    ws.Columns("P").Hidden = True


def orientacao_item(ws) -> None:
    validacao = ws.Range("A2:A200").Validation
    validacao.InputTitle = TITULO_ORIENTACAO_ITEM
    validacao.InputMessage = ORIENTACAO_ITEM
    validacao.ShowInput = True


def _enderecos(rng) -> set[str]:
    return {a.Address.replace("$", "") for a in rng.Areas}


def normalizar_formatacao_condicional(ws) -> None:
    app = ws.Application
    suprimir = app.Union(*(ws.Range(f) for f in CF_SUPRESSAO))
    novos = app.Union(*(ws.Range(f) for f in CF_NOVOS_ITENS))
    regras = ws.Cells.FormatConditions
    ajustadas = 0
    for i in range(1, regras.Count + 1):
        regra = regras.Item(i)
        faixa = _enderecos(regra.AppliesTo)
        if faixa == {"A2:M200"}:
            regra.ModifyAppliesToRange(suprimir)
            ajustadas += 1
        elif faixa == {"A2:N200", "P2:P200"}:
            regra.ModifyAppliesToRange(novos)
            ajustadas += 1
    if ajustadas != 3 or regras.Count != 9:
        raise RuntimeError(f"aditivos: formatacao condicional inesperada ({ajustadas}/{regras.Count})")


def aplicar(caminho: Path) -> None:
    import pythoncom
    import win32com.client as com

    validar_ascii_e_parenteses()
    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = _tentar(lambda: excel.Workbooks.Open(str(caminho)))
        time.sleep(1.0)
        _tentar(lambda: setattr(excel, "Calculation", v3.XL_CALCULO_MANUAL))
        ws = _tentar(lambda: wb.Worksheets(ABA))
        _tentar(lambda: conferir(ws))
        antes = fotografar_formulas(wb)
        for etapa in (
            mover_colunas, escrever_formulas, formatar_coluna_nova, cabecalho,
            orientacao_item, normalizar_formatacao_condicional,
        ):
            _tentar(lambda: etapa(ws))
        conferir_formulas(antes, fotografar_formulas(wb))
        excel.Calculation = v3.XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()
        ws.Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        ws.Range("A2").Select()
        wb.Worksheets("CONTROLE").Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        wb.Save()
        v3._fechar(wb)
        wb = None
    finally:
        if wb is not None:
            v3._fechar(wb)
        for _ in range(10):
            try:
                excel.Quit()
                break
            except Exception:
                time.sleep(1.0)
        wb = excel = None
        gc.collect()
        pythoncom.CoUninitialize()


def main() -> int:
    caminho = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else TEMPLATE
    if not caminho.exists():
        print("ERRO: arquivo nao encontrado: %s" % caminho)
        return 1
    aplicar(caminho)
    print("OK: aba aditivos reorganizada (Coleta 12.1) em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
