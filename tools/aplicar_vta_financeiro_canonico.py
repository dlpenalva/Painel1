# -*- coding: utf-8 -*-
"""VTA-M2/VTA-M2.1 — VTA do metodo Financeiro alinhado a identidade canonica no XLS.

Aplica, via Excel COM em copia temporaria (padrao zero-corrupcao ja usado por
`aplicar_vta_posicoes_tabela1.py`):

Identidade para o metodo Financeiro (CONTROLE!B1 -> MEMORIA_RESULTADOS!B4 =
"Financeiro"):

    VTA = passado efetivamente desembolsado (financeiro!C, informado)
        + acertos ainda devidos (B21/B16, retroativo oficial — inalterado)
        + futuro remanescente atualizado (D35, quando B4="Financeiro" ja
          resolve para D32 — inalterado)

* MEMORIA_RESULTADOS:
    - C20/D20 (novas): "Desembolsado (Financeiro)" = SOMA financeiro!C
      informado, sem filtrar por EFEITO_FINANCEIRO (o passado e o que foi
      pago, independente de ja estar dentro do periodo de efeito).
    - B26 (VTA FINAL): ganha um ramo NOVO exclusivamente para
      $B$4="Financeiro" (D20+B21+D35+ajustes). O ramo PC ($B$4="PCs", usa
      T21/T22/T23/T25) e o ramo Itens/Consumido (formula generica B23) NAO
      sao alterados — permanecem byte-a-byte identicos.
    - B28 (referencia comparativa): passa a ser condicional por metodo —
      Financeiro mostra $B$23 (formula generica antiga, agora so
      comparativa); PC e Itens continuam mostrando $B$26 exatamente como
      hoje (nenhuma mudanca de comportamento para esses dois metodos).
* RESULTADOS (apenas linhas novas, apos a linha 66 ja ocupada — nenhuma
  estrutura existente e deslocada):
    - Bloco "7. METODOLOGIA DO VTA" — linha 1 e texto fixo (a identidade
      vale para os tres metodos); linha 2 e FORMULA condicional a
      MEMORIA_RESULTADOS!$B$4 (Financeiro/PCs/Itens), nunca texto fixo
      mencionando "Financeiro" quando o metodo selecionado for outro.
    - Bloco "8. CONFERENCIA DA EXECUCAO" — titulo condicional (so se
      apresenta como auditoria do Financeiro quando $B$4="Financeiro").
      Quando o metodo NAO e Financeiro, cada celula de dado (colunas
      B-E) mostra o rotulo "Nao aplicavel ao metodo selecionado" (nunca
      zero, nunca numero do Financeiro atribuido a PC/Consumido).
      Quando o metodo E Financeiro: desembolsado informado x execucao
      teorica pelo quantitativo (VU do proprio ciclo, nunca VU_C0), por
      ciclo, com status OK/REVISAR/NAO COMPARAVEL (comparacao exata, sem
      tolerancia percentual arbitraria). "Execucao teorica" usa uma
      CADEIA DE FALLBACK do checkpoint anterior (REM_BASE do ciclo n-1,
      n-2, ... ate a quantidade contratada C0 em E) — necessaria porque
      o caso real pode ter ciclos intermediarios nao coletados
      (posicao_contratual!J/N em branco) sem que isso invalide a
      comparacao do ciclo seguinte que TEM checkpoint (ex.: C3 com R
      preenchido).

Trava anti-regressao: T21/T22/T23/T25, o ramo PC dentro de B26 (a
subexpressao literal do IF($B$4="PCs", ...)) e o ramo Itens/Consumido
(a subexpressao literal do ELSE generico) precisam permanecer
identicos antes/depois, verificado por assinatura de texto. A area de
RESULTADOS linhas 68-77 e validada como vazia (ou ja no formato
VTA-M2/M2.1 esperado) antes da escrita, para nao sobrescrever
silenciosamente conteudo futuro de outra frente.
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_CONTINUOUS = 1
XL_THIN = 2

ABA_MEMORIA = "MEMORIA_RESULTADOS"
ABA_RESULTADOS = "RESULTADOS"

MOEDA_BR = "R$ #.##0,00"

_PROT_FLAGS = (
    "AllowFormattingCells",
    "AllowFormattingColumns",
    "AllowFormattingRows",
    "AllowInsertingColumns",
    "AllowInsertingRows",
    "AllowInsertingHyperlinks",
    "AllowDeletingColumns",
    "AllowDeletingRows",
    "AllowSorting",
    "AllowFiltering",
    "AllowUsingPivotTables",
)


def _cor(hex_rgb: str) -> int:
    texto = hex_rgb.strip().lstrip("#")
    r, g, b = (int(texto[i:i + 2], 16) for i in (0, 2, 4))
    return r + (g << 8) + (b << 16)


CORES = {
    "azul_muito_claro": _cor("EEF4F8"),
    "borda": _cor("B0C4D8"),
}

# Ramo PC dentro de B26 — precisa permanecer literalmente identico.
_RAMO_PC_B26 = (
    'IF($T$25="CALCULO MANUAL REQUERIDO","",ROUND($T$25+IF(ISNUMBER(B24),B24,0),2))'
)
_B26_ANTIGO = (
    '=IF(AND(B24<>"",B25<>""),"",IF(ISNUMBER(B25),B25,IF($B$4="PCs",'
    + _RAMO_PC_B26
    + ',IF(OR(B23="",AND(B24<>"",NOT(ISNUMBER(B24)))),"",'
    'ROUND(B23+IF(ISNUMBER(B24),B24,0)+IF(ISNUMBER($N$263),$N$263,0),2)))))'
)
_RAMO_ITENS_B26 = (
    'IF(OR(B23="",AND(B24<>"",NOT(ISNUMBER(B24)))),"",'
    'ROUND(B23+IF(ISNUMBER(B24),B24,0)+IF(ISNUMBER($N$263),$N$263,0),2))'
)
_RAMO_FINANCEIRO_B26 = (
    'IF(OR($D$20="",B21="",D35="",AND(B24<>"",NOT(ISNUMBER(B24)))),"",'
    'ROUND($D$20+B21+D35+IF(ISNUMBER(B24),B24,0)+IF(ISNUMBER($N$263),$N$263,0),2))'
)
_B26_NOVO = (
    '=IF(AND(B24<>"",B25<>""),"",IF(ISNUMBER(B25),B25,IF($B$4="PCs",'
    + _RAMO_PC_B26
    + ',IF($B$4="Financeiro",'
    + _RAMO_FINANCEIRO_B26
    + ','
    + _RAMO_ITENS_B26
    + '))))'
)

_D20_FORMULA = (
    '=IF(COUNTIF(financeiro!$C$2:$C$73,"<>")=0,"",'
    'ROUND(SUM(financeiro!$C$2:$C$73),2))'
)

_B28_NOVO = '=IF($B$4="Financeiro",$B$23,$B$26)'


def _nomes_abas(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


def _nomes_definidos(wb) -> set[str]:
    return {str(nome.Name).split("!")[-1] for nome in wb.Names}


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


def _bordas(rng) -> None:
    for indice in (7, 8, 9, 10, 11, 12):
        borda = rng.Borders(indice)
        borda.LineStyle = XL_CONTINUOUS
        borda.Weight = XL_THIN
        borda.Color = CORES["borda"]


def _validar_origem(wb) -> None:
    abas = _nomes_abas(wb)
    obrigatorias = {
        ABA_MEMORIA, ABA_RESULTADOS, "financeiro", "posicao_contratual",
        "historico_VU", "itens_Remanesc", "CONTROLE",
    }
    ausentes = sorted(obrigatorias.difference(abas))
    if ausentes:
        raise ValueError(f"Abas obrigatorias ausentes: {', '.join(ausentes)}")
    mem = wb.Worksheets(ABA_MEMORIA)
    for endereco in ("B4", "B16", "B20", "B21", "B22", "B23", "B26", "B28", "T25"):
        if not str(mem.Range(endereco).Formula or ""):
            raise ValueError(f"MEMORIA_RESULTADOS!{endereco} ausente; layout inesperado.")
    atual = str(mem.Range("B26").Formula)
    if atual != _B26_ANTIGO:
        raise ValueError(
            "MEMORIA_RESULTADOS!B26 nao corresponde ao formato esperado "
            "(template ja mudou desde a auditoria VTA-M1). Formula atual:\n"
            f"{atual}"
        )
    if "VTA_FINAL" not in _nomes_definidos(wb):
        raise ValueError("Nome definido VTA_FINAL ausente.")
    if mem.Range("C20").Value not in (None, ""):
        raise ValueError("MEMORIA_RESULTADOS!C20 ja ocupada; escolher outra celula.")
    if mem.Range("D20").Value not in (None, ""):
        raise ValueError("MEMORIA_RESULTADOS!D20 ja ocupada; escolher outra celula.")
    _validar_area_resultados(wb)


def _validar_area_resultados(wb) -> None:
    """Trava: a area RESULTADOS!A68:E77 deve estar vazia (template
    pristino) antes da escrita — protege contra sobrescrita silenciosa de
    conteudo futuro de outra frente que tenha ocupado essas linhas."""
    res = wb.Worksheets(ABA_RESULTADOS)
    linha_final = _LINHA_CONFERENCIA_C0 + 5 - 1
    bloco = res.Range(f"A{_LINHA_METODOLOGIA}:E{linha_final}")
    ocupadas = [
        celula.Address()
        for celula in bloco.Cells
        if celula.Value not in (None, "")
    ]
    if ocupadas:
        raise ValueError(
            "RESULTADOS!A68:E77 nao esta vazia (encontrado conteudo em "
            f"{', '.join(ocupadas[:5])}{'...' if len(ocupadas) > 5 else ''}); "
            "area reservada ao bloco VTA-M2/M2.1 foi ocupada por outra frente."
        )


def _snapshot_travas(wb) -> dict[str, str]:
    mem = wb.Worksheets(ABA_MEMORIA)
    return {e: str(mem.Range(e).Formula) for e in ("T21", "T22", "T23", "T25")}


def _aplicar_memoria(wb) -> None:
    mem = wb.Worksheets(ABA_MEMORIA)
    estado, selecao = _capturar_protecao(mem)
    try:
        mem.Range("C20").Value = "Desembolsado (Financeiro)"
        mem.Range("D20").Formula = _D20_FORMULA
        mem.Range("B26").Formula = _B26_NOVO
        mem.Range("B28").Formula = _B28_NOVO
    finally:
        _restaurar_protecao(mem, estado, selecao)


_LINHA_METODOLOGIA = 68
_LINHA_CONFERENCIA_TITULO = 71
_LINHA_CONFERENCIA_CABECALHO = 72
_LINHA_CONFERENCIA_C0 = 73

_TEXTO_METODOLOGIA_1 = (
    "VTA = execucao ja realizada + acertos ainda devidos + saldo futuro atualizado."
)

_NAO_APLICAVEL = "Nao aplicavel ao metodo selecionado"


def _condicionar_financeiro(corpo_formula: str, nao_aplicavel: str = _NAO_APLICAVEL) -> str:
    """Envolve `corpo_formula` (sem o '=' inicial) para so ser avaliada
    quando MEMORIA_RESULTADOS!$B$4="Financeiro"; caso contrario, mostra o
    rotulo `nao_aplicavel` (nunca zero, nunca numero do Financeiro
    atribuido a PC/Consumido)."""
    return (
        '=IF(MEMORIA_RESULTADOS!$B$4<>"Financeiro","' + nao_aplicavel + '",'
        + corpo_formula.lstrip("=") + ")"
    )


def _formula_metodologia_execucao() -> str:
    return (
        '=IF(MEMORIA_RESULTADOS!$B$4="Financeiro",'
        '"Metodo Financeiro: execucao ja realizada = valores efetivamente '
        'desembolsados informados no financeiro.",'
        'IF(MEMORIA_RESULTADOS!$B$4="PCs",'
        '"Metodo PCs: execucao ja realizada = valores apurados pelos '
        'Pedidos de Compra conforme as regras do metodo.",'
        'IF(MEMORIA_RESULTADOS!$B$4="Itens",'
        '"Metodo Consumido: execucao ja realizada = quantidades '
        'consumidas x valores unitarios aplicaveis.",'
        '"Selecione o metodo em CONTROLE!B1.")))'
    )


def _formula_titulo_conferencia() -> str:
    return (
        '=IF(MEMORIA_RESULTADOS!$B$4="Financeiro",'
        '"8. CONFERENCIA DA EXECUCAO (Financeiro) - camada de auditoria; '
        'nao recalcula o VTA oficial",'
        '"8. CONFERENCIA DA EXECUCAO - auditoria especifica do metodo '
        'Financeiro; nao aplicavel ao metodo selecionado")'
    )


def _formula_desembolsado_ciclo(ciclo: str) -> str:
    corpo = (
        f'IF(COUNTIFS(financeiro!$B$2:$B$73,"{ciclo}",'
        'financeiro!$C$2:$C$73,"<>")=0,"",'
        f'ROUND(SUMIFS(financeiro!$C$2:$C$73,financeiro!$B$2:$B$73,"{ciclo}"),2))'
    )
    return _condicionar_financeiro(corpo)


# Cadeia de fallback do checkpoint "anterior" por ciclo, do mais recente ao
# mais antigo, terminando sempre na quantidade contratada C0 (E) — provado
# contra os cabecalhos reais de posicao_contratual: REM_BASE_Cn (F/J/N/R/V)
# e um snapshot bruto vindo de itens_Remanesc a cada rodada de coleta, por
# isso pode faltar quando um ciclo intermediario nao foi coletado (caso real:
# C1/C2 em branco, mas C3 com dado). A cadeia evita marcar NAO COMPARAVEL por
# erro de formula quando existe base quantitativa anterior disponivel, ainda
# que nao seja a do ciclo imediatamente anterior.
_MAPA_ANTERIOR_CANDIDATOS = {
    "C0": ["E"],
    "C1": ["F", "E"],
    "C2": ["J", "F", "E"],
    "C3": ["N", "J", "F", "E"],
    "C4": ["R", "N", "J", "F", "E"],
}
_MAPA_ATUAL = {"C0": "F", "C1": "J", "C2": "N", "C3": "R", "C4": "V"}
# Colunas de historico_VU (VU_C0..VU_C4) — provado contra o cabecalho real.
_MAPA_VU = {"C0": "C", "C1": "D", "C2": "E", "C3": "F", "C4": "G"}


def _col_range(nome_aba: str, coluna: str) -> str:
    return f"{nome_aba}!${coluna}$2:${coluna}$201"


def _anterior_efetivo_expr(candidatos: list[str]) -> str:
    *resto, base = candidatos
    expr = f"N({_col_range('posicao_contratual', base)})"
    for coluna in reversed(resto):
        rng = _col_range("posicao_contratual", coluna)
        expr = f"IF(ISNUMBER({rng}),N({rng}),{expr})"
    return expr


def _formula_execucao_teorica_ciclo(ciclo: str) -> str:
    candidatos = _MAPA_ANTERIOR_CANDIDATOS[ciclo]
    atual = _MAPA_ATUAL[ciclo]
    vu_col = _MAPA_VU[ciclo]
    atual_rng = _col_range("posicao_contratual", atual)
    vu_rng = _col_range("historico_VU", vu_col)

    # Sem fallback possivel para o checkpoint ATUAL do ciclo (ele define o
    # fim do intervalo observado); se ausente, o ciclo nao foi coletado e a
    # comparacao nao pode ser fabricada. O ANTERIOR tem fallback ate E, entao
    # "sem base" so ocorre no caso teorico de TODOS os candidatos ausentes.
    anterior_todos_ausentes = "*".join(
        f"(1-ISNUMBER({_col_range('posicao_contratual', c)}))" for c in candidatos
    )
    falta_dados = (
        'SUMPRODUCT((posicao_contratual!$A$2:$A$201<>"")*'
        f'(((1-ISNUMBER({atual_rng}))+({anterior_todos_ausentes}))>0))'
    )
    anterior_expr = _anterior_efetivo_expr(candidatos)
    # N() neutraliza celulas "vazias por formula" (IF(...,"",...) devolve
    # texto "", nao branco real) — sem isso, SUMPRODUCT quebra com #VALUE!
    # ao subtrair texto de numero nas linhas alem dos itens reais.
    calculo = (
        f'ROUND(SUMPRODUCT(({anterior_expr}-N({atual_rng}))*N({vu_rng})),2)'
    )
    corpo = f'IF({falta_dados}>0,"NAO COMPARAVEL",{calculo})'
    return _condicionar_financeiro(corpo)


def _formula_diferenca(linha: int) -> str:
    corpo = (
        f'IF(OR(B{linha}="",C{linha}="",C{linha}="NAO COMPARAVEL"),"",'
        f'ROUND(B{linha}-C{linha},2))'
    )
    return _condicionar_financeiro(corpo)


def _formula_status(linha: int) -> str:
    corpo = (
        f'IF(OR(B{linha}="",C{linha}="",C{linha}="NAO COMPARAVEL"),'
        f'"NAO COMPARAVEL",IF(ROUND(D{linha},2)=0,"OK","REVISAR"))'
    )
    return _condicionar_financeiro(corpo, nao_aplicavel="NAO APLICAVEL")


def _aplicar_resultados(wb) -> None:
    res = wb.Worksheets(ABA_RESULTADOS)
    estado, selecao = _capturar_protecao(res)
    try:
        res.Range(f"A{_LINHA_METODOLOGIA}").Value = "7. METODOLOGIA DO VTA"
        res.Range(f"A{_LINHA_METODOLOGIA}").Font.Bold = True
        res.Range(f"A{_LINHA_METODOLOGIA + 1}").Value = _TEXTO_METODOLOGIA_1
        res.Range(f"A{_LINHA_METODOLOGIA + 2}").Formula = _formula_metodologia_execucao()

        res.Range(f"A{_LINHA_CONFERENCIA_TITULO}").Formula = _formula_titulo_conferencia()
        res.Range(f"A{_LINHA_CONFERENCIA_TITULO}").Font.Bold = True

        cabecalho = (
            "Ciclo", "Desembolsado informado",
            "Execucao teorica pelo quantitativo", "Diferenca", "Status",
        )
        for offset, texto in enumerate(cabecalho):
            celula = res.Cells(_LINHA_CONFERENCIA_CABECALHO, 1 + offset)
            celula.Value = texto
            celula.Font.Bold = True

        for indice, ciclo in enumerate(("C0", "C1", "C2", "C3", "C4")):
            linha = _LINHA_CONFERENCIA_C0 + indice
            res.Range(f"A{linha}").Value = ciclo
            res.Range(f"B{linha}").Formula = _formula_desembolsado_ciclo(ciclo.lower())
            res.Range(f"C{linha}").Formula = _formula_execucao_teorica_ciclo(ciclo)
            res.Range(f"D{linha}").Formula = _formula_diferenca(linha)
            res.Range(f"E{linha}").Formula = _formula_status(linha)
            res.Range(f"B{linha}:D{linha}").NumberFormat = MOEDA_BR

        linha_final = _LINHA_CONFERENCIA_C0 + 5
        bloco = res.Range(
            f"A{_LINHA_CONFERENCIA_TITULO}:E{linha_final - 1}"
        )
        bloco.Interior.Color = CORES["azul_muito_claro"]
        _bordas(res.Range(f"A{_LINHA_CONFERENCIA_CABECALHO}:E{linha_final - 1}"))
    finally:
        _restaurar_protecao(res, estado, selecao)


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_vta_fin_"))
    tmp_xlsx = tmp_dir / origem.name
    shutil.copyfile(origem, tmp_xlsx)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
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
        travas_antes = _snapshot_travas(wb)

        _aplicar_memoria(wb)
        _aplicar_resultados(wb)

        travas_depois = _snapshot_travas(wb)
        if travas_antes != travas_depois:
            difs = {
                k: (travas_antes[k], travas_depois[k])
                for k in travas_antes if travas_antes[k] != travas_depois[k]
            }
            raise RuntimeError(f"TRAVA VIOLADA: T21/T22/T23/T25 alterados: {difs}")
        nova_b26 = str(wb.Worksheets(ABA_MEMORIA).Range("B26").Formula)
        if _RAMO_PC_B26 not in nova_b26:
            raise RuntimeError("TRAVA VIOLADA: ramo PC ausente/alterado dentro de B26.")
        if _RAMO_ITENS_B26 not in nova_b26:
            raise RuntimeError("TRAVA VIOLADA: ramo Itens/Consumido ausente/alterado dentro de B26.")

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()

        if aba_ativa in _nomes_abas(wb):
            wb.Worksheets(aba_ativa).Activate()
        wb.Save()
        salvo = True
        wb.Close(SaveChanges=False)
        wb = None

        # Reabre sem reparo para provar zero-corrupcao.
        wb = excel.Workbooks.Open(
            str(tmp_xlsx), UpdateLinks=0, ReadOnly=True, CorruptLoad=0
        )
        abas = _nomes_abas(wb)
        for obrig in (ABA_MEMORIA, ABA_RESULTADOS):
            if obrig not in abas:
                raise RuntimeError(f"Aba {obrig} ausente apos reabertura.")
        if "VTA_FINAL" not in _nomes_definidos(wb):
            raise RuntimeError("Nome VTA_FINAL ausente apos reabertura.")
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


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("origem", type=Path)
    parser.add_argument("destino", type=Path)
    args = parser.parse_args()
    aplicar(args.origem, args.destino)
    print("VTA Financeiro canonico aplicado:", args.destino)


if __name__ == "__main__":
    main()
