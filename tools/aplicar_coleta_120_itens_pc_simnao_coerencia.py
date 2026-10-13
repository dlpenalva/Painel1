# -*- coding: utf-8 -*-
"""Coleta 12.0 — itens_PC com filtro, PC_PAGO normalizado e coerencia do metodo.

Aplica no template oficial (ou numa copia de Coleta gerada), via Excel COM
(openpyxl destroi a CF x14), somente:

* itens_PC — AutoFilter A1:L5001; AD (PC_PAGO_NORMALIZADO, oculta) com o valor
  de G normalizado; H/I/J/K e o quadro N12:Q17 passam a ler AD; avisos M31
  (grafia diferente reconhecida) e M32 (valor invalido); destaque em G.
* MEMORIA_RESULTADOS — textos AG43:AG64; S54:T65 (contagens das fontes,
  coerencia metodo x fonte, mensagens e contagens de grafia); C10:C14 em AD;
  B26 e T40 bloqueados quando a fonte do metodo esta vazia.
* RESULTADOS_DETALHE — B16:B20 em AD; D22, H14, B3, J5 e B7 com o bloqueio
  de coerencia e a causa na lista de pendencias.
* CONTROLE — C1 com o aviso curto ao lado do metodo.

As formulas vem de `_coleta_oficial` (fonte unica). Cada celula precisa estar
na forma 11.9 ou ja na 12.0; qualquer outra coisa aborta sem gravar.
REGRA ZERO CORRUPCAO XLSX: formulas em ingles, ASCII e parenteses balanceados.

uso: python tools/aplicar_coleta_120_itens_pc_simnao_coerencia.py [caminho.xlsx]
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

import _coleta_oficial as co  # noqa: E402
import aplicar_ajustes_xls_ux_pos174 as ux  # noqa: E402
from aplicar_26h_template import _PROT_FLAGS  # noqa: E402

# Ordem posicional de Worksheet.Protect apos (Password, DrawingObjects,
# Contents, Scenarios, UserInterfaceOnly). Argumentos NOMEADOS sao ignorados em
# silencio pelo COM late-bound (as permissoes voltariam todas a False).
_ORDEM_PROTECT = (
    "AllowFormattingCells", "AllowFormattingColumns", "AllowFormattingRows",
    "AllowInsertingColumns", "AllowInsertingRows", "AllowInsertingHyperlinks",
    "AllowDeletingColumns", "AllowDeletingRows", "AllowSorting",
    "AllowFiltering", "AllowUsingPivotTables",
)
assert set(_ORDEM_PROTECT) == set(_PROT_FLAGS)


def _capturar_protecao(ws):
    """Estado da protecao da aba (None se desprotegida); desprotege a aba."""
    if not bool(ws.ProtectContents):
        return None, None
    protecao = ws.Protection
    estado = {flag: bool(getattr(protecao, flag)) for flag in _ORDEM_PROTECT}
    estado["_objetos"] = bool(ws.ProtectDrawingObjects)
    estado["_cenarios"] = bool(ws.ProtectScenarios)
    selecao = ws.EnableSelection
    ws.Unprotect()
    return estado, selecao


def _restaurar_protecao(ws, estado, selecao) -> None:
    if estado is None:
        return
    ws.Protect(
        "", estado["_objetos"], True, estado["_cenarios"], False,
        *(estado[flag] for flag in _ORDEM_PROTECT),
    )
    for flag in _ORDEM_PROTECT:
        if bool(getattr(ws.Protection, flag)) != estado[flag]:
            raise RuntimeError(f"{ws.Name}: protecao restaurada sem {flag}; nada gravado")
    ws.EnableSelection = selecao

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

XL_CALCULO_MANUAL, XL_CALCULO_AUTOMATICO = -4135, -4105
XL_SEM_PREENCHIMENTO = -4142
PRIMEIRA_PC, ULTIMA_PC = 2, 5001
AD = co.COLUNA_PC_PAGO_NORMALIZADO
COR_ALERTA = "C00000"
COR_AVISO = "9C5700"
AVISOS_PC = (
    ("M31:T31", co._FORMULA_AVISO_GRAFIA_PC, COR_AVISO),
    ("M32:T32", co._FORMULA_AVISO_INVALIDO_PC, COR_ALERTA),
)


def validar_formula(formula: str) -> None:
    if not formula.startswith("="):
        raise ValueError("formula sem '='")
    if formula.count("(") != formula.count(")"):
        raise ValueError("formula com parenteses desbalanceados")
    if any(ord(ch) > 127 for ch in formula):
        raise ValueError("formula com caractere nao ASCII")


def _com(formula: str) -> str:
    """Forma lida/gravada pelo COM: sem o prefixo _xlfn. do arquivo."""
    return formula.replace("_xlfn.", "")


def _copiar_formato(wb, origem, destino) -> None:
    """Formato da celula de origem no destino, sem area de transferencia."""
    o = origem.Cells(1, 1)
    destino.NumberFormat = o.NumberFormat
    destino.Font.Name = o.Font.Name
    destino.Font.Size = o.Font.Size
    destino.Font.Bold = o.Font.Bold
    destino.Font.Italic = o.Font.Italic
    destino.Font.Color = o.Font.Color
    if o.Interior.Pattern == XL_SEM_PREENCHIMENTO:
        destino.Interior.Pattern = XL_SEM_PREENCHIMENTO
    else:
        destino.Interior.Pattern = o.Interior.Pattern
        destino.Interior.Color = o.Interior.Color
    destino.HorizontalAlignment = o.HorizontalAlignment
    destino.VerticalAlignment = o.VerticalAlignment
    destino.WrapText = o.WrapText
    destino.Locked = o.Locked


def _reescrever_coluna(ws, coluna: str, anterior, nova) -> None:
    faixa = ws.Range(f"{coluna}{PRIMEIRA_PC}:{coluna}{ULTIMA_PC}")
    novas = []
    for deslocamento, (atual,) in enumerate(faixa.Formula):
        linha = PRIMEIRA_PC + deslocamento
        formula = nova(linha)
        validar_formula(formula)
        if atual not in (_com(anterior(linha)), _com(formula)):
            raise RuntimeError(f"itens_PC!{coluna}{linha} fora do formato esperado; nada aplicado")
        novas.append((_com(formula),))
    faixa.Formula = tuple(novas)


def _formas_pago_sim(ws, faixa: str) -> dict[str, str]:
    novas = {}
    for celula in ws.Range(faixa).Cells:
        atual = str(celula.Formula)
        if co._TROCA_PC_PAGO_SIM in atual and co._TRECHO_PC_PAGO_SIM not in atual:
            continue  # ja na 12.0
        try:
            nova = co._formula_pc_pago_sim_normalizado(atual)
        except ValueError:
            raise RuntimeError(
                f"{ws.Name}!{celula.Address} sem o criterio PC_PAGO = Sim; nada aplicado"
            ) from None
        validar_formula(nova)
        novas[celula.Address.replace("$", "")] = nova
    return novas


def _formas_coerencia(wb, aba: str) -> dict[str, str]:
    novas = {}
    for (aba_alvo, celula) in co._TRECHOS_COERENCIA_120:
        if aba_alvo != aba:
            continue
        atual = str(wb.Worksheets(aba).Range(celula).Formula)
        try:
            novas[celula] = co._formula_coerencia_120(aba, celula, atual)
        except ValueError:
            raise RuntimeError(f"{aba}!{celula} fora do formato 11.9; nada aplicado") from None
    return novas


def _remover_regras_g(faixa) -> None:
    # Reaplicacao idempotente. O COM devolve Formula1 relativo a celula ativa,
    # entao a regra e reconhecida pelo marcador exclusivo "INVALIDO" (as regras
    # de linha de A2:L5001 nao o contem e ficam intactas).
    marcador = f'"{co.PC_PAGO_INVALIDO}"'
    for i in range(faixa.FormatConditions.Count, 0, -1):
        if marcador in str(faixa.FormatConditions(i).Formula1):
            faixa.FormatConditions(i).Delete()


def _regra_g(faixa, formula: str, cor: str) -> None:
    validar_formula("=" + formula)
    local = ux._local(faixa, "=" + formula)
    regra = faixa.FormatConditions.Add(ux.XL_EXPRESSION, None, local)
    regra.StopIfTrue = False
    regra.Interior.Color = ux._bgr(cor)
    regra.SetFirstPriority()


def conferir(wb) -> None:
    """Confere todas as frentes antes de gravar qualquer celula."""
    pc = wb.Worksheets("itens_PC")
    if pc.Range(f"{AD}1").Formula not in ("", co.CABECALHO_PC_PAGO_NORMALIZADO):
        raise RuntimeError(f"itens_PC!{AD}1 ocupada; nada aplicado")
    for (atual,) in pc.Range(f"{AD}{PRIMEIRA_PC}:{AD}{ULTIMA_PC}").Formula:
        if atual != "" and not str(atual).startswith("=IFERROR(IF(UPPER(TRIM("):
            raise RuntimeError(f"itens_PC!{AD} ocupada; nada aplicado")
    for endereco, formula, _ in AVISOS_PC:
        atual = pc.Range(endereco).Cells(1, 1).Formula
        if atual not in ("", _com(formula)):
            raise RuntimeError(f"itens_PC!{endereco} ocupada; nada aplicado")
    if pc.AutoFilterMode and (
        pc.AutoFilter.Range.Address.replace("$", "") != co.FAIXA_FILTRO_ITENS_PC
    ):
        raise RuntimeError("itens_PC com AutoFilter em outra faixa; nada aplicado")
    mem = wb.Worksheets("MEMORIA_RESULTADOS")
    for celula, texto in co._TEXTOS_COLETA_120.items():
        if mem.Range(celula).Formula not in ("", texto):
            raise RuntimeError(f"MEMORIA_RESULTADOS!{celula} nao esta livre; nada aplicado")
    for celula, (rotulo, formula) in co._CELULAS_COLETA_120.items():
        validar_formula(formula)
        linha = celula[1:]
        if mem.Range(f"S{linha}").Formula not in ("", rotulo) or mem.Range(celula).Formula not in (
            "", _com(formula),
        ):
            raise RuntimeError(f"MEMORIA_RESULTADOS!S{linha}:{celula} nao esta livre; nada aplicado")
    if wb.Worksheets("CONTROLE").Range("C1").Formula not in ("", co._FORMULA_AVISO_METODO_CONTROLE):
        raise RuntimeError("CONTROLE!C1 nao esta livre; nada aplicado")
    for aba, faixa in co._FAIXAS_PC_PAGO_SIM.items():
        _formas_pago_sim(wb.Worksheets(aba), faixa)
    _formas_coerencia(wb, "MEMORIA_RESULTADOS")
    _formas_coerencia(wb, "RESULTADOS_DETALHE")


def frente_itens_pc(wb) -> None:
    ws = wb.Worksheets("itens_PC")
    estado, selecao = _capturar_protecao(ws)
    # AD: valor normalizado (antes de H/I/J/K, que passam a le-lo).
    cabecalho = ws.Range(f"{AD}1")
    faixa_ad = ws.Range(f"{AD}{PRIMEIRA_PC}:{AD}{ULTIMA_PC}")
    faixa_ad.Formula = tuple(
        (_com(co._formula_pc_pago_normalizado(linha)),)
        for linha in range(PRIMEIRA_PC, ULTIMA_PC + 1)
    )
    cabecalho.Value = co.CABECALHO_PC_PAGO_NORMALIZADO
    _copiar_formato(wb, ws.Range("AC1"), cabecalho)
    _copiar_formato(wb, ws.Range(f"AC{PRIMEIRA_PC}"), faixa_ad)
    ws.Columns(AD).ColumnWidth = ws.Columns("AC").ColumnWidth
    ws.Columns(AD).Hidden = True

    _reescrever_coluna(ws, "H", co._formula_retro_pc_11_8, co._formula_retro_pc)
    _reescrever_coluna(ws, "I", co._formula_em_analise_pc_11_8, co._formula_em_analise_pc)
    _reescrever_coluna(ws, "J", co._formula_potencial_pc_11_8, co._formula_potencial_pc)
    _reescrever_coluna(ws, "K", co._formula_check_pc_11_8, co._formula_check_pc)
    for celula, nova in _formas_pago_sim(ws, co._FAIXAS_PC_PAGO_SIM["itens_PC"]).items():
        ws.Range(celula).Formula = nova

    # Avisos abaixo da legenda "COMO OS PCs SAO TRATADOS" (M20:T29).
    for endereco, formula, cor in AVISOS_PC:
        faixa = ws.Range(endereco)
        if not faixa.MergeCells:
            faixa.Merge()
        faixa.Cells(1, 1).Formula = formula
        faixa.WrapText = True
        faixa.VerticalAlignment = ux.XL_VCENTER
        faixa.Font.Bold = True
        faixa.Font.Color = ux._bgr(cor)
        ws.Rows(faixa.Row).RowHeight = max(ws.Rows(faixa.Row).RowHeight, 30)

    coluna_g = ws.Range(f"G{PRIMEIRA_PC}:G{ULTIMA_PC}")
    _remover_regras_g(coluna_g)
    _regra_g(coluna_g, co._FORMULA_DESTAQUE_INVALIDO_PC, co._COR_DESTAQUE_INVALIDO_PC)
    _regra_g(coluna_g, co._FORMULA_DESTAQUE_GRAFIA_PC, co._COR_DESTAQUE_GRAFIA_PC)

    if not ws.AutoFilterMode:
        ws.Range(co.FAIXA_FILTRO_ITENS_PC).AutoFilter()
    if estado is not None:
        estado["AllowFiltering"] = True
    _restaurar_protecao(ws, estado, selecao)


def frente_memoria(wb) -> None:
    mem = wb.Worksheets("MEMORIA_RESULTADOS")
    novas = _formas_coerencia(wb, "MEMORIA_RESULTADOS")
    novas.update(_formas_pago_sim(mem, co._FAIXAS_PC_PAGO_SIM["MEMORIA_RESULTADOS"]))
    for celula, texto in co._TEXTOS_COLETA_120.items():
        mem.Range(celula).Value = texto
        _copiar_formato(wb, mem.Range("AG42"), mem.Range(celula))
    for celula, (rotulo, formula) in co._CELULAS_COLETA_120.items():
        linha = celula[1:]
        mem.Range(f"S{linha}").Value = rotulo
        _copiar_formato(wb, mem.Range("S48"), mem.Range(f"S{linha}"))
        _copiar_formato(wb, mem.Range("T48"), mem.Range(celula))
        mem.Range(celula).Formula = _com(formula)
    for celula, nova in novas.items():
        mem.Range(celula).Formula = nova


def frente_detalhe(wb) -> None:
    ws = wb.Worksheets("RESULTADOS_DETALHE")
    estado, selecao = _capturar_protecao(ws)
    novas = _formas_coerencia(wb, "RESULTADOS_DETALHE")
    novas.update(_formas_pago_sim(ws, co._FAIXAS_PC_PAGO_SIM["RESULTADOS_DETALHE"]))
    for celula, nova in novas.items():
        ws.Range(celula).Formula = nova
    _restaurar_protecao(ws, estado, selecao)


def frente_controle(wb) -> None:
    ws = wb.Worksheets("CONTROLE")
    estado, selecao = _capturar_protecao(ws)
    alvo = ws.Range("C1")
    alvo.Formula = co._FORMULA_AVISO_METODO_CONTROLE
    alvo.Font.Bold = True
    alvo.Font.Color = ux._bgr(COR_ALERTA)
    alvo.WrapText = False
    alvo.HorizontalAlignment = ux.XL_LEFT
    local = ux._local(alvo, '=LEFT($C$1,8)="CONFERIR"')
    for i in range(alvo.FormatConditions.Count, 0, -1):
        if alvo.FormatConditions(i).Formula1 == local:
            alvo.FormatConditions(i).Delete()
    regra = alvo.FormatConditions.Add(ux.XL_EXPRESSION, None, local)
    regra.StopIfTrue = False
    regra.Font.Color = ux._bgr(COR_AVISO)
    _restaurar_protecao(ws, estado, selecao)


def aplicar(caminho: Path) -> None:
    import pythoncom
    import win32com.client as com

    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        excel.Calculation = XL_CALCULO_MANUAL
        conferir(wb)
        frente_itens_pc(wb)
        frente_memoria(wb)
        frente_detalhe(wb)
        frente_controle(wb)
        excel.Calculation = XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()
        wb.Worksheets("CONTROLE").Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        wb.Save()
        ux._fechar(wb)
        wb = None
    finally:
        if wb is not None:
            ux._fechar(wb)
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
    caminho = Path(sys.argv[1] if len(sys.argv) > 1 else TEMPLATE).resolve()
    if not caminho.exists():
        print("ERRO: arquivo nao encontrado: %s" % caminho)
        return 1
    aplicar(caminho)
    print("OK: Coleta 12.0 aplicada em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
