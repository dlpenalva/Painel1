# -*- coding: utf-8 -*-
"""VTA-M2/VTA-M2.1/VTA-M2.2 — VTA do metodo Financeiro alinhado a identidade
canonica no XLS.

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
      teorica pelo quantitativo, por ciclo, com status
      OK/REVISAR/NAO COMPARAVEL/NAO APLICAVEL.

VTA-M2.2 — regra temporal da conferencia (correcao sobre a M2.1):
    A M2.1 usava uma cadeia de fallback do checkpoint anterior (REM_BASE
    do ciclo n-1, n-2, ... ate E) para nao perder a comparacao de C3
    quando ciclos intermediarios nao foram coletados. Essa cadeia foi
    identificada como temporalmente invalida: podia comparar uma
    execucao ACUMULADA de varios ciclos (ex.: C0->C3) contra o
    Financeiro de apenas C3, produzindo uma diferenca numericamente
    precisa mas conceitualmente errada.
    A M2.2 elimina qualquer fallback/encadeamento. "Execucao teorica
    pelo quantitativo" de cada ciclo passa a somar, entre os itens, as
    colunas de execucao JA existentes em itens_Remanesc — calculadas
    pelo proprio template com semantica estritamente adjacente (par
    unico de checkpoints, nenhum encadeamento):
        C0: itens_Remanesc!AC (VALOR_EXECUTADO_C0 = MAX(E-J,0)*VU_C0)
        C1: itens_Remanesc!N  (VALOR_EXECUTADO_C1 = MAX(J+H-N,0)*VU_C1)
        C2: itens_Remanesc!P  (VALOR_EXECUTADO_C2 = MAX(N+L-R,0)*VU_C2)
        C3: itens_Remanesc!R  (VALOR_EXECUTADO_C3 = MAX(R+P-V,0)*VU_C3)
    Cada uma dessas celulas ja e "" quando falta qualquer um dos DOIS
    checkpoints adjacentes que ela exige (nenhum fallback embutido). Se
    QUALQUER item real carecer do par, o ciclo inteiro fica NAO
    COMPARAVEL (somar so os itens disponiveis produziria total parcial,
    nao comparavel ao Financeiro que cobre todos os itens). C4 nao tem
    par nesta versao do schema (nao existe REM_BASE_C5 que feche o
    ciclo) — e sempre NAO COMPARAVEL, nunca por fallback, por ausencia
    estrutural de checkpoint de fechamento.
    Nao ha, nesta versao do schema, uma posicao quantitativa "vigente"
    com a MESMA data de corte do Financeiro que permita comparacao
    parcial do ciclo em curso — por isso essa comparacao parcial nao foi
    implementada (fica fora de escopo, conforme instrucao explicita).
    Comparacao de intervalos agregados (ex.: C0->C3 somado) tambem fica
    fora de escopo desta correcao.

Trava anti-regressao: T21/T22/T23/T25, o ramo PC dentro de B26 (a
subexpressao literal do IF($B$4="PCs", ...)) e o ramo Itens/Consumido
(a subexpressao literal do ELSE generico) precisam permanecer
identicos antes/depois, verificado por assinatura de texto. A area de
RESULTADOS linhas 68-77 e validada como vazia (ou ja no formato
VTA-M2/M2.1/M2.2 esperado) antes da escrita, para nao sobrescrever
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


# VTA-M2.2: em vez de reconstruir a execucao por diferenca entre checkpoints
# de posicao_contratual (que pode misturar ciclos nao adjacentes e violar a
# homogeneidade temporal apontada na auditoria), reaproveita-se as colunas
# de execucao JA existentes em itens_Remanesc, calculadas pelo proprio
# template com semantica adjacente-apenas (nenhum fallback embutido):
#   C0: AB=QTD_EXECUTADA_C0=MAX(posicao_contratual!E-posicao_contratual!J,0)
#       AC=VALOR_EXECUTADO_C0=AB*historico_VU!C (VU_C0)
#   C1: M=MAX(posicao_contratual!J+H-N,0); N=M*historico_VU!D (VU_C1)
#   C2: O=MAX(posicao_contratual!N+L-R,0); P=O*historico_VU!E (VU_C2)
#   C3: Q=MAX(posicao_contratual!R+P-V,0); R=Q*historico_VU!F (VU_C3)
# Cada uma dessas formulas ja retorna "" (nao 0) quando falta qualquer um
# dos DOIS checkpoints adjacentes que ela exige — nao ha encadeamento para
# checkpoints mais antigos. C4 nao tem par nesta versao do schema (nao ha
# "C5" que feche o ciclo C4), por isso e sempre NAO COMPARAVEL.
_MAPA_EXECUCAO_VALOR = {"C0": "AC", "C1": "N", "C2": "P", "C3": "R"}


def _col_range(nome_aba: str, coluna: str) -> str:
    return f"{nome_aba}!${coluna}$2:${coluna}$201"


def _formula_execucao_teorica_ciclo(ciclo: str) -> str:
    val_col = _MAPA_EXECUCAO_VALOR.get(ciclo)
    if val_col is None:
        # C4: nao ha checkpoint de fechamento (nao existe REM_BASE_C5) nesta
        # versao do schema — nenhuma base temporalmente valida e possivel.
        return _condicionar_financeiro('"NAO COMPARAVEL"')
    val_rng = _col_range("itens_Remanesc", val_col)
    # Conservador (regra do item 7): se QUALQUER item real nao tiver o par
    # adjacente necessario (a propria celula VALOR_EXECUTADO_Cn ja vem
    # vazia nesse caso), o ciclo inteiro fica NAO COMPARAVEL — somar so os
    # itens com dado disponivel produziria um total parcial nao comparavel
    # ao Financeiro (que cobre todos os itens).
    # O ">0" fica FORA do SUMPRODUCT (compara o escalar resultante), nunca
    # dentro do argumento unico — testado empiricamente contra o caso real:
    # SUMPRODUCT((A<>"")*(1-ISNUMBER(range))>0) nao soma corretamente um
    # array logico derivado de uma multiplicacao anterior (retorna 0 mesmo
    # quando ha itens ausentes); SUMPRODUCT((A<>"")*(1-ISNUMBER(range)))>0,
    # com a comparacao aplicada ao total ja somado, e o padrao correto.
    falta_dados = (
        'SUMPRODUCT((itens_Remanesc!$A$2:$A$201<>"")*'
        f'(1-ISNUMBER({val_rng})))'
    )
    # N() neutraliza celulas "vazias por formula" (IF(...,"",...) devolve
    # texto "", nao branco real) nas linhas alem dos itens reais.
    calculo = f'ROUND(SUM(N({val_rng})),2)'
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
