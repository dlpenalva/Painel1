# -*- coding: utf-8 -*-
"""Coleta 12.4 — precisao financeira canonica + SALDO REMANESCENTE — EVOLUCAO.

PRECISAO (frente 1)
  Somente agregacoes MONETARIAS que ainda nao fechavam em 2 casas recebem
  ROUND(...,2). Os componentes ja eram arredondados item a item; a soma de
  milhares de parcelas de 2 casas deixava residuo binario (39.727.893,699999996)
  e diferencas que deveriam ser zero saiam como -0,00000002 ("-R$ 0,00").
  Nenhum percentual, fator, quantidade ou contagem e tocado; nenhuma regra muda.
  O efeito numerico e < 0,000001 por celula.

    financeiro          C74, E74, F74 (linha TOTAL)
    itens_Consumidos    P2:P200 (CONS_VALOR_TOTAL; O = quantidade, fora)
    itens_PC            O3:S9 e O12:T18 (quadros por ciclo; N/T3:T9 = contagens)
    MEMORIA_RESULTADOS  T21, T22, T23, W53, W54, W66, W67
    RESULTADOS_DETALHE  B85 (ramo PCs com posicao atual)

SALDO REMANESCENTE — EVOLUCAO (frente 2)
  Nova tabela em RESULTADOS!B76:G83, depois das tabelas existentes, so com
  espelhos de fontes canonicas (nenhum calculo financeiro novo):
    saldo da abertura Cn   RESULTADOS_DETALHE!C26:C30, somente se a abertura
                           estiver completa (MEMORIA_RESULTADOS!W41:W45 = 1)
    data da abertura Cn    parametros!I2:I6 (DATA_ABERTURA_FISICA_EXATA)
    ultima posicao         CICLO_EM_EXECUCAO!D5/A9 quando a posicao atual e
                           valida (MEMORIA_RESULTADOS!W49 = 1); senao a ultima
                           abertura completa (MEMORIA_RESULTADOS!W46). A escolha
                           fica em MEMORIA_RESULTADOS!V68:W70 (a executiva nao
                           calcula nem usa INDIRECT) e a RESULTADOS so espelha.
  Ausencia de posicao nunca vira R$ 0,00: aparece "—" e "Não informado".
  Textos acentuados ficam como constantes em MEMORIA_RESULTADOS!AF81:AG83
  (AG80 segue vazio: fim do bloco de orientacao da 12.1).

uso: python tools/aplicar_coleta_124_precisao_saldo.py [caminho.xlsx]
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

import aplicar_resultados_executivo_v3 as v3  # noqa: E402

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
MEM = "MEMORIA_RESULTADOS"
DET = "RESULTADOS_DETALHE"
EXE = "RESULTADOS"

# ------------------------------------------------------------- frente 1: ROUND
ROUND_CELULAS_INTEIRAS: dict[str, tuple[str, ...]] = {
    "financeiro": ("C74", "E74", "F74"),
    "itens_PC": tuple(
        f"{col}{lin}" for lin in range(3, 10) for col in "OPQRS"
    ) + tuple(
        f"{col}{lin}" for lin in range(12, 19) for col in "OPQRST"
    ),
    MEM: ("T22", "W53", "W54"),
}

_SOMA_CICLO_EM_EXECUCAO_G = 'SUM(INDIRECT("CICLO_EM_EXECUCAO!G13:G211"))'
_W67_PCS = (
    "$T$21+SUMPRODUCT((ROW(itens_PC!$O$3:$O$7)-3>=1)*(ROW(itens_PC!$O$3:$O$7)-3<$W$46)"
    "*(itens_PC!$O$3:$O$7+itens_PC!$Q$3:$Q$7))"
)
ROUND_TRECHOS: dict[tuple[str, str], tuple[tuple[str, str], ...]] = {
    (MEM, "T21"): ((",SUM($X$2:$X$201))", ",ROUND(SUM($X$2:$X$201),2))"),),
    (MEM, "T23"): ((",SUM($Y$2:$Y$201))", ",ROUND(SUM($Y$2:$Y$201),2))"),),
    (MEM, "W66"): (
        ("SUMPRODUCT((ROW($W$61:$W$65)-ROW($W$61)<$T$20)*$W$61:$W$65)",
         "ROUND(SUMPRODUCT((ROW($W$61:$W$65)-ROW($W$61)<$T$20)*$W$61:$W$65),2)"),
        ('"",$T$21+$T$22)', '"",ROUND($T$21+$T$22,2))'),
    ),
    (MEM, "W67"): (
        ("SUMPRODUCT((ROW($W$61:$W$65)-ROW($W$61)<$W$46)*$W$61:$W$65)",
         "ROUND(SUMPRODUCT((ROW($W$61:$W$65)-ROW($W$61)<$W$46)*$W$61:$W$65),2)"),
        (_W67_PCS, f"ROUND({_W67_PCS},2)"),
    ),
    (DET, "B85"): (
        (f",{_SOMA_CICLO_EM_EXECUCAO_G},", f",ROUND({_SOMA_CICLO_EM_EXECUCAO_G},2),"),
    ),
}

LINHAS_CONSUMIDOS = range(2, 201)


def formula_consumidos_p(linha: int) -> tuple[str, str]:
    """(forma 12.3, forma 12.4) de itens_Consumidos!P{linha}."""
    soma = f"SUM(F{linha},H{linha},J{linha},L{linha},N{linha})"
    return (f'=IF(A{linha}="","",{soma})', f'=IF(A{linha}="","",ROUND({soma},2))')


def arredondar_formula(atual: str) -> str:
    """=X -> =ROUND(X,2) (idempotente)."""
    if atual.startswith("=ROUND(") and atual.endswith(",2)"):
        return atual
    return "=ROUND(" + atual[1:] + ",2)"


def aplicar_trechos(aba: str, celula: str, atual: str) -> str:
    """Troca literal fail-closed; reaplicar devolve a mesma formula."""
    nova = atual
    for trecho, troca in ROUND_TRECHOS[(aba, celula)]:
        if troca in nova:
            continue
        if nova.count(trecho) != 1:
            raise ValueError(f"{aba}!{celula} fora da forma 12.3 esperada")
        nova = nova.replace(trecho, troca)
    return nova


# ------------------------------------------------------ frente 2: SALDO (RESULTADOS)
TITULO_SALDO = "SALDO REMANESCENTE — EVOLUÇÃO"
CABECALHO_SALDO = ("REFERÊNCIA", "DATA", "SALDO REMANESCENTE ATUALIZADO", "SITUAÇÃO")
ROTULO_ULTIMA = "Última posição informada"
L_SALDO_TITULO = 76
L_SALDO_CAB = 77
L_SALDO_C0 = 78
L_SALDO_ULTIMA = 83
AREA_IMPRESSAO = "$B$1:$G$83"

# Constantes acentuadas (AF = chave ASCII, AG = texto)
CONSTANTES_SALDO = {
    81: ("SALDO_SIT_ABERTURA", "Posição de abertura"),
    82: ("SALDO_SIT_NAO_INFORMADO", "Não informado"),
    83: ("SALDO_SIT_ATUAL", "Posição atual"),
}
TRACO = f"{MEM}!$AG$2"
SIT_ABERTURA = f"{MEM}!$AG$81"
SIT_NAO_INFORMADO = f"{MEM}!$AG$82"
SIT_ATUAL = f"{MEM}!$AG$83"

_CEE_DATA = 'INDIRECT("CICLO_EM_EXECUCAO!$D$5")'
_CEE_TOTAL = 'INDIRECT("CICLO_EM_EXECUCAO!$A$9")'

# A RESULTADOS executiva nao calcula nem usa INDIRECT (contrato do v3): a
# escolha da ultima posicao fica na camada de apoio MEMORIA_RESULTADOS!V68:W70
# e a RESULTADOS so espelha W69 (saldo) e W70 (data). E so selecao entre
# fontes canonicas existentes; nenhum valor financeiro novo e calculado.
APOIO_ULTIMA = {
    68: ("APOIO - SALDO REMANESCENTE: ULTIMA POSICAO INFORMADA (RESULTADOS!B83)", None),
    69: ("Saldo da ultima posicao (CICLO_EM_EXECUCAO!A9 se W49 = 1; senao abertura W46)",
         '=IFERROR(IF($W$49=1,IF(ISNUMBER(' + _CEE_TOTAL + '),' + _CEE_TOTAL + ',""),'
         'IF($W$46="","",IF(ISNUMBER(INDEX(RESULTADOS_DETALHE!$C$26:$C$30,$W$46+1)),'
         'INDEX(RESULTADOS_DETALHE!$C$26:$C$30,$W$46+1),""))),"")'),
    70: ("Data da ultima posicao (CICLO_EM_EXECUCAO!D5 ou parametros!I da abertura W46)",
         '=IFERROR(IF(NOT(ISNUMBER($W$69)),"",IF($W$49=1,IF(ISNUMBER(' + _CEE_DATA + '),'
         + _CEE_DATA + ',""),IF(ISNUMBER(INDEX(parametros!$I$2:$I$6,$W$46+1)),'
         'INDEX(parametros!$I$2:$I$6,$W$46+1),""))),"")'),
}


def formulas_saldo() -> dict[str, str]:
    """Formulas (ASCII) da tabela SALDO REMANESCENTE — EVOLUCAO."""
    f: dict[str, str] = {}
    for n in range(5):
        linha = L_SALDO_C0 + n
        completa = f"{MEM}!$W${41 + n}=1"
        saldo = f"{DET}!$C${26 + n}"
        data = f"parametros!$I${2 + n}"
        f[f"D{linha}"] = f"=IF(AND({completa},ISNUMBER({saldo})),{saldo},{TRACO})"
        f[f"C{linha}"] = f"=IF(AND(ISNUMBER(D{linha}),ISNUMBER({data})),{data},{TRACO})"
        f[f"E{linha}"] = f"=IF(ISNUMBER(D{linha}),{SIT_ABERTURA},{SIT_NAO_INFORMADO})"
    u = L_SALDO_ULTIMA
    f[f"D{u}"] = f"=IF(ISNUMBER({MEM}!$W$69),{MEM}!$W$69,{TRACO})"
    f[f"C{u}"] = f"=IF(AND(ISNUMBER(D{u}),ISNUMBER({MEM}!$W$70)),{MEM}!$W$70,{TRACO})"
    f[f"E{u}"] = f"=IF(ISNUMBER(D{u}),{SIT_ATUAL},{SIT_NAO_INFORMADO})"
    return f


def rotulos_saldo() -> dict[str, str]:
    r = {f"B{L_SALDO_TITULO}": TITULO_SALDO}
    for col, texto in zip("BCDE", CABECALHO_SALDO):
        r[f"{col}{L_SALDO_CAB}"] = texto
    for n in range(5):
        r[f"B{L_SALDO_C0 + n}"] = f"Abertura C{n}"
    r[f"B{L_SALDO_ULTIMA}"] = ROTULO_ULTIMA
    return r


# Estilo: copia atributo a atributo de celulas existentes da propria RESULTADOS
# (nenhum estilo novo; formatacao condicional nao e copiada).
#   titulo   <- B44:G44 (secao 2)          cabecalho <- B45 / C37 / C45 / G37
#   1a linha <- B46 / C38 / C46 / G38      demais    <- B47 / C39 / C47 / G39
#   ultima   <- B52:E52 (linha TOTAL), com data/situacao alinhadas como C38/G38
FONTE_ESTILO: dict[int, dict[str, str]] = {
    L_SALDO_TITULO: {c: f"{c}44" for c in "BCDEFG"},
    L_SALDO_CAB: {"B": "B45", "C": "C37", "D": "C45", "E": "G37"},
    L_SALDO_C0: {"B": "B46", "C": "C38", "D": "C46", "E": "G38"},
}
for _linha in range(L_SALDO_C0 + 1, L_SALDO_ULTIMA):
    FONTE_ESTILO[_linha] = {"B": "B47", "C": "C39", "D": "C47", "E": "G39"}
FONTE_ESTILO[L_SALDO_ULTIMA] = {"B": "B52", "C": "C52", "D": "C52", "E": "C52"}
AJUSTE_ULTIMA = {"C": "C38", "E": "G38"}  # formato/alinhamento da coluna
ALTURA_FONTE = {L_SALDO_TITULO: 44, L_SALDO_CAB: 45}
ALTURA_DADOS = 46


def validar_ascii_e_parenteses() -> None:
    formulas = {f"{EXE}!{c}": v for c, v in formulas_saldo().items()}
    formulas["itens_Consumidos!P2"] = formula_consumidos_p(2)[1]
    for linha, (_, formula) in APOIO_ULTIMA.items():
        if formula:
            formulas[f"{MEM}!W{linha}"] = formula
    v3.validar_ascii_e_parenteses(formulas)


def _gravar(rng, endereco: str, formula: str) -> None:
    """Toda formula nova passa pela regra ZERO CORRUPCAO antes de gravar."""
    v3.validar_ascii_e_parenteses({endereco: formula})
    rng.Formula = formula


# ------------------------------------------------------------- helpers COM
_COM_OCUPADO = -2147418111
XL_NONE = -4142


def _com_seguro(funcao):
    def chamada(*args, **kwargs):
        for _ in range(60):
            try:
                return funcao(*args, **kwargs)
            except Exception as exc:  # pragma: no cover - somente Excel COM
                if getattr(exc, "hresult", None) != _COM_OCUPADO:
                    raise
                time.sleep(0.3)
        return funcao(*args, **kwargs)
    return chamada


@_com_seguro
def _copiar_estilo(src, dst, *, formato_e_alinhamento_de=None) -> None:
    dst.NumberFormat = src.NumberFormat
    for attr in ("Name", "Size", "Bold", "Italic", "Color"):
        setattr(dst.Font, attr, getattr(src.Font, attr))
    if src.Interior.ColorIndex == XL_NONE:
        dst.Interior.ColorIndex = XL_NONE
    else:
        dst.Interior.Pattern = src.Interior.Pattern
        dst.Interior.Color = src.Interior.Color
    for attr in ("HorizontalAlignment", "VerticalAlignment", "WrapText", "IndentLevel"):
        setattr(dst, attr, getattr(src, attr))
    for indice in (7, 8, 9, 10):  # esquerda, topo, base, direita
        origem, destino = src.Borders(indice), dst.Borders(indice)
        if origem.LineStyle == XL_NONE:
            destino.LineStyle = XL_NONE
        else:
            destino.LineStyle = origem.LineStyle
            destino.Weight = origem.Weight
            destino.Color = origem.Color
    if formato_e_alinhamento_de is not None:
        ref = formato_e_alinhamento_de
        dst.NumberFormat = ref.NumberFormat
        for attr in ("HorizontalAlignment", "WrapText", "IndentLevel"):
            setattr(dst, attr, getattr(ref, attr))


def conferir(wb) -> None:
    """Forma 12.3 exigida antes da migracao (fail-closed)."""
    ws = wb.Worksheets(EXE)
    for linha in range(L_SALDO_TITULO, L_SALDO_ULTIMA + 1):
        for col in "BCDEFG":
            if ws.Range(f"{col}{linha}").Formula not in ("", None):
                if ws.Range(f"B{L_SALDO_TITULO}").Value == TITULO_SALDO:
                    return  # ja aplicado: a forma 12.4 e conferida no final
                raise RuntimeError(f"{EXE}!{col}{linha} ocupada; nada aplicado")
    if ws.PageSetup.PrintArea.replace("$", "") != "B1:G74":
        raise RuntimeError("area de impressao da RESULTADOS fora da forma 12.3")
    mem = wb.Worksheets(MEM)
    livres = [f"{c}{l}" for l in range(80, 84) for c in ("AF", "AG")]
    livres += [f"{c}{l}" for l in APOIO_ULTIMA for c in ("V", "W")]
    for celula in livres:
        if mem.Range(celula).Formula not in ("", None):
            raise RuntimeError(f"{MEM}!{celula} ocupada; nada aplicado")


def frente_precisao(wb) -> None:
    for aba, celulas in ROUND_CELULAS_INTEIRAS.items():
        ws = wb.Worksheets(aba)
        for celula in celulas:
            atual = ws.Range(celula).Formula
            if not str(atual).startswith("="):
                raise RuntimeError(f"{aba}!{celula} sem formula; nada aplicado")
            _gravar(ws.Range(celula), f"{aba}!{celula}", arredondar_formula(atual))
    for (aba, celula) in ROUND_TRECHOS:
        rng = wb.Worksheets(aba).Range(celula)
        _gravar(rng, f"{aba}!{celula}", aplicar_trechos(aba, celula, rng.Formula))
    ws = wb.Worksheets("itens_Consumidos")
    for linha in LINHAS_CONSUMIDOS:
        antiga, nova = formula_consumidos_p(linha)
        atual = ws.Range(f"P{linha}").Formula
        if atual == nova:
            continue
        if atual != antiga:
            raise RuntimeError(f"itens_Consumidos!P{linha} fora da forma 12.3")
        _gravar(ws.Range(f"P{linha}"), f"itens_Consumidos!P{linha}", nova)


def frente_saldo(wb) -> None:
    mem = wb.Worksheets(MEM)
    for linha, (chave, texto) in CONSTANTES_SALDO.items():
        mem.Range(f"AF{linha}").Value = chave
        mem.Range(f"AG{linha}").Value = texto
        _copiar_estilo(mem.Range("AF79"), mem.Range(f"AF{linha}"))
        _copiar_estilo(mem.Range("AG79"), mem.Range(f"AG{linha}"))
    for linha, (rotulo, formula) in APOIO_ULTIMA.items():
        _copiar_estilo(mem.Range("V60" if formula is None else "V66"), mem.Range(f"V{linha}"))
        mem.Range(f"V{linha}").Value = rotulo
        if formula:
            _copiar_estilo(mem.Range("W66"), mem.Range(f"W{linha}"))
            _gravar(mem.Range(f"W{linha}"), f"{MEM}!W{linha}", formula)

    ws = wb.Worksheets(EXE)
    for linha, origem in FONTE_ESTILO.items():
        for col, celula in origem.items():
            ajuste = AJUSTE_ULTIMA.get(col) if linha == L_SALDO_ULTIMA else None
            _copiar_estilo(
                ws.Range(celula), ws.Range(f"{col}{linha}"),
                formato_e_alinhamento_de=ws.Range(ajuste) if ajuste else None,
            )
        fonte_altura = ALTURA_FONTE.get(linha, ALTURA_DADOS)
        ws.Rows(linha).RowHeight = ws.Rows(fonte_altura).RowHeight
    for celula, texto in rotulos_saldo().items():
        ws.Range(celula).Value = texto
    for celula, formula in formulas_saldo().items():
        ws.Range(celula).Formula = formula
    ws.PageSetup.PrintArea = AREA_IMPRESSAO


def conferir_resultado(wb) -> None:
    ws = wb.Worksheets(EXE)
    for celula, formula in formulas_saldo().items():
        if ws.Range(celula).Formula != formula:
            raise RuntimeError(f"{EXE}!{celula} nao recebeu a formula 12.4")
    if ws.Range(f"B{L_SALDO_TITULO}").Value != TITULO_SALDO:
        raise RuntimeError("titulo da tabela de saldo ausente")
    mem = wb.Worksheets(MEM)
    if mem.Range("AG80").Formula not in ("", None):
        raise RuntimeError("MEMORIA_RESULTADOS!AG80 deve seguir vazia")
    for linha, (_, formula) in APOIO_ULTIMA.items():
        if formula and mem.Range(f"W{linha}").Formula != formula:
            raise RuntimeError(f"{MEM}!W{linha} nao recebeu a formula 12.4")


def aplicar(caminho: Path) -> None:
    import pythoncom
    import win32com.client as com

    validar_ascii_e_parenteses()
    pythoncom.CoInitialize()
    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        excel.Calculation = v3.XL_CALCULO_MANUAL
        conferir(wb)
        frente_precisao(wb)
        frente_saldo(wb)
        conferir_resultado(wb)
        excel.Calculation = v3.XL_CALCULO_AUTOMATICO
        excel.CalculateFullRebuild()
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
    print("OK: Coleta 12.4 (precisao + saldo remanescente) aplicada em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
