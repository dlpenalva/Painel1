# -*- coding: utf-8 -*-
"""Alinha a camada de formulas do XLSX ao VALOR CONSIDERADO por PC.

Motivacao
---------
`itens_PC!P` (VALOR_ATUALIZADO_TOTAL por ciclo) e a soma de `F = D * E`, com
`E = FATOR_ACUMULADO` historico do ciclo. Essa medida permanece valida como
referencia tecnica (Etapa 26C), mas NAO e a que compoe resultado, retroativo,
VTA e documentos.

A medida que governa e o VALOR HISTORICO CONSIDERADO: base efetivamente
executada mais o retroativo reconhecido a pagar. O PC anterior ao inicio do
efeito financeiro do seu ciclo pertence ao ciclo, usa fator efetivo 1 e
permanece no valor original -- por isso nao pode entrar no VTA multiplicado
pelo fator integral.

Por ciclo, o valor considerado ja existe na propria aba:
``itens_PC!O`` (VALOR_PC_TOTAL) + ``itens_PC!Q`` (RETROATIVO_RECONHECIDO).

Este utilitario troca a fonte de `MEMORIA_RESULTADOS!T21`, `T22` e `W48` de
`P` para `O+Q`, publica o valor considerado por PC em `itens_PC!U` e atualiza
os rotulos correspondentes. `T23`, `T25`, `W50`, `W51`, `W52` e `B26` sao
travados: suas formulas nao podem mudar.

Uso:
    python tools/aplicar_valor_considerado_vta.py <origem.xlsx> <destino.xlsx>
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

ABA_MEMORIA = "MEMORIA_RESULTADOS"
ABA_RESULTADOS = "RESULTADOS"
ABA_ITENS_PC = "itens_PC"

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
XL_CELLTYPE_FORMULAS = -4123
XL_CELLTYPE_CONSTANTS = 2
XL_ERRORS = 16

ULTIMA_LINHA_PC = 5001

# Soma, por ciclo, do VALOR CONSIDERADO (VALOR_PC_TOTAL + RETROATIVO).
_CONSIDERADO_POR_CICLO = "(itens_PC!$O$2:$O$6+itens_PC!$Q$2:$Q$6)"
_FAIXA_CICLOS = "(ROW(itens_PC!$O$2:$O$6)-2"

FORMULA_T21 = (
    "=IF(itens_PC!$O$2>0,ROUND(itens_PC!$O$2+itens_PC!$Q$2,2),"
    "SUM($X$2:$X$201))"
)
FORMULA_T22 = (
    "=SUMPRODUCT(" + _FAIXA_CICLOS + ">=1)*" + _FAIXA_CICLOS + "<$T$20)*"
    + _CONSIDERADO_POR_CICLO + ")"
)
FORMULA_W48 = (
    '=IF($W$46="","",ROUND($T$21+SUMPRODUCT('
    + _FAIXA_CICLOS + ">=1)*" + _FAIXA_CICLOS + "<$W$46)*"
    + _CONSIDERADO_POR_CICLO + ")+$W$53+$W$54,2))"
)
FORMULA_U_PC = '=IF(OR($D{linha}="",$H{linha}=""),"",ROUND($D{linha}+$H{linha},2))'

# C0 executado por item: movimentacao entre as ABERTURAS TEMPORALMENTE
# CORRETAS. A abertura de um ciclo e a quantidade declarada MENOS os deltas
# com data de efeito posterior a sua abertura (DELTA_POSTERIOR_ABERTURA_Cn):
# um aditivo do meio de C1 nao pode retroagir sobre a fotografia de abertura
# de C1 e, por consequencia, encolher a execucao de C0.
FORMULA_X_C0 = (
    '=IF(posicao_contratual!$A2="",0,'
    'IF(AND(ISNUMBER(posicao_contratual!$Y2),posicao_contratual!$Y2>0),0,'
    "IF(AND(ISNUMBER(posicao_contratual!$G2),ISNUMBER(posicao_contratual!$K2),"
    "ISNUMBER(historico_VU!$C2)),"
    "((posicao_contratual!$G2-N(posicao_contratual!$AB2))"
    "-(posicao_contratual!$K2-N(posicao_contratual!$AC2)))*historico_VU!$C2,"
    '"")))'
)
ULTIMA_LINHA_POSICAO = 201
ROTULO_X1 = (
    "VTA-PC aux: C0 fisico por item ((abertura C0 - abertura C1) x VU_C0)"
)

ROTULO_S21 = "C0 executado (valor considerado: PC C0 ou movimentacao fisica x VU_C0)"
ROTULO_S22 = "Execucao PC (ciclos 1..vigente-1, VALOR CONSIDERADO = base + retroativo)"
ROTULO_U1 = "VALOR_CONSIDERADO"
FONTE_F11 = (
    "itens_PC!O+Q (valor considerado) + posicao_contratual (G/K/O/S/W) x historico_VU"
)

# Formulas que NAO podem mudar nesta aplicacao.
TRAVAS = ("T23", "T25", "W50", "W51", "W52", "B26")

# --------------------------------------------------------------------------- #
# Posicao fisica do ciclo em execucao e status global
# --------------------------------------------------------------------------- #
# `MEMORIA_RESULTADOS!W49` vale 1 quando CICLO_EM_EXECUCAO!A9 nao esta vazia --
# e A9 so produz valor quando ha data de posicao, ha itens e nenhum item esta
# em ERRO ou INCOMPLETO. E, portanto, a prova de posicao fisica COMPLETA e
# VALIDA exigida para que ela prevaleca sobre a estimativa por PCs.
COND_FISICO = "MEMORIA_RESULTADOS!$W$49=1"
SOMA_EXECUCAO_FISICA = (
    'ROUND(IFERROR(SUM(INDIRECT("CICLO_EM_EXECUCAO!F13:F211")),0),2)'
)
MARCA_FISICO = "CICLO_EM_EXECUCAO!F13:F211"

FORMULA_H33 = (
    '=IF(OR($B$35="",$B$36="",$B$37="",$B$38="",$B$38<0),"REVISE",'
    "IF(OR(posicao_referencia!$I$2=TRUE," + COND_FISICO + '),"VALIDADO","ESTIMADO"))'
)
FORMULA_A39 = (
    "=IF(" + COND_FISICO + ',"Posicao fisica itemizada informada pelo fiscal '
    '(CICLO_EM_EXECUCAO): execucao e remanescente do ciclo atual sao medidos, '
    'nao estimados.",'
    'IF(AND($B$36<>"",posicao_referencia!$I$2<>TRUE),'
    '"Estimativa: assume execucao registrada pelo metodo oficial ("&$B$5&") '
    'aproximadamente igual ao consumo fisico do ciclo.",""))'
)

# Trava anti-dupla-contagem dos aditivos: o calculo manual so continua exigido
# enquanto NAO houver prova de que a base contratual ja incorporou o aditivo.
# A prova sao as duas travas independentes fechando em zero: a reconciliacao
# das decomposicoes do VTA (W51) e a trava anti-dupla-contagem (W55).
GUARDA_ADITIVOS_ANTIGA = "$B$45>0"
GUARDA_ADITIVOS_NOVA = (
    "AND($B$45>0,NOT(AND(ISNUMBER($W$51),ROUND($W$51,2)=0,"
    "ISNUMBER($W$55),ROUND($W$55,2)=0)))"
)
NOTA_A27 = (
    "Aditivos marcados como considerados exigem prova de que a base contratual "
    "ja os incorpora. A planilha so dispensa o calculo manual quando a "
    "reconciliacao das duas decomposicoes do VTA e a trava anti-dupla-contagem "
    "fecham ambas em 0,00; caso contrario, segue exigindo decisao humana."
)

# --------------------------------------------------------------------------- #
# DATA DE CORTE e TOTAIS CANONICOS DE PCs
# --------------------------------------------------------------------------- #
# CONTROLE!B3 e a data de corte unica do contrato. PC com DATA_PC posterior
# permanece no inventario do arquivo e sai de todo resultado "ate o corte".
# A exclusao exige PROVA de posterioridade: corte ausente ou data ilegivel
# nunca excluem -- por isso o limite cai para 31/12/9999 quando B3 nao e data,
# e os totais sao obtidos SUBTRAINDO os comprovadamente posteriores.
LIMITE_CORTE = "T31"
FORMULA_LIMITE_CORTE = (
    "=IF(ISNUMBER(CONTROLE!$B$3),CONTROLE!$B$3,DATE(9999,12,31))"
)
ROTULO_LIMITE_CORTE = "Limite da data de corte (CONTROLE!B3; sem corte = 31/12/9999)"

_D = "itens_PC!$D$2:$D$5001"
_B = "itens_PC!$B$2:$B$5001"
_C = "itens_PC!$C$2:$C$5001"
_L = "itens_PC!$L$2:$L$5001"
_POSTERIOR = f'{_B},">"&$T$31'


def _posteriores_por_ciclo() -> str:
    return "+".join(
        f'SUMIFS({_D},{_C},"C{n}",{_POSTERIOR})' for n in range(5)
    )


TOTAIS_CANONICOS = {
    # (celula, rotulo, formula)
    "T33": (
        "Total cadastrado de PCs (inventario integral do arquivo)",
        f"=ROUND(SUM({_D}),2)",
    ),
    "T34": (
        "Total considerado ate a data de corte",
        f'=ROUND($T$33-SUMIFS({_D},{_POSTERIOR}),2)',
    ),
    "T35": (
        "Total enquadrado nos ciclos (ate o corte)",
        "=ROUND(SUM(itens_PC!$O$2:$O$6)-(" + _posteriores_por_ciclo() + "),2)",
    ),
    "T36": (
        "Total com efeito financeiro (ate o corte)",
        f'=ROUND(SUMIFS({_D},{_L},"Sim")-SUMIFS({_D},{_L},"Sim",{_POSTERIOR}),2)',
    ),
    "T37": (
        "Total sem efeito financeiro (ate o corte)",
        "=ROUND($T$35-$T$36,2)",
    ),
}

# Retroativo por ciclo (metodo PC) tambem passa a respeitar a data de corte.
FORMULA_RETRO_CICLO = (
    '=IF(COUNTIFS({c},"{ciclo}",{d},">0")=0,"",'
    'ROUND(SUMIFS(itens_PC!$H$2:$H$5001,{c},"{ciclo}",'
    'itens_PC!$G$2:$G$5001,"Sim",{b},"<="&$T$31),2))'
)

# RESULTADOS: base considerada por ciclo (ramo PCs) filtrada pelo corte.
FILTRO_PC_ANTIGO = 'itens_PC!$G$2:$G$5001,"Sim"'
FILTRO_PC_NOVO = (
    'itens_PC!$G$2:$G$5001,"Sim",itens_PC!$B$2:$B$5001,'
    '"<="&MEMORIA_RESULTADOS!$T$31'
)

# Secao 5 de RESULTADOS: as doze medidas, com nomes claros.
LINHA_SECAO_TOTAIS = 53
MEDIDAS_CANONICAS = [
    ("Total cadastrado de PCs", "=IF($B$5<>\"PCs\",\"\",MEMORIA_RESULTADOS!$T$33)",
     "itens_PC!D (inventario integral, sem corte)"),
    ("Total considerado ate a data de corte",
     "=IF($B$5<>\"PCs\",\"\",MEMORIA_RESULTADOS!$T$34)",
     "MEMORIA!T34 = T33 - PCs com DATA_PC > CONTROLE!B3"),
    ("Total enquadrado nos ciclos",
     "=IF($B$5<>\"PCs\",\"\",MEMORIA_RESULTADOS!$T$35)",
     "itens_PC!O (por ciclo) - posteriores ao corte"),
    ("Total com efeito financeiro",
     "=IF($B$5<>\"PCs\",\"\",MEMORIA_RESULTADOS!$T$36)",
     "itens_PC!L=Sim, ate o corte"),
    ("Total sem efeito financeiro",
     "=IF($B$5<>\"PCs\",\"\",MEMORIA_RESULTADOS!$T$37)",
     "enquadrado - com efeito (competencias anteriores ao inicio do efeito)"),
    ("Retroativo reconhecido", "=$D$22", "MEMORIA!B16 (retroativo oficial)"),
    ("Execucao fisica do ciclo atual", "=$B$36",
     "CICLO_EM_EXECUCAO!F13:F211 quando a posicao fisica esta completa"),
    ("Remanescente fisico atual", "=$B$38",
     "abertura do ciclo vigente - execucao fisica"),
    ("VTA oficial", "=$B$10", "MEMORIA!W50 (FORMA 1 - posicao atual)"),
    ("Referencias auditaveis", '="Linhas 10 a 12 desta aba"',
     "posicao atual / ultima abertura / contrato integralmente reajustado"),
    ("Reconciliacao", "=$B$13", "MEMORIA!W51 (FORMA 1 - FORMA 2); 0,00 reconcilia"),
    ("Status", "=$B$3", "STATUS GLOBAL desta aba"),
]
TITULO_SECAO_TOTAIS = "5. TOTAIS CANONICOS DE PCs — MEDIDAS COM NOMES CLAROS"


def _nomes_abas(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


def _validar_origem(wb) -> None:
    abas = _nomes_abas(wb)
    for obrig in (ABA_MEMORIA, ABA_RESULTADOS, ABA_ITENS_PC):
        if obrig not in abas:
            raise RuntimeError(f"Aba {obrig} ausente na origem.")
    mem = wb.Worksheets(ABA_MEMORIA)
    for endereco in ("T21", "T22", "T23", "T25", "W48", "W50"):
        if not str(mem.Range(endereco).Formula or ""):
            raise RuntimeError(
                f"{ABA_MEMORIA}!{endereco} vazio: origem incompativel."
            )
    if str(wb.Worksheets(ABA_ITENS_PC).Range("O1").Value or "") != "VALOR_PC_TOTAL":
        raise RuntimeError("itens_PC!O1 nao e VALOR_PC_TOTAL; layout inesperado.")


def _snapshot_travas(wb) -> dict[str, str]:
    mem = wb.Worksheets(ABA_MEMORIA)
    return {e: str(mem.Range(e).Formula) for e in TRAVAS}


def _aplicar_memoria(wb) -> None:
    mem = wb.Worksheets(ABA_MEMORIA)
    mem.Range("T21").Formula = FORMULA_T21
    mem.Range("T22").Formula = FORMULA_T22
    mem.Range("W48").Formula = FORMULA_W48
    mem.Range("S21").Value = ROTULO_S21
    mem.Range("S22").Value = ROTULO_S22
    mem.Range("X1").Value = ROTULO_X1
    mem.Range(f"X2:X{ULTIMA_LINHA_POSICAO}").Formula = FORMULA_X_C0


def _aplicar_itens_pc(wb) -> None:
    ws = wb.Worksheets(ABA_ITENS_PC)
    ws.Range("U1").Value = ROTULO_U1
    ws.Range(f"U2:U{ULTIMA_LINHA_PC}").Formula = FORMULA_U_PC.format(linha=2)


def _aplicar_resultados(wb) -> None:
    ws = wb.Worksheets(ABA_RESULTADOS)
    ws.Range("F11").Value = FONTE_F11

    # Ciclo atual: a execucao fisica itemizada PREVALECE sobre a estimativa
    # pelo metodo oficial. A formula estimativa e preservada integralmente como
    # ramo alternativo (fallback homologado quando nao ha posicao fisica).
    atual = str(ws.Range("B36").Formula or "")
    if MARCA_FISICO not in atual:
        if not atual.startswith("="):
            raise RuntimeError("RESULTADOS!B36 nao e formula; origem inesperada.")
        ws.Range("B36").Formula = (
            "=IF(" + COND_FISICO + "," + SOMA_EXECUCAO_FISICA
            + ",(" + atual[1:] + "))"
        )
    ws.Range("H33").Formula = FORMULA_H33
    ws.Range("A39").Formula = FORMULA_A39


def _aplicar_corte(wb) -> None:
    """Aplica CONTROLE!B3 aos resultados 'ate a data de corte'."""
    mem = wb.Worksheets(ABA_MEMORIA)
    mem.Range(f"S{LIMITE_CORTE[1:]}").Value = ROTULO_LIMITE_CORTE
    mem.Range(LIMITE_CORTE).Formula = FORMULA_LIMITE_CORTE

    for n in range(5):
        mem.Range(f"C{10 + n}").Formula = FORMULA_RETRO_CICLO.format(
            c=_C, d=_D, b=_B, ciclo=f"C{n}"
        )

    for celula, (rotulo, formula) in TOTAIS_CANONICOS.items():
        mem.Range("S" + celula[1:]).Value = rotulo
        mem.Range(celula).Formula = formula

    res = wb.Worksheets(ABA_RESULTADOS)
    for linha in range(16, 21):
        atual = str(res.Range(f"B{linha}").Formula or "")
        if FILTRO_PC_NOVO in atual:
            continue
        if FILTRO_PC_ANTIGO not in atual:
            raise RuntimeError(
                f"{ABA_RESULTADOS}!B{linha} sem o filtro de PC esperado."
            )
        res.Range(f"B{linha}").Formula = atual.replace(
            FILTRO_PC_ANTIGO, FILTRO_PC_NOVO
        )


def _aplicar_totais_canonicos(wb) -> None:
    """Publica as doze medidas com nomes claros, sem criar aba ou painel."""
    res = wb.Worksheets(ABA_RESULTADOS)
    linha = LINHA_SECAO_TOTAIS
    res.Range(f"A{linha}").Value = TITULO_SECAO_TOTAIS
    res.Range(f"A{linha}").Font.Bold = True
    res.Range(f"A{linha + 1}").Value = "Medida"
    res.Range(f"B{linha + 1}").Value = "Valor"
    res.Range(f"C{linha + 1}").Value = "Composicao auditavel"
    res.Range(f"A{linha + 1}:C{linha + 1}").Font.Bold = True
    for i, (rotulo, formula, fonte) in enumerate(MEDIDAS_CANONICAS, start=2):
        alvo = linha + i
        res.Range(f"A{alvo}").Value = f"{i - 1}. {rotulo}"
        res.Range(f"B{alvo}").Formula = formula
        res.Range(f"C{alvo}").Value = fonte


def _aplicar_guarda_aditivos(wb) -> None:
    mem = wb.Worksheets(ABA_MEMORIA)
    formula = str(mem.Range("E26").Formula or "")
    if GUARDA_ADITIVOS_NOVA in formula:
        return
    if GUARDA_ADITIVOS_ANTIGA not in formula:
        raise RuntimeError(
            f"{ABA_MEMORIA}!E26 sem a guarda {GUARDA_ADITIVOS_ANTIGA!r}; "
            "origem incompativel."
        )
    mem.Range("E26").Formula = formula.replace(
        GUARDA_ADITIVOS_ANTIGA, GUARDA_ADITIVOS_NOVA, 1
    )
    mem.Range("A27").Value = NOTA_A27


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


def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_valor_considerado_"))
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
        _aplicar_itens_pc(wb)
        _aplicar_resultados(wb)
        _aplicar_corte(wb)
        _aplicar_totais_canonicos(wb)
        _aplicar_guarda_aditivos(wb)

        travas_depois = _snapshot_travas(wb)
        if travas_antes != travas_depois:
            difs = {k: (travas_antes[k], travas_depois[k])
                    for k in travas_antes if travas_antes[k] != travas_depois[k]}
            raise RuntimeError(f"TRAVA VIOLADA: {difs}")

        excel.Calculation = XL_CALC_AUTOMATIC
        excel.CalculateFullRebuild()
        _verificar_sem_erros(wb)

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
        for obrig in (ABA_MEMORIA, ABA_RESULTADOS, ABA_ITENS_PC):
            if obrig not in abas:
                raise RuntimeError(f"Aba {obrig} ausente apos reabertura.")
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
    print("Valor considerado aplicado ao VTA:", args.destino)


if __name__ == "__main__":
    main()
