# -*- coding: utf-8 -*-
"""Cabecalhos de itens_Remanesc: fotografia quantitativa na DATA EXATA.

APRESENTACAO PURA. As 16 formulas de itens_Remanesc!E1:T1 continuam ancoradas
em parametros!I (DATA_ABERTURA_FISICA_EXATA); nada da cadeia de calculo muda.

  ANTES   QTD_REM_BASE_C1
          Inicio: 15/04/2026

  DEPOIS  QTD. REMANESCENTE - C1
          Informe a quantidade existente em 15/04/2026

Rotulos separam entrada manual de coluna automatica: so as colunas de QTD
REMANESCENTE (E/G/I/K) sao digitadas pelo fiscal e por isso recebem a data;
as demais declaram-se calculadas. Sem data definida para o ciclo, o cabecalho
fica apenas com o rotulo — nunca com data projetada.

parametros!I permanece intacto (fronteira de aditivos de posicao_contratual!
AB:AF, posicao_referencia!I6 e cobertura_temporal!B6 nao sao tocadas), assim
como parametros!H, que segue exclusivo da competencia do efeito financeiro.

REGRA ZERO CORRUPCAO XLSX: formulas em ingles maiusculo, separador virgula,
ASCII puro (o travessao do texto aprovado vira hifen) e parenteses balanceados.
Aplicacao por Excel COM — openpyxl remove a formatacao condicional x14.

Uso:  python tools/aplicar_cabecalhos_itens_remanesc_fotografia.py [caminho.xlsx]
"""
from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

# ciclo -> linha de parametros (C1=I3 ... C4=I6)
CICLOS = {"C1": 3, "C2": 4, "C3": 5, "C4": 6}

# celula -> (ciclo, rotulo, segunda linha; None = usa a data exata)
CABECALHOS = {
    "E1": ("C1", "QTD. REMANESCENTE - C1", None),
    "F1": ("C1", "VALOR REMANESCENTE - C1", "Calculado automaticamente"),
    "G1": ("C2", "QTD. REMANESCENTE - C2", None),
    "H1": ("C2", "VALOR REMANESCENTE - C2", "Calculado automaticamente"),
    "I1": ("C3", "QTD. REMANESCENTE - C3", None),
    "J1": ("C3", "VALOR REMANESCENTE - C3", "Calculado automaticamente"),
    "K1": ("C4", "QTD. REMANESCENTE - C4", None),
    "L1": ("C4", "VALOR REMANESCENTE - C4", "Calculado automaticamente"),
    "M1": ("C1", "QTD. EXECUTADA - C1", "Calculada automaticamente"),
    "N1": ("C1", "VALOR EXECUTADO - C1", "Calculado automaticamente"),
    "O1": ("C2", "QTD. EXECUTADA - C2", "Calculada automaticamente"),
    "P1": ("C2", "VALOR EXECUTADO - C2", "Calculado automaticamente"),
    "Q1": ("C3", "QTD. EXECUTADA - C3", "Calculada automaticamente"),
    "R1": ("C3", "VALOR EXECUTADO - C3", "Calculado automaticamente"),
    "S1": ("C4", "QTD. EXECUTADA - C4", "Calculada automaticamente"),
    "T1": ("C4", "VALOR EXECUTADO - C4", "Calculado automaticamente"),
}

ORIENTACAO = "Informe a quantidade existente em "


def formula(celula: str) -> str:
    """Formula do cabecalho. Coluna manual traz a data exata da fotografia."""
    ciclo, rotulo, fixa = CABECALHOS[celula]
    if fixa is not None:
        return f'="{rotulo}"&CHAR(10)&"{fixa}"'
    ref = f"parametros!$I${CICLOS[ciclo]}"
    data = (
        f'RIGHT("0"&DAY({ref}),2)&"/"'
        f'&RIGHT("0"&MONTH({ref}),2)&"/"'
        f'&YEAR({ref})'
    )
    return (
        f'=IF({ref}="","{rotulo}",'
        f'"{rotulo}"&CHAR(10)&"{ORIENTACAO}"&{data})'
    )


def validar(texto: str) -> None:
    """Guardas da regra ZERO CORRUPCAO, antes de qualquer escrita."""
    if any(ord(c) > 127 for c in texto):
        raise ValueError(f"formula com caractere nao-ASCII: {texto!r}")
    if texto.count("(") != texto.count(")"):
        raise ValueError(f"parenteses desbalanceados: {texto!r}")
    if ";" in texto:
        raise ValueError(f"separador ponto e virgula: {texto!r}")


def aplicar(caminho: Path) -> None:
    import win32com.client as com

    formulas = {celula: formula(celula) for celula in CABECALHOS}
    for texto in formulas.values():
        validar(texto)

    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(caminho))
        try:
            ws = wb.Worksheets("itens_Remanesc")
            for celula, texto in formulas.items():
                ws.Range(celula).Formula = texto
            wb.Save()
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
    print(f"16 cabecalhos aplicados em {caminho}")


if __name__ == "__main__":
    aplicar(Path(sys.argv[1]) if len(sys.argv) > 1 else TEMPLATE)
