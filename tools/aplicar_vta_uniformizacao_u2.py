# -*- coding: utf-8 -*-
"""VTA-U2 — uniformizacao natural do VTA na aba RESULTADOS.

Aplica, via Excel COM em copia temporaria (mesmo padrao zero-corrupcao de
`aplicar_vta_financeiro_canonico.py`), quatro mudancas de APRESENTACAO na aba
RESULTADOS. Nenhuma formula de negocio do VTA e alterada: MEMORIA_RESULTADOS
permanece byte-a-byte identica (travada e conferida no fim).

1) Card "VTA OFICIAL" (C5) passa a exibir a saida canonica calculada do metodo
   (nome definido VTA_FINAL = MEMORIA_RESULTADOS!B26), e nao mais a referencia
   fisica B10. O chip de situacao do card (C4 = H8) ja validava contra
   VTA_FINAL: card e status voltam a falar da mesma grandeza.

2) B10/B11 sao PRESERVADAS (valores e formulas intactos) e apenas rerotuladas
   como referencias auditaveis; deixam de se chamar "VTA OFICIAL". O mesmo vale
   para a medida 9 do bloco 6 (B63), que apontava para B10.

3) Bloco novo "9. COMO E FORMADO O VTA?" (linhas 79-87): quadro dinamico de
   quatro parcelas derivadas das fontes reais do metodo selecionado, mais uma
   linha de conferencia que prova a reconciliacao com o VTA Oficial. Nenhum
   valor digitado.

4) Bloco 8 (CONFERENCIA DA EXECUCAO, linhas 71-78) ganha linguagem didatica:
   rotulos de coluna, o texto de ausencia de base ("Sem historico fisico...")
   no lugar de "NAO COMPARAVEL", e a nota de que a conferencia nao altera o VTA
   Oficial. A coluna D continua vazia quando nao ha base comparavel — a
   verificacao passa a ser NOT(ISNUMBER(C)), mais robusta que comparar texto.

Uso:
    python tools/aplicar_vta_uniformizacao_u2.py <origem.xlsx> <destino.xlsx>
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

XL_CALC_MANUAL = -4135
XL_CALC_AUTOMATIC = -4105
MOEDA_BR = "R$ #.##0,00"

CORES = {
    "azul_muito_claro": 0xF7F0E8,  # BGR
}

# ---------------------------------------------------------------- constantes

LINHA_REF_ATUAL = 10
LINHA_REF_ABERTURA = 11
LINHA_MEDIDA_VTA = 63
LINHA_CONF_CABECALHO = 72
LINHA_CONF_C0 = 73
LINHA_CONF_NOTA = 78
LINHA_BLOCO_TITULO = 79

ROTULO_REF_ATUAL = "Referência auditável — posição física atual"
ROTULO_REF_ABERTURA = (
    "Referência auditável — última posição de abertura disponível"
)

SEM_HISTORICO_C = "Sem historico fisico suficiente"
SEM_HISTORICO_E = "Sem historico fisico para comparar"
NAO_APLICAVEL = "Nao aplicavel ao metodo selecionado"

NOTA_CONFERENCIA = (
    "Esta conferencia compara o Financeiro informado com uma reconstrucao pelo "
    "quantitativo. Ela nao altera o VTA Oficial."
)

TITULO_BLOCO = "9. COMO E FORMADO O VTA?"
TEXTO_BLOCO = (
    "O VTA reune o valor ja executado, os ajustes ainda devidos (quando houver) "
    "e o saldo que ainda falta executar, ja atualizado."
)
TEXTO_CONFERIR = (
    "Para conferir: VTA Oficial = Executado apurado + Ajustes ainda devidos "
    "+ Remanescente atualizado."
)

VTA_CANONICO = '=IF(VTA_FINAL="","",VTA_FINAL)'

# Colunas de execucao por ciclo ja resolvidas pelo proprio template
# (par adjacente unico, sem encadeamento) — mantidas exatamente como a VTA-M2.2
# as definiu; aqui so muda o rotulo de ausencia de base.
COLUNA_EXECUCAO = {"C0": "AC", "C1": "N", "C2": "P", "C3": "R"}


def _nomes_abas(wb) -> list[str]:
    return [ws.Name for ws in wb.Worksheets]


def _nomes_definidos(wb) -> set[str]:
    return {n.Name for n in wb.Names}


def _capturar_protecao(ws):
    estado = bool(ws.ProtectContents)
    selecao = None
    if estado:
        selecao = ws.EnableSelection
        ws.Unprotect()
    return estado, selecao


def _restaurar_protecao(ws, estado, selecao) -> None:
    if not estado:
        return
    ws.Protect(DrawingObjects=True, Contents=True, Scenarios=True)
    if selecao is not None:
        ws.EnableSelection = selecao


def _bordas(rng) -> None:
    for indice in (7, 8, 9, 10, 11, 12):
        try:
            borda = rng.Borders(indice)
            borda.LineStyle = 1
            borda.Weight = 2
            borda.Color = 0xBFBFBF
        except Exception:  # pragma: no cover - bordas sao cosmeticas
            pass


# ------------------------------------------------------------- validacoes

def _validar_origem(wb) -> None:
    abas = _nomes_abas(wb)
    for obrig in (ABA_MEMORIA, ABA_RESULTADOS):
        if obrig not in abas:
            raise ValueError(f"Aba {obrig} ausente na origem.")
    if "VTA_FINAL" not in _nomes_definidos(wb):
        raise ValueError("Nome definido VTA_FINAL ausente.")

    res = wb.Worksheets(ABA_RESULTADOS)
    if str(res.Range(f"A{LINHA_CONF_CABECALHO}").Value or "").strip() != "Ciclo":
        raise ValueError(
            "RESULTADOS!A72 nao e o cabecalho do bloco 8; origem fora do "
            "estado esperado (aplicar antes o aplicar_vta_financeiro_canonico)."
        )
    bloco = res.Range(f"A{LINHA_CONF_NOTA}:E{LINHA_BLOCO_TITULO + 8}")
    ocupadas = [c.Address() for c in bloco.Cells if c.Value not in (None, "")]
    if ocupadas:
        raise ValueError(
            "RESULTADOS!A78:E87 nao esta vazia (conteudo em "
            f"{', '.join(ocupadas[:5])}{'...' if len(ocupadas) > 5 else ''})."
        )


def _snapshot_memoria(wb) -> dict[str, str]:
    """Formulas de negocio do VTA que esta frente NAO pode tocar."""
    mem = wb.Worksheets(ABA_MEMORIA)
    return {
        e: str(mem.Range(e).Formula)
        for e in ("B16", "B20", "B21", "B22", "B23", "B26", "B28",
                  "D20", "F20", "D35", "T21", "T22", "T23", "T25")
    }


# ------------------------------------------------------------- aplicacoes

def _aplicar_card_e_referencias(res) -> None:
    # 1) Card VTA OFICIAL -> saida canonica calculada do metodo.
    res.Range("C5").Formula = VTA_CANONICO

    # 2) B10/B11 preservadas; apenas deixam de se chamar VTA OFICIAL.
    res.Range(f"A{LINHA_REF_ATUAL}").Value = ROTULO_REF_ATUAL
    res.Range(f"A{LINHA_REF_ABERTURA}").Value = ROTULO_REF_ABERTURA

    # Medida 9 do bloco 6 apontava para B10 (referencia fisica).
    res.Range(f"B{LINHA_MEDIDA_VTA}").Formula = VTA_CANONICO
    res.Range(f"C{LINHA_MEDIDA_VTA}").Value = (
        "MEMORIA!B26 (VTA_FINAL) - saida canonica do metodo selecionado"
    )


def _formula_execucao_teorica(ciclo: str) -> str:
    if ciclo == "C4":
        corpo = f'"{SEM_HISTORICO_C}"'
    else:
        coluna = COLUNA_EXECUCAO[ciclo]
        alvo = f"itens_Remanesc!${coluna}$2:${coluna}$201"
        corpo = (
            'IF(SUMPRODUCT((itens_Remanesc!$A$2:$A$201<>"")*'
            f"(1-ISNUMBER({alvo})))>0,"
            f'"{SEM_HISTORICO_C}",ROUND(SUM(N({alvo})),2))'
        )
    return (
        f'=IF(MEMORIA_RESULTADOS!$B$4<>"Financeiro","{NAO_APLICAVEL}",{corpo})'
    )


def _formula_diferenca(linha: int) -> str:
    return (
        f'=IF(MEMORIA_RESULTADOS!$B$4<>"Financeiro","{NAO_APLICAVEL}",'
        f'IF(OR(B{linha}="",C{linha}="",NOT(ISNUMBER(C{linha}))),"",'
        f"ROUND(B{linha}-C{linha},2)))"
    )


def _formula_conferencia(linha: int) -> str:
    return (
        '=IF(MEMORIA_RESULTADOS!$B$4<>"Financeiro","NAO APLICAVEL",'
        f'IF(OR(B{linha}="",C{linha}="",NOT(ISNUMBER(C{linha}))),'
        f'"{SEM_HISTORICO_E}",'
        f'IF(ROUND(D{linha},2)=0,"OK","REVISAR")))'
    )


def _aplicar_conferencia(res) -> None:
    cabecalho = (
        "Ciclo",
        "Desembolsado informado",
        "Execucao estimada pelo quantitativo",
        "Diferenca para o Financeiro informado",
        "Conferencia",
    )
    for offset, texto in enumerate(cabecalho):
        celula = res.Cells(LINHA_CONF_CABECALHO, 1 + offset)
        celula.Value = texto
        celula.Font.Bold = True

    for indice, ciclo in enumerate(("C0", "C1", "C2", "C3", "C4")):
        linha = LINHA_CONF_C0 + indice
        res.Range(f"C{linha}").Formula = _formula_execucao_teorica(ciclo)
        res.Range(f"D{linha}").Formula = _formula_diferenca(linha)
        res.Range(f"E{linha}").Formula = _formula_conferencia(linha)
        res.Range(f"B{linha}:D{linha}").NumberFormat = MOEDA_BR

    nota = res.Range(f"A{LINHA_CONF_NOTA}")
    nota.Value = NOTA_CONFERENCIA
    nota.Font.Italic = True


_F_EXECUTADO = (
    '=IF(MEMORIA_RESULTADOS!$B$4="Financeiro",'
    'IF(MEMORIA_RESULTADOS!$D$20="","",MEMORIA_RESULTADOS!$D$20),'
    'IF(MEMORIA_RESULTADOS!$B$4="Itens",'
    'IF(MEMORIA_RESULTADOS!$F$20="","",MEMORIA_RESULTADOS!$F$20),'
    'IF(MEMORIA_RESULTADOS!$B$4="PCs",'
    'IF(MEMORIA_RESULTADOS!$T$25="CALCULO MANUAL REQUERIDO","",'
    'ROUND(MEMORIA_RESULTADOS!$T$21+MEMORIA_RESULTADOS!$T$22,2)),"")))'
)

_F_EXECUTADO_FONTE = (
    '=IF(MEMORIA_RESULTADOS!$B$4="Financeiro",'
    '"Valores efetivamente pagos, conforme informados na aba Financeiro.",'
    'IF(MEMORIA_RESULTADOS!$B$4="Itens",'
    '"Quantidades consumidas x valores unitarios aplicaveis, ja atualizados.",'
    'IF(MEMORIA_RESULTADOS!$B$4="PCs",'
    '"Valor considerado dos Pedidos de Compra anteriores ao ciclo vigente.",'
    '"Selecione o metodo em CONTROLE!B1.")))'
)

_F_AJUSTES = (
    '=IF(MEMORIA_RESULTADOS!$B$4="Financeiro",'
    'IF(MEMORIA_RESULTADOS!$B$21="","",MEMORIA_RESULTADOS!$B$21),'
    'IF(OR(MEMORIA_RESULTADOS!$B$4="Itens",MEMORIA_RESULTADOS!$B$4="PCs"),0,""))'
)

_F_AJUSTES_FONTE = (
    '=IF(MEMORIA_RESULTADOS!$B$4="Financeiro",'
    '"Reajuste ja reconhecido e ainda nao contido no valor pago.",'
    'IF(MEMORIA_RESULTADOS!$B$4="Itens",'
    '"Nao aplicavel: o reajuste ja esta dentro da execucao atualizada.",'
    'IF(MEMORIA_RESULTADOS!$B$4="PCs",'
    '"Nao aplicavel: o retroativo ja esta dentro do valor considerado dos PCs.",'
    '"")))'
)

_F_REMANESCENTE = '=IF(MEMORIA_RESULTADOS!$D$35="","",MEMORIA_RESULTADOS!$D$35)'

_F_CONFERE = (
    '=IF(OR($B$83="",$B$85="",$B$86=""),"",'
    'ROUND($B$86-($B$83+N($B$84)+$B$85),2))'
)

_F_CONFERE_TEXTO = (
    '=IF($B$87="","Aguardando base para conferir.",'
    'IF(ABS($B$87)<=MEMORIA_RESULTADOS!$D$4,'
    '"As parcelas fecham com o VTA Oficial.",'
    'IF(OR(ISNUMBER(MEMORIA_RESULTADOS!$B$24),ISNUMBER(MEMORIA_RESULTADOS!$B$25)),'
    '"Diferenca corresponde ao ajuste manual registrado na secao 5.",'
    '"REVISE - as parcelas nao fecham com o VTA Oficial.")))'
)


def _aplicar_bloco_didatico(res) -> None:
    titulo = res.Range(f"A{LINHA_BLOCO_TITULO}")
    titulo.Value = TITULO_BLOCO
    titulo.Font.Bold = True
    res.Range(f"A{LINHA_BLOCO_TITULO + 1}").Value = TEXTO_BLOCO
    res.Range(f"A{LINHA_BLOCO_TITULO + 2}").Value = TEXTO_CONFERIR

    linha_cab = LINHA_BLOCO_TITULO + 3  # 82
    for offset, texto in enumerate(("Parcela", "Valor", "De onde vem")):
        celula = res.Cells(linha_cab, 1 + offset)
        celula.Value = texto
        celula.Font.Bold = True

    linhas = (
        ("Executado apurado", _F_EXECUTADO, _F_EXECUTADO_FONTE),
        ("(+) Ajustes ainda devidos", _F_AJUSTES, _F_AJUSTES_FONTE),
        ("(+) Remanescente atualizado", _F_REMANESCENTE,
         "Saldo que ainda falta executar, ja atualizado (MEMORIA!D35)."),
        ("(=) VTA Oficial", VTA_CANONICO,
         "Saida canonica do metodo (MEMORIA_RESULTADOS!B26 = VTA_FINAL)."),
    )
    for indice, (rotulo, formula, fonte) in enumerate(linhas):
        linha = linha_cab + 1 + indice  # 83..86
        res.Range(f"A{linha}").Value = rotulo
        res.Range(f"B{linha}").Formula = formula
        if str(fonte).startswith("="):
            res.Range(f"C{linha}").Formula = fonte
        else:
            res.Range(f"C{linha}").Value = fonte
        res.Range(f"B{linha}").NumberFormat = MOEDA_BR
    res.Range(f"A{linha_cab + 4}").Font.Bold = True
    res.Range(f"B{linha_cab + 4}").Font.Bold = True

    linha_conf = linha_cab + 5  # 87
    res.Range(f"A{linha_conf}").Value = "Conferencia (deve ser 0,00)"
    res.Range(f"B{linha_conf}").Formula = _F_CONFERE
    res.Range(f"B{linha_conf}").NumberFormat = MOEDA_BR
    res.Range(f"C{linha_conf}").Formula = _F_CONFERE_TEXTO

    bloco = res.Range(f"A{LINHA_BLOCO_TITULO}:C{linha_conf}")
    bloco.Interior.Color = CORES["azul_muito_claro"]
    _bordas(res.Range(f"A{linha_cab}:C{linha_conf}"))


def _aplicar_resultados(wb) -> None:
    res = wb.Worksheets(ABA_RESULTADOS)
    estado, selecao = _capturar_protecao(res)
    try:
        _aplicar_card_e_referencias(res)
        _aplicar_conferencia(res)
        _aplicar_bloco_didatico(res)
    finally:
        _restaurar_protecao(res, estado, selecao)


# ------------------------------------------------------------------ driver

def aplicar(origem: Path, destino: Path) -> None:
    origem = Path(origem).resolve()
    destino = Path(destino).resolve()
    if not origem.is_file():
        raise FileNotFoundError(origem)
    if origem == destino:
        raise ValueError("Origem e destino devem ser diferentes.")

    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_vta_u2_"))
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
        memoria_antes = _snapshot_memoria(wb)

        _aplicar_resultados(wb)

        memoria_depois = _snapshot_memoria(wb)
        if memoria_antes != memoria_depois:
            difs = {
                k: (memoria_antes[k], memoria_depois[k])
                for k in memoria_antes if memoria_antes[k] != memoria_depois[k]
            }
            raise RuntimeError(
                f"TRAVA VIOLADA: MEMORIA_RESULTADOS alterada pela VTA-U2: {difs}"
            )

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
        res = wb.Worksheets(ABA_RESULTADOS)
        if "VTA_FINAL" not in str(res.Range("C5").Formula):
            raise RuntimeError("C5 nao ficou apontando para VTA_FINAL.")
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
    print("VTA-U2 aplicado:", args.destino)


if __name__ == "__main__":
    main()
