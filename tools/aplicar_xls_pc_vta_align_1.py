# -*- coding: utf-8 -*-
"""XLS-PC-VTA-ALIGN-1: alinha o XLS a regra ja homologada do VTA-POT-1.

Regra unica, nao reaberta aqui:

    VTA OFICIAL = VTA SEM POTENCIAL + RETROATIVO POTENCIAL INCORPORADO
    RETROATIVO CONSIDERADO NO VTA = RECONHECIDO + POTENCIAL INCORPORADO

O VTA-POT-1 ja fechava o quadro 9 (B84/C86). Este aplicador propaga a MESMA
decomposicao para as demais areas VISIVEIS do metodo PC, sempre LENDO as
fontes canonicas (VTA_FINAL, VTA_SEM_POTENCIAL, RETROATIVO_POTENCIAL_VTA e
RESULTADOS!$D$22) — nenhuma delas recalcula o potencial.

itens_PC
    R2/R11  cabecalho passa a dizer "NAO PAGOS" (a coluna I zera quando
            PC_PAGO_A_CONTRATADA="Sim": e literalmente o valor em analise
            dos PCs nao pagos).
    19      linha JA existente e vazia, imediatamente antes de "COMO OS PCs
            SAO TRATADOS": recebe o fechamento explicito do retroativo.
    V:AC    reasseguradas ocultas (template e copia entregue).

RESULTADOS
    A4/A6   o card do VTA declara que o valor ja inclui a parcela POTENCIAL.
    D4      "RETROATIVO RECONHECIDO A PAGAR" — o card NAO passa a somar o
            potencial: publicar reconhecido+potencial sob "a pagar" faria o
            POTENCIAL parecer obrigacao constituida.
    8       faixa livre do bloco superior: card ambar do POTENCIAL e o
            fechamento "RETROATIVO CONSIDERADO NO VTA". Fora do metodo PC a
            linha permanece vazia, exatamente como hoje.
    55:67   bloco 6 passa a demonstrar reconhecido / potencial / considerado.
    81/83   o texto de conferencia do bloco 9 passa a nomear as parcelas.

Ancoras preservadas: EXECUTADO_APURADO (B83), AJUSTES_DEVIDOS (B84),
VTA_FINAL (B86) e CONFERENCIA_FORMACAO_VTA (B87). itens_PC!A:L intacta.
Financeiro e Itens Consumidos: nenhuma celula com valor novo.

O aplicador usa Excel COM para preservar recursos x14 que o openpyxl remove.

Uso: <python> tools/aplicar_xls_pc_vta_align_1.py [caminho.xlsx]
"""
from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

XL_LEFT, XL_RIGHT, XL_CENTER = -4131, -4152, -4108
XL_VCENTER = -4108
XL_CONTINUOUS, XL_THIN = 1, 2
XL_EXPRESSION = 2

AZUL_ESCURO = "1F4E78"
AZUL_MUITO_CLARO = "F2F7FC"
CINZA_BORDA = "BFBFBF"
TEXTO_PADRAO = "1F3864"
TEXTO_AUXILIAR = "595959"
BRANCO = "FFFFFF"
CARD_CLARO = "FAFAFA"

# Mesmo par do VTA-POT-1: a parcela POTENCIAL tem uma unica cor no produto.
POTENCIAL_BG = "FFF4CC"
POTENCIAL_TEXTO = "7F6000"

MOEDA_LOCAL = '"R$" #.##0,00;-"R$" #.##0,00;"R$" 0,00;"—"'
MOEDA_INVARIANTE = '"R$" #,##0.00;-"R$" #,##0.00;"R$" 0.00;"—"'

METODO = 'MEMORIA_RESULTADOS!$B$4'
E_PC = f'{METODO}="PCs"'
NAO_E_PC = f'{METODO}<>"PCs"'

# ------------------------------------------------------------------ itens_PC
R2_NOVO = "VALOR EM ANÁLISE - NÃO PAGOS (ÁREA GEST.)"
R11_NOVO = "Valor em análise — não pagos (regra vigente)"

# Fechamento do Quadro 2 na linha 19 (ja existente e vazia).
# Q18 = TOTAL do retroativo RECONHECIDO. A parcela POTENCIAL somada e a
# canonica (RETROATIVO_POTENCIAL_VTA), nao S18: S18 e o potencial de TODOS os
# ciclos ate o corte — inclusive o vigente e o residual —, enquanto o VTA so
# incorpora o potencial dos ciclos JA ENCERRADOS. Fechar por S18 publicaria um
# total que nao existe em lugar nenhum do VTA.
F_PC_19_ROTULO = (
    '="RETROATIVO CONSIDERADO NO VTA = RECONHECIDO R$ "&'
    'TEXT(N($Q$18),"#.##0,00")&"  +  POTENCIAL INCORPORADO R$ "&'
    'TEXT(N(MEMORIA_RESULTADOS!$T$39),"#.##0,00")&'
    '"   (só o potencial dos ciclos já encerrados entra no VTA)"'
)
F_PC_19_VALOR = '=ROUND(N($Q$18)+N(MEMORIA_RESULTADOS!$T$39),2)'

# --------------------------------------------------------------- RESULTADOS
F_A4 = (
    f'=IF({E_PC},'
    '"VTA OFICIAL — inclui retroativo reconhecido + retroativo potencial",'
    '"VTA OFICIAL")'
)
F_A6 = (
    f'=IF({NAO_E_PC},"",'
    '"Inclui R$ "&TEXT(N(RETROATIVO_POTENCIAL_VTA),"#.##0,00")&'
    '" de retroativo POTENCIAL (sujeito à confirmação da área gestora)")'
)
D4_NOVO = "RETROATIVO RECONHECIDO A PAGAR"

F_A8 = (
    f'=IF({NAO_E_PC},"",'
    '"RETROATIVO POTENCIAL — POTENCIAL (incorporado ao VTA)")'
)
F_C8 = f'=IF({NAO_E_PC},"",IFERROR(RETROATIVO_POTENCIAL_VTA,""))'
F_D8 = f'=IF({NAO_E_PC},"","RETROATIVO CONSIDERADO NO VTA")'
F_E8 = (
    f'=IF(OR({NAO_E_PC},$D$22=""),"",'
    'ROUND(N($D$22)+N(RETROATIVO_POTENCIAL_VTA),2))'
)

# Bloco 6 — 13 medidas em A55:C67. As cinco primeiras e a 6a nao mudam; as
# medidas 7 e 8 sao novas (potencial e considerado) e as demais descem uma
# linha. A antiga medida 10 ("Informacoes para conferencia" -> "Disponiveis
# para consulta") era um item sem valor apuravel e cede o lugar.
BLOCO6 = [
    (
        "1. Total cadastrado de Pedidos de Compra",
        '=IF($B$5<>"PCs","",MEMORIA_RESULTADOS!$T$33)',
        "Todos os PCs informados, sem aplicar a data de corte.",
    ),
    (
        "2. Total até a data de corte",
        '=IF($B$5<>"PCs","",MEMORIA_RESULTADOS!$T$34)',
        "Total dos PCs com data igual ou anterior à data de corte.",
    ),
    (
        "3. Total distribuído nos ciclos",
        '=IF($B$5<>"PCs","",MEMORIA_RESULTADOS!$T$35)',
        "PCs localizados nos ciclos contratuais e dentro da data de corte.",
    ),
    (
        "4. Total com efeito financeiro",
        '=IF($B$5<>"PCs","",MEMORIA_RESULTADOS!$T$36)',
        "PCs alcançados pelos efeitos financeiros do respectivo ciclo.",
    ),
    (
        "5. Total sem efeito financeiro",
        '=IF($B$5<>"PCs","",MEMORIA_RESULTADOS!$T$37)',
        "PCs anteriores ao início dos efeitos financeiros do ciclo.",
    ),
    (
        "6. Retroativo reconhecido",
        "=$D$22",
        "Diferença reconhecida dos PCs pagos e alcançados pelos efeitos "
        "financeiros. É o retroativo a pagar.",
    ),
    (
        "7. Retroativo potencial — POTENCIAL",
        '=IF($B$5<>"PCs","",IFERROR(RETROATIVO_POTENCIAL_VTA,""))',
        "Parcela dos PCs ainda em análise pela área gestora que o VTA "
        "incorpora por critério prudencial. Não é retroativo reconhecido "
        "nem valor a pagar.",
    ),
    (
        "8. Retroativo considerado no VTA",
        '=IF(OR($B$5<>"PCs",$D$22=""),"",'
        'ROUND(N($D$22)+N(RETROATIVO_POTENCIAL_VTA),2))',
        "Soma da medida 6 com a medida 7: o retroativo total que o Valor "
        "Total Atualizado já considera.",
    ),
    (
        "9. Execução do ciclo atual",
        "=$B$36",
        "Execução registrada para o ciclo em andamento.",
    ),
    (
        "10. Saldo remanescente atual",
        "=$B$38",
        "Saldo atualizado que ainda falta executar.",
    ),
    (
        "11. VTA oficial",
        '=IF(VTA_FINAL="","",VTA_FINAL)',
        "Valor Total Atualizado apurado pelo método selecionado. No método "
        "de Pedidos de Compra ele já inclui a medida 7.",
    ),
    (
        "12. Diferença entre as formas de cálculo",
        "=$B$13",
        "R$ 0,00 indica que as duas formas de cálculo chegaram ao mesmo "
        "valor.",
    ),
    (
        "13. Resultado da apuração",
        "=$B$3",
        "Situação final apresentada por esta aba.",
    ),
]
LINHA_POTENCIAL_BLOCO6 = 61          # medida 7
LINHA_VTA_BLOCO6 = 65                # medida 11 (negrito)
LINHA_STATUS_BLOCO6 = 67             # medida 13 (texto, nunca moeda)

F_A81 = (
    f'=IF({E_PC},'
    '"Para conferir: VTA oficial = execução apurada (já inclui o retroativo '
    'reconhecido) + retroativo POTENCIAL incorporado + saldo remanescente '
    'atualizado.",'
    '"Para conferir: VTA oficial = execução apurada + parcela adicional do '
    'método + saldo remanescente atualizado.")'
)
F_C83 = (
    f'=IF({METODO}="Financeiro",'
    '"Valores efetivamente pagos, conforme informados na aba Financeiro.",'
    f'IF({METODO}="Itens",'
    '"Quantidades consumidas x valores unitarios aplicaveis, ja atualizados.",'
    f'IF({E_PC},'
    '"Valor considerado dos Pedidos de Compra anteriores ao ciclo vigente, '
    'ja com o retroativo reconhecido.",'
    '"Selecione o metodo em CONTROLE!B1.")))'
)


def _bgr(hex_rgb: str) -> int:
    return int(hex_rgb[4:6] + hex_rgb[2:4] + hex_rgb[0:2], 16)


def _fechar(wb) -> None:
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass


def _moeda(rng) -> None:
    try:
        rng.NumberFormatLocal = MOEDA_LOCAL
    except Exception:
        rng.NumberFormat = MOEDA_INVARIANTE


def _formato_geral(rng) -> None:
    """Formato "Geral" — o Excel pt-BR recusa ``NumberFormat="General"``."""
    try:
        rng.NumberFormatLocal = "Geral"
    except Exception:
        rng.NumberFormat = "General"


def _bordas(rng) -> None:
    for indice in (7, 8, 9, 10, 11, 12):
        borda = rng.Borders(indice)
        borda.LineStyle = XL_CONTINUOUS
        borda.Weight = XL_THIN
        borda.Color = _bgr(CINZA_BORDA)


def _remover_cf_propria(rng, formula: str) -> None:
    """Remove APENAS a regra deste aplicador (idempotencia).

    NUNCA usar ``FormatConditions.Delete()`` na faixa inteira: as abas
    oficiais carregam regras x14 homologadas que seriam destruidas.
    """
    condicoes = rng.FormatConditions
    for indice in range(condicoes.Count, 0, -1):
        try:
            atual = condicoes.Item(indice)
            if str(atual.Formula1 or "").replace(" ", "") == formula.replace(" ", ""):
                atual.Delete()
        except Exception:
            continue


def _cf(rng, formula: str, fundo: str, texto: str | None = None) -> None:
    _remover_cf_propria(rng, formula)
    cond = rng.FormatConditions.Add(XL_EXPRESSION, 0, formula)
    cond.Interior.Color = _bgr(fundo)
    if texto:
        cond.Font.Color = _bgr(texto)


# ------------------------------------------------------------------ itens_PC
def _aplicar_itens_pc(ws) -> None:
    principal_antes = ws.Range("A1:L5001").Formula

    ws.Range("R2").Value = R2_NOVO
    ws.Range("R11").Value = R11_NOVO

    # Linha 19: ja existente, vazia e imediatamente antes de "COMO OS PCs SAO
    # TRATADOS". Nenhuma linha fisica e inserida.
    for endereco in ("M19:T19", "M19:R19", "S19:T19"):
        if ws.Range(endereco).MergeCells:
            ws.Range(endereco).UnMerge()
    ws.Range("M19:T19").ClearContents()

    ws.Range("M19:R19").Merge()
    ws.Range("M19").Formula = F_PC_19_ROTULO
    rotulo = ws.Range("M19:R19")
    rotulo.WrapText = True
    rotulo.HorizontalAlignment = XL_LEFT
    rotulo.VerticalAlignment = XL_VCENTER
    rotulo.Font.Bold = True
    rotulo.Font.Size = 9
    rotulo.Font.Color = _bgr(TEXTO_PADRAO)

    ws.Range("S19:T19").Merge()
    ws.Range("S19").Formula = F_PC_19_VALOR
    valor = ws.Range("S19:T19")
    valor.HorizontalAlignment = XL_RIGHT
    valor.VerticalAlignment = XL_VCENTER
    valor.Font.Bold = True
    valor.Font.Size = 11
    valor.Font.Color = _bgr(TEXTO_PADRAO)
    _moeda(valor)

    ws.Range("M19:T19").Interior.Color = _bgr(AZUL_MUITO_CLARO)
    _bordas(ws.Range("M19:T19"))
    ws.Rows("19:19").RowHeight = 32

    # V:AC seguem ocultas — a faixa ja existe como um unico grupo no template.
    ws.Range("V:AC").EntireColumn.Hidden = True

    if ws.Range("A1:L5001").Formula != principal_antes:
        raise RuntimeError("A faixa operacional itens_PC!A:L foi alterada.")


# ---------------------------------------------------------------- RESULTADOS
def _aplicar_bloco_superior(ws) -> None:
    ws.Range("A4").Formula = F_A4
    ws.Range("D4").Value = D4_NOVO

    # Rodape do card do VTA (faixa livre, mesmo lugar dos rodapes E6/F6).
    ws.Range("A6").Formula = F_A6

    # Linha 8: livre no bloco superior. Sem metodo PC ela devolve "" e
    # continua sendo o separador branco de hoje.
    ws.Range("A8").Formula = F_A8
    ws.Range("C8").Formula = F_C8
    ws.Range("D8").Formula = F_D8
    ws.Range("E8").Formula = F_E8

    ws.Range("A8:E8").Font.Size = 9
    ws.Range("A8:E8").VerticalAlignment = XL_VCENTER
    ws.Range("A8:B8").HorizontalAlignment = XL_LEFT
    ws.Range("D8").HorizontalAlignment = XL_LEFT
    ws.Range("C8").HorizontalAlignment = XL_RIGHT
    ws.Range("E8").HorizontalAlignment = XL_RIGHT
    for endereco in ("A8", "C8", "D8", "E8"):
        ws.Range(endereco).Font.Bold = True
    _moeda(ws.Range("C8"))
    _moeda(ws.Range("E8"))
    ws.Range("A8:C8").Font.Color = _bgr(POTENCIAL_TEXTO)
    ws.Range("D8:E8").Font.Color = _bgr(TEXTO_PADRAO)

    # Ambar SO na parcela que e especificamente POTENCIAL; o fechamento
    # (D8:E8) fica neutro. Fora do metodo PC nada e pintado.
    _cf(ws.Range("A8:C8"), '=$C$8<>""', POTENCIAL_BG, POTENCIAL_TEXTO)
    _cf(ws.Range("D8:E8"), '=$E$8<>""', CARD_CLARO, TEXTO_PADRAO)


def _aplicar_bloco6(ws) -> None:
    # Linha 67 herda por copia o estilo, as bordas e o merge C:H da linha 66.
    ws.Range("A66:H66").Copy(ws.Range("A67:H67"))
    ws.Rows("67:67").RowHeight = 30

    for indice, (rotulo, formula, explicacao) in enumerate(BLOCO6):
        linha = 55 + indice
        ws.Cells(linha, 1).Value = rotulo
        ws.Cells(linha, 2).Formula = formula
        ws.Cells(linha, 3).Value = explicacao

    ws.Range("A55:A67").HorizontalAlignment = XL_LEFT
    ws.Range("B55:B67").HorizontalAlignment = XL_RIGHT
    _moeda(ws.Range(f"B55:B{LINHA_STATUS_BLOCO6 - 1}"))
    _formato_geral(ws.Range(f"B{LINHA_STATUS_BLOCO6}"))
    ws.Range("B55:B67").Font.Bold = False
    ws.Range(f"B{LINHA_VTA_BLOCO6}").Font.Bold = True
    ws.Range(f"B{LINHA_STATUS_BLOCO6}").Font.Bold = True
    # A grade comeca no cabecalho (54): o Excel guarda a aresta compartilhada
    # uma vez so, e pintar apenas a partir da 55 apagaria a base da linha 54.
    _bordas(ws.Range("A54:C67"))

    linha_pot = LINHA_POTENCIAL_BLOCO6
    _cf(
        ws.Range(f"A{linha_pot}:B{linha_pot}"),
        f'=$B${linha_pot}<>""',
        POTENCIAL_BG,
        POTENCIAL_TEXTO,
    )


def _aplicar_bloco9(ws) -> None:
    ws.Range("A81").Formula = F_A81
    ws.Range("C83").Formula = F_C83


def aplicar(caminho: Path) -> None:
    import win32com.client as com

    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        cf_antes = {
            nome: wb.Worksheets(nome).Cells.FormatConditions.Count
            for nome in ("RESULTADOS", "itens_PC", "MEMORIA_RESULTADOS")
        }
        resultados = wb.Worksheets("RESULTADOS")
        _aplicar_itens_pc(wb.Worksheets("itens_PC"))
        _aplicar_bloco_superior(resultados)
        _aplicar_bloco6(resultados)
        _aplicar_bloco9(resultados)
        excel.CutCopyMode = False
        for nome, antes in cf_antes.items():
            depois = wb.Worksheets(nome).Cells.FormatConditions.Count
            if depois < antes:
                raise RuntimeError(
                    f"{nome}: formatacao condicional perdida "
                    f"({antes} -> {depois})."
                )
        excel.CalculateFullRebuild()
        wb.Worksheets("CONTROLE").Activate()
        wb.Save()
        _fechar(wb)
        wb = None
    finally:
        if wb is not None:
            _fechar(wb)
        excel.Quit()


def main() -> int:
    caminho = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else TEMPLATE
    if not caminho.exists():
        print(f"ERRO: arquivo não encontrado: {caminho}")
        return 1
    aplicar(caminho)
    print(f"OK: XLS-PC-VTA-ALIGN-1 aplicado em {caminho}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
