# -*- coding: utf-8 -*-
"""PC-UX-1: padroniza a apresentação do método PC no XLS oficial.

O aplicador usa Excel COM para preservar recursos x14 que o openpyxl remove.
Não altera A:L de ``itens_PC``, ``MEMORIA_RESULTADOS`` nem fórmulas econômicas.
O quadro M:T é deslocado apenas dentro da área secundária, com todas as suas
dependências conhecidas reancoradas e verificadas.

Uso: python tools/aplicar_pc_ux_1.py [caminho.xlsx]
"""
from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

XL_LEFT, XL_RIGHT, XL_CENTER = -4131, -4152, -4108
XL_VCENTER = -4108
XL_CONTINUOUS, XL_THIN = 1, 2
XL_NONE = -4142

AZUL_ESCURO = "1F4E78"
AZUL_CLARO = "D9EAF7"
AZUL_MUITO_CLARO = "F2F7FC"
BRANCO = "FFFFFF"
CINZA_BORDA = "BFBFBF"
CINZA_TEXTO = "595959"
TEXTO_PADRAO = "1F3864"

MOEDA_LOCAL = '"R$" #.##0,00;-"R$" #.##0,00;"R$" 0,00;"—"'
MOEDA_INVARIANTE = '"R$" #,##0.00;-"R$" #,##0.00;"R$" 0.00;"—"'

TITULO_TODOS = "TODOS OS PCs CADASTRADOS POR CICLO"
TITULO_CONSIDERADOS = "PCs CONSIDERADOS NA APURAÇÃO ATÉ A DATA DE CORTE"


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


def _bordas(rng) -> None:
    for indice in (7, 8, 9, 10, 11, 12):
        borda = rng.Borders(indice)
        borda.LineStyle = XL_CONTINUOUS
        borda.Weight = XL_THIN
        borda.Color = _bgr(CINZA_BORDA)


def _titulo(ws, endereco: str, texto: str, tamanho: int = 11) -> None:
    rng = ws.Range(endereco)
    rng.Merge()
    rng.Value = texto
    rng.Interior.Color = _bgr(AZUL_ESCURO)
    rng.Font.Color = _bgr(BRANCO)
    rng.Font.Bold = True
    rng.Font.Size = tamanho
    rng.HorizontalAlignment = XL_LEFT
    rng.VerticalAlignment = XL_VCENTER


def _cabecalho(rng) -> None:
    rng.Interior.Color = _bgr(AZUL_CLARO)
    rng.Font.Color = _bgr(TEXTO_PADRAO)
    rng.Font.Bold = True
    rng.Font.Size = 9
    rng.HorizontalAlignment = XL_CENTER
    rng.VerticalAlignment = XL_VCENTER
    rng.WrapText = True


def _aplicar_itens_pc(ws, memoria) -> None:
    # A:L é o contrato operacional. A leitura matricial torna qualquer mudança
    # nessa faixa observável antes de salvar.
    principal_antes = ws.Range("A1:L5001").Formula

    headers = tuple(ws.Range("M1:T1").Value[0])
    esperado = (
        "CICLO", "QTD_PC", "VALOR_PC_TOTAL", "VALOR_ATUALIZADO_TOTAL",
        "RETROATIVO_RECONHECIDO_A_PAGAR", "VALOR_ATUALIZADO_EM_ANALISE",
        "DELTA_POTENCIAL", "QTD_COM_CHECK",
    )
    ja_aplicado = str(ws.Range("M1").Value or "") == TITULO_TODOS
    if not ja_aplicado and headers != esperado:
        raise RuntimeError(f"M:T divergiu do quadro homologado: {headers!r}")

    # Move somente o quadro lateral; o Excel atualiza referências externas.
    if not ja_aplicado:
        ws.Range("M1:T7").Cut(ws.Range("M2:T8"))
    _titulo(ws, "M1:T1", TITULO_TODOS)
    novos_headers = (
        "CICLO", "QTD. DE PCs", "VALOR ORIGINAL TOTAL",
        "VALOR DOS PCs COM FATOR DO CICLO", "RETROATIVO RECONHECIDO",
        "VALOR EM ANÁLISE (ÁREA GEST.)", "RETROATIVO POTENCIAL",
        "QTD. COM ALERTA",
    )
    for coluna, valor in enumerate(novos_headers, start=13):
        ws.Cells(2, coluna).Value = valor
    _cabecalho(ws.Range("M2:T2"))
    _bordas(ws.Range("M2:T8"))
    ws.Rows("1:2").RowHeight = 30
    ws.Range("M3:M8").HorizontalAlignment = XL_CENTER
    ws.Range("N3:N8").HorizontalAlignment = XL_CENTER
    _moeda(ws.Range("O3:S8"))
    ws.Range("O3:S8").HorizontalAlignment = XL_RIGHT

    # O deslocamento do quadro atualiza endereços, mas não a constante aritmética
    # do ROW(...)-2. Reancorar -3 preserva exatamente o significado anterior.
    if not ja_aplicado:
        for endereco in ("T22", "W67"):
            formula = str(memoria.Range(endereco).Formula)
            formula = formula.replace("-2>=1", "-3>=1").replace("-2<", "-3<")
            memoria.Range(endereco).Formula = formula

    _titulo(ws, "M10:T10", TITULO_CONSIDERADOS)
    headers_considerados = (
        "Ciclo", "PCs pagos/reconhecidos", "Valor original", "Valor atualizado",
        "Retroativo reconhecido", "Valor em análise (área gest.)",
        "Retroativo potencial", "Fora da data de corte",
    )
    for coluna, valor in enumerate(headers_considerados, start=13):
        ws.Cells(11, coluna).Value = valor
    _cabecalho(ws.Range("M11:T11"))
    ws.Rows("10:11").RowHeight = 34

    for linha, ciclo in enumerate(("C0", "C1", "C2", "C3", "C4"), start=12):
        ws.Cells(linha, 13).Value = ciclo
        base = f'$C$2:$C$5001,$M{linha}'
        pago = '$G$2:$G$5001,"Sim"'
        corte = 'MEMORIA_RESULTADOS!$T$31'
        ws.Cells(linha, 14).Formula = (
            f'=IF({corte}="",COUNTIFS({base},{pago}),'
            f'COUNTIFS({base},{pago},$B$2:$B$5001,"<="&{corte}))'
        )
        for coluna, fonte in ((15, "D"), (16, "U"), (17, "H")):
            ws.Cells(linha, coluna).Formula = (
                f'=IF({corte}="",SUMIFS(${fonte}$2:${fonte}$5001,{base},{pago}),'
                f'SUMIFS(${fonte}$2:${fonte}$5001,{base},{pago},'
                f'$B$2:$B$5001,"<="&{corte}))'
            )
        for coluna, fonte in ((18, "I"), (19, "J")):
            ws.Cells(linha, coluna).Formula = (
                f'=IF({corte}="",SUMIFS(${fonte}$2:${fonte}$5001,{base}),'
                f'SUMIFS(${fonte}$2:${fonte}$5001,{base},'
                f'$B$2:$B$5001,"<="&{corte}))'
            )
        ws.Cells(linha, 20).Formula = (
            f'=IF({corte}="",0,SUMIFS($D$2:$D$5001,{base},'
            f'$B$2:$B$5001,">"&{corte}))'
        )

    ws.Range("M17").Value = "TOTAL"
    for coluna in range(14, 21):
        letra = chr(64 + coluna)
        ws.Cells(17, coluna).Formula = f"=SUM({letra}12:{letra}16)"
    ws.Range("M12:N17").HorizontalAlignment = XL_CENTER
    _moeda(ws.Range("O12:T17"))
    ws.Range("O12:T17").HorizontalAlignment = XL_RIGHT
    ws.Range("M17:T17").Font.Bold = True
    ws.Range("M17:T17").Interior.Color = _bgr(AZUL_MUITO_CLARO)
    _bordas(ws.Range("M11:T17"))

    explicacoes = (
        "COMO OS PCs SÃO TRATADOS",
        "PC pago e dentro da data de corte: integra a execução considerada.",
        "PC não pago e dentro da data de corte: permanece como valor em análise pela área gestora.",
        "C0: não recebe reajuste e não gera retroativo.",
        "C1 em diante: pode receber reajuste conforme os efeitos financeiros.",
        "PC pago + efeito financeiro: retroativo reconhecido.",
        "PC não pago + efeito financeiro: retroativo potencial.",
        "PC posterior à data de corte: não participa desta apuração.",
    )
    for linha, texto in enumerate(explicacoes, start=19):
        ws.Range(f"M{linha}:T{linha}").Merge()
        ws.Range(f"M{linha}").Value = texto
        ws.Range(f"M{linha}:T{linha}").WrapText = True
        ws.Range(f"M{linha}:T{linha}").VerticalAlignment = XL_VCENTER
        ws.Range(f"M{linha}:T{linha}").HorizontalAlignment = XL_LEFT
        ws.Rows(f"{linha}:{linha}").RowHeight = 28
    _titulo(ws, "M19:T19", explicacoes[0])
    ws.Range("M20:T26").Interior.Color = _bgr(AZUL_MUITO_CLARO)
    ws.Range("M20:T26").Font.Color = _bgr(TEXTO_PADRAO)
    ws.Range("M20:T26").Font.Size = 9

    larguras = {"M": 10, "N": 18, "O": 16, "P": 17, "Q": 18,
                "R": 20, "S": 18, "T": 18}
    for coluna, largura in larguras.items():
        ws.Columns(coluna).ColumnWidth = largura

    if ws.Range("A1:L5001").Formula != principal_antes:
        raise RuntimeError("A faixa operacional itens_PC!A:L foi alterada.")


def _limpar_bordas(rng) -> None:
    for indice in (7, 8, 9, 10, 11, 12):
        rng.Borders(indice).LineStyle = XL_NONE


def _aplicar_resultados(ws) -> None:
    for endereco in ("C9:E9", "F9:G9"):
        if ws.Range(endereco).MergeCells:
            ws.Range(endereco).UnMerge()
    _titulo(ws, "A9:H9", "1. COMO O VTA FOI CALCULADO")
    ws.Rows("9:9").RowHeight = 24

    # Linhas 10:13 seguem ocultas e são contrato técnico do leitor. Preservar
    # seus rótulos e o status bruto evita contaminar a API interna com a camada
    # de linguagem simples aplicada nas áreas visíveis.
    rotulos_a = {
        10: "[AUDITORIA INTERNA] Posicao fisica atual - nao e VTA",
        11: "[AUDITORIA INTERNA] Ultima posicao de abertura - nao e VTA",
        12: "[AUDITORIA INTERNA] Contrato original integralmente reajustado - comparativo",
        13: "[AUDITORIA INTERNA] Reconciliacao (posicao atual - ultima abertura)",
    }
    textos_c = {
        10: '=IF(MEMORIA_RESULTADOS!$W$50="","POSICAO ATUAL NAO INFORMADA (CICLO_EM_EXECUCAO ausente/incompleta)",IF(MEMORIA_RESULTADOS!$B$4="Financeiro","execucao historica financeiro!E ciclos encerrados (MEMORIA!W66) + execucao ate a data + remanescente atual (CICLO_EM_EXECUCAO!F/G)","C0 exec (MEMORIA!T21) + ciclos encerrados (MEMORIA!T22) + execucao ate a data + remanescente atual (CICLO_EM_EXECUCAO!F/G)"))',
        11: '=IF(MEMORIA_RESULTADOS!$B$4="Financeiro","execucao historica financeiro!E ate a abertura adotada (MEMORIA!W67)","C0 exec (MEMORIA!T21) + ciclos encerrados (MEMORIA!T22)")&" + remanescente na abertura C"&IF(MEMORIA_RESULTADOS!$W$46="","?",MEMORIA_RESULTADOS!$W$46)&" (MEMORIA!AB)"&IF(MEMORIA_RESULTADOS!$W$47=""," "," — "&MEMORIA_RESULTADOS!$W$47)',
        12: "SUMPRODUCT(posicao_contratual!B x C) x RESULTADOS!H5 (fator historico integral)",
        13: "Diferenca entre FORMA 1 e FORMA 2 (0 = reconciliado)",
    }
    textos_f = {
        10: "CICLO_EM_EXECUCAO (fiscal) + itens_PC (C0/encerrados)",
        11: "itens_PC!O+Q (valor considerado) + posicao_contratual (G/K/O/S/W) x historico_VU",
        12: "comparativo_VTA!B208",
        13: "MEMORIA!W50 − MEMORIA!W48",
    }
    for linha in range(10, 14):
        ws.Range(f"A{linha}").Value = rotulos_a[linha]
        if str(textos_c[linha]).startswith("="):
            ws.Range(f"C{linha}").Formula = textos_c[linha]
        else:
            ws.Range(f"C{linha}").Value = textos_c[linha]
        ws.Range(f"F{linha}").Value = textos_f[linha]
    ws.Range("H13").Formula = "=MEMORIA_RESULTADOS!$W$52"

    ws.Range("A15").Value = "2. EXECUÇÃO E RETROATIVO POR CICLO"
    ws.Range("B15").Value = "Valor original"
    ws.Range("C15").Value = "Valor atualizado"
    ws.Range("D15").Value = "Retroativo"
    ws.Range("A24").Value = "3. SALDO REMANESCENTE POR CICLO"
    ws.Range("B25").Value = "Saldo sem reajuste"
    ws.Range("C25").Value = "Saldo atualizado"
    ws.Range("D25").Value = "Diferença"
    ws.Range("E23").Formula = (
        '=IF($A$23="","","PCs anteriores ao início dos efeitos financeiros "'
        '&"não geram retroativo.")'
    )

    if ws.Range("A53:H53").MergeCells:
        ws.Range("A53:H53").UnMerge()
    _titulo(ws, "A53:H53", "6. RESUMO DOS PRINCIPAIS VALORES DA APURAÇÃO")
    for linha in range(54, 67):
        rng = ws.Range(f"C{linha}:H{linha}")
        if rng.MergeCells:
            rng.UnMerge()
        rng.Merge()
    ws.Range("C54").Value = "O QUE ESTE VALOR REPRESENTA"
    _cabecalho(ws.Range("A54:H54"))
    ws.Range("A54").HorizontalAlignment = XL_LEFT
    ws.Range("B54").HorizontalAlignment = XL_RIGHT
    resumo = {
        55: ("1. Total cadastrado de Pedidos de Compra", "Todos os PCs informados, sem aplicar a data de corte."),
        56: ("2. Total até a data de corte", "Total dos PCs com data igual ou anterior à data de corte."),
        57: ("3. Total distribuído nos ciclos", "PCs localizados nos ciclos contratuais e dentro da data de corte."),
        58: ("4. Total com efeito financeiro", "PCs alcançados pelos efeitos financeiros do respectivo ciclo."),
        59: ("5. Total sem efeito financeiro", "PCs anteriores ao início dos efeitos financeiros do ciclo."),
        60: ("6. Retroativo reconhecido", "Diferença reconhecida dos PCs pagos e alcançados pelos efeitos financeiros."),
        61: ("7. Execução do ciclo atual", "Execução registrada para o ciclo em andamento."),
        62: ("8. Saldo remanescente atual", "Saldo atualizado que ainda falta executar."),
        63: ("9. VTA oficial", "Valor Total Atualizado apurado pelo método selecionado."),
        64: ("10. Informações para conferência", "Detalhes das posições contratuais disponíveis para consulta."),
        65: ("11. Diferença entre as formas de cálculo", "R$ 0,00 indica que as duas formas de cálculo chegaram ao mesmo valor."),
        66: ("12. Resultado da apuração", "Situação final apresentada por esta aba."),
    }
    for linha, (rotulo, explicacao) in resumo.items():
        ws.Range(f"A{linha}").Value = rotulo
        ws.Range(f"C{linha}").Value = explicacao
        ws.Range(f"C{linha}:H{linha}").WrapText = True
        ws.Rows(f"{linha}:{linha}").RowHeight = 30 if linha != 60 else 36
    ws.Range("B64").Formula = '="Disponíveis para consulta"'
    _bordas(ws.Range("A54:H66"))

    if ws.Range("A68:H68").MergeCells:
        ws.Range("A68:H68").UnMerge()
    _titulo(ws, "A68:H68", "7. METODOLOGIA DO VTA")
    if ws.Range("A69:H69").MergeCells:
        ws.Range("A69:H69").UnMerge()
    ws.Range("A69:H69").Merge()
    ws.Range("A69").Formula = (
        '="COMO O VTA É CALCULADO — execução já realizada + ajustes aplicáveis + "'
        '&"saldo remanescente atualizado. MÉTODO: "&IF(MEMORIA_RESULTADOS!$B$4="Financeiro",'
        '"Financeiro",IF(MEMORIA_RESULTADOS!$B$4="PCs","Pedidos de Compra",'
        'IF(MEMORIA_RESULTADOS!$B$4="Itens","Consumido","não selecionado")))'
    )
    ws.Range("A69:H69").WrapText = True
    ws.Range("A69:H69").Interior.Color = _bgr(AZUL_MUITO_CLARO)
    ws.Rows("69:69").RowHeight = 34
    ws.Range("A70:H70").ClearContents()
    _limpar_bordas(ws.Range("A70:H70"))

    if ws.Range("A71:H71").MergeCells:
        ws.Range("A71:H71").UnMerge()
    _titulo(ws, "A71:H71", "8. DIFERENÇA ENTRE AS FORMAS DE CÁLCULO (FINANCEIRO)")
    ws.Range("B72").Value = "Valor informado"
    ws.Range("C72").Value = "Valor estimado"
    ws.Range("D72").Value = "Diferença"
    ws.Range("E72").Value = "Resultado"
    for linha in range(73, 78):
        formula = str(ws.Range(f"E{linha}").Formula)
        ws.Range(f"E{linha}").Formula = formula.replace('"OK"', '"SEM DIFERENÇA"').replace(
            '"REVISAR"', '"VERIFICAR DIFERENÇA"'
        )
    ws.Range("A78:H78").ClearContents()
    _limpar_bordas(ws.Range("A78:H78"))

    if ws.Range("A79:H79").MergeCells:
        ws.Range("A79:H79").UnMerge()
    _titulo(ws, "A79:H79", "9. VALOR TOTAL ATUALIZADO DO CONTRATO", tamanho=12)
    for linha in (80, 81):
        rng = ws.Range(f"A{linha}:H{linha}")
        if rng.MergeCells:
            rng.UnMerge()
        rng.Merge()
    ws.Range("A80").Value = (
        "O VTA reúne a execução já realizada, os ajustes contratuais aplicáveis "
        "e o saldo que ainda falta executar, todos nos valores correspondentes."
    )
    ws.Range("A81").Value = (
        "Para conferir: VTA oficial = execução apurada + ajustes aplicáveis + "
        "saldo remanescente atualizado."
    )
    ws.Range("A80:H81").WrapText = True
    ws.Rows("80:81").RowHeight = 28
    ws.Range("A87").Value = "Diferença entre as formas de cálculo (deve ser R$ 0,00)"
    ws.Range("C87").Formula = (
        '=IF($B$87="","Aguardando base para conferir.",'
        'IF(ABS($B$87)<=MEMORIA_RESULTADOS!$D$4,"SEM DIFERENÇA",'
        '"VERIFICAR DIFERENÇA"))'
    )


def aplicar(caminho: Path) -> None:
    import win32com.client as com

    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        _aplicar_itens_pc(wb.Worksheets("itens_PC"), wb.Worksheets("MEMORIA_RESULTADOS"))
        _aplicar_resultados(wb.Worksheets("RESULTADOS"))
        excel.CalculateFullRebuild()
        wb.Worksheets("itens_PC").Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
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
    print(f"OK: PC-UX-1 aplicado em {caminho}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
