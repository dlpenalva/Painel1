"""Rótulos e referências da aba RESULTADOS; não altera fórmulas econômicas."""

NOTA_EXECUCAO = "Já contém o reajuste/retroativo reconhecido incorporado à execução."
NOTA_POTENCIAL = "Parcela prudencial sujeita à confirmação da área gestora."
SALDO = "Saldo remanescente final atualizado"
ALTURA_LINHA_86 = 144.5


def valores_apresentacao_pc(ws):
    def condicional(texto, anterior):
        ramo = anterior[1:] if isinstance(anterior, str) and anterior.startswith("=") else '"' + str(anterior or "").replace('"', '""') + '"'
        return '=IF($B$5="PCs","' + texto.replace('"', '""') + '",' + ramo + ')'
    rotulos = {
        "A15": "2. EXECUÇÃO RECONHECIDA EM PCs POR CICLO",
        "B15": "Valor executado original",
        "C15": "Valor executado atualizado",
        "D15": "Reajuste incorporado",
        "A24": "3. REMANESCENTE — REFERÊNCIAS POR CICLO",
        "B25": "Saldo sem reajuste de referência do ciclo",
        "C25": "Saldo atualizado de referência do ciclo",
        "D25": "Diferença de referência do ciclo",
        "E25": "Referências do ciclo; não são necessariamente saldos finais após execução. Os ciclos não se somam.",
        "A83": "Execução realizada — valor atualizado",
        "C83": NOTA_EXECUCAO,
        "A85": SALDO,
        "C85": "Parcela final utilizada no VTA; não somar os saldos históricos.",
        "A86": "VALOR TOTAL ATUALIZADO DO CONTRATO",
        "A80": "Execução atualizada consolidada + saldo remanescente final atualizado + retroativo potencial = VTA oficial. O quadro de PCs reconhecidos acima não substitui a execução consolidada, que pode incluir posição física.",
    }
    valores = {}
    for celula, texto in rotulos.items():
        anterior = ws[celula].value
        prefixo = '=IF($B$5="PCs","' + texto.replace('"', '""') + '",'
        valores[celula] = anterior if isinstance(anterior, str) and anterior.startswith(prefixo) else condicional(texto, anterior)
    # Somente referências de apresentação. O VTA e seus auxiliares continuam
    # intactos. Na posição física, F/G já contêm as parcelas canônicas.
    for celula, expressao in {
        "B83": 'IF(MEMORIA_RESULTADOS!$T$25="CALCULO MANUAL REQUERIDO","",ROUND(MEMORIA_RESULTADOS!$T$21+MEMORIA_RESULTADOS!$T$22+IF(MEMORIA_RESULTADOS!$W$49=1,$B$36,0),2))',
        "B85": 'IF(MEMORIA_RESULTADOS!$T$25="CALCULO MANUAL REQUERIDO","",IF(MEMORIA_RESULTADOS!$W$49=1,SUM(INDIRECT("CICLO_EM_EXECUCAO!G13:G211")),MEMORIA_RESULTADOS!$T$23))',
    }.items():
        prefixo = '=IF($B$5="PCs",' + expressao + ','
        anterior = ws[celula].value
        valores[celula] = anterior if anterior.startswith(prefixo) else prefixo + anterior[1:] + ')'
    # Mantém a observação prudencial detalhada (inclusive potencial negativo).
    valores["C84"] = ws["C84"].value
    if NOTA_POTENCIAL not in valores["C84"]:
        valores["C84"] = '=IF($B$5="PCs","' + NOTA_POTENCIAL + ' ","")&(' + valores["C84"][1:] + ')'
    return valores


def garantir_apresentacao_pc(wb):
    if "RESULTADOS" not in wb.sheetnames:
        return
    ws = wb["RESULTADOS"]
    for celula, valor in valores_apresentacao_pc(ws).items():
        ws[celula] = valor
    # Mantém as larguras; cabeçalhos maiores usam quebra nas mesmas células.
    from copy import copy
    for celula in ("B15", "C15", "D15", "B25", "C25", "D25"):
        alinhamento = copy(ws[celula].alignment)
        alinhamento.wrap_text = True
        ws[celula].alignment = alinhamento
    ws.row_dimensions[25].height = max(ws.row_dimensions[25].height or 0, 48)
    ws.row_dimensions[86].height = max(
        ws.row_dimensions[86].height or 0, ALTURA_LINHA_86
    )
