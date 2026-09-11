"""Quadros PC compartilhados pelos consumidores; não altera a apuração."""
from math import isfinite

EXECUCAO = "Execução realizada — valor atualizado"
SALDO = "Saldo remanescente final atualizado"
POTENCIAL = "Retroativo potencial"
TOTAL = "VALOR TOTAL ATUALIZADO DO CONTRATO"
NOTA_EXECUCAO = "Já contém o reajuste/retroativo reconhecido incorporado à execução."
NOTA_POTENCIAL = "Parcela prudencial sujeita à confirmação da área gestora."
NOTA_REFERENCIA = "Valores de referência do ciclo; não representam necessariamente o saldo final após a execução. Os ciclos não se somam."
NOTA_REAJUSTE = "Reajuste incorporado corresponde à diferença entre o valor executado atualizado e o valor executado original e já está incluído no valor atualizado."


def _numero(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and isfinite(v)


def ler_referencias_pc(wb):
    """Transporta o quadro histórico já calculado do XLS, sem fórmulas novas."""
    if "RESULTADOS" not in wb.sheetnames:
        return []
    ws = wb["RESULTADOS"]
    linhas = []
    for r in range(26, 31):
        valores = [ws.cell(r, c).value for c in range(1, 5)]
        if valores[0] in ("C0", "C1", "C2", "C3", "C4") and all(_numero(v) for v in valores[1:]):
            linhas.append(valores)
    return linhas


def montar_quadros_pc(composicao, memoria=None, vta=None, referencias_xls=None):
    """Seleciona valores publicados pelo motor; diferenças são demonstrativas.

    Não infere original de posição física/aditivo a partir de valor_base:
    nessas parcelas o motor publica como base um valor já atualizado.
    Saldo ausente numa composição disponível significa parcela zero, conforme
    a própria composição (que soma zero quando saldo_remanescente é None).
    A conferência abaixo apenas controla a exibição; nunca corrige o VTA.
    """
    comp = composicao or {}
    if comp.get("metodo") != "pc" or not comp.get("disponivel") or comp.get("bloqueia_formalizacao"):
        return {}
    execucao = comp.get("total_execucao_atualizada")
    saldo = (comp.get("saldo_remanescente") or {}).get("valor_atualizado", 0.0)
    potencial = comp.get("retroativo_potencial_vta")
    oficial = comp.get("vta_composicao") if vta is None else vta
    if not all(_numero(v) for v in (execucao, saldo, potencial, oficial)):
        return {}
    if abs(round(execucao + saldo + potencial - oficial, 2)) > 0.01:
        return {}
    parcelas = comp.get("execucao_por_ciclo") or []
    if any(not _numero(p.get("valor_atualizado")) for p in parcelas):
        return {}
    original = bool(parcelas) and all(
        p.get("fonte") in ("pc", "fisico") and _numero(p.get("valor_base"))
        for p in parcelas
    ) and _numero(comp.get("total_execucao_base"))
    cab_exec = ["Ciclo"]
    if original:
        cab_exec.append("Valor executado original")
    cab_exec.append("Valor executado atualizado")
    if original:
        cab_exec.append("Reajuste incorporado")
    cab_exec.append("Situação")
    linhas_exec = []
    for p in parcelas:
        linha = [p.get("ciclo", "")]
        if original:
            linha.append(p["valor_base"])
        linha.append(p["valor_atualizado"])
        if original:
            linha.append(round(p["valor_atualizado"] - p["valor_base"], 2))
        linha.append(p.get("descricao", ""))
        linhas_exec.append(linha)
    total_exec = ["TOTAL"]
    if original:
        total_exec.append(comp["total_execucao_base"])
    total_exec.append(execucao)
    if original:
        total_exec.append(round(execucao - comp["total_execucao_base"], 2))
    total_exec.append("Execução considerada no VTA")
    linhas_exec.append(total_exec)
    referencias = []
    for c in (memoria or {}).get("ciclos") or []:
        residual = c.get("residuais") or {}
        base, atualizado = residual.get("valor_original"), residual.get("valor_atualizado")
        if residual.get("itens") and _numero(base) and _numero(atualizado):
            referencias.append([c["ciclo"], base, atualizado, round(atualizado - base, 2)])
    if referencias_xls:
        referencias = [list(linha) for linha in referencias_xls]
    return {
        "execucao_cabecalho": cab_exec,
        "execucao": linhas_exec,
        "nota_execucao": NOTA_REAJUSTE if original else NOTA_EXECUCAO,
        "referencia_cabecalho": ["Ciclo", "Saldo sem reajuste de referência do ciclo", "Saldo atualizado de referência do ciclo", "Diferença de referência do ciclo"],
        "referencias": referencias,
        "saldo_final": saldo,
        "composicao_cabecalho": ["Componente", "Valor", "Observação"],
        "composicao": [
            [EXECUCAO, execucao, NOTA_EXECUCAO],
            [SALDO, saldo, "Parcela final utilizada no VTA; não somar os saldos históricos."],
            [POTENCIAL, potencial, NOTA_POTENCIAL],
            [TOTAL, oficial, "Execução atualizada + saldo final atualizado + retroativo potencial."],
        ],
    }
