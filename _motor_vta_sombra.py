"""Motor sombra de consolidacao do VTA por parcelas.

Fase 3 v10.5.3: calcula uma composicao paralela para auditoria, sem alterar
formulas oficiais do XLS nem o VTA lido da aba historico.
"""
from __future__ import annotations

from typing import Any


TIPOS_FINANCEIROS_COMPUTAVEIS = {
    "Execucao Atualizada",
    "Saldo Remanescente",
    "Retroativo Reconhecido",
    "Aditivo Computavel",
}

# Triangulacao: cada fonte_parcela pertence a um grupo de evidencia.
# "comum" entra em todos os metodos (saldo remanescente e aditivos valem
# tanto para o caminho financeiro quanto para o caminho por PC).
_GRUPO_POR_FONTE = {
    "Financeiro": "financeiro",
    "Historico financeiro": "financeiro",
    "PC": "pc",
    "Itens consumidos": "estoque",
    "Itens remanescentes": "comum",
    "Aditivo": "comum",
}

LIMIAR_DIVERGENCIA_TRIANGULACAO = 0.05


def _to_float(valor: Any, default: float = 0.0) -> float:
    try:
        if valor in (None, ""):
            return default
        return float(valor)
    except (TypeError, ValueError):
        return default


def _valor_parcela_pc(item: dict[str, Any]) -> float:
    efeito = str(item.get("efeito_financeiro_pc") or "").strip()
    valor_pc = _to_float(item.get("valor_pc"))
    if efeito == "Nao":
        return valor_pc
    if efeito != "Sim":
        return 0.0
    valor_atualizado = _to_float(item.get("valor_atualizado"))
    if valor_atualizado:
        return valor_atualizado

    fator = _to_float(item.get("fator_acumulado"), default=1.0) or 1.0
    return round(valor_pc * fator, 2)


def _parcela_base(item: dict[str, Any], valor: float) -> dict[str, Any]:
    campos = item.get("campos_vta") or {}
    return {
        "linha": item.get("linha"),
        "identificador": item.get("numero_pc") or item.get("item_ou_grupo"),
        "origem_dado": campos.get("origem_dado"),
        "tipo_parcela": campos.get("tipo_parcela"),
        "tipo_financeiro": campos.get("tipo_financeiro"),
        "fonte_parcela": campos.get("fonte_parcela"),
        "computa_vta": campos.get("computa_vta"),
        "ja_refletido_em": campos.get("ja_refletido_em"),
        "status_consolidacao": campos.get("status_consolidacao"),
        "justificativa_vta": campos.get("justificativa_vta"),
        "valor": valor,
    }


def _parcela_generica(item: dict[str, Any]) -> dict[str, Any]:
    return {
        "linha": item.get("linha"),
        "identificador": item.get("identificador"),
        "ciclo": item.get("ciclo"),
        "data_referencia": item.get("data_referencia"),
        "confianca": item.get("confianca"),
        "origem_dado": item.get("origem_dado"),
        "tipo_parcela": item.get("tipo_parcela"),
        "tipo_financeiro": item.get("tipo_financeiro"),
        "fonte_parcela": item.get("fonte_parcela"),
        "computa_vta": item.get("computa_vta", "Sim"),
        "ja_refletido_em": item.get("ja_refletido_em", "Nao"),
        "status_consolidacao": item.get("status_consolidacao", "COMPUTADO"),
        "justificativa_vta": item.get("justificativa_vta", ""),
        "valor": _to_float(item.get("valor")),
        # Composicao do VTA: devido = base x fator do ciclo (quando materializado
        # pelo leitor); a base segue em "valor" para retroativo e reconciliacao.
        "fator_acumulado": item.get("fator_acumulado"),
        "valor_atualizado": item.get("valor_atualizado"),
        "motivo": item.get("motivo", "Parcela base do VTA sombra integral."),
    }


def _inconsistencia(campos: dict[str, str], valor: float) -> str:
    computa = campos.get("computa_vta")
    status = campos.get("status_consolidacao")
    tipo_fin = campos.get("tipo_financeiro")

    if valor <= 0:
        return "Valor da parcela ausente ou nao positivo."
    if status == "INCONSISTENTE":
        return "STATUS_CONSOLIDACAO informado como INCONSISTENTE."
    if status == "COMPUTADO" and computa != "Sim":
        return "STATUS_CONSOLIDACAO=COMPUTADO sem COMPUTA_VTA=Sim."
    if status == "COMPUTADO" and tipo_fin not in TIPOS_FINANCEIROS_COMPUTAVEIS:
        return f"TIPO_FINANCEIRO={tipo_fin} nao pode ser computado."
    if computa == "Sim" and status not in ("COMPUTADO", "DESCARTADO_DUPLICIDADE"):
        return f"COMPUTA_VTA=Sim com STATUS_CONSOLIDACAO={status}."
    return ""


def triangular_vta_por_fonte(vta_sombra: dict[str, Any]) -> dict[str, Any]:
    """Triangula o VTA sombra agrupando as parcelas computadas por fonte.

    Nao recalcula parcela nenhuma: apenas reagrupa o que o motor sombra ja
    computou, compondo o VTA por dois caminhos independentes de evidencia:

    - metodo financeiro: parcelas de Financeiro + Historico financeiro;
    - metodo PC: parcelas de PC;

    ambos somados as parcelas comuns (saldo remanescente e aditivos), que
    valem para qualquer caminho. Consumo por estoque, quando materializado
    como parcela, forma um terceiro subtotal informativo.
    """
    parcelas = (vta_sombra or {}).get("parcelas_computadas") or []
    por_fonte: dict[str, dict[str, Any]] = {}
    por_grupo = {"financeiro": 0.0, "pc": 0.0, "estoque": 0.0, "comum": 0.0, "outros": 0.0}

    for parcela in parcelas:
        fonte = str(parcela.get("fonte_parcela") or "Nao informado").strip() or "Nao informado"
        valor = _to_float(parcela.get("valor"))
        registro = por_fonte.setdefault(fonte, {"parcelas": 0, "valor": 0.0})
        registro["parcelas"] += 1
        registro["valor"] = round(registro["valor"] + valor, 2)
        por_grupo[_GRUPO_POR_FONTE.get(fonte, "outros")] += valor

    comum = por_grupo["comum"]
    tem_financeiro = por_grupo["financeiro"] > 0.0
    tem_pc = por_grupo["pc"] > 0.0
    vta_financeiro = round(por_grupo["financeiro"] + comum, 2) if tem_financeiro else None
    vta_pc = round(por_grupo["pc"] + comum, 2) if tem_pc else None

    saida: dict[str, Any] = {
        "por_fonte": por_fonte,
        "vta_metodo_financeiro": vta_financeiro,
        "vta_metodo_pc": vta_pc,
        "subtotal_estoque": round(por_grupo["estoque"], 2) or None,
        "subtotal_comum": round(comum, 2),
        "subtotal_outros": round(por_grupo["outros"], 2) or None,
        "comparavel": tem_financeiro and tem_pc,
        "delta_conciliacao": None,
        "divergencia_pct": None,
        "alertas": [],
    }

    if saida["comparavel"]:
        delta = round(vta_pc - vta_financeiro, 2)
        base = max(abs(vta_financeiro), abs(vta_pc)) or 1.0
        divergencia = abs(delta) / base
        saida["delta_conciliacao"] = delta
        saida["divergencia_pct"] = round(divergencia, 4)
        if divergencia > LIMIAR_DIVERGENCIA_TRIANGULACAO:
            saida["alertas"].append(
                f"Triangulacao: VTA por PC diverge {divergencia:.1%} do VTA "
                f"financeiro (delta {delta:+.2f}); conciliar fontes antes de "
                "formalizar."
            )
    elif tem_financeiro or tem_pc:
        metodo = "financeiro" if tem_financeiro else "PC"
        saida["alertas"].append(
            f"Triangulacao: apenas o metodo {metodo} tem evidencia computada; "
            "sem segunda fonte para conciliacao."
        )

    return saida


def calcular_vta_sombra(
    resumo_oficial: dict[str, Any] | None,
    itens_pc: dict[str, Any] | None,
    parcelas_base: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Calcula VTA sombra a partir das parcelas oficiais e de itens_PC.

    Sem parcelas_base, preserva a Fase 3: parte do VTA oficial e acrescenta PC
    autorizado. Com parcelas_base, opera em modo integral: soma as parcelas
    oficiais materializadas e compara com o VTA oficial.
    """
    resumo_oficial = resumo_oficial or {}
    itens_pc = itens_pc or {}
    parcelas_base = parcelas_base or []
    modo_integral = bool(parcelas_base)
    oficial = _to_float(resumo_oficial.get("valor_total_atualizado"))

    saida: dict[str, Any] = {
        "modo": "sombra_integral" if modo_integral else "sombra",
        "vta_oficial": oficial,
        "vta_sombra": 0.0 if modo_integral else oficial,
        "diferenca": 0.0,
        "parcelas_computadas": [],
        "parcelas_nao_computadas": [],
        "parcelas_descartadas_duplicidade": [],
        "impacto_potencial": [],
        "inconsistencias": [],
        "alertas": [],
        "resumo": {
            "computadas": 0,
            "nao_computadas": 0,
            "descartadas_duplicidade": 0,
            "impacto_potencial": 0,
            "inconsistencias": 0,
        },
    }

    if resumo_oficial.get("valor_total_atualizado") in (None, ""):
        saida["alertas"].append(
            "VTA oficial ausente no historico; comparacao sombra usa 0 como base."
        )

    for item in parcelas_base:
        parcela = _parcela_generica(item)
        valor = _to_float(parcela.get("valor"))
        if valor <= 0:
            parcela["motivo"] = "Parcela base ausente, zerada ou negativa."
            saida["parcelas_nao_computadas"].append(parcela)
            continue
        if parcela.get("ja_refletido_em") != "Nao":
            parcela["motivo"] = (
                f"Parcela base ja refletida em {parcela.get('ja_refletido_em')}."
            )
            saida["parcelas_descartadas_duplicidade"].append(parcela)
            continue
        if parcela.get("computa_vta") != "Sim":
            parcela["motivo"] = "Parcela base marcada para nao computar."
            saida["parcelas_nao_computadas"].append(parcela)
            continue
        if parcela.get("status_consolidacao") == "INCONSISTENTE":
            parcela["motivo"] = "Parcela base inconsistente."
            saida["inconsistencias"].append(parcela)
            continue

        parcela["motivo"] = parcela.get("motivo") or "Parcela base computada."
        saida["parcelas_computadas"].append(parcela)
        saida["vta_sombra"] += valor

    for item in itens_pc.get("itens") or []:
        campos = item.get("campos_vta") or {}
        valor = _valor_parcela_pc(item)
        parcela = _parcela_base(item, valor)

        motivo_inconsistencia = _inconsistencia(campos, valor)
        if motivo_inconsistencia:
            parcela["motivo"] = motivo_inconsistencia
            saida["inconsistencias"].append(parcela)
            saida["alertas"].append(
                f"PC {parcela.get('identificador')!r}: {motivo_inconsistencia}"
            )
            continue

        if campos.get("ja_refletido_em") != "Nao":
            parcela["motivo"] = (
                f"Parcela ja refletida em {campos.get('ja_refletido_em')}."
            )
            saida["parcelas_descartadas_duplicidade"].append(parcela)
            continue

        if campos.get("computa_vta") != "Sim":
            if campos.get("tipo_financeiro") == "Impacto Potencial" or (
                campos.get("status_consolidacao") == "EM_ANALISE"
            ):
                parcela["motivo"] = "Impacto potencial ou parcela em analise."
                saida["impacto_potencial"].append(parcela)
            else:
                parcela["motivo"] = "COMPUTA_VTA diferente de Sim."
                saida["parcelas_nao_computadas"].append(parcela)
            continue

        if campos.get("tipo_financeiro") == "Impacto Potencial":
            parcela["motivo"] = "Impacto potencial nao compoe VTA sombra."
            saida["impacto_potencial"].append(parcela)
            continue

        if campos.get("tipo_financeiro") not in TIPOS_FINANCEIROS_COMPUTAVEIS:
            parcela["motivo"] = (
                f"TIPO_FINANCEIRO={campos.get('tipo_financeiro')} nao computavel."
            )
            saida["parcelas_nao_computadas"].append(parcela)
            continue

        parcela["motivo"] = "Parcela autorizada no motor sombra."
        saida["parcelas_computadas"].append(parcela)
        saida["vta_sombra"] += valor

    saida["vta_sombra"] = round(saida["vta_sombra"], 2)
    saida["diferenca"] = round(saida["vta_sombra"] - saida["vta_oficial"], 2)
    saida["resumo"] = {
        "computadas": len(saida["parcelas_computadas"]),
        "nao_computadas": len(saida["parcelas_nao_computadas"]),
        "descartadas_duplicidade": len(saida["parcelas_descartadas_duplicidade"]),
        "impacto_potencial": len(saida["impacto_potencial"]),
        "inconsistencias": len(saida["inconsistencias"]),
    }
    saida["triangulacao"] = triangular_vta_por_fonte(saida)
    saida["alertas"].extend(saida["triangulacao"]["alertas"])
    return saida
