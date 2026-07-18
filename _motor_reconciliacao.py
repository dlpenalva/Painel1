"""Motor de Reconciliacao (VTA) — camada paralela e auditavel.

Implementa a RFC docs/RFC_MOTOR_RECONCILIACAO.md:
- hierarquia de prevalencia Financeiro > PC > Consumo;
- grao ciclo (bases agregadas; fator do ciclo e unico);
- soft-block com escalonamento por faixa de divergencia;
- evidencia nao aditiva rotulada CORROBORANTE ("Evidencia corroborante").

Nao altera vta_sombra, historico!B51 nem parsing existente.
"""
from __future__ import annotations

from typing import Any

# Ponto unico de configuracao (RFC §5). O valor aplicado e sempre gravado
# no registro (tolerancia_aplicada) para rastreabilidade.
TOLERANCIA_ABSOLUTA = 0.05
TOLERANCIA_RELATIVA = 0.005
LIMIAR_SOFT_BLOCK = 0.05

HIERARQUIA_PREVALENCIA = ["financeiro", "pc", "consumo"]

ROTULO_CORROBORANTE = "Evidencia corroborante (nao soma ao VTA)"

_FONTES_FINANCEIRO = {"Financeiro", "Historico financeiro"}


def _tofl(valor: Any, default: float = 0.0) -> float:
    try:
        if valor in (None, ""):
            return default
        return float(valor)
    except (TypeError, ValueError):
        return default


def _ciclo_da_parcela(parcela: dict[str, Any]) -> str:
    ciclo = str(parcela.get("ciclo") or "").strip().upper()
    if ciclo:
        return ciclo
    # Fallback: parcelas da aba oculta carregam o ciclo no identificador
    # (ex.: "financeiro:C1:base:3").
    ident = str(parcela.get("identificador") or "")
    for parte in ident.split(":"):
        parte = parte.strip().upper()
        if parte in {f"C{i}" for i in range(5)}:
            return parte
    return ""


def _totais_por_fonte_e_ciclo(leitura: dict[str, Any]) -> tuple[dict[str, dict[str, float]], list[dict[str, Any]]]:
    """Agrega bases por ciclo: financeiro (parcelas), pc (valor_pc), consumo."""
    totais: dict[str, dict[str, float]] = {"financeiro": {}, "pc": {}, "consumo": {}}
    vinculos_explicitos: list[dict[str, Any]] = []

    vta = leitura.get("vta_sombra") or {}
    for parcela in vta.get("parcelas_computadas") or []:
        if parcela.get("fonte_parcela") not in _FONTES_FINANCEIRO:
            continue
        ciclo = _ciclo_da_parcela(parcela)
        if not ciclo:
            continue
        valor = _tofl(parcela.get("valor"))
        totais["financeiro"][ciclo] = round(
            totais["financeiro"].get(ciclo, 0.0) + valor, 2
        )

    for item in (leitura.get("itens_pc_v10") or {}).get("itens") or []:
        campos = item.get("campos_vta") or {}
        ciclo = str(item.get("ciclo") or item.get("ciclo_pc") or "").strip().upper()
        valor = _tofl(item.get("valor_pc"))
        if not ciclo or not valor:
            continue
        refletido = str(campos.get("ja_refletido_em") or "Nao").strip()
        if refletido not in ("", "Nao"):
            vinculos_explicitos.append({
                "ciclo": ciclo,
                "identificador": item.get("numero_pc") or item.get("item_ou_grupo"),
                "valor": valor,
                "refletido_em": refletido,
            })
            continue
        totais["pc"][ciclo] = round(totais["pc"].get(ciclo, 0.0) + valor, 2)

    for item in (leitura.get("itens_consumidos_v10") or {}).get("itens") or []:
        ciclo = str(item.get("ciclo_inferido") or "").strip().upper()
        valor = _tofl(item.get("valor_total"))
        if not ciclo or not valor:
            continue
        totais["consumo"][ciclo] = round(totais["consumo"].get(ciclo, 0.0) + valor, 2)

    return totais, vinculos_explicitos


def _tolerancia(valor_principal: float) -> float:
    return round(max(TOLERANCIA_ABSOLUTA, abs(valor_principal) * TOLERANCIA_RELATIVA), 2)


def _classificar(diferenca: float, valor_principal: float, tolerancia: float) -> str:
    if diferenca <= tolerancia:
        return "RECONCILIADO"
    base = abs(valor_principal) or 1.0
    if diferenca / base <= LIMIAR_SOFT_BLOCK:
        return "RECONCILIADO_COM_ALERTA"
    return "DIVERGENTE"


def reconciliar_execucoes(
    leitura: dict[str, Any],
    decisoes: dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Monta os registros de reconciliacao por ciclo (RFC §6).

    ``decisoes`` mapeia registro_id -> decisao do log apartado; quando
    presente, o registro DIVERGENTE vira RESOLVIDO_GCC e o soft-block cai.
    """
    decisoes = decisoes or {}
    totais, vinculos = _totais_por_fonte_e_ciclo(leitura)
    ciclos = sorted(
        {c for fonte in totais.values() for c in fonte}
    )

    registros: list[dict[str, Any]] = []
    alertas: list[str] = []
    resumo = {
        "RECONCILIADO": 0, "RECONCILIADO_COM_ALERTA": 0,
        "DIVERGENTE": 0, "INCONCLUSIVO": 0, "RESOLVIDO_GCC": 0,
    }
    total_computado = 0.0
    total_em_analise = 0.0

    for ciclo in ciclos:
        valores = {
            fonte: totais[fonte][ciclo]
            for fonte in HIERARQUIA_PREVALENCIA
            if ciclo in totais[fonte]
        }
        fonte_principal = next(
            f for f in HIERARQUIA_PREVALENCIA if f in valores
        )
        valor_principal = valores[fonte_principal]
        tolerancia = _tolerancia(valor_principal)

        secundarias = []
        maior_diferenca = 0.0
        for fonte in HIERARQUIA_PREVALENCIA:
            if fonte == fonte_principal or fonte not in valores:
                continue
            diferenca = round(valores[fonte] - valor_principal, 2)
            maior_diferenca = max(maior_diferenca, abs(diferenca))
            secundarias.append({
                "fonte": fonte,
                "papel_reconciliacao": "CORROBORANTE",
                "nao_aditiva": True,
                "rotulo": ROTULO_CORROBORANTE,
                "valor": valores[fonte],
                "diferenca": diferenca,
            })

        if not secundarias:
            status = "INCONCLUSIVO"
            metodo = f"FONTE_UNICA_{fonte_principal.upper()}"
        else:
            status = _classificar(maior_diferenca, valor_principal, tolerancia)
            metodo = f"PREVALENCIA_{fonte_principal.upper()}"

        registro_id = f"reconciliacao:{ciclo}"
        valor_em_analise = (
            round(maior_diferenca, 2) if status == "DIVERGENTE" else 0.0
        )
        decisao = decisoes.get(registro_id)
        if decisao and status == "DIVERGENTE":
            status = "RESOLVIDO_GCC"
            valor_em_analise = 0.0

        registro = {
            "id": registro_id,
            "ciclo": ciclo,
            "valores_por_fonte": valores,
            "fonte_principal": fonte_principal,
            "fontes_secundarias": secundarias,
            "criterio_selecao": (
                "agregacao por ciclo (fator unico); "
                "hierarquia Financeiro > PC > Consumo"
            ),
            "status_reconciliacao": status,
            "diferenca_identificada": round(maior_diferenca, 2) if secundarias else None,
            "metodo_apuracao": metodo,
            "tolerancia_aplicada": tolerancia,
            "valor_computado": valor_principal,
            "valor_em_analise": valor_em_analise,
            "bloqueia_formalizacao": status == "DIVERGENTE",
            "decisao_gcc": decisao,
        }
        registros.append(registro)
        resumo[status] += 1
        total_computado += valor_principal
        total_em_analise += valor_em_analise

        if status == "RECONCILIADO_COM_ALERTA":
            alertas.append(
                f"Reconciliacao {ciclo}: diferenca de R$ {maior_diferenca:,.2f} "
                f"entre fontes (acima da tolerancia de R$ {tolerancia:,.2f}, "
                "dentro do limiar); prevalencia aplicada."
            )
        elif status == "DIVERGENTE":
            alertas.append(
                f"Reconciliacao {ciclo}: DIVERGENTE — diferenca de "
                f"R$ {maior_diferenca:,.2f} acima de {LIMIAR_SOFT_BLOCK:.0%}; "
                "valor prevalente computado e diferenca em analise "
                "(soft-block ate decisao da GCC)."
            )

    for vinculo in vinculos:
        registro_id = f"reconciliacao:vinculo:{vinculo['identificador']}"
        registros.append({
            "id": registro_id,
            "ciclo": vinculo["ciclo"],
            "valores_por_fonte": {"pc": vinculo["valor"]},
            "fonte_principal": str(vinculo["refletido_em"]).strip().lower(),
            "fontes_secundarias": [{
                "fonte": "pc",
                "papel_reconciliacao": "CORROBORANTE",
                "nao_aditiva": True,
                "rotulo": ROTULO_CORROBORANTE,
                "valor": vinculo["valor"],
                "diferenca": None,
            }],
            "criterio_selecao": "vinculo explicito (JA_REFLETIDO_EM)",
            "status_reconciliacao": "RECONCILIADO",
            "diferenca_identificada": None,
            "metodo_apuracao": "VINCULO_EXPLICITO",
            "tolerancia_aplicada": None,
            "valor_computado": 0.0,
            "valor_em_analise": 0.0,
            "bloqueia_formalizacao": False,
            "decisao_gcc": decisoes.get(registro_id),
        })
        resumo["RECONCILIADO"] += 1

    return {
        "ok": bool(registros),
        "registros": registros,
        "resumo": resumo,
        "total_computado_prevalencia": round(total_computado, 2),
        "total_em_analise": round(total_em_analise, 2),
        "bloqueia_formalizacao": any(
            r["bloqueia_formalizacao"] for r in registros
        ),
        "alertas": alertas,
        "configuracao": {
            "tolerancia_absoluta": TOLERANCIA_ABSOLUTA,
            "tolerancia_relativa": TOLERANCIA_RELATIVA,
            "limiar_soft_block": LIMIAR_SOFT_BLOCK,
            "hierarquia": list(HIERARQUIA_PREVALENCIA),
        },
    }
