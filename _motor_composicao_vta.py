"""Motor de Composicao do VTA — espelha o Quadro de memoria fiscal das apostilas.

Compoe o Valor Total Atualizado por parcelas auditaveis, linha a linha:

    VTA = soma(executado do ciclo x fator acumulado do ciclo)
        + saldo remanescente do corte (base original) x fator acumulado vigente
        + soma(aditivo/supressao: valor na assinatura x fator do ciclo-marco)

vedada a dupla contagem (JA_REFLETIDO_EM). A execucao por ciclo vem da
reconciliacao (fonte prevalente Financeiro > PC > Consumo), entao a parcela
composta ja e a base unica por ciclo, nunca soma de fontes redundantes.

Camada aditiva: nao altera vta_sombra, historico!B51, formulas oficiais nem
parsing existente. O cenario "valor original x fator acumulado" e exposto
apenas como teorico, com alerta de superestimacao — nunca como VTA.
"""
from __future__ import annotations

from typing import Any

ROTULO_CENARIO_TEORICO = (
    "Cenario teorico (valor original x fator acumulado) — superestima o VTA "
    "porque reprecifica execucao ja paga a precos antigos; nao usar como VTA."
)


def _tofl(valor: Any, default: float = 0.0) -> float:
    try:
        if valor in (None, ""):
            return default
        return float(valor)
    except (TypeError, ValueError):
        return default


def _fator_do_ciclo(por_ciclo: dict[str, Any], ciclo: str) -> float | None:
    reg = por_ciclo.get(str(ciclo or "").strip().upper()) or {}
    fator = _tofl(reg.get("fator_acumulado"), default=0.0)
    return fator or None


def _execucao_por_ciclo(
    leitura: dict[str, Any],
    por_ciclo: dict[str, Any],
    alertas: list[str],
) -> list[dict[str, Any]]:
    registros = (leitura.get("reconciliacao") or {}).get("registros") or []
    linhas: list[dict[str, Any]] = []
    for reg in registros:
        if reg.get("metodo_apuracao") == "VINCULO_EXPLICITO":
            continue
        ciclo = str(reg.get("ciclo") or "").strip().upper()
        base = _tofl(reg.get("valor_computado"))
        if not ciclo or not base:
            continue
        fator = _fator_do_ciclo(por_ciclo, ciclo)
        if fator is None:
            alertas.append(
                f"Composicao VTA: ciclo {ciclo} sem fator acumulado "
                "parametrizado; execucao composta pela base, sem atualizacao."
            )
            atualizado = round(base, 2)
        else:
            atualizado = round(base * fator, 2)
        linhas.append({
            "ciclo": ciclo,
            "descricao": (
                f"{ciclo} executado" if (fator or 1.0) == 1.0
                else f"{ciclo} executado atualizado"
            ),
            "valor_base": round(base, 2),
            "fator_acumulado": fator,
            "valor_atualizado": atualizado,
            "fonte": reg.get("fonte_principal"),
            "status_reconciliacao": reg.get("status_reconciliacao"),
            "bloqueia_formalizacao": bool(reg.get("bloqueia_formalizacao")),
        })
    return sorted(linhas, key=lambda l: l["ciclo"])


def _saldo_remanescente(
    leitura: dict[str, Any],
    alertas: list[str],
) -> dict[str, Any] | None:
    potencial = leitura.get("potencial_futuro") or {}
    saldo = _tofl(potencial.get("saldo_remanescente_base"))
    if not saldo:
        return None
    fator = _tofl(potencial.get("fator_vigente"), default=0.0) or None
    atualizado = _tofl(potencial.get("valor_atualizado_vigente"), default=0.0)
    if not atualizado:
        atualizado = round(saldo, 2)
        alertas.append(
            "Composicao VTA: saldo remanescente sem fator vigente; "
            "composto pela base, sem atualizacao."
        )
    return {
        "descricao": "Saldo remanescente atualizado no corte",
        "valor_base": round(saldo, 2),
        "fator_acumulado": fator,
        "valor_atualizado": round(atualizado, 2),
        "ciclo": potencial.get("ciclo_vigente") or "",
        "fonte": "remanescente",
    }


def _aditivos(
    leitura: dict[str, Any],
    alertas: list[str],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    visiveis = leitura.get("aditivos_visiveis") or {}
    computados: list[dict[str, Any]] = []
    nao_computados: list[dict[str, Any]] = []

    if visiveis.get("ok"):
        for item in visiveis.get("itens") or []:
            registro = {
                "descricao": item.get("evento") or "Aditivo/supressao",
                "ciclo": item.get("ciclo_marco") or "",
                "valor_base": _tofl(item.get("valor_assinatura")),
                "fator_acumulado": item.get("fator_acumulado"),
                "valor_atualizado": _tofl(
                    item.get("valor_atualizado"),
                    default=_tofl(item.get("valor_assinatura")),
                ),
                "fonte": "aditivo",
                "ja_refletido_em": item.get("ja_refletido_em"),
            }
            if str(item.get("ja_refletido_em") or "Nao").strip() not in ("", "Nao"):
                registro["motivo"] = (
                    f"Ja refletido em {item.get('ja_refletido_em')}; "
                    "nao soma ao VTA (vedada dupla contagem)."
                )
                nao_computados.append(registro)
            else:
                computados.append(registro)
        return computados, nao_computados

    # Fallback legado: aditivos consolidados da aba oculta (motor sombra),
    # que ja chegam a valor atualizado.
    parcelas = (leitura.get("vta_sombra") or {}).get("parcelas_computadas") or []
    for parcela in parcelas:
        if parcela.get("fonte_parcela") != "Aditivo":
            continue
        computados.append({
            "descricao": parcela.get("identificador") or "Aditivo (aba oculta)",
            "ciclo": "",
            "valor_base": None,
            "fator_acumulado": None,
            "valor_atualizado": _tofl(parcela.get("valor")),
            "fonte": "aditivo",
            "ja_refletido_em": "Nao",
        })
    if computados:
        alertas.append(
            "Composicao VTA: aditivos lidos do bloco oculto homologado "
            "(sem aba visivel ENTRADA_XLS_ADITIVOS); valores ja atualizados."
        )
    return computados, nao_computados


def _cenario_teorico(
    leitura: dict[str, Any],
    por_ciclo: dict[str, Any],
) -> dict[str, Any] | None:
    itens = (leitura.get("itens_contrato") or {}).get("itens") or []
    valor_original = round(sum(
        _tofl(i.get("qtd_contratada")) * _tofl(i.get("vu_original"))
        for i in itens
    ), 2)
    if not valor_original:
        return None
    ciclo_vigente = str(
        (leitura.get("controle") or {}).get("ciclo_vigente") or ""
    ).strip().upper()
    fator = _fator_do_ciclo(por_ciclo, ciclo_vigente)
    if fator is None:
        return None
    return {
        "rotulo": ROTULO_CENARIO_TEORICO,
        "valor_original": valor_original,
        "ciclo_vigente": ciclo_vigente,
        "fator_acumulado": fator,
        "valor_teorico": round(valor_original * fator, 2),
        "e_vta": False,
    }


def montar_composicao_vta(leitura: dict[str, Any]) -> dict[str, Any]:
    """Monta o quadro COMPOSICAO_VTA a partir da leitura ja consolidada.

    Requer que a leitura ja tenha reconciliacao, potencial_futuro e (quando
    existir) aditivos_visiveis — por isso roda ao final do leitor.
    """
    alertas: list[str] = []
    por_ciclo = (leitura.get("parametros_v10") or {}).get("por_ciclo") or {}

    execucao = _execucao_por_ciclo(leitura, por_ciclo, alertas)
    saldo = _saldo_remanescente(leitura, alertas)
    aditivos, aditivos_fora = _aditivos(leitura, alertas)

    saida: dict[str, Any] = {
        "disponivel": False,
        "motivo": "",
        "execucao_por_ciclo": execucao,
        "saldo_remanescente": saldo,
        "aditivos": aditivos,
        "aditivos_nao_computados": aditivos_fora,
        "total_execucao_atualizada": 0.0,
        "total_execucao_base": 0.0,
        "retroativo_implicito": 0.0,
        "total_aditivos_atualizados": 0.0,
        "vta_composicao": None,
        "cenario_teorico": _cenario_teorico(leitura, por_ciclo),
        "bloqueia_formalizacao": any(
            l.get("bloqueia_formalizacao") for l in execucao
        ),
        "linhas": [],
        "alertas": alertas,
    }

    if not execucao and saldo is None and not aditivos:
        saida["motivo"] = (
            "Sem execucao reconciliada, saldo remanescente ou aditivos; "
            "nada a compor."
        )
        return saida

    total_exec = round(sum(l["valor_atualizado"] for l in execucao), 2)
    total_base = round(sum(l["valor_base"] for l in execucao), 2)
    total_aditivos = round(sum(a["valor_atualizado"] for a in aditivos), 2)
    valor_saldo = round(saldo["valor_atualizado"], 2) if saldo else 0.0

    saida.update({
        "disponivel": True,
        "total_execucao_atualizada": total_exec,
        "total_execucao_base": total_base,
        "retroativo_implicito": round(total_exec - total_base, 2),
        "total_aditivos_atualizados": total_aditivos,
        "vta_composicao": round(total_exec + valor_saldo + total_aditivos, 2),
    })

    linhas = [dict(l, tipo="execucao") for l in execucao]
    if saldo:
        linhas.append(dict(saldo, tipo="saldo_remanescente"))
    linhas.extend(dict(a, tipo="aditivo") for a in aditivos)
    for idx, linha in enumerate(linhas):
        linha["ref"] = chr(ord("A") + idx) if idx < 26 else str(idx + 1)
    saida["linhas"] = linhas

    if saida["bloqueia_formalizacao"]:
        alertas.append(
            "Composicao VTA: ha ciclo DIVERGENTE na reconciliacao; o VTA "
            "composto usa o valor prevalente e a formalizacao segue bloqueada "
            "ate decisao da GCC."
        )
    return saida
