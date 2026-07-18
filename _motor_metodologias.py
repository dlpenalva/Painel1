"""Motor de Metodologias do Claus New.

Camada de decisao de produto, somente leitura: identifica evidencias,
calcula alternativas comparaveis, detecta risco de dupla contagem e entrega
uma unica metodologia recomendada para confirmacao operacional.
"""
from __future__ import annotations

import copy
from typing import Any


MET_FINANCEIRO = "Financeiro"
MET_PC = "Pedidos de Compra (PC)"
MET_CONSUMIDOS = "Itens Consumidos"
MET_FIN_REMANESC = "Financeiro + Remanescentes"

CONF_ALTA = "alta"
CONF_MEDIA = "media"
CONF_BAIXA = "baixa"
CONF_NULA = "insuficiente"

ACAO_CONFIRMAR = "Confirmar metodologia recomendada"
ACAO_SOLICITAR_EVIDENCIA = "Solicitar evidencia de execucao"
ACAO_REVISAR_DUPLA = "Revisar possivel dupla contagem"
ACAO_REVISAR_PENDENCIAS = "Revisar pendencias antes de confirmar"


def _f(valor: Any) -> float:
    try:
        return float(valor or 0.0)
    except (TypeError, ValueError):
        return 0.0


def _round(valor: Any) -> float:
    return round(_f(valor), 2)


def _bool_sim(valor: Any) -> bool:
    return str(valor or "").strip().lower() in {"sim", "s", "true", "1", "yes"}


def _eventos(leitura: dict[str, Any]) -> list[dict[str, Any]]:
    return list((leitura.get("event_log_sombra") or {}).get("eventos") or [])


def _pcs(painel: dict[str, Any]) -> list[dict[str, Any]]:
    return list((painel.get("situacao_pcs") or {}).get("pcs") or [])


def _resumo_oficial(painel: dict[str, Any]) -> dict[str, Any]:
    return dict((painel.get("situacao_financeira") or {}).get("resumo_oficial") or {})


def _totais_motor(painel: dict[str, Any]) -> dict[str, Any]:
    return dict((painel.get("situacao_financeira") or {}).get("totais") or {})


def _valor_consumidos(leitura: dict[str, Any]) -> float:
    consumidos = leitura.get("itens_consumidos_v10") or {}
    totais = consumidos.get("totais") if isinstance(consumidos, dict) else {}
    for chave in ("valor_total", "valor_total_atualizado", "valor"):
        if totais and totais.get(chave) is not None:
            return _round(totais.get(chave))
    total = 0.0
    for item in (consumidos.get("itens") if isinstance(consumidos, dict) else []) or []:
        valor = None
        for chave in ("valor_total", "valor_total_atualizado", "valor_consumido", "valor"):
            if item.get(chave) is not None:
                valor = item.get(chave)
                break
        total += _f(valor)
    return _round(total)


def _valor_financeiro(leitura: dict[str, Any]) -> dict[str, float]:
    total = 0.0
    reconhecido = 0.0
    pago = 0.0
    for ev in _eventos(leitura):
        fonte = str(ev.get("fonte_parcela") or ev.get("origem_dado") or "").strip().lower()
        if fonte != "financeiro":
            continue
        valor = _f(ev.get("valor"))
        total += valor
        tipo = str(ev.get("tipo_financeiro") or "").lower()
        if "reconhecido" in tipo:
            reconhecido += valor
        if "execucao" in tipo or "pago" in tipo or "pag" in tipo:
            pago += valor
    return {
        "total": _round(total),
        "reconhecido": _round(reconhecido),
        "pago": _round(pago),
    }


def _evidencias(leitura: dict[str, Any], painel: dict[str, Any]) -> dict[str, Any]:
    eventos = _eventos(leitura)
    pcs = _pcs(painel)
    consumidos = (leitura.get("itens_consumidos_v10") or {}).get("itens") or []
    exec_saldo = (leitura.get("execucao_saldo") or {}).get("itens") or []
    financeiro = _valor_financeiro(leitura)

    parcelas_financeiro = 0
    parcelas_remanescentes = 0
    parcelas_aditivos = 0
    for ev in eventos:
        fonte = str(ev.get("fonte_parcela") or ev.get("origem_dado") or "").strip().lower()
        if fonte == "financeiro":
            parcelas_financeiro += 1
        elif "remanescente" in fonte:
            parcelas_remanescentes += 1
        elif fonte == "aditivo":
            parcelas_aditivos += 1

    return {
        "financeiro": {
            "presente": parcelas_financeiro > 0,
            "parcelas": parcelas_financeiro,
            "valor": financeiro["total"],
            "reconhecido": financeiro["reconhecido"],
            "pago": financeiro["pago"],
        },
        "remanescentes": {
            "presente": parcelas_remanescentes > 0 or len(exec_saldo) > 0,
            "parcelas": parcelas_remanescentes,
            "itens_execucao_saldo": len(exec_saldo),
            "valor": _round(_resumo_oficial(painel).get("saldo_remanescente")),
        },
        "pcs": {
            "presente": len(pcs) > 0,
            "quantidade": len(pcs),
            "sem_data": sum(1 for pc in pcs if not pc.get("data_pc")),
            "pagos": sum(1 for pc in pcs if _f(pc.get("valor_pago")) > 0.0),
            "pagos_elegiveis": sum(
                1 for pc in pcs
                if pc.get("elegivel_retroativo_pc") is not False
                and _f(pc.get("valor_pago")) > 0.0
            ),
            "valor": _round(sum(_f(pc.get("valor_pc")) for pc in pcs)),
            "devido": _round(sum(_f(pc.get("valor_devido")) for pc in pcs)),
            "pago": _round(sum(_f(pc.get("valor_pago")) for pc in pcs)),
            "executado_atualizado_pago": _round(sum(
                _f(pc.get("valor_pago")) + _f(pc.get("retroativo"))
                for pc in pcs
                if pc.get("elegivel_retroativo_pc") is not False
                and _f(pc.get("valor_pago")) > 0.0
            )),
            "em_analise": _round(sum(
                _f(pc.get("retroativo")) for pc in pcs
                if str(pc.get("natureza_delta") or "") == "potencial"
            )),
            "reconhecido": _round(sum(
                _f(pc.get("retroativo")) for pc in pcs
                if str(pc.get("natureza_delta") or "") == "reconhecido"
            )),
            "origens_retroativo": [
                {
                    "numero_pc": pc.get("numero_pc"),
                    "ciclo": pc.get("ciclo_temporal"),
                    "valor": _round(pc.get("retroativo")),
                    "classificacao": "execucao",
                }
                for pc in pcs if _f(pc.get("retroativo")) > 0.0
            ],
        },
        "consumidos": {
            "presente": len(consumidos) > 0,
            "itens": len(consumidos),
            "valor": _valor_consumidos(leitura),
        },
        "aditivos": {
            "presentes_no_event_log": parcelas_aditivos,
            "observacao": "PC nao e aditivo; PC entra apenas como execucao.",
        },
    }


def _detectar_dupla_contagem(leitura: dict[str, Any], painel: dict[str, Any]) -> dict[str, Any]:
    itens: list[dict[str, Any]] = []
    vta = leitura.get("vta_sombra") or {}
    for p in vta.get("parcelas_descartadas_duplicidade") or []:
        itens.append({
            "origem": p.get("identificador") or p.get("numero_pc") or "parcela",
            "valor": _round(p.get("valor")),
            "motivo": "Parcela ja refletida em outra fonte; excluida da conferencia sombra.",
        })

    for ev in _eventos(leitura):
        if str(ev.get("tipo_evento") or "") == "PC_DESCARTADO_DUPLICIDADE":
            itens.append({
                "origem": ev.get("identificador") or ev.get("numero_pc") or "PC",
                "valor": _round(ev.get("valor")),
                "motivo": "PC descartado da conferencia por ja estar refletido em outra fonte.",
            })
        refletido = str(ev.get("ja_refletido_em") or "").strip()
        if refletido and refletido.lower() not in {"nao", "não", "n", "0", "false"}:
            itens.append({
                "origem": ev.get("identificador") or ev.get("numero_pc") or "parcela",
                "valor": _round(ev.get("valor")),
                "motivo": f"Declarado como ja refletido em {refletido}.",
            })

    for alerta in painel.get("alertas") or []:
        texto = (str(alerta.get("codigo") or "") + " " +
                 str(alerta.get("mensagem_negocio") or alerta.get("mensagem") or ""))
        if "duplic" in texto.lower() or "ja_refletido" in texto.lower():
            itens.append({
                "origem": alerta.get("identificador") or "alerta",
                "valor": 0.0,
                "motivo": str(alerta.get("mensagem_negocio") or alerta.get("mensagem") or ""),
            })

    vistos: set[tuple[str, float, str]] = set()
    unicos: list[dict[str, Any]] = []
    for item in itens:
        chave = (str(item["origem"]), _round(item["valor"]), str(item["motivo"]))
        if chave not in vistos:
            vistos.add(chave)
            unicos.append(item)

    return {
        "existe": bool(unicos),
        "quantidade": len(unicos),
        "itens": unicos,
        "acao": ACAO_REVISAR_DUPLA if unicos else "",
    }


def _pendencias(painel: dict[str, Any], dupla: dict[str, Any]) -> list[str]:
    pendencias: list[str] = []
    for alerta in painel.get("alertas") or []:
        nivel = str(alerta.get("nivel") or "").upper()
        if nivel in {"ERRO_GRAVE", "ALERTA"}:
            msg = str(alerta.get("mensagem_negocio") or alerta.get("mensagem") or "").strip()
            if msg:
                pendencias.append(msg)
    if dupla.get("existe"):
        pendencias.append("Revisar possivel dupla contagem antes de formalizar valores.")
    return pendencias


def _alternativa(
    nome: str,
    disponivel: bool,
    score: int,
    confiabilidade: str,
    valor_recomendado: float | None,
    valores: dict[str, Any],
    justificativa: str,
    fonte: str,
    riscos: list[str] | None = None,
) -> dict[str, Any]:
    return {
        "nome": nome,
        "disponivel": bool(disponivel),
        "score": score if disponivel else 0,
        "confiabilidade": confiabilidade if disponivel else CONF_NULA,
        "valor_recomendado": valor_recomendado if disponivel else None,
        "valores": valores if disponivel else {},
        "justificativa": justificativa,
        "fonte": fonte,
        "riscos": list(riscos or []),
    }


def _montar_alternativas(
    ev: dict[str, Any],
    painel: dict[str, Any],
    dupla: dict[str, Any],
    pendencias: list[str],
) -> list[dict[str, Any]]:
    resumo = _resumo_oficial(painel)
    totais = _totais_motor(painel)
    vta = painel.get("vta") or {}
    tem_pendencia = bool(pendencias)
    risco_dupla = ["Ha risco de dupla contagem a conferir."] if dupla.get("existe") else []
    riscos_gerais = risco_dupla + (["Ha pendencias no painel."] if tem_pendencia else [])

    fin = ev["financeiro"]
    rem = ev["remanescentes"]
    pcs = ev["pcs"]
    consumidos = ev["consumidos"]

    return [
        _alternativa(
            MET_FIN_REMANESC,
            fin["presente"] and rem["presente"],
            110 - (15 if dupla.get("existe") else 0) - (10 if tem_pendencia else 0),
            CONF_ALTA,
            _round((fin["pago"] or fin["valor"]) + rem["valor"]),
            {
                "valor_em_analise": pcs["em_analise"],
                "valor_reconhecido": fin["reconhecido"],
                "valor_pago": fin["pago"],
                "saldo_remanescente": rem["valor"],
                "vta": _round(vta.get("oficial")),
            },
            "Usa financeiro para o que ja ocorreu e remanescente para a parcela futura.",
            "Event Log sombra + resumo oficial + execucao_saldo.",
            riscos_gerais,
        ),
        _alternativa(
            MET_FINANCEIRO,
            fin["presente"],
            100 - (15 if dupla.get("existe") else 0) - (10 if tem_pendencia else 0),
            CONF_ALTA,
            fin["valor"],
            {
                "valor_em_analise": 0.0,
                "valor_reconhecido": fin["reconhecido"],
                "valor_pago": fin["pago"],
                "vta": _round(vta.get("oficial")),
            },
            "Usa a execucao financeira como evidencia mais forte de valor realizado.",
            "Event Log sombra derivado da aba financeiro.",
            riscos_gerais,
        ),
        _alternativa(
            MET_PC,
            pcs["pagos_elegiveis"] > 0,
            70 - (20 if pcs["sem_data"] else 0) - (15 if dupla.get("existe") else 0)
            - (10 if tem_pendencia else 0),
            CONF_MEDIA if not pcs["sem_data"] else CONF_BAIXA,
            pcs["executado_atualizado_pago"],
            {
                "valor_em_analise": pcs["em_analise"],
                "valor_reconhecido": pcs["reconhecido"],
                "valor_pago": pcs["pago"],
                "valor_pc": pcs["valor"],
                "valor_devido": pcs["devido"],
                "retroativo": _round(totais.get("retroativo")),
                "vta": _round(vta.get("metodologia_sombra") or vta.get("oficial")),
                "pcs_retroativo": pcs["origens_retroativo"],
            },
            "Fallback excepcional: usa somente execucao de PCs definitivamente pagos; historicos, pendentes e nao pagos ficam fora.",
            "Motor Temporal: DATA_PC contra linha temporal dos ciclos.",
            risco_dupla + (["Ha PC sem DATA_PC."] if pcs["sem_data"] else []),
        ),
        _alternativa(
            MET_CONSUMIDOS,
            consumidos["presente"],
            65 - (15 if dupla.get("existe") else 0) - (10 if tem_pendencia else 0),
            CONF_MEDIA,
            consumidos["valor"],
            {
                "valor_em_analise": consumidos["valor"],
                "valor_reconhecido": 0.0,
                "valor_pago": 0.0,
                "vta": _round(vta.get("metodologia_sombra_integral") or vta.get("oficial")),
            },
            "Fallback excepcional: itens consumidos representam execucao efetivamente paga quando nao ha financeiro suficiente.",
            "Aba itens_Consumidos lida pelo leitor v10.",
            riscos_gerais,
        ),
    ]


def _comparar(alternativas: list[dict[str, Any]]) -> list[dict[str, Any]]:
    disponiveis = [a for a in alternativas if a.get("disponivel")]
    comparacao: list[dict[str, Any]] = []
    for base in disponiveis:
        for outra in disponiveis:
            if base["nome"] >= outra["nome"]:
                continue
            valor_base = base.get("valor_recomendado")
            valor_outra = outra.get("valor_recomendado")
            if valor_base is None or valor_outra is None:
                continue
            comparacao.append({
                "metodologia_a": base["nome"],
                "metodologia_b": outra["nome"],
                "diferenca": _round(_f(valor_base) - _f(valor_outra)),
            })
    return comparacao


def _acao_recomendada(recomendada: dict[str, Any] | None, pendencias: list[str],
                      dupla: dict[str, Any]) -> dict[str, str]:
    if recomendada is None:
        return {
            "acao": ACAO_SOLICITAR_EVIDENCIA,
            "motivo": "Nao ha evidencia suficiente para recomendar metodologia.",
        }
    if dupla.get("existe"):
        return {
            "acao": ACAO_REVISAR_DUPLA,
            "motivo": "Ha indicio de valor ja refletido em outra fonte.",
        }
    if pendencias:
        return {
            "acao": ACAO_REVISAR_PENDENCIAS,
            "motivo": "Ha pendencias de leitura ou consistencia antes da confirmacao.",
        }
    return {
        "acao": ACAO_CONFIRMAR,
        "motivo": "A metodologia recomendada tem evidencia suficiente para confirmacao operacional.",
    }


def montar_motor_metodologias(leitura: dict[str, Any], painel: dict[str, Any]) -> dict[str, Any]:
    """Decide a metodologia recomendada sem alterar leitura nem painel."""
    leitura_ref = copy.deepcopy(leitura)
    painel_ref = copy.deepcopy(painel)

    if not isinstance(leitura, dict) or not isinstance(painel, dict) or not painel.get("disponivel"):
        return {
            "disponivel": False,
            "motivo": "Leitura ou painel indisponivel.",
            "evidencias": {},
            "alternativas": [],
            "recomendada": None,
            "resultado_recomendado": None,
        }

    dupla = _detectar_dupla_contagem(leitura, painel)
    pend = _pendencias(painel, dupla)
    ev = _evidencias(leitura, painel)
    alternativas = _montar_alternativas(ev, painel, dupla, pend)
    disponiveis = [a for a in alternativas if a.get("disponivel")]
    recomendada = max(disponiveis, key=lambda a: (a["score"], a["valor_recomendado"] or 0.0), default=None)

    assert leitura == leitura_ref
    assert painel == painel_ref

    resultado = None
    if recomendada:
        valores = recomendada.get("valores") or {}
        resultado = {
            "metodologia": recomendada["nome"],
            "valor_recomendado": recomendada.get("valor_recomendado"),
            "valor_em_analise": valores.get("valor_em_analise", 0.0),
            "valor_reconhecido": valores.get("valor_reconhecido", 0.0),
            "valor_pago": valores.get("valor_pago", 0.0),
            "vta": valores.get("vta"),
            "confiabilidade": recomendada.get("confiabilidade"),
            "justificativa": recomendada.get("justificativa"),
            "pc_gerador_retroativo": valores.get("pcs_retroativo") or ev["pcs"]["origens_retroativo"],
            "pendencias": pend,
            "confirmacao_requerida": True,
        }

    return {
        "disponivel": True,
        "evidencias": ev,
        "alternativas": alternativas,
        "comparacao": _comparar(alternativas),
        "dupla_contagem": dupla,
        "recomendada": recomendada,
        "resultado_recomendado": resultado,
        "proxima_acao": _acao_recomendada(recomendada, pend, dupla),
        "garantias": {
            "pc_classificacao": "execucao",
            "data_pc_enquadramento": "linha temporal dos ciclos",
            "sem_alterar_xls": True,
        },
    }
