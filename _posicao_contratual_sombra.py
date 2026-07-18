"""Adaptador sombra para o novo Motor de Posicao Contratual.

Converte o resultado ja lido pelo leitor v10 em entradas normalizadas para
_motor_posicao_contratual. Nao le Excel, nao escreve Excel e nao altera o
resultado oficial.
"""
from __future__ import annotations

from dataclasses import asdict, is_dataclass
from typing import Any

from _motor_posicao_contratual import (
    ErroGraveMotorPosicaoContratual,
    calcular_posicao_contratual,
)


def _to_float(valor: Any) -> float | None:
    if isinstance(valor, bool) or valor in (None, ""):
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _asdict(obj: Any) -> Any:
    if is_dataclass(obj):
        return asdict(obj)
    if isinstance(obj, tuple):
        return [_asdict(v) for v in obj]
    if isinstance(obj, list):
        return [_asdict(v) for v in obj]
    if isinstance(obj, dict):
        return {k: _asdict(v) for k, v in obj.items()}
    return obj


def _alerta(codigo: str, mensagem: str, nivel: str = "ALERTA") -> dict[str, Any]:
    return {
        "nivel": nivel,
        "codigo": codigo,
        "mensagem": mensagem,
        "ciclo": "",
        "identificador": "",
    }


def adaptar_ciclos(res_leitor: dict[str, Any]) -> list[dict[str, Any]]:
    return list((res_leitor.get("parametros_v10") or {}).get("ciclos") or [])


def adaptar_itens(res_leitor: dict[str, Any]) -> list[dict[str, Any]]:
    itens: dict[str, dict[str, Any]] = {}

    for item in (res_leitor.get("execucao_saldo") or {}).get("itens") or []:
        nome = str(item.get("item") or "").strip()
        if not nome:
            continue
        itens[nome] = {
            "item": nome,
            "quantidade_original": _to_float(item.get("qtd_contratada")) or 0.0,
            "valor_unitario_historico": _to_float(item.get("vu_original")),
            "origem": {"fonte": "execucao_saldo"},
        }

    for item in (res_leitor.get("historico_vu") or {}).get("itens") or []:
        nome = str(item.get("item") or "").strip()
        if not nome:
            continue
        registro = itens.setdefault(nome, {
            "item": nome,
            "quantidade_original": 0.0,
            "origem": {"fonte": "historico_vu"},
        })
        if registro.get("valor_unitario_historico") is None:
            registro["valor_unitario_historico"] = _to_float(item.get("vu_original"))

    return list(itens.values())


def adaptar_historico_vu(res_leitor: dict[str, Any]) -> dict[str, dict[str, Any]]:
    saida: dict[str, dict[str, Any]] = {}
    for item in (res_leitor.get("historico_vu") or {}).get("itens") or []:
        nome = str(item.get("item") or "").strip()
        if not nome:
            continue
        ciclos = dict(item.get("vu_ciclos") or {})
        saida[nome] = {
            "C0": ciclos.get("VU_C0", item.get("vu_original")),
            "C1": ciclos.get("VU_C1"),
            "C2": ciclos.get("VU_C2"),
            "C3": ciclos.get("VU_C3"),
            "C4": ciclos.get("VU_C4"),
        }
    return saida


def _tipo_movimento_de_evento(evento: dict[str, Any]) -> str:
    # REGRA PERMANENTE: Pedido de Compra e EXECUCAO do contrato; nunca
    # aditivo/acrescimo/alteracao estrutural. PC nao pode virar movimento
    # de quantidade no motor de posicao.
    tipo = str(evento.get("tipo_financeiro") or evento.get("tipo_evento") or "").lower()
    origem = str(evento.get("origem_dado") or "").lower()
    if "pc" in origem or "pedido" in origem or tipo == "pc_lido":
        return "EXECUCAO"
    if "supress" in tipo:
        return "SUPRESSAO"
    if "aditivo" in tipo:
        return "ACRESCIMO"
    return "DESCONHECIDO"


def adaptar_movimentos(res_leitor: dict[str, Any]) -> list[dict[str, Any]]:
    movimentos: list[dict[str, Any]] = []

    for evento in (res_leitor.get("event_log_sombra") or {}).get("eventos") or []:
        item = str(evento.get("identificador") or "").strip()
        if not item:
            continue
        if _tipo_movimento_de_evento(evento) == "EXECUCAO":
            # Execucao (PC) nao altera quantidade contratual; rastreio
            # permanece no Event Log, fora dos movimentos estruturais.
            continue
        movimentos.append({
            "identificador": item,
            "item": item,
            "tipo_movimento": _tipo_movimento_de_evento(evento),
            "quantidade": None,
            "data_efeito": None,
            "ciclo_conferencia": evento.get("ciclo") or "",
            "origem": {
                "fonte": "event_log_sombra",
                "sequencia": evento.get("sequencia"),
                "origem_dado": evento.get("origem_dado"),
                "linha": evento.get("linha"),
            },
        })

    return movimentos


def montar_posicao_contratual_sombra(res_leitor: dict[str, Any]) -> dict[str, Any]:
    """Monta a chave posicao_contratual_sombra para o retorno do leitor."""
    ciclos = adaptar_ciclos(res_leitor)
    itens = adaptar_itens(res_leitor)
    movimentos = adaptar_movimentos(res_leitor)
    historico_vu = adaptar_historico_vu(res_leitor)

    if not ciclos:
        return {
            "ok": False,
            "linhas_quantidade": [],
            "linhas_valor_vigente": [],
            "evidencias": [],
            "alertas": [_alerta("SEM_CICLOS", "Sem ciclos normalizados para posicao contratual sombra.")],
            "alertas_bloqueantes": [],
            "alertas_informativos": [],
            "resumo_por_item_ciclo": [],
            "adaptador": {
                "ciclos": 0,
                "itens": len(itens),
                "movimentos": len(movimentos),
                "fonte": "Estado/Event Log",
            },
        }

    try:
        resultado = calcular_posicao_contratual(
            ciclos,
            itens,
            movimentos,
            historico_vu=historico_vu,
            event_log=res_leitor.get("event_log_sombra") or {},
        )
        saida = _asdict(resultado)
        saida["ok"] = True
    except ErroGraveMotorPosicaoContratual as exc:
        saida = {
            "ok": False,
            "linhas_quantidade": [],
            "linhas_valor_vigente": [],
            "evidencias": [],
            "alertas": [_asdict(a) for a in exc.alertas],
            "alertas_bloqueantes": [_asdict(a) for a in exc.alertas if a.bloqueante],
            "alertas_informativos": [_asdict(a) for a in exc.alertas if not a.bloqueante],
            "resumo_por_item_ciclo": [],
        }

    saida["adaptador"] = {
        "ciclos": len(ciclos),
        "itens": len(itens),
        "movimentos": len(movimentos),
        "fonte": "Estado/Event Log",
    }
    return saida
