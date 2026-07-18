"""Estado Contratual sombra derivado de Event Log.

Fase 4.2: reconstrucao deterministica e auditavel em memoria, sem
persistencia, sem cache e sem alteracao do VTA oficial.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Any, Iterable


_CICLO_ORDEM = {"C0": 0, "C1": 1, "C2": 2, "C3": 3, "C4": 4}


@dataclass(frozen=True)
class EventoContratual:
    sequencia: int
    tipo_evento: str
    origem_dado: str
    tipo_financeiro: str
    status_consolidacao: str
    identificador: str
    ciclo: str = ""
    linha: int | None = None
    valor: float = 0.0
    computa_vta: str = ""
    ja_refletido_em: str = ""
    fonte_parcela: str = ""
    justificativa: str = ""
    rastreabilidade: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class EventLog:
    eventos: tuple[EventoContratual, ...]


@dataclass(frozen=True)
class EstadoContratual:
    marco: str
    eventos_processados: int
    eventos_por_origem: dict[str, int]
    eventos_por_tipo_financeiro: dict[str, int]
    eventos_por_status: dict[str, int]
    valores_por_origem: dict[str, float]
    valores_por_tipo_financeiro: dict[str, float]
    eventos: tuple[EventoContratual, ...]
    rastreabilidade: tuple[dict[str, Any], ...]


def _to_float(valor: Any) -> float:
    try:
        if valor in (None, ""):
            return 0.0
        return float(valor)
    except (TypeError, ValueError):
        return 0.0


def _ciclo_no_marco(ciclo: str, marco: str) -> bool:
    if not marco:
        return True
    ciclo_norm = str(ciclo or "").strip().upper()
    marco_norm = str(marco or "").strip().upper()
    if not ciclo_norm:
        return True
    if ciclo_norm not in _CICLO_ORDEM or marco_norm not in _CICLO_ORDEM:
        return ciclo_norm == marco_norm
    return _CICLO_ORDEM[ciclo_norm] <= _CICLO_ORDEM[marco_norm]


def reconstruir_estado_contratual(
    event_log: EventLog,
    marco: str = "",
) -> EstadoContratual:
    """Reconstroi o Estado Contratual como funcao pura de EventLog e Marco."""
    eventos = tuple(
        evento
        for evento in sorted(event_log.eventos, key=lambda e: e.sequencia)
        if _ciclo_no_marco(evento.ciclo, marco)
    )

    por_origem: dict[str, int] = {}
    por_tipo: dict[str, int] = {}
    por_status: dict[str, int] = {}
    valores_origem: dict[str, float] = {}
    valores_tipo: dict[str, float] = {}
    rastreabilidade: list[dict[str, Any]] = []

    for evento in eventos:
        origem = evento.origem_dado or "Nao informado"
        tipo = evento.tipo_financeiro or "Nao informado"
        status = evento.status_consolidacao or "Nao informado"
        valor = _to_float(evento.valor)

        por_origem[origem] = por_origem.get(origem, 0) + 1
        por_tipo[tipo] = por_tipo.get(tipo, 0) + 1
        por_status[status] = por_status.get(status, 0) + 1
        valores_origem[origem] = round(valores_origem.get(origem, 0.0) + valor, 2)
        valores_tipo[tipo] = round(valores_tipo.get(tipo, 0.0) + valor, 2)
        rastreabilidade.append({
            "sequencia": evento.sequencia,
            "tipo_evento": evento.tipo_evento,
            "identificador": evento.identificador,
            "origem_dado": evento.origem_dado,
            "tipo_financeiro": evento.tipo_financeiro,
            "status_consolidacao": evento.status_consolidacao,
            "ciclo": evento.ciclo,
            "linha": evento.linha,
            "valor": valor,
            "justificativa": evento.justificativa,
        })

    return EstadoContratual(
        marco=str(marco or ""),
        eventos_processados=len(eventos),
        eventos_por_origem=dict(sorted(por_origem.items())),
        eventos_por_tipo_financeiro=dict(sorted(por_tipo.items())),
        eventos_por_status=dict(sorted(por_status.items())),
        valores_por_origem=dict(sorted(valores_origem.items())),
        valores_por_tipo_financeiro=dict(sorted(valores_tipo.items())),
        eventos=eventos,
        rastreabilidade=tuple(rastreabilidade),
    )


def estado_contratual_para_dict(estado: EstadoContratual) -> dict[str, Any]:
    return asdict(estado)


def montar_event_log_sombra(
    parcelas_base: Iterable[dict[str, Any]],
    itens_pc: dict[str, Any] | None = None,
) -> EventLog:
    """Monta Event Log em memoria a partir das leituras ja materializadas."""
    eventos: list[EventoContratual] = []

    def proxima_seq() -> int:
        return len(eventos) + 1

    for parcela in parcelas_base:
        identificador = str(parcela.get("identificador") or "").strip()
        ciclo = _ciclo_de_parcela(parcela)
        eventos.append(EventoContratual(
            sequencia=proxima_seq(),
            tipo_evento="parcela_base_lida",
            origem_dado=str(parcela.get("origem_dado") or ""),
            tipo_financeiro=str(parcela.get("tipo_financeiro") or ""),
            status_consolidacao=str(parcela.get("status_consolidacao") or "COMPUTADO"),
            identificador=identificador,
            ciclo=ciclo,
            linha=parcela.get("linha"),
            valor=_to_float(parcela.get("valor")),
            computa_vta=str(parcela.get("computa_vta") or "Sim"),
            ja_refletido_em=str(parcela.get("ja_refletido_em") or "Nao"),
            fonte_parcela=str(parcela.get("fonte_parcela") or ""),
            justificativa=str(parcela.get("justificativa_vta") or parcela.get("motivo") or ""),
            rastreabilidade={"fonte": "parcelas_sombra"},
        ))

    for item in (itens_pc or {}).get("itens") or []:
        campos = item.get("campos_vta") or {}
        identificador = str(item.get("numero_pc") or item.get("item_ou_grupo") or "").strip()
        valor = item.get("valor_atualizado") or item.get("valor_pc") or 0
        eventos.append(EventoContratual(
            sequencia=proxima_seq(),
            tipo_evento="pc_lido",
            origem_dado=str(campos.get("origem_dado") or "Pedido de Compra"),
            tipo_financeiro=str(campos.get("tipo_financeiro") or ""),
            status_consolidacao=str(campos.get("status_consolidacao") or ""),
            identificador=identificador,
            ciclo=str(item.get("ciclo") or ""),
            linha=item.get("linha"),
            valor=_to_float(valor),
            computa_vta=str(campos.get("computa_vta") or ""),
            ja_refletido_em=str(campos.get("ja_refletido_em") or ""),
            fonte_parcela=str(campos.get("fonte_parcela") or ""),
            justificativa=str(campos.get("justificativa_vta") or ""),
            rastreabilidade={"fonte": "itens_PC"},
        ))

    return EventLog(eventos=tuple(eventos))


def _ciclo_de_parcela(parcela: dict[str, Any]) -> str:
    ciclo = str(parcela.get("ciclo") or "").strip().upper()
    if ciclo:
        return ciclo
    identificador = str(parcela.get("identificador") or "").upper()
    for parte in identificador.replace(":", " ").split():
        if parte in _CICLO_ORDEM:
            return parte
    return ""
