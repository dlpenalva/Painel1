"""Novo Motor de Posicao Contratual - Fase 4.3.2.

Este modulo implementa em fatias incrementais:

- Parte 1: regras estruturais da linha temporal (PCONTR-001, 002, 003, 004,
  005 e 026).
- Parte 2: regras de quantidades e movimentos (PCONTR-006, 007, 008, 009,
  010, 011, 023, 024, 025 e 027).
- Parte 3: regras de valor unitario e valor vigente (PCONTR-012, 013, 014 e
  015).
- Parte 4: regras de rastreabilidade e alertas (PCONTR-021, 022 e 030).

Entrada esperada: dados normalizados de Event Log / Estado Contratual, nunca
workbook, bytes XLSX ou celulas Excel. A camada experimental antiga
`_posicao_contratual.py` permanece apenas como referencia funcional, sem
dependencia de execucao aqui.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, Iterable
import unicodedata


CICLOS = ("C0", "C1", "C2", "C3", "C4")
_ORDEM_CICLOS = {ciclo: i for i, ciclo in enumerate(CICLOS)}


@dataclass(frozen=True)
class AlertaMotorPosicao:
    nivel: str
    codigo: str
    mensagem: str
    ciclo: str = ""
    identificador: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "nivel": self.nivel,
            "codigo": self.codigo,
            "mensagem": self.mensagem,
            "ciclo": self.ciclo,
            "identificador": self.identificador,
        }

    @property
    def bloqueante(self) -> bool:
        return self.nivel == "ERRO_GRAVE"


class ErroGraveMotorPosicaoContratual(ValueError):
    """Erro impeditivo para a linha temporal do novo motor."""

    def __init__(self, mensagem: str, alertas: Iterable[AlertaMotorPosicao] = ()):
        super().__init__(mensagem)
        self.alertas = tuple(alertas)


@dataclass(frozen=True)
class CicloContratual:
    ciclo: str
    data_inicio: date
    data_fim: date
    origem: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ResultadoCiclo:
    ciclo: str | None
    alertas: tuple[AlertaMotorPosicao, ...] = ()


@dataclass(frozen=True)
class MovimentoTemporal:
    identificador: str
    tipo_evento: str
    data_efeito: Any
    ciclo_conferencia: str = ""
    origem: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ProjecaoTemporal:
    identificador: str
    tipo_evento: str
    data_efeito: Any
    ciclo_calculado: str | None
    ciclo_conferencia: str
    ciclos_afetados: tuple[str, ...]
    alertas: tuple[AlertaMotorPosicao, ...]
    origem: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class ItemContratual:
    item: str
    quantidade_original: float
    origem: dict[str, Any] = field(default_factory=dict)
    valor_unitario_historico: float | int | None = None


@dataclass(frozen=True)
class MovimentoQuantidade:
    item: str
    tipo_movimento: str
    quantidade: float | int | None
    data_efeito: Any = None
    ciclo_conferencia: str = ""
    identificador: str = ""
    origem: dict[str, Any] = field(default_factory=dict)
    valor_unitario: float | int | None = None


@dataclass(frozen=True)
class LinhaQuantidade:
    item: str
    ciclo: str
    quantidade_original: float
    inclusoes_acum: float
    acrescimos_acum: float
    supressoes_acum: float
    quantidade_vigente: float
    origem: str
    check: str


@dataclass(frozen=True)
class LinhaValorVigente:
    item: str
    ciclo: str
    quantidade_vigente: float
    valor_unitario_vigente: float | None
    valor_vigente: float | None
    origem_vu: str
    check: str


@dataclass(frozen=True)
class EvidenciaMovimento:
    identificador: str
    item: str
    tipo_evento: str
    tipo_movimento: str
    data_efeito: Any
    ciclo_calculado: str | None
    ciclo_informado: str
    divergencia_ciclo: bool
    ciclos_afetados: tuple[str, ...]
    origem: dict[str, Any]
    alertas: tuple[AlertaMotorPosicao, ...]
    alertas_bloqueantes: tuple[AlertaMotorPosicao, ...]
    alertas_informativos: tuple[AlertaMotorPosicao, ...]


@dataclass(frozen=True)
class ResultadoMotorPosicao:
    ciclos: tuple[CicloContratual, ...]
    projecoes: tuple[ProjecaoTemporal, ...]
    alertas: tuple[AlertaMotorPosicao, ...]
    regras_implementadas: tuple[str, ...] = (
        "PCONTR-001",
        "PCONTR-002",
        "PCONTR-003",
        "PCONTR-004",
        "PCONTR-005",
        "PCONTR-026",
    )


@dataclass(frozen=True)
class ResultadoQuantidades:
    ciclos: tuple[CicloContratual, ...]
    linhas: tuple[LinhaQuantidade, ...]
    alertas: tuple[AlertaMotorPosicao, ...]
    regras_implementadas: tuple[str, ...] = (
        "PCONTR-006",
        "PCONTR-007",
        "PCONTR-008",
        "PCONTR-009",
        "PCONTR-010",
        "PCONTR-011",
        "PCONTR-023",
        "PCONTR-024",
        "PCONTR-025",
        "PCONTR-027",
    )


@dataclass(frozen=True)
class ResultadoValoresVigentes:
    ciclos: tuple[CicloContratual, ...]
    linhas: tuple[LinhaValorVigente, ...]
    alertas: tuple[AlertaMotorPosicao, ...]
    regras_implementadas: tuple[str, ...] = (
        "PCONTR-012",
        "PCONTR-013",
        "PCONTR-014",
        "PCONTR-015",
    )


@dataclass(frozen=True)
class ResultadoRastreabilidade:
    ciclos: tuple[CicloContratual, ...]
    evidencias: tuple[EvidenciaMovimento, ...]
    alertas: tuple[AlertaMotorPosicao, ...]
    alertas_bloqueantes: tuple[AlertaMotorPosicao, ...]
    alertas_informativos: tuple[AlertaMotorPosicao, ...]
    regras_implementadas: tuple[str, ...] = (
        "PCONTR-021",
        "PCONTR-022",
        "PCONTR-030",
    )


@dataclass(frozen=True)
class ResultadoPosicaoContratual:
    ciclos: tuple[CicloContratual, ...]
    linhas_quantidade: tuple[LinhaQuantidade, ...]
    linhas_valor_vigente: tuple[LinhaValorVigente, ...]
    evidencias: tuple[EvidenciaMovimento, ...]
    alertas: tuple[AlertaMotorPosicao, ...]
    alertas_bloqueantes: tuple[AlertaMotorPosicao, ...]
    alertas_informativos: tuple[AlertaMotorPosicao, ...]
    resumo_por_item_ciclo: tuple[dict[str, Any], ...]
    regras_implementadas: tuple[str, ...] = (
        "PCONTR-001",
        "PCONTR-002",
        "PCONTR-003",
        "PCONTR-004",
        "PCONTR-005",
        "PCONTR-006",
        "PCONTR-007",
        "PCONTR-008",
        "PCONTR-009",
        "PCONTR-010",
        "PCONTR-011",
        "PCONTR-012",
        "PCONTR-013",
        "PCONTR-014",
        "PCONTR-015",
        "PCONTR-021",
        "PCONTR-022",
        "PCONTR-023",
        "PCONTR-024",
        "PCONTR-025",
        "PCONTR-026",
        "PCONTR-027",
        "PCONTR-030",
    )


def _alerta(
    nivel: str,
    codigo: str,
    mensagem: str,
    ciclo: str = "",
    identificador: str = "",
) -> AlertaMotorPosicao:
    return AlertaMotorPosicao(
        nivel=nivel,
        codigo=codigo,
        mensagem=mensagem,
        ciclo=ciclo,
        identificador=identificador,
    )


def _separar_alertas(
    alertas: Iterable[AlertaMotorPosicao],
) -> tuple[tuple[AlertaMotorPosicao, ...], tuple[AlertaMotorPosicao, ...]]:
    bloqueantes: list[AlertaMotorPosicao] = []
    informativos: list[AlertaMotorPosicao] = []
    for alerta in alertas:
        if alerta.bloqueante:
            bloqueantes.append(alerta)
        else:
            informativos.append(alerta)
    return tuple(bloqueantes), tuple(informativos)


def _valor(obj: Any, chave: str, default: Any = None) -> Any:
    if isinstance(obj, dict):
        return obj.get(chave, default)
    return getattr(obj, chave, default)


def _data_real(valor: Any) -> date | None:
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    return None


def _to_float(valor: Any) -> float | None:
    if isinstance(valor, bool) or valor in (None, ""):
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _moeda(valor: float) -> float:
    return round(valor + 0.0000000001, 2)


def _ciclo_normalizado(valor: Any) -> str:
    return str(valor or "").strip().upper()


def _ciclo_conferencia(valor: Any) -> str:
    if not isinstance(valor, str):
        return ""
    texto = valor.strip().upper()
    if not texto or texto.startswith("="):
        return ""
    return texto


def _ciclo_proibido(ciclo: str) -> bool:
    return (
        len(ciclo) == 2
        and ciclo.startswith("C")
        and ciclo[1].isdigit()
        and ciclo not in CICLOS
    )


def _normalizar_ciclos(ciclos: Any) -> list[Any]:
    if isinstance(ciclos, dict):
        if "por_ciclo" in ciclos:
            return list((ciclos.get("por_ciclo") or {}).values())
        return list(ciclos.values())
    return list(ciclos or [])


def validar_linha_temporal(ciclos: Any) -> tuple[CicloContratual, ...]:
    """Valida e ordena C0-C4 sem aceitar ciclos extras ou datas textuais."""
    registros = _normalizar_ciclos(ciclos)
    por_ciclo: dict[str, Any] = {}
    alertas: list[AlertaMotorPosicao] = []

    for reg in registros:
        ciclo = _ciclo_normalizado(_valor(reg, "ciclo"))
        if not ciclo:
            continue
        if _ciclo_proibido(ciclo):
            alertas.append(_alerta(
                "ERRO_GRAVE",
                "CICLO_PROIBIDO",
                f"Ciclo proibido na linha temporal: {ciclo!r}; apenas C0-C4.",
                ciclo=ciclo,
            ))
            raise ErroGraveMotorPosicaoContratual(alertas[-1].mensagem, alertas)
        if ciclo not in CICLOS:
            alertas.append(_alerta(
                "ERRO_GRAVE",
                "CICLO_INVALIDO",
                f"Ciclo desconhecido na linha temporal: {ciclo!r}.",
                ciclo=ciclo,
            ))
            raise ErroGraveMotorPosicaoContratual(alertas[-1].mensagem, alertas)
        por_ciclo[ciclo] = reg

    faltantes = [ciclo for ciclo in CICLOS if ciclo not in por_ciclo]
    if faltantes:
        alerta = _alerta(
            "ERRO_GRAVE",
            "CICLO_AUSENTE",
            "Linha temporal incompleta; faltam: " + ", ".join(faltantes),
        )
        raise ErroGraveMotorPosicaoContratual(alerta.mensagem, (alerta,))

    saida: list[CicloContratual] = []
    for ciclo in CICLOS:
        reg = por_ciclo[ciclo]
        inicio = _data_real(_valor(reg, "data_inicio"))
        fim = _data_real(_valor(reg, "data_fim"))
        if inicio is None or fim is None:
            alerta = _alerta(
                "ERRO_GRAVE",
                "DATA_CICLO_INVALIDA",
                f"{ciclo} exige DATA_INICIO e DATA_FIM como datas reais.",
                ciclo=ciclo,
            )
            raise ErroGraveMotorPosicaoContratual(alerta.mensagem, (alerta,))
        if fim < inicio:
            alerta = _alerta(
                "ERRO_GRAVE",
                "DATA_CICLO_INVERTIDA",
                f"{ciclo} tem DATA_FIM anterior a DATA_INICIO.",
                ciclo=ciclo,
            )
            raise ErroGraveMotorPosicaoContratual(alerta.mensagem, (alerta,))
        saida.append(CicloContratual(
            ciclo=ciclo,
            data_inicio=inicio,
            data_fim=fim,
            origem={
                "origem_aba": _valor(reg, "origem_aba", ""),
                "origem_linha": _valor(reg, "origem_linha", None),
                "fator_acumulado": _valor(reg, "fator_acumulado", None),
            },
        ))

    return tuple(saida)


def determinar_ciclo_por_data(
    data_efeito: Any,
    ciclos: Any,
    ciclo_conferencia: Any = "",
) -> ResultadoCiclo:
    """Calcula o ciclo pela data de efeito; conferencia nunca sobrescreve."""
    linha = validar_linha_temporal(ciclos)
    por_nome = {c.ciclo: c for c in linha}
    alertas: list[AlertaMotorPosicao] = []

    conferencia = _ciclo_conferencia(ciclo_conferencia)
    if _ciclo_proibido(conferencia):
        alertas.append(_alerta(
            "ERRO_GRAVE",
            "CICLO_PROIBIDO",
            f"Conferencia indica ciclo proibido {conferencia!r}; apenas C0-C4.",
            ciclo=conferencia,
        ))
        return ResultadoCiclo(ciclo=None, alertas=tuple(alertas))

    data = _data_real(data_efeito)
    if data is None:
        alertas.append(_alerta(
            "ALERTA",
            "ADITIVO_SEM_DATA",
            f"Evento sem data de efeito real: {data_efeito!r}.",
        ))
        return ResultadoCiclo(ciclo=None, alertas=tuple(alertas))

    matches = [
        ciclo.ciclo
        for ciclo in linha
        if ciclo.data_inicio <= data <= ciclo.data_fim
    ]

    calculado: str | None = None
    if matches:
        calculado = matches[0]
        if (
            "C0" in matches
            and "C1" in matches
            and por_nome["C0"].data_inicio == por_nome["C1"].data_inicio
            and data == por_nome["C0"].data_inicio
        ):
            calculado = "C0" if conferencia == "C0" else "C1"
    elif data < por_nome["C1"].data_inicio:
        calculado = "C0"

    if calculado is None:
        alertas.append(_alerta(
            "ALERTA",
            "CICLO_INVALIDO",
            f"Data de efeito {data.isoformat()} fora das faixas C0-C4.",
        ))
        return ResultadoCiclo(ciclo=None, alertas=tuple(alertas))

    if conferencia in CICLOS and conferencia != calculado:
        alertas.append(_alerta(
            "ALERTA",
            "CICLO_DIVERGENTE",
            f"Ciclo calculado={calculado} difere da conferencia={conferencia}.",
            ciclo=calculado,
        ))

    return ResultadoCiclo(ciclo=calculado, alertas=tuple(alertas))


def ciclos_afetados_a_partir(ciclo: str | None) -> tuple[str, ...]:
    """Materializa PCONTR-005: evento vale do ciclo de efeito em diante."""
    if ciclo not in _ORDEM_CICLOS:
        return ()
    indice = _ORDEM_CICLOS[ciclo]
    return tuple(CICLOS[indice:])


def normalizar_tipo_movimento(texto: Any) -> str:
    """Classifica movimentos formais sem depender de Excel ou modulo legado."""
    if not isinstance(texto, str):
        return "DESCONHECIDO"
    sem_acento = "".join(
        ch for ch in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(ch)
    ).lower()
    if "inclus" in sem_acento:
        return "INCLUSAO"
    if "acresc" in sem_acento:
        return "ACRESCIMO"
    if "supress" in sem_acento:
        return "SUPRESSAO"
    return "DESCONHECIDO"


def _eventos_de(event_log: Any) -> list[Any]:
    if event_log is None:
        return []
    if isinstance(event_log, dict):
        return list(event_log.get("eventos") or [])
    return list(_valor(event_log, "eventos", []) or [])


def _movimento_temporal_de_evento(evento: Any) -> MovimentoTemporal | None:
    tipo = str(_valor(evento, "tipo_evento", "") or "")
    if tipo not in {"ADITIVO_CONSOLIDADO", "ADITIVO", "MOVIMENTO_CONTRATUAL"}:
        return None

    data_efeito = _valor(evento, "data_efeito", None)
    if data_efeito is None:
        data_efeito = _valor(evento, "data_evento", None)

    identificador = (
        str(_valor(evento, "identificador", "") or "")
        or str(_valor(evento, "item", "") or "")
        or str(_valor(evento, "observacao", "") or "")
    )

    return MovimentoTemporal(
        identificador=identificador,
        tipo_evento=tipo,
        data_efeito=data_efeito,
        ciclo_conferencia=str(_valor(evento, "ciclo", "") or ""),
        origem={
            "origem_aba": _valor(evento, "origem_aba", ""),
            "origem_linha": _valor(evento, "origem_linha", None),
            "fonte": "event_log",
        },
    )


def calcular_linha_temporal(
    ciclos: Any,
    movimentos: Iterable[MovimentoTemporal | dict[str, Any]] = (),
    event_log: Any = None,
    estado_contratual: Any = None,
) -> ResultadoMotorPosicao:
    """Executa apenas a Parte 1 do novo motor sobre dados normalizados.

    `estado_contratual` e aceito como insumo arquitetural para manter a
    assinatura alinhada a Fase 4, mas esta etapa nao calcula saldos ou
    quantidades a partir dele.
    """
    del estado_contratual  # Parte 1 nao consome saldos/quantidades.

    linha = validar_linha_temporal(ciclos)
    todos_movimentos: list[MovimentoTemporal] = []

    for mov in movimentos:
        if isinstance(mov, MovimentoTemporal):
            todos_movimentos.append(mov)
        else:
            todos_movimentos.append(MovimentoTemporal(
                identificador=str(mov.get("identificador") or mov.get("item") or ""),
                tipo_evento=str(mov.get("tipo_evento") or "MOVIMENTO_CONTRATUAL"),
                data_efeito=mov.get("data_efeito", mov.get("data_evento")),
                ciclo_conferencia=str(mov.get("ciclo_conferencia") or mov.get("ciclo") or ""),
                origem=dict(mov.get("origem") or {}),
            ))

    for evento in _eventos_de(event_log):
        mov = _movimento_temporal_de_evento(evento)
        if mov is not None:
            todos_movimentos.append(mov)

    projecoes: list[ProjecaoTemporal] = []
    alertas: list[AlertaMotorPosicao] = []

    for mov in todos_movimentos:
        resultado = determinar_ciclo_por_data(
            mov.data_efeito,
            linha,
            ciclo_conferencia=mov.ciclo_conferencia,
        )
        ciclos_afetados = ciclos_afetados_a_partir(resultado.ciclo)
        mov_alertas = tuple(
            _alerta(
                a.nivel,
                a.codigo,
                a.mensagem,
                ciclo=a.ciclo,
                identificador=mov.identificador,
            )
            for a in resultado.alertas
        )
        alertas.extend(mov_alertas)
        projecoes.append(ProjecaoTemporal(
            identificador=mov.identificador,
            tipo_evento=mov.tipo_evento,
            data_efeito=mov.data_efeito,
            ciclo_calculado=resultado.ciclo,
            ciclo_conferencia=_ciclo_conferencia(mov.ciclo_conferencia),
            ciclos_afetados=ciclos_afetados,
            alertas=mov_alertas,
            origem=mov.origem,
        ))

    return ResultadoMotorPosicao(
        ciclos=linha,
        projecoes=tuple(projecoes),
        alertas=tuple(alertas),
    )


def _item_de(reg: ItemContratual | dict[str, Any]) -> ItemContratual:
    if isinstance(reg, ItemContratual):
        return reg
    item = str(reg.get("item") or reg.get("identificador") or "").strip()
    qtd = _to_float(reg.get("quantidade_original", reg.get("qtd_contratada")))
    vu = _to_float(
        reg.get(
            "valor_unitario_historico",
            reg.get("vu_historico", reg.get("valor_unitario", reg.get("vu"))),
        )
    )
    return ItemContratual(
        item=item,
        quantidade_original=0.0 if qtd is None else qtd,
        valor_unitario_historico=vu,
        origem=dict(reg.get("origem") or {}),
    )


def _movimento_quantidade_de(
    reg: MovimentoQuantidade | dict[str, Any],
) -> MovimentoQuantidade:
    if isinstance(reg, MovimentoQuantidade):
        return reg
    item = str(reg.get("item") or reg.get("identificador") or "").strip()
    return MovimentoQuantidade(
        item=item,
        tipo_movimento=str(
            reg.get("tipo_movimento")
            or reg.get("tipo")
            or reg.get("tipo_evento")
            or ""
        ),
        quantidade=reg.get("quantidade", reg.get("qtd")),
        valor_unitario=reg.get("valor_unitario", reg.get("vu")),
        data_efeito=reg.get("data_efeito", reg.get("data_evento")),
        ciclo_conferencia=str(reg.get("ciclo_conferencia") or reg.get("ciclo") or ""),
        identificador=str(reg.get("identificador") or item),
        origem=dict(reg.get("origem") or {}),
    )


def _movimento_evidencia_de(reg: MovimentoTemporal | MovimentoQuantidade | dict[str, Any]) -> dict[str, Any]:
    if isinstance(reg, MovimentoTemporal):
        return {
            "identificador": reg.identificador,
            "item": "",
            "tipo_evento": reg.tipo_evento,
            "tipo_movimento": "",
            "data_efeito": reg.data_efeito,
            "ciclo_conferencia": reg.ciclo_conferencia,
            "origem": dict(reg.origem or {}),
        }
    if isinstance(reg, MovimentoQuantidade):
        return {
            "identificador": reg.identificador or reg.item,
            "item": reg.item,
            "tipo_evento": "MOVIMENTO_CONTRATUAL",
            "tipo_movimento": normalizar_tipo_movimento(reg.tipo_movimento),
            "data_efeito": reg.data_efeito,
            "ciclo_conferencia": reg.ciclo_conferencia,
            "origem": dict(reg.origem or {}),
        }

    item = str(reg.get("item") or reg.get("identificador") or "").strip()
    tipo_movimento = normalizar_tipo_movimento(
        reg.get("tipo_movimento")
        or reg.get("tipo")
        or reg.get("tipo_evento")
        or ""
    )
    return {
        "identificador": str(reg.get("identificador") or item),
        "item": item,
        "tipo_evento": str(reg.get("tipo_evento") or "MOVIMENTO_CONTRATUAL"),
        "tipo_movimento": tipo_movimento,
        "data_efeito": reg.get("data_efeito", reg.get("data_evento")),
        "ciclo_conferencia": str(reg.get("ciclo_conferencia") or reg.get("ciclo") or ""),
        "origem": dict(reg.get("origem") or {}),
    }


def _historico_vu_do_item(
    item: str,
    ciclo: str,
    item_base: ItemContratual | None,
    historico_vu: dict[str, dict[str, Any]],
    ciclos: tuple[CicloContratual, ...],
) -> float | None:
    vus = historico_vu.get(item) or {}
    direto = _to_float(vus.get(ciclo))
    if direto is not None:
        return direto

    vu_c0 = _to_float(vus.get("C0"))
    fator = _to_float(next(
        (
            ciclo_reg.origem.get("fator_acumulado")
            for ciclo_reg in ciclos
            if ciclo_reg.ciclo == ciclo
        ),
        None,
    ))
    if vu_c0 is not None and fator is not None:
        return _moeda(vu_c0 * fator)

    if item_base is not None:
        return _to_float(item_base.valor_unitario_historico)
    return None


def _linha(
    item: str,
    ciclo: str,
    quantidade_original: float,
    inclusoes: float,
    acrescimos: float,
    supressoes: float,
    origem: str,
    check: str,
) -> LinhaQuantidade:
    vigente = round(quantidade_original + inclusoes + acrescimos - supressoes, 10)
    return LinhaQuantidade(
        item=item,
        ciclo=ciclo,
        quantidade_original=quantidade_original,
        inclusoes_acum=inclusoes,
        acrescimos_acum=acrescimos,
        supressoes_acum=supressoes,
        quantidade_vigente=vigente,
        origem=origem,
        check=check,
    )


def calcular_quantidades(
    ciclos: Any,
    itens_base: Iterable[ItemContratual | dict[str, Any]],
    movimentos: Iterable[MovimentoQuantidade | dict[str, Any]] = (),
) -> ResultadoQuantidades:
    """Calcula somente quantidades contratuais por item/ciclo.

    Nao calcula valor unitario, valor vigente, VTA, retroativo, documentos ou
    qualquer saida Excel. Movimentos sem data/ciclo calculavel viram alerta e
    nao alteram a posicao.
    """
    linha_temporal = validar_linha_temporal(ciclos)
    alertas: list[AlertaMotorPosicao] = []

    base: dict[str, ItemContratual] = {}
    for item in (_item_de(reg) for reg in itens_base):
        if not item.item:
            continue
        if item.quantidade_original < 0:
            alerta = _alerta(
                "ERRO_GRAVE",
                "QTD_VIGENTE_NEGATIVA",
                f"Quantidade original negativa para item {item.item!r}.",
                identificador=item.item,
            )
            raise ErroGraveMotorPosicaoContratual(alerta.mensagem, (alerta,))
        base[item.item] = item

    movimentos_norm = [_movimento_quantidade_de(m) for m in movimentos]
    movimentos_processados: list[dict[str, Any]] = []
    inclusoes_por_item: dict[str, str] = {}

    for mov in movimentos_norm:
        tipo = normalizar_tipo_movimento(mov.tipo_movimento)
        qtd = _to_float(mov.quantidade)
        resultado = determinar_ciclo_por_data(
            mov.data_efeito,
            linha_temporal,
            ciclo_conferencia=mov.ciclo_conferencia,
        )
        for alerta in resultado.alertas:
            alertas.append(_alerta(
                alerta.nivel,
                alerta.codigo,
                alerta.mensagem,
                ciclo=alerta.ciclo,
                identificador=mov.identificador or mov.item,
            ))
        if resultado.ciclo is None:
            continue
        if tipo == "DESCONHECIDO":
            alertas.append(_alerta(
                "ALERTA",
                "MOVIMENTO_DESCONHECIDO",
                f"Movimento de item {mov.item!r} nao classificado.",
                ciclo=resultado.ciclo,
                identificador=mov.identificador or mov.item,
            ))
            continue
        if qtd is None:
            alertas.append(_alerta(
                "ALERTA",
                "MOVIMENTO_SEM_QUANTIDADE",
                f"Movimento de item {mov.item!r} sem quantidade.",
                ciclo=resultado.ciclo,
                identificador=mov.identificador or mov.item,
            ))
            continue
        if qtd < 0:
            alerta = _alerta(
                "ERRO_GRAVE",
                "QTD_VIGENTE_NEGATIVA",
                f"Movimento de item {mov.item!r} tem quantidade negativa.",
                ciclo=resultado.ciclo,
                identificador=mov.identificador or mov.item,
            )
            raise ErroGraveMotorPosicaoContratual(alerta.mensagem, (alerta,))

        if mov.item not in base and tipo != "INCLUSAO":
            alertas.append(_alerta(
                "ALERTA",
                "MOVIMENTO_ITEM_INEXISTENTE",
                f"Movimento {tipo} para item inexistente {mov.item!r}; item nao criado.",
                ciclo=resultado.ciclo,
                identificador=mov.identificador or mov.item,
            ))
            continue

        if tipo == "INCLUSAO" and mov.item not in base:
            inclusoes_por_item.setdefault(mov.item, resultado.ciclo)

        movimentos_processados.append({
            "item": mov.item,
            "tipo": tipo,
            "quantidade": qtd,
            "ciclo": resultado.ciclo,
            "identificador": mov.identificador or mov.item,
        })

    itens_ordem = list(base)
    for item in movimentos_processados:
        nome = item["item"]
        if item["tipo"] == "INCLUSAO" and nome not in itens_ordem and nome not in base:
            itens_ordem.append(nome)

    linhas: list[LinhaQuantidade] = []
    for item in itens_ordem:
        original = base[item].quantidade_original if item in base else 0.0
        origem = "ORIGINAL" if item in base else f"INCLUIDO_{inclusoes_por_item[item]}"
        inicio_indice = 0 if item in base else _ORDEM_CICLOS[inclusoes_por_item[item]]

        for ciclo in CICLOS[inicio_indice:]:
            acrescimos = 0.0
            supressoes = 0.0
            inclusoes = 0.0
            for mov in movimentos_processados:
                if mov["item"] != item:
                    continue
                if _ORDEM_CICLOS[mov["ciclo"]] > _ORDEM_CICLOS[ciclo]:
                    continue
                if mov["tipo"] == "INCLUSAO":
                    if item in base:
                        acrescimos += mov["quantidade"]
                    else:
                        inclusoes += mov["quantidade"]
                elif mov["tipo"] == "ACRESCIMO":
                    acrescimos += mov["quantidade"]
                elif mov["tipo"] == "SUPRESSAO":
                    supressoes += mov["quantidade"]

            linha = _linha(
                item,
                ciclo,
                original,
                inclusoes,
                acrescimos,
                supressoes,
                origem,
                "OK",
            )
            if linha.quantidade_vigente < 0:
                alerta = _alerta(
                    "ERRO_GRAVE",
                    "SUPRESSAO_MAIOR_QUE_SALDO",
                    f"Supressao maior que saldo do item {item!r} em {ciclo}.",
                    ciclo=ciclo,
                    identificador=item,
                )
                raise ErroGraveMotorPosicaoContratual(alerta.mensagem, alertas + [alerta])
            if linha.quantidade_vigente == 0 and supressoes > 0:
                linha = _linha(
                    item,
                    ciclo,
                    original,
                    inclusoes,
                    acrescimos,
                    supressoes,
                    origem,
                    "SUPRESSAO_INTEGRAL",
                )
            linhas.append(linha)

    return ResultadoQuantidades(
        ciclos=linha_temporal,
        linhas=tuple(linhas),
        alertas=tuple(alertas),
    )


def calcular_valores_vigentes(
    ciclos: Any,
    itens_base: Iterable[ItemContratual | dict[str, Any]],
    movimentos: Iterable[MovimentoQuantidade | dict[str, Any]] = (),
) -> ResultadoValoresVigentes:
    """Calcula VU vigente e valor vigente sobre a quantidade ja apurada.

    Esta etapa permanece pura: nao calcula VTA, retroativo, documentos, app,
    POSICAO_CONTRATUAL ou qualquer saida Excel. PCONTR-021 fica fora desta
    funcao.
    """
    itens_norm = [_item_de(reg) for reg in itens_base]
    movimentos_norm = [_movimento_quantidade_de(reg) for reg in movimentos]
    resultado_qtd = calcular_quantidades(ciclos, itens_norm, movimentos_norm)
    alertas = list(resultado_qtd.alertas)

    vu_base: dict[str, float] = {}
    for item in itens_norm:
        vu = _to_float(item.valor_unitario_historico)
        if item.item and vu is not None:
            vu_base[item.item] = vu

    vu_inclusao: dict[str, tuple[str, float | None]] = {}
    for mov in movimentos_norm:
        if normalizar_tipo_movimento(mov.tipo_movimento) != "INCLUSAO":
            continue
        resultado = determinar_ciclo_por_data(
            mov.data_efeito,
            resultado_qtd.ciclos,
            ciclo_conferencia=mov.ciclo_conferencia,
        )
        if resultado.ciclo is None or mov.item in vu_base:
            continue
        if mov.item in vu_inclusao:
            continue
        vu = _to_float(mov.valor_unitario)
        vu_inclusao[mov.item] = (resultado.ciclo, vu)
        if vu is None:
            alertas.append(_alerta(
                "ALERTA",
                "ITEM_NOVO_SEM_VU",
                f"Item novo {mov.item!r} sem valor unitario de inclusao.",
                ciclo=resultado.ciclo,
                identificador=mov.identificador or mov.item,
            ))

    linhas: list[LinhaValorVigente] = []
    for linha_qtd in resultado_qtd.linhas:
        origem_vu = ""
        vu_vigente: float | None = None
        if linha_qtd.item in vu_base:
            origem_vu = "HISTORICO"
            vu_vigente = vu_base[linha_qtd.item]
        elif linha_qtd.item in vu_inclusao:
            origem_vu = f"INCLUSAO_{vu_inclusao[linha_qtd.item][0]}"
            vu_vigente = vu_inclusao[linha_qtd.item][1]

        valor_vigente = (
            None
            if vu_vigente is None
            else _moeda(linha_qtd.quantidade_vigente * vu_vigente)
        )
        linhas.append(LinhaValorVigente(
            item=linha_qtd.item,
            ciclo=linha_qtd.ciclo,
            quantidade_vigente=linha_qtd.quantidade_vigente,
            valor_unitario_vigente=vu_vigente,
            valor_vigente=valor_vigente,
            origem_vu=origem_vu,
            check="VU_AUSENTE" if vu_vigente is None else "OK",
        ))

    return ResultadoValoresVigentes(
        ciclos=resultado_qtd.ciclos,
        linhas=tuple(linhas),
        alertas=tuple(alertas),
    )


def calcular_rastreabilidade(
    ciclos: Any,
    movimentos: Iterable[MovimentoTemporal | MovimentoQuantidade | dict[str, Any]] = (),
    event_log: Any = None,
) -> ResultadoRastreabilidade:
    """Materializa evidencias de origem, ciclo calculado e alertas.

    Esta funcao nao calcula VTA, retroativo, documentos, app, Excel ou
    POSICAO_CONTRATUAL. Alertas nao impeditivos permanecem na evidencia para
    auditoria da posicao contratual.
    """
    linha_temporal = validar_linha_temporal(ciclos)
    regs: list[dict[str, Any]] = []

    for mov in movimentos:
        regs.append(_movimento_evidencia_de(mov))

    for evento in _eventos_de(event_log):
        mov_temporal = _movimento_temporal_de_evento(evento)
        if mov_temporal is not None:
            regs.append(_movimento_evidencia_de(mov_temporal))

    evidencias: list[EvidenciaMovimento] = []
    todos_alertas: list[AlertaMotorPosicao] = []

    for reg in regs:
        ciclo_informado = _ciclo_conferencia(reg["ciclo_conferencia"])
        resultado = determinar_ciclo_por_data(
            reg["data_efeito"],
            linha_temporal,
            ciclo_conferencia=ciclo_informado,
        )
        alertas = tuple(
            _alerta(
                alerta.nivel,
                alerta.codigo,
                alerta.mensagem,
                ciclo=alerta.ciclo,
                identificador=reg["identificador"],
            )
            for alerta in resultado.alertas
        )
        bloqueantes, informativos = _separar_alertas(alertas)
        todos_alertas.extend(alertas)
        divergencia = any(alerta.codigo == "CICLO_DIVERGENTE" for alerta in alertas)
        evidencias.append(EvidenciaMovimento(
            identificador=reg["identificador"],
            item=reg["item"],
            tipo_evento=reg["tipo_evento"],
            tipo_movimento=reg["tipo_movimento"],
            data_efeito=reg["data_efeito"],
            ciclo_calculado=resultado.ciclo,
            ciclo_informado=ciclo_informado,
            divergencia_ciclo=divergencia,
            ciclos_afetados=ciclos_afetados_a_partir(resultado.ciclo),
            origem=reg["origem"],
            alertas=alertas,
            alertas_bloqueantes=bloqueantes,
            alertas_informativos=informativos,
        ))

    bloqueantes, informativos = _separar_alertas(todos_alertas)
    return ResultadoRastreabilidade(
        ciclos=linha_temporal,
        evidencias=tuple(evidencias),
        alertas=tuple(todos_alertas),
        alertas_bloqueantes=bloqueantes,
        alertas_informativos=informativos,
    )


def calcular_posicao_contratual(
    ciclos: Any,
    itens_base: Iterable[ItemContratual | dict[str, Any]],
    movimentos: Iterable[MovimentoQuantidade | dict[str, Any]] = (),
    historico_vu: dict[str, dict[str, Any]] | None = None,
    event_log: Any = None,
) -> ResultadoPosicaoContratual:
    """Calcula uma saida consolidada para homologacao e futura UX web.

    A funcao reune quantidade, valor vigente, evidencias e alertas sem ler ou
    escrever Excel, sem VTA, retroativo, documentos, app ou POSICAO_CONTRATUAL.
    """
    itens_norm = [_item_de(reg) for reg in itens_base]
    movimentos_norm = [_movimento_quantidade_de(reg) for reg in movimentos]
    hist = dict(historico_vu or {})

    resultado_qtd = calcular_quantidades(ciclos, itens_norm, movimentos_norm)
    resultado_rastro = calcular_rastreabilidade(
        resultado_qtd.ciclos,
        movimentos_norm,
        event_log=event_log,
    )

    itens_por_nome = {item.item: item for item in itens_norm}
    vu_inclusao: dict[str, tuple[str, float | None]] = {}
    for mov in movimentos_norm:
        if normalizar_tipo_movimento(mov.tipo_movimento) != "INCLUSAO":
            continue
        resultado = determinar_ciclo_por_data(
            mov.data_efeito,
            resultado_qtd.ciclos,
            ciclo_conferencia=mov.ciclo_conferencia,
        )
        if resultado.ciclo is None or mov.item in itens_por_nome or mov.item in vu_inclusao:
            continue
        vu_inclusao[mov.item] = (resultado.ciclo, _to_float(mov.valor_unitario))

    alertas: list[AlertaMotorPosicao] = []
    alertas.extend(resultado_qtd.alertas)
    alertas.extend(resultado_rastro.alertas)

    linhas_valor: list[LinhaValorVigente] = []
    for linha in resultado_qtd.linhas:
        origem_vu = ""
        vu_vigente: float | None = None
        item_base = itens_por_nome.get(linha.item)
        if item_base is not None:
            origem_vu = "HISTORICO"
            vu_vigente = _historico_vu_do_item(
                linha.item,
                linha.ciclo,
                item_base,
                hist,
                resultado_qtd.ciclos,
            )
        elif linha.item in vu_inclusao:
            origem_vu = f"INCLUSAO_{vu_inclusao[linha.item][0]}"
            vu_vigente = vu_inclusao[linha.item][1]
            if vu_vigente is None and not any(
                a.codigo == "ITEM_NOVO_SEM_VU" and a.identificador == linha.item
                for a in alertas
            ):
                alertas.append(_alerta(
                    "ALERTA",
                    "ITEM_NOVO_SEM_VU",
                    f"Item novo {linha.item!r} sem valor unitario de inclusao.",
                    ciclo=vu_inclusao[linha.item][0],
                    identificador=linha.item,
                ))

        linhas_valor.append(LinhaValorVigente(
            item=linha.item,
            ciclo=linha.ciclo,
            quantidade_vigente=linha.quantidade_vigente,
            valor_unitario_vigente=vu_vigente,
            valor_vigente=None if vu_vigente is None else _moeda(linha.quantidade_vigente * vu_vigente),
            origem_vu=origem_vu,
            check="VU_AUSENTE" if vu_vigente is None else linha.check,
        ))

    alertas_por_item: dict[str, list[str]] = {}
    for alerta in alertas:
        item = alerta.identificador
        if not item:
            continue
        codigos = alertas_por_item.setdefault(item, [])
        if alerta.codigo not in codigos:
            codigos.append(alerta.codigo)

    valores_por_chave = {(linha.item, linha.ciclo): linha for linha in linhas_valor}
    resumo: list[dict[str, Any]] = []
    for linha in resultado_qtd.linhas:
        valor = valores_por_chave[(linha.item, linha.ciclo)]
        codigos = alertas_por_item.get(linha.item)
        resumo.append({
            "ITEM": linha.item,
            "CICLO": linha.ciclo,
            "QTD_ORIGINAL": linha.quantidade_original,
            "ACRESCIMOS_ACUM": linha.inclusoes_acum + linha.acrescimos_acum,
            "SUPRESSOES_ACUM": linha.supressoes_acum,
            "QTD_VIGENTE": linha.quantidade_vigente,
            "VU_VIGENTE": valor.valor_unitario_vigente,
            "VALOR_TOTAL_VIGENTE": valor.valor_vigente,
            "ORIGEM": linha.origem,
            "CHECK": "; ".join(codigos) if codigos else linha.check,
        })

    bloqueantes, informativos = _separar_alertas(alertas)
    return ResultadoPosicaoContratual(
        ciclos=resultado_qtd.ciclos,
        linhas_quantidade=resultado_qtd.linhas,
        linhas_valor_vigente=tuple(linhas_valor),
        evidencias=resultado_rastro.evidencias,
        alertas=tuple(alertas),
        alertas_bloqueantes=bloqueantes,
        alertas_informativos=informativos,
        resumo_por_item_ciclo=tuple(resumo),
    )
