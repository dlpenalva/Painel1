"""Motor Temporal do Claus New — Sprint 1.

Camada de RACIOCINIO de negocio que CONSOME as estruturas ja consolidadas
(Event Log, Estado Contratual, Motor de Posicao Contratual). NAO cria motor
paralelo e NAO duplica regras: reutiliza `determinar_ciclo_por_data` para o
enquadramento temporal e o Estado Contratual sombra para a memoria/rastro.

Responde, em uma unica saida estruturada, as perguntas de negocio da Sprint 1:

  Q1  O reajuste e aplicavel?
  Q2  Qual indice/percentual deve ser utilizado?
  Q3  Qual e o interregno correto?
  Q4  Em qual ciclo cada Pedido de Compra pertence?
  Q5  Quanto foi efetivamente pago?
  Q6  Quanto deveria ter sido pago?
  Q7  Qual e o delta entre pago e devido?
  Q8  Esse delta e potencial, reconhecido ou ja pago?
  Q9  Qual retroativo pertence a cada PC?
  Q10 Qual memoria temporal explica o resultado?

REGRAS PERMANENTES honradas aqui:
  * O enquadramento de ciclo usa EXCLUSIVAMENTE a linha temporal dos ciclos
    (datas de inicio/fim). Nunca usa formalizacao nem COMPUTAR_NESTA_APURACAO.
  * Linha temporal, Formalizacao e Financeiro sao conceitos independentes:
    COMPUTAR_NESTA_APURACAO afeta apenas fator/financeiro, jamais o ciclo.
  * Pedido de Compra representa execucao; nunca aditivo/acrescimo/alteracao.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from typing import Any

from _motor_posicao_contratual import (
    determinar_ciclo_por_data,
    ErroGraveMotorPosicaoContratual,
)
from _estado_contratual_sombra import (
    montar_event_log_sombra,
    reconstruir_estado_contratual,
    estado_contratual_para_dict,
)
from _motor_vta_sombra import calcular_vta_sombra
from _efeitos_financeiros_pc import efeito_financeiro_pc

CICLOS = ("C0", "C1", "C2", "C3", "C4")

# Naturezas do delta (Q8).
DELTA_POTENCIAL = "potencial"
DELTA_RECONHECIDO = "reconhecido"
DELTA_JA_PAGO = "ja_pago"

# Cobertura de regras de negocio desta sprint (uma por pergunta).
REGRAS_COBERTAS = (
    "MT-Q1-reajuste-aplicavel",
    "MT-Q2-indice-percentual",
    "MT-Q3-interregno",
    "MT-Q4-ciclo-temporal",
    "MT-Q5-valor-pago",
    "MT-Q6-valor-devido",
    "MT-Q7-delta",
    "MT-Q8-natureza-delta",
    "MT-Q9-retroativo-por-pc",
    "MT-Q10-memoria-temporal",
    "MT-Q11-valor-contrato-hoje",
    "MT-Q12-vta-por-metodologia",
    "MT-Q13-saldo-por-ciclo",
)


# --------------------------------------------------------------------------- #
# Estruturas de saida
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class ClassificacaoPC:
    numero_pc: str
    data_pc: date | None
    # Q4 — ciclo pela linha temporal (nunca por formalizacao).
    ciclo_temporal: str | None
    ciclo_informado: str
    divergencia_ciclo: bool
    # Q1/Q2/Q3 — reajuste do ciclo.
    reajuste_aplicavel: bool
    indice_percentual: float | None
    fator_aplicado: float | None
    interregno_inicio: date | None
    interregno_fim: date | None
    computa_nesta_apuracao: bool
    efeito_financeiro_pc: str | None
    # Q5/Q6/Q7 — financeiro.
    valor_pc: float | None
    valor_devido: float | None
    valor_pago: float | None
    elegivel_retroativo_pc: bool
    delta: float | None
    # Q8/Q9 — natureza e retroativo.
    natureza_delta: str
    retroativo: float | None
    # Q10 — memoria temporal (rastro legivel).
    memoria_temporal: tuple[str, ...] = ()
    alertas: tuple[dict[str, Any], ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "numero_pc": self.numero_pc,
            "data_pc": _iso(self.data_pc),
            "ciclo_temporal": self.ciclo_temporal,
            "ciclo_informado": self.ciclo_informado,
            "divergencia_ciclo": self.divergencia_ciclo,
            "reajuste_aplicavel": self.reajuste_aplicavel,
            "indice_percentual": self.indice_percentual,
            "fator_aplicado": self.fator_aplicado,
            "interregno_inicio": _iso(self.interregno_inicio),
            "interregno_fim": _iso(self.interregno_fim),
            "computa_nesta_apuracao": self.computa_nesta_apuracao,
            "efeito_financeiro_pc": self.efeito_financeiro_pc,
            "valor_pc": self.valor_pc,
            "valor_devido": self.valor_devido,
            "valor_pago": self.valor_pago,
            "elegivel_retroativo_pc": self.elegivel_retroativo_pc,
            "delta": self.delta,
            "natureza_delta": self.natureza_delta,
            "retroativo": self.retroativo,
            "memoria_temporal": list(self.memoria_temporal),
            "alertas": [dict(a) for a in self.alertas],
        }


@dataclass(frozen=True)
class ResultadoMotorTemporal:
    marco: str
    ciclos: tuple[dict[str, Any], ...]
    pcs: tuple[ClassificacaoPC, ...]
    totais: dict[str, float]
    estado_contratual: dict[str, Any]
    alertas: tuple[dict[str, Any], ...]
    rastreabilidade: tuple[dict[str, Any], ...]
    # Q11 — quanto o contrato vale hoje (por ciclo + referencia).
    valor_contrato: dict[str, Any] = None  # type: ignore[assignment]
    # Q12 — VTA em cada metodologia (oficial, sombra, sombra_integral).
    vta: dict[str, Any] = None  # type: ignore[assignment]
    # Q13 — quanto resta do contrato em cada ciclo.
    saldo_por_ciclo: tuple[dict[str, Any], ...] = ()
    regras_cobertas: tuple[str, ...] = REGRAS_COBERTAS

    def to_dict(self) -> dict[str, Any]:
        return {
            "marco": self.marco,
            "ciclos": [dict(c) for c in self.ciclos],
            "pcs": [pc.to_dict() for pc in self.pcs],
            "totais": dict(self.totais),
            "estado_contratual": self.estado_contratual,
            "alertas": [dict(a) for a in self.alertas],
            "rastreabilidade": [dict(r) for r in self.rastreabilidade],
            "valor_contrato": dict(self.valor_contrato or {}),
            "vta": dict(self.vta or {}),
            "saldo_por_ciclo": [dict(s) for s in self.saldo_por_ciclo],
            "regras_cobertas": list(self.regras_cobertas),
        }


# --------------------------------------------------------------------------- #
# Helpers puros
# --------------------------------------------------------------------------- #
def _iso(valor: date | None) -> str | None:
    return valor.isoformat() if isinstance(valor, date) else None


def _como_data(valor: Any) -> date | None:
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    return None


def _como_float(valor: Any) -> float | None:
    if isinstance(valor, bool) or valor in (None, ""):
        return None
    if isinstance(valor, str) and valor.strip().startswith("="):
        return None  # formula nao avaliada
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _sim(valor: Any) -> bool:
    return str(valor or "").strip().lower() in {"sim", "s", "true", "1"}


def _alerta(codigo: str, mensagem: str, nivel: str = "ALERTA",
            identificador: str = "") -> dict[str, Any]:
    return {"nivel": nivel, "codigo": codigo, "mensagem": mensagem,
            "identificador": identificador}


def _extrair_ciclos(res_leitor: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Recupera o mapa por_ciclo do leitor, tolerando aninhamentos comuns."""
    ciclos_bloco = res_leitor.get("ciclos")
    if isinstance(ciclos_bloco, dict) and ciclos_bloco.get("por_ciclo"):
        return dict(ciclos_bloco["por_ciclo"])
    if isinstance(ciclos_bloco, list):
        return {str(c.get("ciclo")): c for c in ciclos_bloco if c.get("ciclo")}
    por_ciclo = res_leitor.get("por_ciclo")
    if isinstance(por_ciclo, dict):
        return dict(por_ciclo)
    return {}


def _linha_temporal_para_motor(por_ciclo: dict[str, dict[str, Any]]) -> list[dict[str, Any]]:
    """Monta a estrutura de ciclos que o motor de posicao valida (C0-C4)."""
    linha: list[dict[str, Any]] = []
    for ciclo in CICLOS:
        reg = por_ciclo.get(ciclo, {})
        linha.append({
            "ciclo": ciclo,
            "data_inicio": _como_data(reg.get("data_inicio")),
            "data_fim": _como_data(reg.get("data_fim")),
        })
    return linha


def enquadrar_data_pc(
    data_pc: Any,
    por_ciclo: dict[str, dict[str, Any]],
    ciclo_conferencia: str = "",
) -> str | None:
    """Fonte unica para enquadrar DATA_PC na linha temporal C0-C4.

    Retorna ``None`` quando a data ou o calendario nao permitem classificacao.
    Nao aplica fator, COMPUTAR ou qualquer efeito financeiro.
    """
    data_norm = _como_data(data_pc)
    if data_norm is None:
        return None
    linha = _linha_temporal_para_motor(por_ciclo or {})
    try:
        resultado = determinar_ciclo_por_data(
            data_norm,
            linha,
            ciclo_conferencia=(ciclo_conferencia or "").strip().upper(),
        )
    except ErroGraveMotorPosicaoContratual:
        return None
    return resultado.ciclo


def _itens_pc(res_leitor: dict[str, Any]) -> list[dict[str, Any]]:
    bloco = res_leitor.get("itens_pc")
    if isinstance(bloco, dict):
        return list(bloco.get("itens") or [])
    if isinstance(bloco, list):
        return list(bloco)
    return []


# --------------------------------------------------------------------------- #
# Classificacao de um PC
# --------------------------------------------------------------------------- #
def _classificar_pc(
    item: dict[str, Any],
    linha_motor: list[dict[str, Any]],
    por_ciclo: dict[str, dict[str, Any]],
    linha_temporal_valida: bool,
) -> ClassificacaoPC:
    numero_pc = str(item.get("numero_pc") or item.get("item_ou_grupo") or "").strip()
    data_pc = _como_data(item.get("data_pc"))
    ciclo_informado = str(item.get("ciclo") or "").strip().upper()
    # Integracao real: celula de ciclo pode vir com conteudo invalido (ex.:
    # planilha legada desalinhada trazendo uma data). Nao exibir lixo: o
    # ciclo informado so vale se for C0-C4; o enquadramento continua 100%
    # pela linha temporal (regra permanente).
    if ciclo_informado and ciclo_informado not in CICLOS:
        alertas_previos = [_alerta(
            "PC_CICLO_INFORMADO_INVALIDO",
            f"PC {numero_pc!r}: ciclo informado {ciclo_informado!r} nao e um "
            f"ciclo valido (C0-C4); considerado nao informado.",
            "INFO", numero_pc)]
        ciclo_informado = ""
    else:
        alertas_previos = []
    valor_pc = _como_float(item.get("valor_pc"))
    campos = item.get("campos_vta") or {}
    memoria: list[str] = []
    alertas: list[dict[str, Any]] = list(alertas_previos)

    # --- Q4: ciclo pela LINHA TEMPORAL (reuso do motor de posicao) ---
    ciclo_temporal: str | None = None
    if data_pc is None:
        alertas.append(_alerta("PC_SEM_DATA",
                               f"PC {numero_pc!r} sem DATA_PC; ciclo indeterminado.",
                               identificador=numero_pc))
        memoria.append("Sem DATA_PC: enquadramento temporal impossivel.")
    elif not linha_temporal_valida:
        alertas.append(_alerta("LINHA_TEMPORAL_INCOMPLETA",
                               "Calendario de ciclos incompleto; ciclo indeterminado.",
                               "ERRO_GRAVE", numero_pc))
        memoria.append("Linha temporal C0-C4 incompleta: nao foi possivel enquadrar.")
    else:
        resultado = determinar_ciclo_por_data(
            data_pc, linha_motor, ciclo_conferencia=ciclo_informado,
        )
        ciclo_temporal = resultado.ciclo
        for a in resultado.alertas:
            alertas.append(_alerta(a.codigo, a.mensagem, a.nivel, numero_pc))
        if ciclo_temporal:
            reg = por_ciclo.get(ciclo_temporal, {})
            di, df = _como_data(reg.get("data_inicio")), _como_data(reg.get("data_fim"))
            memoria.append(
                f"DATA_PC {data_pc.isoformat()} enquadrada em {ciclo_temporal} "
                f"(intervalo {_iso(di)}..{_iso(df)}) — linha temporal, "
                f"independente de formalizacao."
            )

    divergencia = bool(
        ciclo_temporal and ciclo_informado in CICLOS
        and ciclo_informado != ciclo_temporal
    )
    if divergencia:
        memoria.append(
            f"Ciclo informado ({ciclo_informado}) difere do temporal "
            f"({ciclo_temporal}); prevalece o temporal."
        )

    reg = por_ciclo.get(ciclo_temporal or "", {})
    computa = _sim(reg.get("computar_nesta_apuracao"))
    percentual = _como_float(reg.get("percentual_reajuste"))
    fator_acum = _como_float(reg.get("fator_acumulado"))
    interregno_inicio = _como_data(reg.get("data_inicio"))
    interregno_fim = _como_data(reg.get("data_fim"))
    efeito_pc = efeito_financeiro_pc(data_pc, ciclo_temporal, reg)
    efeito_informado = str(item.get("efeito_financeiro_pc") or "").strip()
    if (
        efeito_informado in {"Sim", "Nao"}
        and efeito_pc is not None
        and efeito_informado != efeito_pc
    ):
        alertas.append(_alerta(
            "EFEITO_FINANCEIRO_PC_DIVERGENTE",
            f"PC {numero_pc!r}: marcador {efeito_informado} diverge da data "
            f"canonica; prevalece {efeito_pc}.",
            "ERRO_GRAVE", numero_pc,
        ))
    if ciclo_temporal not in (None, "C0") and computa and efeito_pc is None:
        alertas.append(_alerta(
            "INICIO_EFEITO_FINANCEIRO_AUSENTE",
            f"PC {numero_pc!r} - {ciclo_temporal}: inicio do efeito financeiro "
            "ausente ou inconsistente; calculo bloqueado.",
            "ERRO_GRAVE", numero_pc,
        ))
        memoria.append(
            "Inicio do efeito financeiro indeterminado: nenhum valor atualizado "
            "foi calculado."
        )

    # --- Q1: reajuste aplicavel? (temporal + existe percentual, ciclo != C0) ---
    reajuste_aplicavel = bool(
        ciclo_temporal and ciclo_temporal != "C0"
        and efeito_pc == "Sim"
        and (percentual not in (None, 0.0) or fator_acum not in (None, 1.0))
    )

    # --- Q6: valor devido. COMPUTAR afeta so o fator, nunca o ciclo. ---
    # Se o ciclo do PC nao computa nesta apuracao, o fator e neutro (1.0):
    # o reajuste desse ciclo nao entra no devido corrente (regra permanente).
    if ciclo_temporal == "C0":
        fator_aplicado: float | None = 1.0
    elif ciclo_temporal is None:
        fator_aplicado = None
        memoria.append(
            "Sem ciclo temporal valido: fator e valor devido permanecem indeterminados."
        )
    elif not computa:
        fator_aplicado = 1.0
        memoria.append(
            f"{ciclo_temporal} fora da apuracao (COMPUTAR=Nao): fator neutro 1,0 "
            f"— formalizacao/financeiro nao alteram o ciclo."
        )
    elif efeito_pc == "Nao":
        fator_aplicado = 1.0
        memoria.append(
            f"{ciclo_temporal} preservado cronologicamente, mas DATA_PC anterior "
            "ao inicio dos efeitos: fator neutro 1,0."
        )
    elif efeito_pc is None:
        fator_aplicado = None
    elif fator_acum is not None:
        fator_aplicado = fator_acum
    else:
        fator_aplicado = None
        alertas.append(_alerta("FATOR_INDETERMINADO",
                               f"Ciclo {ciclo_temporal} sem FATOR_ACUMULADO numerico.",
                               identificador=numero_pc))

    valor_devido: float | None = None
    if valor_pc is not None and fator_aplicado is not None:
        valor_devido = round(valor_pc * fator_aplicado, 2)
        memoria.append(
            f"Devido = valor_pc {valor_pc:.2f} x fator {fator_aplicado:g} "
            f"= {valor_devido:.2f}."
        )

    # --- Q5: valor pago (execucao ja liquidada) ---
    pago_flag = campos.get("pc_pago_a_contratada")
    if pago_flag is None:
        pago_flag = item.get("pc_pago_a_contratada")
    valor_pago: float | None = None
    pago_conhecido = pago_flag is not None
    pagamento_definitivo = bool(
        _sim(pago_flag) and campos.get("elegivel_retroativo_pc", True) is not False
    )
    if _sim(pago_flag):
        valor_pago = _como_float(campos.get("valor_pago"))
        if valor_pago is None:
            valor_pago = valor_pc  # pago = execucao original quando liquidada
        memoria.append(f"PC marcado como pago; valor pago = {valor_pago}.")
    elif pago_conhecido:
        valor_pago = 0.0
        memoria.append("PC nao pago; valor pago = 0.")
    else:
        alertas.append(_alerta("PAGO_INDETERMINADO",
                               f"PC {numero_pc!r} sem marcador de pagamento.",
                               "INFO", numero_pc))

    elegivel_retroativo_pc = bool(
        pagamento_definitivo
        and ciclo_temporal in CICLOS
        and fator_aplicado is not None
    )

    # --- Q7: delta ---
    delta: float | None = None
    if valor_devido is not None and _sim(pago_flag) and valor_pago is not None:
        delta = round(valor_devido - valor_pago, 2)
    elif valor_devido is not None and pago_conhecido and valor_pc is not None:
        # PC ainda nao pago: J representa somente o incremento potencial;
        # o valor integral em analise permanece separado em I.
        delta = round(valor_devido - valor_pc, 2)

    # --- Q8: natureza do delta ---
    tipo_fin = str(campos.get("tipo_financeiro") or "").strip().lower()
    if valor_pago and delta is not None and abs(delta) < 0.005:
        natureza = DELTA_JA_PAGO
    elif "reconhec" in tipo_fin:
        natureza = DELTA_RECONHECIDO
    else:
        natureza = DELTA_POTENCIAL

    # --- Q9: retroativo por PC ---
    # Metodo PC e excepcional: retroativo existe somente sobre pagamento
    # definitivo. PC historico, pendente ou nao pago nao prova execucao. Se
    # houve glosa, valor_pago ja contem a base final e substitui valor_pc.
    retroativo: float | None = None
    if (
        elegivel_retroativo_pc
        and valor_pago is not None and valor_pago > 0.0
        and fator_aplicado is not None
    ):
        retroativo = round(valor_pago * (fator_aplicado - 1.0), 2)
        if retroativo <= 0.0:
            memoria.append("Sem incremento de reajuste nesta apuracao: retroativo = 0.")
        else:
            memoria.append(
                f"Retroativo do PC pago = valor_pago {valor_pago:.2f} x (fator "
                f"{fator_aplicado:g} - 1) = {retroativo:.2f} (natureza {natureza})."
            )
    elif pagamento_definitivo and fator_aplicado is None:
        retroativo = None
        memoria.append(
            "Pagamento definitivo existe, mas o calendario nao permite calcular "
            "o retroativo; resultado indeterminado e metodo PC indisponivel."
        )
    elif pago_conhecido:
        retroativo = 0.0
        memoria.append(
            "PC sem pagamento definitivo elegivel: nao gera retroativo pelo metodo PC."
        )
    else:
        retroativo = 0.0
        memoria.append("Retroativo PC indeterminado (pagamento ou fator ausente).")

    return ClassificacaoPC(
        numero_pc=numero_pc,
        data_pc=data_pc,
        ciclo_temporal=ciclo_temporal,
        ciclo_informado=ciclo_informado,
        divergencia_ciclo=divergencia,
        reajuste_aplicavel=reajuste_aplicavel,
        indice_percentual=percentual,
        fator_aplicado=fator_aplicado,
        interregno_inicio=interregno_inicio,
        interregno_fim=interregno_fim,
        computa_nesta_apuracao=computa,
        efeito_financeiro_pc=efeito_pc,
        valor_pc=valor_pc,
        valor_devido=valor_devido,
        valor_pago=valor_pago,
        elegivel_retroativo_pc=elegivel_retroativo_pc,
        delta=delta,
        natureza_delta=natureza,
        retroativo=retroativo,
        memoria_temporal=tuple(memoria),
        alertas=tuple(alertas),
    )


# --------------------------------------------------------------------------- #
# Entrada publica
# --------------------------------------------------------------------------- #
def montar_motor_temporal(
    res_leitor: dict[str, Any],
    marco: str = "",
) -> ResultadoMotorTemporal:
    """Executa o Motor Temporal sobre a leitura ja materializada (res_leitor).

    Nao le/escreve Excel, nao gera documentos, Streamlit, PDF ou SIGA.
    """
    if not isinstance(res_leitor, dict):
        raise TypeError("res_leitor deve ser o dict retornado pelo leitor v10.")

    por_ciclo = _extrair_ciclos(res_leitor)
    linha_motor = _linha_temporal_para_motor(por_ciclo)
    alertas: list[dict[str, Any]] = []

    # Valida a linha temporal uma unica vez; reuso do motor de posicao.
    linha_temporal_valida = True
    try:
        determinar_ciclo_por_data(date(2000, 1, 1), linha_motor)
    except ErroGraveMotorPosicaoContratual as exc:
        linha_temporal_valida = False
        alertas.append(_alerta("LINHA_TEMPORAL_INVALIDA", str(exc), "ERRO_GRAVE"))

    # Ciclos expostos (Q1/Q2/Q3 por ciclo).
    ciclos_saida: list[dict[str, Any]] = []
    for ciclo in CICLOS:
        reg = por_ciclo.get(ciclo, {})
        percentual = _como_float(reg.get("percentual_reajuste"))
        ciclos_saida.append({
            "ciclo": ciclo,
            "data_inicio": _iso(_como_data(reg.get("data_inicio"))),
            "data_fim": _iso(_como_data(reg.get("data_fim"))),
            "interregno": [
                _iso(_como_data(reg.get("data_inicio"))),
                _iso(_como_data(reg.get("data_fim"))),
            ],
            "indice_percentual": percentual,
            "fator_acumulado": _como_float(reg.get("fator_acumulado")),
            "computa_nesta_apuracao": _sim(reg.get("computar_nesta_apuracao")),
            "reajuste_aplicavel": bool(
                ciclo != "C0" and percentual not in (None, 0.0)
            ),
        })

    # Classificacao PC a PC (Q4-Q10).
    pcs: list[ClassificacaoPC] = []
    for item in _itens_pc(res_leitor):
        pc = _classificar_pc(item, linha_motor, por_ciclo, linha_temporal_valida)
        pcs.append(pc)
        alertas.extend(pc.alertas)

    # Totais consolidados.
    def _soma(attr: str) -> float:
        return round(sum(getattr(p, attr) or 0.0 for p in pcs), 2)

    totais = {
        "count_pcs": float(len(pcs)),
        "valor_pc": _soma("valor_pc"),
        "valor_devido": _soma("valor_devido"),
        "valor_pago": _soma("valor_pago"),
        "delta": _soma("delta"),
        "retroativo": _soma("retroativo"),
        "pcs_retroativo_indeterminado": float(sum(
            1 for p in pcs
            if p.retroativo is None and (p.valor_pago or 0.0) > 0.0
        )),
    }

    # --- Estado Contratual / memoria (reuso do Event Log sombra) — Q10 ---
    estado_dict: dict[str, Any] = {}
    rastreabilidade: tuple[dict[str, Any], ...] = ()
    try:
        event_log = montar_event_log_sombra(
            parcelas_base=res_leitor.get("parcelas_sombra") or [],
            itens_pc=res_leitor.get("itens_pc"),
        )
        estado = reconstruir_estado_contratual(event_log, marco=marco)
        estado_dict = estado_contratual_para_dict(estado)
        rastreabilidade = estado.rastreabilidade
    except Exception as exc:  # reuso tolerante; nao derruba o motor
        alertas.append(_alerta("ESTADO_CONTRATUAL_INDISPONIVEL", str(exc), "INFO"))

    # --- Q11: quanto o contrato vale hoje (reuso da posicao sombra) ---
    posicao = res_leitor.get("posicao_contratual_sombra") or {}
    resumo_ic = posicao.get("resumo_por_item_ciclo") or []
    vigente_por_ciclo: dict[str, float] = {}
    for linha in resumo_ic:
        ciclo = str(linha.get("CICLO") or "")
        valor = _como_float(linha.get("VALOR_TOTAL_VIGENTE"))
        if ciclo in CICLOS and valor is not None:
            vigente_por_ciclo[ciclo] = round(
                vigente_por_ciclo.get(ciclo, 0.0) + valor, 2)
    ciclo_ref = next(
        (c for c in reversed(CICLOS) if vigente_por_ciclo.get(c)), None)
    valor_contrato = {
        "por_ciclo": vigente_por_ciclo,
        "ciclo_referencia": ciclo_ref,
        "valor_hoje": vigente_por_ciclo.get(ciclo_ref or "", None),
        "fonte": "posicao_contratual_sombra" if resumo_ic else "indisponivel",
    }
    if not resumo_ic:
        alertas.append(_alerta(
            "VALOR_CONTRATO_INDISPONIVEL",
            "posicao_contratual_sombra ausente; valor do contrato nao derivado.",
            "INFO"))

    # --- Q13: quanto resta do contrato em cada ciclo ---
    # Saldo = valor vigente do ciclo (posicao) - execucao (PCs) enquadrada
    # nele pela LINHA TEMPORAL. PC e execucao: consome, nunca altera o teto.
    executado_por_ciclo: dict[str, float] = {}
    for p in pcs:
        if p.ciclo_temporal and p.valor_pc is not None:
            executado_por_ciclo[p.ciclo_temporal] = round(
                executado_por_ciclo.get(p.ciclo_temporal, 0.0) + p.valor_pc, 2)
    saldo_por_ciclo: list[dict[str, Any]] = []
    for ciclo in CICLOS:
        vigente = vigente_por_ciclo.get(ciclo)
        executado = executado_por_ciclo.get(ciclo, 0.0)
        saldo_por_ciclo.append({
            "ciclo": ciclo,
            "valor_vigente": vigente,
            "executado": executado,
            "saldo": None if vigente is None else round(vigente - executado, 2),
        })

    # --- Q12: VTA em cada metodologia (reuso do motor VTA sombra) ---
    resumo_oficial = (
        res_leitor.get("resumo") or res_leitor.get("historico_resumo") or {})
    vta: dict[str, Any] = {}
    try:
        vta_s = calcular_vta_sombra(resumo_oficial, res_leitor.get("itens_pc"))
        vta_i = calcular_vta_sombra(
            resumo_oficial, res_leitor.get("itens_pc"),
            list(res_leitor.get("parcelas_sombra") or []))
        vta = {
            "oficial": vta_s.get("vta_oficial"),
            "metodologia_sombra": vta_s.get("vta_sombra"),
            "metodologia_sombra_integral": vta_i.get("vta_sombra"),
            "diferenca_sombra": vta_s.get("diferenca"),
            "diferenca_integral": vta_i.get("diferenca"),
            "alertas": list(vta_s.get("alertas") or []) + list(vta_i.get("alertas") or []),
        }
    except Exception as exc:  # reuso tolerante
        alertas.append(_alerta("VTA_INDISPONIVEL", str(exc), "INFO"))

    return ResultadoMotorTemporal(
        marco=str(marco or ""),
        ciclos=tuple(ciclos_saida),
        pcs=tuple(pcs),
        totais=totais,
        estado_contratual=estado_dict,
        alertas=tuple(alertas),
        rastreabilidade=tuple(rastreabilidade),
        valor_contrato=valor_contrato,
        vta=vta,
        saldo_por_ciclo=tuple(saldo_por_ciclo),
    )
