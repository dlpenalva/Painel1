"""Motor de Cobertura Temporal do Claus New — posicao fisica no ciclo em execucao.

Camada de DIAGNOSTICO (sombra/auditoria) que CONSOME a leitura ja materializada
e responde, sem tocar o VTA oficial nem duplicar conceito existente, a pergunta
central do eixo "cobertura temporal / remanescente no ciclo em execucao".

DISTINCAO OBRIGATORIA (hotfix): ULTIMA EVIDENCIA != COBERTURA COMPLETA.
  * ULTIMA EVIDENCIA = automatica, derivada dos dados que existem (MAX da data).
    NUNCA e "completo ate": pode haver registros anteriores nao informados.
  * COBERTURA CONFIRMADA = informacao GCC (nunca inferida por simples MAX).
  * COBERTURA INFERIDA (somente Financeiro) = quando a grade de competencias
    demonstra continuidade rigorosa (todos os meses do intervalo informados,
    inclusive os meses com valor zero; vazio != zero).
  * PC NAO admite inferencia automatica de completude: MAX(DATA_PC) e apenas
    ultima evidencia; "PC confirmado completo ate" e exclusivamente GCC.

PRIORIDADE da cobertura adotada:
  Financeiro: confirmada_gcc > inferida (grade continua) > (so ultima evidencia).
  PC:         confirmada_gcc > (so ultima evidencia).

PROJECAO fail-closed: a projecao so e autorizada a partir do periodo seguinte a
cobertura ADOTADA (fisica + confirmadas/inferidas), NUNCA a partir da ultima
evidencia de uma fonte cuja completude nao esteja confirmada/inferida.

REUSO (nao reinventa): `enquadrar_data_pc` (motor temporal) para enquadramento;
a regra homologada de `posicao_referencia` (posicao atual completa vs. fallback
global para a fotografia historica); a hierarquia Financeiro > PC > Consumo
(motor de reconciliacao) para fonte principal vs. conferencia.

REGRAS PERMANENTES honradas:
  * Posicao FISICA != cobertura FINANCEIRA != evidencia PC. Datas distintas.
  * Financeiro/PC posteriores NAO reduzem a quantidade fisica automaticamente.
  * Projecao e DIAGNOSTICO: nunca vira fato observado nem cria retroativo.
  * Nao le/escreve Excel; nao altera B23/B25/B26; nao gera documentos.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Any

from _motor_temporal import enquadrar_data_pc
from _motor_reconciliacao import HIERARQUIA_PREVALENCIA

CICLOS = ("C0", "C1", "C2", "C3", "C4")

# Modos temporais (Bloco C).
MODO_POSICAO_DE_CORTE = "POSICAO_DE_CORTE"
MODO_POSICAO_ATUAL = "POSICAO_ATUAL"
MODO_FINANCEIRO_POSTERIOR = "FINANCEIRO_POSTERIOR"
MODO_PC_POSTERIOR = "PC_POSTERIOR"
MODO_HIBRIDO_TEMPORAL = "HIBRIDO_TEMPORAL"
MODO_PROJETADO = "PROJETADO"


# --------------------------------------------------------------------------- #
# Helpers puros
# --------------------------------------------------------------------------- #
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
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _iso(valor: date | None) -> str | None:
    return valor.isoformat() if isinstance(valor, date) else None


def _mes_seguinte(ano: int, mes: int) -> tuple[int, int]:
    return (ano + 1, 1) if mes == 12 else (ano, mes + 1)


def _extrair_por_ciclo(res: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Recupera por_ciclo tolerando os aninhamentos usuais do leitor v10."""
    bloco = res.get("ciclos")
    if isinstance(bloco, dict) and bloco.get("por_ciclo"):
        return dict(bloco["por_ciclo"])
    if isinstance(bloco, list):
        return {str(c.get("ciclo")): c for c in bloco if c.get("ciclo")}
    if isinstance(res.get("por_ciclo"), dict):
        return dict(res["por_ciclo"])
    param = res.get("parametros_v10") or res.get("parametros") or {}
    if isinstance(param, dict) and isinstance(param.get("por_ciclo"), dict):
        return dict(param["por_ciclo"])
    return {}


def _itens_pc(res: dict[str, Any]) -> list[dict[str, Any]]:
    bloco = res.get("itens_pc")
    if isinstance(bloco, dict):
        return list(bloco.get("itens") or [])
    if isinstance(bloco, list):
        return list(bloco)
    return []


def _financeiro(res: dict[str, Any]) -> list[dict[str, Any]]:
    bloco = res.get("financeiro")
    if isinstance(bloco, dict):
        return list(bloco.get("itens") or bloco.get("linhas") or [])
    if isinstance(bloco, list):
        return list(bloco)
    return []


def _itens_base(res: dict[str, Any]) -> list[str]:
    """Codigos de item da base (itens_Remanesc)."""
    bloco = res.get("itens_base") or res.get("itens_remanesc") or []
    itens: list[str] = []
    for reg in bloco:
        if isinstance(reg, dict):
            cod = str(reg.get("item") or reg.get("ITEM") or "").strip()
        else:
            cod = str(reg or "").strip()
        if cod:
            itens.append(cod)
    return itens


def _remanescente_atual(res: dict[str, Any]) -> dict[str, float]:
    """Fotografia recente informada pelo fiscal (posicao_referencia!B)."""
    bloco = res.get("remanescente_atual") or {}
    saida: dict[str, float] = {}
    if isinstance(bloco, dict):
        pares = list(bloco.items())
    else:
        pares = [
            (r.get("item") or r.get("ITEM"), r.get("qtd", r.get("QTD_REM_ATUAL")))
            for r in bloco if isinstance(r, dict)
        ]
    for item, qtd in pares:
        cod = str(item or "").strip()
        val = _como_float(qtd)
        if cod and val is not None:
            saida[cod] = val
    return saida


def _confirmacao_gcc(res: dict[str, Any]) -> tuple[date | None, date | None]:
    """Datas de cobertura CONFIRMADA pela GCC (nunca inferida por MAX)."""
    bloco = res.get("confirmacao_gcc") or {}
    fin = _como_data(bloco.get("financeiro_ate") or bloco.get("financeiro_confirmada_ate"))
    pc = _como_data(bloco.get("pc_ate") or bloco.get("pc_confirmada_ate"))
    return fin, pc


def _ciclo_por_data(data_ref: date | None, por_ciclo: dict[str, dict[str, Any]]) -> str | None:
    """Enquadra uma data na linha temporal C0-C4.

    Fonte primaria: `enquadrar_data_pc` (canonica). Ela exige o calendario
    completo (C0-C4) e rejeita linhas parciais; como `por_ciclo` do leitor pode
    omitir ciclos ainda vazios, mantemos um fallback tolerante que apenas
    localiza a data no intervalo [data_inicio, data_fim] de um ciclo definido.
    """
    if data_ref is None:
        return None
    ciclo = enquadrar_data_pc(data_ref, por_ciclo)
    if ciclo in CICLOS:
        return ciclo
    for c in CICLOS:
        reg = por_ciclo.get(c) or {}
        ini, fim = _como_data(reg.get("data_inicio")), _como_data(reg.get("data_fim"))
        if ini and fim and ini <= data_ref <= fim:
            return c
    return None


def _fotografias_ciclo(res: dict[str, Any], por_ciclo: dict[str, dict]) -> list[str]:
    """Ciclos que possuem fotografia historica (QTD_REM_BASE preenchida)."""
    bloco = res.get("fotografias_ciclo")
    if isinstance(bloco, dict):
        return [c for c in CICLOS if bloco.get(c)]
    if isinstance(bloco, list):
        return [str(c).strip().upper() for c in bloco if str(c).strip().upper() in CICLOS]
    return [c for c in CICLOS if _como_data((por_ciclo.get(c) or {}).get("data_inicio"))]


def _grade_competencias(financeiro: list[dict[str, Any]]) -> tuple[
        list[date], dict[tuple[int, int], date]]:
    """Competencias financeiras efetivamente informadas (vazio != zero).

    Considera informada apenas a linha que EXISTE com competencia valida — um
    valor zero conta como informado; a AUSENCIA de linha e lacuna, nunca zero.
    Retorna (datas ordenadas, mapa (ano,mes)->maior data do mes).
    """
    datas = [_como_data(f.get("competencia") or f.get("COMPETENCIA")) for f in financeiro]
    datas = sorted(d for d in datas if d is not None)
    por_mes: dict[tuple[int, int], date] = {}
    for d in datas:
        chave = (d.year, d.month)
        por_mes[chave] = max(por_mes.get(chave, d), d)
    return datas, por_mes


def _cobertura_inferida_financeiro(por_mes: dict[tuple[int, int], date]) -> date | None:
    """Ultima data da MAIOR sequencia contigua de meses a partir do primeiro.

    Continuidade rigorosa: exige que cada mes seguinte esteja presente. Uma
    lacuna (mes ausente) interrompe a inferencia — a cobertura inferida vai
    apenas ate o fim do trecho continuo inicial. NAO extrapola a lacuna.
    """
    if not por_mes:
        return None
    meses = sorted(por_mes)
    fim = meses[0]
    for atual in meses[1:]:
        if atual == _mes_seguinte(*fim):
            fim = atual
        else:
            break
    return por_mes[fim]


def _tem_lacuna_ate(por_mes: dict[tuple[int, int], date], alvo: date) -> bool:
    """Ha algum mes ausente entre o primeiro informado e o mes de `alvo`?"""
    if not por_mes:
        return False
    y, m = sorted(por_mes)[0]
    alvo_ym = (alvo.year, alvo.month)
    while (y, m) <= alvo_ym:
        if (y, m) not in por_mes:
            return True
        y, m = _mes_seguinte(y, m)
    return False


# --------------------------------------------------------------------------- #
# Estrutura de saida
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class ResultadoCoberturaTemporal:
    # Bloco A — Marcos.
    data_analise: date | None
    ciclo_atual: str | None
    inicio_ciclo_atual: date | None
    data_fotografia_corte: date | None
    data_fotografia_recente: date | None
    # Bloco B — Ultima evidencia vs. cobertura confirmada/inferida.
    posicao_fisica_conhecida_ate: date | None
    financeiro_ultima_evidencia: date | None
    financeiro_cobertura_inferida_ate: date | None
    financeiro_cobertura_confirmada_ate: date | None
    financeiro_cobertura_adotada_ate: date | None
    pc_ultima_evidencia: date | None
    pc_cobertura_confirmada_ate: date | None
    pc_cobertura_adotada_ate: date | None
    ultima_evidencia_por_fonte: dict[str, str | None]
    projecao_autorizada_a_partir_de: date | None
    # Bloco C — Decisao.
    modo_temporal: str
    fonte_principal: str | None
    fontes_conferencia: tuple[str, ...]
    posicao_observada: dict[str, Any]
    posicao_projetada: dict[str, Any] | None
    # Garantias / diagnostico.
    posicao_atual_completa: bool
    ciclo_referencia: str | None
    origem_posicao: str
    dupla_contagem_prevenida: bool
    reconciliacao_fisica: dict[str, Any]
    alertas: tuple[dict[str, Any], ...] = ()
    fontes_somadas: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "bloco_a_marcos": {
                "data_analise": _iso(self.data_analise),
                "ciclo_atual": self.ciclo_atual,
                "inicio_ciclo_atual": _iso(self.inicio_ciclo_atual),
                "data_fotografia_corte": _iso(self.data_fotografia_corte),
                "data_fotografia_recente": _iso(self.data_fotografia_recente),
            },
            "bloco_b_cobertura": {
                "posicao_fisica_conhecida_ate": _iso(self.posicao_fisica_conhecida_ate),
                "financeiro_ultima_evidencia": _iso(self.financeiro_ultima_evidencia),
                "financeiro_cobertura_inferida_ate": _iso(self.financeiro_cobertura_inferida_ate),
                "financeiro_cobertura_confirmada_ate": _iso(self.financeiro_cobertura_confirmada_ate),
                "financeiro_cobertura_adotada_ate": _iso(self.financeiro_cobertura_adotada_ate),
                "pc_ultima_evidencia": _iso(self.pc_ultima_evidencia),
                "pc_cobertura_confirmada_ate": _iso(self.pc_cobertura_confirmada_ate),
                "pc_cobertura_adotada_ate": _iso(self.pc_cobertura_adotada_ate),
                "ultima_evidencia_por_fonte": dict(self.ultima_evidencia_por_fonte),
                "projecao_autorizada_a_partir_de": _iso(self.projecao_autorizada_a_partir_de),
            },
            "bloco_c_decisao": {
                "modo_temporal": self.modo_temporal,
                "fonte_principal": self.fonte_principal,
                "fontes_conferencia": list(self.fontes_conferencia),
                "posicao_observada": dict(self.posicao_observada),
                "posicao_projetada": dict(self.posicao_projetada) if self.posicao_projetada else None,
            },
            "posicao_atual_completa": self.posicao_atual_completa,
            "ciclo_referencia": self.ciclo_referencia,
            "origem_posicao": self.origem_posicao,
            "dupla_contagem_prevenida": self.dupla_contagem_prevenida,
            "fontes_somadas": list(self.fontes_somadas),
            "reconciliacao_fisica": dict(self.reconciliacao_fisica),
            "alertas": [dict(a) for a in self.alertas],
        }


def _alerta(codigo: str, mensagem: str, nivel: str = "INFO") -> dict[str, Any]:
    return {"nivel": nivel, "codigo": codigo, "mensagem": mensagem}


# --------------------------------------------------------------------------- #
# Motor
# --------------------------------------------------------------------------- #
def montar_cobertura_temporal(res_leitor: dict[str, Any]) -> ResultadoCoberturaTemporal:
    """Diagnostico de cobertura temporal sobre a leitura (res_leitor).

    Contrato de entrada (todos tolerantes/opcionais):
      controle: {data_corte, ciclo_vigente, data_analise}
      por_ciclo (ou parametros_v10.por_ciclo / ciclos.por_ciclo):
        {C0..C4: {data_inicio, data_fim, fator_acumulado, ...}}
      financeiro: [{competencia: date, ciclo, valor}, ...]
      itens_pc:  [{numero_pc, data_pc, valor_pc, ciclo}, ...]  (ou {"itens": [...]})
      itens_base: [{item}] ou [item]                (itens_Remanesc)
      remanescente_atual: {item: qtd} ou [{item, qtd}]  (posicao_referencia!B)
      fotografias_ciclo: [ciclos] / {ciclo: bool}   (QTD_REM_BASE preenchida)
      confirmacao_gcc: {financeiro_ate: date, pc_ate: date}  (cobertura GCC)
    """
    if not isinstance(res_leitor, dict):
        raise TypeError("res_leitor deve ser o dict retornado pelo leitor v10.")

    controle = res_leitor.get("controle") or {}
    por_ciclo = _extrair_por_ciclo(res_leitor)
    alertas: list[dict[str, Any]] = []

    data_corte = _como_data(controle.get("data_corte"))
    data_analise = _como_data(controle.get("data_analise"))
    itens_base = _itens_base(res_leitor)
    remanescente = _remanescente_atual(res_leitor)
    fotos = _fotografias_ciclo(res_leitor, por_ciclo)
    fin_conf_gcc, pc_conf_gcc = _confirmacao_gcc(res_leitor)

    # ---- Posicao fisica de referencia (regra homologada posicao_referencia) --
    ciclo_corte = _ciclo_por_data(data_corte, por_ciclo)
    completa = bool(
        data_corte is not None
        and ciclo_corte in CICLOS
        and itens_base
        and all(item in remanescente for item in itens_base)
    )
    ciclo_fallback = next((c for c in reversed(CICLOS) if c in fotos), None)

    if completa:
        ciclo_referencia = ciclo_corte
        data_fotografia_recente = data_corte
        posicao_fisica_ate = data_corte
        data_fotografia_corte = _como_data((por_ciclo.get(ciclo_corte) or {}).get("data_inicio"))
        origem = f"POSICAO ATUAL INFORMADA - {_iso(data_corte)}"
    else:
        ciclo_referencia = ciclo_fallback
        data_fotografia_recente = None
        data_fotografia_corte = _como_data((por_ciclo.get(ciclo_fallback) or {}).get("data_inicio"))
        posicao_fisica_ate = data_fotografia_corte
        if ciclo_referencia is None:
            origem = "POSICAO DE REFERENCIA INDISPONIVEL"
            alertas.append(_alerta(
                "POSICAO_FISICA_INDISPONIVEL",
                "Sem fotografia atual completa e sem fotografia historica valida."))
        elif remanescente:
            origem = (f"POSICAO ATUAL INCOMPLETA - UTILIZADA ABERTURA {ciclo_referencia} "
                      f"({_iso(posicao_fisica_ate)})")
        else:
            origem = f"ABERTURA DO CICLO {ciclo_referencia} - {_iso(posicao_fisica_ate)}"

    # ---- Marcos: ciclo atual e inicio ---------------------------------------
    ciclo_atual = str(controle.get("ciclo_vigente") or "").strip().upper() or ciclo_referencia
    if ciclo_atual not in CICLOS:
        ciclo_atual = ciclo_referencia
    inicio_ciclo_atual = _como_data((por_ciclo.get(ciclo_atual) or {}).get("data_inicio"))

    # ---- FINANCEIRO: ultima evidencia, inferida, confirmada, adotada --------
    _fin_datas, fin_por_mes = _grade_competencias(_financeiro(res_leitor))
    financeiro_ultima = _fin_datas[-1] if _fin_datas else None
    financeiro_inferida = _cobertura_inferida_financeiro(fin_por_mes)
    if fin_conf_gcc is not None:
        financeiro_adotada = fin_conf_gcc
        # A confirmacao GCC prevalece, mas a lacuna conhecida vira conferencia.
        if _tem_lacuna_ate(fin_por_mes, fin_conf_gcc):
            alertas.append(_alerta(
                "FINANCEIRO_LACUNA_SOB_CONFIRMACAO",
                "Cobertura financeira confirmada pela GCC abrange meses ausentes "
                "na grade (lacuna preservada para conferencia).", "ALERTA"))
    else:
        financeiro_adotada = financeiro_inferida

    # ---- PC: ultima evidencia e confirmada (SEM inferencia automatica) ------
    pc_datas = [_como_data(p.get("data_pc") or p.get("DATA_PC")) for p in _itens_pc(res_leitor)]
    pc_datas = [d for d in pc_datas if d is not None]
    pc_ultima = max(pc_datas) if pc_datas else None
    pc_adotada = pc_conf_gcc  # PC nao admite inferencia de completude por ausencia.

    ultima_evidencia = {
        "fisica": _iso(posicao_fisica_ate),
        "financeiro": _iso(financeiro_ultima),
        "pc": _iso(pc_ultima),
    }

    # ---- PROJECAO fail-closed: a partir da cobertura ADOTADA, nunca da MAX ---
    adotadas = [d for d in (posicao_fisica_ate, financeiro_adotada, pc_adotada) if d is not None]
    cobertura_adotada_geral = max(adotadas) if adotadas else None
    projecao_autorizada = None
    if (data_analise and cobertura_adotada_geral
            and data_analise > cobertura_adotada_geral):
        projecao_autorizada = cobertura_adotada_geral + timedelta(days=1)

    # ---- Decisao (Bloco C): modo temporal (pela ULTIMA EVIDENCIA) -----------
    posterior_fin = bool(financeiro_ultima and posicao_fisica_ate and financeiro_ultima > posicao_fisica_ate)
    posterior_pc = bool(pc_ultima and posicao_fisica_ate and pc_ultima > posicao_fisica_ate)

    if completa and not (posterior_fin or posterior_pc):
        modo = MODO_POSICAO_ATUAL
    elif posterior_fin and posterior_pc:
        modo = MODO_HIBRIDO_TEMPORAL
    elif posterior_fin:
        modo = MODO_FINANCEIRO_POSTERIOR
    elif posterior_pc:
        modo = MODO_PC_POSTERIOR
    else:
        modo = MODO_POSICAO_DE_CORTE

    # ---- Fonte principal vs conferencia (hierarquia; sem dupla contagem) -----
    fontes_presentes = []
    if financeiro_ultima is not None:
        fontes_presentes.append("financeiro")
    if pc_ultima is not None:
        fontes_presentes.append("pc")
    ordenadas = [f for f in HIERARQUIA_PREVALENCIA if f in fontes_presentes]
    fonte_principal = ordenadas[0] if ordenadas else None
    fontes_conferencia = tuple(f for f in ordenadas[1:])
    fontes_somadas = (fonte_principal,) if fonte_principal else ()

    # ---- Posicao observada x projetada --------------------------------------
    posicao_observada = {
        "data": _iso(posicao_fisica_ate),
        "ciclo": ciclo_referencia,
        "origem": origem,
        "base": "fotografia atual" if completa else "fotografia historica (abertura)",
    }
    posicao_projetada = None
    if projecao_autorizada is not None:
        posicao_projetada = {
            "a_partir_de": _iso(projecao_autorizada),
            "ate": _iso(data_analise),
            "natureza": "ESTIMATIVA - nao observada, nao cria retroativo a pagar",
        }
        alertas.append(_alerta(
            "PROJECAO_AUTORIZADA",
            f"Projecao autorizada a partir de {_iso(projecao_autorizada)} (periodo "
            f"seguinte a cobertura adotada {_iso(cobertura_adotada_geral)}); posicao "
            f"observada mantida na fotografia ({_iso(posicao_fisica_ate)})."))

    # ---- Reconciliacao fisica (VTA sombra ao nivel da posicao) ---------------
    base_fotografia = sum(remanescente.values()) if (completa and remanescente) else None
    reconciliacao_fisica = {
        "reconciliado": ciclo_referencia is not None,
        "data_posicao_utilizada": _iso(posicao_fisica_ate),
        "execucao_desde_corte": 0.0 if not completa else None,
        "base_fotografia_atual": base_fotografia,
        "observacao": (
            "Fallback: fotografia = abertura do ciclo; execucao desde o corte = 0."
            if not completa else
            "Posicao atual completa: reconcilia contra a fotografia informada."
        ),
    }

    return ResultadoCoberturaTemporal(
        data_analise=data_analise,
        ciclo_atual=ciclo_atual,
        inicio_ciclo_atual=inicio_ciclo_atual,
        data_fotografia_corte=data_fotografia_corte,
        data_fotografia_recente=data_fotografia_recente,
        posicao_fisica_conhecida_ate=posicao_fisica_ate,
        financeiro_ultima_evidencia=financeiro_ultima,
        financeiro_cobertura_inferida_ate=financeiro_inferida,
        financeiro_cobertura_confirmada_ate=fin_conf_gcc,
        financeiro_cobertura_adotada_ate=financeiro_adotada,
        pc_ultima_evidencia=pc_ultima,
        pc_cobertura_confirmada_ate=pc_conf_gcc,
        pc_cobertura_adotada_ate=pc_adotada,
        ultima_evidencia_por_fonte=ultima_evidencia,
        projecao_autorizada_a_partir_de=projecao_autorizada,
        modo_temporal=modo,
        fonte_principal=fonte_principal,
        fontes_conferencia=fontes_conferencia,
        posicao_observada=posicao_observada,
        posicao_projetada=posicao_projetada,
        posicao_atual_completa=completa,
        ciclo_referencia=ciclo_referencia,
        origem_posicao=origem,
        dupla_contagem_prevenida=True,
        reconciliacao_fisica=reconciliacao_fisica,
        alertas=tuple(alertas),
        fontes_somadas=fontes_somadas,
    )
