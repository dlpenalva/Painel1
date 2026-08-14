"""Motor de dominio da Adequacao Orcamentaria (Etapa 4).

Camada UNICA de matematica. Reproduz fielmente o golden normativo
`10.adequacao_orcamentaria_v4_.xlsx` (abas RESUMO, FINANCEIRO_REFERENCIA,
PEDIDOS_COMPRA, ADEQUACAO_ORCAMENTARIA, TEXTO). A UI (pages/12) e os leitores
apenas estruturam entradas e apresentam saidas; toda a regra vive aqui.

Regras confirmadas diretamente no golden:

- Origem "Pedidos de compra": media = TOTAL dos PCs considerados / JANELA (meses),
  com os meses sem pedido permanecendo no denominador (PEDIDOS_COMPRA!L11 =
  L10/RESUMO!B11). Janela = 1..60 meses; termina no ultimo dia do mes da ultima
  competencia (EOMONTH(B8,0)); inicia (janela-1) meses antes (dia 1). PC
  considerado = data dentro da janela E "Considerar" != Nao.
- Origem "Financeiro mensal": media = AVERAGE dos meses informados (ignora vazios).
- Fator = 1 + percentual. Referencia mensal reajustada = ROUND(media * fator, 2).
- Projecao futura: do mes seguinte a ultima competencia (EDATE(B8,1)) ate o mes de
  termino da vigencia (EOMONTH(B9,0)), mes a mes; mes final contado integralmente;
  sem pro-rata diario.
- Por mes: base F = override (convertido para base se "ja reajustado", i.e. C/fator)
  ou a referencia automatica; base apos saldo G = MIN(F, saldo - soma dos G
  anteriores), limitada a 0 (cap cumulativo); H = ROUND(G*fator, 2);
  I (diferenca) = ROUND(H - G, 2).
- Diferenca futura = SUM(I). Complemento estimado = retroativo + diferenca futura.
- Programacao por exercicio: para cada ano de YEAR(inicio_projecao) a YEAR(fim
  vigencia), soma das diferencas (I) dos meses daquele ano MAIS o retroativo
  somente no primeiro exercicio. Soma dos exercicios = complemento.

Arredondamento: ROUND do Excel (metade para cima) via Decimal ROUND_HALF_UP,
aplicado nos mesmos pontos do golden (referencia reajustada, H e I por mes). As
somas preservam o comportamento binario do Excel (nao re-arredondam).
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from typing import Any, Iterable

from dateutil.relativedelta import relativedelta

ORIGEM_FINANCEIRO = "financeiro"
ORIGEM_PCS = "pcs"
PREMISSA_JA_REAJUSTADO = "Valor ja reajustado"
JANELA_MIN = 1
JANELA_MAX = 60


# ---------------------------------------------------------------- utilitarios

def _round2(x: float | Decimal | int | None) -> float:
    """ROUND(x, 2) do Excel (metade para cima), reproduzido sobre o double."""
    if x is None:
        return 0.0
    return float(Decimal(str(float(x))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def _as_date(v: Any) -> date | None:
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%m/%Y", "%Y-%m"):
        try:
            return datetime.strptime(str(v).strip(), fmt).date()
        except ValueError:
            continue
    return None


def _mes1(d: date) -> date:
    """Primeiro dia do mes de d."""
    return d.replace(day=1)


def _add_meses(d: date, n: int) -> date:
    """Primeiro dia do mes n meses apos o mes de d (EDATE sobre o dia 1)."""
    return _mes1(d) + relativedelta(months=n)


def _eomonth(d: date, n: int) -> date:
    """Ultimo dia do mes n meses apos o mes de d (EOMONTH do Excel)."""
    base = _mes1(d) + relativedelta(months=n + 1)
    return base - relativedelta(days=1)


def _num(v: Any) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- entradas

@dataclass
class Pedido:
    """Pedido de Compra historico (origem PCs)."""
    identificacao: Any = ""
    data: date | None = None
    valor: float = 0.0
    considerar: bool = True

    @classmethod
    def de_dict(cls, d: dict) -> "Pedido":
        cons = d.get("considerar", d.get("Considerar", True))
        if isinstance(cons, str):
            cons = cons.strip().lower() not in ("nao", "não", "n", "false", "0")
        return cls(
            identificacao=d.get("identificacao", d.get("id", d.get("numero_pc", ""))),
            data=_as_date(d.get("data", d.get("data_pc", d.get("DATA_PC")))),
            valor=_num(d.get("valor", d.get("valor_pc", d.get("VALOR_PC")))) or 0.0,
            considerar=bool(cons),
        )


@dataclass
class OverrideMes:
    """Override mensal opcional do fiscal (aba ADEQUACAO, colunas C e D)."""
    valor: float | None = None      # vazio => automatico; 0 => sem execucao
    ja_reajustado: bool = False     # True => converter para base dividindo pelo fator


# ---------------------------------------------------------------- calculo

def media_pedidos_compra(
    pedidos: Iterable[Pedido | dict],
    ultima_competencia: date,
    janela_meses: int,
) -> dict:
    """Media mensal pela origem PCs: total considerado / janela (meses-calendario).

    Meses sem pedido permanecem no denominador (divisor = janela, nao meses com PC).
    """
    peds = [p if isinstance(p, Pedido) else Pedido.de_dict(p) for p in (pedidos or [])]
    comp = _mes1(ultima_competencia)
    fim = _eomonth(comp, 0)
    inicio = _add_meses(comp, -(janela_meses - 1))

    considerados = [
        p for p in peds
        if p.considerar and p.data is not None and inicio <= p.data <= fim
    ]
    total = sum(p.valor for p in considerados)
    meses_com = {(_mes1(p.data).year, _mes1(p.data).month) for p in considerados if p.valor}
    n_meses_com = len(meses_com)
    media = (total / janela_meses) if janela_meses else 0.0
    return {
        "inicio_janela": inicio,
        "fim_janela": fim,
        "pedidos_considerados": len(considerados),
        "meses_com_pedido": n_meses_com,
        "meses_sem_pedido": max(0, janela_meses - n_meses_com),
        "total_historico": total,
        "media_mensal": media,
    }


def media_financeiro(valores_mensais: Iterable[Any]) -> dict:
    """Media mensal pela origem Financeiro: AVERAGE dos meses informados (ignora vazios)."""
    nums = [n for n in (_num(v) for v in (valores_mensais or [])) if n is not None]
    media = (sum(nums) / len(nums)) if nums else 0.0
    return {
        "meses_com_valor": len(nums),
        "total_historico": sum(nums),
        "media_mensal": media,
    }


def valor_original_foi_informado(valor: Any) -> bool:
    """True se o valor ORIGINAL foi informado, INCLUINDO zero explicito.

    Regra da Adequacao: ZERO != VAZIO. Nao deve usar conversao numerica
    (parse) para decidir preenchimento, pois vazio viraria 0.0. Aqui olhamos
    o valor cru:
      None / "" / "   " / NaN / pd.NA / pd.NaT / "—" -> False (sem informacao)
      0 / 0.0 / "0" / "0,00" / "R$ 0,00" / "1000" -> True (informado)
    Pandas nao e importado aqui: pd.NA/pd.NaT sao detectados por str().
    """
    if valor is None or isinstance(valor, bool):
        return False
    if isinstance(valor, (int, float)):
        return not (isinstance(valor, float) and math.isnan(valor))
    texto = str(valor).strip()
    if not texto:
        return False
    if texto.lower() in ("nan", "none", "null", "nat", "<na>", "—", "-", "–"):
        return False
    return any(ch.isdigit() for ch in texto)


def janela_financeira_competencias(por_competencia: Iterable[Any], n: int = 6) -> dict:
    """Janela financeira de n competencias-CALENDARIO terminando na ULTIMA
    competencia INFORMADA (zero informado conta como informada).

    Entrada: iteravel de (ano, mes, valor) onde valor e float (inclui 0.0) ou
    None (sem informacao). Meses ausentes/None entram como "Sem informação" e
    NAO puxam competencia anterior para completar a janela. A media considera
    apenas as competencias com valor informado.

    Saida: {competencias:[{ano,mes,valor,situacao}], media_mensal,
            competencias_informadas, competencias_sem_info, total}.
    """
    mapa: dict[tuple[int, int], float | None] = {}
    for ano, mes, valor in por_competencia:
        mapa[(int(ano), int(mes))] = None if valor is None else float(valor)
    informadas_ord = [a * 12 + (m - 1) for (a, m), v in mapa.items() if v is not None]
    if not informadas_ord:
        return {"competencias": [], "media_mensal": 0.0,
                "competencias_informadas": 0, "competencias_sem_info": 0, "total": 0}
    fim = max(informadas_ord)
    inicio = fim - (n - 1)
    competencias = []
    valores_informados = []
    for ordinal in range(inicio, fim + 1):
        ano, mes0 = divmod(ordinal, 12)
        mes = mes0 + 1
        valor = mapa.get((ano, mes))
        if valor is None:
            competencias.append({"ano": ano, "mes": mes, "valor": None,
                                 "situacao": "Sem informação"})
        else:
            situacao = "Zero informado" if abs(valor) < 0.005 else "Informado"
            competencias.append({"ano": ano, "mes": mes, "valor": valor,
                                 "situacao": situacao})
            valores_informados.append(valor)
    media = (sum(valores_informados) / len(valores_informados)) if valores_informados else 0.0
    return {
        "competencias": competencias,
        "media_mensal": media,
        "competencias_informadas": len(valores_informados),
        "competencias_sem_info": sum(1 for c in competencias if c["valor"] is None),
        "total": len(competencias),
    }


def pedidos_de_itens_pc(registros: Iterable[dict], exclusoes: Iterable[Any] | None = None) -> list[Pedido]:
    """Mapeia registros estruturados de itens_PC (NUMERO_PC/DATA_PC/VALOR_PC) em
    Pedido, sem redigitacao. `exclusoes` = identificadores marcados como nao
    considerados na Adequacao (estado especifico da Adequacao, nao altera itens_PC).
    """
    excl = {str(e) for e in (exclusoes or [])}
    peds: list[Pedido] = []
    for r in registros or []:
        if not isinstance(r, dict):
            continue
        data = _as_date(r.get("data_pc") or r.get("DATA_PC") or r.get("data"))
        valor = _num(r.get("valor_pc", r.get("VALOR_PC", r.get("valor"))))
        if data is None or valor is None:
            continue
        ident = (r.get("numero_pc") or r.get("NUMERO_PC") or r.get("id")
                 or r.get("identificacao") or "")
        cons = r.get("considerar_na_adequacao", r.get("considerar", True))
        if isinstance(cons, str):
            cons = cons.strip().lower() not in ("nao", "não", "n", "false", "0")
        if str(ident) in excl:
            cons = False
        peds.append(Pedido(identificacao=ident, data=data, valor=valor, considerar=bool(cons)))
    return peds


def classificar_pedidos(
    pedidos: Iterable[Pedido | dict],
    ultima_competencia: date,
    janela_meses: int,
) -> dict:
    """Classifica cada PC em Considerado / Fora da janela / Excluido para a UI
    (espelha PEDIDOS_COMPRA!F). Nao altera o calculo da media."""
    comp = _mes1(ultima_competencia)
    fim = _eomonth(comp, 0)
    inicio = _add_meses(comp, -(janela_meses - 1))
    linhas = []
    for x in (pedidos or []):
        p = x if isinstance(x, Pedido) else Pedido.de_dict(x)
        if not p.considerar:
            sit = "Excluido"
        elif p.data is None:
            sit = "Sem data"
        elif inicio <= p.data <= fim:
            sit = "Considerado"
        else:
            sit = "Fora da janela"
        linhas.append({"identificacao": p.identificacao, "data": p.data,
                       "valor": p.valor, "situacao": sit})
    return {"inicio_janela": inicio, "fim_janela": fim, "pedidos": linhas}


# ------------------------------------------------------- cadencia (Etapa 51B)
#
# Problema corrigido: a projecao presumia MENSALIDADE — qualquer media
# historica era replicada em todos os meses futuros, transformando um gasto
# anual/por ciclo em 12+ "mensalidades". A cadencia separa QUANDO o gasto
# ocorre de QUANTO vale a ocorrencia.
#
# Criterio (deterministico e auditavel, sem estatistica opaca):
#   1. Eventos = meses-calendario com gasto real (> 0) da origem selecionada.
#   2. MENSAL: densidade de calendario >= MENSAL_DENSIDADE_MINIMA no intervalo
#      entre o primeiro e o ultimo evento (>= MENSAL_SPAN_MINIMO meses). O
#      denominador e o CALENDARIO (nao "linhas informadas"): o produtor
#      descarta competencias zeradas, entao um unico gasto nao pode virar
#      "100% dos meses informados".
#   3. Periodico: com o calendario de ciclos, conta-se ocorrencias por ciclo
#      COMPLETO (k_i). Padrao consistente = amplitude(k_i) <= 1. k = mediana.
#      k=1 por ciclo/anual, k=2 semestral, k=3 quadrimestral, k=4 trimestral,
#      k>=5 periodico generico, k>=12 mensal.
#   4. Posicao no ciclo: para cada slot j, mediana das posicoes relativas
#      (meses desde o inicio do ciclo) da j-esima ocorrencia de cada ciclo
#      base. Valor por slot = mediana dos valores da j-esima ocorrencia.
#      Nada e "hardcodado" em mes especifico.
#   5. C0 pode conter implantacao/investimento inicial: so entra na base de
#      inferencia como FALLBACK, quando nenhum ciclo C1+ completo tem evento
#      (confianca reduzida). Ciclos PRECLUSOS entram normalmente como
#      evidencia HISTORICA de execucao — preclusao juridica nao apaga o
#      perfil financeiro (e nada aqui gera retroativo novo).
#   6. Sem padrao consistente: IRREGULAR — a projecao automatica NAO inventa
#      mensalidade nem espalha media; exige premissa do usuario.

CADENCIA_MENSAL = "mensal"
CADENCIA_POR_CICLO = "por_ciclo"
CADENCIA_SEMESTRAL = "semestral"
CADENCIA_QUADRIMESTRAL = "quadrimestral"
CADENCIA_TRIMESTRAL = "trimestral"
CADENCIA_PERIODICA = "periodica"
CADENCIA_IRREGULAR = "irregular"

MENSAL_DENSIDADE_MINIMA = 0.7
MENSAL_SPAN_MINIMO = 4

_ROTULOS_CADENCIA = {
    CADENCIA_MENSAL: "Mensal (recorrente mês a mês)",
    CADENCIA_POR_CICLO: "1 ocorrência por ciclo (aprox. anual)",
    CADENCIA_SEMESTRAL: "2 ocorrências por ciclo (aprox. semestral)",
    CADENCIA_QUADRIMESTRAL: "3 ocorrências por ciclo (aprox. quadrimestral)",
    CADENCIA_TRIMESTRAL: "4 ocorrências por ciclo (aprox. trimestral)",
    CADENCIA_PERIODICA: "{k} ocorrências por ciclo",
    CADENCIA_IRREGULAR: "Histórico sem periodicidade suficiente",
}


@dataclass
class CicloCadencia:
    """Janela de calendario de um ciclo contratual, para fins de cadencia.

    `precluso` e informativo: ciclo precluso CONTINUA evidencia historica de
    execucao (nao gera retroativo — isso e regra de outra camada).
    """
    nome: str
    inicio: date
    fim: date
    precluso: bool = False


def _ordinal_mes(d: date) -> int:
    return d.year * 12 + (d.month - 1)


def _mes_de_ordinal(n: int) -> date:
    ano, mes0 = divmod(int(n), 12)
    return date(ano, mes0 + 1, 1)


def _mediana(valores: list[float]) -> float:
    ordenados = sorted(valores)
    n = len(ordenados)
    if n == 0:
        return 0.0
    meio = n // 2
    if n % 2:
        return float(ordenados[meio])
    return (float(ordenados[meio - 1]) + float(ordenados[meio])) / 2.0


def eventos_mensais(pares: Iterable[tuple]) -> list[dict]:
    """Normaliza (ano, mes, valor) em eventos mensais reais (valor > 0).

    Multiplos lancamentos no mesmo mes somam-se em um unico evento. Zeros e
    vazios nao sao eventos (zero informado = mes observado sem gasto).
    """
    por_mes: dict[int, float] = {}
    for ano, mes, valor in (pares or []):
        v = _num(valor)
        if v is None or v <= 0:
            continue
        chave = int(ano) * 12 + (int(mes) - 1)
        por_mes[chave] = por_mes.get(chave, 0.0) + v
    return [
        {"mes": _mes_de_ordinal(k), "valor": v}
        for k, v in sorted(por_mes.items())
    ]


def _ciclo_do_mes(ciclos: list[CicloCadencia], mes: date) -> CicloCadencia | None:
    for c in ciclos:
        if _mes1(c.inicio) <= mes <= _mes1(c.fim):
            return c
    return None


def inferir_cadencia(
    pares: Iterable[tuple],
    ciclos: Iterable[CicloCadencia] | None,
    ultima_competencia: Any,
) -> dict:
    """Classifica o padrao historico de execucao (ver criterio no cabecalho).

    pares: (ano, mes, valor) de TODOS os lancamentos historicos da origem.
    ciclos: calendario dos ciclos contratuais (inclui preclusos e C0).
    ultima_competencia: fim do periodo observado (define ciclo completo).
    """
    eventos = eventos_mensais(pares)
    ciclos = sorted(list(ciclos or []), key=lambda c: c.inicio)
    ultima = _as_date(ultima_competencia)

    base = {
        "padrao": CADENCIA_IRREGULAR,
        "rotulo": _ROTULOS_CADENCIA[CADENCIA_IRREGULAR],
        "ocorrencias_por_ciclo": 0,
        "posicoes": [],
        "valores_por_slot": [],
        "valor_referencia": 0.0,
        "ciclos_base": [],
        "usa_c0": False,
        "confianca": "insuficiente",
        "duracao_ciclo_meses": 12,
        "eventos": eventos,
        "explicacao": "",
    }
    if not eventos:
        base["explicacao"] = "Nenhum gasto historico identificado na origem."
        return base

    # --- 1) MENSAL por densidade de calendario -----------------------------
    primeiro = _ordinal_mes(eventos[0]["mes"])
    ultimo = _ordinal_mes(eventos[-1]["mes"])
    span = ultimo - primeiro + 1
    densidade = len(eventos) / span if span > 0 else 0.0
    if span >= MENSAL_SPAN_MINIMO and densidade >= MENSAL_DENSIDADE_MINIMA:
        base.update({
            "padrao": CADENCIA_MENSAL,
            "rotulo": _ROTULOS_CADENCIA[CADENCIA_MENSAL],
            "ocorrencias_por_ciclo": 12,
            "valor_referencia": _mediana([e["valor"] for e in eventos]),
            "confianca": "alta",
            "explicacao": (
                f"Gasto presente em {len(eventos)} de {span} meses-calendario "
                f"({densidade:.0%}): comportamento recorrente mensal."
            ),
        })
        return base

    if not ciclos or ultima is None:
        base["explicacao"] = (
            "Sem calendario de ciclos disponivel para inferir periodicidade; "
            "gastos nao sao mensais. Informe a premissa de projecao."
        )
        return base

    # --- 2) ocorrencias por ciclo COMPLETO ---------------------------------
    duracoes = [
        (_ordinal_mes(c.fim) - _ordinal_mes(c.inicio) + 1) for c in ciclos
    ]
    duracao = int(round(_mediana([float(d) for d in duracoes]))) if duracoes else 12
    por_ciclo: dict[str, list[dict]] = {c.nome: [] for c in ciclos}
    for e in eventos:
        c = _ciclo_do_mes(ciclos, e["mes"])
        if c is not None:
            por_ciclo[c.nome].append(e)

    completos = [c for c in ciclos if _mes1(c.fim) <= _mes1(ultima)]
    base_c1 = [c for c in completos if c.nome.upper() != "C0" and por_ciclo[c.nome]]
    usa_c0 = False
    if base_c1:
        ciclos_base = base_c1
    else:
        ciclos_base = [c for c in completos if por_ciclo[c.nome]]
        usa_c0 = any(c.nome.upper() == "C0" for c in ciclos_base)
    if not ciclos_base:
        base["explicacao"] = (
            "Nenhum ciclo historico completo com gastos: base insuficiente "
            "para inferir a cadencia."
        )
        return base

    ks = [len(por_ciclo[c.nome]) for c in ciclos_base]
    k = int(round(_mediana([float(x) for x in ks])))
    if (max(ks) - min(ks)) > 1 or k < 1:
        base["explicacao"] = (
            f"Ocorrencias por ciclo inconsistentes ({ks}): sem padrao "
            "periodico confiavel."
        )
        return base

    # --- 3) posicoes e valores por slot ------------------------------------
    posicoes: list[int] = []
    valores: list[float] = []
    for j in range(k):
        pos_j = []
        val_j = []
        for c in ciclos_base:
            evs = sorted(por_ciclo[c.nome], key=lambda e: e["mes"])
            if j < len(evs):
                pos_j.append(float(_ordinal_mes(evs[j]["mes"]) - _ordinal_mes(_mes1(c.inicio))))
                val_j.append(evs[j]["valor"])
        posicoes.append(int(round(_mediana(pos_j))) if pos_j else 0)
        valores.append(_mediana(val_j) if val_j else 0.0)

    if k >= 12:
        padrao = CADENCIA_MENSAL
    else:
        padrao = {
            1: CADENCIA_POR_CICLO,
            2: CADENCIA_SEMESTRAL,
            3: CADENCIA_QUADRIMESTRAL,
            4: CADENCIA_TRIMESTRAL,
        }.get(k, CADENCIA_PERIODICA)
    rotulo = _ROTULOS_CADENCIA[padrao]
    if padrao == CADENCIA_PERIODICA:
        rotulo = rotulo.format(k=k)

    confianca = "alta" if len(ciclos_base) >= 2 else "media"
    if usa_c0:
        confianca = "baixa"
    base.update({
        "padrao": padrao,
        "rotulo": rotulo,
        "ocorrencias_por_ciclo": k,
        "posicoes": posicoes,
        "valores_por_slot": valores,
        "valor_referencia": _mediana(valores),
        "ciclos_base": [c.nome for c in ciclos_base],
        "usa_c0": usa_c0,
        "confianca": confianca,
        "duracao_ciclo_meses": duracao,
        "explicacao": (
            f"{k} ocorrencia(s) por ciclo em {', '.join(c.nome for c in ciclos_base)}"
            + (" (fallback C0: sem historico posterior suficiente)" if usa_c0 else "")
            + f"; posicoes relativas medianas {posicoes} (meses do ciclo)."
        ),
    })
    return base


def cadencia_por_ciclo_forcada(pares: Iterable[tuple],
                               ciclos: Iterable[CicloCadencia] | None) -> dict:
    """Premissa POR CICLO imposta pelo usuario quando a inferencia automatica
    nao encontra padrao: 1 ocorrencia por ciclo, na posicao relativa mediana
    dos eventos observados, com o valor mediano por ocorrencia. Continua
    deterministica e baseada apenas em eventos reais."""
    eventos = eventos_mensais(pares)
    ciclos = sorted(list(ciclos or []), key=lambda c: c.inicio)
    duracoes = [
        (_ordinal_mes(c.fim) - _ordinal_mes(c.inicio) + 1) for c in ciclos
    ]
    duracao = int(round(_mediana([float(d) for d in duracoes]))) if duracoes else 12
    posicoes_rel = []
    for e in eventos:
        c = _ciclo_do_mes(ciclos, e["mes"])
        if c is not None:
            posicoes_rel.append(float(_ordinal_mes(e["mes"]) - _ordinal_mes(_mes1(c.inicio))))
    if not eventos or not ciclos:
        return {
            "padrao": CADENCIA_IRREGULAR,
            "rotulo": _ROTULOS_CADENCIA[CADENCIA_IRREGULAR],
            "ocorrencias_por_ciclo": 0, "posicoes": [], "valores_por_slot": [],
            "valor_referencia": 0.0, "ciclos_base": [], "usa_c0": False,
            "confianca": "insuficiente", "duracao_ciclo_meses": duracao,
            "eventos": eventos,
            "explicacao": "Sem eventos ou sem calendario de ciclos para a premissa por ciclo.",
        }
    posicao = int(round(_mediana(posicoes_rel))) if posicoes_rel else 0
    valor = _mediana([e["valor"] for e in eventos])
    return {
        "padrao": CADENCIA_POR_CICLO,
        "rotulo": _ROTULOS_CADENCIA[CADENCIA_POR_CICLO],
        "ocorrencias_por_ciclo": 1,
        "posicoes": [posicao],
        "valores_por_slot": [valor],
        "valor_referencia": valor,
        "ciclos_base": [c.nome for c in ciclos if any(
            _mes1(c.inicio) <= e["mes"] <= _mes1(c.fim) for e in eventos)],
        "usa_c0": False,
        "confianca": "premissa do usuario",
        "duracao_ciclo_meses": duracao,
        "eventos": eventos,
        "explicacao": (
            "Premissa POR CICLO definida pelo usuario: 1 ocorrencia por ciclo "
            f"na posicao mediana {posicao} com valor mediano por ocorrencia."
        ),
    }


def projetar_por_cadencia(
    cadencia: dict,
    ciclos: Iterable[CicloCadencia] | None,
    inicio_projecao: Any,
    fim_projecao: Any,
) -> dict:
    """Projeta as ocorrencias futuras conforme a cadencia inferida.

    Devolve {date(ano, mes, 1) -> valor base}. Somente ocorrencias que caem
    dentro de [inicio_projecao, fim_projecao] entram — nada e multiplicado
    por "todos os meses restantes". Ciclos futuros sao extrapolados a partir
    do ultimo ciclo conhecido, mantendo a duracao mediana observada. Se o
    ciclo corrente ja teve as k ocorrencias esperadas, nada mais e projetado
    nele (o perfil segue no ciclo seguinte).
    """
    inicio = _as_date(inicio_projecao)
    fim = _as_date(fim_projecao)
    if inicio is None or fim is None:
        return {}
    if cadencia.get("padrao") in (CADENCIA_IRREGULAR, CADENCIA_MENSAL):
        return {}
    k = int(cadencia.get("ocorrencias_por_ciclo") or 0)
    posicoes = list(cadencia.get("posicoes") or [])
    valores = list(cadencia.get("valores_por_slot") or [])
    if k < 1 or not posicoes:
        return {}
    duracao = int(cadencia.get("duracao_ciclo_meses") or 12)
    eventos = cadencia.get("eventos") or []

    janela: list[CicloCadencia] = sorted(list(ciclos or []), key=lambda c: c.inicio)
    if not janela:
        return {}
    # extrapola ciclos futuros ate cobrir o fim da projecao
    ultimo = janela[-1]
    seq = 1
    while _mes1(ultimo.fim) < _mes1(fim):
        ini = _add_meses(ultimo.fim, 1)
        ultimo = CicloCadencia(
            nome=f"{janela[-1].nome}+{seq}",
            inicio=ini,
            fim=_add_meses(ini, duracao - 1),
        )
        janela.append(ultimo)
        seq += 1

    ini_ord = _ordinal_mes(_mes1(inicio))
    fim_ord = _ordinal_mes(_mes1(fim))
    base: dict[date, float] = {}
    for c in janela:
        observados = sum(
            1 for e in eventos
            if _mes1(c.inicio) <= e["mes"] <= _mes1(c.fim)
            and _ordinal_mes(e["mes"]) < ini_ord
        )
        restantes = max(0, k - observados)
        if restantes <= 0:
            continue
        agendados = []
        for j, pos in enumerate(posicoes):
            mes_ord = _ordinal_mes(_mes1(c.inicio)) + int(pos)
            agendados.append((mes_ord, valores[j] if j < len(valores) else 0.0))
        agendados.sort()
        futuros = [(m, v) for m, v in agendados if ini_ord <= m <= fim_ord]
        for m, v in futuros[-restantes:]:
            base[_mes_de_ordinal(m)] = base.get(_mes_de_ordinal(m), 0.0) + v
    return base


def projetar_por_ciclo_proporcional(
    cadencia: dict,
    inicio_projecao: Any,
    fim_projecao: Any,
) -> dict:
    """Base ORCAMENTARIA proporcional do padrao POR CICLO (1 ocorrencia/ciclo).

    A cadencia continua explicando QUANDO o gasto historicamente ocorre; para
    fins orcamentarios, porem, projetar a ocorrencia apenas no mes historico
    pode deixar exercicios da vigencia sem cobertura (ex.: ocorrencia em
    ago/set e contrato terminando em maio). Aqui o valor recorrente do ciclo
    (valor_referencia, mediana das ocorrencias dos ciclos completos pos-C0 ja
    apurada pela inferencia) e diluido pela duracao do ciclo em meses e
    aplicado a TODOS os meses do horizonte [inicio_projecao, fim_projecao]:
    cada exercicio recebe cobertura proporcional aos seus meses restantes.

    Somente o padrao POR CICLO usa esta proporcionalizacao; os demais padroes
    seguem projetar_por_cadencia e IRREGULAR permanece fail-closed ({}).
    """
    inicio = _as_date(inicio_projecao)
    fim = _as_date(fim_projecao)
    if inicio is None or fim is None:
        return {}
    if cadencia.get("padrao") != CADENCIA_POR_CICLO:
        return {}
    valor_ciclo = _num(cadencia.get("valor_referencia")) or 0.0
    duracao = int(cadencia.get("duracao_ciclo_meses") or 12)
    if valor_ciclo <= 0 or duracao < 1:
        return {}
    base_mensal = valor_ciclo / duracao
    ini_ord = _ordinal_mes(_mes1(inicio))
    fim_ord = _ordinal_mes(_mes1(fim))
    return {_mes_de_ordinal(m): base_mensal for m in range(ini_ord, fim_ord + 1)}


def calcular_adequacao_orcamentaria(
    *,
    origem: str,
    percentual: float,
    ultima_competencia: Any,
    data_fim_vigencia: Any,
    retroativo: float = 0.0,
    janela_meses: int = 39,
    saldo_contratual: float | None = None,
    pedidos: Iterable[Pedido | dict] | None = None,
    financeiro_mensal: Iterable[Any] | None = None,
    overrides: dict | None = None,
) -> dict:
    """Executa a Adequacao Orcamentaria conforme o golden normativo.

    overrides: {competencia(date do mes) -> OverrideMes | {"valor":..,"ja_reajustado":..}}.
    Retorna estrutura serializavel com toda a memoria de calculo e checks.
    """
    checks: list[str] = []
    origem = ORIGEM_PCS if str(origem).strip().lower() in (
        "pcs", "pedidos de compra", "pedidos") else ORIGEM_FINANCEIRO
    comp = _as_date(ultima_competencia)
    fim_vig = _as_date(data_fim_vigencia)
    perc = _num(percentual) or 0.0
    retro = _num(retroativo) or 0.0
    saldo = _num(saldo_contratual)  # None => sem cap
    try:
        janela = int(janela_meses)
    except (TypeError, ValueError):
        janela = 0

    if not (JANELA_MIN <= janela <= JANELA_MAX):
        checks.append(f"JANELA FORA DE 1..60 (informado: {janela_meses})")
    if comp is None:
        checks.append("ULTIMA COMPETENCIA INVALIDA")
    if fim_vig is None:
        checks.append("DATA FINAL DA VIGENCIA INVALIDA")

    fator = 1.0 + perc

    if origem == ORIGEM_PCS:
        base_hist = media_pedidos_compra(pedidos or [], comp, janela) if comp and janela else {
            "inicio_janela": None, "fim_janela": None, "pedidos_considerados": 0,
            "meses_com_pedido": 0, "meses_sem_pedido": 0, "total_historico": 0.0,
            "media_mensal": 0.0}
    else:
        base_hist = media_financeiro(financeiro_mensal or [])

    referencia_mensal = base_hist.get("media_mensal", 0.0)
    referencia_reajustada = _round2(referencia_mensal * fator)

    # --- projecao futura mes a mes ---
    overrides = overrides or {}
    ov_norm: dict[tuple[int, int], OverrideMes] = {}
    for k, val in overrides.items():
        d = _as_date(k)
        if d is None:
            continue
        if isinstance(val, OverrideMes):
            ov_norm[(d.year, d.month)] = val
        elif isinstance(val, dict):
            ov_norm[(d.year, d.month)] = OverrideMes(
                valor=_num(val.get("valor")),
                ja_reajustado=bool(val.get("ja_reajustado", False)),
            )
        else:
            ov_norm[(d.year, d.month)] = OverrideMes(valor=_num(val))

    memoria: list[dict] = []
    diferenca_futura = 0.0
    base_futura = 0.0
    soma_g = 0.0
    competencia_inicial = None
    if comp is not None and fim_vig is not None:
        competencia_inicial = _add_meses(comp, 1)
        limite = _eomonth(fim_vig, 0)
        mes = competencia_inicial
        idx = 0
        while mes <= limite:
            idx += 1
            ov = ov_norm.get((mes.year, mes.month))
            e = referencia_mensal
            if ov is not None and ov.valor is not None:
                if ov.ja_reajustado:
                    f_base = (ov.valor / fator) if fator else 0.0
                else:
                    f_base = ov.valor
                situacao = "Valor informado pelo fiscal"
            else:
                f_base = e
                situacao = "Projecao automatica"
            if saldo is None:
                g = f_base
            else:
                g = max(0.0, min(f_base, saldo - soma_g))
                if g < f_base - 1e-9:
                    situacao = "Limitado ao saldo"
            soma_g += g
            h = _round2(g * fator)
            i = _round2(h - g)
            diferenca_futura += i
            base_futura += g
            memoria.append({
                "indice": idx,
                "competencia": mes,
                "referencia_automatica": e,
                "override": (ov.valor if ov else None),
                "override_ja_reajustado": (ov.ja_reajustado if ov else False),
                "base_sem_reajuste": f_base,
                "base_considerada": g,
                "valor_reajustado": h,
                "diferenca": i,
                "situacao": situacao,
            })
            mes = _add_meses(mes, 1)

    complemento_estimado = retro + diferenca_futura

    # --- programacao por exercicio ---
    programacao: list[dict] = []
    if competencia_inicial is not None and fim_vig is not None:
        ano_ini = competencia_inicial.year
        for ano in range(ano_ini, fim_vig.year + 1):
            dif_ano = sum(m["diferenca"] for m in memoria if m["competencia"].year == ano)
            valor = dif_ano + (retro if ano == ano_ini else 0.0)
            programacao.append({
                "exercicio": ano,
                "valor": valor,
                "composicao": ("Retroativo + projecao futura" if ano == ano_ini
                               else "Projecao futura"),
            })
    soma_prog = sum(p["valor"] for p in programacao)
    if programacao and abs(soma_prog - complemento_estimado) > 0.005:
        checks.append("SOMA DOS EXERCICIOS DIVERGE DO COMPLEMENTO")

    if origem == ORIGEM_PCS and base_hist.get("pedidos_considerados", 0) == 0:
        checks.append("NENHUM PC CONSIDERADO NA JANELA")
    if not memoria and comp is not None and fim_vig is not None:
        checks.append("SEM MESES DE PROJECAO (data final anterior a competencia inicial)")

    return {
        "origem": origem,
        "percentual": perc,
        "fator": fator,
        "ultima_competencia": comp,
        "data_fim_vigencia": fim_vig,
        "janela_meses": janela,
        "saldo_contratual": saldo,
        "competencia_inicial_projecao": competencia_inicial,
        "base_historica": base_hist,
        "total_historico": base_hist.get("total_historico", 0.0),
        "media_mensal": referencia_mensal,
        "referencia_mensal": referencia_mensal,
        "referencia_reajustada": referencia_reajustada,
        "meses_projetados": len(memoria),
        "memoria_mensal": memoria,
        "base_futura": base_futura,
        "retroativo": retro,
        "diferenca_futura": diferenca_futura,
        "complemento_estimado": complemento_estimado,
        "programacao_por_exercicio": programacao,
        "soma_programacao": soma_prog,
        "status": ("OK" if not checks else "; ".join(checks)),
        "checks": checks,
    }
