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
