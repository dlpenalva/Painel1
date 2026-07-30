"""Motor da aba comparativo_VTA (Coleta Oficial; nao aparece na web).

BLOCO 1 — EVOLUCAO DO REMANESCENTE A PRECOS ORIGINAIS
    Por item e por ciclo: VU_ORIGINAL x QTD_REM_AJUSTADA_Cn (da posicao_contratual,
    respeitando aditivos). SEM reajuste. Mostra so a evolucao do saldo fisico/
    economico a precos originais.

BLOCO 2 — COMPARACAO DE REFERENCIA (NAO E VTA NEM BASE DE PAGAMENTO)
    VTA pelo metodo aplicado  x  contrato original integralmente reajustado.
    O contrafactual usa o FATOR ACUMULADO HISTORICO INTEGRAL do contrato ate o
    ciclo atual (inclui o historico de ciclos ja formalizados, ex.: C1), nunca
    apenas o fator dos ciclos da apuracao corrente.
"""
from __future__ import annotations

from typing import Any

CICLOS = ("C0", "C1", "C2", "C3", "C4")

ALERTA_CONTRAFACTUAL = (
    "O valor integralmente reajustado e apenas um cenario comparativo. Ele "
    "pressupoe que todo o contrato original permanecesse sujeito ao fator "
    "acumulado final, desconsiderando a execucao ocorrida em ciclos distintos "
    "e a reducao do saldo remanescente. Nao utilizar este valor como VTA, "
    "retroativo ou base de pagamento."
)
TITULO_BLOCO2 = (
    "COMPARACAO DE REFERENCIA - NAO UTILIZAR COMO VTA OU VALOR DE PAGAMENTO"
)


def _tofl(v: Any) -> float:
    try:
        if v in (None, ""):
            return 0.0
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _fator_historico_integral(leitura: dict[str, Any], vigente: str) -> float | None:
    """Fator acumulado historico integral ate o ciclo vigente.

    Prioriza parametros!F do ciclo vigente (acumulacao de TODOS os ciclos, nao
    so os da apuracao). Fallback: VU_vigente/VU_C0 do historico_VU.
    """
    por_ciclo = (leitura.get("parametros_v10") or {}).get("por_ciclo") or {}
    fator = _tofl((por_ciclo.get(vigente) or {}).get("fator_acumulado"))
    if fator:
        return fator
    for h in (leitura.get("historico_vu") or {}).get("itens") or []:
        vu = h.get("vu_ciclos") or {}
        vu0 = _tofl(vu.get("VU_C0"))
        vuv = _tofl(vu.get(f"VU_{vigente}"))
        if vu0 and vuv:
            return round(vuv / vu0, 10)
    return None


def montar_comparativo_vta(leitura: dict[str, Any]) -> dict[str, Any]:
    posc = ((leitura.get("posicao_contratual") or {}).get("itens")) or []
    vigente = str((leitura.get("controle") or {}).get("ciclo_vigente")
                  or "C4").strip().upper()

    # BLOCO 1 — precos originais, por item e ciclo.
    itens: list[dict[str, Any]] = []
    totais = {c: 0.0 for c in CICLOS}
    for p in posc:
        vu = _tofl(p.get("VU_ORIGINAL"))
        linha: dict[str, Any] = {"item": p.get("ITEM"), "vu_original": vu,
                                 "ciclos": {}}
        for c in CICLOS:
            q = _tofl(p.get(f"QTD_REM_AJUSTADA_{c}"))
            val = round(vu * q, 2)
            linha["ciclos"][c] = {"qtd": q, "valor": val}
            totais[c] = round(totais[c] + val, 2)
        itens.append(linha)

    # BLOCO 2 — comparacao de referencia.
    from _motor_composicao_vta import montar_composicao_vta
    vta = montar_composicao_vta(leitura).get("vta_composicao")
    fator = _fator_historico_integral(leitura, vigente)
    valor_original = round(sum(
        _tofl(p.get("VU_ORIGINAL")) * _tofl(p.get("QTD_BASE_ORIGINAL"))
        for p in posc
    ), 2)
    contrato_cheio = round(valor_original * fator, 2) if fator else None

    return {
        "disponivel": bool(posc),
        "ciclos": list(CICLOS),
        "bloco1_itens": itens,
        "bloco1_totais_por_ciclo": totais,
        "bloco2": {
            "titulo": TITULO_BLOCO2,
            "vta_metodo_aplicado": vta,
            "valor_original_contrato": valor_original,
            "fator_historico_integral": fator,
            "contrato_cheio_reajustado": contrato_cheio,
            "alerta": ALERTA_CONTRAFACTUAL,
            "e_vta": False,
        },
    }
