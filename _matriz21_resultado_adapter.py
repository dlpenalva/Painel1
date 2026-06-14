# -*- coding: utf-8 -*-
"""
Adaptador de resultado — Matriz 2.1 → resultado compatível com documentos — v2

Objetivo:
- Enriquecer o resultado da Matriz 2.1 para os documentos oficiais.
- Melhorar PDF/Relatório Executivo, que na v1 funcionava, mas ficava pobre em campos como:
  índice, ciclos, valores executivos e composição.
- Não altera cálculo do VTA.
- Não substitui Matriz 2.0/v10.

Uso:
    from _matriz21_resultado_adapter import adaptar_resultado_matriz21
"""

from __future__ import annotations

from copy import deepcopy
from typing import Any, Dict, List

try:
    import pandas as pd
except Exception:  # pragma: no cover
    pd = None


def _num(v: Any, default: float = 0.0) -> float:
    if v is None:
        return default
    if isinstance(v, bool):
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s:
        return default
    s = s.replace("R$", "").replace(" ", "")
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return default


def _df(obj: Any):
    if pd is None:
        return obj
    if obj is None:
        return pd.DataFrame()
    if isinstance(obj, pd.DataFrame):
        return obj.copy()
    if isinstance(obj, list):
        return pd.DataFrame(obj)
    if isinstance(obj, dict):
        return pd.DataFrame([obj])
    return pd.DataFrame()


def _rows(obj: Any) -> List[Dict[str, Any]]:
    if obj is None:
        return []
    if pd is not None and isinstance(obj, pd.DataFrame):
        return obj.to_dict(orient="records")
    if isinstance(obj, list):
        return [x for x in obj if isinstance(x, dict)]
    if isinstance(obj, dict):
        return [obj]
    return []


def _moeda(v: Any) -> str:
    n = _num(v, 0.0)
    s = f"R$ {n:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _pct(v: Any) -> str:
    n = _num(v, 0.0)
    if abs(n) <= 1:
        n *= 100
    return f"{n:.2f}%".replace(".", ",")


def _data_br(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    # YYYY-MM-DD
    if len(s) >= 10 and s[4:5] == "-" and s[7:8] == "-":
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"
    return s


def _periodo_br(inicio: Any, fim: Any) -> str:
    a = str(inicio or "").strip()
    b = str(fim or "").strip()
    if len(a) >= 7 and a[4:5] == "-":
        a_fmt = f"{a[5:7]}/{a[0:4]}"
    else:
        a_fmt = a
    if len(b) >= 7 and b[4:5] == "-":
        b_fmt = f"{b[5:7]}/{b[0:4]}"
    else:
        b_fmt = b
    return f"{a_fmt} a {b_fmt}" if a_fmt or b_fmt else ""


def _maior_fator(ciclos: Dict[str, Any]) -> float:
    fatores = []
    if isinstance(ciclos, dict):
        for info in ciclos.values():
            if isinstance(info, dict):
                fatores.append(_num(info.get("fator_acumulado"), 1.0))
    return max(fatores) if fatores else 1.0


def _montar_alertas(validacoes: List[Dict[str, Any]]) -> List[str]:
    alertas = []
    for v in validacoes or []:
        resultado = str(v.get("Resultado", "")).upper()
        nome = str(v.get("Validação", "")).strip()
        detalhe = str(v.get("Detalhe", "")).strip()
        if resultado in {"ERRO", "ALERTA", "DIVERGENTE"}:
            msg = f"{resultado}: {nome}"
            if detalhe:
                msg += f" — {detalhe}"
            alertas.append(msg)
    return alertas


def _valor_componente(memoria: List[Dict[str, Any]], nome: str) -> float:
    for row in memoria or []:
        if str(row.get("Componente", "")).strip() == nome:
            return _num(row.get("Valor"), 0.0)
    return 0.0


def _montar_df_ciclos(ciclos: Dict[str, Any]):
    rows = []
    if not isinstance(ciclos, dict):
        return _df(rows)

    for ciclo, info in ciclos.items():
        info = info or {}
        if not isinstance(info, dict):
            info = {}
        percentual = _num(info.get("percentual"), 0.0)
        fator_acumulado = _num(info.get("fator_acumulado"), 1.0)
        rows.append({
            "Ciclo": ciclo,
            "Data-base": _data_br(info.get("inicio")),
            "Data do pedido": "não se aplica" if ciclo == "C0" else "",
            "Classificação": "base sem reajuste" if ciclo == "C0" else "Matriz 2.1",
            "Percentual aplicado": percentual,
            "Percentual aplicado (%)": _pct(percentual),
            "Fator acumulado": fator_acumulado,
            "Período": _periodo_br(info.get("inicio"), info.get("fim")),
            "Observação": (
                "Período inicial do contrato, sem aplicação de reajuste."
                if ciclo == "C0"
                else "Ciclo informado na Matriz 2.1."
            ),
        })

    return _df(rows)


def _montar_composicao_legivel(memoria: List[Dict[str, Any]]):
    rows = []
    for row in memoria or []:
        comp = str(row.get("Componente", "")).strip()
        if not comp:
            continue
        rows.append({
            "Componente": comp,
            "Fonte": row.get("Fonte", ""),
            "Valor": _num(row.get("Valor"), 0.0),
            "Valor formatado": _moeda(row.get("Valor")),
            "Status": row.get("Status", ""),
            "Observação": row.get("Observação", ""),
        })
    return _df(rows)


def adaptar_resultado_matriz21(res_m21: Dict[str, Any]) -> Dict[str, Any]:
    """
    Converte resultado bruto da Matriz 2.1 para o formato esperado pelos documentos.

    Não recalcula o VTA. Apenas mapeia e enriquece as chaves já calculadas pelo leitor.
    """
    r = deepcopy(res_m21 or {})

    total = round(_num(r.get("total_vta_m21", r.get("valor_atualizado_contrato", 0.0))), 2)
    total_aditivos = round(_num(r.get("total_aditivos", 0.0)), 2)

    ciclos = r.get("ciclos", {}) or {}
    entradas = r.get("entradas_criticas", {}) or {}
    memoria = r.get("memoria_vta_m21", []) or []
    validacoes = r.get("validacoes_m21", []) or []
    conciliacao = r.get("conciliacao_referencia", []) or []

    valor_c0 = _valor_componente(memoria, "C0")
    valor_c1 = _valor_componente(memoria, "C1")
    valor_c2 = _valor_componente(memoria, "C2")
    valor_c3 = _valor_componente(memoria, "C3 executado")
    valor_c4 = _valor_componente(memoria, "C4 executado")
    saldo_rem = _valor_componente(memoria, "Saldo remanescente após última competência")
    valor_executado_atualizado = round(valor_c0 + valor_c1 + valor_c2 + valor_c3 + valor_c4, 2)

    df_comp = _montar_composicao_legivel(memoria)
    df_valid = _df(r.get("df_validacoes_m21", validacoes))
    df_conc = _df(r.get("df_conciliacao_referencia", conciliacao))

    df_fin = _df(r.get("financeiro_linhas", []))
    df_itens = _df(r.get("itens_linhas", []))
    df_adit = _df(r.get("aditivos_linhas", []))
    df_ciclos = _montar_df_ciclos(ciclos)

    fator_acum = _maior_fator(ciclos)
    variacao = round(fator_acum - 1.0, 6)

    alertas = _montar_alertas(_rows(validacoes))

    indice_nome = entradas.get("Índice aplicado") or entradas.get("Indice aplicado") or "Não informado"

    params_v10 = {
        "origem": "matriz_2_1_experimental",
        "indice": indice_nome,
        "indice_aplicado": indice_nome,
        "data_inicio_c0": entradas.get("Data início C0 / início contratual"),
        "data_corte": entradas.get("Data de corte da apuração"),
        "ha_ciclo_em_execucao": entradas.get("Há ciclo em execução?"),
        "ciclo_em_execucao": entradas.get("Ciclo em execução"),
        "ultima_competencia_financeira": entradas.get("Última competência financeira informada"),
        "forma_c0": entradas.get("Forma de informar C0"),
        "forma_saldo": entradas.get("Forma de informar saldo remanescente"),
        "marco_saldo": entradas.get("Marco do saldo remanescente"),
        "metodologia_corte": "Matriz 2.1 — execução atualizada por ciclos + saldo remanescente + aditivos/supressões.",
    }

    resultado = {
        "ok": True,
        "versao": "matriz_2_1",
        "origem_coleta": "matriz_2_1_experimental",
        "tipo": "Matriz 2.1",
        "modo_detectado": "Matriz 2.1 experimental",
        "modo_calculo": "Matriz 2.1 experimental",
        "arquivo_origem": r.get("arquivo", ""),

        # Identificação
        "indice": indice_nome,
        "indice_aplicado": indice_nome,
        "data_corte": entradas.get("Data de corte da apuração"),
        "metodologia_corte": "Matriz 2.1 — execução atualizada por ciclos + saldo remanescente + aditivos/supressões.",

        # Chaves centrais consumidas pela UI/documentos
        "valor_atualizado_contrato": total,
        "valor_global_estoque": total,
        "valor_global_atualizado": total,
        "valor_global_contrato": total,
        "valor_represado_a_pagar": 0.0,

        # Indicadores executivos auxiliares
        "valor_original": valor_c0,
        "valor_pago_efetivo": valor_executado_atualizado,
        "valor_teorico_calculado": valor_executado_atualizado,
        "valor_executado_atualizado_ciclos": valor_executado_atualizado,
        "saldo_remanescente_atualizado": saldo_rem,
        "valor_total_aditivos_supressoes": total_aditivos,
        "total_aditivos_supressoes": total_aditivos,
        "total_aditivos_atualizados": total_aditivos,

        # Fatores/ciclos
        "fator_acumulado": fator_acum,
        "variacao_acumulada": variacao,
        "ciclos": ciclos,
        "df_ciclos": df_ciclos,

        # Bases lidas
        "parametros": entradas,
        "params": entradas,
        "params_v10": params_v10,
        "financeiro": {
            "linhas": r.get("financeiro_linhas", []),
            "totais": r.get("financeiro_totais", {}),
            "total_delta": 0.0,
        },
        "itens": {
            "linhas": r.get("itens_linhas", []),
            "totais": r.get("itens_totais", {}),
        },
        "aditivos": {
            "linhas": r.get("aditivos_linhas", []),
            "total": total_aditivos,
        },
        "historico": [],

        # DataFrames auxiliares
        "df_composicao_valor_total": df_comp,
        "df_financeiro": df_fin,
        "df_execucao_financeira": df_fin,
        "df_itens": df_itens,
        "df_aditivos": df_adit,

        # Memórias e validações específicas
        "memoria_vta_m21": memoria,
        "total_vta_m21": total,
        "validacoes_m21": validacoes,
        "conciliacao_referencia": conciliacao,
        "df_validacoes_m21": df_valid,
        "df_conciliacao_referencia": df_conc,

        # Alertas
        "alertas": alertas,
        "_alertas_info": alertas,
        "_matriz21_resultado_bruto": r,
        "_adaptado_para_documentos": True,
    }

    return resultado


criar_resultado_documental_matriz21 = adaptar_resultado_matriz21

# >>> PATCH_M21_RETROATIVO_DOCS_V4
# Enriquecimento documental da Matriz 2.1:
# - calcula retroativo financeiro quando houver valor_base/valor_atualizado;
# - preserva o VTA calculado pelo leitor;
# - melhora chaves consumidas por cards, PDF, DOCX, TXT e XLSX.
def _m21v4_num(v, default=0.0):
    try:
        if v is None or isinstance(v, bool):
            return default
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip().replace("R$", "").replace(" ", "")
        if not s:
            return default
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        return float(s)
    except Exception:
        return default


def _m21v4_rows(obj):
    if obj is None:
        return []
    try:
        import pandas as _pd_m21v4
        if isinstance(obj, _pd_m21v4.DataFrame):
            return obj.to_dict(orient="records")
    except Exception:
        pass
    if isinstance(obj, list):
        return [x for x in obj if isinstance(x, dict)]
    if isinstance(obj, dict):
        if isinstance(obj.get("linhas"), list):
            return [x for x in obj.get("linhas", []) if isinstance(x, dict)]
        return [obj]
    return []


def _m21v4_get(row, candidatos):
    if not isinstance(row, dict):
        return None
    mapa = {str(k).lower().strip(): k for k in row.keys()}
    for cand in candidatos:
        k = mapa.get(str(cand).lower().strip())
        if k is not None:
            return row.get(k)
    for k in row.keys():
        kn = str(k).lower().strip()
        if any(str(c).lower().strip() in kn for c in candidatos):
            return row.get(k)
    return None


def _m21v4_financeiro_rows(bruto, resultado):
    fontes = []
    if isinstance(bruto, dict):
        fontes.extend([
            bruto.get("financeiro_linhas"),
            bruto.get("financeiro"),
            bruto.get("df_financeiro"),
            bruto.get("df_execucao_financeira"),
        ])
    if isinstance(resultado, dict):
        fontes.extend([
            resultado.get("df_financeiro"),
            resultado.get("df_execucao_financeira"),
            resultado.get("financeiro"),
        ])
    rows = []
    for fonte in fontes:
        rows.extend(_m21v4_rows(fonte))
    return rows


def _m21v4_calcular_retroativo(bruto, resultado):
    total = 0.0
    for row in _m21v4_financeiro_rows(bruto, resultado):
        base = _m21v4_get(row, [
            "valor_base", "valor base", "valor original", "valor_original",
            "valor bruto", "valor_bruto", "valor reconhecido", "valor_reconhecido",
            "valor medido", "valor_medido", "valor pago", "valor_pago"
        ])
        atualizado = _m21v4_get(row, [
            "valor_atualizado", "valor atualizado", "valor reajustado", "valor_reajustado",
            "valor corrigido", "valor_corrigido", "valor com reajuste"
        ])
        b = _m21v4_num(base, None)
        a = _m21v4_num(atualizado, None)
        if b is None or a is None:
            continue
        total += (a - b)
    if abs(total) <= 0.004 and isinstance(resultado, dict):
        fin = resultado.get("financeiro") if isinstance(resultado.get("financeiro"), dict) else {}
        total = _m21v4_num(fin.get("total_delta"), _m21v4_num(resultado.get("valor_retroativo"), 0.0))
    return round(total, 2)


def _m21v4_enriquecer(resultado, bruto=None):
    if not isinstance(resultado, dict):
        return resultado
    bruto = bruto if isinstance(bruto, dict) else resultado.get("_matriz21_resultado_bruto", {})
    retro = _m21v4_calcular_retroativo(bruto, resultado)
    if abs(retro) > 0.004:
        resultado["valor_represado_a_pagar"] = retro
        resultado["valor_retroativo"] = retro
        resultado["retroativo"] = retro
        resultado["retroativo_total"] = retro
        resultado["total_retroativo"] = retro
        if isinstance(resultado.get("financeiro"), dict):
            resultado["financeiro"]["total_delta"] = retro
    entradas = bruto.get("entradas_criticas", {}) if isinstance(bruto, dict) else {}
    if isinstance(entradas, dict):
        idx = (
            entradas.get("Índice aplicado") or entradas.get("Indice aplicado") or
            entradas.get("indice") or entradas.get("Índice")
        )
        if idx and str(resultado.get("indice", "")).strip().lower() in {"", "não informado", "nao informado", "none"}:
            resultado["indice"] = idx
            resultado["indice_aplicado"] = idx
    resultado["origem_coleta"] = "matriz_2_1"
    resultado["tipo"] = "Matriz 2.1"
    resultado["ok"] = bool(resultado.get("ok", True))
    return resultado


try:
    _adaptar_resultado_matriz21_base_v4 = adaptar_resultado_matriz21

    def adaptar_resultado_matriz21(res_m21):
        resultado = _adaptar_resultado_matriz21_base_v4(res_m21)
        return _m21v4_enriquecer(resultado, res_m21)

    criar_resultado_documental_matriz21 = adaptar_resultado_matriz21
except Exception:
    pass
# <<< PATCH_M21_RETROATIVO_DOCS_V4
# >>> PATCH_M21_INDICE_ADAPTER_V5
try:
    _adaptar_resultado_matriz21_base_v5 = adaptar_resultado_matriz21

    def _m21_buscar_indice_recursivo_v5(obj):
        nomes_validos = {
            "indice",
            "índice",
            "indice_utilizado",
            "índice utilizado",
            "indice contratual",
            "índice contratual",
            "tipo_indice",
            "tipo índice",
            "indice_nome",
        }
        marcadores = ("IPCA", "IST", "ICTI", "INPC", "IGP", "IGP-M", "INCC")
        if isinstance(obj, dict):
            for chave, valor in obj.items():
                chave_norm = str(chave).strip().lower()
                if chave_norm in nomes_validos and valor not in (None, ""):
                    return str(valor).strip()
            for valor in obj.values():
                achado = _m21_buscar_indice_recursivo_v5(valor)
                if achado:
                    return achado
        elif isinstance(obj, (list, tuple)):
            for i, valor in enumerate(obj):
                if isinstance(valor, str) and valor.strip().lower() in nomes_validos:
                    if i + 1 < len(obj) and obj[i + 1] not in (None, ""):
                        return str(obj[i + 1]).strip()
                    if i + 2 < len(obj) and obj[i + 2] not in (None, ""):
                        return str(obj[i + 2]).strip()
                achado = _m21_buscar_indice_recursivo_v5(valor)
                if achado:
                    return achado
        elif isinstance(obj, str):
            texto = obj.strip()
            up = texto.upper()
            if any(m in up for m in marcadores) and len(texto) <= 80:
                return texto
        return ""

    def adaptar_resultado_matriz21(diag):
        resultado = _adaptar_resultado_matriz21_base_v5(diag)
        if not isinstance(resultado, dict):
            resultado = {"ok": True, "valor_atualizado_contrato": 0, "origem_coleta": "matriz_2_1"}
        indice = (
            resultado.get("indice")
            or resultado.get("indice_utilizado")
            or resultado.get("indice_contratual")
            or _m21_buscar_indice_recursivo_v5(diag)
        )
        if indice:
            resultado["indice"] = indice
            resultado["indice_utilizado"] = indice
            resultado["indice_contratual"] = indice
        resultado["origem_coleta"] = "matriz_2_1"
        resultado["tipo"] = resultado.get("tipo") or "Matriz 2.1"
        resultado["ok"] = bool(resultado.get("ok", True))
        return resultado
except Exception:
    pass
# <<< PATCH_M21_INDICE_ADAPTER_V5

# >>> PATCH_M21_INDICE_CARD_V6_ADAPTER
# Garante que o resultado consolidado usado pelos cards/documentos receba o indice da Matriz 2.1.
def _m21_v6_normalizar_indice_resultado(valor):
    if valor is None:
        return ""
    texto = str(valor).strip()
    if not texto or texto.lower() in {"não informado", "nao informado", "none", "nan"}:
        return ""
    up = texto.upper()
    mapa = [
        ("ICTI", "ICTI"),
        ("IST", "IST"),
        ("IPCA", "IPCA"),
        ("INPC", "INPC"),
        ("IGP-M", "IGP-M"),
        ("IGPM", "IGP-M"),
    ]
    for chave, nome in mapa:
        if chave in up:
            return nome
    return ""


def _m21_v6_buscar_indice(obj, profundidade=0):
    if profundidade > 6:
        return ""
    if isinstance(obj, dict):
        for chave, valor in obj.items():
            chave_txt = str(chave).lower()
            if "indice" in chave_txt or "índice" in chave_txt:
                achou = _m21_v6_normalizar_indice_resultado(valor)
                if achou:
                    return achou
            achou = _m21_v6_buscar_indice(valor, profundidade + 1)
            if achou:
                return achou
    elif isinstance(obj, (list, tuple)):
        for valor in obj:
            achou = _m21_v6_buscar_indice(valor, profundidade + 1)
            if achou:
                return achou
    else:
        return _m21_v6_normalizar_indice_resultado(obj)
    return ""


try:
    _adaptar_resultado_matriz21_base_v6 = adaptar_resultado_matriz21
except NameError:
    _adaptar_resultado_matriz21_base_v6 = None


def adaptar_resultado_matriz21(diag):
    if _adaptar_resultado_matriz21_base_v6 is None:
        resultado = dict(diag or {}) if isinstance(diag, dict) else {}
    else:
        resultado = _adaptar_resultado_matriz21_base_v6(diag)
        if not isinstance(resultado, dict):
            resultado = {}

    indice_atual = _m21_v6_normalizar_indice_resultado(
        resultado.get("indice")
        or resultado.get("indice_utilizado")
        or resultado.get("indice_nome")
        or (resultado.get("parametros") or {}).get("indice") if isinstance(resultado.get("parametros"), dict) else ""
    )
    indice_diag = _m21_v6_buscar_indice(diag)
    indice = indice_atual or indice_diag

    if indice:
        resultado["indice"] = indice
        resultado["indice_utilizado"] = indice
        resultado["indice_nome"] = indice
        params = resultado.get("parametros")
        if not isinstance(params, dict):
            params = {}
        params["indice"] = indice
        params["indice_utilizado"] = indice
        resultado["parametros"] = params

    return resultado
# <<< PATCH_M21_INDICE_CARD_V6_ADAPTER

