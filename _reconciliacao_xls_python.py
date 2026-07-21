"""Reconciliacao entre a aba RESULTADOS do XLS e o motor Python.

Os intervalos nomeados de RESULTADOS sao referencia de AUDITORIA — nunca
substituem silenciosamente o calculo do motor. Cada campo comparavel e
classificado; divergencia relevante bloqueia documentos formais (via
politica de entrega segura), exibindo os dois valores sem escolher um.
"""
from __future__ import annotations

from typing import Any

TOLERANCIA_MONETARIA = 0.01  # R$ — maxima; regra mais restritiva prevalece

STATUS_CONCILIADO = "CONCILIADO"
STATUS_DENTRO_TOLERANCIA = "DIVERGENCIA_DENTRO_DA_TOLERANCIA"
STATUS_RELEVANTE = "DIVERGENCIA_RELEVANTE"
STATUS_SEM_CACHE = "RESULTADO_XLS_INDISPONIVEL_POR_CACHE"


def _num(valor: Any) -> float | None:
    try:
        if valor in (None, ""):
            return None
        return float(valor)
    except (TypeError, ValueError):
        return None


def _classificar(xls: float | None, py: float | None, tolerancia: float) -> str | None:
    if xls is None and py is None:
        return None  # campo não aplicável: nenhuma das fontes produziu valor
    if xls is None:
        return STATUS_SEM_CACHE
    if py is None:
        return None  # sem contraparte Python — comparacao nao aplicavel
    delta = abs(round(xls - py, 4))
    if delta == 0:
        return STATUS_CONCILIADO
    if delta <= tolerancia:
        return STATUS_DENTRO_TOLERANCIA
    return STATUS_RELEVANTE


# Etapa 5b: mapa de dependencia direta entre campos oficiais. Uma divergencia
# relevante em um campo torna nao-confiaveis o proprio campo e os que dependem
# diretamente dele. Retroativo e VTA/remanescente sao dimensoes independentes:
# uma divergencia so de retroativo nao contamina VTA, e vice-versa.
_DEPENDENTES_DIRETOS: dict[str, tuple[str, ...]] = {
    "QTD_REM_OFICIAL": ("REM_BASE_OFICIAL", "REM_ATUALIZADO_OFICIAL", "VTA_FINAL"),
    "REM_BASE_OFICIAL": ("REM_ATUALIZADO_OFICIAL", "VTA_FINAL"),
    "REM_ATUALIZADO_OFICIAL": ("VTA_FINAL",),
    "VTA_FINAL": (),
    "RETRO_FIN": ("RETRO_OFICIAL",),
    "RETRO_PC": ("RETRO_OFICIAL",),
    "RETRO_ITENS": ("RETRO_OFICIAL",),
    "RETRO_OFICIAL": (),
}


def campos_nao_confiaveis_para_documentos(reconciliacao: dict[str, Any] | None) -> set[str]:
    """Campos oficiais que NAO devem ser preenchidos nos documentos liberados.

    Recebe o dicionario de reconciliacao_xls_python e retorna o conjunto de
    campos canonicos (o proprio divergente + dependentes diretos) que precisam
    ficar vazios nos 3 documentos liberados apesar da divergencia (Etapa 5b).
    Nunca adota XLS nem Python; apenas marca o que deixar em branco.
    """
    nao_confiaveis: set[str] = set()
    for div in (reconciliacao or {}).get("divergencias_relevantes") or []:
        campo = str(div.get("campo") or "")
        if not campo:
            continue
        nao_confiaveis.add(campo)
        nao_confiaveis.update(_DEPENDENTES_DIRETOS.get(campo, ()))
    return nao_confiaveis


def reconciliar_xls_python(leitura: dict[str, Any]) -> dict[str, Any]:
    """Compara os valores nomeados de RESULTADOS com o motor Python."""
    resultado: dict[str, Any] = {
        "disponivel": False,
        "tolerancia": TOLERANCIA_MONETARIA,
        "campos": [],
        "divergencias_relevantes": [],
        "sem_cache": False,
        "status_geral": None,
    }
    rx = leitura.get("resultados_xls") or {}
    if not rx.get("disponivel"):
        return resultado
    resultado["disponivel"] = True
    valores = rx.get("valores") or {}
    resultado["sem_cache"] = bool(rx.get("cache_ausente"))

    tol_xls = _num(valores.get("TOLERANCIA_DIVERGENCIA"))
    tolerancia = TOLERANCIA_MONETARIA
    if tol_xls is not None and 0 < tol_xls < TOLERANCIA_MONETARIA:
        tolerancia = tol_xls  # regra existente mais restritiva prevalece
    resultado["tolerancia"] = tolerancia

    memoria = (leitura.get("objeto_processo") or {}).get("memoria_por_ciclo") or {}
    ciclos = memoria.get("ciclos") or []

    def _soma(metodo: str, campo: str) -> float:
        return round(sum(
            _num(((c.get("retroativo") or {}).get(metodo) or {}).get(campo)) or 0.0
            for c in ciclos
        ), 2)

    def _evid(metodo: str) -> int:
        return sum(
            int(((c.get("retroativo") or {}).get(metodo) or {}).get("evidencias") or 0)
            for c in ciclos
        )

    retro_fin_py = _soma("financeiro", "retroativo") if _evid("financeiro") else None
    retro_pc_py = _soma("pc", "retroativo") if _evid("pc") else None
    retro_itens_py = _soma("consumidos", "retroativo") if _evid("consumidos") else None

    # Retroativo oficial Python: por ciclo, primeiro metodo com evidencia
    # (hierarquia financeiro > pc > consumidos) — espelha a politica.
    retro_oficial_py: float | None = None
    parcelas_oficial = []
    for c in ciclos:
        retro = c.get("retroativo") or {}
        for metodo in ("financeiro", "pc", "consumidos"):
            reg = retro.get(metodo) or {}
            if int(reg.get("evidencias") or 0) > 0:
                parcelas_oficial.append(_num(reg.get("retroativo")) or 0.0)
                break
    if parcelas_oficial:
        retro_oficial_py = round(sum(parcelas_oficial), 2)

    qtd_rem_py: float | None = None
    rem_base_py: float | None = None
    rem_atualizado_py: float | None = None
    residuais = [c.get("residuais") or {} for c in ciclos]
    if any(int(r.get("itens") or 0) for r in residuais):
        qtd_rem_py = round(sum(_num(r.get("quantidade")) or 0.0 for r in residuais), 4)
        rem_base_py = round(sum(_num(r.get("valor_original")) or 0.0 for r in residuais), 2)
        rem_atualizado_py = round(
            sum(_num(r.get("valor_atualizado")) or 0.0 for r in residuais), 2
        )

    vta_py = _num((memoria.get("vta") or {}).get("valor_total_atualizado"))

    modo = str((leitura.get("controle") or {}).get("modo") or "").lower()
    compara_remanescente = modo != "pc"
    comparacoes = [
        ("RETRO_FIN", "Retroativo financeiro", retro_fin_py),
        ("RETRO_PC", "Retroativo por PCs", retro_pc_py),
        ("RETRO_ITENS", "Retroativo por itens", retro_itens_py),
        ("RETRO_OFICIAL", "Retroativo oficial", retro_oficial_py),
        ("QTD_REM_OFICIAL", "Quantidade remanescente oficial", qtd_rem_py if compara_remanescente else None),
        ("REM_BASE_OFICIAL", "Valor remanescente base", rem_base_py if compara_remanescente else None),
        ("REM_ATUALIZADO_OFICIAL", "Valor remanescente atualizado", rem_atualizado_py if compara_remanescente else None),
        ("VTA_FINAL", "VTA final", vta_py),
    ]

    for nome, rotulo, py in comparacoes:
        if nome not in (rx.get("nomes_presentes") or []):
            continue
        xls = _num(valores.get(nome))
        status = _classificar(xls, py, tolerancia)
        campo = {
            "campo": nome, "rotulo": rotulo, "xls": xls, "python": py,
            "status": status,
            "nota": (
                "Sem contraparte Python disponivel — comparacao nao aplicavel."
                if status is None else ""
            ),
        }
        resultado["campos"].append(campo)
        if status == STATUS_RELEVANTE:
            resultado["divergencias_relevantes"].append(campo)

    status_efetivos = [c["status"] for c in resultado["campos"] if c["status"]]
    if resultado["divergencias_relevantes"]:
        resultado["status_geral"] = STATUS_RELEVANTE
    elif resultado["sem_cache"]:
        resultado["status_geral"] = STATUS_SEM_CACHE
    elif status_efetivos and all(s == STATUS_SEM_CACHE for s in status_efetivos):
        resultado["status_geral"] = STATUS_SEM_CACHE
    elif any(s == STATUS_DENTRO_TOLERANCIA for s in status_efetivos):
        resultado["status_geral"] = STATUS_DENTRO_TOLERANCIA
    elif any(s == STATUS_CONCILIADO for s in status_efetivos):
        resultado["status_geral"] = STATUS_CONCILIADO
    return resultado
