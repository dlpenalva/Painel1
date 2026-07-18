from __future__ import annotations

from typing import Any


STATUS_INSUFICIENTE = "INFORMACAO_INSUFICIENTE"
STATUS_PARCIAL = "APURACAO_PARCIAL"
STATUS_BLOQUEADO = "BLOQUEADO_PARA_FORMALIZACAO"
STATUS_VALIDACAO = "PRONTO_PARA_VALIDACAO_FISCAL"
STATUS_FORMALIZACAO = "APTO_PARA_FORMALIZACAO"

_METODOS = ("financeiro", "pc", "consumidos")


def _numero(valor: Any) -> float | None:
    try:
        if valor in (None, ""):
            return None
        return float(valor)
    except (TypeError, ValueError):
        return None


def _sim(valor: Any) -> bool:
    return str(valor or "").strip().lower() in {"sim", "s", "yes", "true", "1"}


def _unicos(valores: list[str]) -> list[str]:
    saida: list[str] = []
    for valor in valores:
        texto = str(valor or "").strip()
        if texto and texto not in saida:
            saida.append(texto)
    return saida


def avaliar_entrega_segura(
    leitura: dict[str, Any],
    confirmacao_fiscal: bool = False,
    confirmacao_gcc: bool = False,
) -> dict[str, Any]:
    """Classifica o que pode ser exibido ou formalizado sem inventar lacunas."""
    if not isinstance(leitura, dict) or not leitura.get("ok"):
        return {
            "status": STATUS_INSUFICIENTE,
            "pode_confirmar": False,
            "pode_formalizar": False,
            "bloqueios": ["O arquivo não foi lido com sucesso."],
            "pendencias": [],
            "indices": [],
            "retroativo": {},
            "remanescentes": {},
        }

    objeto = leitura.get("objeto_processo") or {}
    memoria = objeto.get("memoria_por_ciclo") or {}
    por_memoria = {
        str(c.get("ciclo") or "").upper(): c
        for c in memoria.get("ciclos") or []
    }
    parametros = (leitura.get("parametros_v10") or {}).get("por_ciclo") or {}

    indices: list[dict[str, Any]] = []
    ciclos_apuracao: list[str] = []
    ciclos_indice_incompleto: list[str] = []
    for ciclo in ("C1", "C2", "C3", "C4"):
        reg = parametros.get(ciclo) or {}
        computar_raw = reg.get("computar_nesta_apuracao")
        computar = _sim(computar_raw) if computar_raw not in (None, "") else False
        if not computar:
            continue
        ciclos_apuracao.append(ciclo)
        indice = _numero(reg.get("percentual_reajuste"))
        if indice is None:
            indice = _numero(reg.get("fator_proprio"))
        fator = _numero(reg.get("fator_acumulado"))
        completo = indice is not None and fator is not None and fator > 0
        if not completo:
            ciclos_indice_incompleto.append(ciclo)
        indices.append({
            "ciclo": ciclo,
            "indice_ciclo": indice,
            "fator_acumulado": fator,
            "indice_acumulado": fator - 1.0 if fator is not None else None,
            "completo": completo,
        })

    evidencias_por_metodo = {metodo: 0 for metodo in _METODOS}
    for ciclo in por_memoria.values():
        retro = ciclo.get("retroativo") or {}
        for metodo in _METODOS:
            evidencias_por_metodo[metodo] += int(
                (retro.get(metodo) or {}).get("evidencias") or 0
            )
    retro_ciclos: list[dict[str, Any]] = []
    ciclos_sem_evidencia: list[str] = []
    for ciclo in ciclos_apuracao:
        retro_metodos = (por_memoria.get(ciclo) or {}).get("retroativo") or {}
        metodo_ciclo = next((
            m for m in _METODOS
            if int((retro_metodos.get(m) or {}).get("evidencias") or 0) > 0
        ), None)
        if metodo_ciclo is None:
            ciclos_sem_evidencia.append(ciclo)
            continue
        reg = retro_metodos.get(metodo_ciclo) or {}
        retro_ciclos.append({
            "ciclo": ciclo,
            "metodo": metodo_ciclo,
            "evidencias": int(reg.get("evidencias") or 0),
            "base_original": _numero(reg.get("base_original")),
            "valor_atualizado": _numero(reg.get("valor_atualizado")),
            "retroativo": _numero(reg.get("retroativo")),
        })
    metodos_usados = _unicos([r["metodo"] for r in retro_ciclos])
    metodo = (
        metodos_usados[0] if len(metodos_usados) == 1
        else "misto_por_ciclo" if metodos_usados else None
    )
    subtotal_retroativo = round(
        sum(r.get("retroativo") or 0.0 for r in retro_ciclos), 2
    ) if retro_ciclos else None

    residuais = [
        (por_memoria.get(ciclo) or {}).get("residuais") or {}
        for ciclo in ("C0", "C1", "C2", "C3", "C4")
    ]
    itens_residuais = sum(int(r.get("itens") or 0) for r in residuais)
    valor_residual = round(sum(_numero(r.get("valor_atualizado")) or 0.0 for r in residuais), 2)
    potencial = memoria.get("vta") or {}
    vta = _numero(potencial.get("valor_total_atualizado"))

    pend_obj = objeto.get("pendencias") or {}
    bloqueios = list(pend_obj.get("bloqueantes") or [])
    pendencias = list(pend_obj.get("advertencias") or [])
    pendencias.extend(pend_obj.get("ressalvas") or [])

    reconciliacao = leitura.get("reconciliacao") or {}
    if reconciliacao.get("bloqueia_formalizacao"):
        bloqueios.append("Há divergência material de execução pendente de decisão da GCC.")
    rec_evidencias = leitura.get("reconciliacao_evidencias_sombra") or {}
    if rec_evidencias.get("divergencias_materiais"):
        bloqueios.append("Há divergência material entre evidências ainda não equalizada pela GCC.")
    composicao = leitura.get("composicao_vta") or {}
    if composicao.get("bloqueia_formalizacao"):
        bloqueios.append("A composição do valor contratual possui diferença material pendente.")
    posicao = leitura.get("posicao_contratual") or {}
    if posicao.get("cache_ausente"):
        bloqueios.append(
            "A aba posicao_contratual não tem valores calculados: o arquivo não foi "
            "recalculado pelo Excel. Abra o XLS no Excel, salve e reenvie antes de "
            "qualquer formalização."
        )
    reconc_xls = leitura.get("reconciliacao_xls_python") or {}
    for div in reconc_xls.get("divergencias_relevantes") or []:
        bloqueios.append(
            f"Divergência relevante XLS × Python em {div.get('campo')} "
            f"({div.get('rotulo')}): XLS={div.get('xls')} vs Python={div.get('python')}. "
            "Nenhum dos resultados é adotado automaticamente; equalizar antes de formalizar."
        )
    if ciclos_indice_incompleto:
        bloqueios.append(
            "Índice ou fator acumulado ausente em: " + ", ".join(ciclos_indice_incompleto) + "."
        )
    if not ciclos_apuracao:
        bloqueios.append("Nenhum ciclo foi marcado para esta apuração.")

    if ciclos_sem_evidencia:
        pendencias.append(
            "Sem evidência de execução registrada em " + ", ".join(ciclos_sem_evidencia)
            + "; o fiscal deve confirmar que não houve execução ou complementar o XLS."
        )
    if not itens_residuais:
        pendencias.append(
            "Remanescentes não demonstrados; o valor futuro e o VTA permanecem indisponíveis."
        )
    if metodo is None:
        pendencias.append(
            "Nenhuma evidência financeira, PC pago definitivo ou consumo pago permite apurar retroativo."
        )

    bloqueios = _unicos(bloqueios)
    pendencias = _unicos(pendencias)
    pode_confirmar = bool(
        ciclos_apuracao and not ciclos_indice_incompleto and metodo and not bloqueios
    )
    pode_formalizar = bool(pode_confirmar and confirmacao_fiscal and confirmacao_gcc)

    if pode_formalizar:
        status = STATUS_FORMALIZACAO
    elif bloqueios:
        status = STATUS_BLOQUEADO
    elif not ciclos_apuracao or metodo is None:
        status = STATUS_INSUFICIENTE
    elif ciclos_sem_evidencia or not itens_residuais or vta is None:
        status = STATUS_PARCIAL
    else:
        status = STATUS_VALIDACAO

    nao_divulgar: list[str] = []
    if not pode_formalizar:
        nao_divulgar.extend([
            "retroativo total como valor definitivo",
            "VTA como valor formalizado",
            "e-mail final ao fornecedor",
            "documentos destinados à formalização",
        ])
    if vta is None:
        nao_divulgar.append("valor total atualizado por ausência de remanescente seguro")

    return {
        "status": status,
        "pode_confirmar": pode_confirmar,
        "pode_formalizar": pode_formalizar,
        "confirmacao_fiscal": bool(confirmacao_fiscal),
        "confirmacao_gcc": bool(confirmacao_gcc),
        "bloqueios": bloqueios,
        "pendencias": pendencias,
        "indices": indices,
        "ciclos_apuracao": ciclos_apuracao,
        "retroativo": {
            "metodo": metodo,
            "evidencias_por_metodo": evidencias_por_metodo,
            "ciclos_com_evidencia": retro_ciclos,
            "ciclos_sem_evidencia": ciclos_sem_evidencia,
            "subtotal_evidencias_recebidas": subtotal_retroativo,
            "total_definitivo": subtotal_retroativo if pode_formalizar else None,
        },
        "remanescentes": {
            "itens_com_fotografia": itens_residuais,
            "valor_atualizado_identificado": valor_residual if itens_residuais else None,
            "vta_calculado": vta,
            "vta_divulgavel": vta if pode_formalizar else None,
        },
        "nao_divulgar": _unicos(nao_divulgar),
    }
