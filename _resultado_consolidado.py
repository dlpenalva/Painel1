"""Camada pura de consolidação dos resultados já calculados.

Este módulo não calcula grandezas financeiras. Ele apenas seleciona fontes
canônicas existentes, preserva a diferença entre ausência e zero e prepara um
contrato estável para a interface de Upload e resultados.
"""

from __future__ import annotations

from numbers import Real
from typing import Any


STATUS_CONFIAVEL = "CONFIÁVEL"
STATUS_RESSALVAS = "CONFIÁVEL COM RESSALVAS"
STATUS_PENDENTE = "PENDENTE DE CONFIRMAÇÃO"
STATUS_BLOQUEADO = "BLOQUEADO"

_STATUS_POLITICA_INCOMPLETOS = {
    "INFORMACAO_INSUFICIENTE",
    "APURACAO_PARCIAL",
}

_ROTULOS_METODO = {
    "financeiro": "Financeiro",
    "pc": "Pedidos de Compra",
    "consumidos": "Itens consumidos",
    "misto_por_ciclo": "Misto por ciclo",
    "indeterminado": "Indeterminado",
}


def _primeiro_informado(*valores: Any) -> Any:
    """Retorna o primeiro valor diferente de None sem confundir zero com ausência."""
    return next((valor for valor in valores if valor is not None), None)


def _numero(valor: Any) -> Real | None:
    if isinstance(valor, Real) and not isinstance(valor, bool):
        return valor
    return None


def _normalizar_metodo(valor: Any) -> str | None:
    texto = str(valor or "").strip().lower()
    if not texto:
        return None
    if texto in {"pc", "pcs", "pedido de compra", "pedidos de compra"}:
        return "pc"
    if texto in {"principal", "financeiro", "base financeira"}:
        return "financeiro"
    if texto in {"d", "consumidos", "itens consumidos"}:
        return "consumidos"
    if texto == "misto_por_ciclo":
        return texto
    if texto in {"indeterminado", "none", "nao informado", "não informado"}:
        return "indeterminado"
    return texto


def _unicos(valores: list[Any]) -> list[str]:
    saida: list[str] = []
    for valor in valores:
        texto = str(valor or "").strip()
        if texto and texto not in saida:
            saida.append(texto)
    return saida


def _alertas_materiais(alertas: list[Any]) -> list[str]:
    marcadores = (
        "diverg", "inconsist", "indetermin", "indispon",
        "sem fator", "calculo manual", "cálculo manual",
    )
    return [
        str(alerta).strip()
        for alerta in alertas
        if str(alerta or "").strip()
        and any(marcador in str(alerta).lower() for marcador in marcadores)
    ]


def montar_resultado_consolidado(
    resultado: dict[str, Any] | None,
    diagnostico: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Seleciona e classifica sinais existentes, sem recalcular valores."""
    resultado = resultado or {}
    diagnostico = diagnostico or resultado.get("diagnostico_coleta") or {}
    politica = resultado.get("politica_entrega_segura") or {}
    reconciliacao = resultado.get("reconciliacao_xls_python") or {}
    composicao_origem = resultado.get("composicao_vta") or {}
    totais_pc = resultado.get("totais_canonicos_pc") or {}
    controle = resultado.get("controle") or {}
    memoria = resultado.get("memoria_por_ciclo") or {}
    referencias_informadas = isinstance(resultado.get("referencias_vta"), dict)
    referencias_vta = resultado.get("referencias_vta") or {}

    metodo_controle = _normalizar_metodo(controle.get("modo"))
    metodo_efetivo = _normalizar_metodo(
        ((memoria.get("vta") or {}).get("metodo"))
        or ((politica.get("retroativo") or {}).get("metodo"))
        or metodo_controle
    )
    metodo_codigo = metodo_efetivo or metodo_controle or "indeterminado"
    metodo_pc = metodo_controle == "pc" or metodo_codigo == "pc"
    metodo_consumidos = metodo_controle == "d" or metodo_codigo == "consumidos"

    status_resultados = (
        ((diagnostico.get("metadados") or {}).get("status_resultados") or {})
        .get("valores") or {}
    )
    bloco_corte = totais_pc.get("ate_o_corte") or {}
    if metodo_pc:
        retroativo_reconhecido = _primeiro_informado(
            bloco_corte.get("retroativo"),
            status_resultados.get("retroativo_oficial"),
            resultado.get("valor_represado_a_pagar"),
        )
    elif metodo_consumidos:
        # VTA-C2 (item 12): no metodo Consumido nao ha fonte independente de
        # pagamento — nem resultado["valor_represado_a_pagar"] (ja suprimido
        # na origem) nem status_resultados["retroativo_oficial"] (mesmo
        # numero de decomposicao do reajuste, cacheado do XLS) sustentam a
        # rotulacao de "retroativo reconhecido". Fica indisponivel, nunca
        # fabricado a partir dessa decomposicao.
        retroativo_reconhecido = None
    else:
        retroativo_reconhecido = _primeiro_informado(
            resultado.get("valor_represado_a_pagar"),
            status_resultados.get("retroativo_oficial"),
        )

    if metodo_pc:
        valor_em_analise = bloco_corte.get("valor_atualizado_em_analise")
        retroativo_potencial = bloco_corte.get("delta_potencial")
        posterior = totais_pc.get("posterior_ao_corte") or {}
        posterior_disponivel = bool(posterior)
        fora_do_corte = {
            "aplicavel": True,
            "quantidade": posterior.get("quantidade") if posterior_disponivel else None,
            "valor_informado": posterior.get("valor_pc") if posterior_disponivel else None,
            "data_corte": _primeiro_informado(
                totais_pc.get("data_corte"), controle.get("data_corte")
            ),
        }
    else:
        valor_em_analise = None
        retroativo_potencial = None
        fora_do_corte = {
            "aplicavel": False,
            "quantidade": None,
            "valor_informado": None,
            "data_corte": controle.get("data_corte"),
        }

    linhas_origem = composicao_origem.get("linhas") or []
    composicao_disponivel = bool(composicao_origem.get("disponivel")) and bool(linhas_origem)
    composicao_conciliada = (
        not bool(composicao_origem.get("bloqueia_formalizacao"))
        if composicao_disponivel else None
    )
    composicao_exibivel = bool(composicao_disponivel and composicao_conciliada)
    if composicao_exibivel:
        mensagem_composicao = ""
    elif composicao_disponivel:
        mensagem_composicao = "A composição detalhada não está conciliada."
    else:
        mensagem_composicao = "A composição detalhada do VTA não está disponível."
    composicao = {
        "disponivel": composicao_disponivel,
        "conciliada": composicao_conciliada,
        "exibivel": composicao_exibivel,
        "linhas": [dict(linha) for linha in linhas_origem if isinstance(linha, dict)]
        if composicao_exibivel else [],
        "mensagem": mensagem_composicao,
        "ha_aditivos_nao_computados": bool(
            composicao_origem.get("aditivos_nao_computados")
        ),
    }

    bloqueios = _unicos(list(resultado.get("bloqueios_formalizacao") or []))
    bloqueado = bool(resultado.get("formalizacao_bloqueada")) or bool(bloqueios)

    pendencias_politica = list(politica.get("pendencias") or [])
    campos_nao_confiaveis = list(
        resultado.get("campos_nao_confiaveis_documentos") or []
    )
    ressalvas: list[Any] = list(pendencias_politica)

    potencial_num = _numero(retroativo_potencial)
    if potencial_num is not None and potencial_num != 0:
        ressalvas.append(
            "Há valor potencial sujeito à aceitação pela área gestora e à condução "
            "do eventual pagamento por essa área."
        )

    qtd_fora = _numero(fora_do_corte.get("quantidade"))
    if fora_do_corte["aplicavel"] and qtd_fora is not None and qtd_fora > 0:
        ressalvas.append("Há pedido(s) de compra posterior(es) à data de corte.")

    if campos_nao_confiaveis:
        ressalvas.append("Há campos materiais que dependem de confirmação.")

    status_reconciliacao = str(reconciliacao.get("status_geral") or "").upper()
    if status_reconciliacao == "DIVERGENCIA_DENTRO_DA_TOLERANCIA":
        ressalvas.append("Há divergência dentro da tolerância vigente.")
    elif status_reconciliacao == "RESULTADO_XLS_INDISPONIVEL_POR_CACHE":
        ressalvas.append("Parte das conferências automáticas está indisponível.")
    elif status_reconciliacao == "DIVERGENCIA_RELEVANTE" and not bloqueado:
        # A política existente decide se a divergência bloqueia. Aqui ela é
        # somente ressalva quando nenhum bloqueio explícito foi produzido.
        ressalvas.append("Há divergência relevante pendente de confirmação.")

    if not composicao_disponivel and resultado.get("valor_atualizado_contrato") is not None:
        ressalvas.append("A composição detalhada do VTA não está disponível.")
    ressalvas.extend(_alertas_materiais(list(composicao_origem.get("alertas") or [])))
    ressalvas = _unicos(ressalvas)

    status_politica = str(politica.get("status") or "").strip().upper()
    status_base = str(diagnostico.get("status_base") or "").strip().upper()
    vta_atual = resultado.get("valor_atualizado_contrato")
    posicao_atual_valida = bool(referencias_vta.get("posicao_atual_disponivel"))
    vta_ultima_posicao = _numero(referencias_vta.get("forma2_ultima_abertura"))
    if not referencias_informadas or posicao_atual_valida:
        vta = vta_atual
        vta_origem = "posicao_atual" if vta is not None else "indisponivel"
    elif vta_ultima_posicao is not None:
        vta = vta_ultima_posicao
        vta_origem = "ultima_posicao_disponivel"
    else:
        vta = None
        vta_origem = "indisponivel"
    resultado_incompleto = bool(
        vta is None
        or retroativo_reconhecido is None
        or metodo_codigo == "indeterminado"
        or status_politica in _STATUS_POLITICA_INCOMPLETOS
        or status_base == "ANALISE_PARCIAL_INFORMACOES_INSUFICIENTES"
        or (
            metodo_pc
            and composicao_origem.get("disponivel") is False
        )
    )

    if bloqueado:
        status = STATUS_BLOQUEADO
        mensagem_status = bloqueios[0] if bloqueios else "Há bloqueio explícito à formalização."
    elif resultado_incompleto:
        status = STATUS_PENDENTE
        mensagem_status = "O resultado central depende de complementação ou confirmação."
    elif ressalvas:
        status = STATUS_RESSALVAS
        somente_potencial = len(ressalvas) == 1 and potencial_num not in (None, 0)
        mensagem_status = (
            ressalvas[0]
            if somente_potencial
            else "Resultado apurado com ressalvas que não impedem sua leitura."
        )
    else:
        status = STATUS_CONFIAVEL
        mensagem_status = "Resultado apurado sem ressalvas materiais identificadas."

    formalizacao = {
        "bloqueada": bloqueado,
        "status": (
            "BLOQUEADA"
            if bloqueado else "AGUARDA CONFIRMAÇÃO"
            if status == STATUS_PENDENTE else "SEM BLOQUEIO"
        ),
        "mensagem": (
            bloqueios[0]
            if bloqueios else "Sem bloqueio explícito à formalização."
        ),
    }

    return {
        "vta": vta,
        "vta_origem": vta_origem,
        "vta_usa_ultima_posicao": vta_origem == "ultima_posicao_disponivel",
        "retroativo_reconhecido": retroativo_reconhecido,
        "valor_atualizado_em_analise": valor_em_analise,
        "retroativo_potencial": retroativo_potencial,
        "medidas_pc_aplicaveis": metodo_pc,
        "fora_do_corte": fora_do_corte,
        "metodo": {
            "codigo": metodo_codigo,
            "rotulo": _ROTULOS_METODO.get(metodo_codigo, str(metodo_codigo)),
        },
        "ciclo_vigente": controle.get("ciclo_vigente"),
        "status_confiabilidade": status,
        "mensagem_status": mensagem_status,
        "formalizacao": formalizacao,
        "bloqueios": bloqueios,
        "ressalvas": ressalvas,
        "campos_nao_confiaveis": campos_nao_confiaveis,
        "composicao_vta": composicao,
    }
