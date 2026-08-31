"""Camada pura de consolidação dos resultados já calculados.

Este módulo não calcula grandezas financeiras. Ele apenas seleciona fontes
canônicas existentes, preserva a diferença entre ausência e zero e prepara um
contrato estável para a interface de Upload e resultados.
"""

from __future__ import annotations

from numbers import Real
from typing import Any


# STATUS-CANON-1: o status apresentado nao e recalculado aqui. O cache de
# RESULTADOS!B3 tem precedencia; quando apenas esse cache falta, a conclusao ja
# emitida pela politica canonica Python pode ser usada. O vocabulario visivel
# permanece o vocabulario oficial do produto.
STATUS_CONFIAVEL = "VALIDADO"
STATUS_RESSALVAS = "VALIDADO COM RESSALVAS"
STATUS_ESTIMADO = "ESTIMADO"
STATUS_PENDENTE = "PENDENTE DE CONFIRMAÇÃO"
# Compatibilidade: BLOQUEADO deixou de ser status DA APURACAO. Bloqueio e um
# estado da FORMALIZACAO (formalizacao["bloqueada"]/["status"]), separado da
# conclusao do calculo. A constante permanece exportada porque integra o
# contrato publico ja consumido por testes e chamadores.
STATUS_BLOQUEADO = "BLOQUEADO"

# VTA-U2.2: origens possiveis do VTA entregue ao web e aos documentos. Existe um
# unico caminho valido — o calculo canonico da metodologia selecionada
# (== MEMORIA_RESULTADOS!B26 == memoria_por_ciclo.vta.valor_total_atualizado).
# Referencias fisicas (RESULTADOS!B10/B11) nunca sao origem de VTA.
ORIGEM_VTA_CANONICA = "vta_canonico"
ORIGEM_VTA_INDISPONIVEL = "indisponivel"

# STATUS-CANON-1: vocabulario emitido por RESULTADOS!B3 e sua traducao para o
# painel. "VALIDADO COM RESSALVAS" nao e emitido pelo template atual; fica
# mapeado porque a regra canonica e preservar a nomenclatura oficial do XLS
# caso ela passe a existir, nunca inventar um estado novo no Python.
_STATUS_OFICIAL_PARA_PAINEL = {
    "VALIDADO": STATUS_CONFIAVEL,
    "VALIDADO COM RESSALVAS": STATUS_RESSALVAS,
    "ESTIMADO": STATUS_ESTIMADO,
    "REVISE": STATUS_PENDENTE,
}
# Conclusoes de RESULTADOS que encerram a apuracao. "REVISE" e o proprio XLS
# pedindo revisao — nao e conclusao.
_STATUS_OFICIAL_CONCLUSIVOS = {"VALIDADO", "VALIDADO COM RESSALVAS", "ESTIMADO"}

ORIGEM_STATUS_RESULTADOS = "resultados_xls"
ORIGEM_STATUS_MOTOR_PYTHON = "motor_canonico_python"
ORIGEM_STATUS_INDISPONIVEL = "indisponivel"

_STATUS_POLITICA_CONCLUSIVOS = {
    "PRONTO_PARA_VALIDACAO_FISCAL",
    "APTO_PARA_FORMALIZACAO",
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


# VTA-U2 (fase 7): a camada visivel ao usuario nao fala em "fator acumulado
# parametrizado" nem em nomes de registro interno. O diagnostico tecnico segue
# intacto em composicao_vta["alertas"]; aqui so a traducao apresentada.
def _traduzir_alerta_composicao(alerta: str) -> str:
    texto = alerta.lower()
    if "sem fator acumulado" in texto:
        return (
            "Os valores já pagos foram considerados conforme informados. Como "
            "esse total não pertence a um único ciclo de reajuste, nenhum novo "
            "fator foi aplicado sobre ele. Isso não impede o cálculo do VTA."
        )
    if "agregador" in texto:
        return (
            "A linha de total das abas de entrada não foi somada junto das "
            "parcelas de cada ciclo, para não contar a mesma execução duas vezes."
        )
    return alerta


def _alertas_materiais(alertas: list[Any]) -> list[str]:
    marcadores = (
        "diverg", "inconsist", "indetermin", "indispon",
        "sem fator", "calculo manual", "cálculo manual",
    )
    return [
        _traduzir_alerta_composicao(str(alerta).strip())
        for alerta in alertas
        if str(alerta or "").strip()
        and any(marcador in str(alerta).lower() for marcador in marcadores)
    ]


# STATUS-CANON-1: leitura semantica — nao reproduz a formula de RESULTADOS!B3
# nem recalcula seus eixos. Le a conclusao que o XLS ja gravou e a classifica.
def _status_oficial_resultados(status_resultados: dict[str, Any]) -> dict[str, Any]:
    """Classifica a conclusao oficial da aba RESULTADOS, sem fabricar VALIDADO.

    Qualquer coisa fora do vocabulario oficial — ausencia, celula vazia, cache
    nao calculado ou texto desconhecido — e indisponibilidade, nunca aprovacao.
    """
    bruto = (status_resultados or {}).get("geral")
    texto = str(bruto).strip().upper() if bruto is not None else ""
    if texto in _STATUS_OFICIAL_PARA_PAINEL:
        return {
            "codigo": texto,
            "bruto": bruto,
            "disponivel": True,
            "conclusivo": texto in _STATUS_OFICIAL_CONCLUSIVOS,
            "origem": ORIGEM_STATUS_RESULTADOS,
        }
    return {
        "codigo": None,
        "bruto": bruto,
        "disponivel": False,
        "conclusivo": False,
        "origem": ORIGEM_STATUS_INDISPONIVEL,
    }


def _status_canonico_apuracao(
    status_resultados: dict[str, Any],
    politica: dict[str, Any],
    formula_status_presente: bool = False,
) -> dict[str, Any]:
    """Seleciona uma conclusao existente sem recalcular a formula do XLS.

    O cache valido de RESULTADOS continua tendo precedencia, inclusive para
    REVISE e ESTIMADO. O fallback so existe para a situacao precisamente
    distinguivel de formula sem cache: a formula foi comprovada no workbook,
    mas seu valor esta vazio, enquanto a politica canonica Python ja concluiu
    a apuracao e declarou que ela pode ser confirmada. Formula/bloco ausente,
    vocabulario desconhecido e politica parcial/bloqueada permanecem
    fail-closed.
    """
    status_xls = _status_oficial_resultados(status_resultados)
    if status_xls["disponivel"]:
        return status_xls

    bruto = (status_resultados or {}).get("geral")
    formula_presente = formula_status_presente is True
    cache_vazio = bruto is None or not str(bruto).strip()
    status_politica = str((politica or {}).get("status") or "").strip().upper()
    politica_conclusiva = (
        status_politica in _STATUS_POLITICA_CONCLUSIVOS
        and (politica or {}).get("pode_confirmar") is True
        and not list((politica or {}).get("bloqueios") or [])
    )
    if formula_presente and cache_vazio and politica_conclusiva:
        return {
            "codigo": "VALIDADO",
            "bruto": bruto,
            "disponivel": True,
            "conclusivo": True,
            "origem": ORIGEM_STATUS_MOTOR_PYTHON,
        }
    return status_xls


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

    # STATUS-CANON-1: `status_resultados_xls` e o bloco inteiro lido da aba
    # RESULTADOS (inclui a conclusao oficial em "geral" == RESULTADOS!B3);
    # `status_resultados` continua sendo apenas o sub-dicionario de valores.
    status_resultados_xls = (
        (diagnostico.get("metadados") or {}).get("status_resultados") or {}
    )
    status_resultados = status_resultados_xls.get("valores") or {}
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

    # STATUS-CANON-1: a politica de entrega segura continua detectando riscos
    # reais, mas o status interno dela deixou de ser uma segunda regua que
    # substitui a conclusao oficial. O que ela apurou entra como RESSALVA.
    status_politica = str(politica.get("status") or "").strip().upper()
    status_base = str(diagnostico.get("status_base") or "").strip().upper()
    if status_base == "ANALISE_PARCIAL_INFORMACOES_INSUFICIENTES":
        ressalvas.append(
            "A leitura da Coleta registrou informações insuficientes em parte "
            "dos blocos."
        )
    if metodo_consumidos and retroativo_reconhecido is None:
        # VTA-C2: supressao deliberada, nao lacuna de informacao. Fica visivel
        # como ressalva em vez de derrubar a apuracao inteira.
        ressalvas.append(
            "No método Itens consumidos não há fonte independente de pagamento "
            "para rotular retroativo reconhecido."
        )
    ressalvas = _unicos(ressalvas)

    vta_atual = resultado.get("valor_atualizado_contrato")
    # Referencias fisicas do XLS, preservadas e rotuladas — nunca viram VTA.
    referencias_auditaveis = {
        "posicao_fisica_atual": _numero(referencias_vta.get("forma1_posicao_atual")),
        "ultima_abertura_disponivel": _numero(
            referencias_vta.get("forma2_ultima_abertura")
        ),
        "posicao_atual_disponivel": bool(
            referencias_vta.get("posicao_atual_disponivel")
        ),
        "rotulos": {
            "posicao_fisica_atual": (
                "Referência auditável — posição física atual"
            ),
            "ultima_abertura_disponivel": (
                "Referência auditável — última posição de abertura disponível"
            ),
        },
    }
    # VTA-U2 (achado A): o VTA entregue ao web e aos documentos e SEMPRE o VTA
    # canonico do metodo selecionado (== MEMORIA_RESULTADOS!B26 ==
    # memoria_por_ciclo.vta.valor_total_atualizado). As referencias fisicas de
    # RESULTADOS!B10/B11 (posicao atual / ultima abertura) permanecem expostas
    # em `referencias_vta` como REFERENCIA AUDITAVEL, mas nunca substituem,
    # completam nem servem de fallback do VTA: sem VTA canonico o resultado fica
    # indisponivel (fail-closed) e a politica existente degrada ou bloqueia.
    # VTA-U2.2: a origem passa a dizer o que o valor de fato e. Enquanto existia
    # o fallback pela posicao fisica, "posicao_atual" distinguia os dois
    # caminhos; sem ele, so ha um caminho — o calculo canonico da metodologia.
    vta = vta_atual
    vta_origem = ORIGEM_VTA_CANONICA if vta is not None else ORIGEM_VTA_INDISPONIVEL

    # ------------------------------------------------------------------
    # STATUS-CANON-1 — STATUS DA APURACAO
    # ------------------------------------------------------------------
    # Fonte primaria: conclusao da aba RESULTADOS. Sem o cache dessa formula, a
    # politica canonica Python pode fornecer a conclusao que ela ja produziu;
    # este modulo nao reproduz a formula nem cria uma regua paralela.
    #
    # Fail-closed: quando a conclusao oficial nao existe, o painel NAO fabrica
    # VALIDADO. Mas a indisponibilidade tem de vir da ausencia do STATUS OFICIAL
    # (ou da ausencia do proprio VTA/metodo), jamais da ausencia isolada de
    # PC/Financeiro/Consumo em um ciclo — execucao legitimamente zero e um valor
    # apurado, nao uma lacuna (ZERO REAL != AUSENCIA).
    status_oficial = _status_canonico_apuracao(
        status_resultados_xls,
        politica,
        (diagnostico.get("metadados") or {}).get(
            "formula_status_resultados_presente"
        ),
    )
    # Lacunas materiais do proprio nucleo, independentes da aba RESULTADOS.
    # Nao incluem "ciclo sem PC": esse sinal e ressalva, nunca indisponibilidade.
    nucleo_indisponivel = vta is None or metodo_codigo == "indeterminado"

    if not status_oficial["disponivel"]:
        status = STATUS_PENDENTE
        mensagem_status = (
            "O status oficial da apuração (aba RESULTADOS) não está disponível "
            "neste arquivo. Abra o XLS no Excel, recalcule, salve e reenvie."
        )
    elif nucleo_indisponivel:
        status = STATUS_PENDENTE
        mensagem_status = (
            "O VTA ou o método da apuração não estão disponíveis nesta leitura."
            if vta is None
            else "O método da apuração não pôde ser determinado."
        )
    else:
        status = _STATUS_OFICIAL_PARA_PAINEL[status_oficial["codigo"]]
        if status == STATUS_PENDENTE:
            mensagem_status = (
                "A aba RESULTADOS concluiu REVISE: o próprio cálculo do XLS "
                "aponta eixos a revisar."
            )
        elif status == STATUS_ESTIMADO:
            mensagem_status = (
                "A aba RESULTADOS concluiu ESTIMADO: o remanescente foi tratado "
                "por estimativa."
            )
        elif status_oficial["origem"] == ORIGEM_STATUS_MOTOR_PYTHON:
            mensagem_status = (
                "Conclusão produzida pelo motor canônico da apuração; a fórmula "
                "de status do XLS está presente, mas sem valor cacheado."
            )
        else:
            mensagem_status = "Conclusão reproduzida da aba RESULTADOS do XLS."

    status_conclusivo = status not in (STATUS_PENDENTE,)

    # ------------------------------------------------------------------
    # STATUS-CANON-1.1 — INFORMACOES (execucao zero conhecida)
    # ------------------------------------------------------------------
    # A politica ja separou ZERO CONHECIDO de EXECUCAO DESCONHECIDA olhando a
    # fonte de execucao do metodo. Aqui entra a ultima condicao do contrato: o
    # resultado oficial precisa estar disponivel. Sem conclusao oficial o fato
    # volta a ser ressalva — nunca o contrario.
    informacoes = _unicos(list(politica.get("informacoes") or []))
    if informacoes and not status_conclusivo:
        ressalvas = _unicos(ressalvas + informacoes)
        informacoes = []

    # ------------------------------------------------------------------
    # STATUS-CANON-1 — FORMALIZACAO (eixo separado)
    # ------------------------------------------------------------------
    # A formalizacao pode estar bloqueada com a apuracao validada, desde que
    # exista causa objetiva e explicita. O inverso — derrubar a apuracao sem
    # causa material — e exatamente o que esta etapa eliminou.
    formalizacao = {
        "bloqueada": bloqueado,
        "status": (
            "BLOQUEADA"
            if bloqueado else "SEM BLOQUEIO"
            if status_conclusivo else "AGUARDA CONFIRMAÇÃO"
        ),
        "mensagem": (
            bloqueios[0]
            if bloqueios else "Há bloqueio explícito à formalização."
            if bloqueado else "Sem bloqueio explícito à formalização."
            if status_conclusivo else mensagem_status
        ),
    }

    return {
        "vta": vta,
        "vta_origem": vta_origem,
        # Compatibilidade: o fallback pela ultima posicao de abertura deixou de
        # existir (VTA-U2, achado A), entao esta chave e sempre False. Mantida
        # porque integra o contrato publico ja consumido por testes.
        "vta_usa_ultima_posicao": False,
        "retroativo_reconhecido": retroativo_reconhecido,
        "valor_atualizado_em_analise": valor_em_analise,
        "retroativo_potencial": retroativo_potencial,
        "medidas_pc_aplicaveis": metodo_pc,
        "referencias_auditaveis": referencias_auditaveis,
        "fora_do_corte": fora_do_corte,
        "metodo": {
            "codigo": metodo_codigo,
            "rotulo": _ROTULOS_METODO.get(metodo_codigo, str(metodo_codigo)),
        },
        "ciclo_vigente": controle.get("ciclo_vigente"),
        # STATUS-CANON-1: conceito explicito e separado. `status_apuracao` diz
        # de onde veio a conclusao; `status_confiabilidade` guarda o mesmo texto
        # e permanece por compatibilidade com o contrato ja consumido.
        "status_apuracao": {
            "codigo": status_oficial["codigo"],
            "rotulo": status,
            "origem": status_oficial["origem"],
            "disponivel": status_oficial["disponivel"],
            "conclusivo": status_conclusivo,
            "bruto": status_oficial["bruto"],
            "mensagem": mensagem_status,
            # Diagnostico da politica de seguranca, preservado e visivel, mas
            # sem poder de substituir a conclusao oficial da apuracao.
            "status_politica": status_politica or None,
        },
        "status_confiabilidade": status,
        "mensagem_status": mensagem_status,
        "formalizacao": formalizacao,
        "bloqueios": bloqueios,
        "ressalvas": ressalvas,
        # Fatos apurados, sem linguagem de pendencia. Nao alteram o status da
        # apuracao nem a formalizacao.
        "informacoes": informacoes,
        "campos_nao_confiaveis": campos_nao_confiaveis,
        "composicao_vta": composicao,
    }
