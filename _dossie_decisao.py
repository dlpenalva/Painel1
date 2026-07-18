"""Dossie de Decisao do Claus New.

Camada de conclusao operacional, somente leitura. Nao calcula VTA, nao altera
MasterFile e nao cria documento; transforma o painel e o assistente em uma
decisao recomendada para confirmacao operacional.
"""
from __future__ import annotations

import copy
from typing import Any


STATUS_APTO = "apto_para_confirmacao"
STATUS_AGUARDAR = "aguardar_evidencia_ou_ato"
STATUS_BLOQUEADO = "bloqueado"
STATUS_INSUFICIENTE = "evidencia_insuficiente"

ATO_APOSTILAMENTO = "Preparar apostilamento"
ATO_RECONHECIMENTO = "Reconhecer retroativo antes de formalizar"
ATO_AGUARDAR_PAGAMENTO = "Aguardar pagamento ou comprovacao de execucao"
ATO_CORRIGIR_EVIDENCIA = "Corrigir ou complementar evidencias no MasterFile"
ATO_REVISAR_DUPLA = "Revisar possivel dupla contagem"
ATO_SEM_ATO = "Nenhum ato imediato"


def _f(valor: Any) -> float:
    try:
        return float(valor or 0.0)
    except (TypeError, ValueError):
        return 0.0


def _brl(valor: Any) -> str:
    try:
        return "R$ " + f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "valor nao disponivel"


def _inc_por_gravidade(assistente: dict[str, Any], gravidade: str) -> list[dict[str, Any]]:
    return [
        i for i in assistente.get("inconsistencias") or []
        if str(i.get("gravidade") or "").lower() == gravidade
    ]


def _pendencias_do_motor(assistente: dict[str, Any]) -> list[str]:
    resultado = assistente.get("resultado_recomendado") or {}
    return [str(p) for p in resultado.get("pendencias") or [] if str(p).strip()]


def _tem_dupla_contagem(assistente: dict[str, Any]) -> bool:
    dupla = (assistente.get("motor_metodologias") or {}).get("dupla_contagem") or {}
    return bool(dupla.get("existe"))


def _pcs_geradores(assistente: dict[str, Any]) -> list[dict[str, Any]]:
    resultado = assistente.get("resultado_recomendado") or {}
    pcs = []
    for pc in resultado.get("pc_gerador_retroativo") or []:
        pcs.append({
            "numero_pc": pc.get("numero_pc"),
            "ciclo": pc.get("ciclo"),
            "valor": pc.get("valor"),
            "classificacao": "execucao",
        })
    return pcs


def _evidencias(assistente: dict[str, Any]) -> list[str]:
    evidencias: list[str] = []
    metodologia = assistente.get("metodologia") or {}
    resultado = assistente.get("resultado_recomendado") or {}
    retro = assistente.get("retroativo") or {}
    vta = assistente.get("vta") or {}
    valor_contrato = assistente.get("valor_contrato") or {}
    saldo = assistente.get("saldo") or {}
    conducao = assistente.get("modo_conducao_gcc") or {}

    if metodologia.get("escolhida"):
        evidencias.append(
            f"Metodologia recomendada: {metodologia.get('escolhida')} "
            f"(confiabilidade {metodologia.get('confiabilidade')})."
        )
    if resultado:
        evidencias.append(
            "Resultado recomendado: "
            f"{_brl(resultado.get('valor_recomendado'))}; "
            f"em analise {_brl(resultado.get('valor_em_analise'))}; "
            f"reconhecido {_brl(resultado.get('valor_reconhecido'))}; "
            f"pago {_brl(resultado.get('valor_pago'))}."
        )
    if retro.get("existe"):
        evidencias.append(
            f"Retroativo apurado: {_brl(retro.get('valor'))} "
            f"em estado {retro.get('estado')}."
        )
    if vta.get("valor") is not None:
        evidencias.append(
            "Valor historico do arquivo preservado para conferencia "
            f"(nao e VTA): {_brl(vta.get('valor'))}."
        )
    if valor_contrato.get("valor") is not None:
        evidencias.append(
            "Valor do contrato hoje localizado: "
            f"{_brl(valor_contrato.get('valor'))}."
        )
    if saldo.get("valor") is not None:
        evidencias.append(f"Saldo remanescente localizado: {_brl(saldo.get('valor'))}.")
    if conducao.get("rotulo"):
        evidencias.append(
            f"Modo de conducao GCC: {conducao.get('rotulo')} "
            f"({conducao.get('grau_evidencia')})."
        )
    manual = conducao.get("resultado_manual_controlado") or {}
    if manual.get("registrado"):
        evidencias.append(
            "Resultado manual controlado registrado pela GCC: "
            f"{_brl(manual.get('valor'))}; motivo: {manual.get('motivo')}; "
            f"ressalva: {manual.get('ressalva')}."
        )
    return evidencias


def _perguntas_resolvidas(assistente: dict[str, Any], status: str) -> list[dict[str, str]]:
    resultado = assistente.get("resultado_recomendado") or {}
    retro = assistente.get("retroativo") or {}
    metodologia = assistente.get("metodologia") or {}
    dupla = (assistente.get("motor_metodologias") or {}).get("dupla_contagem") or {}
    conducao = assistente.get("modo_conducao_gcc") or {}

    return [
        {
            "pergunta": "Qual o modo de conducao da GCC?",
            "resposta": conducao.get("rotulo") or "Nao diagnosticado.",
        },
        {
            "pergunta": "Qual metodologia deve ser usada?",
            "resposta": metodologia.get("escolhida") or "Evidencia insuficiente para recomendar.",
        },
        {
            "pergunta": "Existe retroativo e qual PC gerou?",
            "resposta": (
                ", ".join(
                    f"{pc.get('numero_pc')} ({_brl(pc.get('valor'))}, ciclo {pc.get('ciclo')})"
                    for pc in _pcs_geradores(assistente)
                )
                if _pcs_geradores(assistente)
                else ("Nao identificado." if not retro.get("existe")
                      else f"Retroativo de {_brl(retro.get('valor'))} sem PC individualizado.")
            ),
        },
        {
            "pergunta": "Quanto esta em analise, reconhecido e pago?",
            "resposta": (
                f"Em analise {_brl(resultado.get('valor_em_analise'))}; "
                f"reconhecido {_brl(resultado.get('valor_reconhecido'))}; "
                f"pago {_brl(resultado.get('valor_pago'))}."
                if resultado else "Nao ha resultado recomendado."
            ),
        },
        {
            "pergunta": "Ha risco de dupla contagem?",
            "resposta": (
                f"Sim, {dupla.get('quantidade', 0)} item(ns) exigem revisao."
                if dupla.get("existe") else "Nao identificado pelo Event Log/VTA sombra."
            ),
        },
        {
            "pergunta": "O processo pode seguir dentro do Claus New?",
            "resposta": "Sim." if status == STATUS_APTO else "Ainda nao.",
        },
    ]


def _decidir_status(assistente: dict[str, Any]) -> tuple[str, str, str]:
    bloqueantes = _inc_por_gravidade(assistente, "bloqueante")
    atencoes = _inc_por_gravidade(assistente, "atencao")
    evidencia = assistente.get("evidencia") or {}
    resultado = assistente.get("resultado_recomendado")
    retro = assistente.get("retroativo") or {}

    if bloqueantes:
        return (
            STATUS_BLOQUEADO,
            ATO_CORRIGIR_EVIDENCIA,
            "Ha inconsistencia bloqueante; concluir agora aumentaria risco de erro.",
        )
    if not evidencia.get("suficiente") or not resultado:
        return (
            STATUS_INSUFICIENTE,
            ATO_CORRIGIR_EVIDENCIA,
            "Faltam evidencias suficientes para recomendar uma conclusao segura.",
        )
    if _tem_dupla_contagem(assistente):
        return (
            STATUS_BLOQUEADO,
            ATO_REVISAR_DUPLA,
            "Ha indicio de valor ja refletido em outra fonte.",
        )
    if atencoes:
        return (
            STATUS_AGUARDAR,
            ATO_CORRIGIR_EVIDENCIA,
            "Ha pendencias de consistencia que devem ser resolvidas antes da conclusao.",
        )
    por_estado = retro.get("por_estado") or {}
    if _f(por_estado.get("em analise") or por_estado.get("em análise")) > 0.0:
        return (
            STATUS_AGUARDAR,
            ATO_AGUARDAR_PAGAMENTO,
            "Ha retroativo em analise; reconhecimento depende de evidencia ou ato administrativo.",
        )
    if _f(por_estado.get("reconhecido")) > 0.0:
        return (
            STATUS_APTO,
            ATO_APOSTILAMENTO,
            "Ha retroativo reconhecido e evidencias suficientes para preparar a formalizacao.",
        )
    if retro.get("existe"):
        return (
            STATUS_AGUARDAR,
            ATO_RECONHECIMENTO,
            "Ha retroativo apurado, mas o estado ainda nao permite formalizacao direta.",
        )
    return (
        STATUS_APTO,
        ATO_SEM_ATO,
        "Nao ha retroativo pendente nem inconsistencia impeditiva.",
    )


def montar_dossie_decisao(assistente: dict[str, Any], painel: dict[str, Any]) -> dict[str, Any]:
    """Monta uma decisao operacional completa sem alterar entradas."""
    if assistente.get("objeto_processo_id") and isinstance(assistente.get("dossie_decisao"), dict):
        return copy.deepcopy(assistente["dossie_decisao"])

    if not assistente.get("disponivel"):
        return {
            "disponivel": False,
            "motivo": assistente.get("motivo") or "Assistente indisponivel.",
        }

    status, ato, motivo = _decidir_status(assistente)
    bloqueantes = _inc_por_gravidade(assistente, "bloqueante")
    atencoes = _inc_por_gravidade(assistente, "atencao")
    pend_motor = _pendencias_do_motor(assistente)
    conducao = assistente.get("modo_conducao_gcc") or {}
    pendencias_impeditivas = [
        str(i.get("descricao") or "") for i in bloqueantes + atencoes
        if str(i.get("descricao") or "").strip()
    ]
    for p in pend_motor:
        if p not in pendencias_impeditivas:
            pendencias_impeditivas.append(p)
    if _tem_dupla_contagem(assistente):
        pendencias_impeditivas.append("Revisar possivel dupla contagem antes de concluir.")
    for alerta in conducao.get("alertas") or []:
        if alerta not in pendencias_impeditivas:
            pendencias_impeditivas.append(str(alerta))

    resultado = assistente.get("resultado_recomendado") or {}
    mfi = assistente.get("masterfile_inteligente") or painel.get("masterfile_inteligente") or {}
    campos_nao_precisa = (
        list(mfi.get("campos_deixaram_de_ser_preenchidos") or [])
        if mfi.get("layout_inteligente") else [
            "metodologia sugerida",
            "ciclo de cada PC pela DATA_PC",
            "PC gerador do retroativo",
            "valor em analise, reconhecido e pago",
            "VTA e saldo remanescente quando disponiveis no MasterFile",
            "risco de dupla contagem sinalizado pelo Event Log/VTA sombra",
        ]
    )
    resumo = (
        f"Conclusao recomendada: {ato}. {motivo} "
        f"Metodologia: {(assistente.get('metodologia') or {}).get('escolhida') or 'nao definida'}. "
        f"Resultado: {_brl(resultado.get('valor_recomendado')) if resultado else 'sem resultado recomendado'}."
    )

    precisa_abrir_excel = (
        status == STATUS_INSUFICIENTE
        or ato == ATO_CORRIGIR_EVIDENCIA
    )

    return {
        "disponivel": True,
        "status": status,
        "pode_concluir_no_claus_new": status == STATUS_APTO,
        "precisa_abrir_excel": precisa_abrir_excel,
        "ato_ou_providencia": ato,
        "motivo": motivo,
        "resumo_executivo": resumo,
        "evidencias_utilizadas": _evidencias(assistente),
        "pendencias_impeditivas": pendencias_impeditivas,
        "pendencias_nao_impeditivas": [
            str(i.get("descricao") or "") for i in assistente.get("inconsistencias") or []
            if str(i.get("gravidade") or "").lower() == "informacao"
        ],
        "perguntas_resolvidas": _perguntas_resolvidas(assistente, status),
        "modo_conducao_gcc": conducao,
        "campos_que_o_usuario_nao_precisa_preencher": campos_nao_precisa,
        "masterfile_inteligente": mfi,
        "confirmacao_requerida": status == STATUS_APTO,
        "garantias": {
            "pc_classificacao": "execucao",
            "data_pc_enquadramento": "linha temporal dos ciclos",
            "nao_altera_xls": True,
            "nao_inventa_dados": True,
            "vta_oficial_preservado": (painel.get("vta") or {}).get("oficial"),
        },
    }
