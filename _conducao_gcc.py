"""Modo de Conducao da GCC para o Claus New.

Camada de produto, sem calculo oficial: diagnostica se a GCC pode seguir em
automatico, se precisa equalizar fatos ou se deve registrar excecao/manual
controlado sem fingir certeza.
"""
from __future__ import annotations

from typing import Any


MODO_AUTOMATICO = "automatico"
MODO_ASSISTIDO = "assistido_gcc"
MODO_EXCEPCIONAL = "excepcional_manual_controlado"

EVIDENCIA_SUFICIENTE = "suficiente"
EVIDENCIA_PARCIAL = "parcial"
EVIDENCIA_INSUFICIENTE = "insuficiente"
EVIDENCIA_EXCEPCIONAL = "excepcional"


def _texto(valor: Any) -> str:
    return str(valor or "").strip()


def _f(valor: Any) -> float | None:
    if valor in (None, ""):
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _lista_textos(valores: Any) -> list[str]:
    if not isinstance(valores, list):
        return []
    return [str(v).strip() for v in valores if str(v).strip()]


def _pendencias_do_dossie(assistente: dict[str, Any] | None) -> list[str]:
    dossie = (assistente or {}).get("dossie_decisao") or {}
    pendencias: list[str] = []
    pendencias.extend(_lista_textos(dossie.get("pendencias_impeditivas")))
    pendencias.extend(_lista_textos(dossie.get("pendencias_nao_impeditivas")))
    resultado = (assistente or {}).get("resultado_recomendado") or {}
    pendencias.extend(_lista_textos(resultado.get("pendencias")))
    vistos: set[str] = set()
    unicas: list[str] = []
    for pendencia in pendencias:
        if pendencia not in vistos:
            vistos.add(pendencia)
            unicas.append(pendencia)
    return unicas


def _tem_dado_xls_nao_estruturado(leitura: dict[str, Any]) -> bool:
    entrada = leitura.get("masterfile_fiscal_definitivo") or {}
    if entrada.get("tem_observacoes_ressalvas"):
        return True
    linhas = entrada.get("linhas_por_aba") or {}
    return bool(linhas.get("FISCAL_OBSERVACOES") or linhas.get("FISCAL_RESSALVAS"))


def _normalizar_equalizacoes(intervencao_gcc: dict[str, Any]) -> list[dict[str, Any]]:
    equalizacoes = intervencao_gcc.get("equalizacoes")
    if isinstance(equalizacoes, dict):
        equalizacoes = [equalizacoes]
    if not isinstance(equalizacoes, list):
        return []

    normalizadas: list[dict[str, Any]] = []
    for item in equalizacoes:
        if not isinstance(item, dict):
            continue
        fato = {
            "tipo_fato": _texto(item.get("tipo_fato") or item.get("tipo")),
            "descricao": _texto(item.get("descricao")),
            "valor": _f(item.get("valor")),
            "data_referencia": item.get("data_referencia") or item.get("data"),
            "fonte": _texto(item.get("fonte")),
            "ressalva": _texto(item.get("ressalva")),
            "origem": "Equalizacao GCC",
        }
        if any(fato.get(k) for k in ("tipo_fato", "descricao", "valor", "data_referencia", "fonte")):
            normalizadas.append(fato)
    return normalizadas


def _normalizar_manual(intervencao_gcc: dict[str, Any]) -> dict[str, Any]:
    bruto = intervencao_gcc.get("resultado_manual_controlado") or {}
    if not isinstance(bruto, dict):
        bruto = {}
    valor = _f(bruto.get("valor"))
    motivo = _texto(bruto.get("motivo"))
    ressalva = _texto(bruto.get("ressalva"))
    informado = any([valor is not None, motivo, ressalva, bruto.get("excepcional")])
    pendencias = []
    if informado:
        if not motivo:
            pendencias.append("Resultado manual controlado exige motivo.")
        if not ressalva:
            pendencias.append("Resultado manual controlado exige ressalva.")
    return {
        "informado": informado,
        "registrado": informado and not pendencias,
        "valor": valor,
        "motivo": motivo,
        "ressalva": ressalva,
        "origem": "Manual controlado GCC" if informado else "",
        "pendencias": pendencias,
        "nao_altera_vta_oficial": True,
        "nao_substitui_documento_final": True,
    }


def _classificar(
    *,
    resultado: dict[str, Any] | None,
    evidencia_suficiente: bool,
    pendencias: list[str],
    dupla_contagem: bool,
    equalizacoes: list[dict[str, Any]],
    manual: dict[str, Any],
    excepcional_solicitado: bool,
    entrada_xls_nao_estruturada: bool,
) -> tuple[str, str]:
    if excepcional_solicitado or manual.get("informado"):
        return MODO_EXCEPCIONAL, EVIDENCIA_EXCEPCIONAL
    if not resultado or not evidencia_suficiente:
        return MODO_EXCEPCIONAL, EVIDENCIA_INSUFICIENTE
    if pendencias or dupla_contagem or equalizacoes or entrada_xls_nao_estruturada:
        return MODO_ASSISTIDO, EVIDENCIA_PARCIAL
    return MODO_AUTOMATICO, EVIDENCIA_SUFICIENTE


def diagnosticar_modo_conducao_gcc(
    leitura: dict[str, Any],
    painel: dict[str, Any],
    assistente: dict[str, Any] | None = None,
    intervencao_gcc: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Diagnostica o trilho operacional sem alterar leitura, painel ou XLS."""
    intervencao = dict(intervencao_gcc or {})
    motor = painel.get("motor_metodologias") or {}
    resultado = (assistente or {}).get("resultado_recomendado") or motor.get("resultado_recomendado")
    evidencia = (assistente or {}).get("evidencia") or {}
    evidencia_suficiente = bool(evidencia.get("suficiente") or resultado)
    dupla = motor.get("dupla_contagem") or {}
    pendencias = _pendencias_do_dossie(assistente)
    equalizacoes = _normalizar_equalizacoes(intervencao)
    manual = _normalizar_manual(intervencao)
    if manual.get("registrado"):
        pendencias = [
            p for p in pendencias
            if "manual valido" not in p.lower()
        ]
    entrada_xls_nao_estruturada = _tem_dado_xls_nao_estruturado(leitura)
    excepcional_solicitado = bool(intervencao.get("excepcional") or intervencao.get("fora_do_modelo"))

    modo, grau = _classificar(
        resultado=resultado,
        evidencia_suficiente=evidencia_suficiente,
        pendencias=pendencias,
        dupla_contagem=bool(dupla.get("existe")),
        equalizacoes=equalizacoes,
        manual=manual,
        excepcional_solicitado=excepcional_solicitado,
        entrada_xls_nao_estruturada=entrada_xls_nao_estruturada,
    )

    if modo == MODO_AUTOMATICO:
        rotulo = "Automatico"
        acao = "Validar recomendacao e seguir para formalizacao quando cabivel."
        calculo = "Claus New calcula e recomenda com evidencia suficiente."
    elif modo == MODO_ASSISTIDO:
        rotulo = "Assistido GCC"
        acao = "Equalizar dados recebidos e validar calculo com ressalva."
        calculo = "Claus New calcula com ressalva e registra origem Equalizacao GCC quando houver ajuste."
    else:
        rotulo = "Excepcional / Manual Controlado"
        acao = "Registrar limitacao, motivo e ressalva; nao inventar resultado."
        calculo = "Claus New nao substitui o caso excepcional por calculo automatico."

    alertas = list(manual.get("pendencias") or [])
    if modo == MODO_EXCEPCIONAL and not manual.get("registrado"):
        alertas.append("Caso excepcional sem resultado manual valido; manter limitacao registrada.")
    if modo == MODO_ASSISTIDO and entrada_xls_nao_estruturada:
        alertas.append("Ha observacoes ou ressalvas na entrada XLS que exigem equalizacao da GCC.")

    return {
        "modo": modo,
        "rotulo": rotulo,
        "grau_evidencia": grau,
        "diagnostico": calculo,
        "proxima_acao": acao,
        "automatizacao": {
            "automatico": modo == MODO_AUTOMATICO,
            "assistido": modo == MODO_ASSISTIDO,
            "manual_controlado": modo == MODO_EXCEPCIONAL,
        },
        "equalizacao_gcc": {
            "origem": "Equalizacao GCC",
            "campos": [
                "tipo_fato",
                "descricao",
                "valor",
                "data_referencia",
                "fonte",
                "ressalva",
            ],
            "fatos_equalizados": equalizacoes,
            "calcular_com_ressalva": modo == MODO_ASSISTIDO,
            "pendencias_origem": pendencias,
        },
        "resultado_manual_controlado": manual,
        "alertas": alertas,
        "rastreabilidade": {
            "web_exclusiva_gcc": True,
            "origem_equalizacao": "Equalizacao GCC" if equalizacoes else "",
            "pc_classificacao": "execucao",
            "data_pc": "linha temporal dos ciclos",
            "nao_inventa_vta_retroativo_pagamento": True,
        },
    }
