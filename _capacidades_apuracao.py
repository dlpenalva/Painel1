"""Motor unico de capacidades para o processamento progressivo do XLS.

O modulo nao recalcula o workbook. Ele decide, a partir das evidencias
preenchidas e dos valores em cache gravados pelo Excel, quais blocos e
documentos podem ser usados com seguranca.
"""

from __future__ import annotations

from typing import Any, Iterable


ESTADO_COMPLETO = "completo"
ESTADO_PARCIAL = "parcial"
ESTADO_NAO_INFORMADO = "nao_informado"
ESTADO_BLOQUEADO = "bloqueado"


def _numero_disponivel(valor: Any) -> bool:
    return valor not in (None, "") and not isinstance(valor, bool)


def _texto(valor: Any) -> str:
    return str(valor or "").strip().upper()


def _registro(
    nome: str,
    estado: str,
    rotulo: str,
    detalhe: str,
    *,
    disponivel: bool = False,
    valor: Any = None,
    origem: str = "",
) -> dict[str, Any]:
    return {
        "nome": nome,
        "estado": estado,
        "rotulo": rotulo,
        "detalhe": detalhe,
        "disponivel": bool(disponivel),
        "valor": valor,
        "origem": origem,
    }


def _documento(
    nome: str,
    estado: str,
    rotulo: str,
    motivo: str,
    *,
    habilitado: bool,
) -> dict[str, Any]:
    return {
        "nome": nome,
        "estado": estado,
        "rotulo": rotulo,
        "classificacao": rotulo.upper(),
        "motivo": motivo,
        "habilitado": bool(habilitado),
    }


def _primeiro_valor(candidatos: Iterable[tuple[Any, str]]) -> tuple[Any, str]:
    for valor, origem in candidatos:
        if _numero_disponivel(valor):
            return valor, origem
    return None, ""


def avaliar_capacidades_apuracao(
    contagens: dict[str, Any] | None,
    metadados: dict[str, Any] | None,
    bloqueios_estruturais: Iterable[str] | None = None,
    lacunas_apuracao: Iterable[str] | None = None,
) -> dict[str, Any]:
    """Avalia cada bloco sem transformar ausencia de dados em bloqueio global."""

    contagens = contagens or {}
    metadados = metadados or {}
    bloqueios = list(bloqueios_estruturais or [])
    lacunas = list(lacunas_apuracao or [])
    estruturalmente_valido = not bloqueios

    qtd_financeiro = int(contagens.get("competencias_com_valor") or 0)
    qtd_remanescentes = int(contagens.get("itens_remanescentes") or 0)
    qtd_consumidos = int(contagens.get("itens_consumidos") or 0)
    qtd_pcs = int(contagens.get("pedidos_de_compra") or 0)
    qtd_aditivos = int(contagens.get("aditivos") or 0)
    qtd_posicao = int(contagens.get("posicao_contratual_itens") or 0)
    qtd_posicao_calculada = int(contagens.get("posicao_contratual_calculada") or 0)
    qtd_vu = int(contagens.get("historico_vu_itens") or 0)
    qtd_vu_calculado = int(contagens.get("historico_vu_calculado") or 0)
    ciclos = list(metadados.get("ciclos_em_analise") or [])

    tem_financeiro = qtd_financeiro > 0
    tem_remanescentes = qtd_remanescentes > 0
    tem_consumidos = qtd_consumidos > 0
    tem_pcs = qtd_pcs > 0
    tem_itens = tem_remanescentes or tem_consumidos or qtd_posicao > 0 or qtd_vu > 0
    tem_aditivos = qtd_aditivos > 0
    tem_ciclos = bool(ciclos)
    tem_alguma_evidencia = any(
        (tem_financeiro, tem_itens, tem_pcs, tem_remanescentes, tem_consumidos, tem_aditivos)
    )

    if not estruturalmente_valido:
        detalhe_bloqueio = "A estrutura do XLS precisa ser corrigida antes da leitura dos blocos."
        blocos = {
            chave: _registro(nome, ESTADO_BLOQUEADO, "Estrutura inválida", detalhe_bloqueio)
            for chave, nome in (
                ("financeiro", "Financeiro"),
                ("itens", "Itens"),
                ("pcs", "PCs"),
                ("consumidos", "Consumidos"),
                ("remanescentes", "Remanescentes"),
            )
        }
    else:
        blocos = {
            "financeiro": _registro(
                "Financeiro",
                ESTADO_COMPLETO if tem_financeiro else ESTADO_NAO_INFORMADO,
                "Completo" if tem_financeiro else "Não informado",
                f"{qtd_financeiro} competência(s) com valor." if tem_financeiro else "Nenhuma competência financeira preenchida.",
                disponivel=tem_financeiro,
            ),
            "itens": _registro(
                "Itens",
                ESTADO_COMPLETO if tem_itens else ESTADO_NAO_INFORMADO,
                "Completo" if tem_itens else "Não informado",
                (
                    f"{qtd_remanescentes} remanescente(s), {qtd_consumidos} consumido(s) e "
                    f"{qtd_posicao or qtd_vu} item(ns) na evolução contratual."
                    if tem_itens
                    else "Nenhum item contratual identificado."
                ),
                disponivel=tem_itens,
            ),
            "pcs": _registro(
                "PCs",
                ESTADO_COMPLETO if tem_pcs else ESTADO_NAO_INFORMADO,
                "Completo" if tem_pcs else "Não informado",
                f"{qtd_pcs} pedido(s) de compra informado(s)." if tem_pcs else "Nenhum pedido de compra informado.",
                disponivel=tem_pcs,
            ),
            "consumidos": _registro(
                "Consumidos",
                ESTADO_COMPLETO if tem_consumidos else ESTADO_NAO_INFORMADO,
                "Completo" if tem_consumidos else "Não informado",
                f"{qtd_consumidos} item(ns) consumido(s)." if tem_consumidos else "Nenhum item consumido informado.",
                disponivel=tem_consumidos,
            ),
            "remanescentes": _registro(
                "Remanescentes",
                ESTADO_COMPLETO if tem_remanescentes else ESTADO_NAO_INFORMADO,
                "Completo" if tem_remanescentes else "Não informado",
                f"{qtd_remanescentes} item(ns) remanescente(s)." if tem_remanescentes else "Nenhum saldo remanescente informado.",
                disponivel=tem_remanescentes,
            ),
        }

    status = metadados.get("status_resultados") or {}
    valores = status.get("valores") or {}
    status_retro = _texto(status.get("retroativo"))
    status_vta = _texto(status.get("vta"))
    status_remanescente = _texto(status.get("remanescente"))

    retro_metodo, origem_retro = _primeiro_valor(
        (
            (valores.get("retroativo_financeiro"), "Financeiro") if tem_financeiro else (None, ""),
            (valores.get("retroativo_pc"), "PCs") if tem_pcs else (None, ""),
            (valores.get("retroativo_itens"), "Itens") if tem_itens else (None, ""),
        )
    )
    retro_oficial = valores.get("retroativo_oficial")
    retro_oficial_seguro = _numero_disponivel(retro_oficial) and any(
        termo in status_retro for termo in ("VALIDADO", "CONFERIR", "CALCULADO")
    )
    if retro_oficial_seguro:
        retroativo = _registro(
            "Retroativo", ESTADO_COMPLETO, "Calculado", "Valor oficial preservado no XLS.",
            disponivel=True, valor=retro_oficial, origem="RESULTADOS",
        )
    elif _numero_disponivel(retro_metodo):
        retroativo = _registro(
            "Retroativo", ESTADO_PARCIAL, "Parcial",
            f"Metodologia {origem_retro} calculada; a seleção do valor oficial ainda requer conclusão.",
            disponivel=True, valor=retro_metodo, origem=origem_retro,
        )
    elif tem_financeiro or tem_pcs or tem_itens:
        retroativo = _registro(
            "Retroativo", ESTADO_PARCIAL, "Aguardando cálculo do XLS",
            "Há base para apuração, mas o arquivo ainda não contém valor calculado em cache.",
        )
    else:
        retroativo = _registro(
            "Retroativo", ESTADO_NAO_INFORMADO, "Não informado",
            "Depende de Financeiro, PCs ou Itens.",
        )

    vta_oficial = valores.get("vta_oficial")
    vta_calculado = valores.get("vta_calculado")
    vta_oficial_seguro = _numero_disponivel(vta_oficial) and any(
        termo in status_vta for termo in ("VALIDADO", "CONFERIR", "CALCULADO")
    )
    if vta_oficial_seguro:
        vta = _registro(
            "VTA", ESTADO_COMPLETO, "Calculado", "Valor Total Atualizado oficial preservado no XLS.",
            disponivel=True, valor=vta_oficial, origem="RESULTADOS",
        )
    elif _numero_disponivel(vta_calculado) and (tem_financeiro or tem_remanescentes or tem_itens):
        vta = _registro(
            "VTA", ESTADO_PARCIAL, "Calculado com ressalva",
            "A composição automática está disponível; eventual ajuste/manualização oficial permanece pendente.",
            disponivel=True, valor=vta_calculado, origem="Composição automática do XLS",
        )
    elif tem_financeiro or tem_remanescentes or tem_itens:
        vta = _registro(
            "VTA", ESTADO_PARCIAL, "Aguardando cálculo do XLS",
            "Há dados relacionados ao contrato, mas o VTA ainda não foi gravado em cache.",
        )
    else:
        vta = _registro("VTA", ESTADO_NAO_INFORMADO, "Não informado", "Depende da composição contratual.")

    rem_valor = valores.get("remanescente_atualizado")
    rem_status_seguro = _numero_disponivel(rem_valor) and any(
        termo in status_remanescente for termo in ("VALIDADO", "CONFERIR", "CALCULADO")
    )
    if rem_status_seguro:
        valor_remanescente = _registro(
            "Valor remanescente", ESTADO_COMPLETO, "Calculado",
            "Saldo remanescente atualizado preservado no XLS.",
            disponivel=True, valor=rem_valor, origem="RESULTADOS",
        )
    elif tem_remanescentes and _numero_disponivel(rem_valor):
        valor_remanescente = _registro(
            "Valor remanescente", ESTADO_PARCIAL, "Calculado com ressalva",
            "O valor foi calculado, mas o status oficial do saldo ainda requer conferência.",
            disponivel=True, valor=rem_valor, origem="RESULTADOS",
        )
    elif tem_remanescentes:
        valor_remanescente = _registro(
            "Valor remanescente", ESTADO_PARCIAL, "Aguardando cálculo do XLS",
            "Os itens estão informados, mas o saldo atualizado ainda não foi gravado em cache.",
        )
    else:
        valor_remanescente = _registro(
            "Valor remanescente", ESTADO_NAO_INFORMADO, "Não informado",
            "Depende de itens remanescentes.",
        )

    if qtd_posicao > 0 and qtd_posicao_calculada > 0:
        posicao = _registro(
            "Posição contratual", ESTADO_COMPLETO, "Calculada",
            f"{qtd_posicao_calculada} item(ns) com posição vigente por ciclo.", disponivel=True,
        )
    elif tem_itens:
        posicao = _registro(
            "Posição contratual", ESTADO_PARCIAL, "Aguardando cálculo do XLS",
            "Há itens, mas a evolução quantitativa ainda não foi gravada em cache.",
        )
    else:
        posicao = _registro(
            "Posição contratual", ESTADO_NAO_INFORMADO, "Não informada", "Depende dos itens contratuais."
        )

    if qtd_vu > 0 and qtd_vu_calculado > 0:
        valores_unitarios = _registro(
            "Valores unitários", ESTADO_COMPLETO, "Calculados",
            f"{qtd_vu_calculado} item(ns) com evolução de VU.", disponivel=True,
        )
    elif tem_itens or tem_pcs:
        valores_unitarios = _registro(
            "Valores unitários", ESTADO_PARCIAL, "Aguardando cálculo do XLS",
            "Há base de itens, mas a evolução dos valores ainda não foi gravada em cache.",
        )
    else:
        valores_unitarios = _registro(
            "Valores unitários", ESTADO_NAO_INFORMADO, "Não informados", "Depende de Itens ou PCs."
        )

    calculos = {
        "retroativo": retroativo,
        "vta": vta,
        "posicao_contratual": posicao,
        "valor_remanescente": valor_remanescente,
        "valores_unitarios": valores_unitarios,
    }

    algum_calculo = any(item["disponivel"] for item in calculos.values())
    base_documental = tem_alguma_evidencia or tem_ciclos
    memoria_disponivel = posicao["disponivel"] or valores_unitarios["disponivel"] or retroativo["disponivel"]
    formalizacao_pronta = retroativo["estado"] == ESTADO_COMPLETO and vta["estado"] == ESTADO_COMPLETO

    documentos = {
        "planilha_executiva": _documento(
            "Planilha Executiva",
            ESTADO_COMPLETO if algum_calculo else ESTADO_PARCIAL if base_documental else ESTADO_NAO_INFORMADO,
            "Disponível" if algum_calculo else "Disponível com ressalvas" if base_documental else "Pendente de dados",
            "Reúne somente os blocos já apurados." if base_documental else "Envie uma coleta com ao menos um bloco preenchido.",
            habilitado=base_documental,
        ),
        "valores_unitarios": _documento(
            "Itens por Ciclo",
            ESTADO_COMPLETO if valores_unitarios["disponivel"] else ESTADO_PARCIAL if (tem_itens or tem_pcs) else ESTADO_NAO_INFORMADO,
            "Disponível" if valores_unitarios["disponivel"] else "Pendente de dados",
            valores_unitarios["detalhe"], habilitado=valores_unitarios["disponivel"],
        ),
        "relatorio_executivo": _documento(
            "Relatório Executivo",
            ESTADO_COMPLETO if formalizacao_pronta else ESTADO_PARCIAL if base_documental else ESTADO_NAO_INFORMADO,
            "Disponível" if formalizacao_pronta else "Disponível com ressalvas" if base_documental else "Pendente de dados",
            "Os blocos pendentes serão identificados no relatório." if base_documental else "Depende de evidências da apuração.",
            habilitado=base_documental,
        ),
        "minuta_apostilamento": _documento(
            "Termo de Apostila",
            ESTADO_COMPLETO if formalizacao_pronta else ESTADO_BLOQUEADO,
            "Disponível" if formalizacao_pronta else "Pendente de dados",
            "Retroativo e VTA oficiais estão concluídos." if formalizacao_pronta else "Depende da conclusão do retroativo e do VTA oficial.",
            habilitado=formalizacao_pronta,
        ),
        "mapa_marcos": _documento(
            "Memória de Cálculo e Marcos",
            ESTADO_COMPLETO if memoria_disponivel else ESTADO_PARCIAL if tem_ciclos else ESTADO_NAO_INFORMADO,
            "Disponível" if memoria_disponivel else "Disponível com ressalvas" if tem_ciclos else "Pendente de dados",
            "Registra os marcos e apenas os cálculos disponíveis." if tem_ciclos else "Depende da identificação dos ciclos.",
            habilitado=tem_ciclos,
        ),
        "checklist_processual": _documento(
            "Checklist Processual", ESTADO_COMPLETO, "Disponível",
            "Independe da conclusão dos cálculos e sinaliza as pendências.", habilitado=True,
        ),
        "garantia_contratual": _documento(
            "Garantia Contratual",
            ESTADO_COMPLETO if vta["estado"] == ESTADO_COMPLETO else ESTADO_PARCIAL if vta["disponivel"] else ESTADO_BLOQUEADO,
            "Disponível" if vta["estado"] == ESTADO_COMPLETO else "Disponível com ressalvas" if vta["disponivel"] else "Pendente de dados",
            "Utiliza o VTA disponível e explicita eventual ressalva." if vta["disponivel"] else "Depende de um Valor Total Atualizado calculado.",
            habilitado=vta["disponivel"],
        ),
        "dou": _documento(
            "DOU",
            ESTADO_COMPLETO if formalizacao_pronta else ESTADO_PARCIAL if base_documental else ESTADO_NAO_INFORMADO,
            "Disponível" if formalizacao_pronta else "Disponível com ressalvas" if base_documental else "Pendente de dados",
            "A minuta pode ser preparada; valores pendentes permanecem sinalizados." if base_documental else "Depende de dados do processo.",
            habilitado=base_documental,
        ),
        "avaliacao_aditivos": _documento(
            "Avaliação de Aditivos", ESTADO_COMPLETO, "Disponível",
            "Módulo independente para registrar e avaliar alterações contratuais.", habilitado=True,
        ),
        "infos_previas": _documento(
            "Infos Prévias", ESTADO_COMPLETO, "Disponível",
            "Módulo independente para organizar as informações de instrução.", habilitado=True,
        ),
        "saneador": _documento(
            "Saneador", ESTADO_COMPLETO if formalizacao_pronta else ESTADO_PARCIAL,
            "Disponível" if formalizacao_pronta else "Disponível com ressalvas",
            "Integra os dados existentes e mantém as lacunas explicitamente marcadas.", habilitado=True,
        ),
    }

    completos = sum(item["estado"] == ESTADO_COMPLETO for item in {**blocos, **calculos}.values())
    pendentes = sum(item["estado"] in (ESTADO_PARCIAL, ESTADO_NAO_INFORMADO) for item in {**blocos, **calculos}.values())
    return {
        "estruturalmente_valido": estruturalmente_valido,
        "processamento_progressivo": True,
        "blocos": blocos,
        "calculos": calculos,
        "documentos": documentos,
        "resumo": {
            "completos": completos,
            "pendentes": pendentes,
            "documentos_habilitados": sum(item["habilitado"] for item in documentos.values()),
            "documentos_total": len(documentos),
            "tem_alguma_evidencia": tem_alguma_evidencia,
            "tem_algum_calculo": algum_calculo,
            "apuracao_integral": formalizacao_pronta,
        },
        "bloqueios_estruturais": bloqueios,
        "lacunas_apuracao": lacunas,
    }
