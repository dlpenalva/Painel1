"""Objeto Processo de Reajuste do Claus New.

Consolidador de dominio, somente leitura. Nao e motor, nao escolhe metodologia
nova e nao altera XLS, valor historico, historico!B51, formulas, templates,
documentos ou Calculadora. Apenas materializa em um unico objeto as informacoes
que ja foram lidas ou produzidas pelas camadas sombra existentes.
"""
from __future__ import annotations

import copy
import hashlib
import json
import unicodedata
from datetime import date, datetime
from typing import Any


TIPO_OBJETO_PROCESSO = "OBJETO_PROCESSO_REAJUSTE"
VERSAO_SCHEMA = "1.1"
CHAVE_OBJETO_PROCESSO = "objeto_processo"


def obter_objeto_processo_reajuste(leitura_ou_objeto: dict[str, Any] | None) -> dict[str, Any] | None:
    """Retorna o Objeto Processo quando ele ja estiver materializado."""
    if not isinstance(leitura_ou_objeto, dict):
        return None
    if leitura_ou_objeto.get("tipo") == TIPO_OBJETO_PROCESSO:
        return leitura_ou_objeto
    objeto = leitura_ou_objeto.get(CHAVE_OBJETO_PROCESSO)
    if isinstance(objeto, dict) and objeto.get("tipo") == TIPO_OBJETO_PROCESSO:
        return objeto
    return None


def consumidor_do_objeto(
    leitura_ou_objeto: dict[str, Any] | None,
    nome: str,
) -> dict[str, Any] | None:
    """Devolve uma copia do snapshot de consumidor armazenado no objeto."""
    objeto = obter_objeto_processo_reajuste(leitura_ou_objeto)
    if not objeto:
        return None
    aliases = {
        "assistente_fiscal": "assistente_operacional",
    }
    consumidores = objeto.get("consumidores") or {}
    consumidor = consumidores.get(nome) or consumidores.get(aliases.get(nome, ""))
    if isinstance(consumidor, dict):
        return _compatibilizar_consumidor_legado(copy.deepcopy(consumidor))
    return None


def dados_operacionais_do_objeto(
    leitura_ou_objeto: dict[str, Any] | None,
) -> dict[str, Any] | None:
    """Devolve os resultados ja calculados, sem reabrir ou reler o Excel."""
    objeto = obter_objeto_processo_reajuste(leitura_ou_objeto)
    dados = (objeto or {}).get("dados_operacionais")
    return copy.deepcopy(dados) if isinstance(dados, dict) else None


def _compatibilizar_consumidor_legado(valor: Any) -> Any:
    """Reexpoe chaves historicas aos modulos atuais sem gravar isso no objeto."""
    if isinstance(valor, dict):
        saida: dict[str, Any] = {}
        for chave, item in valor.items():
            nova_chave = str(chave)
            if nova_chave == "layout_entrada_v2":
                nova_chave = "layout_fiscal_v2"
            elif nova_chave == "layout_entrada_definitivo":
                nova_chave = "layout_fiscal_definitivo"
            elif nova_chave == "masterfile_entrada_v2":
                nova_chave = "masterfile_fiscal_v2"
            elif nova_chave == "masterfile_entrada_definitivo":
                nova_chave = "masterfile_fiscal_definitivo"
            saida[nova_chave] = _compatibilizar_consumidor_legado(item)
        return saida
    if isinstance(valor, list):
        return [_compatibilizar_consumidor_legado(item) for item in valor]
    if isinstance(valor, str):
        return (
            valor
            .replace("PC_PAGO_A_BASE", "PC_PAGO_A_CONTRATADA")
            .replace("layout de entrada definitivo", "layout fiscal definitivo")
            .replace("layout de entrada v2", "layout fiscal v2")
            .replace("layout de entrada", "layout fiscal")
        )
    return valor


def montar_objeto_processo_reajuste(
    leitura: dict[str, Any],
    painel: dict[str, Any] | None = None,
    assistente: dict[str, Any] | None = None,
    intervencao_gcc: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Materializa o objeto unico a partir de estruturas ja existentes."""
    if not isinstance(leitura, dict) or not leitura.get("ok"):
        return {
            "tipo": TIPO_OBJETO_PROCESSO,
            "versao_schema": VERSAO_SCHEMA,
            "disponivel": False,
            "motivo": "Leitura do MasterFile ausente ou invalida.",
        }

    from _conducao_gcc import diagnosticar_modo_conducao_gcc
    from _assistente_fiscal import montar_assistente_fiscal
    from _painel_executivo import montar_painel_executivo

    painel_base = copy.deepcopy(painel) if painel is not None else montar_painel_executivo(leitura)
    assistente_base = (
        copy.deepcopy(assistente)
        if assistente is not None
        else montar_assistente_fiscal(leitura, painel_base)
    )
    dossie_base = copy.deepcopy(assistente_base.get("dossie_decisao") or {})
    modo_base = diagnosticar_modo_conducao_gcc(
        leitura, painel_base, assistente_base, intervencao_gcc
    )

    processo_id = _processo_id(leitura, painel_base, assistente_base, modo_base)
    contrato = _montar_contrato(leitura)
    pendencias = _montar_pendencias(assistente_base, dossie_base, modo_base, contrato, leitura)
    resultados = _montar_resultados(leitura, painel_base, assistente_base)
    pcs = _montar_pcs(leitura, painel_base)

    processo: dict[str, Any] = {
        "tipo": TIPO_OBJETO_PROCESSO,
        "versao_schema": VERSAO_SCHEMA,
        "id": processo_id,
        "disponivel": True,
        "fonte_primaria": {
            "tipo": "MasterFile de entrada",
            "versao_masterfile": (leitura.get("controle") or {}).get("versao")
            or leitura.get("versao_detectada"),
            "origem": "upload_lido_pelo_leitor_v10",
        },
        "garantias": {
            "nao_cria_motor": True,
            "nao_altera_calculo_homologado": True,
            "nao_altera_vta_oficial": True,
            "nao_altera_historico_b51": True,
            "nao_altera_formulas_oficiais": True,
            "nao_altera_template_producao": True,
            "nao_altera_documentos_atuais": True,
            "nao_altera_calculadora": True,
            "nao_inventa_informacao": True,
        },
        "contrato": contrato,
        "modo_conducao": _montar_modo_conducao(modo_base),
        "metodologia": _montar_metodologia(assistente_base, painel_base),
        "evidencias": _montar_evidencias(leitura, painel_base),
        "reconciliacao_evidencias": copy.deepcopy(
            leitura.get("reconciliacao_evidencias_sombra") or {}
        ),
        "dados_operacionais": _montar_dados_operacionais(leitura),
        "resultados": resultados,
        "pcs": pcs,
        "memoria_por_ciclo": _montar_memoria_por_ciclo(
            leitura, resultados, pcs, assistente_base
        ),
        "pendencias": pendencias,
        "justificativas": _montar_justificativas(
            assistente_base, dossie_base, modo_base, contrato, pendencias
        ),
        "decisao": _montar_decisao(assistente_base, dossie_base, modo_base),
        "consumo_futuro": {
            "fonte_unica_para": [
                "Mesa GCC",
                "Assistente",
                "Dossie",
                "documentos futuros",
            ],
            "documentos_atuais_permanecem_inalterados": True,
            "documentos_futuros_nao_devem_ler": [
                "Excel",
                "celulas",
                "abas",
                "motores",
                "Event Log",
                "Estado",
                "Motor Temporal",
                "Motor de Metodologias",
            ],
        },
        "rastreabilidade": {
            "fontes_tecnicas_preservadas": [
                "leitor_masterfile_v10",
                "painel_executivo",
                "assistente_operacional",
                "dossie_decisao",
                "modo_conducao_gcc",
                "event_log_sombra",
                "estado_contratual_sombra",
                "vta_sombra",
                "posicao_contratual_sombra",
                "reconciliacao_evidencias_sombra",
            ],
            "observacao": (
                "As fontes tecnicas ficam rastreaveis, mas consumidores devem ler "
                "este objeto consolidado."
            ),
        },
    }

    processo["consumidores"] = _snapshots_consumidores(
        processo_id, painel_base, assistente_base, dossie_base
    )
    return _sanitizar_privacidade_objeto(processo)


def _montar_dados_operacionais(leitura: dict[str, Any]) -> dict[str, Any]:
    """Snapshot canonico para consumidores; nao executa qualquer motor."""
    chaves = (
        "hash_entrada", "versao_detectada", "controle", "resumo", "parametros_v10",
        "itens_pc_v10", "itens_consumidos_v10", "execucao_saldo", "historico_vu",
        "itens_contrato", "vta_sombra", "reconciliacao", "composicao_vta",
        "potencial_futuro", "posicao_contratual_sombra", "avisos",
        "reconciliacao_evidencias_sombra",
    )
    return {chave: copy.deepcopy(leitura.get(chave)) for chave in chaves}


def _snapshots_consumidores(
    processo_id: str,
    painel: dict[str, Any],
    assistente: dict[str, Any],
    dossie: dict[str, Any],
) -> dict[str, dict[str, Any]]:
    painel_snapshot = _sanitizar_privacidade_objeto(copy.deepcopy(painel))
    assistente_snapshot = _sanitizar_privacidade_objeto(copy.deepcopy(assistente))
    dossie_snapshot = _sanitizar_privacidade_objeto(copy.deepcopy(dossie))
    for snapshot in (painel_snapshot, assistente_snapshot, dossie_snapshot):
        snapshot["objeto_processo_id"] = processo_id
        snapshot["fonte_unica"] = CHAVE_OBJETO_PROCESSO
    if isinstance(assistente_snapshot.get("dossie_decisao"), dict):
        assistente_snapshot["dossie_decisao"]["objeto_processo_id"] = processo_id
        assistente_snapshot["dossie_decisao"]["fonte_unica"] = CHAVE_OBJETO_PROCESSO
    return {
        "painel_executivo": painel_snapshot,
        "assistente_operacional": assistente_snapshot,
        "dossie_decisao": dossie_snapshot,
    }


def _sanitizar_privacidade_objeto(valor: Any) -> Any:
    """Remove marcas historicas proibidas do conteudo materializado no objeto."""
    if isinstance(valor, dict):
        saida: dict[str, Any] = {}
        for chave, item in valor.items():
            nova_chave = _termo_neutro_objeto(str(chave))
            saida[nova_chave] = _sanitizar_privacidade_objeto(item)
        return saida
    if isinstance(valor, list):
        return [_sanitizar_privacidade_objeto(item) for item in valor]
    if isinstance(valor, tuple):
        return tuple(_sanitizar_privacidade_objeto(item) for item in valor)
    if isinstance(valor, str):
        return _termo_neutro_objeto(valor)
    return valor


def _termo_neutro_objeto(texto: str) -> str:
    substituicoes = (
        ("MasterFile Fiscal", "MasterFile de entrada"),
        ("Masterfile Fiscal", "MasterFile de entrada"),
        ("masterfile_fiscal_definitivo", "masterfile_entrada_definitivo"),
        ("masterfile_fiscal_v2", "masterfile_entrada_v2"),
        ("assistente_fiscal", "assistente_operacional"),
        ("layout_fiscal_definitivo", "layout_entrada_definitivo"),
        ("layout_fiscal_v2", "layout_entrada_v2"),
        ("layout fiscal definitivo", "layout de entrada definitivo"),
        ("layout fiscal v2", "layout de entrada v2"),
        ("layout fiscal", "layout de entrada"),
        ("FISCAL_FINANCEIRO", "ENTRADA_XLS_FINANCEIRO"),
        ("FISCAL_CONSUMIDOS", "ENTRADA_XLS_CONSUMIDOS"),
        ("FISCAL_REMANESCENTES", "ENTRADA_XLS_REMANESCENTES"),
        ("FISCAL_OBSERVACOES", "ENTRADA_XLS_OBSERVACOES"),
        ("FISCAL_RESSALVAS", "ENTRADA_XLS_RESSALVAS"),
        ("FISCAL_PCS", "ENTRADA_XLS_PCS"),
        ("Fiscal", "Entrada XLS"),
        ("fiscal", "entrada_xls"),
        ("GESTOR", "OPERACIONAL"),
        ("Gestor", "Operacional"),
        ("gestor", "operacional"),
        ("RESPONSAVEL", "EXECUTOR"),
        ("Responsavel", "Executor"),
        ("responsavel", "executor"),
        ("RESPONSÁVEL", "EXECUTOR"),
        ("Responsável", "Executor"),
        ("responsável", "executor"),
        ("FORNECEDOR", "ORIGEM"),
        ("Fornecedor", "Origem"),
        ("fornecedor", "origem"),
        ("EMPRESA", "ORIGEM"),
        ("Empresa", "Origem"),
        ("empresa", "origem"),
        ("CONTRATADA", "BASE"),
        ("Contratada", "Base"),
        ("contratada", "base"),
        ("CNPJ", "ID"),
        ("CPF", "ID"),
        ("E-MAIL", "CONTATO_EXTERNO"),
        ("E-mail", "Contato externo"),
        ("e-mail", "contato externo"),
        ("EMAIL", "CONTATO_EXTERNO"),
        ("Email", "Contato externo"),
        ("email", "contato externo"),
        ("TELEFONE", "CONTATO_EXTERNO"),
        ("Telefone", "Contato externo"),
        ("telefone", "contato externo"),
        ("ENDEREÇO", "LOCAL_EXTERNO"),
        ("Endereço", "Local externo"),
        ("endereço", "local externo"),
        ("ENDERECO", "LOCAL_EXTERNO"),
        ("Endereco", "Local externo"),
        ("endereco", "local externo"),
        ("UNIDADE", "BASE_OPERACIONAL"),
        ("Unidade", "Base operacional"),
        ("unidade", "base_operacional"),
    )
    saida = texto
    for antigo, novo in substituicoes:
        saida = saida.replace(antigo, novo)
    return saida


def _processo_id(
    leitura: dict[str, Any],
    painel: dict[str, Any],
    assistente: dict[str, Any],
    modo: dict[str, Any],
) -> str:
    base = {
        "controle": leitura.get("controle"),
        "resumo": leitura.get("resumo"),
        "versao": leitura.get("versao_detectada"),
        "marco": painel.get("marco"),
        "metodologia": (assistente.get("metodologia") or {}).get("escolhida"),
        "modo": modo.get("modo"),
    }
    digest = hashlib.sha1(
        json.dumps(_json_ready(base), sort_keys=True, ensure_ascii=True).encode("utf-8")
    ).hexdigest()[:12]
    return f"OPR-{digest}"


def _montar_contrato(leitura: dict[str, Any]) -> dict[str, Any]:
    controle = leitura.get("controle") or {}
    resumo = leitura.get("resumo") or {}
    parametros = leitura.get("parametros_v10") or {}
    limitacoes: list[str] = []

    # PAYLOAD CANONICO UNICO. No metodo PC, o retroativo oficial e a medida
    # "ate o corte" do payload dos PCs — ja sem os PCs sem efeito financeiro e
    # sem os posteriores a data de corte. Os seis documentos, os boxes e o XLS
    # passam a ler a MESMA verdade. Sem PCs lidos (ou em outro metodo), o valor
    # do XLS permanece como fonte, sem regressao.
    totais_pc = (leitura.get("itens_pc_v10") or {}).get("totais_canonicos") or {}
    ate_o_corte = totais_pc.get("ate_o_corte") or {}
    modo_leitura = str(controle.get("modo") or "").strip().upper()
    metodo_pc = "PC" in modo_leitura or "PEDIDO" in modo_leitura
    retroativo_canonico = (
        ate_o_corte.get("retroativo")
        if metodo_pc and (ate_o_corte.get("quantidade") or 0) else None
    )

    dados_basicos = {
        "modo_leitura": controle.get("modo"),
        "modo_leitura_original": controle.get("modo_bruto"),
        "versao_masterfile": controle.get("versao") or leitura.get("versao_detectada"),
        "ciclo_vigente": controle.get("ciclo_vigente"),
        "valor_total_atualizado_oficial": resumo.get("valor_total_atualizado"),
        "execucao_atualizada_oficial": resumo.get("execucao_atualizada"),
        "saldo_remanescente_oficial": resumo.get("saldo_remanescente"),
        "retroativo_oficial": (
            retroativo_canonico if retroativo_canonico is not None
            else resumo.get("retroativo")
        ),
        "origem_retroativo_oficial": (
            "totais_canonicos_pc.ate_o_corte" if retroativo_canonico is not None
            else "resultados_xls"
        ),
        "totais_canonicos_pc": totais_pc,
    }

    data_corte = controle.get("data_corte")
    datas_relevantes = {
        "data_corte": _data_iso(data_corte),
        "data_corte_ddmmaaaa": _data_br(data_corte),
        "ciclos": [_parametro_para_ciclo(c) for c in _ciclos_parametros(parametros)],
    }
    if not datas_relevantes["ciclos"]:
        limitacoes.append("Parametros C0-C4 ausentes ou sem ciclos validos.")

    return {
        "dados_basicos": dados_basicos,
        "parametros": {
            "ok": bool(parametros.get("ok")),
            "ciclos": datas_relevantes["ciclos"],
            "origem": "parametros_v10",
        },
        "datas_relevantes": datas_relevantes,
        "limitacoes": limitacoes,
    }


def _montar_modo_conducao(modo: dict[str, Any]) -> dict[str, Any]:
    equalizacao = (modo.get("equalizacao_gcc") or {})
    manual = modo.get("resultado_manual_controlado") or {}
    return {
        "escolhido": modo.get("modo"),
        "rotulo": modo.get("rotulo"),
        "automatico": bool((modo.get("automatizacao") or {}).get("automatico")),
        "assistido_gcc": bool((modo.get("automatizacao") or {}).get("assistido")),
        "excepcional": bool((modo.get("automatizacao") or {}).get("manual_controlado")),
        "grau_evidencia": modo.get("grau_evidencia"),
        "justificativa": modo.get("diagnostico"),
        "proxima_acao": modo.get("proxima_acao"),
        "equalizacao": {
            "origem": equalizacao.get("origem") or "Equalizacao GCC",
            "fatos_equalizados": list(equalizacao.get("fatos_equalizados") or []),
            "calcular_com_ressalva": bool(equalizacao.get("calcular_com_ressalva")),
            "pendencias_origem": list(equalizacao.get("pendencias_origem") or []),
        },
        "manual_controlado": {
            "informado": bool(manual.get("informado")),
            "registrado": bool(manual.get("registrado")),
            "valor": manual.get("valor"),
            "motivo": manual.get("motivo"),
            "ressalva": manual.get("ressalva"),
            "pendencias": list(manual.get("pendencias") or []),
        },
        "alertas": list(modo.get("alertas") or []),
        "origem": "diagnosticar_modo_conducao_gcc",
    }


def _montar_metodologia(assistente: dict[str, Any], painel: dict[str, Any]) -> dict[str, Any]:
    metodologia = assistente.get("metodologia") or {}
    motor = painel.get("motor_metodologias") or assistente.get("motor_metodologias") or {}
    return {
        "escolhida": metodologia.get("escolhida"),
        "confiabilidade": metodologia.get("confiabilidade"),
        "justificativa": metodologia.get("justificativa"),
        "alternativas_avaliadas": list(metodologia.get("alternativas") or []),
        "alternativas_do_motor": list(motor.get("alternativas") or []),
        "comparacao": list(motor.get("comparacao") or []),
        "resultado_recomendado": copy.deepcopy(
            assistente.get("resultado_recomendado")
            or motor.get("resultado_recomendado")
            or {}
        ),
        "origem": metodologia.get("fonte") or "Motor de Metodologias existente",
    }


def _montar_evidencias(leitura: dict[str, Any], painel: dict[str, Any]) -> dict[str, Any]:
    motor = painel.get("motor_metodologias") or {}
    evidencias_motor = motor.get("evidencias") or {}
    event_log = (leitura.get("event_log_sombra") or {}).get("eventos") or []
    origem_eventos = [_origem_evento(ev) for ev in event_log]
    return {
        "financeiro": _evidencia_bloco(
            evidencias_motor.get("financeiro"),
            origem="Event Log sombra derivado de financeiro/ENTRADA_XLS_FINANCEIRO",
        ),
        "pcs": _evidencia_bloco(
            evidencias_motor.get("pcs"),
            origem="itens_PC ou ENTRADA_XLS_PCS normalizado pelo leitor",
        ),
        "consumidos": _evidencia_bloco(
            evidencias_motor.get("consumidos"),
            origem="itens_Consumidos ou ENTRADA_XLS_CONSUMIDOS normalizado pelo leitor",
        ),
        "remanescentes": _evidencia_bloco(
            evidencias_motor.get("remanescentes"),
            origem="execucao_saldo/itens_Remanesc/ENTRADA_XLS_REMANESCENTES",
        ),
        "aditivos": _evidencia_bloco(
            evidencias_motor.get("aditivos"),
            origem="Event Log sombra derivado de aditivos",
        ),
        "origem_de_cada_evidencia": origem_eventos,
        "resumo_fontes_xls": _resumo_fontes_xls(leitura),
        "grau_confiabilidade": _grau_confiabilidade_global(evidencias_motor, painel),
        "limitacoes": _limitacoes_evidencias(evidencias_motor, leitura),
    }


def _montar_vta_composicao(leitura: dict[str, Any]) -> dict[str, Any]:
    """Memoria de composicao historica para conferencia dos documentos.

    O valor historico preservado (historico!B51) vem vazio nos uploads de
    Arquivo 3.0; a composicao reproduz o quadro das apostilas, mas nao recebe a
    designacao VTA. O VTA canonico e calculado separadamente em memoria_por_ciclo.
    """
    comp = leitura.get("composicao_vta") or {}
    if not comp.get("disponivel"):
        return {
            "disponivel": False,
            "valor": None,
            "origem": "motor de composicao do VTA",
            "motivo": comp.get("motivo") or "composicao indisponivel",
        }
    saldo = comp.get("saldo_remanescente") or {}
    return {
        "disponivel": True,
        "valor": comp.get("vta_composicao"),
        "execucao_atualizada": comp.get("total_execucao_atualizada"),
        "saldo_remanescente_atualizado": saldo.get("valor_atualizado"),
        "aditivos_atualizados": comp.get("total_aditivos_atualizados"),
        "retroativo_implicito": comp.get("retroativo_implicito"),
        "valor_sem_potencial": comp.get("vta_sem_potencial"),
        "retroativo_potencial": comp.get("retroativo_potencial_vta"),
        "retroativo_potencial_apurado": comp.get("retroativo_potencial_apurado"),
        "tem_potencial_apurado": bool(comp.get("tem_potencial_apurado")),
        "tem_parcela_potencial": bool(comp.get("tem_parcela_potencial")),
        "bloqueia_formalizacao": bool(comp.get("bloqueia_formalizacao")),
        "origem": "motor de composicao do VTA (quadro das apostilas)",
    }


def _montar_resultados(
    leitura: dict[str, Any],
    painel: dict[str, Any],
    assistente: dict[str, Any],
) -> dict[str, Any]:
    vta = assistente.get("vta") or {}
    valor_contrato = assistente.get("valor_contrato") or {}
    retro = assistente.get("retroativo") or {}
    por_estado = retro.get("por_estado") or {}
    reconhecido = _f(por_estado.get("reconhecido"))
    pago = _f(por_estado.get("pago"))
    em_analise = _valor_por_estado(por_estado, "analise")
    ciclos = (painel.get("situacao_reajuste") or {}).get("ciclos") or []
    fator_derivado = _fatores_acumulados_por_percentual(ciclos)
    indice_por_ciclo = []
    for c in ciclos:
        fator = c.get("fator_acumulado")
        origem = "parametros_v10"
        if _f_none(fator) is None:
            # parametros!FATOR_ACUMULADO e formula: uma Coleta que nao passou
            # pelo Excel nao traz valor calculado nela. O percentual do ciclo
            # e entrada literal e continua disponivel, entao o fator segue
            # sendo um dado da Coleta — apenas nao lido. Reconstruido aqui
            # somente para a narrativa documental do indice acumulado; as
            # cadeias de VTA, retroativo e posicao contratual nao usam este
            # bloco e permanecem inalteradas.
            derivado = fator_derivado.get(c.get("ciclo"))
            if derivado is not None:
                fator = derivado
                origem = "parametros_v10.percentual_do_ciclo (fator recomposto)"
        indice_por_ciclo.append({
            "ciclo": c.get("ciclo"),
            "indice_percentual": c.get("indice_percentual"),
            "fator_acumulado": fator,
            "indice_acumulado": _indice_acumulado(fator),
            "origem": origem,
        })
    ultimo_indice = next(
        (c for c in reversed(indice_por_ciclo) if c.get("fator_acumulado") is not None),
        {},
    )
    return {
        "vta_oficial": {
            "valor": vta.get("valor") or (painel.get("vta") or {}).get("oficial"),
            "origem": "historico/resumo lido; valor preservado para conferencia, nao e VTA",
        },
        "vta_composicao": _montar_vta_composicao(leitura),
        "valor_contrato_atualizado": {
            "valor": valor_contrato.get("valor"),
            "ciclo_referencia": valor_contrato.get("ciclo_referencia"),
            "origem": valor_contrato.get("fonte"),
        },
        "retroativo": {
            "total": retro.get("valor"),
            "estado": retro.get("estado"),
            "com_efeitos_financeiros": round(reconhecido + pago, 2),
            "sem_efeitos_financeiros": round(em_analise, 2),
            "por_estado": copy.deepcopy(por_estado),
            "origens": copy.deepcopy(retro.get("origens") or []),
            "origem": retro.get("fonte"),
        },
        "indice_por_ciclo": indice_por_ciclo,
        "indice_acumulado": {
            "ciclo_referencia": ultimo_indice.get("ciclo"),
            "fator_acumulado": ultimo_indice.get("fator_acumulado"),
            "indice_acumulado": ultimo_indice.get("indice_acumulado"),
            "origem": "parametros_v10.fator_acumulado",
        },
        "vu_por_ciclo": _vu_por_ciclo(leitura),
        "remanescentes_por_ciclo": list(
            (painel.get("posicao_contratual") or {}).get("saldo_por_ciclo") or []
        ),
        "posicao_contratual": copy.deepcopy(
            painel.get("posicao_contratual")
            or leitura.get("posicao_contratual_sombra")
            or {}
        ),
    }


def _montar_pcs(leitura: dict[str, Any], painel: dict[str, Any]) -> list[dict[str, Any]]:
    pcs_painel = (painel.get("situacao_pcs") or {}).get("pcs") or []
    pcs_lidos = (leitura.get("itens_pc_v10") or {}).get("itens") or []
    entrada_xls = leitura.get("masterfile_fiscal_definitivo") or {}
    origem_aba = "ENTRADA_XLS_PCS" if entrada_xls.get("layout_fiscal_definitivo") else "itens_PC"

    por_numero = {
        str(pc.get("numero_pc") or pc.get("item_ou_grupo") or ""): pc
        for pc in pcs_lidos
    }
    resultado: list[dict[str, Any]] = []
    for pc in pcs_painel:
        numero = str(pc.get("numero_pc") or "")
        bruto = por_numero.get(numero) or {}
        valor_retro = _f(pc.get("retroativo"))
        valor_pago = pc.get("valor_pago")
        elegivel_declarado = bruto.get("elegivel_retroativo_pc")
        elegivel_retroativo = (
            bool(elegivel_declarado)
            if "elegivel_retroativo_pc" in bruto
            else _f(valor_pago) > 0.0
        )
        resultado.append({
            "numero": numero or None,
            "data": _data_iso(pc.get("data_pc") or bruto.get("data_pc")),
            "data_ddmmaaaa": _data_br(pc.get("data_pc") or bruto.get("data_pc")),
            "valor": pc.get("valor_pc") if pc.get("valor_pc") is not None else bruto.get("valor_pc"),
            "ciclo_calculado": pc.get("ciclo_temporal"),
            "ciclo_informado": pc.get("ciclo_informado"),
            "situacao": pc.get("natureza_delta"),
            "status_pagamento": bruto.get("status_pagamento_pc") or (
                "PAGO_LEGADO" if _f(valor_pago) > 0.0 else "NAO_CONFIRMADO"
            ),
            "pagamento_definitivo": elegivel_retroativo,
            "elegivel_retroativo_pc": elegivel_retroativo,
            "origem": {
                "aba": origem_aba,
                "linha": bruto.get("linha"),
                "grau_confiabilidade": "medio" if pc.get("alertas") else "alto",
            },
            "se_gerou_retroativo": elegivel_retroativo and valor_retro > 0.0,
            "valor_retroativo": pc.get("retroativo"),
            "valor_devido": pc.get("valor_devido"),
            "efeito_financeiro_pc": (
                pc.get("efeito_financeiro_pc")
                or bruto.get("efeito_financeiro_pc")
            ),
            "valor_pago": valor_pago,
            "valor_pc_historico": pc.get("valor_pc") if pc.get("valor_pc") is not None else bruto.get("valor_pc"),
            "alertas": list(pc.get("alertas") or []),
        })
    return resultado


def _montar_memoria_por_ciclo(
    leitura: dict[str, Any],
    resultados: dict[str, Any],
    pcs: list[dict[str, Any]],
    assistente: dict[str, Any],
) -> dict[str, Any]:
    """Consolida bases e valores atualizados por ciclo, sem somar metodos."""
    por_ciclo = (leitura.get("parametros_v10") or {}).get("por_ciclo") or {}
    ciclos: dict[str, dict[str, Any]] = {}
    for nome in ("C0", "C1", "C2", "C3", "C4"):
        reg = por_ciclo.get(nome) or {}
        ciclos[nome] = {
            "ciclo": nome,
            "data_inicio": _data_iso(reg.get("data_inicio")),
            "data_fim": _data_iso(reg.get("data_fim")),
            "fator_acumulado": _f_none(reg.get("fator_acumulado")) or (1.0 if nome == "C0" else None),
            "retroativo": {
                "financeiro": {"base_original": 0.0, "valor_atualizado": 0.0, "retroativo": 0.0, "evidencias": 0},
                "pc": {"base_original": 0.0, "valor_atualizado": 0.0, "retroativo": 0.0, "evidencias": 0},
                "consumidos": {"base_original": 0.0, "valor_atualizado": 0.0, "retroativo": 0.0, "evidencias": 0},
            },
            "residuais": {"quantidade": 0.0, "valor_original": 0.0, "valor_atualizado": 0.0, "itens": 0, "fotografias": []},
        }

    def acumular(ciclo: str, metodo: str, base: Any, atualizado: Any) -> None:
        if ciclo not in ciclos:
            return
        b = _f(base); a = _f(atualizado)
        alvo = ciclos[ciclo]["retroativo"][metodo]
        alvo["base_original"] = round(alvo["base_original"] + b, 2)
        alvo["valor_atualizado"] = round(alvo["valor_atualizado"] + a, 2)
        alvo["retroativo"] = round(alvo["valor_atualizado"] - alvo["base_original"], 2)
        alvo["evidencias"] += 1

    def fator_pc_na_apuracao(ciclo_alvo: str) -> float | None:
        """Reproduz a regra do XLS: acumula somente ciclos marcados para apuração.

        O fator histórico total de ``parametros!F`` inclui reajustes já concedidos.
        Para um PC pertencente ao objeto atual, porém, ``itens_PC`` aplica apenas os
        percentuais das linhas com COMPUTAR_NESTA_APURACAO=Sim até o ciclo do PC.
        Usar o fator histórico total aqui duplicaria ciclos já formalizados.
        """
        try:
            limite = int(str(ciclo_alvo).upper().removeprefix("C"))
        except (TypeError, ValueError):
            return None
        fator = 1.0
        for numero in range(1, limite + 1):
            reg = por_ciclo.get(f"C{numero}") or {}
            computar = str(reg.get("computar_nesta_apuracao") or "").strip().lower()
            if computar not in {"sim", "s", "true", "1", "yes"}:
                continue
            percentual = _f_none(reg.get("percentual_reajuste"))
            if percentual is not None:
                fator *= 1.0 + percentual
        return round(fator, 12)

    for parcela in (leitura.get("vta_sombra") or {}).get("parcelas_computadas") or []:
        if parcela.get("fonte_parcela") not in {"Financeiro", "Historico financeiro"}:
            continue
        ident = str(parcela.get("identificador") or "")
        if ident.startswith("financeiro:") and ":delta:" in ident:
            # Parcela-delta do financeiro sombra: retroativo ja refletido no
            # valor_atualizado da parcela-base da mesma competencia — somar
            # as duas dobraria o retroativo.
            continue
        ciclo = str(parcela.get("ciclo") or "").upper()
        fator = ciclos.get(ciclo, {}).get("fator_acumulado") or 1.0
        base = parcela.get("valor")
        atualizado = parcela.get("valor_atualizado")
        if atualizado is None:
            atualizado = round(_f(base) * fator, 2)
        acumular(ciclo, "financeiro", base, atualizado)

    pcs_elegiveis = 0
    pcs_nao_elegiveis = 0
    for pc in pcs:
        if not pc.get("elegivel_retroativo_pc"):
            pcs_nao_elegiveis += 1
            continue
        ciclo_pc = str(pc.get("ciclo_calculado") or "").upper()
        efeito_pc = str(pc.get("efeito_financeiro_pc") or "").strip()
        fator_pc = (
            1.0 if efeito_pc == "Nao"
            else fator_pc_na_apuracao(ciclo_pc) if efeito_pc == "Sim"
            else None
        )
        valor_pago_pc = _f_none(pc.get("valor_pago"))
        if fator_pc is None or valor_pago_pc is None or valor_pago_pc <= 0.0:
            pcs_nao_elegiveis += 1
            continue
        pcs_elegiveis += 1
        acumular(
            ciclo_pc,
            "pc",
            valor_pago_pc,
            round(valor_pago_pc * fator_pc, 2),
        )

    # VTA-C2: item["consumos"] e multi-ciclo por item ({"C0": {"qtd","valor"},
    # "C1": {...}, ...}) — nao existe (nem nunca existiu) item["ciclo"] unico
    # no schema real de _ler_itens_consumidos_v10. Itera cada ciclo do item
    # diretamente. Ausencia (qtd None/"") nao gera evidencia; qtd=0 explicito
    # e evidencia valida (base=0, atualizado=0) — teste numerico explicito via
    # _f_none, nunca truthiness.
    # VTA-C2.1 (item 3): uma QTD_CONS_Cn numerica informada exige valoracao
    # suficiente (VU_original e, na ausencia de valor cacheado, fator do
    # ciclo). Se faltar, NAO somar apenas as parcelas validas — sinaliza
    # falha de completude que fecha (fail-closed) a conferencia inteira do
    # metodo Consumido mais abaixo, em vez de produzir um VTA parcial
    # silencioso.
    consumidos_falha_completude = False
    for item in (leitura.get("itens_consumidos_v10") or {}).get("itens") or []:
        if item.get("item") in (None, ""):
            continue
        vu_original = _f_none(item.get("vu_original"))
        consumos = item.get("consumos") or {}
        for ciclo in ("C0", "C1", "C2", "C3", "C4"):
            consumo = consumos.get(ciclo) or {}
            qtd = _f_none(consumo.get("qtd"))
            if qtd is None:
                continue
            if vu_original is None:
                consumidos_falha_completude = True
                continue
            base = round(qtd * vu_original, 2)
            atualizado = _f_none(consumo.get("valor"))
            if atualizado is None:
                fator = ciclos.get(ciclo, {}).get("fator_acumulado")
                if fator is None:
                    consumidos_falha_completude = True
                    continue
                atualizado = round(base * fator, 2)
            acumular(ciclo, "consumidos", base, atualizado)

    fotos_escolhidas: dict[tuple[str, str], dict[str, Any]] = {}
    for foto in (leitura.get("execucao_saldo") or {}).get("fotografias_ciclo") or []:
        chave = (str(foto.get("item") or ""), str(foto.get("ciclo") or "").upper())
        anterior = fotos_escolhidas.get(chave)
        if anterior is None or foto.get("tipo_fotografia") == "CICLO_ATUAL_EM_EXECUCAO":
            fotos_escolhidas[chave] = foto
    for foto in fotos_escolhidas.values():
        ciclo = str(foto.get("ciclo") or "").upper()
        if ciclo not in ciclos:
            continue
        fator = ciclos[ciclo]["fator_acumulado"] or 1.0
        qtd = _f(foto.get("qtd_remanescente"))
        original = _f(foto.get("valor_original"))
        residual = ciclos[ciclo]["residuais"]
        residual["quantidade"] = round(residual["quantidade"] + qtd, 4)
        residual["valor_original"] = round(residual["valor_original"] + original, 2)
        residual["valor_atualizado"] = round(residual["valor_atualizado"] + original * fator, 2)
        residual["itens"] += 1
        residual["fotografias"].append(copy.deepcopy(foto))

    # Novo modelo oficial: sem itens_Execucao_Saldo/historico, a fotografia
    # quantitativa vem de posicao_contratual (100% formulas recalculadas
    # pelo Excel; cache ausente ja bloqueia a formalizacao na politica).
    # Entra apenas quando nenhuma fotografia tradicional existir, e somente
    # para o ciclo vigente (posicao atual do contrato) — sem dado sintetico.
    posicao = leitura.get("posicao_contratual") or {}
    vigente = str((leitura.get("controle") or {}).get("ciclo_vigente") or "").upper()
    tem_fotografia = any(c["residuais"]["itens"] for c in ciclos.values())
    if (not tem_fotografia and posicao.get("ok")
            and not posicao.get("cache_ausente") and vigente in ciclos):
        fator_vigente = ciclos[vigente]["fator_acumulado"]
        historico_por_item = {
            str(item.get("item") or "").strip().upper(): item
            for item in (leitura.get("historico_vu") or {}).get("itens") or []
        }
        for reg in posicao.get("itens") or []:
            vu = _f_none(reg.get("VU_ORIGINAL"))
            qtd = _f_none(reg.get(f"QTD_REM_AJUSTADA_{vigente}"))
            hist = historico_por_item.get(
                str(reg.get("ITEM") or "").strip().upper()
            ) or {}
            vu_vigente = _f_none(
                (hist.get("vu_ciclos") or {}).get(f"VU_{vigente}")
            )
            if (vu is None or qtd is None or vu_vigente is None
                    or fator_vigente is None):
                continue
            original = round(qtd * vu, 2)
            residual = ciclos[vigente]["residuais"]
            residual["quantidade"] = round(residual["quantidade"] + qtd, 4)
            residual["valor_original"] = round(residual["valor_original"] + original, 2)
            residual["valor_atualizado"] = round(
                residual["valor_atualizado"] + round(qtd * vu_vigente, 2), 2
            )
            residual["itens"] += 1
            residual["fotografias"].append({
                "item": reg.get("ITEM"),
                "ciclo": vigente,
                "qtd_remanescente": qtd,
                "valor_original": original,
                "tipo_fotografia": "POSICAO_CONTRATUAL",
                "origem": "posicao_contratual (XLS recalculado pelo Excel)",
                "check": reg.get("CHECK_POSICAO_CONTRATUAL"),
            })

    vu_itens = []
    for item in (leitura.get("itens_contrato") or {}).get("itens") or []:
        vu_original = _f_none(item.get("vu_original"))
        if vu_original is None:
            continue
        vu_ciclos = {
            ciclo: round(vu_original * (ciclos[ciclo]["fator_acumulado"] or 1.0), 6)
            for ciclo in ("C1", "C2", "C3", "C4")
        }
        vu_itens.append({
            "item": item.get("item"), "descricao": item.get("descricao"),
            "quantidade_contratada": item.get("qtd_contratada"),
            "vu_original": vu_original,
            "valor_total_original": item.get("valor_total_original") or round(_f(item.get("qtd_contratada")) * vu_original, 2),
            "vu_ciclos": vu_ciclos,
        })

    # Consumidores documentais (VU por item): sem ITENS_CONTRATO no novo
    # modelo, deriva de posicao_contratual (ITEM, VU_ORIGINAL, QTD_BASE).
    if not vu_itens and posicao.get("ok") and not posicao.get("cache_ausente"):
        historico_por_item = {
            str(item.get("item") or "").strip().upper(): item
            for item in (leitura.get("historico_vu") or {}).get("itens") or []
        }
        for reg in posicao.get("itens") or []:
            vu_original = _f_none(reg.get("VU_ORIGINAL"))
            if vu_original is None:
                continue
            # Etapa 26H: base zero automatica p/ novo item (Nxxx) com base
            # vazia; item normal permanece com vazio (nunca vira 0 aqui).
            from _reajuste_utils import qtd_base_efetiva
            qtd_base = qtd_base_efetiva(
                reg.get("ITEM"), reg.get("QTD_BASE_ORIGINAL")
            )
            qtd_vigente = (
                _f_none(reg.get(f"QTD_CONTRATADA_{vigente}"))
                if vigente in ciclos else None
            )
            hist = historico_por_item.get(
                str(reg.get("ITEM") or "").strip().upper()
            ) or {}
            vu_por_ciclo = {
                c: _f_none((hist.get("vu_ciclos") or {}).get(f"VU_{c}"))
                for c in ("C0", "C1", "C2", "C3", "C4")
            }
            qtd_por_ciclo = {
                c: _f_none(
                    (hist.get("qtd_vigente_ciclos") or {}).get(c)
                )
                for c in ("C0", "C1", "C2", "C3", "C4")
            }
            rem_por_ciclo = {
                c: _f_none(
                    (hist.get("qtd_rem_ajustada_ciclos") or {}).get(c)
                )
                for c in ("C0", "C1", "C2", "C3", "C4")
            }
            totais_por_ciclo = {
                c: (
                    round(qtd_por_ciclo[c] * vu_por_ciclo[c], 2)
                    if qtd_por_ciclo[c] is not None
                    and vu_por_ciclo[c] is not None else None
                )
                for c in ("C0", "C1", "C2", "C3", "C4")
            }
            # 26J.1: ciclo de nascimento (mesma regra da coluna tecnica Y do
            # template: primeiro ciclo com quantidade contratada positiva).
            # Permite distinguir pre-vigencia de dado ausente nos documentos.
            ciclo_nascimento = next(
                (
                    n for n in range(0, 5)
                    if (_f_none(reg.get(f"QTD_CONTRATADA_C{n}")) or 0) > 0
                ),
                None,
            )
            vu_itens.append({
                "item": reg.get("ITEM"),
                "descricao": None,
                "quantidade_contratada": (
                    qtd_vigente if qtd_vigente is not None else qtd_base
                ),
                "vu_original": vu_original,
                "valor_total_original": totais_por_ciclo["C0"],
                "vu_c0": vu_por_ciclo["C0"],
                "vu_ciclos": {
                    c: vu_por_ciclo[c]
                    for c in ("C1", "C2", "C3", "C4")
                },
                "vu_por_ciclo": vu_por_ciclo,
                "quantidade_por_ciclo": qtd_por_ciclo,
                "remanescente_por_ciclo": rem_por_ciclo,
                "total_por_ciclo": totais_por_ciclo,
                "ciclo_nascimento": ciclo_nascimento,
                "origem": "posicao_contratual",
            })

    potencial_info = leitura.get("potencial_futuro") or {}
    potencial_raw = potencial_info.get("valor_atualizado_vigente")
    potencial_disponivel = bool(potencial_info.get("disponivel") and potencial_raw is not None)
    if not potencial_disponivel and vigente in ciclos:
        residual_vigente = ciclos[vigente]["residuais"]
        if int(residual_vigente.get("itens") or 0) > 0:
            potencial_raw = residual_vigente.get("valor_atualizado")
            potencial_disponivel = potencial_raw is not None
    potencial = _f(potencial_raw)
    # VTA-C2: potencial_futuro (leitura["potencial_futuro"]) depende de
    # resumo["saldo_remanescente"], que NAO e alimentado pelo metodo
    # Consumido (so por ENTRADA_XLS_REMANESCENTES/fiscal legado). O metodo
    # Consumido tem fonte propria e independente para o remanescente futuro,
    # calculada diretamente de itens_consumidos_v10 (mesma semantica do
    # MEMORIA_RESULTADOS!C33/D33 apos a correcao do XLS): quantidade
    # contratada menos quantidade total ja consumida, por item, valorada a
    # VU_original e escalada pelo fator acumulado do ciclo vigente.
    # VTA-C2.1: tres desfechos possiveis para o remanescente do Consumido:
    #   - None: nenhum item real do metodo Consumido (evidencias de
    #     execucao tambem serao 0; a conferencia ja fica indisponivel por
    #     falta de execucao, independente deste calculo).
    #   - "FALHA": ha pelo menos um item real com dado obrigatorio ausente
    #     (qtd_contratada/vu_original/qtd_total) OU sinalizado como
    #     divergente pelo proprio CHECK do template (ex.: sobreconsumo,
    #     itens_Consumidos!Q="DIVERGENCIA: CONSUMO MAIOR QUE CONTRATADO")
    #     — fecha (fail-closed) o metodo inteiro; nunca soma so os itens
    #     validos, o que mascararia o problema como saldo zero legitimo.
    #   - (original, atualizado): todos os itens reais completos e sem
    #     divergencia sinalizada — zero real (contrato 100% consumido) e
    #     um resultado valido (MAX(diferenca,0) pode ser 0), nunca vira
    #     None so por a soma dar zero.
    remanescente_consumidos: tuple[float, float] | str | None = None
    fator_vigente_consumidos = (ciclos.get(vigente) or {}).get("fator_acumulado")
    if fator_vigente_consumidos is not None:
        original_consumidos = 0.0
        tem_item_real = False
        falha_remanescente = False
        for item in (leitura.get("itens_consumidos_v10") or {}).get("itens") or []:
            if item.get("item") in (None, ""):
                continue
            tem_item_real = True
            qtd_contratada = _f_none(item.get("qtd_contratada"))
            vu_original = _f_none(item.get("vu_original"))
            qtd_total = _f_none(item.get("qtd_total"))
            if qtd_contratada is None or vu_original is None or qtd_total is None:
                falha_remanescente = True
                continue
            check = str(item.get("check") or "").strip().upper()
            if check and check != "OK":
                falha_remanescente = True
                continue
            diferenca = qtd_contratada - qtd_total
            original_consumidos += max(diferenca, 0.0) * vu_original
        if falha_remanescente:
            remanescente_consumidos = "FALHA"
        elif tem_item_real:
            original_consumidos = round(original_consumidos, 2)
            remanescente_consumidos = (
                original_consumidos,
                round(original_consumidos * fator_vigente_consumidos, 2),
            )
    falha_consumidos_geral = (
        consumidos_falha_completude or remanescente_consumidos == "FALHA"
    )
    conferencias = []
    for metodo in ("financeiro", "pc", "consumidos"):
        executado = round(sum(c["retroativo"][metodo]["valor_atualizado"] for c in ciclos.values()), 2)
        evidencias = sum(c["retroativo"][metodo]["evidencias"] for c in ciclos.values())
        bloqueado_metodo = metodo == "consumidos" and falha_consumidos_geral
        if metodo == "consumidos" and isinstance(remanescente_consumidos, tuple):
            potencial_metodo = remanescente_consumidos[1]
            potencial_disponivel_metodo = True
        elif metodo == "consumidos":
            potencial_metodo = None
            potencial_disponivel_metodo = False
        else:
            potencial_metodo = potencial
            potencial_disponivel_metodo = potencial_disponivel
        disponivel_metodo = evidencias > 0 and not bloqueado_metodo
        conferencias.append({
            "metodo": metodo, "disponivel": disponivel_metodo,
            "executado_atualizado": executado if evidencias and not bloqueado_metodo else None,
            "potencial_restante_atualizado": (
                potencial_metodo if potencial_disponivel_metodo and not bloqueado_metodo else None
            ),
            "valor_total_atualizado": (
                round(executado + potencial_metodo, 2)
                if evidencias and potencial_disponivel_metodo and not bloqueado_metodo else None
            ),
            "natureza": "CONFERENCIA_METODOLOGICA",
            "prioridade": {"financeiro": 1, "pc": 2, "consumidos": 3}[metodo],
            "regra_elegibilidade": (
                "fonte padrao preenchida pelo fiscal"
                if metodo == "financeiro" else (
                    "somente PCs com pagamento definitivo e valor efetivamente pago"
                    if metodo == "pc" else
                    # VTA-C2: sem fonte independente de pagamento no metodo
                    # Consumido — nao afirmar "execucao paga"/"retroativo
                    # reconhecido" (item 12 da tarefa).
                    "quantidade consumida x VU aplicavel ao ciclo, sem "
                    "fonte independente de confirmacao de pagamento"
                )
            ),
        })
    disponiveis = [c for c in conferencias if c["disponivel"]]
    # Regra de negocio: Financeiro e sempre a fonte primaria. PC e Consumidos
    # sao fallbacks excepcionais, nessa ordem, nunca alternativas concorrentes.
    # VTA-C2: excecao minima e explicita — quando o metodo configurado em
    # CONTROLE!B1 e "Itens Consumidos" (leitura["controle"]["modo"]=="d") e
    # ha evidencia Consumido valida, a prioridade historica (financeiro->pc)
    # NAO pode vencer por dado residual em outra aba. Financeiro/PC mantêm a
    # prioridade inalterada em qualquer outro cenario.
    _metodo_configurado = {"principal": "financeiro", "pc": "pc", "d": "consumidos"}.get(
        str((leitura.get("controle") or {}).get("modo") or "")
    )
    if _metodo_configurado == "consumidos" and any(
        c["metodo"] == "consumidos" and c["disponivel"] for c in conferencias
    ):
        metodo_escolhido = "consumidos"
    else:
        metodo_escolhido = next(
            (metodo for metodo in ("financeiro", "pc", "consumidos")
             if any(c["metodo"] == metodo and c["disponivel"] for c in conferencias)),
            None,
        )
    selecao_automatica_unica = len(disponiveis) == 1
    # Etapa 27B: a conferencia "pc" e pago-dependente (soma somente PCs
    # pagos) e nao pode fabricar o VTA do metodo PC. O VTA por PC vem
    # exclusivamente da composicao canonica (estrutural + linha de corte);
    # indisponivel -> fail-closed com motivo tecnico, nunca proxy por
    # pagamento nem zero. Fallbacks de financeiro/consumidos preservados.
    vta = next((
        c for c in conferencias
        if c["metodo"] == metodo_escolhido
        and c["metodo"] != "pc"
        and c["disponivel"]
        and c["valor_total_atualizado"] is not None
    ), None)
    vta_pc_fail_closed = metodo_escolhido == "pc"
    composicao = leitura.get("composicao_vta") or {}
    if (str(composicao.get("metodo") or "").lower() == "pc"
            and composicao.get("disponivel")
            and _f_none(composicao.get("vta_composicao")) is not None):
        # A composicao PC e o espelho do XLS: PCs anteriores ao ciclo vigente
        # mais o remanescente itemizado do mesmo corte. O caminho antigo
        # somava tambem PCs do ciclo vigente ao mesmo remanescente, duplicando
        # a execucao e gerando a divergencia 26C.
        metodo_escolhido = "pc"
        selecao_automatica_unica = True
        vta_pc_fail_closed = False
        saldo_comp = composicao.get("saldo_remanescente") or {}
        vta = {
            "metodo": "pc",
            "disponivel": True,
            "executado_atualizado": composicao.get(
                "total_execucao_atualizada"
            ),
            "potencial_restante_atualizado": saldo_comp.get(
                "valor_atualizado"
            ),
            "valor_total_atualizado": composicao.get("vta_composicao"),
            # VTA-POT-1: decomposicao prudencial exposta aqui para que web, XLS
            # e documentos LEIAM o mesmo numero em vez de recalcular.
            "vta_sem_potencial": composicao.get("vta_sem_potencial"),
            # Apurado (pode ser < 0) x incorporado ao VTA (piso zero).
            "retroativo_potencial_apurado": composicao.get(
                "retroativo_potencial_apurado"
            ),
            "retroativo_potencial_vta": composicao.get("retroativo_potencial_vta"),
            "tem_potencial_apurado": bool(composicao.get("tem_potencial_apurado")),
            "tem_parcela_potencial": bool(composicao.get("tem_parcela_potencial")),
            "potencial_por_ciclo": list(composicao.get("potencial_por_ciclo") or []),
            "prioridade": 2,
            "regra_elegibilidade": (
                "PCs anteriores ao ciclo vigente + remanescente do mesmo corte "
                "+ retroativo potencial dos mesmos PCs (criterio prudencial)"
            ),
        }
    return {
        "ciclos": [ciclos[c] for c in ("C0", "C1", "C2", "C3", "C4")],
        "vu_itens": vu_itens,
        "conferencias_metodologicas": conferencias,
        "vta": ({
            **vta,
            "natureza": (
                "VALOR_TOTAL_ATUALIZADO_COM_RESSALVA"
                if selecao_automatica_unica else "VALOR_TOTAL_ATUALIZADO"
            ),
            "criterio_selecao": (
                "fonte padrao financeiro"
                if metodo_escolhido == "financeiro" else (
                    "composicao canonica PC (estrutural + linha de corte)"
                    if metodo_escolhido == "pc" else
                    "fallback consumidos como execucao paga"
                )
            ),
        } if vta else {
            "valor_total_atualizado": None, "metodo": metodo_escolhido,
            "natureza": "INDETERMINADO",
            "motivo": (
                "VTA pelo metodo PC requer a composicao canonica (PCs "
                "validos anteriores a linha de corte + remanescente do "
                "mesmo corte); sem essa base o VTA nao e fabricado a "
                "partir de PCs pagos (fail-closed)."
                if vta_pc_fail_closed else
                "Potencial restante nao informado; ha apenas execucao atualizada."
                if metodo_escolhido else "Metodologia sem evidencias suficientes."
            )
        }),
        "definicao_vta": "execucao atualizada pelo metodo aplicavel + potencial restante atualizado",
        "hierarquia_retroativo": ["financeiro", "pc_pago_definitivo", "consumidos_pagos"],
        "controle_pcs": {
            "elegiveis_retroativo": pcs_elegiveis,
            "historicos_pendentes_ou_nao_pagos": pcs_nao_elegiveis,
            "regra": "PC historico, pendente ou sem valor final nao gera retroativo",
        },
    }


def _montar_pendencias(
    assistente: dict[str, Any],
    dossie: dict[str, Any],
    modo: dict[str, Any],
    contrato: dict[str, Any],
    leitura: dict[str, Any] | None = None,
) -> dict[str, list[str]]:
    bloqueantes = [
        str(i.get("descricao") or "")
        for i in assistente.get("inconsistencias") or []
        if str(i.get("gravidade") or "").lower() == "bloqueante"
    ]
    advertencias = [
        str(i.get("descricao") or "")
        for i in assistente.get("inconsistencias") or []
        if _norm_texto(i.get("gravidade")) in {"atencao", "atencao"}
    ]
    for item in dossie.get("pendencias_impeditivas") or []:
        texto = str(item or "").strip()
        if texto and texto not in bloqueantes and texto not in advertencias:
            advertencias.append(texto)
    for alerta in modo.get("alertas") or []:
        texto = str(alerta or "").strip()
        if texto and texto not in advertencias:
            advertencias.append(texto)
    ressalvas = [
        str(i.get("descricao") or "")
        for i in assistente.get("inconsistencias") or []
        if str(i.get("gravidade") or "").lower() == "informacao"
    ]
    ressalvas.extend(str(p) for p in dossie.get("pendencias_nao_impeditivas") or [])
    ressalvas.extend((contrato.get("limitacoes") or []))
    if leitura is not None and not _vu_por_ciclo(leitura):
        ressalvas.append("VU por ciclo nao localizado no Objeto Processo; nao foi estimado.")
    return {
        "bloqueantes": _unicos(bloqueantes),
        "advertencias": _unicos(advertencias),
        "ressalvas": _unicos(ressalvas),
    }


def _montar_justificativas(
    assistente: dict[str, Any],
    dossie: dict[str, Any],
    modo: dict[str, Any],
    contrato: dict[str, Any],
    pendencias: dict[str, list[str]],
) -> dict[str, Any]:
    return {
        "metodologia": (assistente.get("metodologia") or {}).get("justificativa"),
        "modo_conducao": modo.get("diagnostico"),
        "limitacoes": list(contrato.get("limitacoes") or []) + list(pendencias.get("ressalvas") or []),
        "excecoes": list(modo.get("alertas") or []),
        "conclusao": dossie.get("motivo"),
    }


def _resumo_fontes_xls(leitura: dict[str, Any]) -> dict[str, Any]:
    entrada = leitura.get("masterfile_fiscal_definitivo") or {}
    return {
        "layout": entrada.get("layout") or "",
        "abas_presentes": list(entrada.get("abas_visiveis_fiscal") or []),
        "linhas_por_aba": copy.deepcopy(entrada.get("linhas_por_aba") or {}),
        "normalizacoes": copy.deepcopy(entrada.get("normalizacoes") or {}),
    }


def _montar_decisao(
    assistente: dict[str, Any],
    dossie: dict[str, Any],
    modo: dict[str, Any],
) -> dict[str, Any]:
    proxima = assistente.get("proxima_acao") or {}
    return {
        "situacao_processo": dossie.get("status"),
        "proxima_acao": proxima.get("acao") or modo.get("proxima_acao"),
        "motivo_proxima_acao": proxima.get("motivo") or dossie.get("motivo"),
        "conclusao_operacional": dossie.get("ato_ou_providencia"),
        "pode_concluir_no_claus_new": bool(dossie.get("pode_concluir_no_claus_new")),
        "precisa_abrir_excel": bool(dossie.get("precisa_abrir_excel")),
        "confirmacao_requerida": bool(dossie.get("confirmacao_requerida")),
        "resumo_executivo": dossie.get("resumo_executivo"),
    }


def _ciclos_parametros(parametros: dict[str, Any]) -> list[dict[str, Any]]:
    ciclos = parametros.get("ciclos")
    if isinstance(ciclos, list) and ciclos:
        return [dict(c) for c in ciclos if isinstance(c, dict)]
    por_ciclo = parametros.get("por_ciclo") or {}
    return [dict(por_ciclo[c]) for c in sorted(por_ciclo) if isinstance(por_ciclo.get(c), dict)]


def _parametro_para_ciclo(reg: dict[str, Any]) -> dict[str, Any]:
    return {
        "ciclo": reg.get("ciclo"),
        "computar_nesta_apuracao": reg.get("computar_nesta_apuracao"),
        "periodo": reg.get("periodo"),
        "data_inicio": _data_iso(reg.get("data_inicio")),
        "data_inicio_ddmmaaaa": _data_br(reg.get("data_inicio")),
        "data_fim": _data_iso(reg.get("data_fim")),
        "data_fim_ddmmaaaa": _data_br(reg.get("data_fim")),
        "percentual_reajuste": reg.get("percentual_reajuste"),
        "fator_acumulado": reg.get("fator_acumulado"),
        "situacao": reg.get("situacao"),
        "origem": {
            "aba": reg.get("origem_aba") or "parametros",
            "linha": reg.get("origem_linha"),
            "grau_confiabilidade": "alto" if reg.get("data_inicio") and reg.get("data_fim") else "medio",
        },
    }


def _evidencia_bloco(bloco: Any, origem: str) -> dict[str, Any]:
    base = copy.deepcopy(bloco) if isinstance(bloco, dict) else {}
    presente = bool(base.get("presente") or base.get("quantidade") or base.get("parcelas") or base.get("itens"))
    base["presente"] = presente
    base["origem"] = origem
    base["grau_confiabilidade"] = "alto" if presente else "insuficiente"
    return base


def _origem_evento(ev: dict[str, Any]) -> dict[str, Any]:
    rastreio = ev.get("rastreabilidade") or {}
    return {
        "sequencia": ev.get("sequencia"),
        "tipo_evento": ev.get("tipo_evento"),
        "identificador": ev.get("identificador"),
        "ciclo": ev.get("ciclo"),
        "valor": ev.get("valor"),
        "origem": ev.get("origem_dado") or rastreio.get("fonte"),
        "fonte_parcela": ev.get("fonte_parcela"),
        "linha": ev.get("linha"),
        "status_consolidacao": ev.get("status_consolidacao"),
        "grau_confiabilidade": _normalizar_confianca(ev.get("grau_confianca") or "medio"),
        "justificativa": ev.get("justificativa"),
    }


def _grau_confiabilidade_global(evidencias: dict[str, Any], painel: dict[str, Any]) -> str:
    semaforo = (painel.get("semaforo") or {}).get("status")
    if semaforo == "critico":
        return "baixo"
    if any((evidencias.get(chave) or {}).get("presente") for chave in ("financeiro", "pcs", "consumidos", "remanescentes")):
        return "alto" if semaforo == "ok" else "medio"
    return "insuficiente"


def _limitacoes_evidencias(evidencias: dict[str, Any], leitura: dict[str, Any]) -> list[str]:
    limitacoes: list[str] = []
    for chave in ("financeiro", "pcs", "consumidos", "remanescentes"):
        bloco = evidencias.get(chave) or {}
        if not bloco.get("presente"):
            limitacoes.append(f"Evidencia de {chave} nao localizada no MasterFile lido.")
    for aviso in leitura.get("avisos") or []:
        texto = str(aviso).strip()
        if texto:
            limitacoes.append(texto)
    return _unicos(limitacoes)


def _vu_por_ciclo(leitura: dict[str, Any]) -> list[dict[str, Any]]:
    resultado = []
    for item in (leitura.get("historico_vu") or {}).get("itens") or []:
        resultado.append({
            "item": item.get("item"),
            "descricao": item.get("descricao"),
            "vu_original": item.get("vu_original"),
            "vu_ciclos": copy.deepcopy(item.get("vu_ciclos") or {}),
            "vu_vigente": item.get("vu_vigente"),
            "fator_acumulado": item.get("fator_acumulado"),
            "origem": {
                "aba": "historico_VU",
                "grau_confiabilidade": "alto" if item.get("vu_ciclos") else "medio",
            },
        })
    return resultado


def _indice_acumulado(fator: Any) -> float | None:
    valor = _f_none(fator)
    if valor is None:
        return None
    return round(valor - 1.0, 10)


def _fatores_acumulados_por_percentual(ciclos: list[dict[str, Any]]) -> dict[str, float]:
    """Recompoe a cadeia de FATOR_ACUMULADO a partir do percentual de cada ciclo.

    Espelha a formula da coluna parametros!F, que encadeia
    ``F{r} = F{r-1} * (1 + E{r})`` e para de produzir valor no primeiro ciclo
    sem percentual numerico. Serve apenas como recomposicao de um dado que ja
    esta na Coleta (o percentual e entrada literal) quando o valor calculado
    da formula nao foi gravado no arquivo.
    """
    derivados: dict[str, float] = {}
    anterior = 1.0
    for c in ciclos:
        ciclo = c.get("ciclo")
        pct = _f_none(c.get("indice_percentual"))
        if pct is None:
            break
        anterior = anterior * (1.0 + pct)
        if ciclo:
            derivados[ciclo] = anterior
    return derivados


def _valor_por_estado(por_estado: dict[str, Any], marcador: str) -> float:
    total = 0.0
    alvo = _norm_texto(marcador)
    for chave, valor in por_estado.items():
        if alvo in _norm_texto(chave):
            total += _f(valor)
    return round(total, 2)


def _data_iso(valor: Any) -> str | None:
    if isinstance(valor, datetime):
        return valor.date().isoformat()
    if isinstance(valor, date):
        return valor.isoformat()
    texto = str(valor or "").strip()
    return texto or None


def _data_br(valor: Any) -> str | None:
    if isinstance(valor, datetime):
        return valor.strftime("%d/%m/%Y")
    if isinstance(valor, date):
        return valor.strftime("%d/%m/%Y")
    texto = str(valor or "").strip()
    if not texto:
        return None
    try:
        parsed = datetime.fromisoformat(texto)
        return parsed.strftime("%d/%m/%Y")
    except ValueError:
        return texto


def _f(valor: Any) -> float:
    try:
        return float(valor or 0.0)
    except (TypeError, ValueError):
        return 0.0


def _f_none(valor: Any) -> float | None:
    if valor in (None, ""):
        return None
    try:
        return float(valor)
    except (TypeError, ValueError):
        return None


def _normalizar_confianca(valor: Any) -> str:
    texto = str(valor or "").strip().lower()
    if "alto" in texto or "alta" in texto:
        return "alto"
    if "baix" in texto:
        return "baixo"
    if "insuf" in texto:
        return "insuficiente"
    return "medio"


def _norm_texto(valor: Any) -> str:
    texto = str(valor or "").strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    return "".join(ch for ch in texto if not unicodedata.combining(ch))


def _unicos(valores: list[str]) -> list[str]:
    vistos: set[str] = set()
    saida: list[str] = []
    for valor in valores:
        texto = str(valor or "").strip()
        if texto and texto not in vistos:
            vistos.add(texto)
            saida.append(texto)
    return saida


def _json_ready(valor: Any) -> Any:
    if isinstance(valor, dict):
        return {str(k): _json_ready(v) for k, v in valor.items()}
    if isinstance(valor, (list, tuple)):
        return [_json_ready(v) for v in valor]
    if isinstance(valor, datetime):
        return valor.isoformat()
    if isinstance(valor, date):
        return valor.isoformat()
    return valor
