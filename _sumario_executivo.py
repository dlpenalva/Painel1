"""Sumario Executivo em PDF (Etapa 5).

Camada de APRESENTACAO somente leitura. Consome exclusivamente o Objeto
Processo de Reajuste (fonte unica canonica) e a memoria de calculo ja
persistida no XLS (Etapa 4). NAO recalcula VTA, retroativo, variacoes nem
efeitos financeiros: apenas exibe valores ja consolidados pelos motores
oficiais. Ausencia de dado vira "Nao informado"; regra inaplicavel vira
"Nao aplicavel"; ausencia nunca vira zero.

Interface tecnica (para conexao aos cards na Etapa 7):

    dados = montar_dados_sumario_executivo(leitura, identificacao=...)
    pdf_bytes = gerar_sumario_executivo_pdf(dados)

`leitura` e o dicionario devolvido por `ler_masterfile_v10` (ou o proprio
objeto do processo ja materializado). `identificacao` transporta metadados
da Calculadora nao presentes no XLS (empresa, contrato, processo, datas de
pedido por ciclo); quando ausentes, os campos saem como "Nao informado".
"""
from __future__ import annotations

import copy
from datetime import date, datetime
from io import BytesIO
from typing import Any

from _objeto_processo_reajuste import (
    montar_objeto_processo_reajuste,
    obter_objeto_processo_reajuste,
)
from _reconciliacao_xls_python import campos_nao_confiaveis_para_documentos

NAO_INFORMADO = "Não informado"
NAO_APLICAVEL = "Não aplicável"

# Etapa 5b: quando ha divergencia relevante XLS x Python, o campo divergente e
# seus dependentes diretos NAO sao exibidos (nem XLS nem Python sao adotados);
# ficam vazios, mas a divergencia permanece sinalizada no sistema.
MOTIVO_DIVERGENCIA_XLS_PYTHON = (
    "Valor não exibido: divergência relevante XLS × Python pendente de "
    "equalização antes da formalização."
)

ROTULO_METODO = {
    "financeiro": "Financeiro",
    "pc": "Pedido de Compras",
    "consumidos": "Consumidos",
}


# ---------------------------------------------------------------------------
# Normalizacao de dados (parte 1: sem PDF)
# ---------------------------------------------------------------------------

def montar_dados_sumario_executivo(
    leitura_ou_objeto: dict[str, Any] | None,
    identificacao: dict[str, Any] | None = None,
    memoria_calculo: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Consolida os dados canonicos do Sumario Executivo, sem renderizar.

    `memoria_calculo` permite injetar a memoria lida do XLS quando o chamador
    trabalha direto com o objeto do processo (que nao a transporta); com a
    leitura completa do leitor v10, ela e extraida automaticamente.
    """
    objeto = obter_objeto_processo_reajuste(leitura_ou_objeto)
    leitura = leitura_ou_objeto if isinstance(leitura_ou_objeto, dict) else {}
    if objeto is None and leitura.get("ok"):
        objeto = montar_objeto_processo_reajuste(leitura)
    if not isinstance(objeto, dict) or not objeto.get("disponivel"):
        return {
            "disponivel": False,
            "motivo": (objeto or {}).get("motivo")
            or "Leitura do MasterFile ausente ou invalida.",
        }

    if memoria_calculo is None:
        memoria_calculo = leitura.get("memoria_calculo")
    if not isinstance(memoria_calculo, dict):
        memoria_calculo = {}

    dados_op = objeto.get("dados_operacionais") or {}
    parametros = dados_op.get("parametros_v10") or {}
    resultados = objeto.get("resultados") or {}
    memoria_ciclo = objeto.get("memoria_por_ciclo") or {}
    metodologia = objeto.get("metodologia") or {}
    pendencias = objeto.get("pendencias") or {}

    # Etapa 5b: campos nao-confiaveis por divergencia relevante XLS x Python.
    # A secao financeira e montada uma unica vez e mascarada antes de alimentar
    # tanto a sintese quanto o topo, para que os 3 documentos herdem o mesmo
    # mascaramento (VTA e retroativo).
    campos_nc = campos_nao_confiaveis_para_documentos(
        leitura.get("reconciliacao_xls_python")
    )
    financeiro_sec = _montar_secao_financeira(resultados, memoria_ciclo)
    _mascarar_financeiro_por_divergencia(financeiro_sec, campos_nc)
    sintese = _montar_sintese(
        objeto, resultados, memoria_ciclo, metodologia, financeiro_sec
    )
    _mascarar_sintese_por_divergencia(sintese, campos_nc)
    # Etapa 26H (politica documental da PREVIA): existindo numero XLS oficial
    # sem resultado definitivo, o VTA e exibido como "R$ x — PREVIA". Nunca
    # declara VALIDADO e nao resolve a divergencia XLS x Python (residual 26C).
    _aplicar_previa_vta(
        sintese,
        leitura.get("resultados_xls"),
        leitura.get("status_resultados"),
    )

    ciclos_sec = _montar_secao_ciclos(parametros, resultados, identificacao)

    return {
        "disponivel": True,
        "identificacao": _montar_identificacao(
            objeto, metodologia, memoria_calculo, identificacao
        ),
        "sintese": sintese,
        "ciclos": ciclos_sec,
        "financeiro": financeiro_sec,
        "itens": _montar_secao_itens(memoria_ciclo, parametros),
        # Etapa 7: historico de VU por ciclo (C0 sempre; ate o ultimo analisado).
        # Fonte canonica unica: leitura["historico_vu"] (aba historico_VU).
        "historico_vu": _montar_secao_historico_vu(
            leitura.get("historico_vu") or {}, ciclos_sec
        ),
        "memoria_calculo": _montar_secao_memoria(memoria_calculo),
        "aditivos": _montar_secao_aditivos(dados_op, parametros),
        "observacoes": _montar_observacoes(objeto, pendencias),
        "campos_nao_confiaveis": sorted(campos_nc),
        # Tres referencias do VTA (auditoria): posicao atual / ultima abertura /
        # integral reajustado. Espelho de RESULTADOS!Tabela 1; nunca oficial.
        "referencias_vta": leitura.get("referencias_vta") or {},
    }


def _ultimo_ciclo_analisado(ciclos: list[dict[str, Any]]) -> int:
    """Indice (0..4) do ultimo ciclo efetivamente analisado nesta apuracao.

    C0 e sempre a base (indice 0). Cada C1..C4 conta como analisado quando
    marcado para computar nesta apuracao (computar == "Sim"). Retorna o maior
    indice analisado; se nenhum reajuste foi computado, retorna 0 (so C0).
    """
    ultimo = 0
    for reg in ciclos or []:
        nome = str(reg.get("ciclo") or "").upper()
        if nome in ("C1", "C2", "C3", "C4") and str(reg.get("computar")) == "Sim":
            ultimo = max(ultimo, int(nome[1]))
    return ultimo


def _montar_secao_historico_vu(
    historico_vu: dict[str, Any], ciclos: list[dict[str, Any]]
) -> dict[str, Any]:
    """Historico de Valores Unitarios por ciclo para Saneador e Apostila.

    Estrutura canonica unica (sem segundo parser): consome leitura["historico_vu"]
    (aba historico_VU, colunas VU_C0..VU_C4). Exibe C0 sempre e os ciclos ate o
    ultimo efetivamente analisado; ciclos futuros fisicamente presentes na
    planilha NAO sao carregados. Nao inventa zeros para celulas sem valor.

    Observacao (Etapa 5): o VU por item e dimensao propria; nenhum campo de
    reconciliacao XLS x Python corresponde a um VU, portanto historico_VU nunca
    e mascarado pelo mapa de dependencia atual.
    """
    ultimo = _ultimo_ciclo_analisado(ciclos)
    colunas = [f"C{i}" for i in range(ultimo + 1)]  # C0..C{ultimo}
    itens_saida: list[dict[str, Any]] = []
    for reg in historico_vu.get("itens") or []:
        vu_ciclos = reg.get("vu_ciclos") or {}
        vus: dict[str, float | None] = {}
        # Etapa 26H: item nascido apos C0 tem VU_C0 legitimamente vazio na aba
        # historico_VU; o fallback p/ vu_original vale apenas para planilhas
        # legadas sem NENHUM VU por ciclo (nunca inventa VU pre-nascimento).
        tem_vu_por_ciclo = any(v is not None for v in vu_ciclos.values())
        for i in range(ultimo + 1):
            valor = vu_ciclos.get(f"VU_C{i}")
            if i == 0 and valor is None and not tem_vu_por_ciclo:
                valor = reg.get("vu_original")  # C0 = VU original (legado)
            vus[f"C{i}"] = _num_ou_none(valor)
        itens_saida.append({
            "item": reg.get("item"),
            "descricao": reg.get("descricao"),
            "vus": vus,
        })
    return {
        "disponivel": bool(itens_saida),
        "ultimo_ciclo": f"C{ultimo}",
        "ciclos": colunas,
        "itens": itens_saida,
    }


def _mascarar_sintese_por_divergencia(sintese: dict[str, Any], campos_nc: set[str]) -> None:
    """Esvazia VTA e retroativo total quando o campo oficial e nao-confiavel.

    VTA_FINAL cobre a cadeia de remanescente (qtd/base/atualizado) via mapa de
    dependencia; RETRO_OFICIAL cobre os retroativos por metodo. Nenhum valor
    (XLS ou Python) e adotado: o campo fica vazio com motivo explicito.
    """
    if "VTA_FINAL" in campos_nc:
        sintese["vta"] = None
        sintese["vta_motivo"] = MOTIVO_DIVERGENCIA_XLS_PYTHON
    if "RETRO_OFICIAL" in campos_nc:
        sintese["retroativo_total"] = None
        sintese["retroativo_estado"] = MOTIVO_DIVERGENCIA_XLS_PYTHON


def _mascarar_financeiro_por_divergencia(financeiro: dict[str, Any], campos_nc: set[str]) -> None:
    """Esvazia os totais de retroativo da secao financeira sob divergencia.

    RETRO_FIN -> delta_total_financeiro; RETRO_PC -> delta_total_pc;
    RETRO_OFICIAL -> retroativo consolidado (e ambos os deltas, pois os
    documentos derivam o retroativo do documento a partir deles). Os valores
    por ciclo (componentes) permanecem, pois nao sao dependentes do total.
    """
    if "RETRO_FIN" in campos_nc or "RETRO_OFICIAL" in campos_nc:
        financeiro["delta_total_financeiro"] = None
    if "RETRO_PC" in campos_nc or "RETRO_OFICIAL" in campos_nc:
        financeiro["delta_total_pc"] = None
    if "RETRO_OFICIAL" in campos_nc:
        financeiro["retroativo_total"] = None
        financeiro["retroativo_estado"] = MOTIVO_DIVERGENCIA_XLS_PYTHON


def _aplicar_previa_vta(
    sintese: dict[str, Any],
    resultados_xls: dict[str, Any] | None,
    status_resultados: dict[str, Any] | None = None,
) -> None:
    """Etapa 26H: PREVIA documental do VTA a partir do numero XLS oficial.

    Aplica-se somente quando o VTA oficial nao pode ser exibido (mascarado por
    divergencia ou indisponivel) e o XLS possui VTA_FINAL numerico. O valor e
    apresentado como PREVIA — nunca como VALIDADO — e a divergencia permanece
    registrada (residual 26C intocado).
    """
    status = str((status_resultados or {}).get("geral") or "").strip().upper()
    if sintese.get("vta") is not None:
        # Igualdade XLS x Python torna o numero calculavel, mas nao transforma
        # um processo ainda marcado como REVISE/ESTIMADO em resultado
        # definitivo. Nessa situacao, preserva o numero e muda somente o selo
        # documental para PREVIA.
        if status in {"REVISE", "ESTIMADO"}:
            sintese["vta_previa"] = round(float(sintese["vta"]), 2)
            sintese["vta"] = None
        return
    valores = (resultados_xls or {}).get("valores") or {}
    vta_xls = valores.get("VTA_FINAL")
    if isinstance(vta_xls, (int, float)) and not isinstance(vta_xls, bool):
        sintese["vta_previa"] = round(float(vta_xls), 2)


def _montar_identificacao(
    objeto: dict[str, Any],
    metodologia: dict[str, Any],
    memoria_calculo: dict[str, Any],
    identificacao: dict[str, Any] | None,
) -> dict[str, Any]:
    externo = identificacao or {}
    controle = (objeto.get("dados_operacionais") or {}).get("controle") or {}
    # Etapa 26H: o indice contratual canonico e CONTROLE!B7 (propagado pelo
    # leitor em controle.indice); METODO_FONTE da memoria segue como fallback.
    indice = (
        str(controle.get("indice") or "").strip()
        or _indice_contratual(metodologia, memoria_calculo)
        or externo.get("indice")
    )
    metodo = metodologia.get("escolhida") or externo.get("metodo")
    # Campos empresa, contrato, processo, versao_masterfile e objeto_processo_id
    # nao sao transportados para o PDF (requisito de privacidade Etapa 5 v2).
    return {
        "indice": _texto_ou_nao_informado(indice),
        "metodo": _texto_ou_nao_informado(metodo),
        "ciclo_vigente": _texto_ou_nao_informado(controle.get("ciclo_vigente")),
        "data_corte": _fmt_data(controle.get("data_corte")),
        "gerado_em": datetime.now().strftime("%d/%m/%Y"),
    }


def _indice_contratual(
    metodologia: dict[str, Any], memoria_calculo: dict[str, Any]
) -> str | None:
    """Indice contratual canonico: METODO_FONTE gravado na memoria (Etapa 4)."""
    fontes: list[str] = []
    for registros in memoria_calculo.values():
        for reg in registros or []:
            if str(reg.get("tipo") or "").upper() == "RESULTADO":
                fonte = str(reg.get("metodo_fonte") or "").strip()
                if fonte and fonte not in fontes:
                    fontes.append(fonte)
    if fontes:
        return "; ".join(fontes)
    return None


def _montar_sintese(
    objeto: dict[str, Any],
    resultados: dict[str, Any],
    memoria_ciclo: dict[str, Any],
    metodologia: dict[str, Any],
    financeiro: dict[str, Any],
) -> dict[str, Any]:
    vta = memoria_ciclo.get("vta") or {}
    # Mesma regra de exibicao da secao financeira: sem evidencia e sem
    # retroativo consolidado, ausencia nao vira zero.
    retro = {
        "total": financeiro.get("retroativo_total"),
        "estado": financeiro.get("retroativo_estado"),
    }
    acumulado = resultados.get("indice_acumulado") or {}
    decisao = objeto.get("decisao") or {}
    if str(acumulado.get("ciclo_referencia") or "").upper() == "C0":
        # C0 e a base contratual, nao um ciclo de reajuste: sem ciclo
        # computado com fator, nao ha variacao acumulada a exibir.
        acumulado = {}
    return {
        "metodo_vta": ROTULO_METODO.get(str(vta.get("metodo") or ""), vta.get("metodo")),
        "vta": vta.get("valor_total_atualizado"),
        "vta_execucao_atualizada": vta.get("executado_atualizado"),
        "vta_saldo_remanescente_atualizado": vta.get(
            "potencial_restante_atualizado"
        ),
        "vta_natureza": vta.get("natureza"),
        "vta_motivo": vta.get("motivo"),
        "variacao_acumulada": acumulado.get("indice_acumulado"),
        "ciclo_referencia_acumulado": acumulado.get("ciclo_referencia"),
        "retroativo_total": retro.get("total"),
        "retroativo_estado": retro.get("estado"),
        "situacao_processo": decisao.get("situacao_processo"),
        "resumo_executivo": decisao.get("resumo_executivo"),
        "justificativa_metodologia": metodologia.get("justificativa"),
    }


def _montar_secao_ciclos(
    parametros: dict[str, Any],
    resultados: dict[str, Any],
    identificacao: dict[str, Any] | None,
) -> list[dict[str, Any]]:
    """Requisitos 2-6: ciclos, periodos, pedidos, meses sem efeito, variacao."""
    pedidos = ((identificacao or {}).get("datas_pedido") or {})
    por_ciclo = parametros.get("por_ciclo") or {}
    indice_por_ciclo = {
        str(reg.get("ciclo") or ""): reg
        for reg in resultados.get("indice_por_ciclo") or []
    }
    saida: list[dict[str, Any]] = []
    for nome in ("C0", "C1", "C2", "C3", "C4"):
        reg = por_ciclo.get(nome)
        if not isinstance(reg, dict):
            continue
        indice = indice_por_ciclo.get(nome) or {}
        saida.append({
            "ciclo": nome,
            "eh_base": nome == "C0",
            "computar": _sim_nao(reg.get("computar_nesta_apuracao")),
            "data_inicio": _fmt_data(reg.get("data_inicio")),
            "data_fim": _fmt_data(reg.get("data_fim")),
            "data_pedido": _fmt_data(pedidos.get(nome)),
            # Apresentacao documental: inicio real do efeito financeiro do ciclo
            # (INICIO_EFEITO_FINANCEIRO). Nunca cai para data_inicio por conveniencia.
            "inicio_efeito_financeiro": _fmt_data(
                reg.get("inicio_efeito_financeiro")
                or reg.get("inicio_efeito_financeiro_parametros")
            ),
            "situacao": _texto_ou_nao_informado(reg.get("situacao")),
            "percentual_reajuste": (
                NAO_APLICAVEL if nome == "C0"
                else _num_ou_none(reg.get("percentual_reajuste"))
            ),
            "fator_acumulado": _num_ou_none(reg.get("fator_acumulado")),
            "indice_acumulado": indice.get("indice_acumulado"),
            "meses_sem_efeito": _meses_sem_efeito(nome, reg),
        })
    return saida


def _meses_sem_efeito(nome: str, reg: dict[str, Any]) -> dict[str, Any]:
    """Requisito 5: competencias do ciclo anteriores ao inicio do efeito.

    Espelha a regra ja homologada da Calculadora (01_Calculo_Simples):
    competencias entre o mes de inicio do ciclo e o mes anterior ao
    INICIO_EFEITO_FINANCEIRO ficam sem efeito financeiro. Nada e estimado:
    sem inicio de efeito registrado, o quadro sai como nao informado.
    """
    if nome == "C0":
        return {"status": NAO_APLICAVEL, "quantidade": None, "competencias": []}
    inicio_ciclo = _como_date(reg.get("data_inicio"))
    inicio_efeito = _como_date(
        reg.get("inicio_efeito_financeiro")
        or reg.get("inicio_efeito_financeiro_parametros")
    )
    if inicio_ciclo is None or inicio_efeito is None:
        return {"status": NAO_INFORMADO, "quantidade": None, "competencias": []}
    competencias: list[str] = []
    ano, mes = inicio_ciclo.year, inicio_ciclo.month
    while (ano, mes) < (inicio_efeito.year, inicio_efeito.month):
        competencias.append(f"{mes:02d}/{ano}")
        mes += 1
        if mes > 12:
            mes, ano = 1, ano + 1
    return {
        "status": "ok",
        "quantidade": len(competencias),
        "competencias": competencias,
    }


def _montar_secao_financeira(
    resultados: dict[str, Any], memoria_ciclo: dict[str, Any]
) -> dict[str, Any]:
    """Requisitos 7-10: pago, delta, reconhecido e em analise, por ciclo."""
    retro = resultados.get("retroativo") or {}
    linhas_financeiro: list[dict[str, Any]] = []
    linhas_pc: list[dict[str, Any]] = []
    for ciclo in memoria_ciclo.get("ciclos") or []:
        nome = ciclo.get("ciclo")
        blocos = ciclo.get("retroativo") or {}
        fin = blocos.get("financeiro") or {}
        pc = blocos.get("pc") or {}
        if int(fin.get("evidencias") or 0) > 0:
            linhas_financeiro.append({
                "ciclo": nome,
                "valor_pago": fin.get("base_original"),
                "valor_atualizado": fin.get("valor_atualizado"),
                "delta": fin.get("retroativo"),
                "evidencias": fin.get("evidencias"),
            })
        if int(pc.get("evidencias") or 0) > 0:
            linhas_pc.append({
                "ciclo": nome,
                "valor_pago": pc.get("base_original"),
                "valor_atualizado": pc.get("valor_atualizado"),
                "delta": pc.get("retroativo"),
                "evidencias": pc.get("evidencias"),
            })
    # Sem evidencia do metodo Financeiro e sem retroativo consolidado, o
    # quadro por estado sairia como zeros do assistente para um processo em
    # que nada foi apurado — ausencia nao vira zero.
    sem_retro_consolidado = (
        not linhas_financeiro and _f(retro.get("total")) == 0.0
    )
    if sem_retro_consolidado:
        retro = {}
    return {
        "financeiro_por_ciclo": linhas_financeiro,
        "pc_por_ciclo": linhas_pc,
        "delta_total_financeiro": (
            round(sum(_f(linha["delta"]) for linha in linhas_financeiro), 2)
            if linhas_financeiro else None
        ),
        "delta_total_pc": (
            round(sum(_f(linha["delta"]) for linha in linhas_pc), 2)
            if linhas_pc else None
        ),
        "retroativo_reconhecido_pago": retro.get("com_efeitos_financeiros"),
        "retroativo_em_analise": retro.get("sem_efeitos_financeiros"),
        "retroativo_por_estado": copy.deepcopy(retro.get("por_estado") or {}),
        "retroativo_total": retro.get("total"),
        "retroativo_estado": retro.get("estado"),
        "hierarquia": list(memoria_ciclo.get("hierarquia_retroativo") or []),
        "controle_pcs": copy.deepcopy(memoria_ciclo.get("controle_pcs") or {}),
    }


def _montar_secao_itens(
    memoria_ciclo: dict[str, Any], parametros: dict[str, Any]
) -> list[dict[str, Any]]:
    """Requisito 11: item, VU/total em C0 e VU/total reajustado por ciclo.

    O VU por ciclo vem pronto de memoria_por_ciclo.vu_itens (objeto do
    processo); o total por ciclo apenas aplica a quantidade canonica ao VU
    canonico, sem novo fator. Somente ciclos com fator canonico registrado em
    parametros sao exibidos (o objeto emite VU=original para ciclos sem
    fator, o que nao representa reajuste concedido).
    """
    por_ciclo = parametros.get("por_ciclo") or {}
    ciclos_com_fator = {
        nome for nome in ("C1", "C2", "C3", "C4")
        if _num_ou_none((por_ciclo.get(nome) or {}).get("fator_acumulado"))
        is not None
    }
    saida: list[dict[str, Any]] = []
    for item in memoria_ciclo.get("vu_itens") or []:
        # A sanitizacao de privacidade do objeto renomeia
        # quantidade_contratada para quantidade_base.
        qtd = _num_ou_none(
            item.get("quantidade_contratada")
            if item.get("quantidade_contratada") is not None
            else item.get("quantidade_base")
        )
        ciclos_exibidos = ["C0"] + sorted(ciclos_com_fator)
        vu_por_ciclo_raw = item.get("vu_por_ciclo") or {
            "C0": item.get("vu_c0", item.get("vu_original")),
            **(item.get("vu_ciclos") or {}),
        }
        qtd_por_ciclo_raw = item.get("quantidade_por_ciclo") or {
            ciclo: qtd for ciclo in ciclos_exibidos
        }
        total_por_ciclo_raw = item.get("total_por_ciclo") or {}
        vu_por_ciclo = {
            ciclo: _num_ou_none(vu_por_ciclo_raw.get(ciclo))
            for ciclo in ciclos_exibidos
        }
        qtd_por_ciclo = {
            ciclo: _num_ou_none(qtd_por_ciclo_raw.get(ciclo))
            for ciclo in ciclos_exibidos
        }
        totais = {}
        for ciclo in ciclos_exibidos:
            total = _num_ou_none(total_por_ciclo_raw.get(ciclo))
            if total is None:
                q = qtd_por_ciclo[ciclo]
                vu = vu_por_ciclo[ciclo]
                total = (
                    round(q * vu, 2)
                    if q is not None and vu is not None else None
                )
            totais[ciclo] = total
        ciclo_nascimento = item.get("ciclo_nascimento")
        saida.append({
            "item": item.get("item"),
            "descricao": item.get("descricao"),
            "quantidade": qtd,
            "ciclo_nascimento": (
                int(ciclo_nascimento) if ciclo_nascimento is not None else None
            ),
            "ciclos": ciclos_exibidos,
            "quantidade_ciclos": qtd_por_ciclo,
            "vu_c0": vu_por_ciclo.get("C0"),
            "total_c0": totais.get("C0"),
            "vu_ciclos": {
                c: vu_por_ciclo[c] for c in ciclos_exibidos if c != "C0"
            },
            "total_ciclos": totais,
        })
    return saida


def _montar_secao_memoria(memoria_calculo: dict[str, Any]) -> list[dict[str, Any]]:
    """Requisito 12: memoria homologada da Etapa 4, agrupada por ciclo."""
    saida: list[dict[str, Any]] = []
    for nome in ("C0", "C1", "C2", "C3", "C4"):
        registros = memoria_calculo.get(nome)
        if not registros:
            continue
        linhas = []
        for reg in registros:
            linhas.append({
                "tipo": str(reg.get("tipo") or "").upper(),
                "ordem": reg.get("ordem"),
                "competencia": _fmt_competencia(reg.get("competencia")),
                "valor_indice": _num_ou_none(reg.get("valor_indice")),
                "fator_mensal": _num_ou_none(reg.get("fator_mensal")),
                "fator_acumulado": _num_ou_none(reg.get("fator_acumulado")),
                "variacao_final": _num_ou_none(reg.get("variacao_final")),
                "metodo_fonte": reg.get("metodo_fonte"),
            })
        saida.append({"ciclo": nome, "registros": linhas})
    return saida


def _montar_secao_aditivos(
    dados_op: dict[str, Any], parametros: dict[str, Any]
) -> dict[str, Any]:
    """Requisito 13: aditivos computaveis ja materializados pelo leitor.

    Regra auditada (leitor v10): so viram parcela os aditivos com
    K=CONSIDERADO NO CALCULO FINANCEIRO=Sim e valor atualizado J em cache do
    Excel — o fator aplicado tem fonte canonica no proprio XLS. A superficie
    documental usa rotulo humano derivado de instrumento ou tipo/ciclo/item e
    preserva o identificador tecnico apenas na estrutura interna.
    """
    vta_sombra = dados_op.get("vta_sombra") or {}
    parcelas = list(vta_sombra.get("parcelas_computadas") or [])
    # Supressoes chegam como parcelas negativas e o motor sombra as mantem na
    # lista nao computada por seguranca financeira. Para a superficie
    # documental, contudo, continuam sendo alteracoes contratuais reais e
    # devem aparecer com sinal, sem serem somadas novamente ao VTA.
    parcelas.extend(vta_sombra.get("parcelas_nao_computadas") or [])
    aditivos: list[dict[str, Any]] = []
    for parcela in parcelas:
        if str(parcela.get("fonte_parcela") or "") != "Aditivo":
            continue
        ciclo = str(parcela.get("ciclo") or "").upper()
        tipo = str(parcela.get("tipo_alteracao") or "").strip()
        tipo_norm = tipo.lower()
        if "supr" in tipo_norm:
            tipo = "Supressão"
        elif "acresc" in tipo_norm or "acrésc" in tipo_norm:
            tipo = "Acréscimo"
        else:
            tipo = tipo or "Alteração contratual"
        item = str(parcela.get("item") or "").strip()
        instrumento = str(parcela.get("instrumento") or "").strip()
        if instrumento:
            rotulo = instrumento
        elif ciclo and item:
            rotulo = f"{ciclo} — {tipo} — Item {item}"
        elif ciclo:
            rotulo = (
                f"Alteração contratual do {ciclo} — registro "
                f"{parcela.get('linha') or ''}"
            ).strip()
        else:
            rotulo = f"{tipo} — Item {item}" if item else tipo
        aditivos.append({
            "identificador_interno": parcela.get("identificador"),
            "rotulo_documental": rotulo,
            "instrumento": instrumento or None,
            "item": item or None,
            "tipo_alteracao": tipo,
            "ciclo": ciclo or NAO_INFORMADO,
            "linha": parcela.get("linha"),
            "data_alteracao": _fmt_data(parcela.get("data_aditivo")),
            "quantidade": _num_ou_none(parcela.get("quantidade")),
            "valor_original": _num_ou_none(parcela.get("valor_original")),
            "valor_atualizado": _num_ou_none(parcela.get("valor")),
            "justificativa": parcela.get("justificativa_vta"),
        })
    return {
        "itens": aditivos,
        "regra": (
            "Somente aditivos considerados no calculo financeiro (K=Sim) com "
            "valor atualizado canonico do Excel sao exibidos; nenhum fator "
            "novo e aplicado e nenhuma data ausente e presumida."
        ),
    }


def _montar_observacoes(
    objeto: dict[str, Any], pendencias: dict[str, Any]
) -> dict[str, Any]:
    evidencias = objeto.get("evidencias") or {}
    return {
        "bloqueantes": list(pendencias.get("bloqueantes") or []),
        "advertencias": list(pendencias.get("advertencias") or []),
        "ressalvas": list(pendencias.get("ressalvas") or []),
        "limitacoes_evidencias": list(evidencias.get("limitacoes") or []),
        "grau_confiabilidade": evidencias.get("grau_confiabilidade"),
    }


# ---------------------------------------------------------------------------
# Formatadores de exibicao
# ---------------------------------------------------------------------------

def _f(valor: Any) -> float:
    try:
        return float(valor or 0.0)
    except (TypeError, ValueError):
        return 0.0


def _num_ou_none(valor: Any) -> float | None:
    if valor in (None, "") or isinstance(valor, bool):
        return None
    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return None
    if numero != numero:  # NaN
        return None
    return numero


def _texto_ou_nao_informado(valor: Any) -> str:
    # Sanitiza emoji/pictograma: o Sumario Executivo e um ARQUIVO entregue.
    from _sanitizacao_documental import remover_emojis_leve
    texto = remover_emojis_leve(str(valor or "")).strip()
    return texto or NAO_INFORMADO


def _sim_nao(valor: Any) -> str:
    texto = str(valor or "").strip().lower()
    if texto in {"sim", "s", "true", "1", "yes"}:
        return "Sim"
    if texto in {"nao", "não", "n", "false", "0", "no"}:
        return "Não"
    return NAO_INFORMADO


def _como_date(valor: Any) -> date | None:
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    texto = str(valor or "").strip()
    if not texto:
        return None
    for formato in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto[:10], formato).date()
        except ValueError:
            continue
    return None


def _fmt_data(valor: Any) -> str:
    parsed = _como_date(valor)
    if parsed is None:
        return NAO_INFORMADO
    return parsed.strftime("%d/%m/%Y")


def _fmt_competencia(valor: Any) -> str:
    parsed = _como_date(valor)
    if parsed is None:
        return NAO_INFORMADO
    return parsed.strftime("%m/%Y")


def formatar_moeda(valor: Any) -> str:
    numero = _num_ou_none(valor)
    if numero is None:
        return NAO_INFORMADO
    texto = f"{numero:,.2f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")
    return f"R$ {texto}"


def formatar_percentual(valor: Any, casas: int = 4) -> str:
    """Percentual a partir do decimal canonico (0.0745 -> '7,4500%')."""
    numero = _num_ou_none(valor)
    if numero is None:
        return NAO_INFORMADO
    texto = f"{numero * 100:.{casas}f}".replace(".", ",")
    return f"{texto}%"


def formatar_fator(valor: Any, casas: int = 6) -> str:
    numero = _num_ou_none(valor)
    if numero is None:
        return NAO_INFORMADO
    return f"{numero:.{casas}f}".replace(".", ",")


def formatar_numero(valor: Any, casas: int = 2) -> str:
    numero = _num_ou_none(valor)
    if numero is None:
        return NAO_INFORMADO
    return f"{numero:,.{casas}f}".replace(",", "\x00").replace(".", ",").replace("\x00", ".")


# ---------------------------------------------------------------------------
# Renderizacao PDF (parte 2)
# ---------------------------------------------------------------------------

_COR_TEXTO = "#18324A"
_COR_TEXTO_SUAVE = "#52697D"
_COR_AZUL = "#1F5F8B"
_COR_BORDA = "#CBD8E2"
_COR_FUNDO_ALT = "#F8FAFC"


def gerar_sumario_executivo_pdf(dados: dict[str, Any]) -> bytes:
    """Renderiza o Sumario Executivo em PDF pesquisavel (A4)."""
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.platypus import (
        SimpleDocTemplate, Spacer,
    )

    estilos = _estilos_pdf()
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=20 * mm,
        title="Sumário Executivo — Reajuste Contratual",
        author="Cl8us",
    )
    historia: list[Any] = []
    if not dados.get("disponivel"):
        historia.append(_paragrafo("Sumário Executivo", estilos["titulo"]))
        historia.append(Spacer(1, 8))
        historia.append(_paragrafo(
            f"Documento não gerado: {dados.get('motivo') or NAO_INFORMADO}",
            estilos["normal"],
        ))
        rodape = _rodape_factory(dados)
        doc.build(historia, onFirstPage=rodape, onLaterPages=rodape)
        return buffer.getvalue()

    _bloco_identificacao(historia, dados, estilos)
    _bloco_sintese(historia, dados, estilos)
    _bloco_referencias_vta_pdf(historia, dados, estilos)
    _bloco_ciclos(historia, dados, estilos)
    _bloco_financeiro(historia, dados, estilos)
    _bloco_itens(historia, dados, estilos)
    _bloco_memoria(historia, dados, estilos)
    _bloco_aditivos(historia, dados, estilos)
    # Secao "8. Observacoes de consistencia" removida (requisito Etapa 5 v2).

    rodape = _rodape_factory(dados)
    doc.build(historia, onFirstPage=rodape, onLaterPages=rodape)
    return buffer.getvalue()


def _estilos_pdf() -> dict[str, Any]:
    from reportlab.lib.enums import TA_LEFT, TA_RIGHT
    from reportlab.lib.styles import ParagraphStyle

    base = dict(fontName="Helvetica", fontSize=8.5, leading=11.5, textColor=_COR_TEXTO)
    return {
        "titulo": ParagraphStyle("titulo", fontName="Helvetica-Bold", fontSize=16,
                                 leading=20, textColor=_COR_AZUL),
        "subtitulo": ParagraphStyle("subtitulo", fontName="Helvetica", fontSize=9.5,
                                    leading=12, textColor=_COR_TEXTO_SUAVE),
        "secao": ParagraphStyle("secao", fontName="Helvetica-Bold", fontSize=11.5,
                                leading=15, textColor=_COR_AZUL, spaceBefore=10,
                                spaceAfter=4, keepWithNext=1),
        "subsecao": ParagraphStyle("subsecao", fontName="Helvetica-Bold", fontSize=9.5,
                                   leading=12.5, textColor=_COR_TEXTO, spaceBefore=8,
                                   spaceAfter=2, keepWithNext=1),
        "normal": ParagraphStyle("normal", **base),
        "celula": ParagraphStyle("celula", **base),
        "celula_dir": ParagraphStyle("celula_dir", alignment=TA_RIGHT, **base),
        "cabecalho_tabela": ParagraphStyle(
            "cabecalho_tabela", fontName="Helvetica-Bold", fontSize=8.5,
            leading=11, textColor="#FFFFFF", alignment=TA_LEFT,
        ),
        "cabecalho_tabela_dir": ParagraphStyle(
            "cabecalho_tabela_dir", fontName="Helvetica-Bold", fontSize=8.5,
            leading=11, textColor="#FFFFFF", alignment=TA_RIGHT,
        ),
        "nota": ParagraphStyle("nota", fontName="Helvetica-Oblique", fontSize=8,
                               leading=10.5, textColor=_COR_TEXTO_SUAVE,
                               spaceBefore=3),
    }


def _rodape_factory(dados: dict[str, Any]):
    from reportlab.lib import colors
    from reportlab.lib.units import mm

    gerado_em = (dados.get("identificacao") or {}).get("gerado_em") \
        or datetime.now().strftime("%d/%m/%Y")
    # Garante formato dd/mm/aaaa (sem hora, sem texto adicional).
    data_rodape = gerado_em[:10] if len(gerado_em) >= 10 else gerado_em

    def _rodape(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor(_COR_BORDA))
        canvas.setLineWidth(0.5)
        canvas.line(
            doc.leftMargin, 13 * mm,
            doc.pagesize[0] - doc.rightMargin, 13 * mm,
        )
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(colors.HexColor(_COR_TEXTO_SUAVE))
        canvas.drawString(
            doc.leftMargin, 9 * mm,
            f"Gerado em {data_rodape}.",
        )
        canvas.restoreState()

    return _rodape


def _tabela(
    linhas: list[list[Any]],
    larguras: list[float],
    estilos: dict[str, Any],
    alinhamentos_direita: set[int] | None = None,
    *,
    permitir_quebra: bool = True,
    alturas: list[float] | None = None,
) -> Any:
    """LongTable com cabecalho repetido nas quebras de pagina."""
    from reportlab.lib import colors
    from reportlab.platypus import LongTable, Paragraph, Table, TableStyle

    direita = alinhamentos_direita or set()
    corpo: list[list[Any]] = []
    for i, linha in enumerate(linhas):
        conv: list[Any] = []
        for j, celula in enumerate(linha):
            if i == 0:
                estilo = estilos[
                    "cabecalho_tabela_dir" if j in direita else "cabecalho_tabela"
                ]
            else:
                estilo = estilos["celula_dir" if j in direita else "celula"]
            conv.append(Paragraph(_escapar(celula), estilo))
        corpo.append(conv)

    classe_tabela = LongTable if permitir_quebra else Table
    tabela = classe_tabela(
        corpo,
        colWidths=larguras,
        rowHeights=alturas,
        repeatRows=1 if permitir_quebra else 0,
        splitByRow=1 if permitir_quebra else 0,
    )
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(_COR_AZUL)),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor(_COR_FUNDO_ALT)]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor(_COR_BORDA)),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 2.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
    ]))
    return tabela


def _largura_util() -> float:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    return A4[0] - 36 * mm


def _bloco_identificacao(historia, dados, estilos) -> None:
    from reportlab.platypus import Spacer

    ident = dados.get("identificacao") or {}
    historia.append(_paragrafo(
        "Sumário Executivo — Reajuste Contratual", estilos["titulo"],
    ))
    historia.append(_paragrafo(
        "Documento executivo consolidado a partir do objeto canônico do "
        "processo e do XLS processado. Nenhum valor é recalculado.",
        estilos["subtitulo"],
    ))
    historia.append(Spacer(1, 6))
    historia.append(_paragrafo("1. Identificação", estilos["secao"]))
    largura = _largura_util()
    # Etapa 26H.1: cabecalho da identificacao trocado de "Campo | Valor" para
    # "Informação | Resultado" por requisito do Gate.
    linhas = [
        ["Informação", "Resultado", "Informação", "Resultado"],
        ["Índice contratual", ident.get("indice"),
         "Método da apuração", ident.get("metodo")],
        ["Ciclo vigente", ident.get("ciclo_vigente"),
         "Data de corte", ident.get("data_corte")],
    ]
    historia.append(_tabela(
        linhas,
        [largura * 0.20, largura * 0.30, largura * 0.20, largura * 0.30],
        estilos,
    ))


def _bloco_sintese(historia, dados, estilos) -> None:
    sintese = dados.get("sintese") or {}
    historia.append(_paragrafo("2. Síntese da apuração", estilos["secao"]))
    largura = _largura_util()
    vta_txt = formatar_moeda(sintese.get("vta"))
    if sintese.get("vta") is None:
        if sintese.get("vta_previa") is not None:
            # Etapa 26H: numero XLS oficial sem resultado definitivo -> PREVIA.
            vta_txt = f"{formatar_moeda(sintese.get('vta_previa'))} — PRÉVIA"
        elif sintese.get("vta_motivo"):
            vta_txt = f"{NAO_INFORMADO} — {sintese.get('vta_motivo')}"
    variacao_txt = formatar_percentual(sintese.get("variacao_acumulada"))
    if sintese.get("ciclo_referencia_acumulado"):
        variacao_txt += f" (referência {sintese.get('ciclo_referencia_acumulado')})"
    # Etapa 26H: as linhas de estado do retroativo e situacao do processo e a
    # conclusao (resumo executivo) foram retiradas do PDF por requisito.
    linhas = [
        ["Indicador", "Valor"],
        ["Método aplicável", _texto_ou_nao_informado(sintese.get("metodo_vta"))],
        ["Valor total atualizado (VTA)", vta_txt],
        ["Variação acumulada", variacao_txt],
        ["Retroativo total", formatar_moeda(sintese.get("retroativo_total"))],
    ]
    historia.append(_tabela(linhas, [largura * 0.34, largura * 0.66], estilos))


def _bloco_ciclos(historia, dados, estilos) -> None:
    ciclos = dados.get("ciclos") or []
    historia.append(_paragrafo("3. Ciclos e efeitos financeiros", estilos["secao"]))
    if not ciclos:
        historia.append(_paragrafo(
            f"Ciclos: {NAO_INFORMADO} — parâmetros C0-C4 não localizados.",
            estilos["normal"],
        ))
        return
    largura = _largura_util()
    linhas: list[list[Any]] = [[
        "Ciclo", "Período", "Pedido de reajuste", "Situação",
        "Variação do ciclo", "Fator acumulado",
    ]]
    for c in ciclos:
        rotulo = f"{c['ciclo']} (base)" if c.get("eh_base") else c["ciclo"]
        variacao = (
            NAO_APLICAVEL if c.get("percentual_reajuste") == NAO_APLICAVEL
            else formatar_percentual(c.get("percentual_reajuste"))
        )
        linhas.append([
            rotulo,
            f"{c.get('data_inicio')} a {c.get('data_fim')}",
            c.get("data_pedido"), c.get("situacao"),
            variacao, formatar_fator(c.get("fator_acumulado")),
        ])
    historia.append(_tabela(
        linhas,
        [largura * 0.10, largura * 0.28, largura * 0.16, largura * 0.16,
         largura * 0.15, largura * 0.15],
        estilos, alinhamentos_direita={4, 5},
    ))

    historia.append(_paragrafo(
        "Meses sem efeito financeiro, por ciclo", estilos["subsecao"],
    ))
    linhas_sem = [["Ciclo", "Quantidade", "Competências sem efeito financeiro"]]
    for c in ciclos:
        bloco = c.get("meses_sem_efeito") or {}
        if bloco.get("status") == "ok":
            qtd = str(bloco.get("quantidade"))
            comp = ", ".join(bloco.get("competencias") or []) or "Nenhuma"
        else:
            qtd = bloco.get("status") or NAO_INFORMADO
            comp = bloco.get("status") or NAO_INFORMADO
        linhas_sem.append([c["ciclo"], qtd, comp])
    historia.append(_tabela(
        linhas_sem, [largura * 0.12, largura * 0.16, largura * 0.72], estilos,
    ))


def _bloco_financeiro(historia, dados, estilos) -> None:
    fin = dados.get("financeiro") or {}
    historia.append(_paragrafo("4. Valores financeiros", estilos["secao"]))
    largura = _largura_util()

    def _tabela_metodo(titulo: str, linhas_metodo: list[dict[str, Any]],
                       delta_total: Any) -> None:
        historia.append(_paragrafo(titulo, estilos["subsecao"]))
        if not linhas_metodo:
            historia.append(_paragrafo(
                f"{NAO_APLICAVEL} — sem evidências deste método no XLS processado.",
                estilos["normal"],
            ))
            return
        linhas = [["Ciclo", "Valor pago", "Valor atualizado", "Delta", "Evidências"]]
        for linha in linhas_metodo:
            linhas.append([
                linha.get("ciclo"), formatar_moeda(linha.get("valor_pago")),
                formatar_moeda(linha.get("valor_atualizado")),
                formatar_moeda(linha.get("delta")), str(linha.get("evidencias")),
            ])
        linhas.append(["Total (delta)", "", "", formatar_moeda(delta_total), ""])
        historia.append(_tabela(
            linhas,
            [largura * 0.16, largura * 0.24, largura * 0.24, largura * 0.22,
             largura * 0.14],
            estilos, alinhamentos_direita={1, 2, 3, 4},
        ))

    _tabela_metodo(
        "Método Financeiro — valor pago, valor atualizado e delta por ciclo",
        fin.get("financeiro_por_ciclo") or [],
        fin.get("delta_total_financeiro"),
    )
    _tabela_metodo(
        "Método Pedido de Compras — retroativo em análise, por ciclo",
        fin.get("pc_por_ciclo") or [],
        fin.get("delta_total_pc"),
    )

    historia.append(_paragrafo(
        "Situação do retroativo", estilos["subsecao"]
    ))
    if (fin.get("retroativo_total") is None
            and fin.get("retroativo_reconhecido_pago") is None
            and not fin.get("retroativo_por_estado")):
        historia.append(_paragrafo(
            f"{NAO_INFORMADO} — nenhum retroativo consolidado pelos motores "
            "oficiais no XLS processado.",
            estilos["normal"],
        ))
        return
    por_estado = fin.get("retroativo_por_estado") or {}
    linhas_estado = [["Situação", "Valor"]]
    linhas_estado.append([
        "Em análise",
        formatar_moeda(
            por_estado.get("em análise", fin.get("retroativo_em_analise"))
        ),
    ])
    linhas_estado.append([
        "Reconhecido", formatar_moeda(por_estado.get("reconhecido", 0.0)),
    ])
    linhas_estado.append([
        "Pago", formatar_moeda(por_estado.get("pago", 0.0)),
    ])
    historia.append(_tabela(
        linhas_estado, [largura * 0.62, largura * 0.38], estilos,
        alinhamentos_direita={1},
    ))


def _celulas_item_no_ciclo(item: dict[str, Any], ciclo: str) -> list[str]:
    """Formata (qtd, VU, total) de um item para um ciclo da tabela de itens.

    Pre-vigencia: ciclo anterior ao nascimento do item nao e dado ausente —
    celulas numericas ficam vazias (nunca zero nem "Nao informado", que fica
    reservado a ausencia real de dado em ciclo ja vigente).
    """
    nascimento = item.get("ciclo_nascimento")
    if nascimento is not None and int(ciclo[1:]) < int(nascimento):
        return ["", "", ""]
    vu = (
        item.get("vu_c0") if ciclo == "C0"
        else (item.get("vu_ciclos") or {}).get(ciclo)
    )
    total = (
        item.get("total_c0") if ciclo == "C0"
        else (item.get("total_ciclos") or {}).get(ciclo)
    )
    return [
        formatar_numero((item.get("quantidade_ciclos") or {}).get(ciclo)),
        formatar_moeda(vu),
        formatar_moeda(total),
    ]


def _bloco_itens(historia, dados, estilos) -> None:
    from reportlab.platypus import FrameBreak

    estilo_titulo_paginado = estilos["secao"].clone(
        "secao_itens_paginada", keepWithNext=0,
    )
    itens = dados.get("itens") or []
    if not itens:
        historia.append(_paragrafo(
            "5. Itens e valores atualizados", estilos["secao"]
        ))
        historia.append(_paragrafo(
            f"{NAO_INFORMADO} — itens do contrato não localizados no XLS "
            "processado.",
            estilos["normal"],
        ))
        return
    largura = _largura_util()
    ciclos_presentes = [
        c for c in ("C1", "C2", "C3", "C4")
        if any((i.get("vu_ciclos") or {}).get(c) is not None for i in itens)
    ]
    ciclos_tabela = ["C0"] + ciclos_presentes
    linhas: list[list[Any]] = [["Item"] + [
        rotulo
        for ciclo in ciclos_tabela
        for rotulo in (f"Qtd {ciclo}", f"VU {ciclo}", f"Total {ciclo}")
    ]]
    for i in itens:
        rotulo = str(i.get("item") if i.get("item") is not None else NAO_INFORMADO)
        if i.get("descricao"):
            rotulo = f"{rotulo} — {i['descricao']}"
        celulas = [rotulo]
        for ciclo in ciclos_tabela:
            celulas.extend(_celulas_item_no_ciclo(i, ciclo))
        linhas.append(celulas)
    n_valores = 3 * len(ciclos_tabela)
    larg_item = largura * 0.16
    larg_col = (largura - larg_item) / n_valores
    cabecalho, corpo = linhas[0], linhas[1:]
    blocos = [corpo[:10]]
    restante = corpo[10:]
    while restante:
        blocos.append(restante[:10])
        restante = restante[10:]
    for indice, bloco in enumerate(blocos):
        if indice == 0:
            titulo = "5. Itens e valores atualizados"
        else:
            historia.append(FrameBreak())
            titulo = "5. Itens e valores atualizados (continuação)"
        historia.append(_paragrafo(titulo, estilo_titulo_paginado))
        historia.append(_tabela(
            [cabecalho] + bloco,
            [larg_item] + [larg_col] * n_valores,
            estilos,
            alinhamentos_direita=set(range(1, 1 + n_valores)),
            permitir_quebra=False,
        ))


def _bloco_memoria(historia, dados, estilos) -> None:
    memoria = dados.get("memoria_calculo") or []
    historia.append(_paragrafo("6. Memória de cálculo", estilos["secao"]))
    if not memoria:
        historia.append(_paragrafo(
            f"{NAO_INFORMADO} — o XLS processado não possui a memória de "
            "cálculo persistida (arquivos legados anteriores à Etapa 4).",
            estilos["normal"],
        ))
        return
    largura = _largura_util()
    for bloco in memoria:
        historia.append(_paragrafo(
            f"Memória do ciclo {bloco['ciclo']}", estilos["subsecao"],
        ))
        linhas: list[list[Any]] = [[
            "Tipo", "Ord.", "Competência", "Índice", "Fator mensal",
            "Fator acumulado", "Variação final", "Método / fonte",
        ]]
        for reg in bloco.get("registros") or []:
            tipo = reg.get("tipo")
            if reg.get("valor_indice") is None:
                valor_indice = ""
            elif tipo == "INDICE":
                valor_indice = formatar_fator(reg.get("valor_indice"), 4)
            else:
                valor_indice = formatar_percentual(reg.get("valor_indice"))
            competencia = reg.get("competencia")
            if competencia == NAO_INFORMADO and tipo == "RESULTADO":
                competencia = ""
            linhas.append([
                tipo,
                "" if reg.get("ordem") is None else str(int(reg["ordem"])),
                competencia,
                valor_indice,
                formatar_fator(reg.get("fator_mensal"))
                if reg.get("fator_mensal") is not None else "",
                formatar_fator(reg.get("fator_acumulado"))
                if reg.get("fator_acumulado") is not None else "",
                formatar_percentual(reg.get("variacao_final"))
                if reg.get("variacao_final") is not None else "",
                reg.get("metodo_fonte") or "",
            ])
        historia.append(_tabela(
            linhas,
            [largura * 0.13, largura * 0.06, largura * 0.13, largura * 0.11,
             largura * 0.13, largura * 0.14, largura * 0.13, largura * 0.17],
            estilos, alinhamentos_direita={3, 4, 5, 6},
        ))


def _bloco_aditivos(historia, dados, estilos) -> None:
    aditivos = dados.get("aditivos") or {}
    historia.append(_paragrafo(
        "7. Alterações contratuais consideradas", estilos["secao"]
    ))
    itens = aditivos.get("itens") or []
    if not itens:
        historia.append(_paragrafo(
            f"{NAO_APLICAVEL} — nenhum aditivo computável no cálculo "
            "financeiro foi localizado no XLS processado.",
            estilos["normal"],
        ))
        return
    largura = _largura_util()
    linhas = [["Alteração", "Ciclo", "Valor atualizado", "Data da alteração"]]
    for a in itens:
        linhas.append([
            a.get("rotulo_documental"), a.get("ciclo"),
            formatar_moeda(a.get("valor_atualizado")),
            a.get("data_alteracao"),
        ])
    historia.append(_tabela(
        linhas,
        [largura * 0.40, largura * 0.12, largura * 0.24, largura * 0.24],
        estilos, alinhamentos_direita={2},
    ))


def _ref_data_pos(valor: Any) -> str:
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        return ""
    try:
        return valor.strftime("%d/%m/%Y")
    except AttributeError:
        return str(valor).strip()


def _bloco_referencias_vta_pdf(historia, dados, estilos) -> None:
    """Tres referencias do VTA (auditavel/comparativo) — nunca substitui o oficial.

    Consome a posicao mais atual valida (FORMA 1) quando disponivel; caso
    contrario apresenta a ultima abertura completa (FORMA 2) e informa
    explicitamente a condicao, sem apresentar a abertura como posicao atual.
    """
    ref = dados.get("referencias_vta") or {}
    if not ref.get("disponivel"):
        return
    historia.append(_paragrafo(
        "Referências auditáveis do Valor Total Atualizado", estilos["secao"]
    ))
    sintese = dados.get("sintese") or {}
    oficial = sintese.get("vta")
    if oficial is None:
        oficial = sintese.get("vta_previa")

    data_pos = _ref_data_pos(ref.get("data_posicao_atual"))
    ciclo_vig = ref.get("ciclo_vigente")
    ciclo_ab = ref.get("ciclo_ultima_abertura")
    f1 = ref.get("forma1_posicao_atual")
    contexto = []
    if data_pos:
        contexto.append(f"data de corte da posição atual: {data_pos}")
    elif f1 is None:
        contexto.append(
            "posição atual não informada — referência pela última posição de "
            "abertura completa"
        )
    if ciclo_vig:
        contexto.append(f"ciclo vigente: {ciclo_vig}")
    if ciclo_ab is not None:
        contexto.append(f"última abertura disponível: C{ciclo_ab}")
    if contexto:
        historia.append(_paragrafo("; ".join(contexto).capitalize() + ".",
                                   estilos["normal"]))
    historia.append(_paragrafo(
        "O valor oficial, com efeito jurídico, permanece o da cadeia homologada "
        "do Valor Total Atualizado do Contrato. As referências abaixo são apenas "
        "para conferência.", estilos["normal"]))

    largura = _largura_util()
    marc = "Não informada"

    def _m(v: Any) -> str:
        return formatar_moeda(v) if isinstance(v, (int, float)) else marc

    linhas = [
        ["Referência", "Valor", "Classificação"],
        ["Valor Total Atualizado (cadeia homologada)", _m(oficial), "OFICIAL"],
        ["VTA pela posição atual do contrato",
         marc if f1 is None else _m(f1), "REFERÊNCIA AUDITÁVEL"],
        ["VTA pela última posição de abertura disponível",
         _m(ref.get("forma2_ultima_abertura")), "REFERÊNCIA AUDITÁVEL"],
        ["Contrato original integralmente reajustado",
         _m(ref.get("forma3_integral_reajustado")), "COMPARATIVO"],
    ]
    historia.append(_tabela(
        linhas, [largura * 0.50, largura * 0.25, largura * 0.25],
        estilos, alinhamentos_direita={1},
    ))
    rec_v = ref.get("reconciliacao_valor")
    rec_s = ref.get("reconciliacao_status")
    if rec_s or isinstance(rec_v, (int, float)):
        txt = "Reconciliação (posição atual − última abertura): "
        txt += _m(rec_v) if isinstance(rec_v, (int, float)) else marc
        if rec_s:
            txt += f" — status: {rec_s}"
        historia.append(_paragrafo(txt + ".", estilos["normal"]))


def _bloco_observacoes(historia, dados, estilos) -> None:
    obs = dados.get("observacoes") or {}
    historia.append(_paragrafo("8. Observações de consistência", estilos["secao"]))
    historia.append(_paragrafo(
        "Grau de confiabilidade das evidências: "
        f"{_texto_ou_nao_informado(obs.get('grau_confiabilidade'))}.",
        estilos["normal"],
    ))
    blocos = (
        ("Pendências bloqueantes", obs.get("bloqueantes") or []),
        ("Advertências", obs.get("advertencias") or []),
        ("Ressalvas", obs.get("ressalvas") or []),
        ("Limitações de evidência", obs.get("limitacoes_evidencias") or []),
    )
    algum = False
    for titulo, itens in blocos:
        if not itens:
            continue
        algum = True
        historia.append(_paragrafo(titulo, estilos["subsecao"]))
        for item in itens:
            historia.append(_paragrafo(f"• {item}", estilos["normal"]))
    if not algum:
        historia.append(_paragrafo(
            "Nenhuma pendência, advertência ou ressalva registrada.",
            estilos["normal"],
        ))


def _paragrafo(texto: Any, estilo) -> Any:
    from reportlab.platypus import Paragraph
    return Paragraph(_escapar(texto), estilo)


def _escapar(texto: Any) -> str:
    return (
        str("" if texto is None else texto)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def gerar_sumario_executivo(
    leitura_ou_objeto: dict[str, Any] | None,
    identificacao: dict[str, Any] | None = None,
    memoria_calculo: dict[str, Any] | None = None,
) -> bytes:
    """Atalho: monta os dados canonicos e devolve o PDF em bytes."""
    dados = montar_dados_sumario_executivo(
        leitura_ou_objeto, identificacao=identificacao,
        memoria_calculo=memoria_calculo,
    )
    return gerar_sumario_executivo_pdf(dados)
