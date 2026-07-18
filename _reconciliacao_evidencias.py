# -*- coding: utf-8 -*-
"""Reconciliacao de Evidencias — grao evidencia/execucao (RFC v10.5.4).

Slice vertical minimo do arquetipo PC + Remanescentes:
- normaliza PC e Remanescentes em um modelo canonico de evidencia;
- gera chave canonica de execucao;
- classifica grandeza, periodo, ciclo, item, unidade e data de corte;
- determina comparabilidade (mesma grandeza + mesma execucao);
- identifica sobreposicoes e impede dupla contagem (deterministico);
- produz resultado em sombra, sem alterar VTA oficial, historico!B51,
  Calculadora, templates nem documentos homologados.

Camada DISTINTA de `_motor_reconciliacao.py` (grao ciclo, prevalencia
agregada): aqui o grao e a evidencia individual com taxonomia de grandezas.
Nenhum score: classificacao 100% deterministica por elegibilidade/grandeza.

Funcao pura: recebe o dict do leitor v10, devolve dict novo; nao muta a
entrada e nao escreve em workbook algum.
"""
from __future__ import annotations

from typing import Any

# ---------------------------------------------------------------------------
# Taxonomia de grandezas (RFC v10.5.4 §4.1 — elegibilidade fonte x grandeza)
# ---------------------------------------------------------------------------
GRANDEZA_QUANTIDADE_AUTORIZADA = "QUANTIDADE_AUTORIZADA"        # PC
GRANDEZA_SALDO_OPERACIONAL = "SALDO_OPERACIONAL"                # fotografia
GRANDEZA_EXECUCAO_FINANCEIRA = "EXECUCAO_FINANCEIRA"            # financeiro
GRANDEZA_EXECUCAO_FISICA = "EXECUCAO_FISICA"                    # consumidos
GRANDEZA_SALDO_HISTORICO = "SALDO_HISTORICO_RECONSTRUIDO"       # cascata

TAXONOMIA_GRANDEZAS = {
    GRANDEZA_QUANTIDADE_AUTORIZADA: (
        "Demanda autorizada por Pedido de Compra; nao prova pagamento nem consumo."
    ),
    GRANDEZA_SALDO_OPERACIONAL: "Saldo declarado (fotografia) na data de corte.",
    GRANDEZA_EXECUCAO_FINANCEIRA: "Execucao financeira reconhecida (pagamentos).",
    GRANDEZA_EXECUCAO_FISICA: "Execucao fisica (quantidades consumidas).",
    GRANDEZA_SALDO_HISTORICO: "Saldo reconstruido pela cascata de eventos.",
}

# Classificacoes deterministicas (sem score)
CLASSIF_PRINCIPAL = "PRINCIPAL"
CLASSIF_CORROBORANTE = "CORROBORANTE"
CLASSIF_DIVERGENTE = "DIVERGENTE"
CLASSIF_SOBREPOSTA = "SOBREPOSTA"
CLASSIF_NAO_COMPARAVEL = "NAO_COMPARAVEL"
CLASSIF_RESOLVIDA_GCC = "RESOLVIDA_GCC"

# Estados de materialidade (RFC v10.5.4 §6.3 — 4 estados)
MAT_TOLERANCIA_ARREDONDAMENTO = "TOLERANCIA_ARREDONDAMENTO"
MAT_DIVERGENCIA_IMATERIAL = "DIVERGENCIA_IMATERIAL_REGISTRADA"
MAT_DIVERGENCIA_MATERIAL = "DIVERGENCIA_MATERIAL"
MAT_DIVERGENCIA_ESTRUTURAL = "DIVERGENCIA_ESTRUTURAL"
MAT_COMPARABILIDADE_NAO_AVALIADA = "COMPARABILIDADE_NAO_AVALIADA"

# Limites EXPERIMENTAIS, NAO HOMOLOGADOS (RFC v10.5.4 §6.3). Configuraveis
# por parametro; nunca cristalizados no nucleo. Percentuais sobre a base.
LIMITES_EXPERIMENTAIS: dict[str, float] = {
    "arredondamento_abs": 2.00,        # ate R$ 2,00 E ...
    "arredondamento_rel": 0.000001,    # ... ate 0,0001% da base
    "operacional_abs": 100.00,         # acima de R$ 100,00 OU ...
    "operacional_rel": 0.001,          # ... 0,10% da base => MATERIAL
}


def _tofl(valor: Any, default: float | None = None) -> float | None:
    try:
        if valor in (None, ""):
            return default
        return float(valor)
    except (TypeError, ValueError):
        return default


def _txt(valor: Any) -> str:
    return str(valor or "").strip()


def classificar_materialidade(
    diferenca: float,
    base: float,
    limites: dict[str, float] | None = None,
    estrutural: bool = False,
) -> str:
    """Classifica uma diferenca nos 4 estados de materialidade.

    ``estrutural=True`` (unidade, ciclo, periodo, existencia ou elegibilidade
    divergentes) e material INDEPENDENTEMENTE do valor.
    """
    if estrutural:
        return MAT_DIVERGENCIA_ESTRUTURAL
    lim = dict(LIMITES_EXPERIMENTAIS)
    lim.update(limites or {})
    dif = abs(diferenca)
    base_abs = abs(base) or 1.0
    rel = dif / base_abs
    if dif <= lim["arredondamento_abs"] and rel <= lim["arredondamento_rel"]:
        return MAT_TOLERANCIA_ARREDONDAMENTO
    if dif > lim["operacional_abs"] or rel > lim["operacional_rel"]:
        return MAT_DIVERGENCIA_MATERIAL
    return MAT_DIVERGENCIA_IMATERIAL


def chave_canonica(ev: dict[str, Any]) -> str:
    """Chave canonica da execucao/evidencia (RFC v10.5.4 §3.1).

    Campos: processo (anonimizado), item, ciclo, periodo/competencia,
    data de corte, grandeza, unidade, tipo de evento, referencia documental,
    id da execucao, natureza.
    """
    partes = [
        _txt(ev.get("processo_ref")),
        _txt(ev.get("item")),
        _txt(ev.get("ciclo")),
        _txt(ev.get("periodo")),
        _txt(ev.get("data_corte")),
        _txt(ev.get("grandeza")),
        _txt(ev.get("unidade")),
        _txt(ev.get("tipo_evento")),
        _txt(ev.get("referencia_documental")),
        _txt(ev.get("id_execucao")),
        _txt(ev.get("natureza")),
    ]
    return "|".join(partes)


def _nova_evidencia(**campos: Any) -> dict[str, Any]:
    ev: dict[str, Any] = {
        "processo_ref": "",          # identificador anonimizado (hash truncado)
        "item": "",
        "ciclo": "",
        "periodo": "",
        "data_corte": "",
        "grandeza": "",
        "unidade": "",
        "tipo_evento": "",
        "referencia_documental": "",
        "id_execucao": "",
        "natureza": "",              # DECLARADA | DERIVADA | CALCULADA
        "fonte": "",
        "valor": None,
        "quantidade": None,
        "origem_aba": "",
        "origem_linha": None,
        "classificacao": "",
        "justificativa_classificacao": "",
        "comparavel_com": [],
        "efeito_no_calculo": "",
    }
    ev.update(campos)
    ev["chave_canonica"] = chave_canonica(ev)
    ev["id_evidencia"] = (
        f"{ev['fonte']}:{ev['origem_aba']}:{ev['origem_linha']}:{ev['item']}"
    )
    return ev


# ---------------------------------------------------------------------------
# Normalizadores (slice PC + Remanescentes)
# ---------------------------------------------------------------------------

def _evidencias_financeiro(res: dict[str, Any], processo_ref: str) -> list[dict[str, Any]]:
    """Normaliza execução financeira já aceita pelo VTA sombra, sem recalcular."""
    evidencias: list[dict[str, Any]] = []
    for indice, parcela in enumerate(
        (res.get("vta_sombra") or {}).get("parcelas_computadas") or [], 1
    ):
        fonte = _txt(parcela.get("fonte_parcela"))
        if fonte not in {"Financeiro", "Historico financeiro"}:
            continue
        ciclo = _txt(parcela.get("ciclo")).upper()
        if not ciclo:
            for parte in _txt(parcela.get("identificador")).split(":"):
                if parte.strip().upper() in {f"C{i}" for i in range(5)}:
                    ciclo = parte.strip().upper()
                    break
        evidencias.append(_nova_evidencia(
            processo_ref=processo_ref,
            item=_txt(parcela.get("item")),
            ciclo=ciclo,
            periodo=_txt(parcela.get("competencia") or parcela.get("data")),
            grandeza=GRANDEZA_EXECUCAO_FINANCEIRA,
            unidade="BRL",
            tipo_evento="EXECUCAO_FINANCEIRA",
            referencia_documental=_txt(parcela.get("referencia_documental")),
            id_execucao=_txt(parcela.get("identificador")) or f"financeiro:{indice}",
            natureza="CALCULADA",
            fonte="financeiro",
            valor=_tofl(parcela.get("valor")),
            origem_aba=("financeiro" if fonte == "Financeiro" else "CICLOS_PASSADOS"),
            origem_linha=parcela.get("origem_linha") or indice,
            classificacao=CLASSIF_PRINCIPAL,
            justificativa_classificacao=(
                "Evidencia elegivel de EXECUCAO_FINANCEIRA reconhecida pelo VTA sombra; "
                "nao compete com saldo operacional nem com autorizacao do PC."
            ),
            efeito_no_calculo=(
                "Execucao financeira ja aceita; prevalece apenas sobre evidencias da "
                "mesma grandeza e nao gera soma adicional na reconciliacao."
            ),
        ))
    return evidencias


def _evidencias_consumidos(res: dict[str, Any], processo_ref: str) -> list[dict[str, Any]]:
    """Normaliza consumo físico valorizado sem tratá-lo como pagamento."""
    evidencias: list[dict[str, Any]] = []
    for indice, reg in enumerate(
        (res.get("itens_consumidos_v10") or {}).get("itens") or [], 1
    ):
        item = _txt(reg.get("item") or reg.get("item_ou_grupo"))
        evidencias.append(_nova_evidencia(
            processo_ref=processo_ref,
            item=item,
            ciclo=_txt(reg.get("ciclo_inferido") or reg.get("ciclo")).upper(),
            periodo=_txt(reg.get("data_referencia")),
            grandeza=GRANDEZA_EXECUCAO_FISICA,
            unidade="BRL",
            tipo_evento="CONSUMO_FISICO",
            referencia_documental=_txt(reg.get("referencia_documental")),
            id_execucao=_txt(reg.get("identificador")) or f"consumo:{item}:{indice}",
            natureza="DECLARADA",
            fonte="consumidos",
            valor=_tofl(reg.get("valor_total")),
            quantidade=_tofl(reg.get("qtd_total") or reg.get("quantidade") or reg.get("qtd")),
            origem_aba="ENTRADA_XLS_CONSUMIDOS",
            origem_linha=reg.get("linha") or indice,
            classificacao=CLASSIF_PRINCIPAL,
            justificativa_classificacao=(
                "Evidencia elegivel de EXECUCAO_FISICA; comprova consumo, mas nao "
                "prova pagamento e nao compete com a fotografia de saldo."
            ),
            efeito_no_calculo=(
                "Consumo fisico valorizado; permanece nao aditivo diante de evidencia "
                "financeira da mesma execucao."
            ),
        ))
    return evidencias

def _evidencias_pc(res: dict[str, Any], processo_ref: str) -> list[dict[str, Any]]:
    evidencias: list[dict[str, Any]] = []
    for reg in (res.get("itens_pc_v10") or {}).get("itens") or []:
        campos = reg.get("campos_vta") or {}
        refletido = _txt(campos.get("ja_refletido_em")) or "Nao"
        sobreposta = refletido not in ("", "Nao")
        if sobreposta:
            classif = CLASSIF_SOBREPOSTA
            justif = (
                f"PC ja refletido em '{refletido}' (JA_REFLETIDO_EM): "
                "mesma execucao ja coberta por outra fonte."
            )
            efeito = "Nao soma ao VTA (anti-dupla contagem deterministica)."
        else:
            classif = CLASSIF_PRINCIPAL
            justif = (
                "Unica evidencia elegivel da grandeza QUANTIDADE_AUTORIZADA "
                "para esta execucao (PC autoriza; nao prova pagamento nem consumo)."
            )
            efeito = (
                "Demanda autorizada; entra no VTA somente quando a execucao "
                "financeira correspondente for reconhecida (fora deste slice)."
            )
        evidencias.append(_nova_evidencia(
            processo_ref=processo_ref,
            item=_txt(reg.get("item_ou_grupo")),
            ciclo=_txt(reg.get("ciclo")).upper(),
            periodo=_txt(reg.get("data_pc")),
            grandeza=GRANDEZA_QUANTIDADE_AUTORIZADA,
            unidade="BRL",
            tipo_evento="PEDIDO_COMPRA",
            referencia_documental=_txt(reg.get("numero_pc")),
            id_execucao=_txt(reg.get("numero_pc")) or _txt(reg.get("item_ou_grupo")),
            natureza="DECLARADA",
            fonte="pc",
            valor=_tofl(reg.get("valor_pc")),
            origem_aba="itens_PC",
            origem_linha=reg.get("linha"),
            classificacao=classif,
            justificativa_classificacao=justif,
            efeito_no_calculo=efeito,
        ))
    return evidencias


def _evidencias_remanescentes(
    res: dict[str, Any],
    processo_ref: str,
    limites: dict[str, float] | None,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    """Fotografias de saldo (SALDO_OPERACIONAL) + checagem interna de
    coerencia (valor declarado x qtd_saldo * VU), nos 4 estados.
    """
    evidencias: list[dict[str, Any]] = []
    divergencias: list[dict[str, Any]] = []
    data_corte = _txt((res.get("controle") or {}).get("data_corte"))
    for reg in (res.get("execucao_saldo") or {}).get("itens") or []:
        item = _txt(reg.get("item"))
        qtd = _tofl(reg.get("qtd_saldo"))
        valor = _tofl(reg.get("valor_saldo"))
        vu = _tofl(reg.get("vu_original"))
        ev = _nova_evidencia(
            processo_ref=processo_ref,
            item=item,
            data_corte=data_corte,
            grandeza=GRANDEZA_SALDO_OPERACIONAL,
            unidade="BRL",
            tipo_evento="FOTOGRAFIA_SALDO",
            referencia_documental=_txt(reg.get("pc") or reg.get("requisicao_sap")),
            id_execucao=f"saldo:{item}:{data_corte}",
            natureza="DECLARADA",
            fonte="remanescente_informado",
            valor=valor,
            quantidade=qtd,
            origem_aba="ENTRADA_XLS_REMANESCENTES",
            origem_linha=reg.get("linha"),
            classificacao=CLASSIF_PRINCIPAL,
            justificativa_classificacao=(
                "Unica fonte elegivel da grandeza SALDO_OPERACIONAL na data "
                "de corte (fotografia declarada pelo fiscal)."
            ),
            efeito_no_calculo=(
                "Saldo operacional na data de corte; nao se compara com "
                "cascata sem diagnostico de comparabilidade (RFC v10.5.4 §5)."
            ),
        )
        # Coerencia interna: valor declarado x (qtd_saldo * VU).
        if valor is not None and qtd is not None and vu:
            derivado = round(qtd * vu, 2)
            dif = round(valor - derivado, 2)
            if dif != 0.0:
                estado = classificar_materialidade(dif, valor or derivado, limites)
                divergencias.append({
                    "registro_id": f"evidencia:{ev['id_evidencia']}",
                    "id_evidencia": ev["id_evidencia"],
                    "item": item,
                    "tipo": "COERENCIA_INTERNA_FOTOGRAFIA",
                    "valor_declarado": valor,
                    "valor_derivado_qtd_x_vu": derivado,
                    "diferenca": dif,
                    "estado_materialidade": estado,
                    "recomendacao_automatica": (
                        "Manter fotografia declarada como referencia SEM "
                        "computar ate decisao (conservador; nenhuma fonte "
                        "selecionada silenciosamente)."
                    ),
                    "encaminhamento": (
                        "Equalizacao GCC" if estado in (
                            MAT_DIVERGENCIA_MATERIAL, MAT_DIVERGENCIA_ESTRUTURAL
                        ) else "registro"
                    ),
                })
                if estado in (MAT_DIVERGENCIA_MATERIAL, MAT_DIVERGENCIA_ESTRUTURAL):
                    ev["classificacao"] = CLASSIF_DIVERGENTE
                    ev["justificativa_classificacao"] = (
                        f"Valor declarado (R$ {valor:,.2f}) diverge do derivado "
                        f"qtd x VU (R$ {derivado:,.2f}) em R$ {dif:,.2f} — {estado}."
                    )
                    ev["efeito_no_calculo"] = (
                        "Nao computavel ate decisao da Equalizacao GCC "
                        "(nenhuma fonte e selecionada silenciosamente)."
                    )
        evidencias.append(ev)
    return evidencias, divergencias


def _ponte_fotografia_cascata(
    res: dict[str, Any],
    limites: dict[str, float] | None,
) -> tuple[list[dict[str, Any]], list[str]]:
    """Ponte item a item fotografia x cascata (RFC v10.5.4 §5.2).

    Fotografia = qtd_saldo declarada (SALDO_OPERACIONAL).
    Cascata    = quantidade_vigente (posicao contratual, ultimo ciclo)
                 menos consumo acumulado (itens_Consumidos), x VU.
    Comportamento conservador: quando a cascata nao e derivavel para o item,
    registra COMPARABILIDADE_NAO_AVALIADA (pendente GCC) — nunca escolhe
    silenciosamente uma fonte nem declara a divergencia resolvida.
    """
    divergencias: list[dict[str, Any]] = []
    alertas: list[str] = []

    fotos = {
        _txt(r.get("item")): r
        for r in (res.get("execucao_saldo") or {}).get("itens") or []
        if _txt(r.get("item"))
    }
    if not fotos:
        return divergencias, alertas

    # Cascata: quantidade vigente por item (ultimo ciclo disponivel).
    vigente: dict[str, float] = {}
    for linha in (res.get("posicao_contratual_sombra") or {}).get("linhas_quantidade") or []:
        item = _txt(linha.get("item"))
        qtd = _tofl(linha.get("quantidade_vigente"))
        if item and qtd is not None:
            vigente[item] = qtd  # linhas ordenadas por ciclo; a ultima prevalece

    consumido: dict[str, float] = {}
    for reg in (res.get("itens_consumidos_v10") or {}).get("itens") or []:
        item = _txt(reg.get("item") or reg.get("item_ou_grupo"))
        qtd = _tofl(reg.get("quantidade") or reg.get("qtd"))
        if item and qtd is not None:
            consumido[item] = consumido.get(item, 0.0) + qtd

    total_dif = 0.0
    itens_comparados = 0
    itens_nao_avaliaveis = 0
    for item, foto in fotos.items():
        qtd_foto = _tofl(foto.get("qtd_saldo"))
        vu = _tofl(foto.get("vu_original"))
        if item not in vigente or qtd_foto is None or not vu:
            itens_nao_avaliaveis += 1
            divergencias.append({
                "registro_id": f"evidencia:fotografia_x_cascata:{item}",
                "id_evidencia": f"ponte:{item}",
                "item": item,
                "tipo": "FOTOGRAFIA_X_CASCATA",
                "qtd_fotografia": qtd_foto,
                "qtd_cascata": None,
                "valor_declarado": (
                    round(qtd_foto * vu, 2)
                    if qtd_foto is not None and vu is not None else None
                ),
                "valor_derivado_qtd_x_vu": None,
                "diferenca": None,
                "estado_materialidade": MAT_COMPARABILIDADE_NAO_AVALIADA,
                "causa_diagnostico": (
                    "Quantidade vigente, quantidade da fotografia ou VU "
                    "insuficiente para reconstruir a cascata."
                ),
                "recomendacao_automatica": (
                    "Obter ou validar os dados ausentes; nenhuma fonte foi selecionada."
                ),
                "encaminhamento": "Equalizacao GCC",
            })
            continue
        qtd_cascata = round(vigente[item] - consumido.get(item, 0.0), 4)
        dif_qtd = round(qtd_foto - qtd_cascata, 4)
        if dif_qtd == 0.0:
            itens_comparados += 1
            continue
        dif_valor = round(dif_qtd * vu, 2)
        base = round((qtd_cascata or qtd_foto) * vu, 2)
        estado = classificar_materialidade(dif_valor, base, limites)
        registro_id = f"evidencia:fotografia_x_cascata:{item}"
        divergencias.append({
            "registro_id": registro_id,
            "id_evidencia": f"ponte:{item}",
            "item": item,
            "tipo": "FOTOGRAFIA_X_CASCATA",
            "qtd_fotografia": qtd_foto,
            "qtd_cascata": qtd_cascata,
            "valor_declarado": round(qtd_foto * vu, 2),
            "valor_derivado_qtd_x_vu": round(qtd_cascata * vu, 2),
            "diferenca": dif_valor,
            "estado_materialidade": estado,
            "recomendacao_automatica": (
                "Grandezas potencialmente distintas (saldo operacional x "
                "saldo historico reconstruido): manter ambas registradas, "
                "nao computar a diferenca ate decisao da GCC."
            ),
            "encaminhamento": (
                "Equalizacao GCC" if estado in (
                    MAT_DIVERGENCIA_MATERIAL, MAT_DIVERGENCIA_ESTRUTURAL
                ) else "registro"
            ),
        })
        total_dif += dif_valor
        itens_comparados += 1

    if itens_nao_avaliaveis:
        alertas.append(
            f"Ponte fotografia x cascata: {itens_nao_avaliaveis} item(ns) sem "
            "cascata derivavel (COMPARABILIDADE_NAO_AVALIADA) — pendente de "
            "dados/decisao GCC; nenhuma fonte foi selecionada."
        )
    if total_dif:
        alertas.append(
            f"Ponte fotografia x cascata: diferenca agregada de "
            f"R$ {abs(total_dif):,.2f} em {itens_comparados} item(ns) "
            "comparado(s); itens materiais encaminhados a Equalizacao GCC."
        )
    return divergencias, alertas


def _aplicar_decisoes_gcc(
    evidencias: list[dict[str, Any]],
    divergencias: list[dict[str, Any]],
    decisoes: dict[str, dict[str, Any]] | None,
) -> None:
    """Aplica decisoes vigentes da GCC (log apartado JSONL, append-only).

    A decisao NUNCA altera o Excel nem o VTA oficial: apenas reclassifica a
    evidencia (RESOLVIDA_GCC) e registra a decisao na divergencia.
    """
    if not decisoes:
        return
    por_id = {ev["id_evidencia"]: ev for ev in evidencias}
    for div in divergencias:
        dec = decisoes.get(_txt(div.get("registro_id")))
        if not dec:
            continue
        div["decisao_gcc"] = dec
        div["encaminhamento"] = "decidido_gcc"
        ev = por_id.get(_txt(div.get("id_evidencia")))
        if ev is not None:
            ev["classificacao"] = CLASSIF_RESOLVIDA_GCC
            ev["justificativa_classificacao"] = (
                f"Decisao GCC ({dec.get('timestamp')}): {dec.get('decisao')} — "
                f"{dec.get('justificativa') or 'sem justificativa'}."
            )
            ev["efeito_no_calculo"] = (
                f"Conforme decisao GCC registrada: {dec.get('decisao')}."
            )


# ---------------------------------------------------------------------------
# Comparabilidade e sobreposicao
# ---------------------------------------------------------------------------

def _grupos_comparaveis(evidencias: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Agrupa por (grandeza, item): apenas evidencias da MESMA grandeza e
    mesma execucao/item sao diretamente comparaveis. Grandezas distintas
    coexistem como NAO_COMPARAVEL entre si (nunca competem)."""
    grupos: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for ev in evidencias:
        grupos.setdefault((ev["grandeza"], ev["item"]), []).append(ev)

    saida: list[dict[str, Any]] = []
    for (grandeza, item), membros in sorted(grupos.items()):
        ids = [m["id_evidencia"] for m in membros]
        for m in membros:
            m["comparavel_com"] = [i for i in ids if i != m["id_evidencia"]]
        saida.append({
            "grandeza": grandeza,
            "item": item,
            "evidencias": ids,
            "comparaveis_entre_si": len(ids) > 1,
        })
    return saida


def _sobreposicoes(evidencias: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Duplicidade deterministica: mesma chave canonica => mesma execucao.
    A primeira permanece; as demais viram SOBREPOSTA (nada e apagado)."""
    vistos: dict[str, str] = {}
    saida: list[dict[str, Any]] = []
    for ev in evidencias:
        chave = ev["chave_canonica"]
        if chave in vistos and ev["classificacao"] != CLASSIF_SOBREPOSTA:
            ev["classificacao"] = CLASSIF_SOBREPOSTA
            ev["justificativa_classificacao"] = (
                f"Mesma chave canonica de {vistos[chave]}: mesma execucao "
                "registrada em duplicidade."
            )
            ev["efeito_no_calculo"] = "Nao soma (anti-dupla contagem)."
            saida.append({
                "id_evidencia": ev["id_evidencia"],
                "sobrepoe": vistos[chave],
                "chave_canonica": chave,
            })
        else:
            vistos.setdefault(chave, ev["id_evidencia"])
    return saida


# ---------------------------------------------------------------------------
# Entrada publica
# ---------------------------------------------------------------------------

def reconciliar_evidencias(
    res_leitor: dict[str, Any],
    limites: dict[str, float] | None = None,
    decisoes: dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Reconciliacao de evidencias em sombra (slice PC + Remanescentes).

    Nao muta ``res_leitor``; nao altera VTA oficial, vta_sombra,
    historico!B51, Calculadora, templates nem documentos. Saida inspecionavel.
    """
    saida: dict[str, Any] = {
        "ok": False,
        "evidencias": [],
        "grupos_comparaveis": [],
        "sobreposicoes": [],
        "divergencias": [],
        "divergencias_materiais": [],
        "ponte_fotografia_cascata": [],
        "decisoes_gcc_aplicadas": [],
        "resumo": {},
        "alertas": [],
        "configuracao": {
            "limites": {**LIMITES_EXPERIMENTAIS, **(limites or {})},
            "homologado": False,
            "observacao": (
                "Limites experimentais (RFC v10.5.4 §6.3); calibrar nos "
                "arquetipos e aprovar na GCC antes de homologar."
            ),
            "taxonomia_grandezas": dict(TAXONOMIA_GRANDEZAS),
        },
    }
    if not isinstance(res_leitor, dict):
        saida["alertas"].append("Saida do leitor ausente; reconciliacao nao gerada.")
        return saida

    processo_ref = _txt(res_leitor.get("hash_entrada"))[:12] or "sem-hash"

    evidencias = _evidencias_financeiro(res_leitor, processo_ref)
    evidencias += _evidencias_consumidos(res_leitor, processo_ref)
    evidencias += _evidencias_pc(res_leitor, processo_ref)
    ev_rem, divergencias = _evidencias_remanescentes(res_leitor, processo_ref, limites)
    evidencias += ev_rem
    ponte, alertas_ponte = _ponte_fotografia_cascata(res_leitor, limites)
    fotos_por_item = {
        e.get("item"): e for e in evidencias
        if e.get("fonte") == "remanescente_informado"
    }
    for div in ponte:
        ev = fotos_por_item.get(div.get("item"))
        if ev is not None:
            div["id_evidencia"] = ev["id_evidencia"]
            if div.get("estado_materialidade") in (
                MAT_DIVERGENCIA_MATERIAL, MAT_DIVERGENCIA_ESTRUTURAL
            ):
                ev["classificacao"] = CLASSIF_DIVERGENTE
                ev["justificativa_classificacao"] = (
                    "Fotografia diverge materialmente da cascata reconstruida; "
                    "nenhuma fonte foi selecionada sem decisao GCC."
                )
                ev["efeito_no_calculo"] = "Nao computavel ate decisao da GCC."
    divergencias += ponte
    _aplicar_decisoes_gcc(evidencias, divergencias, decisoes)

    saida["sobreposicoes"] = _sobreposicoes(evidencias)
    saida["grupos_comparaveis"] = _grupos_comparaveis(evidencias)
    saida["divergencias"] = divergencias
    saida["ponte_fotografia_cascata"] = ponte
    saida["decisoes_gcc_aplicadas"] = [
        d["decisao_gcc"] for d in divergencias if d.get("decisao_gcc")
    ]
    saida["evidencias"] = evidencias
    saida["alertas"].extend(alertas_ponte)

    por_classif: dict[str, int] = {}
    for ev in evidencias:
        por_classif[ev["classificacao"]] = por_classif.get(ev["classificacao"], 0) + 1
    materiais = [
        d for d in divergencias
        if d.get("encaminhamento") != "decidido_gcc"
        and d["estado_materialidade"] in (
            MAT_DIVERGENCIA_MATERIAL, MAT_DIVERGENCIA_ESTRUTURAL
        )
    ]
    saida["divergencias_materiais"] = materiais
    nao_avaliadas = sum(
        d.get("estado_materialidade") == MAT_COMPARABILIDADE_NAO_AVALIADA
        and d.get("encaminhamento") != "decidido_gcc"
        for d in divergencias
    )
    saida["resumo"] = {
        "total_evidencias": len(evidencias),
        "por_classificacao": por_classif,
        "por_grandeza": {
            g: sum(1 for e in evidencias if e["grandeza"] == g)
            for g in sorted({e["grandeza"] for e in evidencias})
        },
        "divergencias_materiais": len(materiais),
        "encaminhadas_equalizacao": len(materiais),
        "comparabilidades_nao_avaliadas": nao_avaliadas,
        "resolvidas_gcc": len(saida["decisoes_gcc_aplicadas"]),
    }
    for d in materiais:
        rotulo = (
            "ESTRUTURAL"
            if d["estado_materialidade"] == MAT_DIVERGENCIA_ESTRUTURAL
            else "MATERIAL"
        )
        saida["alertas"].append(
            f"DIVERGENCIA_{rotulo} item {d['item']}: "
            f"R$ {abs(d['diferenca']):,.2f} — encaminhada a Equalizacao GCC."
        )
    saida["ok"] = True
    return saida
