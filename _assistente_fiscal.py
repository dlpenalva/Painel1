"""Assistente operacional — Cl8us 3.0.

Camada de ORIENTACAO em linguagem de negocio, somente leitura e sombra:
nao altera o valor historico, historico!B51, template nem documentos. Consome o
view-model do Painel Executivo (que ja consome Motor Temporal, Posicao,
Estado Contratual, Event Log e VTA sombra) e responde as perguntas do
usuario:

  * O contrato possui evidencia suficiente?
  * Qual metodologia foi escolhida e por que? (o sistema escolhe; o
    usuario apenas confirma)
  * Qual indice foi utilizado?
  * O reajuste e aplicavel?
  * Existe retroativo? Qual o valor? Em qual estado do ciclo de vida
    (em analise -> reconhecido -> pago)? O calculo NUNCA promove estado;
    somente ato administrativo.
  * Qual o valor atualizado do contrato? Qual o saldo remanescente?
  * Existem inconsistencias? Qual a proxima acao recomendada?

Regras permanentes honradas: PC e EXECUCAO (nunca aditivo); calendario,
formalizacao e financeiro sao independentes; todo numero informa de onde
veio e onde pode ser conferido.
"""
from __future__ import annotations

from typing import Any

from _conducao_gcc import diagnosticar_modo_conducao_gcc
from _dossie_decisao import montar_dossie_decisao
from _painel_executivo import montar_painel_executivo

# Ciclo de vida do retroativo (Regra permanente 5).
ESTADO_EM_ANALISE = "em análise"
ESTADO_RECONHECIDO = "reconhecido"
ESTADO_PAGO = "pago"
ESTADO_MISTO = "misto"

# Metodologias oficiais do Cl8us 3.0.
MET_FINANCEIRO = "Financeiro"
MET_PC = "Pedidos de Compra (PC)"
MET_CONSUMIDOS = "Itens Consumidos"
MET_FIN_REMANESC = "Financeiro + Remanescentes"

CONF_ALTA = "alta"
CONF_MEDIA = "média"
CONF_NULA = "insuficiente"

# Acoes recomendadas (vocabulario fechado, em linguagem operacional).
ACAO_REVISAR_BLOQUEANTES = "Revisar inconsistências bloqueantes"
ACAO_SOLICITAR_FINANCEIRO = "Solicitar evidência financeira oficial"
ACAO_REVISAR_DUPLA = "Revisar possível dupla contagem"
ACAO_REVISAR_DIVERGENCIA = "Revisar divergência de ciclo"
ACAO_GERAR_APOSTILAMENTO = "Gerar apostilamento"
ACAO_AGUARDAR_PAGAMENTO = "Aguardar pagamento / confirmar execução dos PCs"
ACAO_NENHUMA = "Nenhuma pendência"

_NATUREZA_PARA_ESTADO = {
    "potencial": ESTADO_EM_ANALISE,
    "reconhecido": ESTADO_RECONHECIDO,
    "ja_pago": ESTADO_PAGO,
}


def _f(valor: Any) -> float:
    try:
        return float(valor or 0.0)
    except (TypeError, ValueError):
        return 0.0


# --------------------------------------------------------------------------- #
# Evidencia disponivel (leitura pura do que o leitor ja materializou)
# --------------------------------------------------------------------------- #
def _levantar_evidencia(leitura: dict[str, Any]) -> dict[str, Any]:
    """Conta as fontes de evidencia ja lidas — nao abre Excel, nao recalcula."""
    eventos = (leitura.get("event_log_sombra") or {}).get("eventos") or []
    fontes = {"financeiro": 0, "remanescentes": 0, "aditivos": 0}
    valor_financeiro = 0.0
    retro_reconhecido_financeiro = 0.0
    for ev in eventos:
        fonte = str(ev.get("fonte_parcela") or "").strip().lower()
        if fonte == "financeiro":
            fontes["financeiro"] += 1
            valor_financeiro += _f(ev.get("valor"))
            if "reconhecido" in str(ev.get("tipo_financeiro") or "").lower():
                retro_reconhecido_financeiro += _f(ev.get("valor"))
        elif "remanescente" in fonte:
            fontes["remanescentes"] += 1
        elif fonte == "aditivo":
            fontes["aditivos"] += 1

    itens_pc = (leitura.get("itens_pc_v10") or {}).get("itens") or []
    consumidos = (leitura.get("itens_consumidos_v10") or {}).get("itens") or []
    exec_saldo = (leitura.get("execucao_saldo") or {}).get("itens") or []

    return {
        "parcelas_financeiro": fontes["financeiro"],
        "valor_financeiro": round(valor_financeiro, 2),
        "retro_reconhecido_financeiro": round(retro_reconhecido_financeiro, 2),
        "parcelas_remanescentes": fontes["remanescentes"],
        "parcelas_aditivos": fontes["aditivos"],
        "pcs": len(itens_pc),
        "itens_consumidos": len(consumidos),
        "itens_execucao_saldo": len(exec_saldo),
    }


# --------------------------------------------------------------------------- #
# Escolha da metodologia (o sistema escolhe; o usuario confirma)
# --------------------------------------------------------------------------- #
def _escolher_metodologia(ev: dict[str, Any]) -> dict[str, Any]:
    tem_fin = ev["parcelas_financeiro"] > 0
    tem_rem = ev["parcelas_remanescentes"] > 0 or ev["itens_execucao_saldo"] > 0
    tem_pc = ev["pcs"] > 0
    tem_cons = ev["itens_consumidos"] > 0

    candidatas = [
        (MET_FIN_REMANESC, tem_fin and tem_rem, CONF_ALTA,
         "Há execução financeira realizada e saldo remanescente informado: "
         "a parte realizada usa o financeiro e a parte futura usa o remanescente — "
         "a leitura mais fiel de contrato em execução."),
        (MET_FINANCEIRO, tem_fin, CONF_ALTA,
         f"Há {ev['parcelas_financeiro']} parcela(s) financeira(s) lida(s) "
         "(aba financeiro): o pagamento efetivo é a evidência de maior confiabilidade."),
        (MET_PC, tem_pc, CONF_MEDIA,
         f"Há {ev['pcs']} Pedido(s) de Compra com execução comandada. PC é execução "
         "(nunca aditivo): gera obrigação e retroativo EM ANÁLISE, ainda não reconhecido."),
        (MET_CONSUMIDOS, tem_cons, CONF_MEDIA,
         f"Há {ev['itens_consumidos']} item(ns) de consumo informados: quantidade "
         "consumida × valor unitário reajustado."),
    ]

    escolhida = next((c for c in candidatas if c[1]), None)
    alternativas = []
    for nome, disponivel, conf, _just in candidatas:
        if escolhida and nome == escolhida[0]:
            continue
        alternativas.append({
            "nome": nome,
            "disponivel": bool(disponivel),
            "motivo": ("Evidência presente, mas com confiabilidade menor que a escolhida."
                       if disponivel else "Evidência não encontrada no MasterFile."),
            "confiabilidade": conf if disponivel else CONF_NULA,
        })

    if escolhida is None:
        return {
            "escolhida": None,
            "confiabilidade": CONF_NULA,
            "justificativa": ("Nenhuma evidência de execução localizada "
                              "(sem financeiro, sem PC, sem consumo, sem remanescente)."),
            "alternativas": alternativas,
            "fonte": "Event Log sombra + abas itens_PC / itens_Consumidos / itens_Execucao_Saldo",
        }
    return {
        "escolhida": escolhida[0],
        "confiabilidade": escolhida[2],
        "justificativa": escolhida[3],
        "alternativas": alternativas,
        "fonte": "Event Log sombra + abas itens_PC / itens_Consumidos / itens_Execucao_Saldo",
    }


# --------------------------------------------------------------------------- #
# Retroativo: valor + ciclo de vida (o calculo nunca promove estado)
# --------------------------------------------------------------------------- #
def _classificar_retroativo(painel: dict[str, Any], ev: dict[str, Any]) -> dict[str, Any]:
    por_estado = {ESTADO_EM_ANALISE: 0.0, ESTADO_RECONHECIDO: 0.0, ESTADO_PAGO: 0.0}
    origens: list[dict[str, Any]] = []
    for pc in (painel.get("situacao_pcs") or {}).get("pcs") or []:
        valor = _f(pc.get("retroativo"))
        if valor <= 0.0:
            continue
        estado = _NATUREZA_PARA_ESTADO.get(str(pc.get("natureza_delta")), ESTADO_EM_ANALISE)
        por_estado[estado] = round(por_estado[estado] + valor, 2)
        origens.append({"origem": f"PC {pc.get('numero_pc')}",
                        "ciclo": pc.get("ciclo_temporal"),
                        "valor": valor, "estado": estado})

    reconhecido_fin = ev.get("retro_reconhecido_financeiro") or 0.0
    if reconhecido_fin > 0.0 and not origens:
        # Sem PCs: o retroativo evidenciado e o reconhecido no financeiro.
        por_estado[ESTADO_RECONHECIDO] = round(reconhecido_fin, 2)
        origens.append({"origem": "aba financeiro (EFEITO_FINANCEIRO=Sim)",
                        "ciclo": None, "valor": reconhecido_fin,
                        "estado": ESTADO_RECONHECIDO})

    total = round(sum(por_estado.values()), 2)
    com_valor = [e for e, v in por_estado.items() if v > 0.0]
    estado = com_valor[0] if len(com_valor) == 1 else (ESTADO_MISTO if com_valor else None)
    return {
        "existe": total > 0.0,
        "valor": total,
        "estado": estado,
        "por_estado": por_estado,
        "origens": origens,
        "oficial_resumo": (painel.get("retroativo") or {}).get("oficial_resumo"),
        "fonte": ("Motor Temporal (Q9, natureza do delta por PC) e parcelas "
                  "reconhecidas da aba financeiro; o estado só muda por ato administrativo."),
    }


# --------------------------------------------------------------------------- #
# Inconsistencias em linguagem de negocio + proxima acao
# --------------------------------------------------------------------------- #
# Traducao de codigos tecnicos para linguagem operacional. O codigo original
# permanece como referencia de auditoria; muda apenas a frase exibida.
_TRADUCAO_ALERTAS = {
    "LINHA_TEMPORAL_INVALIDA": (
        "O calendário de ciclos (aba parametros) está incompleto ou com datas "
        "inválidas. Sem ele, nenhum Pedido de Compra pode ser enquadrado no ciclo."),
    "LINHA_TEMPORAL_INCOMPLETA": (
        "O calendário de ciclos (aba parametros) está incompleto; não foi "
        "possível enquadrar o Pedido de Compra no ciclo."),
    "DATA_CICLO_INVALIDA": (
        "Há ciclo com data de início/fim ausente ou inconsistente na aba "
        "parametros. Corrija as datas para permitir o enquadramento temporal."),
    "PC_SEM_DATA": (
        "Há Pedido de Compra sem DATA_PC na aba itens_PC. Sem a data, o PC "
        "não pode ser enquadrado em nenhum ciclo."),
    "FATOR_INDETERMINADO": (
        "O ciclo do PC está sem FATOR_ACUMULADO numérico na aba parametros; "
        "o valor devido não pôde ser calculado."),
    "PC_CICLO_INFORMADO_INVALIDO": (
        "O ciclo informado no PC não é um ciclo válido (C0-C4); prevalece o "
        "enquadramento pela linha temporal."),
    "VALOR_CONTRATO_INDISPONIVEL": (
        "O valor vigente por ciclo não pôde ser derivado (aba historico_VU "
        "sem dados); se houver resumo oficial, ele será usado como referência, "
        "sem estimar valor inexistente."),
}


def _traduzir_inconsistencias(painel: dict[str, Any]) -> list[dict[str, Any]]:
    itens: list[dict[str, Any]] = []
    for a in painel.get("alertas") or []:
        nivel = str(a.get("nivel") or "INFO").upper()
        codigo = str(a.get("codigo") or "")
        msg = str(a.get("mensagem_negocio") or a.get("mensagem") or "")
        if not a.get("mensagem_negocio"):
            msg = _TRADUCAO_ALERTAS.get(codigo, msg)
        texto_l = (codigo + " " + msg).lower()
        if nivel == "ERRO_GRAVE":
            itens.append({"gravidade": "bloqueante", "codigo": codigo,
                          "descricao": msg, "acao": ACAO_REVISAR_BLOQUEANTES})
        elif "dupl" in texto_l or "ja_refletido" in texto_l or "já refletido" in texto_l:
            itens.append({"gravidade": "atenção", "codigo": codigo,
                          "descricao": ("Possível dupla contagem: a mesma execução pode "
                                        "estar no financeiro e em outra fonte. " + msg),
                          "acao": ACAO_REVISAR_DUPLA})
        elif codigo == "CICLO_DIVERGENTE":
            itens.append({"gravidade": "atenção", "codigo": codigo,
                          "descricao": (msg + " Prevalece o enquadramento pela linha "
                                        "temporal dos ciclos (regra permanente)."),
                          "acao": ACAO_REVISAR_DIVERGENCIA})
        elif codigo == "PAGO_INDETERMINADO":
            itens.append({"gravidade": "informação", "codigo": codigo,
                          "descricao": msg + " Sem confirmação de pagamento, o retroativo "
                                             "permanece em análise.",
                          "acao": ACAO_AGUARDAR_PAGAMENTO})
        elif nivel == "INFO" and codigo in _TRADUCAO_ALERTAS:
            itens.append({"gravidade": "informação", "codigo": codigo,
                          "descricao": msg, "acao": ""})
        elif nivel == "ALERTA":
            itens.append({"gravidade": "atenção", "codigo": codigo,
                          "descricao": msg, "acao": ""})
    return itens


def _proxima_acao(
    metodologia: dict[str, Any],
    retro: dict[str, Any],
    inconsistencias: list[dict[str, Any]],
) -> dict[str, str]:
    bloqueantes = [i for i in inconsistencias if i["gravidade"] == "bloqueante"]
    if bloqueantes:
        return {"acao": ACAO_REVISAR_BLOQUEANTES,
                "motivo": f"{len(bloqueantes)} inconsistência(s) bloqueante(s) impedem "
                          "qualquer decisão segura sobre o reajuste."}
    if metodologia.get("escolhida") is None:
        return {"acao": ACAO_SOLICITAR_FINANCEIRO,
                "motivo": "Não há evidência de execução no MasterFile; sem financeiro, "
                          "PC, consumo ou remanescente não é possível apurar reajuste."}
    if any(i["acao"] == ACAO_REVISAR_DUPLA for i in inconsistencias):
        return {"acao": ACAO_REVISAR_DUPLA,
                "motivo": "Antes de reconhecer valores, confirme que a mesma execução "
                          "não foi contada duas vezes."}
    if any(i["acao"] == ACAO_REVISAR_DIVERGENCIA for i in inconsistencias):
        return {"acao": ACAO_REVISAR_DIVERGENCIA,
                "motivo": "Há PC com ciclo informado diferente do enquadramento temporal; "
                          "confirme a correção antes de formalizar."}
    if retro["por_estado"][ESTADO_RECONHECIDO] > 0.0:
        return {"acao": ACAO_GERAR_APOSTILAMENTO,
                "motivo": "Há retroativo reconhecido pendente de formalização."}
    if retro["por_estado"][ESTADO_EM_ANALISE] > 0.0:
        return {"acao": ACAO_AGUARDAR_PAGAMENTO,
                "motivo": "O retroativo apurado é potencial (PCs sem pagamento "
                          "confirmado); reconhecimento exige ato administrativo."}
    return {"acao": ACAO_NENHUMA,
            "motivo": "Sem retroativo pendente e sem inconsistências que exijam ação."}


def _metodologia_do_motor(motor: dict[str, Any], fallback: dict[str, Any]) -> dict[str, Any]:
    """Converte o Motor de Metodologias para o contrato legado do Assistente."""
    rec = motor.get("recomendada") or {}
    if not rec:
        return fallback

    conf = rec.get("confiabilidade")
    if conf == "media":
        conf = CONF_MEDIA
    alternativas = []
    for alt in motor.get("alternativas") or []:
        if alt.get("nome") == rec.get("nome"):
            continue
        alternativas.append({
            "nome": alt.get("nome"),
            "disponivel": bool(alt.get("disponivel")),
            "motivo": (
                "Evidencia presente, mas com pontuacao menor que a recomendada."
                if alt.get("disponivel") else "Evidencia nao encontrada no MasterFile."
            ),
            "confiabilidade": alt.get("confiabilidade") if alt.get("disponivel") else CONF_NULA,
            "valor_recomendado": alt.get("valor_recomendado"),
        })
    return {
        "escolhida": rec.get("nome"),
        "confiabilidade": conf or fallback.get("confiabilidade"),
        "justificativa": rec.get("justificativa") or fallback.get("justificativa"),
        "alternativas": alternativas,
        "fonte": rec.get("fonte") or fallback.get("fonte"),
        "valor_recomendado": rec.get("valor_recomendado"),
        "score": rec.get("score"),
    }


# --------------------------------------------------------------------------- #
# Entrada publica
# --------------------------------------------------------------------------- #
def montar_assistente_fiscal(
    leitura: dict[str, Any],
    painel: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Monta as respostas do Assistente operacional (funcao pura, sombra)."""
    from _objeto_processo_reajuste import consumidor_do_objeto
    assistente_objeto = consumidor_do_objeto(leitura, "assistente_fiscal")
    if assistente_objeto is not None:
        return assistente_objeto

    if painel is None:
        painel = montar_painel_executivo(leitura)
    if not painel.get("disponivel"):
        return {"disponivel": False, "motivo": painel.get("motivo")}

    ev = _levantar_evidencia(leitura)
    metodologia = _escolher_metodologia(ev)
    motor_metodologias = painel.get("motor_metodologias") or {}
    if motor_metodologias.get("disponivel"):
        metodologia = _metodologia_do_motor(motor_metodologias, metodologia)
    retro = _classificar_retroativo(painel, ev)
    inconsistencias = _traduzir_inconsistencias(painel)
    acao = _proxima_acao(metodologia, retro, inconsistencias)

    linha_temporal_ok = not any(
        i["gravidade"] == "bloqueante" for i in inconsistencias)
    evidencia_suficiente = metodologia.get("escolhida") is not None and linha_temporal_ok

    sr = painel.get("situacao_reajuste") or {}
    ciclos_aplicaveis = sr.get("ciclos_aplicaveis") or []
    indice_por_ciclo = [
        {"ciclo": c.get("ciclo"), "percentual": c.get("indice_percentual"),
         "fator_acumulado": c.get("fator_acumulado")}
        for c in sr.get("ciclos") or [] if c.get("reajuste_aplicavel")
    ]

    sc = painel.get("situacao_contrato") or {}
    saldo_ciclos = (painel.get("posicao_contratual") or {}).get("saldo_por_ciclo") or []
    saldo_ref = next(
        (s.get("saldo") for s in reversed(saldo_ciclos)
         if s.get("ciclo") == sc.get("ciclo_referencia") and s.get("saldo") is not None),
        None)
    resumo_oficial = (painel.get("situacao_financeira") or {}).get("resumo_oficial") or {}

    resposta = {
        "disponivel": True,
        "marco": painel.get("marco"),
        "masterfile_inteligente": painel.get("masterfile_inteligente") or {},
        "evidencia": {
            "suficiente": evidencia_suficiente,
            "detalhe": ev,
            "fonte": "Leitura do MasterFile (Event Log sombra e abas de itens).",
        },
        "metodologia": metodologia,
        "motor_metodologias": motor_metodologias,
        "resultado_recomendado": motor_metodologias.get("resultado_recomendado"),
        "indice": {
            "por_ciclo": indice_por_ciclo,
            "fonte": "Aba parametros (PERCENTUAL_REAJUSTE / FATOR_ACUMULADO), via Motor Temporal.",
        },
        "reajuste_aplicavel": {
            "resposta": bool(ciclos_aplicaveis),
            "ciclos": ciclos_aplicaveis,
            "fonte": "Motor Temporal (Q1), pela linha temporal dos ciclos.",
        },
        "retroativo": retro,
        "vta": {
            "valor": (painel.get("vta") or {}).get("oficial"),
            "conferencia_sombra": (painel.get("vta") or {}).get("metodologia_sombra"),
            "diferenca": (painel.get("vta") or {}).get("diferenca_sombra"),
            "fonte": ("Resumo histórico do MasterFile; conferência pela "
                      "metodologia sombra — este valor isolado não é VTA."),
        },
        "valor_contrato": {
            "valor": sc.get("valor_hoje"),
            "ciclo_referencia": sc.get("ciclo_referencia"),
            "fonte": (
                "Valor do contrato hoje não disponível no MasterFile; nenhum valor foi estimado."
                if sc.get("valor_hoje") is None
                else "Resumo oficial do MasterFile (valor total atualizado) — a posição "
                "contratual sombra não pôde derivar o valor vigente."
                if str(sc.get("fonte_valor") or "").startswith("resumo_oficial")
                else "Motor de Posição Contratual (sombra), valor vigente do ciclo de referência."
            ),
        },
        "saldo": {
            "valor": saldo_ref if saldo_ref is not None else resumo_oficial.get("saldo_remanescente"),
            "fonte": ("Posição contratual sombra (vigente - executado no ciclo de referência)"
                      if saldo_ref is not None else "Resumo oficial do MasterFile."),
        },
        "inconsistencias": inconsistencias,
        "proxima_acao": acao,
        "semaforo": (painel.get("semaforo") or {}).get("status"),
    }
    resposta["modo_conducao_gcc"] = diagnosticar_modo_conducao_gcc(
        leitura, painel, resposta
    )
    resposta["dossie_decisao"] = montar_dossie_decisao(resposta, painel)
    return resposta


# --------------------------------------------------------------------------- #
# Renderizacao (somente leitura; "o que faco agora?" no topo)
# --------------------------------------------------------------------------- #
def _brl(valor: Any) -> str:
    try:
        return "R$ " + f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "—"


def _pct(valor: Any) -> str:
    try:
        return f"{float(valor) * 100:.2f}%".replace(".", ",")
    except (TypeError, ValueError):
        return "—"


def render_assistente_fiscal(leitura: dict[str, Any],
                             painel: dict[str, Any] | None = None) -> None:
    import streamlit as st

    ass = montar_assistente_fiscal(leitura, painel)
    if not ass.get("disponivel"):
        st.info(f"Assistente operacional indisponível: {ass.get('motivo')}")
        return

    st.subheader("Assistente operacional")

    # --- 1. O que faco agora? ---
    acao = ass["proxima_acao"]
    if acao["acao"] == ACAO_NENHUMA:
        st.success(f"**Próxima ação: {acao['acao']}.** {acao['motivo']}")
    elif acao["acao"] in (ACAO_REVISAR_BLOQUEANTES, ACAO_SOLICITAR_FINANCEIRO):
        st.error(f"**Próxima ação: {acao['acao']}.** {acao['motivo']}")
    else:
        st.warning(f"**Próxima ação: {acao['acao']}.** {acao['motivo']}")

    dossie = ass.get("dossie_decisao") or {}
    if dossie.get("disponivel"):
        texto = (
            f"**ConclusÃ£o operacional:** {dossie.get('ato_ou_providencia')}. "
            f"{dossie.get('motivo')}"
        )
        if dossie.get("pode_concluir_no_claus_new"):
            st.success(texto)
        else:
            st.warning(texto)
        st.caption(dossie.get("resumo_executivo") or "")

    mfi = ass.get("masterfile_inteligente") or {}
    if mfi.get("layout_inteligente"):
        st.info(
            "**MasterFile Inteligente:** "
            f"{mfi.get('quantidade_campos_obrigatorios')} campos obrigatÃ³rios; "
            f"{mfi.get('reducao_estimada_campos')} campos deixaram de ser "
            "preenchidos manualmente."
        )

    # --- 2. Veredicto em linguagem de negocio ---
    met = ass["metodologia"]
    retro = ass["retroativo"]
    col_a, col_b = st.columns(2)
    with col_a:
        evid = ass["evidencia"]
        st.write("**O contrato possui evidência suficiente?** "
                 + ("Sim." if evid["suficiente"] else "Não."))
        st.write("**Metodologia escolhida:** "
                 + (f"{met['escolhida']} (confiabilidade {met['confiabilidade']})."
                    if met.get("escolhida") else "nenhuma — evidência insuficiente."))
        st.caption(met["justificativa"])
        resultado = ass.get("resultado_recomendado") or {}
        if resultado:
            st.write("**Resultado recomendado:** "
                     f"{_brl(resultado.get('valor_recomendado'))}.")
            st.caption(
                "A recomendacao pode ser confirmada; as demais metodologias ficam "
                "apenas para conferencia."
            )
        aplic = ass["reajuste_aplicavel"]
        st.write("**O reajuste é aplicável?** "
                 + (f"Sim, nos ciclos {', '.join(aplic['ciclos'])}."
                    if aplic["resposta"] else "Não."))
        indices = ass["indice"]["por_ciclo"]
        if indices:
            st.write("**Índice utilizado:** "
                     + "; ".join(f"{i['ciclo']}: {_pct(i['percentual'])}" for i in indices)
                     + ".")
            st.caption(ass["indice"]["fonte"])
    with col_b:
        if retro["existe"]:
            st.write(f"**Existe retroativo?** Sim: **{_brl(retro['valor'])}** "
                     f"— estado: **{retro['estado']}**.")
            for o in retro["origens"]:
                st.caption(f"{_brl(o['valor'])} originado de {o['origem']}"
                           + (f" (ciclo {o['ciclo']})" if o.get("ciclo") else "")
                           + f" — {o['estado']}.")
        else:
            st.write("**Existe retroativo?** Não identificado nesta apuração.")
        st.write(f"**Valor atualizado do contrato (VTA):** {_brl(ass['vta']['valor'])}.")
        st.caption(ass["vta"]["fonte"])
        st.write(f"**Saldo remanescente:** {_brl(ass['saldo']['valor'])}.")
        st.caption(ass["saldo"]["fonte"])

    # --- 3. Inconsistencias ---
    inc = ass["inconsistencias"]
    if inc:
        st.write(f"**Existem inconsistências?** Sim, {len(inc)}:")
        for i in inc:
            texto = f"({i['gravidade']}) {i['descricao']}"
            if i.get("acao"):
                texto += f" → {i['acao']}."
            if i["gravidade"] == "bloqueante":
                st.error(texto)
            elif i["gravidade"] == "atenção":
                st.warning(texto)
            else:
                st.caption(texto)
    else:
        st.write("**Existem inconsistências?** Não.")

    st.caption(
        "Confirmacao operacional: a metodologia foi escolhida pelo sistema a partir "
        "da evidência disponível. O estado do retroativo só muda por ato "
        "administrativo — nunca pelo cálculo."
    )
