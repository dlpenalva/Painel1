"""Painel Executivo do Cl8us 3.0.

Camada de APRESENTACAO somente leitura. Nao calcula nada novo e nao duplica
regra de negocio: consome exclusivamente os motores ja homologados —

  * Motor Temporal (`montar_motor_temporal`) — Q1-Q13;
  * Motor de Posicao Contratual — via `posicao_contratual_sombra` do leitor;
  * Estado Contratual / Event Log — via chaves sombra do leitor;
  * VTA — via `vta_sombra` do leitor e Q12 do Motor Temporal;
  * Integracao sombra — chaves ja materializadas em `ler_masterfile_v10`.

A pergunta que esta camada responde: "Consigo entender o contrato em menos
de 30 segundos?". Nada aqui escreve Excel, session_state ou disco.
"""
from __future__ import annotations

from typing import Any

from _motor_metodologias import montar_motor_metodologias
from _motor_temporal import montar_motor_temporal

NIVEL_CRITICO = "critico"
NIVEL_ATENCAO = "atencao"
NIVEL_OK = "ok"

_TRADUCAO_ALERTAS_NEGOCIO = {
    "LINHA_TEMPORAL_INVALIDA": (
        "O calendario de ciclos da aba parametros esta incompleto ou com datas "
        "invalidas. Sem esse calendario, os Pedidos de Compra nao podem ser "
        "enquadrados no ciclo correto."),
    "LINHA_TEMPORAL_INCOMPLETA": (
        "O calendario de ciclos da aba parametros esta incompleto; nao foi "
        "possivel enquadrar o Pedido de Compra no ciclo."),
    "DATA_CICLO_INVALIDA": (
        "Ha ciclo com data de inicio ou fim ausente/inconsistente na aba "
        "parametros. Corrija as datas para permitir o enquadramento temporal."),
    "PC_SEM_DATA": (
        "Ha Pedido de Compra sem DATA_PC na aba itens_PC. Sem a data, o PC nao "
        "pode ser enquadrado em nenhum ciclo."),
    "FATOR_INDETERMINADO": (
        "O ciclo do PC esta sem FATOR_ACUMULADO numerico na aba parametros; o "
        "valor devido nao pode ser calculado."),
    "PC_CICLO_INFORMADO_INVALIDO": (
        "O ciclo informado no PC nao e valido (C0-C4). O sistema desconsiderou "
        "esse campo e usou somente a DATA_PC contra a linha temporal dos ciclos."),
    "PAGO_INDETERMINADO": (
        "O PC nao informa se houve pagamento. Sem essa confirmacao, o "
        "valor permanece em analise."),
    "VALOR_CONTRATO_INDISPONIVEL": (
        "O valor vigente por ciclo nao esta disponivel no MasterFile. O painel "
        "so mostra valor alternativo se houver resumo oficial preenchido."),
    "ESTADO_CONTRATUAL_INDISPONIVEL": (
        "A memoria contratual sombra nao ficou disponivel para conferencia. A "
        "leitura oficial do MasterFile permanece preservada."),
    "VTA_INDISPONIVEL": (
        "A conferencia metodologica do valor historico nao ficou disponivel. "
        "O dado lido do MasterFile permanece preservado, mas nao e tratado "
        "como VTA sem o potencial restante."),
}


def _mensagem_alerta_negocio(alerta: dict[str, Any]) -> str:
    codigo = str(alerta.get("codigo") or "")
    mensagem = str(alerta.get("mensagem") or "")
    if codigo == "AVISO_LEITOR":
        if "layout anterior de itens_PC" in mensagem:
            return mensagem
        if "DATA_PC nao informada" in mensagem:
            return "Ha Pedido de Compra sem DATA_PC na aba itens_PC."
        if "VALOR_PC nao informado" in mensagem:
            return "Ha Pedido de Compra sem VALOR_PC na aba itens_PC."
        if "DATA_PC fora dos ciclos" in mensagem:
            return (
                "Ha Pedido de Compra cuja DATA_PC nao foi enquadrada nos ciclos "
                "da aba parametros.")
        return mensagem
    return _TRADUCAO_ALERTAS_NEGOCIO.get(codigo, mensagem)


# --------------------------------------------------------------------------- #
# Adaptacao leitura -> Motor Temporal (so renomeia chaves, nao recalcula)
# --------------------------------------------------------------------------- #
def _parcelas_do_event_log(leitura: dict[str, Any]) -> list[dict[str, Any]]:
    """Reconstroi as parcelas-base a partir do Event Log ja materializado.

    O leitor v10 nao expoe `parcelas_sombra` diretamente, mas cada parcela
    virou um evento `parcela_base_lida` no Event Log sombra. Aqui apenas
    devolvemos os mesmos dados ao formato de entrada do motor (adaptacao,
    sem novo calculo).
    """
    eventos = (leitura.get("event_log_sombra") or {}).get("eventos") or []
    parcelas: list[dict[str, Any]] = []
    for ev in eventos:
        if str(ev.get("tipo_evento") or "") != "parcela_base_lida":
            continue
        parcelas.append({
            "identificador": ev.get("identificador"),
            "origem_dado": ev.get("origem_dado"),
            "tipo_financeiro": ev.get("tipo_financeiro"),
            "status_consolidacao": ev.get("status_consolidacao"),
            "ciclo": ev.get("ciclo"),
            "linha": ev.get("linha"),
            "valor": ev.get("valor"),
            "computa_vta": ev.get("computa_vta"),
            "ja_refletido_em": ev.get("ja_refletido_em"),
            "fonte_parcela": ev.get("fonte_parcela"),
            "justificativa_vta": ev.get("justificativa"),
        })
    return parcelas


def _adaptar_para_motor(leitura: dict[str, Any]) -> dict[str, Any]:
    """Copia rasa da leitura com as chaves no formato do Motor Temporal."""
    res = dict(leitura)
    res["ciclos"] = {
        "por_ciclo": dict((leitura.get("parametros_v10") or {}).get("por_ciclo") or {}),
    }
    itens_pc = leitura.get("itens_pc_v10") or {}
    if not (itens_pc.get("itens") if isinstance(itens_pc, dict) else itens_pc):
        itens_pc = leitura.get("itens_pc") or {}
    res["itens_pc"] = itens_pc
    if not res.get("parcelas_sombra"):
        res["parcelas_sombra"] = _parcelas_do_event_log(leitura)
    return res


# --------------------------------------------------------------------------- #
# View-model (funcao pura)
# --------------------------------------------------------------------------- #
def _semaforo(alertas: list[dict[str, Any]]) -> dict[str, Any]:
    por_nivel: dict[str, int] = {}
    for a in alertas:
        nivel = str(a.get("nivel") or "INFO").upper()
        por_nivel[nivel] = por_nivel.get(nivel, 0) + 1
    if por_nivel.get("ERRO_GRAVE"):
        status = NIVEL_CRITICO
    elif por_nivel.get("ALERTA"):
        status = NIVEL_ATENCAO
    else:
        status = NIVEL_OK
    return {"status": status, "por_nivel": por_nivel, "total": len(alertas)}


def montar_painel_executivo(
    leitura: dict[str, Any],
    marco: str = "",
) -> dict[str, Any]:
    """Monta o view-model do Painel Executivo a partir da leitura do upload.

    Somente leitura: nao muta `leitura`, nao persiste nada e nao cria regra
    nova — cada campo abaixo e repasse direto de um motor existente.
    """
    if not marco:
        from _objeto_processo_reajuste import consumidor_do_objeto
        painel_objeto = consumidor_do_objeto(leitura, "painel_executivo")
        if painel_objeto is not None:
            return painel_objeto

    if not isinstance(leitura, dict) or not leitura.get("ok"):
        return {"disponivel": False,
                "motivo": "Leitura do Masterfile ausente ou invalida."}

    por_ciclo = (leitura.get("parametros_v10") or {}).get("por_ciclo") or {}
    if not por_ciclo:
        return {"disponivel": False,
                "motivo": ("Painel Executivo requer Masterfile v10 "
                           "(aba parametros com ciclos C0-C4).")}

    controle = leitura.get("controle") or {}
    marco = str(marco or controle.get("ciclo_vigente") or "")

    motor = montar_motor_temporal(_adaptar_para_motor(leitura), marco=marco)
    resumo = leitura.get("resumo") or {}
    posicao = leitura.get("posicao_contratual_sombra") or {}
    estado = leitura.get("estado_contratual_sombra") or motor.estado_contratual or {}
    vta_leitor = leitura.get("vta_sombra") or {}

    # --- Alertas consolidados (motor + posicao sombra + avisos do leitor) ---
    alertas: list[dict[str, Any]] = [dict(a) for a in motor.alertas]
    for a in posicao.get("alertas_bloqueantes") or []:
        alertas.append({"nivel": "ERRO_GRAVE",
                        "codigo": a.get("codigo") or "POSICAO_BLOQUEANTE",
                        "mensagem": a.get("mensagem") or str(a),
                        "identificador": a.get("identificador") or ""})
    for aviso in leitura.get("avisos") or []:
        alertas.append({"nivel": "INFO", "codigo": "AVISO_LEITOR",
                        "mensagem": str(aviso), "identificador": ""})
    for alerta in alertas:
        alerta["mensagem_negocio"] = _mensagem_alerta_negocio(alerta)

    # --- Situacao do contrato (controle + Estado Contratual + Q11) ---
    # Fallback HONESTO do valor do contrato: quando a posicao sombra nao
    # deriva o valor vigente (sem historico_VU populado), usa o VTA do
    # resumo oficial do Masterfile — declarando a fonte, nunca inventando.
    valor_hoje = (motor.valor_contrato or {}).get("valor_hoje")
    fonte_valor = (motor.valor_contrato or {}).get("fonte")
    if valor_hoje is None:
        valor_hoje = resumo.get("valor_total_atualizado")
        if valor_hoje is not None:
            fonte_valor = "resumo_oficial (fallback)"
    situacao_contrato = {
        "marco": marco,
        "modo": controle.get("modo") or "",
        "ciclo_vigente": controle.get("ciclo_vigente") or "",
        "data_corte": controle.get("data_corte"),
        "versao_masterfile": controle.get("versao") or leitura.get("versao_detectada") or "",
        "valor_hoje": valor_hoje,
        "ciclo_referencia": (motor.valor_contrato or {}).get("ciclo_referencia"),
        "valor_por_ciclo": dict((motor.valor_contrato or {}).get("por_ciclo") or {}),
        "fonte_valor": fonte_valor,
        "eventos_processados": estado.get("eventos_processados", 0),
        "valores_por_origem": dict(estado.get("valores_por_origem") or {}),
    }

    # --- Situacao do reajuste (Q1/Q2/Q3 por ciclo) ---
    ciclos = [dict(c) for c in motor.ciclos]
    situacao_reajuste = {
        "ciclos": ciclos,
        "ciclos_aplicaveis": [c["ciclo"] for c in ciclos if c.get("reajuste_aplicavel")],
        "ciclos_na_apuracao": [c["ciclo"] for c in ciclos if c.get("computa_nesta_apuracao")],
    }

    # --- Situacao financeira (Q5/Q6/Q7 + resumo oficial do Masterfile) ---
    situacao_financeira = {
        "totais": dict(motor.totais),
        "resumo_oficial": {
            "valor_total_atualizado": resumo.get("valor_total_atualizado"),
            "retroativo": resumo.get("retroativo"),
            "execucao_atualizada": resumo.get("execucao_atualizada"),
            "saldo_remanescente": resumo.get("saldo_remanescente"),
        },
    }

    # --- Situacao dos Pedidos de Compra (Q4-Q9) ---
    pcs = [pc.to_dict() for pc in motor.pcs]
    situacao_pcs = {
        "total": len(pcs),
        "com_divergencia_ciclo": sum(1 for p in pcs if p.get("divergencia_ciclo")),
        "sem_data": sum(1 for p in pcs if not p.get("data_pc")),
        "pagos": sum(1 for p in pcs if (p.get("valor_pago") or 0.0) > 0.0),
        "pcs": pcs,
    }

    # --- Posicao contratual (motor de posicao via sombra + Q13) ---
    posicao_contratual = {
        "ok": bool(posicao.get("ok")),
        "saldo_por_ciclo": [dict(s) for s in motor.saldo_por_ciclo],
        "resumo_por_item_ciclo": list(posicao.get("resumo_por_item_ciclo") or []),
        "adaptador": dict(posicao.get("adaptador") or {}),
        "bloqueantes": len(posicao.get("alertas_bloqueantes") or []),
    }

    # --- VTA (Q12 + metodologia sombra do leitor) ---
    vta_motor = dict(motor.vta or {})
    vta = {
        "oficial": vta_motor.get("oficial", vta_leitor.get("vta_oficial")),
        "metodologia_sombra": vta_motor.get("metodologia_sombra",
                                            vta_leitor.get("vta_sombra")),
        "metodologia_sombra_integral": vta_motor.get("metodologia_sombra_integral"),
        "diferenca_sombra": vta_motor.get("diferenca_sombra",
                                          vta_leitor.get("diferenca")),
        "diferenca_integral": vta_motor.get("diferenca_integral"),
        "triangulacao": dict(vta_leitor.get("triangulacao") or {}),
        "potencial_futuro": dict(leitura.get("potencial_futuro") or {}),
    }

    # --- Retroativo (Q9 por PC + total) ---
    retroativo = {
        "total_apurado": motor.totais.get("retroativo"),
        "oficial_resumo": resumo.get("retroativo"),
        "por_pc": [
            {"numero_pc": p.get("numero_pc"), "ciclo": p.get("ciclo_temporal"),
             "retroativo": p.get("retroativo"), "natureza": p.get("natureza_delta")}
            for p in pcs if p.get("retroativo")
        ],
    }

    painel = {
        "disponivel": True,
        "marco": marco,
        "semaforo": _semaforo(alertas),
        "masterfile_inteligente": dict((leitura.get("itens_pc_v10") or {}).get("masterfile_inteligente") or {}),
        "situacao_contrato": situacao_contrato,
        "situacao_reajuste": situacao_reajuste,
        "situacao_financeira": situacao_financeira,
        "situacao_pcs": situacao_pcs,
        "posicao_contratual": posicao_contratual,
        "vta": vta,
        "retroativo": retroativo,
        "alertas": alertas,
        "regras_cobertas": list(motor.regras_cobertas),
    }
    painel["motor_metodologias"] = montar_motor_metodologias(leitura, painel)
    return painel


# --------------------------------------------------------------------------- #
# Renderizacao Streamlit (somente leitura: nenhum input, nenhum botao)
# --------------------------------------------------------------------------- #
def _brl(valor: Any) -> str:
    try:
        return "R$ " + f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (TypeError, ValueError):
        return "—"


def _data_br(valor: Any) -> str:
    if hasattr(valor, "strftime"):
        return valor.strftime("%d/%m/%Y")
    return str(valor or "—")


def _pct(valor: Any) -> str:
    try:
        return f"{float(valor) * 100:.2f}%".replace(".", ",")
    except (TypeError, ValueError):
        return "—"


def render_painel_executivo(leitura: dict[str, Any]) -> None:
    """Renderiza o Painel Executivo logo apos o upload (somente leitura)."""
    import pandas as pd
    import streamlit as st

    painel = montar_painel_executivo(leitura)
    if not painel.get("disponivel"):
        st.info(f"Painel Executivo indisponível: {painel.get('motivo')}")
        return

    # Sprint 3: o Assistente operacional responde "o que faco agora?" no topo;
    # o detalhamento abaixo vira conferencia ("como esse calculo foi produzido?").
    from _assistente_fiscal import render_assistente_fiscal  # import local evita ciclo
    render_assistente_fiscal(leitura, painel)

    sem = painel["semaforo"]
    resumo_niveis = " · ".join(
        f"{n}: {q}" for n, q in sorted(sem["por_nivel"].items())) or "sem apontamentos"

    # --- Linha 1: os 4 numeros que respondem "o contrato em 30 segundos" ---
    sc = painel["situacao_contrato"]
    fin = painel["situacao_financeira"]
    vta = painel["vta"]
    retro = painel["retroativo"]
    m1, m2, m3, m4 = st.columns(4)
    rotulo_valor = (
        f"Valor do contrato hoje ({sc.get('ciclo_referencia')})"
        if sc.get("ciclo_referencia")
        else "Valor do contrato hoje (resumo oficial)"
        if str(sc.get("fonte_valor") or "").startswith("resumo_oficial")
        else "Valor do contrato hoje"
    )
    m1.metric(rotulo_valor, _brl(sc.get("valor_hoje")))
    m2.metric("VTA (oficial)", _brl(vta.get("oficial")))
    m3.metric("Retroativo apurado", _brl(retro.get("total_apurado")))
    m4.metric("Delta pago × devido", _brl(fin["totais"].get("delta")))

    # --- Conferencia: tudo abaixo responde "como esse calculo foi produzido?" ---
    with st.expander(
        "Como esse cálculo foi produzido? — Painel Executivo detalhado",
        expanded=False,
    ):
        col_esq, col_dir = st.columns([1, 1.35])

        # --- Situacao do contrato ---
        with col_esq:
            with st.container(border=True):
                st.markdown("**Situação do contrato**")
                st.write(f"Ciclo vigente (marco): **{sc.get('ciclo_vigente') or '—'}**")
                st.write(f"Modo de leitura: **{sc.get('modo') or '—'}**")
                st.write(f"Data de corte: **{_data_br(sc.get('data_corte'))}**")
                st.write(f"Versão do Masterfile: **{sc.get('versao_masterfile') or '—'}**")
                st.caption(
                    f"Estado Contratual: {sc.get('eventos_processados', 0)} eventos "
                    f"processados no marco (Event Log sombra)."
                )

        # --- Situacao do reajuste ---
        with col_dir:
            with st.container(border=True):
                st.markdown("**Situação do reajuste**")
                sr = painel["situacao_reajuste"]
                df_ciclos = pd.DataFrame([{
                    "Ciclo": c.get("ciclo"),
                    "Interregno": f"{_data_br(c.get('data_inicio'))} a {_data_br(c.get('data_fim'))}",
                    "Índice": _pct(c.get("indice_percentual")),
                    "Fator acum.": c.get("fator_acumulado") if c.get("fator_acumulado") is not None else "—",
                    "Nesta apuração": "Sim" if c.get("computa_nesta_apuracao") else "Não",
                    "Reajuste aplicável": "Sim" if c.get("reajuste_aplicavel") else "Não",
                } for c in sr["ciclos"]])
                st.dataframe(df_ciclos, use_container_width=True, hide_index=True)

        # --- Situacao financeira ---
        with st.container(border=True):
            st.markdown("**Situação financeira**")
            tot = fin["totais"]
            f1, f2, f3, f4 = st.columns(4)
            f1.metric("Executado (PCs)", _brl(tot.get("valor_pc")))
            f2.metric("Devido (com reajuste)", _brl(tot.get("valor_devido")))
            f3.metric("Pago", _brl(tot.get("valor_pago")))
            f4.metric("Delta (devido − pago)", _brl(tot.get("delta")))
            rz = fin["resumo_oficial"]
            st.caption(
                f"Resumo oficial do Masterfile — VTA: {_brl(rz.get('valor_total_atualizado'))} · "
                f"Retroativo: {_brl(rz.get('retroativo'))} · "
                f"Execução atualizada: {_brl(rz.get('execucao_atualizada'))} · "
                f"Saldo remanescente: {_brl(rz.get('saldo_remanescente'))}."
            )

        # --- Situacao dos Pedidos de Compra ---
        with st.container(border=True):
            st.markdown("**Pedidos de Compra**")
            sp = painel["situacao_pcs"]
            p1, p2, p3, p4 = st.columns(4)
            p1.metric("PCs lidos", sp["total"])
            p2.metric("Pagos", sp["pagos"])
            p3.metric("Divergência de ciclo", sp["com_divergencia_ciclo"])
            p4.metric("Sem data", sp["sem_data"])
            if sp["pcs"]:
                df_pcs = pd.DataFrame([{
                    "PC": p.get("numero_pc"),
                    "Data": _data_br(p.get("data_pc")),
                    "Ciclo (temporal)": p.get("ciclo_temporal") or "?",
                    "Ciclo informado": p.get("ciclo_informado") or "—",
                    "Valor": _brl(p.get("valor_pc")),
                    "Devido": _brl(p.get("valor_devido")),
                    "Pago": _brl(p.get("valor_pago")),
                    "Delta": _brl(p.get("delta")),
                    "Natureza": p.get("natureza_delta"),
                    "Retroativo": _brl(p.get("retroativo")),
                } for p in sp["pcs"]])
                st.dataframe(df_pcs, use_container_width=True, hide_index=True)
            else:
                st.caption("Nenhum Pedido de Compra lido neste Masterfile.")

        # --- Posicao contratual + VTA/retroativo ---
        col_pos, col_vta = st.columns([1.35, 1])
        with col_pos:
            with st.container(border=True):
                st.markdown("**Posição contratual (saldo por ciclo)**")
                pc = painel["posicao_contratual"]
                df_saldo = pd.DataFrame([{
                    "Ciclo": s.get("ciclo"),
                    "Valor vigente": _brl(s.get("valor_vigente")),
                    "Executado": _brl(s.get("executado")),
                    "Saldo": _brl(s.get("saldo")),
                } for s in pc["saldo_por_ciclo"]])
                st.dataframe(df_saldo, use_container_width=True, hide_index=True)
                ad = pc.get("adaptador") or {}
                st.caption(
                    f"Fonte: motor de posição contratual (sombra) — "
                    f"{ad.get('itens', 0)} itens, {ad.get('movimentos', 0)} movimentos, "
                    f"{pc.get('bloqueantes', 0)} alerta(s) bloqueante(s)."
                )

        with col_vta:
            with st.container(border=True):
                st.markdown("**Conferências históricas e retroativo**")
                st.write(f"Valor histórico do arquivo (não é VTA): **{_brl(vta.get('oficial'))}**")
                st.write(f"Conferência pela metodologia sombra: **{_brl(vta.get('metodologia_sombra'))}**")
                st.write(f"Conferência sombra integral: **{_brl(vta.get('metodologia_sombra_integral'))}**")
                st.write(f"Diferença entre conferências: **{_brl(vta.get('diferenca_sombra'))}**")
                st.write(f"Retroativo apurado (PCs): **{_brl(retro.get('total_apurado'))}**")
                st.write(f"Retroativo no resumo oficial: **{_brl(retro.get('oficial_resumo'))}**")
                if retro["por_pc"]:
                    st.markdown("**Retroativo por PC**")
                    df_retro = pd.DataFrame([{
                        "PC": r.get("numero_pc"),
                        "Ciclo": r.get("ciclo"),
                        "Retroativo": _brl(r.get("retroativo")),
                        "Natureza": r.get("natureza"),
                    } for r in retro["por_pc"]])
                    st.dataframe(df_retro, use_container_width=True, hide_index=True)

        # --- Alertas ---
        with st.container(border=True):
            st.markdown(f"**Pendências e alertas** — {resumo_niveis}")
            if painel["alertas"]:
                for a in painel["alertas"]:
                    nivel = str(a.get("nivel") or "INFO").upper()
                    texto = str(a.get("mensagem_negocio") or a.get("mensagem") or "")
                    if a.get("identificador"):
                        texto += f" (PC: {a['identificador']})"
                    if nivel == "ERRO_GRAVE":
                        st.error(texto)
                    elif nivel == "ALERTA":
                        st.warning(texto)
                    else:
                        st.caption(texto)
            else:
                st.caption("Nenhum alerta.")
