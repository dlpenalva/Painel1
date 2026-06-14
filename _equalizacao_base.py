"""
_equalizacao_base.py
--------------------
Equalização da Base — uso interno da GCC (cl8us 2.0).

Responsabilidade:
    Exibir, após o upload da Coleta Única, perguntas condicionais
    para que a GCC interprete os dados fornecidos pela área responsável/fiscal.
    NÃO altera cálculo, NÃO recalcula índice, NÃO substitui o processamento.

Uso em 15_Analise_Contratual.py:
    from _equalizacao_base import render_equalizacao_base

    respostas = render_equalizacao_base(dados_coleta)
    # respostas fica em st.session_state["equalizacao_base"]

Retorno (dict em session_state["equalizacao_base"]):
    financeiro_final        str  — Sim / Não / Não sei / N/A
    competencia_final       str  — mm/aaaa ou "Não sei" ou "N/A"
    corte_saldo             str  — Sim / Não / Não sei / N/A
    data_corte_saldo        str  — dd/mm/aaaa ou mm/aaaa ou ""
    aditivos_no_saldo       str  — Sim / Não / Não sei / N/A
    tratamento_aditivos     str  — texto da opção ou "N/A"
    preco_ciclo_cruzado     str  — Sim / Não / Não sei / N/A
    historico_ciclos_form   str  — texto da opção ou "N/A"
    equalizado              bool — True quando GCC confirmar
"""

import streamlit as st


# ──────────────────────────────────────────
# Constante de chave no session_state
# ──────────────────────────────────────────
CHAVE = "equalizacao_base"


def _inicializar():
    """Garante estrutura padrão no session_state."""
    padrao = {
        "financeiro_final":      "N/A",
        "competencia_final":     "N/A",
        "corte_saldo":           "N/A",
        "data_corte_saldo":      "",
        "aditivos_no_saldo":     "N/A",
        "tratamento_aditivos":   "N/A",
        "preco_ciclo_cruzado":   "N/A",
        "historico_ciclos_form": "N/A",
        "equalizado":            False,
    }
    if CHAVE not in st.session_state:
        st.session_state[CHAVE] = padrao.copy()


def _salvar(campo, valor):
    st.session_state[CHAVE][campo] = valor


def _get(campo):
    return st.session_state[CHAVE].get(campo, "N/A")


# ──────────────────────────────────────────
# Bloco público principal
# ──────────────────────────────────────────

def render_equalizacao_base(dados_coleta: dict) -> dict:
    """
    Renderiza o bloco de Equalização da Base na página de Análise e Documentos.

    Parâmetros:
        dados_coleta: dict retornado por ler_coleta_unica() ou ler_coleta_mestre().

    Retorna:
        dict com as respostas (mesmo que já esteja em session_state).
    """
    _inicializar()

    # Detecção robusta: suporta leitor Coleta Única e leitor ColetaMestre v10
    _fin = dados_coleta.get("financeiro", {})
    _its = dados_coleta.get("itens", {})
    _linhas_fin = (
        _fin.get("linhas_preenchidas", 0)
        or _fin.get("linhas_mensais_preenchidas", 0)
    )
    _total_fin = (
        _fin.get("total", 0)
        or _fin.get("total_pago", 0)
        or _fin.get("total_com_efeito", 0)
    )
    tem_financeiro = _linhas_fin > 0 or abs(float(_total_fin or 0)) > 0.01
    _n_its = (
        _its.get("itens_cadastrados", 0)
        or _its.get("linhas_preenchidas", 0)
    )
    tem_itens     = _n_its > 0
    tem_rem_atual = (
        _its.get("itens_com_rem_atual", 0) > 0
        or _its.get("itens_com_rem_corte", 0) > 0
    )
    tem_aditivos  = len(dados_coleta.get("aditivos", [])) > 0
    tem_ciclos    = len(dados_coleta.get("ciclos", [])) > 0

    st.markdown("---")
    st.subheader("🔎 Equalização da Base")
    st.caption(
        "Uso interno da GCC. As respostas orientam o diagnóstico e o relatório síntese. "
        "Não alteram o cálculo de reajuste já processado."
    )

    perguntas_respondidas = 0
    perguntas_aplicaveis  = 0

    # ── Pergunta 1 — Financeiro final ──────────────────────────────
    if tem_financeiro:
        perguntas_aplicaveis += 1
        st.markdown("**1. Os valores financeiros informados representam valores finais reconhecidos para pagamento por competência?**")
        opcoes_p1 = ["Sim", "Não", "Não sei"]
        atual_p1  = _get("financeiro_final")
        idx_p1    = opcoes_p1.index(atual_p1) if atual_p1 in opcoes_p1 else 0
        resp_p1   = st.radio(
            "financeiro_final",
            opcoes_p1,
            index=idx_p1,
            horizontal=True,
            label_visibility="collapsed",
            key="eq_financeiro_final",
        )
        _salvar("financeiro_final", resp_p1)
        if resp_p1 != "Não sei":
            perguntas_respondidas += 1

    # ── Pergunta 2 — Competência final do financeiro ───────────────
    if tem_financeiro:
        perguntas_aplicaveis += 1
        st.markdown("**2. Até qual competência/mês há informação financeira?**")
        col_p2a, col_p2b = st.columns([2, 1])
        with col_p2a:
            atual_comp = _get("competencia_final")
            val_comp   = "" if atual_comp in ("N/A", "Não sei") else atual_comp
            resp_comp  = st.text_input(
                "Competência (mm/aaaa)",
                value=val_comp,
                placeholder="Ex.: 06/2025",
                key="eq_competencia_final",
            )
        with col_p2b:
            nao_sei_p2 = st.checkbox(
                "Não sei",
                value=(atual_comp == "Não sei"),
                key="eq_competencia_nao_sei",
            )
        if nao_sei_p2:
            _salvar("competencia_final", "Não sei")
        elif resp_comp.strip():
            _salvar("competencia_final", resp_comp.strip())
            perguntas_respondidas += 1
        else:
            _salvar("competencia_final", "")

    # ── Pergunta 3 — Corte do saldo remanescente ───────────────────
    if tem_itens and tem_rem_atual:
        perguntas_aplicaveis += 1
        st.markdown("**3. O saldo remanescente informado corresponde a uma data/competência de corte conhecida?**")
        opcoes_p3 = ["Sim", "Não", "Não sei"]
        atual_p3  = _get("corte_saldo")
        idx_p3    = opcoes_p3.index(atual_p3) if atual_p3 in opcoes_p3 else 2
        resp_p3   = st.radio(
            "corte_saldo",
            opcoes_p3,
            index=idx_p3,
            horizontal=True,
            label_visibility="collapsed",
            key="eq_corte_saldo",
        )
        _salvar("corte_saldo", resp_p3)
        if resp_p3 != "Não sei":
            perguntas_respondidas += 1

        if resp_p3 == "Sim":
            atual_dc  = _get("data_corte_saldo")
            resp_dc   = st.text_input(
                "Data ou competência de corte (dd/mm/aaaa ou mm/aaaa)",
                value=atual_dc if atual_dc not in ("N/A", "") else "",
                placeholder="Ex.: 30/06/2025 ou 06/2025",
                key="eq_data_corte_saldo",
            )
            _salvar("data_corte_saldo", resp_dc.strip())
        else:
            _salvar("data_corte_saldo", "")

    # ── Pergunta 4 — Aditivos no saldo ────────────────────────────
    if tem_itens and tem_rem_atual:
        perguntas_aplicaveis += 1
        st.markdown("**4. O saldo remanescente informado já considera aditivos e supressões?**")
        opcoes_p4 = ["Sim", "Não", "Não sei"]
        atual_p4  = _get("aditivos_no_saldo")
        idx_p4    = opcoes_p4.index(atual_p4) if atual_p4 in opcoes_p4 else 2
        resp_p4   = st.radio(
            "aditivos_no_saldo",
            opcoes_p4,
            index=idx_p4,
            horizontal=True,
            label_visibility="collapsed",
            key="eq_aditivos_no_saldo",
        )
        _salvar("aditivos_no_saldo", resp_p4)
        if resp_p4 != "Não sei":
            perguntas_respondidas += 1

    # ── Pergunta 5 — Tratamento dos aditivos ──────────────────────
    if tem_aditivos:
        perguntas_aplicaveis += 1
        st.markdown("**5. Como os aditivos/supressões informados devem ser tratados no Valor Total Atualizado?**")
        opcoes_p5 = [
            "Já estão refletidos no financeiro/itens — não somar novamente",
            "Devem ser computados à parte no VTA",
            "Apenas informativos",
            "Não sei",
        ]
        atual_p5 = _get("tratamento_aditivos")
        idx_p5   = opcoes_p5.index(atual_p5) if atual_p5 in opcoes_p5 else 3
        resp_p5  = st.radio(
            "tratamento_aditivos",
            opcoes_p5,
            index=idx_p5,
            label_visibility="collapsed",
            key="eq_tratamento_aditivos",
        )
        _salvar("tratamento_aditivos", resp_p5)
        if resp_p5 != "Não sei":
            perguntas_respondidas += 1

    # ── Pergunta 6 — Preço formado num ciclo, executado em outro ──
    if tem_financeiro or tem_itens:
        perguntas_aplicaveis += 1
        st.markdown(
            "**6. Há valores formados em uma competência/ciclo, mas executados, "
            "medidos ou atestados em competência/ciclo posterior?**"
        )
        opcoes_p6 = ["Sim", "Não", "Não sei"]
        atual_p6  = _get("preco_ciclo_cruzado")
        idx_p6    = opcoes_p6.index(atual_p6) if atual_p6 in opcoes_p6 else 2
        resp_p6   = st.radio(
            "preco_ciclo_cruzado",
            opcoes_p6,
            index=idx_p6,
            horizontal=True,
            label_visibility="collapsed",
            key="eq_preco_ciclo_cruzado",
        )
        _salvar("preco_ciclo_cruzado", resp_p6)
        if resp_p6 != "Não sei":
            perguntas_respondidas += 1

    # ── Pergunta 7 — Forma do histórico de ciclos anteriores ──────
    if tem_financeiro or tem_itens or tem_ciclos:
        perguntas_aplicaveis += 1
        st.markdown("**7. O histórico de ciclos anteriores foi informado de que forma?**")
        opcoes_p7 = [
            "Mês a mês",
            "Consolidado por ciclo",
            "Apenas saldo atual/remanescente",
            "Não informado",
            "Não sei",
        ]
        atual_p7 = _get("historico_ciclos_form")
        idx_p7   = opcoes_p7.index(atual_p7) if atual_p7 in opcoes_p7 else 4
        resp_p7  = st.radio(
            "historico_ciclos_form",
            opcoes_p7,
            index=idx_p7,
            label_visibility="collapsed",
            key="eq_historico_ciclos",
        )
        _salvar("historico_ciclos_form", resp_p7)
        if resp_p7 != "Não sei":
            perguntas_respondidas += 1

    # ── Barra de progresso e botão de confirmação ─────────────────
    st.markdown("---")
    if perguntas_aplicaveis > 0:
        pct = perguntas_respondidas / perguntas_aplicaveis
        st.progress(pct, text=f"Equalização: {perguntas_respondidas} de {perguntas_aplicaveis} perguntas respondidas")

    col_btn, col_status = st.columns([1, 3])
    with col_btn:
        if st.button("✅ Confirmar Equalização", key="eq_btn_confirmar", use_container_width=True):
            st.session_state[CHAVE]["equalizado"] = True
            st.success("Equalização confirmada. Prossiga com o processamento.")

    if st.session_state[CHAVE].get("equalizado"):
        with col_status:
            st.info("Base equalizada pela GCC.", icon="✔️")

    return st.session_state[CHAVE]


# ──────────────────────────────────────────
# Diagnóstico resumido (Etapa 2)
# ──────────────────────────────────────────

def render_diagnostico_equalizacao(dados_coleta: dict) -> None:
    """
    Exibe resumo objetivo das respostas e gaps detectados.
    Chamar APÓS render_equalizacao_base().
    """
    _inicializar()
    eq = st.session_state.get(CHAVE, {})

    if not eq.get("equalizado"):
        return  # Só exibe após confirmação

    st.markdown("---")
    st.subheader("📋 Diagnóstico da Base")

    fin  = dados_coleta.get("financeiro", {})
    itens = dados_coleta.get("itens", {})

    tem_financeiro = fin.get("linhas_preenchidas", 0) > 0
    tem_itens      = itens.get("itens_cadastrados", 0) > 0
    tem_rem_atual  = itens.get("itens_com_rem_atual", 0) > 0
    tem_aditivos   = len(dados_coleta.get("aditivos", [])) > 0

    def _linha(label, valor, ok=True):
        icone = "✅" if ok else "⚠️"
        st.markdown(f"{icone} **{label}:** {valor}")

    _linha("Base financeira identificada",
           f"Sim — {fin.get('linhas_preenchidas', 0)} competências, total R$ {fin.get('total', 0):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if tem_financeiro else "Não",
           ok=tem_financeiro)

    _linha("Base de itens identificada",
           f"Sim — {itens.get('itens_cadastrados', 0)} itens" if tem_itens else "Não",
           ok=tem_itens)

    _linha("Remanescente atual informado",
           f"Sim — {itens.get('itens_com_rem_atual', 0)} itens com saldo" if tem_rem_atual else "Não",
           ok=tem_rem_atual)

    _linha("Aditivos/supressões identificados",
           f"Sim — {len(dados_coleta.get('aditivos', []))} registros" if tem_aditivos else "Não informados",
           ok=True)

    comp_final = eq.get("competencia_final", "")
    _linha("Competência final informada",
           comp_final if comp_final not in ("", "N/A", "Não sei") else "Não informada",
           ok=comp_final not in ("", "N/A", "Não sei"))

    data_corte = eq.get("data_corte_saldo", "")
    _linha("Data de corte do saldo",
           data_corte if data_corte else "Não informada",
           ok=bool(data_corte))

    # Gaps
    gaps = []
    if not tem_financeiro:
        gaps.append("Sem financeiro histórico — retroativo financeiro definitivo não calculável.")
    if not tem_rem_atual:
        gaps.append("Sem remanescente atual — saldo remanescente não estimável.")
    if comp_final in ("", "N/A", "Não sei"):
        gaps.append("Competência final do financeiro não informada.")
    if not data_corte and tem_rem_atual:
        gaps.append("Data de corte do saldo remanescente não informada.")
    if eq.get("aditivos_no_saldo") == "Não sei" and tem_aditivos:
        gaps.append("Dúvida sobre aditivos já incorporados no saldo.")
    if eq.get("historico_ciclos_form") in ("Não informado", "Não sei", "N/A"):
        gaps.append("Histórico de ciclos anteriores não informado ou incerto.")
    if eq.get("preco_ciclo_cruzado") == "Sim":
        gaps.append("Valores formados em um ciclo e executados em outro — verificar corte correto.")

    if gaps:
        st.markdown("---")
        st.markdown("**⚠️ Gaps e pendências detectados:**")
        for g in gaps:
            st.markdown(f"- {g}")
    else:
        st.success("Nenhum gap crítico detectado na base.")
