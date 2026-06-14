"""
_relatorio_sintese.py
---------------------
Relatório Síntese da Apuração — cl8us FARC.

Uso em 15_Analise_Contratual.py:
    from _relatorio_sintese import render_relatorio_sintese
    render_relatorio_sintese(res, _diag, equalizacao)
"""

import streamlit as st


# ── Formatação ────────────────────────────────────────────────────

def _moeda(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"

def _pct(v):
    """Formata percentuais em padrão brasileiro com duas casas: xx,xx%."""
    try:
        if v is None or str(v).strip() == "":
            return "0,00%"
        s = str(v).replace("%", "").replace(".", "").replace(",", ".") if "," in str(v) else str(v).replace("%", "")
        n = float(s)
        # Se vier em fração, converte para percentual. Ex.: 0.015778 -> 1.5778%
        if abs(n) <= 1:
            n *= 100
        return f"{n:.2f}%".replace(".", ",")
    except Exception:
        return "0,00%"

def _fator(v):
    """Formata fator acumulado com 6 casas decimais em padrão brasileiro."""
    try:
        if v is None or str(v).strip() == "":
            return "1,000000"
        n = float(str(v).replace(",", "."))
        return f"{_fator(n)}".replace(".", ",")
    except Exception:
        return "1,000000"


def _txt_gap(gaps):
    if not gaps:
        return ""
    linhas = "\n".join(f"- {g}" for g in gaps)
    return f"\n\nGAPS E PENDÊNCIAS DETECTADOS:\n{linhas}"


# ── Montar texto do relatório ─────────────────────────────────────

def _montar_texto(res, diag, eq):
    r = res or {}
    d = diag or {}
    eq = eq or {}

    val_total   = float(r.get("valor_atualizado_contrato", r.get("valor_global_estoque", 0)) or 0)
    val_retro   = float(r.get("valor_represado_a_pagar", r.get("valor_retroativo_consumo", 0)) or 0)
    fator       = float(r.get("fator_acumulado", 1.0) or 1.0)
    variacao    = float(r.get("variacao_acumulada", fator - 1.0) or fator - 1.0)
    indice      = str(r.get("indice_contratual", r.get("indice", "")) or "")
    modo        = str(r.get("modo_detectado", r.get("modo_preliminar", "")) or "")

    ciclos_res  = r.get("ciclos", []) or []
    ciclos_diag = d.get("ciclos", []) or []
    ciclos      = ciclos_res if ciclos_res else ciclos_diag

    n_ciclos    = len(ciclos)
    n_efeito    = sum(1 for c in ciclos if str(c.get("efeito","")).lower() == "sim"
                      or str(c.get("tem_efeito_fin","")).lower() == "sim"
                      or c.get("objeto_analise_atual"))
    n_precluso  = sum(1 for c in ciclos if "precluso" in str(c.get("situacao","")).lower()
                      or "precluso" in str(c.get("situacao_aplicada","")).lower())

    # Gaps da equalização
    gaps = eq.get("gaps", []) if isinstance(eq, dict) else []
    premissas = eq.get("premissas", []) if isinstance(eq, dict) else []

    # Competência final
    comp_final = eq.get("competencia_final","") if isinstance(eq,dict) else ""
    comp_final_txt = f" até a competência {comp_final}" if comp_final and comp_final not in ("","N/A","Não sei") else ""

    # Data de corte do saldo
    data_corte = eq.get("data_corte_saldo","") if isinstance(eq,dict) else ""
    corte_txt = f", com data de corte em {data_corte}" if data_corte else ""

    linhas = []
    linhas.append("RELATÓRIO SÍNTESE DA APURAÇÃO")
    linhas.append("=" * 60)
    linhas.append("")

    # 1. Síntese dos ciclos
    linhas.append("1. SÍNTESE DA ANÁLISE")
    linhas.append(f"Foram analisados {n_ciclos} ciclo(s) de reajuste contratual.")
    if n_efeito:
        linhas.append(f"Ciclos com efeitos financeiros identificados: {n_efeito}.")
    if n_precluso:
        linhas.append(f"Ciclos preclusos (sem geração de novos efeitos): {n_precluso}.")
    if indice:
        linhas.append(f"Índice contratual aplicado: {indice}.")
    if variacao:
        linhas.append(f"Variação acumulada apurada: {_pct(variacao)}.")
    if modo:
        linhas.append(f"Base de apuração: {modo}.")
    linhas.append("")

    # 2. Ciclos individuais
    if ciclos:
        linhas.append("2. CICLOS")
        for c in ciclos:
            nome = c.get("ciclo","")
            sit  = (c.get("situacao_aplicada") or c.get("situacao") or "").strip()
            pct  = c.get("percentual_apurado", c.get("percentual", 0)) or 0
            fat  = c.get("fator_acumulado", 1.0) or 1.0
            if "precluso" in sit.lower():
                linhas.append(f"• {nome}: classificado como precluso — sem geração de novos efeitos financeiros.")
            elif float(pct) > 0:
                linhas.append(f"• {nome}: reajuste de {_pct(pct)} (fator acumulado {float(fat):.6f}).")
            else:
                linhas.append(f"• {nome}: {sit if sit else 'sem percentual apurado'}.")
        linhas.append("")

    # 3. Retroativo
    if val_retro and abs(val_retro) > 0.01:
        linhas.append("3. VALOR RETROATIVO")
        linhas.append(f"O retroativo apurado totaliza {_moeda(val_retro)}{comp_final_txt}.")
        linhas.append("")

    # 4. Valor Total Atualizado
    linhas.append("4. VALOR TOTAL ATUALIZADO DO CONTRATO")
    if val_total and abs(val_total) > 0.01:
        linhas.append(f"O Valor Total Atualizado do Contrato foi apurado em {_moeda(val_total)}{corte_txt}.")
    else:
        linhas.append("O Valor Total Atualizado do Contrato não foi calculado nesta apuração.")
    linhas.append("")

    # 5. Metodologia
    linhas.append("5. METODOLOGIA")
    linhas.append(
        "O valor total atualizado corresponde à soma da execução financeira informada "
        "até a última competência disponível com o saldo remanescente atualizado a partir "
        "da competência seguinte, calculado com base nos itens ainda passíveis de execução "
        "até o fim da vigência, multiplicados pelo fator acumulado do índice contratual."
    )
    linhas.append("")

    # 6. Premissas assumidas
    if premissas:
        linhas.append("6. PREMISSAS ASSUMIDAS PELA GCC")
        for p in premissas:
            linhas.append(f"• {p}")
        linhas.append("")

    # 7. Gaps e pedido de validação
    linhas.append("7. GAPS E PEDIDO DE VALIDAÇÃO")
    if gaps:
        for g in gaps:
            linhas.append(f"• {g}")
    else:
        linhas.append("Nenhum gap crítico foi detectado na base de dados utilizada.")
    linhas.append("")
    linhas.append(
        "Solicita-se validação da área responsável quanto à suficiência da base "
        "informada, especialmente em relação à última competência financeira considerada, "
        "ao saldo remanescente utilizado e ao tratamento de eventuais aditivos e supressões."
    )
    linhas.append("")
    linhas.append("=" * 60)
    linhas.append("Gerado automaticamente pelo sistema cl8us — uso interno GCC.")

    return "\n".join(linhas)


# ── Render principal ──────────────────────────────────────────────

def render_relatorio_sintese(res, diag=None, equalizacao=None):
    """
    Renderiza o bloco Relatório Síntese da Apuração.

    Parâmetros:
        res:          dict — resultado do processamento (retorno de calcular_resultado_v10)
        diag:         dict — dados da coleta (retorno do leitor)
        equalizacao:  dict — respostas da equalização (session_state["equalizacao_base"])
    """
    if not res:
        return

    # Tentar carregar equalização do session_state se não fornecida
    if equalizacao is None:
        try:
            from _equalizacao_base import CHAVE
            equalizacao = st.session_state.get(CHAVE, {})
        except Exception:
            equalizacao = {}

    # Tentar carregar classificação da base
    classificacao_base = {}
    try:
        from _middle_layer_coleta import classificar_base_equalizacao
        classificacao_base = classificar_base_equalizacao(diag, equalizacao)
        equalizacao = {**equalizacao, **classificacao_base}
    except Exception:
        pass

    st.divider()
    st.subheader("📄 Relatório Síntese da Apuração")
    st.caption("Uso interno da GCC. Pode ser enviado à área responsável para validação final.")

    texto = _montar_texto(res, diag, equalizacao)

    # Preview expandível
    with st.expander("👁 Visualizar relatório", expanded=False):
        st.text(texto)

    # Classificação da base (se disponível)
    if classificacao_base:
        col1, col2 = st.columns(2)
        with col1:
            cls = classificacao_base.get("classificacao","")
            if cls:
                cor = {"Base completa":"green","Base híbrida":"blue",
                       "Base financeira":"blue","Base itemizada":"orange",
                       "Base insuficiente":"red"}.get(cls,"gray")
                st.markdown(f"**Classificação da base:** :{cor}[{cls}]")
        with col2:
            conf = classificacao_base.get("confianca","")
            if conf:
                cor2 = {"Alta":"green","Média":"orange","Baixa":"red"}.get(conf,"gray")
                st.markdown(f"**Confiança:** :{cor2}[{conf}]")

    # Download .txt
    st.download_button(
        label="⬇ Baixar Relatório Síntese (.txt)",
        data=texto.encode("utf-8"),
        file_name="Relatorio_Sintese_Apuracao.txt",
        mime="text/plain",
        key="btn_relatorio_sintese",
    )

# >>> PATCH_RELATORIO_SINTESE_ROBUSTO_M21_V4
# Fallback para impedir que o bloco do Relatório Síntese desapareça quando
# algum objeto vier como string/lista em vez de dict.
def _m21v4_sintese_num(v, default=0.0):
    try:
        if v is None or isinstance(v, bool):
            return default
        if isinstance(v, (int, float)):
            return float(v)
        s = str(v).strip().replace("R$", "").replace(" ", "")
        if not s:
            return default
        if "," in s:
            s = s.replace(".", "").replace(",", ".")
        return float(s)
    except Exception:
        return default


def _m21v4_sintese_moeda(v):
    n = _m21v4_sintese_num(v, 0.0)
    return f"R$ {n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _m21v4_sintese_texto_fallback(res, diag=None, equalizacao=None, erro=None):
    res = res if isinstance(res, dict) else {}
    diag = diag if isinstance(diag, dict) else {}
    linhas = []
    linhas.append("RELATÓRIO SÍNTESE DA APURAÇÃO")
    linhas.append("=" * 60)
    linhas.append("")
    linhas.append("1. SÍNTESE DA ANÁLISE")
    linhas.append(f"Tipo de coleta: {res.get('tipo', res.get('origem_coleta', 'Matriz 2.1'))}.")
    linhas.append(f"Índice aplicado: {res.get('indice', res.get('indice_aplicado', 'Não informado'))}.")
    linhas.append(f"Valor Total Atualizado: {_m21v4_sintese_moeda(res.get('valor_atualizado_contrato', res.get('valor_global_estoque', 0)))}.")
    linhas.append(f"Retroativo/valor represado: {_m21v4_sintese_moeda(res.get('valor_represado_a_pagar', res.get('valor_retroativo', 0)))}.")
    linhas.append(f"Aditivos/supressões: {_m21v4_sintese_moeda(res.get('total_aditivos_supressoes', res.get('valor_total_aditivos_supressoes', 0)))}.")
    linhas.append("")
    linhas.append("2. OBSERVAÇÃO")
    linhas.append("Relatório gerado em modo robusto para preservar o download documental da Matriz 2.1.")
    if erro:
        linhas.append(f"Aviso técnico interno: {erro}")
    return "\n".join(linhas)


try:
    _render_relatorio_sintese_original_v4 = render_relatorio_sintese

    def render_relatorio_sintese(res, diag=None, equalizacao=None):
        try:
            if isinstance(res, str):
                res = {"tipo": res}
            if not isinstance(res, dict):
                res = {}
            if isinstance(diag, str) or diag is None:
                diag = {}
            if not isinstance(diag, dict):
                diag = {}
            if isinstance(equalizacao, str) or equalizacao is None:
                equalizacao = {}
            if not isinstance(equalizacao, dict):
                equalizacao = {}
            return _render_relatorio_sintese_original_v4(res, diag, equalizacao)
        except Exception as exc:
            st.divider()
            st.subheader("📄 Relatório Síntese da Apuração")
            st.caption("Uso interno da GCC. Pode ser enviado à área responsável para validação final.")
            texto = _m21v4_sintese_texto_fallback(res, diag, equalizacao, exc)
            with st.expander("👁 Visualizar relatório", expanded=False):
                st.text(texto)
            st.download_button(
                label="⬇ Baixar Relatório Síntese (.txt)",
                data=texto.encode("utf-8"),
                file_name="Relatorio_Sintese_Apuracao.txt",
                mime="text/plain",
                key="btn_relatorio_sintese_m21_v4",
            )
except Exception:
    pass
# <<< PATCH_RELATORIO_SINTESE_ROBUSTO_M21_V4

