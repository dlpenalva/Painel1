"""Adequacao Orcamentaria — UX V3 (planilha guiada).

Rearquitetura de INTERACAO (Etapa UX V3). A matematica NAO muda: continua no
motor de dominio (_adequacao_orcamentaria) e nos adaptadores de view-model
(_adequacao_ui), ambos ja homologados (golden Financeiro/PC, projecao,
programacao, XLSX, ZERO x VAZIO, overrides, retroativo oficial).

A pagina deixa de ser linear e passa a reproduzir mentalmente abas de uma
planilha:

    1. Base       -> dados de entrada (retroativo, percentual, vigencia)
    2. Historico  -> formacao da media (Financeiro ou Pedidos de Compra)
    3. Projecao   -> meses futuros (vazio=media, 0=zero, valor=override)
    4. Resultado  -> complementacao necessaria + programacao + XLSX

O que era controle tecnico (fator 1.xxxxx, override, janela, exclusoes por
multiselect, metodologia longa) saiu do fluxo principal para "Opcoes avancadas".
O estado da nova interface vive em namespace proprio (adequacao_v3_*).
"""
from datetime import datetime
from calendar import monthrange
from zoneinfo import ZoneInfo
from html import escape

import pandas as pd
import streamlit as st

from _ui_utils import render_cabecalho_pagina
# A matematica vive no motor de dominio. Os adaptadores de view-model (formatacao
# e preparacao de dados, delegando ao motor) vivem em _adequacao_ui.
from _adequacao_orcamentaria import (
    _as_date,
    janela_automatica_pcs,
    media_pedidos_compra,
    valor_original_foi_informado,
    pedidos_de_itens_pc,
    classificar_pedidos,
    media_financeiro,
    inferir_cadencia,
    projetar_por_cadencia,
    projetar_por_ciclo_proporcional,
    cadencia_por_ciclo_forcada,
    base_prudencial_irregular,
    avaliar_evidencia_media_irregular_pc,
    CADENCIA_MENSAL,
    CADENCIA_IRREGULAR,
    CADENCIA_POR_CICLO,
)
from _adequacao_ui import (
    _round2,
    moeda,
    parse_moeda_br,
    pct,
    percentual_e_fator_da_adequacao,
    texto_seguro,
    extrair_contexto_valores,
    carregar_itens_pc_da_sessao,
    financeiro_por_competencia,
    janela_6_competencias,
    valores_informados_da_janela,
    periodo_para_label,
    normalizar_competencia,
    gerar_periodos_projecao,
    montar_base_editor,
    calcular_projecao,
    cronograma_por_exercicio,
    gerar_xlsx_projecao,
    situacao_financeira_considerada,
    atualizar_exclusoes_manuais_pc,
    ciclos_para_cadencia,
    pares_de_financeiro,
    pares_de_pedidos,
    ROTULO_MEDIA_IRREGULAR,
    NOTA_RESUMO_ESTIMADA,
    NOTA_RESUMO_SALDO,
    NOTA_RESUMO_PARCIAL,
    NAO_DETERMINADA,
    validar_vigencia_para_adequacao,
    explicar_meses_de_movimento_por_ciclo,
    explicar_resultado_adequacao,
    saldo_elegivel_da_apuracao,
    consolidar_resultado_adequacao,
    linhas_resumo_resultado,
)

st.set_page_config(page_icon="assets/cl8us_favicon_512.png", page_title="TLB · cl8us - Adequação Orçamentária", layout="wide")


# ---------------------------------------------------------------- render helpers

def data_hora_brasilia():
    return datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")


def render_card_valor(label, valor, nota="", destaque=False, formato="moeda", alerta=False):
    if alerta:
        # Vermelho suave: comunica "não reconhecido/não incluído" sem
        # aparência de erro crítico (paleta distinta de st.error).
        bg, border, cor_valor = "#FDF2F2", "#F0B8B8", "#8A3A3A"
    else:
        bg = "#EAF2F8" if destaque else "#FFFFFF"
        border = "#9EC5E8" if destaque else "#E5EAF0"
        cor_valor = "#0B1F3A" if destaque else "#0F172A"
    fs = "1.65rem" if destaque else "1.12rem"
    nota_html = f"<div style='color:#64748B; font-size:0.82rem; margin-top:5px;'>{escape(str(nota))}</div>" if nota else ""
    if formato == "inteiro":
        try:
            valor_fmt = f"{int(round(float(valor or 0)))}"
        except Exception:
            valor_fmt = "0"
    elif formato == "texto":
        valor_fmt = escape(str(valor))
    else:
        valor_fmt = moeda(valor)
    st.markdown(
        f"""<div style="background:{bg}; border:1px solid {border}; border-radius:14px; padding:14px 16px; min-height:96px;">
            <div style="color:#475569; font-size:0.84rem; font-weight:700; margin-bottom:7px;">{escape(str(label))}</div>
            <div style="color:{cor_valor}; font-size:{fs}; font-weight:900; line-height:1.2; overflow-wrap:anywhere;">{valor_fmt}</div>
            {nota_html}</div>""",
        unsafe_allow_html=True,
    )


def render_leitura(itens):
    """Composicao compacta de leitura (rotulo -> valor), sem cinco cards grandes.
    itens: lista de (rotulo, valor_texto)."""
    linhas = "".join(
        f"<div style='display:flex; justify-content:space-between; gap:16px; "
        f"padding:7px 2px; border-bottom:1px solid #EEF2F6;'>"
        f"<span style='color:#475569; font-weight:700;'>{escape(str(r))}</span>"
        f"<span style='color:#0F172A; font-weight:800; text-align:right;'>{escape(str(v))}</span></div>"
        for r, v in itens
    )
    st.markdown(
        f"<div style='background:#FFFFFF; border:1px solid #E5EAF0; border-radius:14px; "
        f"padding:8px 16px;'>{linhas}</div>",
        unsafe_allow_html=True,
    )


def input_moeda(label, valor_padrao, key, help=None):
    txt = st.text_input(label, value=moeda(valor_padrao, com_prefixo=False), key=key, help=help)
    valor = parse_moeda_br(txt)
    st.caption(moeda(valor))
    return valor


def _metodologia_valor_importado_texto(resultado_valor_global):
    """Resumo (Opcoes avancadas) da metodologia do Valor Total Atualizado
    importado do modulo Valores. Nao altera calculo."""
    if not isinstance(resultado_valor_global, dict) or not resultado_valor_global:
        return ("Não há resultado do módulo Valores carregado nesta sessão. "
                "Os campos desta página devem ser conferidos/preenchidos manualmente.")
    cfg = resultado_valor_global.get("config_ciclo_em_execucao", {}) or {}
    corte = bool(resultado_valor_global.get("corte_operacional_aplicado")
                 or resultado_valor_global.get("corte_operacional_solicitado") or cfg.get("aplicar"))
    try:
        valor_total = moeda(resultado_valor_global.get("valor_atualizado_contrato", resultado_valor_global.get("valor_global_estoque", 0)))
    except Exception:
        valor_total = str(resultado_valor_global.get("valor_atualizado_contrato", "Não informado"))
    try:
        execucao = moeda(resultado_valor_global.get("valor_executado_atualizado", 0))
    except Exception:
        execucao = str(resultado_valor_global.get("valor_executado_atualizado", "Não informado"))
    try:
        remanescente = moeda(resultado_valor_global.get("remanescente_reajustado", 0))
    except Exception:
        remanescente = str(resultado_valor_global.get("remanescente_reajustado", "Não informado"))
    if corte:
        return (f"Metodologia: corte operacional no ciclo em execução. "
                f"Execução atualizada considerada: {execucao}. "
                f"Saldo remanescente atualizado: {remanescente}. "
                f"Valor Total Atualizado importado: {valor_total}.")
    return (f"Metodologia: corte padrão. Composição: execução atualizada por ciclo + "
            f"saldo remanescente atualizado. Execução: {execucao}. "
            f"Remanescente: {remanescente}. Valor Total Atualizado importado: {valor_total}.")


def _valor_contratual_automatico(resultado, adm, *chaves):
    """Busca o primeiro valor nao vazio entre as fontes ja usadas em outras
    paginas para identificacao contratual (mesmo padrao de
    extrair_apostilado em pages/13_DOU.py e do fallback contrato/
    numero_contrato em pages/04_Relatorio_Global.py). Nunca inventa: se
    nenhuma fonte tiver o dado, devolve string vazia para o campo manual
    decidir o placeholder."""
    fontes = []
    for base in (resultado, adm):
        if not isinstance(base, dict):
            continue
        fontes.append(base)
        contexto = base.get("contexto_contratual_anterior")
        if isinstance(contexto, dict):
            fontes.append(contexto)
    for fonte in fontes:
        for chave in chaves:
            texto = texto_seguro(fonte.get(chave), "")
            if texto:
                return texto
    return ""


def _data_siga(valor):
    d = _as_date(valor)
    return d.strftime("%d/%m/%Y") if d else "[campo a preencher]"


# ---------------------------------------------------------------- cabecalho + contexto

render_cabecalho_pagina("Adequação Orçamentária", "")
if st.button("← Voltar para Central", key="voltar_central_adequacao"):
    st.switch_page("pages/03_Valor_Global.py")

ctx = extrair_contexto_valores(st.session_state.get("resultado_valor_global", {}) or {})
resultado = ctx["resultado"]
diagnostico = st.session_state.get("diagnostico_coleta_v2")
modo_apuracao = resultado.get("modo_apuracao", "Completo") if isinstance(resultado, dict) else "Completo"
modo_reduzido_estoque = modo_apuracao == "Reduzido por Itens/Estoque"
modo_consumo_itens_ciclo = modo_apuracao == "Consumo por Itens/Ciclo"

# Medidas canonicas ja consolidadas pela apuracao (_resultado_consolidado).
# A Adequacao LE daqui e nao recalcula: retroativo reconhecido e potencial sao
# grandezas distintas produzidas la, e a data de corte tem precedencia unica
# (totais_canonicos_pc.data_corte -> controle.data_corte).
consolidado_ad = (resultado.get("resultado_consolidado") or {}) if isinstance(resultado, dict) else {}
medidas_pc = bool(consolidado_ad.get("medidas_pc_aplicaveis"))
data_corte_canonica = _as_date((consolidado_ad.get("fora_do_corte") or {}).get("data_corte"))
retroativo_potencial = consolidado_ad.get("retroativo_potencial") if medidas_pc else None
valor_em_analise_pc = consolidado_ad.get("valor_atualizado_em_analise") if medidas_pc else None
# Saldo contratual elegivel (Cl8us 12.20): limita a projecao futura, na mesma
# base economica dela (sem o reajuste). Nunca derivado do VTA; ausencia != zero.
saldo_info = saldo_elegivel_da_apuracao(resultado)
saldo_elegivel = saldo_info["valor"]

# Base financeira por competencia (ZERO x VAZIO preservados) e janela de 6
# competencias-calendario terminando na ultima competencia INFORMADA.
fin_por_comp, origem_financeira = financeiro_por_competencia(resultado)
ultimos_6 = janela_6_competencias(fin_por_comp, 6)
_vals_informados = valores_informados_da_janela(ultimos_6)
media_6 = media_financeiro(_vals_informados)["media_mensal"]
comp_informadas = len(_vals_informados)
comp_total_janela = len(ultimos_6)
comp_sem_info = comp_total_janela - comp_informadas
ultima_comp_fin = ultimos_6["_periodo"].iloc[-1] if not ultimos_6.empty else None
ultima_comp_fin_txt = periodo_para_label(ultima_comp_fin) if ultima_comp_fin is not None else "[campo a preencher]"

tem_apuracao = bool(ctx["disponivel"])
registros_pc_ctx = carregar_itens_pc_da_sessao(resultado, diagnostico)

# Fontes historicas efetivamente encontradas (nunca inventa).
tem_fin = comp_informadas > 0
tem_pc = len(registros_pc_ctx) > 0
if tem_fin and tem_pc:
    fontes_txt = "Financeiro + Pedidos de Compra"
    metodo_sugerido = "Financeiro"
elif tem_fin:
    fontes_txt = "Financeiro"
    metodo_sugerido = "Financeiro"
elif tem_pc:
    fontes_txt = "Pedidos de Compra"
    metodo_sugerido = "Pedidos de Compra"
else:
    fontes_txt = "Nenhuma"
    metodo_sugerido = "Financeiro"

st.caption("Confira a base, valide o histórico, ajuste a projeção e obtenha o valor da adequação.")

if modo_consumo_itens_ciclo:
    st.info("Modo Consumo por Itens/Ciclo: a base mensal por competência não foi informada. "
            "A adequação utilizará o Retroativo e deve ser tratada como estimativa apoiada na validação fiscal.")
elif modo_reduzido_estoque:
    st.info("Modo Reduzido por Itens/Estoque: a base mensal por competência não foi informada. "
            "A adequação será tratada como estimativa.")

tab_base, tab_hist, tab_proj, tab_result, tab_siga = st.tabs(
    ["1. Base", "2. Histórico", "3. Projeção", "4. Resultado", "5. Texto SIGA"])


# ================================================================ TAB 1 — BASE
with tab_base:
    st.subheader("1. Base da adequação")
    label_retroativo = ("Retroativo (itens consumidos/ciclo)" if modo_consumo_itens_ciclo
                        else ("Retroativo estimado por itens/estoque" if modo_reduzido_estoque
                              else ("Retroativo reconhecido" if medidas_pc
                                    else "Retroativo apurado")))
    linhas_base = [
        (label_retroativo, moeda(ctx["valor_represado"]) if tem_apuracao else "Não localizado"),
    ]
    # Metodo PC: reconhecido e potencial sao grandezas DIFERENTES e aparecem
    # separadas. Ausencia nao vira zero — potencial nao apurado sai como
    # "Não localizado"; potencial zero e exibido como zero, sem alarde.
    if medidas_pc:
        linhas_base.append((
            "Retroativo potencial",
            moeda(retroativo_potencial) if retroativo_potencial is not None else "Não localizado",
        ))
        if valor_em_analise_pc is not None:
            linhas_base.append(("Valor atualizado em análise", moeda(valor_em_analise_pc)))
    linhas_base.append(
        ("Percentual do reajuste", pct(ctx["variacao"]) if tem_apuracao else "Não localizado"))
    # Corte e vigencia sao conceitos distintos e aparecem lado a lado: o corte
    # fecha o HISTORICO; a vigencia (campo abaixo) fecha a PROJECAO.
    if data_corte_canonica is not None:
        linhas_base.append(
            ("Data de corte da apuração", data_corte_canonica.strftime("%d/%m/%Y")))
    linhas_base.extend([
        ("Última competência do histórico financeiro", ultima_comp_fin_txt),
        ("Fontes encontradas", fontes_txt),
        ("Saldo contratual elegível para a projeção",
         moeda(saldo_elegivel) if saldo_elegivel is not None else "Não disponível"),
    ])
    render_leitura(linhas_base)
    st.caption(f"Saldo elegível: {saldo_info['origem']}. Limita apenas a projeção futura "
               "(base sem reajuste); o retroativo reconhecido não é limitado pelo saldo."
               + (f" Saldo remanescente atualizado (conciliação): {moeda(saldo_info['atualizado'])}."
                  if saldo_info.get("atualizado") is not None else ""))
    if medidas_pc:
        # VTA-POT-1: a adequação NÃO parte do VTA — ela soma retroativo
        # reconhecido + diferença futura projetada. Por isso o potencial, que
        # agora integra o VTA por critério prudencial, continua entrando aqui
        # apenas como cenário de planejamento: não há soma dupla possível.
        st.caption("O retroativo potencial ainda não é reconhecido: entra apenas como "
                   "cenário de planejamento e não integra a complementação confirmada. "
                   "O valor atualizado em análise é exposição — não é somado à adequação. "
                   "Essa mesma parcela já integra o Valor Total Atualizado por critério "
                   "prudencial; a adequação parte do retroativo reconhecido e da projeção, "
                   "nunca do VTA, de modo que ela não é contada duas vezes.")

    st.markdown("**Data final da vigência contratual**")
    st.warning("**Confira a vigência antes de prosseguir.** Essa data define o "
               "último mês da projeção, a complementação orçamentária e a "
               "programação por exercício. Não é obtida automaticamente da apuração.")
    # Escopo da cor restrito a este campo; o usuario nunca deve confundir
    # a vigencia informada manualmente com a data de corte canonica.
    st.markdown("""<style>
    .st-key-adequacao_v3_data_final_vigencia {
        border: 2px solid #B45309; border-radius: 12px;
        background: #FFF7ED; padding: 12px 14px;
    }
    .st-key-adequacao_v3_data_final_vigencia input {
        font-weight: 700; font-size: 1.05rem;
    }
    </style>""", unsafe_allow_html=True)
    _vigencia_digitada = st.text_input(
        "Data final da vigência contratual (dd/mm/aaaa)",
        value=st.session_state.get("adequacao_v3_data_final_vigencia", ""),
        placeholder="Ex.: 16/04/2027", key="adequacao_v3_data_final_vigencia",
        help="Informe a data contratual completa; somente dd/mm/aaaa válido é aceito.")
    data_final_vigencia, erro_vigencia = validar_vigencia_para_adequacao(
        _vigencia_digitada, data_corte_canonica)
    if not _vigencia_digitada.strip():
        st.caption("Preencha a vigência contratual para habilitar a projeção.")
    elif erro_vigencia:
        st.error(erro_vigencia + " A projeção futura ficará não determinada.")
    else:
        st.success(f"Vigência considerada no cálculo: {data_final_vigencia}. "
                   "Confirme essa data no contrato.")
        _ultimo_dia = _as_date(data_final_vigencia)
        if _ultimo_dia and _ultimo_dia.day < monthrange(
                _ultimo_dia.year, _ultimo_dia.month)[1]:
            st.warning(
                "**Atenção ao último mês da vigência:** a projeção por média "
                "considera a competência de encerramento inteira, mesmo "
                f"que o contrato termine em {data_final_vigencia}. "
                "Não existe rateio automático por dias. Se a previsão de "
                "consumo do último mês for parcial, ajuste o valor dessa "
                "competência na aba Projeção, com fundamento do fiscal.")

    # Base historica para a projecao (Seção 6: o seletor de origem vive na Base).
    # Escolha orientada, com sugestao automatica pela fonte encontrada. Usamos
    # st.radio (controle clean testavel via AppTest — o segmented_control desta
    # versao do Streamlit nao e dirigivel entre reruns). A matematica de cada
    # origem nao muda; a aba Historico apenas consome a origem escolhida aqui.
    st.markdown("**Base histórica para a projeção**")
    st.caption(f"Método sugerido: {metodo_sugerido}. Fontes encontradas: {fontes_txt}.")
    _opcoes_origem = ["Financeiro", "Pedidos de Compra"]
    origem_hist = st.radio(
        "Base histórica para a projeção",
        _opcoes_origem, horizontal=True,
        index=(0 if metodo_sugerido != "Pedidos de Compra" else 1),
        key="adequacao_v3_origem", label_visibility="collapsed")
    if origem_hist not in _opcoes_origem:
        origem_hist = metodo_sugerido if metodo_sugerido in _opcoes_origem else "Financeiro"
    origem_pc = origem_hist == "Pedidos de Compra"

    # Valores efetivamente usados no calculo (importados por padrao; ajustaveis
    # apenas em Opcoes avancadas, sem contaminar o fluxo normal).
    retroativo = ctx["valor_represado"]
    percentual_txt = pct(ctx["variacao"])

    with st.expander("Opções avançadas", expanded=False):
        st.caption("Ajustes excepcionais. O usuário comum não precisa abrir esta seção.")
        if not tem_apuracao:
            st.caption("Valores não localizados na apuração — informe manualmente.")
        retroativo = input_moeda(
            "Ajuste excepcional do retroativo reconhecido", ctx["valor_represado"],
            "adequacao_v3_retroativo",
            help="O retroativo reconhecido vem da apuração; ajuste apenas em caráter "
                 "excepcional. NÃO some aqui o retroativo potencial: ele é importado "
                 "automaticamente da apuração e permanece em grandeza separada.")
        percentual_txt = st.text_input("Ajustar percentual de reajuste aplicado",
                                       value=pct(ctx["variacao"]), key="adequacao_v3_percentual")
        _percentual_prev, fator_prev = percentual_e_fator_da_adequacao(
            ctx["variacao"], tem_apuracao, percentual_txt
        )
        st.caption(f"Fator usado: {fator_prev:.6f}".replace(".", ",")
                   + " · informação técnica (não é necessária no fluxo normal).")
        st.markdown("**Metodologia do Valor Total Atualizado importado**")
        st.caption(_metodologia_valor_importado_texto(resultado))
        if st.button("Limpar ajustes desta adequação", key="adequacao_v3_reset"):
            for _k in [k for k in list(st.session_state.keys()) if str(k).startswith("adequacao_v3_")]:
                del st.session_state[_k]
            st.rerun()

    # Fator EXATO (Etapa 51B): o percentual canonico da apuracao (float, ex.:
    # 0.028899355) so e substituido quando o usuario EDITA o campo. O
    # round-trip float -> "2,89%" -> float degradava o fator e o valor
    # reajustado; percentual exibido e apresentacao, nunca fonte matematica.
    # REGRA PETREA: ajuste manual passa pela fonte canonica (2 casas).
    percentual_reajuste, fator_reajuste = percentual_e_fator_da_adequacao(
        ctx["variacao"], tem_apuracao, percentual_txt
    )


# ================================================================ TAB 2 — HISTORICO
with tab_hist:
    st.subheader("2. Histórico utilizado")
    # O seletor de origem vive na aba Base (Seção 6). Aqui apenas consumimos a
    # escolha ja feita — sem oferecer novamente o seletor (estado unico:
    # adequacao_v3_origem).
    st.caption(f"Histórico utilizado: {origem_hist}. Método sugerido: {metodo_sugerido}. "
               f"Fontes encontradas: {fontes_txt}.")

    # A matematica vive no motor; aqui apenas escolhemos a referencia mensal
    # (media_ref) conforme a origem. O restante do fluxo e comum as duas origens.
    media_ref = media_6
    origem_hist_rotulo = "Financeiro mensal"
    janela_meses_pc = None
    janela_rotulo = ""
    ultima_comp = ultima_comp_fin
    ultima_comp_txt = ultima_comp_fin_txt
    base_pc_cenario = None

    if not origem_pc:
        # ----- HISTORICO FINANCEIRO: a tabela e o elemento principal -----
        origem_hist_rotulo = "Financeiro mensal"
        if ultimos_6.empty:
            st.warning("Não foi localizado histórico financeiro nesta apuração. "
                       "Use Pedidos de Compra ou carregue uma Coleta com base financeira mensal.")
        else:
            df_fin_vis = pd.DataFrame({
                "Competência": ultimos_6["Competência"].tolist(),
                "Valor": ["—" if v is None else moeda(v) for v in ultimos_6["valor"].tolist()],
                "Situação": ultimos_6["Situação"].tolist(),
            })
            st.dataframe(df_fin_vis, use_container_width=True, hide_index=True)

        ajustar_fin = st.toggle("Ajustar valores somente para esta adequação",
                                value=False, key="adequacao_v3_ajustar_fin",
                                help="Por padrão o histórico importado é apenas leitura e não é alterado.")
        if ajustar_fin and not ultimos_6.empty:
            base_fin_editor = pd.DataFrame({
                "Competência": ultimos_6["Competência"].tolist(),
                "Valor importado": ["" if v is None else moeda(v, com_prefixo=False)
                                    for v in ultimos_6["valor"].tolist()],
                "Valor considerado": ["" if v is None else moeda(v, com_prefixo=False)
                                     for v in ultimos_6["valor"].tolist()],
                "Observação": [""] * len(ultimos_6),
            })
            # O editor NAO exibe a "Situação" IMPORTADA (que contradiria o valor
            # considerado). O valor importado permanece intacto (coluna disabled).
            df_fin_ed = st.data_editor(
                base_fin_editor, hide_index=True, use_container_width=True, num_rows="fixed",
                key="adequacao_v3_fin_editor",
                column_config={
                    "Competência": st.column_config.TextColumn("Competência", disabled=True),
                    "Valor importado": st.column_config.TextColumn("Valor importado", disabled=True),
                    "Valor considerado": st.column_config.TextColumn("Valor considerado"),
                    "Observação": st.column_config.TextColumn("Observação"),
                })
            # Situacao CONSIDERADA derivada do valor EFETIVAMENTE usado (coerencia
            # ZERO x VAZIO): vazio => Sem informação; 0 => Zero informado; !=0 =>
            # Informado. Nunca converte vazio em zero; so o informado entra na media.
            considerados = []
            linhas_sit = []
            for comp_lbl, importado, bruto, obs in zip(
                    df_fin_ed["Competência"].tolist(), df_fin_ed["Valor importado"].tolist(),
                    df_fin_ed["Valor considerado"].tolist(), df_fin_ed["Observação"].tolist()):
                informado = valor_original_foi_informado(bruto)
                linhas_sit.append({
                    "Competência": comp_lbl,
                    "Valor importado": str(importado) if str(importado).strip() else "—",
                    "Valor considerado": moeda(parse_moeda_br(bruto)) if informado else "—",
                    "Situação considerada": situacao_financeira_considerada(bruto),
                    "Observação": obs,
                })
                if informado:
                    considerados.append(parse_moeda_br(bruto))
            st.dataframe(pd.DataFrame(linhas_sit), use_container_width=True, hide_index=True)
            media_ref = media_financeiro(considerados)["media_mensal"]
            comp_informadas = len(considerados)
            st.caption("O valor importado (fiscal) permanece intacto; o valor considerado ajusta "
                       "apenas esta adequação. A situação considerada reflete o valor usado.")

        st.markdown(f"**Média das competências informadas:** {moeda(media_ref)}")
        cA, cB, cC = st.columns(3)
        with cA:
            render_card_valor("Competências utilizadas", comp_total_janela or 0, formato="inteiro")
        with cB:
            render_card_valor("Com informação", comp_informadas, formato="inteiro")
        with cC:
            render_card_valor("Sem informação", max(0, comp_total_janela - comp_informadas), formato="inteiro")
        with st.expander("Opções avançadas do histórico / diagnóstico", expanded=False):
            st.caption(f"Base de execução mensal detectada: {origem_financeira}")
            st.caption("Janela: 6 competências-calendário terminando na última competência informada. "
                       "Zero informado entra na média; sem informação fica fora do denominador.")

    else:
        # ----- HISTORICO PEDIDOS DE COMPRA: tabela com checkbox USAR -----
        origem_hist_rotulo = "Pedidos de compra"
        registros_pc = carregar_itens_pc_da_sessao(resultado, diagnostico)
        # A competencia final do HISTORICO e a data de corte da apuracao. A data
        # final da VIGENCIA e outro conceito, continua informada separadamente na
        # aba Base e nunca substitui o corte: confundir os dois zerava a projecao
        # futura. Precedencia unica: corte canonico -> ultima competencia do
        # historico financeiro -> campo manual (queda para apuracao sem corte).
        origem_comp_final = ""
        if data_corte_canonica is not None:
            ultima_comp_data = data_corte_canonica
            origem_comp_final = "data de corte da apuração"
        elif ultima_comp_fin is not None:
            ultima_comp_data = ultima_comp_fin.to_timestamp().date()
            origem_comp_final = "última competência do histórico financeiro"
        else:
            ultima_comp_data = None

        if ultima_comp_data is not None:
            ultima_comp = pd.Period(ultima_comp_data, freq="M")
            ultima_comp_txt = periodo_para_label(ultima_comp)
            render_leitura([
                ("Competência final do histórico", ultima_comp_txt),
                ("Origem", origem_comp_final),
            ])
            st.caption("A competência final do histórico vem da apuração — não há "
                       "redigitação. A data final da vigência contratual é informada "
                       "separadamente na aba Base e delimita apenas a projeção futura.")
        else:
            comp_ref_txt = st.text_input(
                "Competência final considerada (mm/aaaa)",
                value=st.session_state.get("adequacao_v3_comp_ref_pc", ""),
                placeholder="Ex.: 06/2026", key="adequacao_v3_comp_ref_pc",
                help="Esta apuração não possui data de corte utilizável nem histórico "
                     "financeiro; informe a competência final do histórico. Não informe "
                     "aqui a data final da vigência: são conceitos distintos.")
            comp_ref = normalizar_competencia(comp_ref_txt)
            if comp_ref is not None:
                ultima_comp = comp_ref
                ultima_comp_data = comp_ref.to_timestamp().date()
                ultima_comp_txt = periodo_para_label(comp_ref)

        if not registros_pc:
            st.warning("NÃO HÁ PEDIDOS DE COMPRA DISPONÍVEIS PARA ESTA ADEQUAÇÃO. "
                       "Utilize a origem Financeiro ou carregue uma Coleta com itens_PC.")
            media_ref = 0.0
        elif ultima_comp_data is None:
            st.warning("Informe a competência final considerada (mm/aaaa) para delimitar a "
                       "janela histórica dos Pedidos de Compra.")
            media_ref = 0.0
        else:
            # JANELA. O padrao e o historico realmente disponivel ate o corte:
            # comeca no primeiro mes com PC importado e termina no mes do corte.
            # Nao existe mais bloco fixo de meses como default — nada de meses
            # artificiais anteriores ao primeiro PC so para completar tamanho.
            # A janela olha TODOS os PCs importados: exclusao manual muda a
            # media, nunca desloca o inicio do historico.
            # A janela automatica e ancorada na DATA DE CORTE canonica. Sem
            # corte na apuracao nao existe janela derivavel: o comportamento
            # anterior (janela informada) permanece, sem mudar o golden.
            janela_auto = janela_automatica_pcs(
                pedidos_de_itens_pc(registros_pc), data_corte_canonica)
            janela_manual = bool(st.session_state.get("adequacao_v3_janela_manual", False))
            if janela_auto is not None and not janela_manual:
                janela_meses_pc = janela_auto["meses"]
                janela_rotulo = "Automático — histórico disponível até o corte"
            else:
                janela_meses_pc = int(st.session_state.get(
                    "adequacao_v3_janela", (janela_auto or {}).get("meses") or 39))
                janela_rotulo = ("Ajuste avançado — janela informada manualmente"
                                 if janela_auto is not None
                                 else "Sem data de corte na apuração — janela informada")
            # ELEGIBILIDADE TEMPORAL (janela) e' separada de EXCLUSAO MANUAL. A
            # classificacao PURA (sem exclusoes) da a situacao temporal de cada PC.
            cl_prev = classificar_pedidos(pedidos_de_itens_pc(registros_pc),
                                          ultima_comp_data, janela_meses_pc)
            excl_state = st.session_state.get("adequacao_v3_exclusoes_pc", set())
            if not isinstance(excl_state, (set, list, tuple)):
                excl_state = set()
            excl_state = set(str(e) for e in excl_state)
            # USAR seed: dentro da janela E nao excluido manualmente. Um PC "Fora da
            # janela" aparece desmarcado (nao usavel), mas isso NAO e exclusao manual.
            base_pc_editor = pd.DataFrame([{
                "USAR": (x["situacao"] == "Considerado") and (str(x["identificacao"]) not in excl_state),
                "PC": str(x["identificacao"]),
                "Data": x["data"].strftime("%d/%m/%Y") if x["data"] else "",
                "Valor": moeda(x["valor"]),
                "Situação": x["situacao"],
            } for x in cl_prev["pedidos"]])
            st.caption("Marque em USAR os Pedidos de Compra (dentro da janela) que entram na "
                       "média. Desmarcar um PC elegível o exclui manualmente; PCs 'Fora da "
                       "janela' não entram e não viram exclusão manual.")
            # Editor recria (key com competencia/janela) quando a janela muda: o USAR
            # e re-semeado a partir das exclusoes MANUAIS persistidas — um PC que
            # volta a ser elegivel e nunca foi excluido reaparece marcado.
            df_pc_ed = st.data_editor(
                base_pc_editor, hide_index=True, use_container_width=True, num_rows="fixed",
                key=f"adequacao_v3_pc_editor_{ultima_comp_txt}_{janela_meses_pc}",
                column_config={
                    "USAR": st.column_config.CheckboxColumn("USAR"),
                    "PC": st.column_config.TextColumn("PC", disabled=True),
                    "Data": st.column_config.TextColumn("Data", disabled=True),
                    "Valor": st.column_config.TextColumn("Valor", disabled=True),
                    "Situação": st.column_config.TextColumn("Situação", disabled=True),
                })
            # Atualiza SOMENTE as exclusoes MANUAIS (voluntarias). PC fora da janela
            # preserva seu estado anterior; nao vira exclusao por (in)elegibilidade.
            exclusoes_manuais = atualizar_exclusoes_manuais_pc(
                [{"pc": str(r["PC"]), "eligivel": (str(r["Situação"]) == "Considerado"),
                  "usar": bool(r["USAR"])} for _, r in df_pc_ed.iterrows()],
                excl_state)
            st.session_state["adequacao_v3_exclusoes_pc"] = exclusoes_manuais
            # A janela permanece soberana; a media vem do motor com as exclusoes manuais.
            peds_pc = pedidos_de_itens_pc(registros_pc, exclusoes=exclusoes_manuais)
            base_pc = media_pedidos_compra(peds_pc, ultima_comp_data, janela_meses_pc)
            media_ref = base_pc["media_mensal"]
            base_pc_cenario = {**base_pc, "meses_historico": janela_meses_pc}

            render_leitura([
                ("Período utilizado",
                 f"{base_pc['inicio_janela'].strftime('%d/%m/%Y')} a {base_pc['fim_janela'].strftime('%d/%m/%Y')}"),
                ("Meses", str(janela_meses_pc)),
                ("Janela", janela_rotulo),
                ("PCs considerados", str(base_pc["pedidos_considerados"])),
                ("Meses com PCs", f"{base_pc['meses_com_pedido']} de {janela_meses_pc}"),
                ("Média mensal", moeda(media_ref)),
            ])
            st.caption("A média mensal divide o valor dos PCs pelos meses do histórico, "
                       "inclusive os meses sem compras. Ela pode servir como base de "
                       "planejamento, mas não comprova que haverá compras todos os meses.")
            if base_pc["pedidos_considerados"] == 0:
                st.info("0 PCs considerados na janela escolhida. A janela permanece soberana "
                        "(não é movida para encaixar pedidos).")
            with st.expander("Opções avançadas do histórico", expanded=False):
                st.checkbox(
                    "Ajustar manualmente a janela histórica",
                    key="adequacao_v3_janela_manual",
                    help="Por padrão a janela é o histórico disponível até o corte. "
                         "Marque apenas para arbitrar outra janela em caráter excepcional.")
                nova_janela = st.slider("Janela histórica dos pedidos (meses)", 1, 60,
                                        value=janela_meses_pc, key="adequacao_v3_janela",
                                        disabled=not janela_manual)
                st.caption(f"Meses com PCs: {base_pc['meses_com_pedido']} · "
                           f"Meses sem PCs: {base_pc['meses_sem_pedido']} · "
                           f"Total histórico: {moeda(base_pc['total_historico'])}.")
                if nova_janela != janela_meses_pc:
                    st.caption("Nova janela aplicada ao recarregar.")

    # Referencia mensal reajustada = media da origem escolhida x fator (lugar unico:
    # a formacao da media vive aqui, no Historico).
    referencia_hist = _round2(media_ref * fator_reajuste)
    st.markdown(f"**Referência mensal reajustada:** {moeda(referencia_hist)} (média × fator).")


# ================================================================ TAB 3 — PROJECAO
with tab_proj:
    st.subheader("3. Projeção futura")
    # O encerramento deve suceder a ultima competencia efetivamente observada.
    # A checagem e adicional ao bloqueio pela data de corte (aba Base).
    if data_final_vigencia and ultima_comp is not None:
        _vig = _as_date(data_final_vigencia)
        if _vig is not None and (int(_vig.year), int(_vig.month)) <= (
                int(ultima_comp.year), int(ultima_comp.month)):
            erro_vigencia = ("A vigência deve terminar após a última "
                             "competência do histórico selecionado.")
            data_final_vigencia = ""
    periodos = gerar_periodos_projecao(ultima_comp, data_final_vigencia)
    periodo_inicio_txt = periodo_para_label(periodos[0]) if periodos else "[campo a preencher]"
    periodo_fim_txt = periodo_para_label(periodos[-1]) if periodos else "[campo a preencher]"
    periodo_projecao_txt = f"{periodo_inicio_txt} a {periodo_fim_txt}" if periodos else "[campo a preencher]"

    if erro_vigencia:
        st.warning(f"**Vigência pendente de validação (aba 1).** {erro_vigencia} "
                   "Nenhuma projeção será utilizada.")
    elif not periodos:
        st.info("Não há competências futuras a projetar com os dados informados.")

    # ----- Premissa de projecao (Etapa 51B) --------------------------------
    # A projecao deixa de PRESUMIR mensalidade: a cadencia real observada
    # (mensal, por ciclo, semestral, trimestral...) decide QUANDO ha gasto;
    # o valor por ocorrencia decide QUANTO. Ciclos preclusos entram como
    # evidencia historica de execucao (nao geram retroativo); C0 so como
    # fallback. Toda a matematica vive no motor (_adequacao_orcamentaria).
    ciclos_cad = ciclos_para_cadencia(resultado.get("df_ciclos") if isinstance(resultado, dict) else None)
    if origem_pc:
        pares_hist = pares_de_pedidos(pedidos_de_itens_pc(
            registros_pc_ctx,
            exclusoes=st.session_state.get("adequacao_v3_exclusoes_pc", set())))
    else:
        pares_hist = pares_de_financeiro(fin_por_comp)
    _ultima_comp_data_cad = ultima_comp.to_timestamp().date() if ultima_comp is not None else None
    cadencia = inferir_cadencia(pares_hist, ciclos_cad, _ultima_comp_data_cad)

    OPCOES_PREMISSA = ["Automática (cadência histórica)", "Mensal (média)", "Por ciclo", "Manual"]
    premissa_proj = st.radio("Premissa de projeção", OPCOES_PREMISSA, horizontal=True,
                             key="adequacao_v3_premissa",
                             help="Automática usa a cadência observada ou, com PCs irregulares e "
                                  "amostra suficiente, simula consumo agregado por média histórica "
                                  "(baixa confiabilidade). Mensal aplica a média por decisão explícita. "
                                  "Por ciclo usa o perfil contratual. Manual usa valores do fiscal.")

    _inicio_proj = periodos[0].to_timestamp().date() if periodos else None
    _fim_proj = periodos[-1].to_timestamp().date() if periodos else None
    usar_media = False
    # Cl8us 12.23: historico PC irregular permite CENARIO ORCAMENTARIO pela
    # media, SOMENTE com amostra suficiente, saldo valido e horizonte correto.
    # Rateio mensal ilustrativo nao atesta mensalidade nem pedidos futuros.
    media_irregular = False
    irregular_sem_projecao = False
    base_cadencia = None
    cadencia_aplicada = cadencia
    if premissa_proj == "Mensal (média)":
        usar_media = True
        premissa_rotulo = "Mensal (média histórica) — definida pelo usuário"
    elif premissa_proj == "Manual":
        base_cadencia = {}
        premissa_rotulo = "Manual — programação informada pelo fiscal"
    elif premissa_proj == "Por ciclo":
        if cadencia["padrao"] not in (CADENCIA_MENSAL, CADENCIA_IRREGULAR):
            cadencia_aplicada = cadencia
        else:
            cadencia_aplicada = cadencia_por_ciclo_forcada(pares_hist, ciclos_cad)
        # Padrao POR CICLO: a cadencia segue identificando o padrao historico,
        # mas a base ORCAMENTARIA e proporcionalizada pelos meses restantes da
        # vigencia (nenhum exercicio fica sem cobertura so porque a ocorrencia
        # historica cai em outro mes). Demais padroes seguem o perfil de
        # ocorrencias.
        if cadencia_aplicada["padrao"] == CADENCIA_POR_CICLO:
            base_cadencia = projetar_por_ciclo_proporcional(cadencia_aplicada, _inicio_proj, _fim_proj)
        else:
            base_cadencia = projetar_por_cadencia(cadencia_aplicada, ciclos_cad, _inicio_proj, _fim_proj)
        premissa_rotulo = f"Por ciclo — {cadencia_aplicada['rotulo']}"
    else:  # Automática (cadência histórica)
        if cadencia["padrao"] == CADENCIA_MENSAL:
            usar_media = True
            premissa_rotulo = f"Automática — {cadencia['rotulo']}"
        elif cadencia["padrao"] == CADENCIA_IRREGULAR:
            evidencia_media = avaliar_evidencia_media_irregular_pc(
                base_pc_cenario if origem_pc else None, cadencia,
                saldo_elegivel, len(periodos))
            # A via automatica sem cadencia e restrita a PCs. A avaliacao
            # nunca presume mensalidade apenas porque ha saldo elevado.
            if origem_pc and evidencia_media["apta"] and not erro_vigencia:
                base_por_mes = base_prudencial_irregular(
                    media_ref, len(periodos), saldo_elegivel)
                if base_por_mes is not None:
                    base_cadencia = {
                        p.to_timestamp().date(): base_por_mes for p in periodos}
                    media_irregular = True
                    premissa_rotulo = (
                        "Automática — cenário pela média de PCs; baixa confiabilidade")
                    st.error(
                        "**HISTÓRICO IRREGULAR: PROJEÇÃO POR MÉDIA PARA PLANEJAMENTO.**\n\n"
                        f"{explicar_meses_de_movimento_por_ciclo(cadencia, origem_pc)}\n\n"
                        f"A amostra reúne {evidencia_media['quantidade_pcs']} PCs em "
                        f"{evidencia_media['meses_historico']} meses de histórico, "
                        f"com compras em {evidencia_media['meses_com_pc']} meses. "
                        "Como há dados suficientes para uma estimativa agregada, "
                        "o sistema utiliza a média histórica, respeitando o saldo "
                        "elegível. O rateio pelos meses futuros é apenas um cenário "
                        "orçamentário, **não significa compras mensais previstas**. "
                        "O fiscal deve validar a estimativa antes de utilizá-la.")
            if not media_irregular:
                base_cadencia = {}
                irregular_sem_projecao = True
                premissa_rotulo = "Automática — histórico irregular; futuro não determinado"
                motivos = ("; ".join(evidencia_media["motivos"])
                           if origem_pc else "a origem financeira não atende à regra de PCs")
                st.error(
                    "**HISTÓRICO SEM PERIODICIDADE SUFICIENTE PARA PROJEÇÃO AUTOMÁTICA.**\n\n"
                    f"{explicar_meses_de_movimento_por_ciclo(cadencia, origem_pc)}\n\n"
                    f"Não há evidência suficiente para estimativa automática: {motivos}. "
                    "Isso não significa ausência de consumo. A opção Automática "
                    "não presumirá uma mensalidade nem calculará a diferença futura. "
                    "Selecione Mensal (média) para simular a média expressamente ou "
                    "Manual para informar a programação fundamentada pelo fiscal.")
        elif cadencia["padrao"] == CADENCIA_POR_CICLO:
            base_cadencia = projetar_por_ciclo_proporcional(cadencia, _inicio_proj, _fim_proj)
            premissa_rotulo = f"Automática — {cadencia['rotulo']}"
        else:
            base_cadencia = projetar_por_cadencia(cadencia, ciclos_cad, _inicio_proj, _fim_proj)
            premissa_rotulo = f"Automática — {cadencia['rotulo']}"

    proporcional_ciclo = (not usar_media and base_cadencia is not None and bool(base_cadencia)
                          and (cadencia_aplicada or {}).get("padrao") == CADENCIA_POR_CICLO
                          and premissa_proj != "Manual")

    render_leitura([
        ("Padrão histórico identificado", cadencia["rotulo"]),
        ("Base histórica", ", ".join(cadencia["ciclos_base"]) if cadencia["ciclos_base"] else "—"),
        ("Premissa de projeção", premissa_rotulo),
    ])
    if cadencia["padrao"] == CADENCIA_IRREGULAR:
        # O alerta vermelho ja contem o diagnostico: nao duplicar o texto.
        if not ((irregular_sem_projecao or media_irregular)
                and premissa_proj.startswith("Automática")):
            st.caption(explicar_meses_de_movimento_por_ciclo(cadencia, origem_pc))
    elif cadencia.get("explicacao"):
        st.caption(cadencia["explicacao"])
    if proporcional_ciclo:
        st.caption("Para fins orçamentários, o valor recorrente do ciclo é proporcionalizado "
                   "pelos meses restantes da vigência.")
    if cadencia.get("usa_c0"):
        st.caption("Atenção: cadência inferida a partir de C0 (implantação/investimento inicial) "
                   "por falta de histórico posterior — confiança reduzida.")

    if usar_media:
        st.caption("Deixe vazio para usar a média. Digite 0 se não haverá execução. "
                   "Digite outro valor para substituir a média naquele mês.")
    elif media_irregular:
        st.caption("As bases por competência representam rateio ilustrativo "
                   "da média para fins orçamentários. Deixe Valor a usar vazio "
                   "para manter a sugestão; informe um valor somente quando houver "
                   "fundamento específico do fiscal. O valor informado não será "
                   "reduzido automaticamente se ultrapassar o saldo.")
    elif irregular_sem_projecao and premissa_proj.startswith("Automática"):
        st.caption("Nenhuma base mensal automática foi determinada. A média histórica "
                   "continua disponível na aba Histórico, mas não é presumida para "
                   "os próximos meses. Para informar valores, selecione Manual.")
    elif premissa_proj == "Manual":
        st.caption("Premissa manual: informe os valores projetados por competência. "
                   "Competências sem valor informado ficam com base 0 (sem execução prevista).")
    else:
        st.caption("Deixe vazio para usar a projeção automática pela cadência (meses sem "
                   "ocorrência prevista ficam com base 0). Digite 0 se não haverá execução. "
                   "Digite outro valor para substituir a competência.")

    proj_avancado = st.toggle("Mostrar opções avançadas da projeção", value=False,
                              key="adequacao_v3_proj_avancado",
                              help="Expõe a premissa do valor informado (sem reajuste / já reajustado) por competência.")

    if usar_media:
        base_editor = montar_base_editor(periodos, media_ref)
    else:
        base_editor = montar_base_editor(periodos, media_ref,
                                         base_por_competencia=base_cadencia,
                                         preencher_valor_a_usar=False)
    if irregular_sem_projecao and premissa_proj.startswith("Automática"):
        # Zero aqui significaria previsao de nao executar. Sem cadencia,
        # a base e indeterminada, nao zero.
        base_editor["Base automática pela média"] = "—"
    editor_key = (f"adequacao_v3_editor_{origem_hist}_{premissa_proj}_{ultima_comp_txt}_"
                  f"{data_final_vigencia}_{round(media_ref, 2)}_{round(fator_reajuste, 6)}"
                  f"_{saldo_elegivel}"
                  + ("_cenario_pc_1223" if media_irregular else ""))

    col_premissa = (st.column_config.SelectboxColumn(
        "Premissa do valor informado", options=["Valor sem reajuste", "Valor já reajustado"], required=True)
        if proj_avancado else None)
    df_editor = st.data_editor(
        base_editor, hide_index=True, use_container_width=True, num_rows="fixed", key=editor_key,
        column_config={
            "Competência": st.column_config.TextColumn("Competência", disabled=True),
            "Base automática pela média": st.column_config.TextColumn("Base sugerida", disabled=True),
            "Valor informado pelo fiscal": st.column_config.TextColumn(
                "Valor a usar",
                disabled=(irregular_sem_projecao and premissa_proj.startswith("Automática"))),
            "Premissa do valor informado": col_premissa,
            "Observação": st.column_config.TextColumn("Observação"),
        },
    )

    # Todas as premissas respeitam o saldo elegivel (automatico limitado;
    # valor informado pelo fiscal nunca reduzido — excesso exige correcao).
    if usar_media:
        df_projecao = calcular_projecao(df_editor, media_ref, fator_reajuste,
                                        saldo_elegivel=saldo_elegivel)
    else:
        df_projecao = calcular_projecao(
            df_editor, media_ref, fator_reajuste,
            base_por_competencia=base_cadencia,
            origem_automatica=("Premissa manual" if premissa_proj == "Manual"
                               else ROTULO_MEDIA_IRREGULAR if media_irregular
                               else f"Cadência: {cadencia_aplicada['rotulo']}"),
            vazio_igual_zero=False,
            saldo_elegivel=saldo_elegivel)
    if modo_reduzido_estoque and ultimos_6.empty:
        rem_original = parse_moeda_br(resultado.get("remanescente_original", 0)) if isinstance(resultado, dict) else 0.0
        rem_atualizado = parse_moeda_br(resultado.get("remanescente_reajustado", 0)) if isinstance(resultado, dict) else 0.0
        diferenca_estoque = round(max(rem_atualizado - rem_original, 0.0), 2)
        df_projecao = pd.DataFrame([{
            "Competência": "Estimativa por saldo remanescente", "Origem": "Modo reduzido por itens/estoque",
            "Premissa usada": "Saldo remanescente informado",
            "Valor base considerado": round(rem_original, 2), "Valor reajustado estimado": round(rem_atualizado, 2),
            "Diferença futura a adequar": diferenca_estoque,
            "Observação": "Estimativa sem base mensal; validar antes de formalizar pagamento.",
        }])

    diferenca_futura = float(df_projecao["Diferença futura a adequar"].sum()) if not df_projecao.empty else 0.0
    qtd_meses = 0 if modo_reduzido_estoque and ultimos_6.empty else len(df_projecao)

    # Insuficiencia de dados (Cl8us 12.20): ausencia de projecao nunca vira
    # diferenca futura de R$ 0,00. Valor informado pelo fiscal determina o
    # futuro; zero informado continua sendo zero.
    informou_valor = (not df_projecao.empty and "Origem" in df_projecao.columns
                      and bool((df_projecao["Origem"] == "Valor informado pelo fiscal").any()))
    motivo_futuro = ""
    if modo_reduzido_estoque and ultimos_6.empty:
        motivo_futuro = ""
    elif erro_vigencia:
        motivo_futuro = "data final da vigência inválida ou não informada (aba 1)"
    elif periodos and not informou_valor:
        if premissa_proj == "Manual":
            motivo_futuro = "nenhum valor informado na premissa Manual"
        elif not pares_hist:
            motivo_futuro = "não há histórico de execução para estimar o futuro"
    # Com historico irregular INSUFICIENTE, a opcao Automatica nao gera
    # mensalidade implicita, nem aproveita valores antigos da sessao.
    if periodos and premissa_proj.startswith("Automática") and irregular_sem_projecao:
        motivo_futuro = ("histórico sem cadência confiável; escolha Mensal (média) "
                         "para adotar explicitamente a média ou Manual para informar "
                         "a programação futura")
    # Sem saldo elegivel, nenhuma premissa automatizada pode atestar a
    # diferenca futura como valor conciliado ao limite contratual.
    # Valores expressamente informados na premissa Manual seguem auditaveis
    # e recebem a ressalva de saldo indisponivel da consolidacao.
    if (not motivo_futuro and periodos and saldo_elegivel is None
            and premissa_proj != "Manual"):
        motivo_futuro = ("saldo contratual elegível indisponível para validar "
                         "a projeção futura automática")
    # Uma diferenca futura nao determinada tambem nao pode aparecer como
    # montante calculado nas competencias da tela ou do XLSX.
    if motivo_futuro and not df_projecao.empty:
        for col in ("Valor base considerado", "Valor reajustado estimado",
                    "Diferença futura a adequar"):
            df_projecao[col] = None
        if irregular_sem_projecao and premissa_proj.startswith("Automática"):
            for col in ("Limite do saldo", "Saldo elegível após a competência"):
                if col in df_projecao:
                    df_projecao[col] = None
    resultado_adequacao = consolidar_resultado_adequacao(
        retroativo=retroativo, df_projecao=df_projecao, meses_restantes=qtd_meses,
        saldo_info=saldo_info,
        retroativo_potencial=(retroativo_potencial if medidas_pc else None),
        motivo_indeterminado=motivo_futuro, baixa_confiabilidade=media_irregular,
        limitado_por_saldo=(media_irregular and saldo_elegivel is not None
                            and float(media_ref or 0) * len(periodos) > saldo_elegivel + 0.005))
    ROTULO_COMPLEMENTACAO = resultado_adequacao["rotulo_complementacao"]
    if resultado_adequacao["excesso_saldo"] > 0.004:
        st.error("OS VALORES INFORMADOS EXCEDEM O SALDO CONTRATUAL ELEGÍVEL em "
                 f"{moeda(resultado_adequacao['excesso_saldo'])} (saldo: {moeda(saldo_elegivel)}). "
                 "Os valores do fiscal não foram reduzidos: corrija a projeção. Enquanto isso, "
                 "a diferença futura fica não determinada.")
    elif resultado_adequacao["limitado_saldo"]:
        st.info(f"A projeção automática foi limitada ao saldo contratual elegível "
                f"({moeda(saldo_elegivel)}).")

    # Feedback dinamico compacto (nao antecipa o Resultado final).
    ocorrencias_previstas = 0
    if not df_projecao.empty and "Valor base considerado" in df_projecao.columns:
        ocorrencias_previstas = int((pd.to_numeric(
            df_projecao["Valor base considerado"], errors="coerce").fillna(0) > 0).sum())
    fb1, fb2, fb3 = st.columns(3)
    with fb1:
        if irregular_sem_projecao and premissa_proj.startswith("Automática"):
            render_card_valor("Meses no horizonte", qtd_meses, formato="inteiro")
        elif usar_media or media_irregular:
            render_card_valor("Meses projetados", qtd_meses, formato="inteiro")
        elif proporcional_ciclo:
            render_card_valor("Meses com cobertura proporcional", ocorrencias_previstas,
                              nota=f"{qtd_meses} meses no horizonte", formato="inteiro")
        else:
            render_card_valor("Ocorrências previstas no horizonte", ocorrencias_previstas,
                              nota=f"{qtd_meses} meses no horizonte", formato="inteiro")
    # A media historica segue na aba Historico (auditoria); aqui vale a
    # diferenca media mensal, informacao relevante para a adequacao.
    _futuro_ok = resultado_adequacao["futuro_determinado"]
    with fb2:
        if _futuro_ok:
            render_card_valor("Diferença média mensal projetada",
                              resultado_adequacao["diferenca_media_mensal"],
                              nota=("Baixa confiabilidade" if media_irregular
                                    else "Média equivalente para planejamento"))
        else:
            render_card_valor("Diferença média mensal projetada", NAO_DETERMINADA, formato="texto")
    with fb3:
        if _futuro_ok:
            render_card_valor("Diferença futura projetada", resultado_adequacao["diferenca_futura"])
        else:
            render_card_valor("Diferença futura projetada", NAO_DETERMINADA, formato="texto",
                              nota=resultado_adequacao["motivo_indeterminado"])

    if not (modo_reduzido_estoque and ultimos_6.empty) and not df_projecao.empty:
        with st.expander(f"Ver projeção mês a mês ({qtd_meses} meses · {periodo_projecao_txt})", expanded=False):
            df_proj_vis = df_projecao.copy()
            for col in ["Valor base considerado", "Valor reajustado estimado", "Diferença futura a adequar",
                        "Saldo elegível após a competência"]:
                if col in df_proj_vis.columns:
                    df_proj_vis[col] = df_proj_vis[col].apply(
                        lambda v: NAO_DETERMINADA if v is None or pd.isna(v) else moeda(v))
            st.dataframe(df_proj_vis, use_container_width=True, hide_index=True)


# ================================================================ TAB 4 — RESULTADO
with tab_result:
    st.subheader("4. Resultado")
    # Fonte unica (WEB = XLSX): consolidar_resultado_adequacao, na aba 3.
    complementacao = resultado_adequacao["complementacao"]
    futuro_determinado = resultado_adequacao["futuro_determinado"]
    referencia_reajustada = _round2(media_ref * fator_reajuste)
    # Com premissa por cadencia/manual, a referencia apresentada e o valor por
    # OCORRENCIA (nao uma "media mensal" que o contrato nao tem).
    # Em historico irregular, a media e apenas diagnostico historico;
    # sua presenca no XLSX nao equivale a adota-la para o futuro.
    usa_media_ref = usar_media or media_irregular or (
        irregular_sem_projecao and premissa_proj.startswith("Automática"))
    valor_ref_hist = media_ref if usa_media_ref else (cadencia_aplicada or {}).get("valor_referencia", 0.0)
    if not usa_media_ref:
        referencia_reajustada = _round2(valor_ref_hist * fator_reajuste)
    cronograma = cronograma_por_exercicio(df_projecao, retroativo,
                                          futuro_determinado=futuro_determinado)

    def _card_ou_nd(coluna, rotulo, valor, nota=""):
        with coluna:
            if valor is None:
                render_card_valor(rotulo, NAO_DETERMINADA, formato="texto", nota=nota)
            else:
                render_card_valor(rotulo, valor, nota=nota)

    # PASSADO — retroativo reconhecido em destaque (nunca limitado pelo saldo).
    st.markdown("**PASSADO**")
    render_card_valor("Retroativo reconhecido considerado" if medidas_pc
                      else "Retroativo já apurado", resultado_adequacao["retroativo"],
                      destaque=True, nota="Apurado na apuração original; não é limitado pelo saldo futuro.")

    # FUTURO — meses restantes, diferenca media mensal e diferenca total.
    st.markdown("**FUTURO**")
    _fatura = resultado_adequacao["fatura_media_reajustada"]
    _cols_fut = st.columns(4 if _fatura is not None else 3)
    with _cols_fut[0]:
        render_card_valor("Meses restantes da vigência", resultado_adequacao["meses_restantes"],
                          formato="inteiro", nota=periodo_projecao_txt if periodos else "")
    _card_ou_nd(_cols_fut[1], "Diferença média mensal projetada",
                resultado_adequacao["diferenca_media_mensal"],
                nota="Média equivalente para planejamento, não pagamento mensal uniforme")
    _card_ou_nd(_cols_fut[2], "Diferença futura total projetada",
                resultado_adequacao["diferenca_futura"],
                nota=(f"{qtd_meses} meses" if futuro_determinado
                      else resultado_adequacao["motivo_indeterminado"]))
    if _fatura is not None:
        with _cols_fut[3]:
            render_card_valor("Fatura média reajustada", _fatura,
                              nota="Base média mensal futura + diferença média mensal")

    # TOTAL — complementacao = retroativo reconhecido + diferenca futura.
    st.markdown("**TOTAL**")
    render_card_valor(ROTULO_COMPLEMENTACAO.upper(), complementacao, destaque=True)
    st.markdown("**Entenda este resultado**")
    narrativa_resultado = explicar_resultado_adequacao(
        resultado_adequacao, premissa_proj, media_ref, referencia_reajustada,
        percentual_reajuste, origem_pc=origem_pc, houve_valor_informado=informou_valor)
    # Markdown do Streamlit interpreta o caractere $ como delimitador LaTeX.
    # Escapar apenas na APRESENTACAO preserva o texto numerico oficial e
    # mostra todas as ocorrencias de R$ sem ocultar o cifrao.
    st.markdown(narrativa_resultado.replace("R$", r"R\$"))
    for _ressalva in resultado_adequacao["ressalvas"]:
        st.warning(_ressalva)
    if media_irregular:
        st.warning(f"{ROTULO_MEDIA_IRREGULAR}. Estimativa exclusivamente orçamentária: "
                   "a diferença futura não representa obrigação de pagamento comprovada.")

    # POTENCIAL — cenario de planejamento. O potencial NAO integra o
    # reconhecido, nem a complementacao, nem a programacao por exercicio.
    cenario_potencial = resultado_adequacao["cenario_potencial"]
    if medidas_pc and retroativo_potencial is not None:
        st.markdown("**POTENCIAL — cenário de planejamento com retroativo potencial**")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            render_card_valor("Retroativo potencial (não reconhecido)", retroativo_potencial, alerta=True)
        with col_p2:
            render_card_valor("Cenário com potencial", cenario_potencial)
        st.caption(f"Cenário com potencial = {ROTULO_COMPLEMENTACAO.lower()} + retroativo "
                   "potencial. Valor de planejamento: o potencial não é reconhecido e "
                   "não entra na programação por exercício abaixo.")
    elif medidas_pc:
        st.caption("Retroativo potencial não localizado nesta apuração.")

    st.markdown("**Programação por exercício**")
    if medidas_pc:
        st.caption(f"A programação representa a {ROTULO_COMPLEMENTACAO.upper()}. O retroativo "
                   "potencial não é distribuído aqui.")
    if isinstance(cronograma, pd.DataFrame) and not cronograma.empty:
        total_cron = float(pd.to_numeric(cronograma["Valor"], errors="coerce").sum())
        df_cron = cronograma.copy()
        _linha_total = {"Exercício": "TOTAL", "Valor": moeda(total_cron),
                        "Retroativo reconhecido": moeda(pd.to_numeric(
                            cronograma["Retroativo reconhecido"], errors="coerce").sum()),
                        "Diferença futura projetada": (moeda(pd.to_numeric(
                            cronograma["Diferença futura projetada"], errors="coerce").sum())
                            if futuro_determinado else NAO_DETERMINADA)}
        for _col in ("Retroativo reconhecido", "Diferença futura projetada", "Valor"):
            df_cron[_col] = df_cron[_col].apply(
                lambda v: NAO_DETERMINADA if v is None or pd.isna(v) else moeda(v))
        df_cron = pd.concat([df_cron, pd.DataFrame([_linha_total])], ignore_index=True)
        st.caption("Retroativo reconhecido e diferença futura aparecem separados; "
                   "“Valor” é o total do exercício.")
        st.dataframe(df_cron, use_container_width=True, hide_index=True)
        if abs(total_cron - complementacao) < 0.01:
            st.caption("A soma dos exercícios confere com a complementação necessária.")
        else:
            st.warning("A soma dos exercícios não confere com a complementação — verificar.")
    else:
        st.info("A programação por exercício depende de projeção futura calculada.")

    # Download XLSX (RESUMO / MEDIA / PROJECAO). XLSX-only permanece a regra.
    # Nomenclatura acompanha a cadencia: "media mensal" so quando a premissa e
    # de fato mensal; caso contrario, valor de referencia POR OCORRENCIA.
    resumo_xlsx = [
        ("Origem histórica", origem_hist_rotulo),
        ("Padrão histórico identificado", cadencia["rotulo"]),
        ("Premissa de projeção", premissa_rotulo),
        ("Média mensal histórica" if usa_media_ref else "Valor de referência por ocorrência",
         media_ref if usa_media_ref else valor_ref_hist),
        ("Percentual de reajuste", pct(percentual_reajuste)),
        ("Fator de reajuste (exato)", f"{fator_reajuste:.9f}".replace(".", ",")),
        ("Referência mensal reajustada" if usa_media_ref else "Referência reajustada por ocorrência",
         referencia_reajustada),
        ("Meses com rateio ilustrativo" if media_irregular
         else "Ocorrências previstas no horizonte",
         NAO_DETERMINADA if irregular_sem_projecao and premissa_proj.startswith("Automática")
         else str(ocorrencias_previstas)),
    ]
    # A referencia reajustada e historica quando nao ha cadencia confiavel,
    # nao e valor mensal previsto na modalidade Automatica.
    if media_irregular:
        resumo_xlsx.append(("Interpretação da distribuição",
                            "Rateio ilustrativo da média histórica; não é previsão mensal de PCs"))
    if irregular_sem_projecao and premissa_proj.startswith("Automática"):
        resumo_xlsx.append(("Uso da referência mensal reajustada",
                            "Apenas histórico; não aplicada à projeção automática"))
    # Mesmos valores do Resultado na WEB (fonte unica: resultado_adequacao).
    resumo_xlsx += linhas_resumo_resultado(resultado_adequacao, mostrar_potencial=medidas_pc)
    if origem_pc and janela_meses_pc is not None:
        _base_pc_exp = media_pedidos_compra(
            pedidos_de_itens_pc(carregar_itens_pc_da_sessao(resultado, diagnostico),
                                exclusoes=st.session_state.get("adequacao_v3_exclusoes_pc", set())),
            ultima_comp.to_timestamp().date(), janela_meses_pc)
        resumo_xlsx[3:3] = [
            ("Janela histórica (meses)", str(janela_meses_pc)),
            ("Origem da janela histórica", janela_rotulo),
            ("Competência final do histórico", ultima_comp_txt),
            ("PCs considerados", str(_base_pc_exp["pedidos_considerados"])),
            ("Meses com PCs", str(_base_pc_exp["meses_com_pedido"])),
            ("Meses sem PCs", str(_base_pc_exp["meses_sem_pedido"])),
            ("Total histórico dos PCs", _base_pc_exp["total_historico"]),
        ]
    xlsx_bytes = gerar_xlsx_projecao(ultimos_6, df_projecao, resumo_xlsx,
                                     nota_resumo=(NOTA_RESUMO_PARCIAL if not futuro_determinado
                                                  else NOTA_RESUMO_ESTIMADA if media_irregular
                                                  else NOTA_RESUMO_SALDO),
                                     programacao=cronograma,
                                     ressalvas=resultado_adequacao["ressalvas"])
    st.download_button("Baixar XLSX", data=xlsx_bytes,
        file_name="adequacao_orcamentaria.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False)


# ================================================================ TAB 5 — TEXTO SIGA
with tab_siga:
    # Redacao apenas apresentacional: consome os valores canonicos ja
    # calculados na aba 4 (retroativo, diferenca_futura, complementacao,
    # qtd_meses, retroativo_potencial, cenario_potencial). Nao recalcula
    # nada — mesma regra de negocio da aba 4: o potencial nao integra a
    # complementacao confirmada nem o total a readequar.
    st.subheader("Texto para solicitação no SIGA")

    adm_siga = st.session_state.get("dados_admissibilidade", {}) or {}
    contrato_auto = _valor_contratual_automatico(resultado, adm_siga, "contrato", "numero_contrato")
    contratada_auto = _valor_contratual_automatico(resultado, adm_siga, "contratada", "apostilado", "fornecedor")

    # Limpeza visual: o bloco de campos manuais (Contrato/Contratada/Cláusula
    # de reajuste) foi retirado da interface. Contrato/Contratada preservam a
    # MESMA fonte automática já existente (_valor_contratual_automatico);
    # sem edição manual, o placeholder "[campo a preencher]" de texto_seguro
    # cobre a ausência de fonte, exatamente como já cobria antes. Cláusula
    # não tinha fonte automática — preserva o mesmo default fixo que o campo
    # manual já usava ("Cláusula Oitava"), sem inventar fonte nova.
    contrato_txt = texto_seguro(contrato_auto)
    contratada_txt = texto_seguro(contratada_auto)
    clausula_txt = texto_seguro("Cláusula Oitava")
    vigencia_txt = _data_siga(data_final_vigencia)

    # Mesma condicao da aba 4 (linha do cenario de planejamento): o cenario
    # com potencial so existe quando ha medida de PC e retroativo potencial
    # localizado. cenario_potencial e a variavel ja calculada la em cima.
    tem_potencial = medidas_pc and retroativo_potencial is not None
    retroativo_txt = moeda(retroativo)
    diferenca_txt = (moeda(resultado_adequacao["diferenca_futura"]) if futuro_determinado
                     else "não determinada")
    complementacao_txt = moeda(complementacao)
    potencial_txt = moeda(retroativo_potencial) if tem_potencial else "Não localizado nesta apuração"
    cenario_txt = (moeda(cenario_potencial) if tem_potencial
                   else "Não aplicável (retroativo potencial não localizado)")

    # Mesma fonte canonica da aba 4 (cronograma_por_exercicio ja calculado
    # ali); nao recalcula nem reconstroi o rateio.
    if isinstance(cronograma, pd.DataFrame) and not cronograma.empty:
        total_cron_siga = float(pd.to_numeric(cronograma["Valor"], errors="coerce").sum())
        linhas_programacao = "\n".join(
            f"{row['Exercício']} | {moeda(row['Valor'])}" for _, row in cronograma.iterrows()
        )
        programacao_txt = (
            "Programação por exercício\n\n"
            f"{linhas_programacao}\n"
            f"TOTAL | {moeda(total_cron_siga)}"
        )
    else:
        programacao_txt = (
            "Programação por exercício\n\n"
            "Não disponível (depende de projeção futura calculada)."
        )

    # Somente na projecao por media historica de baixa confiabilidade o total
    # e identificado como estimado; demais modalidades mantem o texto integral.
    rotulo_total_siga = ("TOTAL PARCIAL A READEQUAR (somente retroativo)" if not futuro_determinado
                         else "TOTAL ESTIMADO A READEQUAR" if media_irregular
                         else "TOTAL A READEQUAR")
    nota_estimativa_siga = (
        f"A diferença futura não foi determinada ({resultado_adequacao['motivo_indeterminado']}) "
        "e deverá ser complementada após a definição da projeção.\n\n" if not futuro_determinado
        else "A diferença futura foi estimada pela média histórica, com baixa "
        "confiabilidade, e depende de validação fiscal.\n\n" if media_irregular else "")

    texto_siga = (
        "Solicitação de adequação orçamentária\n\n"
        f"1. Solicita-se adequação orçamentária para o Contrato {contrato_txt}, firmado com a "
        f"{contratada_txt}, com vigência até {vigencia_txt}.\n\n"
        "2. A adequação decorre da atualização dos valores contratuais resultante do reajuste "
        f"previsto na {clausula_txt}.\n\n"
        "3. Para fins de adequação orçamentária, foram considerados os seguintes valores:\n\n"
        "Composição da adequação orçamentária\n\n"
        f"Retroativo reconhecido considerado: {retroativo_txt}\n"
        f"Diferença futura projetada – {qtd_meses} meses: {diferenca_txt}\n\n"
        f"{rotulo_total_siga}: {complementacao_txt}\n\n"
        f"{nota_estimativa_siga}"
        f"{programacao_txt}\n\n"
        "4. Informações complementares:\n\n"
        f"Retroativo reconhecido considerado: {retroativo_txt}\n\n"
        "Retroativo potencial, ainda em aceitação pela área gestora e não incluído "
        f"nesta adequação (já considerado no Valor Total Atualizado por critério "
        f"prudencial): {potencial_txt}\n\n"
        f"Diferença futura projetada para {qtd_meses} meses: {diferenca_txt}\n\n"
        "Cenário de planejamento considerando eventual reconhecimento do retroativo potencial: "
        f"{cenario_txt}\n"
    )

    # SINCRONIZACAO FORCADA (bug reproduzido em producao no PR #83): um
    # st.text_area com key fixa so usa o parametro value= na PRIMEIRA vez
    # que aquele widget e registrado. Em qualquer rerun seguinte, o
    # Streamlit preserva o valor ja associado a key (session_state), e o
    # novo value=texto_siga calculado e silenciosamente ignorado — o texto
    # exibido ficava congelado no primeiro calculo (ex.: antes da vigencia
    # ser informada). Fonte: streamlit/runtime/state/session_state.py,
    # SessionState.register_widget — so grava o valor "on first registration".
    # Este text_area e apenas area de copia (secao 5 da tarefa): nao ha
    # edicao livre do usuario a preservar, so os campos oficiais acima
    # (contrato/contratada/clausula). Por isso a correcao e legitima:
    # sempre que a ASSINATURA dos insumos mudar, sobrescrevemos
    # st.session_state[key] com o texto canonico ANTES de instanciar o
    # widget (sem passar value=), que e o padrao suportado pelo Streamlit
    # para atualizacao programatica de um widget com key.
    assinatura_siga = (
        contrato_txt, contratada_txt, clausula_txt, vigencia_txt,
        retroativo_txt, qtd_meses, diferenca_txt, complementacao_txt,
        potencial_txt, cenario_txt, programacao_txt, media_irregular,
    )
    if st.session_state.get("adequacao_v3_siga_assinatura") != assinatura_siga:
        st.session_state["adequacao_v3_siga_texto_area"] = texto_siga
        st.session_state["adequacao_v3_siga_assinatura"] = assinatura_siga

    st.text_area("Texto pronto para copiar", height=440,
                 key="adequacao_v3_siga_texto_area",
                 help="Selecione tudo (Ctrl+A) e copie (Ctrl+C), ou use Baixar .txt abaixo.")
    st.download_button(
        "Baixar .txt", data=texto_siga.encode("utf-8"),
        file_name="solicitacao_adequacao_orcamentaria_siga.txt", mime="text/plain",
        key="adequacao_v3_siga_download")
