"""
15_Analise_Contratual.py
------------------------
Página unificada de Análise Contratual — cl8us v2.0.

Consolida o fluxo de:
  - Upload e processamento da Coleta Única
  - Exibição dos resultados essenciais (4 métricas)
  - Geração de documentos (PDF, DOCX, XLSX)
  - Detalhes técnicos em expander (sob demanda)

Usa as funções de 03_Valor_Global.py e 04_Relatorio_Global.py
via exec parcial até o marcador # ── INICIO_UI ──, sem
executar o código de interface desses arquivos.
"""

from pathlib import Path
import sys

import pandas as pd
import streamlit as st

from _ui_utils import render_marca_topo, render_aviso_privacidade
from _ponte_coleta_valores import processar_coleta, MODOS_COLETA
try:
    from _leitor_coleta_mestre import ler_coleta_mestre as _ler_v10
    _TEM_LEITOR_V10 = True
except ImportError:
    _TEM_LEITOR_V10 = False

try:
    from _leitor_coleta_unica import ler_coleta_unica as _ler_legado
    _TEM_LEITOR_LEGADO = True
except ImportError:
    _TEM_LEITOR_LEGADO = False


# ─────────────────────────────────────────────────────────────────────
# Carregamento das funções dos módulos existentes
# ─────────────────────────────────────────────────────────────────────

def _carregar_funcoes_modulo(caminho_relativo: str) -> dict:
    """
    Executa apenas a parte de funções de um módulo (até o marcador INICIO_UI),
    retornando o namespace resultante. Não executa código de UI.
    """
    caminho = Path(__file__).resolve().parent / caminho_relativo
    if not caminho.exists():
        return {}
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            fonte = f.read()
        # Corta tudo a partir do marcador de UI
        partes = fonte.split("# ── INICIO_UI ──")
        fonte_funcoes = partes[0]
        env = {}
        exec(compile(fonte_funcoes, str(caminho), "exec"), env)
        return env
    except Exception as exc:
        st.warning(f"Não foi possível carregar funções de {caminho_relativo}: {exc}")
        return {}


@st.cache_resource(show_spinner=False)
def _env_valores():
    return _carregar_funcoes_modulo("03_Valor_Global.py")


@st.cache_resource(show_spinner=False)
def _env_relatorio():
    return _carregar_funcoes_modulo("04_Relatorio_Global.py")


# ─────────────────────────────────────────────────────────────────────
# Helpers de formatação
# ─────────────────────────────────────────────────────────────────────

def _moeda(valor):
    try:
        v = round(float(valor), 2)
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"


def _fator_fmt(valor):
    try:
        return f"{float(valor):,.4f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "—"


def _df_visual(df, moeda_cols=None):
    """Formata DataFrame para exibição."""
    if not isinstance(df, pd.DataFrame) or df.empty:
        return df
    df = df.copy()
    for col in (moeda_cols or []):
        if col in df.columns:
            df[col] = df[col].apply(lambda v: _moeda(v) if isinstance(v, (int, float)) else v)
    return df


# ─────────────────────────────────────────────────────────────────────
# UI principal
# ─────────────────────────────────────────────────────────────────────

render_marca_topo()
st.title("Análise Contratual")
render_aviso_privacidade(tem_upload=True, tem_download=True)

adm = st.session_state.get("dados_admissibilidade")
res = st.session_state.get("resultado_valor_global")

# ── Contexto da Admissibilidade ───────────────────────────────────────
with st.expander("Contexto da Admissibilidade", expanded=not bool(res)):
    if adm:
        c1, c2, c3 = st.columns(3)
        c1.metric("Origem", adm.get("origem") or adm.get("tipo", "Não informado"))
        c2.metric("Índice", adm.get("indice", "Não informado"))
        c3.metric("Ciclos", len(adm.get("ciclos", [])))
    else:
        st.info("Dados de admissibilidade não encontrados na sessão. "
                "Os parâmetros serão lidos do arquivo de coleta.")

# ── Upload e processamento ─────────────────────────────────────────────
st.subheader("Carregar Coleta Única")

tab_nova, tab_legado = st.tabs(["📋 Coleta Única (v2.0)", "📁 Arquivo de Coleta (legado)"])

with tab_nova:
    arquivo_cu = st.file_uploader(
        "Coleta Única preenchida (.xlsx)",
        type=["xlsx"],
        key="ac_upload_cu",
    )
    if arquivo_cu:
        bytes_cu = arquivo_cu.getvalue()
        _diag = ler_coleta_unica(bytes_cu)
        if not _diag.get("ok"):
            st.error(f"Erro na leitura: {_diag.get('erro','')}")
        else:
            _modo_det = _diag.get("modo_preliminar", "Financeiro Histórico")
            _idx = MODOS_COLETA.index(_modo_det) if _modo_det in MODOS_COLETA else 0
            _fin = _diag.get("financeiro", {})
            _its = _diag.get("itens", {})
            col_m, col_i = st.columns([2, 3])
            with col_m:
                modo_confirmado = st.selectbox(
                    "Modo identificado — confirme ou ajuste:",
                    options=MODOS_COLETA,
                    index=_idx,
                    key="ac_modo_confirmado",
                )
            with col_i:
                st.markdown(
                    f"**Financeiro:** {_fin.get('linhas_preenchidas',0)} linhas "
                    f"— {_moeda(_fin.get('total',0))} &nbsp;|&nbsp; "
                    f"**Itens:** {_its.get('itens_cadastrados',0)} "
                    f"&nbsp;|&nbsp; **Ciclos:** {len(_diag.get('ciclos',[]))}",
                    unsafe_allow_html=True,
                )
                for _rv in _diag.get("ressalvas", []):
                    st.warning(_rv, icon="⚠️")

            if st.button("Processar Coleta Única", type="primary", key="ac_btn_cu"):
                _env = _env_valores()
                _processar = _env.get("processar_arquivo_coleta")
                if not _processar:
                    st.error("Função de processamento não carregada. Verifique 03_Valor_Global.py.")
                else:
                    with st.spinner("Processando..."):
                        try:
                            resultado = _processar(bytes_cu)
                            resultado["modo_apuracao"] = modo_confirmado
                            st.session_state["resultado_valor_global"] = resultado
                            st.session_state["modo_coleta_usado"] = modo_confirmado
                            res = resultado
                            st.success(f"Processado — modo: **{modo_confirmado}**")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"Erro no processamento: {exc}")

with tab_legado:
    arquivo_leg = st.file_uploader(
        "Arquivo de Coleta preenchido (.xlsx)",
        type=["xlsx"],
        key="ac_upload_legado",
    )
    if arquivo_leg:
        if st.button("Processar", type="primary", key="ac_btn_legado"):
            _env = _env_valores()
            _processar = _env.get("processar_arquivo_coleta")
            if not _processar:
                st.error("Função de processamento não carregada.")
            else:
                with st.spinner("Processando..."):
                    try:
                        resultado = _processar(arquivo_leg.getvalue())
                        st.session_state["resultado_valor_global"] = resultado
                        res = resultado
                        st.success("Arquivo processado.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"Erro no processamento: {exc}")

# ── Resultados ────────────────────────────────────────────────────────
if not res:
    st.divider()
    st.info("Carregue e processe uma Coleta Única para ver os resultados.")
    st.stop()

st.divider()

# 4 métricas principais
modo_ui = res.get("modo_apuracao", "Completo")
modo_reduzido  = str(modo_ui).lower() in ("reduzido por itens/estoque", "itens/estoque")
modo_consumo   = str(modo_ui).lower() in ("consumo por itens/ciclo",)

# Badge de modo
_cor_modo = {"Completo": "#CCFBF1", "Financeiro Histórico": "#DBEAFE",
             "Reduzido por Itens/Estoque": "#EDE9FE", "Consumo por Itens/Ciclo": "#F6F3EE"}.get(modo_ui, "#F1F5F9")
_txt_modo = {"Completo": "#065F46", "Financeiro Histórico": "#1E3A8A",
             "Reduzido por Itens/Estoque": "#4C1D95", "Consumo por Itens/Ciclo": "#3F4F35"}.get(modo_ui, "#334155")
st.markdown(
    f'<div style="background:{_cor_modo};color:{_txt_modo};border-radius:8px;'
    f'padding:6px 14px;display:inline-block;font-weight:700;font-size:0.88rem;'
    f'margin-bottom:12px;">Modo: {modo_ui}</div>',
    unsafe_allow_html=True,
)

val_total     = res.get("valor_atualizado_contrato", res.get("valor_global_estoque", 0))
val_retro     = res.get("valor_represado_a_pagar", res.get("valor_retroativo_consumo", 0))
fator_acum    = res.get("fator_acumulado", 1.0)
indice_res    = res.get("indice", adm.get("indice", "—") if adm else "—")

m1, m2, m3, m4 = st.columns(4)
m1.metric("Valor Total Atualizado", _moeda(val_total))
m2.metric("Retroativo", _moeda(val_retro))
m3.metric("Fator acumulado", _fator_fmt(fator_acum))
m4.metric("Índice", indice_res)

if modo_reduzido:
    st.warning("Modo Reduzido: retroativo estimado por itens/estoque. "
               "Não substitui apuração financeira definitiva.", icon="⚠️")
elif modo_consumo:
    st.warning("Modo Consumo: retroativo apurado por quantitativos consumidos/ciclo. "
               "Válido quando a fiscalização confirma equivalência com faturamento devido.", icon="⚠️")

# ── Gerar Documentos ──────────────────────────────────────────────────
st.divider()
st.subheader("Gerar Documentos")

tipo_doc = st.radio(
    "Selecione o documento:",
    [
        "Relatório Executivo (PDF)",
        "Minuta de Apostilamento (DOCX)",
        "Sumário Executivo (XLSX)",
        "Itens por Ciclo (XLSX)",
    ],
    horizontal=True,
    key="ac_tipo_doc",
)

if st.button("⬇️ Baixar documento selecionado", key="ac_btn_doc"):
    _env_rel = _env_relatorio()
    if tipo_doc == "Relatório Executivo (PDF)":
        _criar_pdf = _env_rel.get("criar_pdf_relatorio")
        if not _criar_pdf:
            st.error("Função de PDF não carregada. Verifique 04_Relatorio_Global.py.")
        else:
            with st.spinner("Gerando PDF..."):
                try:
                    pdf_bytes = _criar_pdf(adm, res)
                    st.download_button(
                        "📄 Clique para baixar o PDF",
                        data=pdf_bytes,
                        file_name="Relatorio_Executivo.pdf",
                        mime="application/pdf",
                        key="ac_dl_pdf",
                    )
                except Exception as exc:
                    st.error(f"Erro ao gerar PDF: {exc}")

    elif tipo_doc == "Sumário Executivo (XLSX)":
        _gerar_planilha = _env_valores().get("gerar_planilha_executiva")
        if not _gerar_planilha:
            st.error("Função não carregada. Verifique 03_Valor_Global.py.")
        else:
            with st.spinner("Gerando sumário..."):
                try:
                    xlsx_bytes = _gerar_planilha(res)
                    st.download_button(
                        "📊 Clique para baixar o Sumário Executivo",
                        data=xlsx_bytes,
                        file_name="Sumario_Executivo.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="ac_dl_sumario",
                    )
                except Exception as exc:
                    st.error(f"Erro ao gerar sumário: {exc}")

    elif tipo_doc == "Itens por Ciclo (XLSX)":
        _gerar_itens = _env_valores().get("gerar_excel_valores_unitarios_por_ciclo")
        df_vu = res.get("df_valores_unitarios_por_ciclo", res.get("df_itens_consolidado"))
        df_ciclos = res.get("df_ciclos")
        if not _gerar_itens or not isinstance(df_vu, pd.DataFrame) or df_vu.empty:
            st.warning("Planilha de itens não disponível. Verifique se a Coleta Única contém itens preenchidos.")
        else:
            with st.spinner("Gerando planilha de itens..."):
                try:
                    xlsx_itens = _gerar_itens(df_vu, df_ciclos)
                    st.download_button(
                        "📋 Clique para baixar Itens por Ciclo",
                        data=xlsx_itens,
                        file_name="Itens_Por_Ciclo.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="ac_dl_itens",
                    )
                except Exception as exc:
                    st.error(f"Erro ao gerar itens: {exc}")
        _gerar_docx = _env_rel.get("gerar_minuta_apostilamento_docx")
        if not _gerar_docx:
            st.error("Função de DOCX não carregada. Verifique 04_Relatorio_Global.py.")
        else:
            with st.spinner("Gerando minuta..."):
                try:
                    docx_bytes = _gerar_docx(adm, res)
                    st.download_button(
                        "📝 Clique para baixar a Minuta",
                        data=docx_bytes,
                        file_name="Minuta_Apostilamento.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="ac_dl_docx",
                    )
                except Exception as exc:
                    st.error(f"Erro ao gerar minuta: {exc}")

# ── Detalhes Técnicos (expansível) ────────────────────────────────────
st.divider()
with st.expander("▼ Detalhes técnicos da apuração", expanded=False):

    st.markdown("##### Metodologia de corte")
    _env_rel2 = _env_relatorio()
    _aviso_corte = _env_rel2.get("aviso_metodologia_corte_html")
    if _aviso_corte:
        try:
            st.markdown(_aviso_corte(res), unsafe_allow_html=True)
        except Exception:
            pass

    st.markdown("##### Quadro executivo por ciclo")
    df_comp = res.get("df_comparativo")
    if isinstance(df_comp, pd.DataFrame) and not df_comp.empty:
        st.dataframe(
            _df_visual(df_comp, moeda_cols=["Valor", "Antes do Reajuste", "Após Reajuste", "Diferença"]),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("Quadro executivo não disponível.")

    st.markdown("##### Financeiro por ciclo")
    df_fin_ciclo = res.get("df_financeiro_por_ciclo")
    if isinstance(df_fin_ciclo, pd.DataFrame) and not df_fin_ciclo.empty:
        st.dataframe(
            _df_visual(df_fin_ciclo, moeda_cols=[
                "Valor pago efetivo","Valor teórico calculado",
                "Valor pago/faturado","Valor devido reajustado",
                "Delta do ciclo","Delta acumulado",
            ]),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("Financeiro por ciclo não disponível.")

    st.markdown("##### Composição do Valor Total Atualizado")
    df_composicao = res.get("df_composicao_valor_total")
    if isinstance(df_composicao, pd.DataFrame) and not df_composicao.empty:
        st.dataframe(
            _df_visual(df_composicao, moeda_cols=["Valor"]),
            use_container_width=True, hide_index=True,
        )
    else:
        st.info("Composição do valor total não disponível.")

    df_ad = res.get("df_aditivos_executivo", res.get("df_aditivos"))
    if isinstance(df_ad, pd.DataFrame) and not df_ad.empty:
        st.markdown("##### Aditivos e Supressões")
        st.caption("Quadro de controle formal. Valores não somados automaticamente ao total quando já incorporados.")
        cols_ad = [c for c in [
            "Aditivo","Ciclo/Marco","Tratamento do aditivo",
            "Valor do aditivo na assinatura","Fator aplicado","Valor do aditivo reajustado",
        ] if c in df_ad.columns]
        st.dataframe(
            _df_visual(df_ad[cols_ad] if cols_ad else df_ad,
                       moeda_cols=["Valor do aditivo na assinatura","Valor do aditivo reajustado"]),
            use_container_width=True, hide_index=True,
        )
