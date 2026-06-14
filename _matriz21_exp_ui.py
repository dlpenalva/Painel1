# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st

from _leitor_matriz_2_1 import ler_matriz_2_1


ROOT_M21 = Path(__file__).resolve().parent
TESTE_PADTEC_M21 = ROOT_M21 / "_TESTES_MATRIZ21_PADTEC" / "TESTE_01_CONCILIACAO_GMP.xlsx"


def _m21_fmt_moeda(v):
    try:
        v = float(v or 0)
    except Exception:
        v = 0.0
    s = f"R$ {v:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _m21_fmt_num(v):
    try:
        v = float(v or 0)
    except Exception:
        return v
    s = f"{v:,.6f}".rstrip("0").rstrip(".")
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _m21_to_df(obj):
    if obj is None:
        return pd.DataFrame()
    if isinstance(obj, pd.DataFrame):
        return obj.copy()
    if isinstance(obj, list):
        return pd.DataFrame(obj)
    if isinstance(obj, dict):
        return pd.DataFrame([obj])
    return pd.DataFrame()


def _m21_formatar_df_visual(obj, moeda_cols=None, fator_cols=None):
    df = _m21_to_df(obj)
    if df.empty:
        return df

    moeda_cols = moeda_cols or []
    fator_cols = fator_cols or []

    for col in moeda_cols:
        if col in df.columns:
            df[col] = df[col].apply(_m21_fmt_moeda)

    for col in fator_cols:
        if col in df.columns:
            df[col] = df[col].apply(_m21_fmt_num)

    return df


def _m21_processar_upload(uploaded_file):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp.write(uploaded_file.getbuffer())
    tmp.close()
    return ler_matriz_2_1(Path(tmp.name))


def _m21_txt(res):
    linhas = []
    linhas.append("RESULTADO EXPERIMENTAL — MATRIZ 2.1")
    linhas.append("")
    linhas.append(f"Arquivo: {res.get('arquivo', '')}")
    linhas.append(f"Valor Total Atualizado: {_m21_fmt_moeda(res.get('total_vta_m21', 0))}")
    linhas.append("")
    linhas.append("MEMORIA_VTA:")
    for row in res.get("memoria_vta_m21", []):
        linhas.append(
            f"- {row.get('Componente','')}: {_m21_fmt_moeda(row.get('Valor',0))} | "
            f"{row.get('Fonte','')} | {row.get('Status','')} | {row.get('Observação','')}"
        )
    linhas.append("")
    linhas.append("VALIDACOES:")
    for row in res.get("validacoes_m21", []):
        linhas.append(f"- {row.get('Validação','')}: {row.get('Resultado','')} | {row.get('Detalhe','')}")
    if res.get("conciliacao_referencia"):
        linhas.append("")
        linhas.append("CONCILIACAO_REFERENCIA:")
        for row in res.get("conciliacao_referencia", []):
            linhas.append(
                f"- {row.get('Parcela','')}: ref={_m21_fmt_moeda(row.get('Valor referência externa',0))} | "
                f"cl8us={_m21_fmt_moeda(row.get('Valor cl8us',0))} | "
                f"dif={_m21_fmt_moeda(row.get('Diferença',0))} | {row.get('Status','')}"
            )
    return "\n".join(linhas)


def _m21_promover_para_documentos(res):
    from _matriz21_resultado_adapter import adaptar_resultado_matriz21

    resultado = adaptar_resultado_matriz21(res)
    st.session_state["resultado_valor_global"] = resultado
    st.session_state["modo_coleta_usado"] = resultado.get("modo_detectado", "Matriz 2.1 experimental")
    st.session_state["resultado_matriz21_experimental_bruto"] = res
    st.session_state["_painel_ativo"] = None
    return resultado


def render_matriz21_experimental():
    with st.expander("Matriz 2.1 experimental — processamento isolado", expanded=True):
        st.caption(
            "Fluxo experimental. Não substitui a Matriz 2.0/v10 e não altera o processamento principal desta página."
        )

        col_upl, col_teste = st.columns([2, 1])

        with col_upl:
            uploaded_m21 = st.file_uploader(
                "Upload de XLSX Matriz 2.1",
                type=["xlsx"],
                key="upload_matriz21_experimental_isolado",
                help="Use uma coleta Matriz 2.1 ou o arquivo de teste PADTEC/GMP.",
            )

        with col_teste:
            usar_teste_m21 = TESTE_PADTEC_M21.exists() and st.button(
                "Processar teste PADTEC/GMP local",
                key="btn_matriz21_teste_padtec_local",
                use_container_width=True,
                disabled=not TESTE_PADTEC_M21.exists(),
            )
            if not TESTE_PADTEC_M21.exists():
                st.caption("Arquivo de teste PADTEC/GMP não encontrado localmente.")

        res = None

        if uploaded_m21 is not None:
            try:
                res = _m21_processar_upload(uploaded_m21)
            except Exception as e:
                st.error(f"Erro ao processar upload Matriz 2.1: {e}")
                return
        elif usar_teste_m21:
            try:
                res = ler_matriz_2_1(TESTE_PADTEC_M21)
            except Exception as e:
                st.error(f"Erro ao processar teste local Matriz 2.1: {e}")
                return

        if not res:
            st.info("Envie uma Matriz 2.1 ou processe o teste PADTEC/GMP local.")
            return

        st.success("Matriz 2.1 experimental processada com sucesso.")

        c1, c2, c3 = st.columns(3)
        c1.metric("Valor Total Atualizado", _m21_fmt_moeda(res.get("total_vta_m21", 0)))
        c2.metric("Tipo", res.get("tipo", "Matriz 2.1"))
        c3.metric("Aditivos/supressões", _m21_fmt_moeda(res.get("total_aditivos", 0)))

        st.info(
            "Para usar este resultado nos botões oficiais de documentos desta página, "
            "promova o resultado experimental para a sessão. A Matriz 2.0/v10 permanece preservada.",
            icon="ℹ️",
        )

        if st.button(
            "Usar este resultado Matriz 2.1 nos documentos desta página",
            key="btn_matriz21_promover_para_documentos",
            type="primary",
            use_container_width=True,
        ):
            try:
                resultado_doc = _m21_promover_para_documentos(res)
                st.success(
                    "Resultado Matriz 2.1 salvo na sessão. "
                    "Os botões oficiais de documentos abaixo passarão a usar este resultado experimental."
                )
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao preparar resultado Matriz 2.1 para documentos: {e}")

        tabs = st.tabs(["Memória VTA", "Validações", "Conciliação", "Financeiro", "Itens/Saldo", "Aditivos"])

        with tabs[0]:
            df_mem = _m21_formatar_df_visual(
                res.get("memoria_vta_m21"),
                moeda_cols=["Valor"],
            )
            st.dataframe(df_mem, use_container_width=True, hide_index=True)

        with tabs[1]:
            df_val = _m21_to_df(res.get("validacoes_m21"))
            st.dataframe(df_val, use_container_width=True, hide_index=True)
            try:
                resultados = set(str(x).upper() for x in df_val.get("Resultado", []))
            except Exception:
                resultados = set()
            if "ERRO" in resultados:
                st.error("Há validações com ERRO. Não seguir para documentos antes de corrigir.")
            elif "ALERTA" in resultados or "DIVERGENTE" in resultados:
                st.warning("Há alertas/divergências a conferir.")
            else:
                st.success("Validações principais OK.")

        with tabs[2]:
            df_conc = _m21_formatar_df_visual(
                res.get("conciliacao_referencia"),
                moeda_cols=["Valor referência externa", "Valor cl8us", "Diferença"],
            )
            if df_conc.empty:
                st.info("Sem conciliação de referência preenchida.")
            else:
                st.dataframe(df_conc, use_container_width=True, hide_index=True)

        with tabs[3]:
            df_fin = _m21_formatar_df_visual(
                res.get("financeiro_linhas"),
                moeda_cols=["valor_base", "valor_atualizado"],
                fator_cols=["fator"],
            )
            st.dataframe(df_fin, use_container_width=True, hide_index=True)

        with tabs[4]:
            df_itens = _m21_formatar_df_visual(
                res.get("itens_linhas"),
                moeda_cols=[
                    "vu_c0",
                    "valor_c0_itens",
                    "valor_saldo_apos_ultima_competencia",
                    "valor_saldo_data_corte",
                ],
            )
            st.dataframe(df_itens, use_container_width=True, hide_index=True)

        with tabs[5]:
            df_adit = _m21_formatar_df_visual(
                res.get("aditivos_linhas"),
                moeda_cols=[
                    "valor_unitario",
                    "valor_consolidado_informado",
                    "valor_original",
                    "valor_atualizado",
                    "valor_computavel",
                ],
                fator_cols=["fator"],
            )
            st.dataframe(df_adit, use_container_width=True, hide_index=True)

        st.download_button(
            "Baixar resultado experimental Matriz 2.1 em TXT",
            data=_m21_txt(res).encode("utf-8"),
            file_name="resultado_matriz21_experimental.txt",
            mime="text/plain",
            key="download_txt_matriz21_experimental",
            use_container_width=True,
        )
