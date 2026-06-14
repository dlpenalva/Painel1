from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from _ui_utils import render_marca_topo, render_aviso_privacidade


st.set_page_config(page_title="TLB · cl8us - Coleta Única", layout="wide")


from _coleta_unica_gerador import gerar_coleta_unica_inteligente

# =====================================================
# Diagnóstico experimental da Coleta Única preenchida
# =====================================================

def _normalizar_coluna_coleta(valor):
    import re
    import unicodedata
    texto = "" if valor is None else str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    return texto.strip("_")


def _numero_coleta(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if not texto:
        return 0.0
    texto = texto.replace("R$", "").replace("\xa0", "").replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except Exception:
        return 0.0


def _ler_aba_coleta(xls, nome_aba, linha_header):
    try:
        df = pd.read_excel(xls, sheet_name=nome_aba, header=linha_header)
    except Exception:
        return pd.DataFrame()
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~pd.Series(df.columns).astype(str).str.startswith("Unnamed").values]
    return df


def _coluna_por_nome(df, opcoes):
    mapa = {_normalizar_coluna_coleta(c): c for c in df.columns}
    for opcao in opcoes:
        chave = _normalizar_coluna_coleta(opcao)
        if chave in mapa:
            return mapa[chave]
    for chave, original in mapa.items():
        for opcao in opcoes:
            alvo = _normalizar_coluna_coleta(opcao)
            if alvo and alvo in chave:
                return original
    return None


def diagnosticar_coleta_unica(bytes_xlsx):
    from io import BytesIO

    resultado = {
        "ok": False,
        "erro": "",
        "abas": {},
        "indicadores": {},
        "modo_preliminar": "Base insuficiente",
        "ressalvas": [],
    }

    abas_esperadas = [
        "PARAMETROS_CONTRATO",
        "CICLOS",
        "FINANCEIRO_HISTORICO",
        "ITENS_CICLOS",
        "ADITIVOS",
        "CICLO_EM_EXECUCAO",
        "VALIDACOES_FISCAIS",
    ]

    try:
        xls = pd.ExcelFile(BytesIO(bytes_xlsx))
    except Exception as exc:
        resultado["erro"] = f"Não foi possível abrir o XLSX: {exc}"
        return resultado

    abas_existentes = set(xls.sheet_names)
    resultado["abas"] = {aba: (aba in abas_existentes) for aba in abas_esperadas}

    faltantes = [aba for aba, ok in resultado["abas"].items() if not ok]
    if faltantes:
        resultado["ressalvas"].append("Abas ausentes: " + ", ".join(faltantes))

    linhas_financeiras = 0
    total_liquido = 0.0

    if "FINANCEIRO_HISTORICO" in abas_existentes:
        df_fin = _ler_aba_coleta(xls, "FINANCEIRO_HISTORICO", 3)
        col_valor = _coluna_por_nome(df_fin, ["Valor informado pelo fiscal", "Valor líquido considerado", "Valor medido/aprovado", "Valor pago"])
        if col_valor:
            serie_valor = df_fin[col_valor].apply(_numero_coleta)
            linhas_financeiras = int((serie_valor.abs() > 0.004).sum())
            total_liquido = float(serie_valor.sum())
    else:
        resultado["ressalvas"].append("Aba FINANCEIRO_HISTORICO ausente.")

    itens_cadastrados = 0
    itens_com_c0 = 0
    itens_com_rem_atual = 0
    itens_com_rem_anterior = 0

    if "ITENS_CICLOS" in abas_existentes:
        df_itens = _ler_aba_coleta(xls, "ITENS_CICLOS", 3)
        col_item = _coluna_por_nome(df_itens, ["Item"])
        col_qtd_c0 = _coluna_por_nome(df_itens, ["Quantidade contratada C0"])
        col_vu_c0 = _coluna_por_nome(df_itens, ["Valor unitário original C0"])
        col_rem_atual = _coluna_por_nome(df_itens, ["Remanescente ciclo atual/corte"])

        cols_rem_anteriores = [
            c for c in df_itens.columns
            if _normalizar_coluna_coleta(c) in [
                "remanescente_c1",
                "remanescente_c2",
                "remanescente_c3",
                "remanescente_c4",
            ]
        ]

        if col_item:
            itens_cadastrados = int(df_itens[col_item].astype(str).str.strip().replace("nan", "").ne("").sum())

        if col_qtd_c0 and col_vu_c0:
            qtd_c0 = df_itens[col_qtd_c0].apply(_numero_coleta)
            vu_c0 = df_itens[col_vu_c0].apply(_numero_coleta)
            itens_com_c0 = int(((qtd_c0.abs() > 0.004) & (vu_c0.abs() > 0.004)).sum())

        if col_rem_atual:
            itens_com_rem_atual = int((df_itens[col_rem_atual].apply(_numero_coleta).abs() > 0.004).sum())

        if cols_rem_anteriores:
            mask = pd.Series(False, index=df_itens.index)
            for col in cols_rem_anteriores:
                mask = mask | (df_itens[col].apply(_numero_coleta).abs() > 0.004)
            itens_com_rem_anterior = int(mask.sum())
    else:
        resultado["ressalvas"].append("Aba ITENS_CICLOS ausente.")

    aditivos_cadastrados = 0

    if "ADITIVOS" in abas_existentes:
        df_ad = _ler_aba_coleta(xls, "ADITIVOS", 2)
        col_id = _coluna_por_nome(df_ad, ["Identificação"])
        col_valor_ad = _coluna_por_nome(df_ad, ["Valor original"])
        if col_id:
            aditivos_cadastrados = int(df_ad[col_id].astype(str).str.strip().replace("nan", "").ne("").sum())
        elif col_valor_ad:
            aditivos_cadastrados = int((df_ad[col_valor_ad].apply(_numero_coleta).abs() > 0.004).sum())

    corte_operacional = "Não"
    ciclo_corte = ""

    if "CICLO_EM_EXECUCAO" in abas_existentes:
        df_corte = _ler_aba_coleta(xls, "CICLO_EM_EXECUCAO", 3)
        col_campo = _coluna_por_nome(df_corte, ["Campo"])
        col_valor = _coluna_por_nome(df_corte, ["Valor"])
        if col_campo and col_valor:
            for _, row in df_corte.iterrows():
                campo = str(row.get(col_campo, "")).strip().lower()
                valor = str(row.get(col_valor, "")).strip()
                if "aplicar corte operacional" in campo:
                    corte_operacional = valor or "Não"
                if "ciclo em execução" in campo:
                    ciclo_corte = valor

    validacoes = {}

    if "VALIDACOES_FISCAIS" in abas_existentes:
        df_val = _ler_aba_coleta(xls, "VALIDACOES_FISCAIS", 2)
        col_pergunta = _coluna_por_nome(df_val, ["Pergunta"])
        col_resposta = _coluna_por_nome(df_val, ["Resposta"])
        if col_pergunta and col_resposta:
            for _, row in df_val.iterrows():
                pergunta = str(row.get(col_pergunta, "")).strip()
                resposta = str(row.get(col_resposta, "")).strip()
                if pergunta:
                    validacoes[pergunta] = resposta

    financeiro_completo_declarado = next(
        (
            resp for pergunta, resp in validacoes.items()
            if "histórico financeiro está completo" in pergunta.lower()
        ),
        "",
    )

    tem_financeiro = linhas_financeiras > 0 and abs(total_liquido) > 0.004
    tem_itens = itens_cadastrados > 0
    tem_rem_atual = itens_com_rem_atual > 0
    tem_itens_anteriores = itens_com_rem_anterior > 0

    if tem_financeiro and tem_itens and tem_rem_atual and tem_itens_anteriores:
        modo = "Completo"
    elif tem_financeiro and tem_rem_atual:
        modo = "Financeiro Histórico com Estoque Atual"
    elif tem_financeiro:
        modo = "Financeiro Histórico"
    elif tem_itens and tem_rem_atual:
        modo = "Itens/Estoque"
    elif tem_itens:
        modo = "Itens parciais com ressalvas"
    else:
        modo = "Base insuficiente"

    if tem_financeiro and not tem_itens_anteriores:
        resultado["ressalvas"].append("Sem memória itemizada dos ciclos anteriores; usar financeiro histórico para execução anterior, se tecnicamente suficiente.")
    if not tem_financeiro:
        resultado["ressalvas"].append("Sem financeiro histórico preenchido; não calcular retroativo financeiro definitivo.")
    if tem_itens and not tem_rem_atual:
        resultado["ressalvas"].append("Itens cadastrados sem remanescente atual/corte; o saldo remanescente pode ficar limitado.")
    if financeiro_completo_declarado and financeiro_completo_declarado.lower() not in ["sim", "", "nan", "none", "não informado"]:
        resultado["ressalvas"].append(f"Fiscal declarou financeiro histórico como: {financeiro_completo_declarado}.")

    resultado["ok"] = True
    resultado["indicadores"] = {
        "Linhas financeiras preenchidas": linhas_financeiras,
        "Valor líquido financeiro total": total_liquido,
        "Itens cadastrados": itens_cadastrados,
        "Itens com C0 completo": itens_com_c0,
        "Itens com remanescente anterior": itens_com_rem_anterior,
        "Itens com remanescente atual/corte": itens_com_rem_atual,
        "Aditivos cadastrados": aditivos_cadastrados,
        "Corte operacional solicitado": corte_operacional or "Não",
        "Ciclo de corte": ciclo_corte or "",
        "Financeiro completo declarado": financeiro_completo_declarado or "Não informado",
    }
    resultado["modo_preliminar"] = modo
    return resultado


render_marca_topo()
st.title("Coleta Única Inteligente")

render_aviso_privacidade(tem_download=True)

st.info(
    "Módulo experimental para gerar uma Coleta Única mais maleável. "
    "Nesta etapa, o XLSX é apenas modelo de preenchimento e diagnóstico. "
    "Ainda não altera o cálculo do módulo Valores nem substitui o Arquivo de Coleta atual."
)

st.markdown(
    """
### Objetivo

Criar uma base única capaz de receber diferentes níveis de informação fiscal:

- financeiro histórico completo;
- itens desde C0 até o ciclo atual;
- itens apenas do ciclo atual;
- financeiro sem itens anteriores;
- aditivos;
- corte operacional no ciclo em execução;
- validações fiscais e ressalvas.
"""
)

xlsx = gerar_coleta_unica_inteligente()

st.download_button(
    label="Baixar Coleta Única Experimental em XLSX",
    data=xlsx,
    file_name="Coleta_Unica_Inteligente_Experimental.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.warning(
    "Esta versão é experimental. Use para validar a estrutura de coleta. "
    "O leitor do módulo Valores será adaptado somente em etapa posterior."
)





st.divider()
st.subheader("Diagnóstico Experimental da Coleta Única")

st.caption(
    "Envie uma Coleta Única preenchida para verificar a qualidade da base. "
    "Esta leitura é apenas diagnóstica e ainda não altera o módulo Valores."
)

arquivo_diagnostico = st.file_uploader(
    "Enviar Coleta Única preenchida para diagnóstico",
    type=["xlsx"],
    key="upload_coleta_unica_diagnostico",
)

if arquivo_diagnostico is not None:
    bytes_diag = arquivo_diagnostico.getvalue()
    diag = diagnosticar_coleta_unica(bytes_diag)

    if not diag.get("ok"):
        st.error(diag.get("erro", "Não foi possível diagnosticar a Coleta Única."))
    else:
        st.markdown("### Abas localizadas")
        df_abas = pd.DataFrame([
            {"Aba": aba, "Status": "OK" if ok else "Ausente"}
            for aba, ok in diag.get("abas", {}).items()
        ])
        st.dataframe(df_abas, use_container_width=True, hide_index=True)

        st.markdown("### Indicadores da base")
        indicadores = diag.get("indicadores", {})

        col_a, col_b, col_c = st.columns(3)
        col_a.metric("Linhas financeiras", indicadores.get("Linhas financeiras preenchidas", 0))
        col_b.metric("Itens cadastrados", indicadores.get("Itens cadastrados", 0))
        col_c.metric("Aditivos cadastrados", indicadores.get("Aditivos cadastrados", 0))

        col_d, col_e, col_f = st.columns(3)
        col_d.metric("Itens com remanescente atual/corte", indicadores.get("Itens com remanescente atual/corte", 0))
        col_e.metric("Corte operacional", indicadores.get("Corte operacional solicitado", "Não"))
        col_f.metric("Modo preliminar", diag.get("modo_preliminar", "Base insuficiente"))

        st.markdown("### Detalhamento")
        df_ind = pd.DataFrame([
            {"Indicador": chave, "Resultado": valor}
            for chave, valor in indicadores.items()
        ])
        st.dataframe(df_ind, use_container_width=True, hide_index=True)

        st.markdown("### Modo preliminar recomendado")
        st.success(diag.get("modo_preliminar", "Base insuficiente"))

        ressalvas = diag.get("ressalvas", [])
        if ressalvas:
            st.markdown("### Ressalvas")
            for item in ressalvas:
                st.warning(item)
        else:
            st.info("Nenhuma ressalva relevante identificada nesta leitura preliminar.")
