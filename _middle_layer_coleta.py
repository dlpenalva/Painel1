"""
_middle_layer_coleta.py
-----------------------
Camada de normalização entre o resultado calculado (ponte/motor)
e os consumidores finais: UI, documentos PDF, DOCX e planilhas XLSX.

Responsabilidades:
    1. Garantir aliases de campos para compatibilidade com motores legados
    2. Complementar DataFrames derivados (valores unitários, execução, composição)
    3. Classificar alertas em impeditivo / atenção / informativo
    4. Avaliar qualidade da base (score 0-5)

Regra de ouro:
    NÃO recalcular valores financeiros. Apenas normalizar, complementar
    e garantir que os campos esperados pelos documentos existam.

Uso:
    from _middle_layer_coleta import (
        normalizar_resultado,
        classificar_alertas,
        avaliar_qualidade,
        montar_df_valores_unitarios,
        classificar_base_equalizacao,
    )

    res_doc   = normalizar_resultado(res_bruto)
    alertas   = classificar_alertas(diag)
    score     = avaliar_qualidade(res_doc, diag)
    df_vu     = montar_df_valores_unitarios(res_doc)
"""

import re
import unicodedata

import pandas as pd


# ─────────────────────────────────────────────────────────────────
# Helpers privados
# ─────────────────────────────────────────────────────────────────

def _normalizar_ciclo(valor):
    """Extrai e normaliza nome de ciclo (C0, C1, C2...) de um texto."""
    if valor is None:
        return ""
    texto = str(valor).strip().upper()
    if not texto:
        return ""
    m = re.search(r"C\s*(\d+)", texto)
    if m:
        return f"C{int(m.group(1))}"
    m = re.search(r"(\d+)", texto)
    if m:
        return f"C{int(m.group(1))}"
    return texto


def _numero(valor, padrao=0.0):
    """Converte valor para float de forma segura."""
    try:
        if valor is None or pd.isna(valor):
            return padrao
    except Exception:
        pass
    try:
        if isinstance(valor, str):
            txt = valor.strip().replace("R$", "").replace("\xa0", "").replace(" ", "")
            if not txt:
                return padrao
            if "," in txt:
                txt = txt.replace(".", "").replace(",", ".")
            return float(txt)
        return float(valor)
    except Exception:
        return padrao


def _col_por_nome(df, opcoes):
    """
    Localiza coluna pelo nome normalizado (sem acento, sem espaço extra).
    Tenta correspondência exata primeiro, depois parcial.
    Retorna o nome original da coluna ou None.
    """
    if not isinstance(df, pd.DataFrame) or df.empty:
        return None

    def _norm(x):
        s = str(x).strip().lower()
        s = unicodedata.normalize("NFKD", s)
        s = "".join(ch for ch in s if not unicodedata.combining(ch))
        return re.sub(r"[^a-z0-9]+", "_", s).strip("_")

    mapa = {_norm(c): c for c in df.columns}
    for opcao in opcoes:
        n = _norm(opcao)
        if n in mapa:
            return mapa[n]
    for ncol, original in mapa.items():
        for opcao in opcoes:
            n = _norm(opcao)
            if n and n in ncol:
                return original
    return None


def _fatores_por_ciclo(df_ciclos):
    """
    Retorna dict {ciclo_nome: fator_acumulado_efetivo} a partir do df_ciclos.
    C0 sempre tem fator 1.0.
    """
    fatores = {"C0": 1.0}
    if not isinstance(df_ciclos, pd.DataFrame) or df_ciclos.empty:
        return fatores
    if "Ciclo" not in df_ciclos.columns:
        return fatores

    for _, row in df_ciclos.iterrows():
        ciclo = _normalizar_ciclo(row.get("Ciclo", ""))
        if not ciclo:
            continue
        fator = None
        for campo in ("Fator acumulado efetivo", "Fator acumulado", "Fator"):
            if campo in df_ciclos.columns:
                candidato = row.get(campo)
                if candidato is not None and str(candidato).strip() != "":
                    fator = candidato
                    break
        fatores[ciclo] = _numero(fator, 1.0)

    return fatores


def _limpar_emoji(texto):
    """Remove emojis e caracteres não-ASCII de controle de um texto."""
    return re.sub(r"[^\w\s,()/.;:—–\-]", "", str(texto)).strip()


# ─────────────────────────────────────────────────────────────────
# 1. normalizar_resultado
# ─────────────────────────────────────────────────────────────────

def normalizar_resultado(res):
    """
    Recebe o dict bruto do motor/ponte e retorna um dict normalizado,
    pronto para alimentar PDF, DOCX e XLSX sem remendos na UI.

    Não altera valores calculados — apenas garante aliases e complementa
    DataFrames derivados que os documentos legados esperam.

    Parâmetro:
        res: dict — resultado bruto de calcular_resultado_v10 / processar_coleta

    Retorno:
        dict normalizado e complementado
    """
    if not isinstance(res, dict):
        return res

    r = dict(res)

    # ── Aliases de valores financeiros principais ──────────────────
    if "valor_atualizado_contrato" in r:
        r.setdefault("valor_global_atualizado", r["valor_atualizado_contrato"])
        r.setdefault("valor_global_estoque",    r["valor_atualizado_contrato"])

    if "valor_represado_a_pagar" in r:
        r.setdefault("retroativo_total", r["valor_represado_a_pagar"])
        r.setdefault("delta_acumulado",  r["valor_represado_a_pagar"])

    if "valor_teorico_calculado" in r:
        r.setdefault("total_devido_reajustado", r["valor_teorico_calculado"])

    if "valor_pago_efetivo" in r:
        r.setdefault("total_pago_faturado", r["valor_pago_efetivo"])

    fat = _numero(r.get("fator_acumulado", 1.0), 1.0)
    r.setdefault("variacao_acumulada", fat - 1.0)

    # ── Alias df_delta_por_ciclo ───────────────────────────────────
    r.setdefault("df_delta_por_ciclo", r.get("df_financeiro_por_ciclo"))

    # ── Complementar totais a partir de df_financeiro_por_ciclo ───
    df_fin = r.get("df_financeiro_por_ciclo")
    if isinstance(df_fin, pd.DataFrame) and not df_fin.empty:
        df_base = df_fin.copy()
        if "Ciclo" in df_base.columns:
            df_base = df_base[
                df_base["Ciclo"].astype(str).str.strip().str.upper().ne("TOTAL")
            ]

        if "Valor pago efetivo" in df_base.columns:
            total_pago = float(
                pd.to_numeric(df_base["Valor pago efetivo"], errors="coerce")
                .fillna(0).sum()
            )
            if not r.get("total_pago_faturado"):
                r["total_pago_faturado"] = total_pago
            if not r.get("valor_pago_efetivo"):
                r["valor_pago_efetivo"] = total_pago

        if "Valor teórico calculado" in df_base.columns:
            total_teorico = float(
                pd.to_numeric(df_base["Valor teórico calculado"], errors="coerce")
                .fillna(0).sum()
            )
            if not r.get("total_devido_reajustado"):
                r["total_devido_reajustado"] = total_teorico
            if not r.get("valor_teorico_calculado"):
                r["valor_teorico_calculado"] = total_teorico

        if "Delta do ciclo" in df_base.columns and not r.get("valor_represado_a_pagar"):
            r["valor_represado_a_pagar"] = float(
                pd.to_numeric(df_base["Delta do ciclo"], errors="coerce")
                .fillna(0).sum()
            )

    # ── Complementar df_execucao_atualizada ───────────────────────
    df_exec = r.get("df_execucao_atualizada")
    if isinstance(df_exec, pd.DataFrame) and not df_exec.empty:
        df_exec = df_exec.copy()
        if ("Valor executado original" not in df_exec.columns
                and "Valor pago efetivo" in df_exec.columns):
            df_exec["Valor executado original"] = df_exec["Valor pago efetivo"]
        if ("Valor executado atualizado" not in df_exec.columns
                and "Valor teórico calculado" in df_exec.columns):
            df_exec["Valor executado atualizado"] = df_exec["Valor teórico calculado"]
        if ("Percentual acumulado aplicado" not in df_exec.columns
                and "Fator acumulado efetivo" in df_exec.columns):
            df_exec["Percentual acumulado aplicado"] = (
                pd.to_numeric(df_exec["Fator acumulado efetivo"], errors="coerce")
                .fillna(1.0) - 1.0
            )
        r["df_execucao_atualizada"] = df_exec
    elif "df_financeiro_por_ciclo" in r:
        r["df_execucao_atualizada"] = r["df_financeiro_por_ciclo"]

    # ── Complementar df_composicao_valor_total ────────────────────
    df_comp = r.get("df_composicao_valor_total")
    if isinstance(df_comp, pd.DataFrame) and not df_comp.empty:
        df_comp = df_comp.copy()
        if "Componente" not in df_comp.columns:
            df_comp["Componente"] = (
                df_comp["Parcela"] if "Parcela" in df_comp.columns
                else df_comp.iloc[:, 0].astype(str)
            )
        if "Parcela" not in df_comp.columns:
            df_comp["Parcela"] = df_comp["Componente"]
        if "Ciclo/Referência" not in df_comp.columns:
            df_comp["Ciclo/Referência"] = ""
        if "Observação" not in df_comp.columns:
            df_comp["Observação"] = ""
        r["df_composicao_valor_total"] = df_comp

    # ── Aliases de DataFrames de valores unitários ─────────────────
    _vu_a = r.get("df_valores_unitarios_ciclo")
    _vu_b = r.get("df_valores_unitarios_por_ciclo")
    if isinstance(_vu_a, pd.DataFrame) and not _vu_a.empty and not isinstance(_vu_b, pd.DataFrame):
        r["df_valores_unitarios_por_ciclo"] = _vu_a
    if isinstance(_vu_b, pd.DataFrame) and not _vu_b.empty and not isinstance(_vu_a, pd.DataFrame):
        r["df_valores_unitarios_ciclo"] = _vu_b

    # ── Construir df_valores_unitarios se ainda não existir ───────
    _tem_vu = (
        isinstance(r.get("df_valores_unitarios_ciclo"), pd.DataFrame)
        and not r["df_valores_unitarios_ciclo"].empty
    )
    if not _tem_vu:
        df_vu = montar_df_valores_unitarios(r)
        if isinstance(df_vu, pd.DataFrame) and not df_vu.empty:
            r["df_valores_unitarios_ciclo"]    = df_vu
            r["df_valores_unitarios_por_ciclo"] = df_vu

    return r


# ─────────────────────────────────────────────────────────────────
# 2. classificar_alertas
# ─────────────────────────────────────────────────────────────────

def classificar_alertas(diag):
    """
    Classifica os alertas/ressalvas do diagnóstico em três níveis:

        criticos    — impeditivos: bloqueiam resultado completo
                      (C0 ausente, dupla contagem, base insuficiente)
        atencao     — ressalvas técnicas relevantes que não bloqueiam
        informativos — rastreabilidade, ajustes GCC

    Parâmetro:
        diag: dict — retorno do leitor (_ler_coleta_mestre / _leitor_coleta_unica)

    Retorno:
        {
            "criticos":     list[str],
            "atencao":      list[str],
            "informativos": list[str],
        }
    """
    resultado = {"criticos": [], "atencao": [], "informativos": []}

    if not isinstance(diag, dict) or not diag.get("ok"):
        return resultado

    alertas_brutos = diag.get("alertas") or diag.get("ressalvas") or []
    ajustes_gcc    = diag.get("ajustes_gcc", 0)

    for rv in alertas_brutos:
        txt       = str(rv)
        txt_lower = txt.lower()
        txt_limpo = _limpar_emoji(txt)

        if (txt.startswith("❌")
                or "c0 não" in txt_lower
                or "dupla contagem" in txt_lower
                or "insuficiente" in txt_lower
                or "bloqueio" in txt_lower):
            resultado["criticos"].append(txt_limpo)

        elif ("ajuste" in txt_lower
              or "rastreab" in txt_lower
              or "gcc" in txt_lower):
            resultado["informativos"].append(txt_limpo)

        else:
            resultado["atencao"].append(txt_limpo)

    if ajustes_gcc and not any(
        "ajuste" in str(x).lower() for x in resultado["informativos"]
    ):
        resultado["informativos"].append(
            f"{ajustes_gcc} célula(s) com ajuste manual GCC registrado(s)."
        )

    return resultado


# ─────────────────────────────────────────────────────────────────
# 3. avaliar_qualidade
# ─────────────────────────────────────────────────────────────────

def avaliar_qualidade(res, diag=None):
    """
    Avalia a qualidade da base em score 0-5.

    Critérios (1 ponto cada):
        1. Financeiro histórico com linhas preenchidas
        2. C0 informado
        3. Itens cadastrados
        4. Remanescente atual preenchido
        5. Sem alertas críticos

    Parâmetros:
        res:  dict — resultado normalizado
        diag: dict — diagnóstico do leitor (opcional)

    Retorno:
        int 0-5
    """
    score = 0
    diag  = diag or {}

    fin = diag.get("financeiro", {}) if diag.get("ok") else {}
    its = diag.get("itens",      {}) if diag.get("ok") else {}

    # 1. Financeiro histórico
    linhas_fin = fin.get(
        "linhas_mensais_preenchidas",
        fin.get("linhas_preenchidas", 0)
    )
    total_fin = fin.get("total_com_efeito", fin.get("total_pago", 0))
    if linhas_fin > 0 or total_fin > 0:
        score += 1

    # 2. C0 informado
    df_fin = res.get("df_financeiro_por_ciclo")
    if isinstance(df_fin, pd.DataFrame) and not df_fin.empty and "Ciclo" in df_fin.columns:
        tem_c0 = df_fin["Ciclo"].astype(str).str.strip().str.upper().eq("C0").any()
    else:
        tem_c0 = _numero(res.get("valor_pago_efetivo", 0)) > 0
    if tem_c0:
        score += 1

    # 3. Itens cadastrados
    if its.get("itens_cadastrados", 0) > 0:
        score += 1

    # 4. Remanescente atual
    if (its.get("itens_com_rem_atual", 0) > 0
            or _numero(res.get("remanescente_reajustado", 0)) > 0):
        score += 1

    # 5. Sem críticos
    alertas = classificar_alertas(diag)
    if not alertas["criticos"]:
        score += 1

    return score


# ─────────────────────────────────────────────────────────────────
# 4. montar_df_valores_unitarios
# ─────────────────────────────────────────────────────────────────

def montar_df_valores_unitarios(res):
    """
    Constrói o DataFrame de valores unitários por ciclo.

    Prioridade:
        1. df_valores_unitarios_ciclo já calculado e com coluna Ciclo
        2. df_valores_unitarios_por_ciclo já calculado e com coluna Ciclo
        3. Construção a partir de df_itens + df_ciclos

    Retorno:
        DataFrame com colunas:
            Item, Ciclo, Valor unitário, Quantidade, Total R$, Ciclo precluso
        ou DataFrame vazio se dados insuficientes.
    """
    # Prioridade 1 e 2
    for chave in ("df_valores_unitarios_ciclo", "df_valores_unitarios_por_ciclo"):
        df = res.get(chave)
        if isinstance(df, pd.DataFrame) and not df.empty and "Ciclo" in df.columns:
            return df

    # Prioridade 3: construir a partir de df_itens
    df_itens  = res.get("df_itens")
    df_ciclos = res.get("df_ciclos")

    if not isinstance(df_itens, pd.DataFrame) or df_itens.empty:
        return pd.DataFrame()

    col_item = _col_por_nome(df_itens, ["Item"])
    col_vu   = _col_por_nome(df_itens, [
        "Valor unitário original C0",
        "Valor unitário original",
        "VU C0",
        "VU Original",
        "Valor unitario",
    ])
    col_qtd  = _col_por_nome(df_itens, [
        "Quantidade contratada C0",
        "Quantidade contratada",
        "Qtd C0",
        "Quantidade",
    ])

    if not col_item or not col_vu:
        return pd.DataFrame()

    fatores   = _fatores_por_ciclo(df_ciclos)
    params_v10 = res.get("params_v10", {}) or {}
    params     = res.get("params",    {}) or {}
    linhas     = []

    for _, row in df_itens.iterrows():
        item        = row.get(col_item, "")
        vu_original = _numero(row.get(col_vu, 0.0))
        qtd_original= _numero(row.get(col_qtd, 0.0)) if col_qtd else 0.0

        # Linha C0
        if abs(vu_original) > 0.004 or abs(qtd_original) > 0.004:
            linhas.append({
                "Item":           item,
                "Ciclo":          "C0",
                "Valor unitário": vu_original,
                "Quantidade":     qtd_original,
                "Total R$":       vu_original * qtd_original,
                "Ciclo precluso": False,
            })

        # Linhas de remanescente por ciclo
        for col in df_itens.columns:
            nome      = str(col)
            nome_norm = nome.lower()
            if "remanescente" not in nome_norm:
                continue
            if "valor" in nome_norm or "total" in nome_norm:
                continue

            ciclo = _normalizar_ciclo(nome)
            if not ciclo:
                # Coluna de remanescente atual sem ciclo explícito — tentar params
                ciclo = _normalizar_ciclo(
                    params_v10.get("ciclo_atual_corte", "")
                    or params_v10.get("ciclo_atual", "")
                    or params.get("ciclo_atual", "")
                )
            if not ciclo:
                continue

            qtd = _numero(row.get(col, 0.0))
            if abs(qtd) <= 0.004:
                continue

            fator    = fatores.get(ciclo, 1.0)
            vu_ciclo = vu_original * fator
            linhas.append({
                "Item":           item,
                "Ciclo":          ciclo,
                "Valor unitário": vu_ciclo,
                "Quantidade":     qtd,
                "Total R$":       vu_ciclo * qtd,
                "Ciclo precluso": False,
            })

    return pd.DataFrame(linhas)


# ─────────────────────────────────────────────────────────────────
# 5. classificar_base_equalizacao
# ─────────────────────────────────────────────────────────────────

def classificar_base_equalizacao(diag, equalizacao=None):
    """
    Classifica a base combinando os dados do leitor (diag) com as
    respostas da Equalização da Base preenchidas pela GCC.

    NÃO altera cálculo. Apenas categoriza para diagnóstico e relatório.

    Parâmetros:
        diag:         dict — retorno do leitor (_ler_coleta_mestre)
        equalizacao:  dict — st.session_state["equalizacao_base"]
                      Se None, tenta ler do session_state automaticamente.

    Retorno:
        {
            "classificacao": str   — rótulo principal da base
            "confianca":     str   — "Alta" / "Média" / "Baixa"
            "gaps":          list  — lacunas detectadas
            "premissas":     list  — premissas assumidas pela GCC
            "equalizado":    bool  — se a GCC confirmou a equalização
        }
    """
    # Tentar carregar equalizacao do session_state se não fornecida
    if equalizacao is None:
        try:
            import streamlit as st
            equalizacao = st.session_state.get("equalizacao_base", {})
        except Exception:
            equalizacao = {}

    eq = equalizacao or {}
    diag = diag or {}
    fin  = diag.get("financeiro", {}) if diag.get("ok") else {}
    its  = diag.get("itens",      {}) if diag.get("ok") else {}

    tem_financeiro  = fin.get("linhas_mensais_preenchidas", fin.get("linhas_preenchidas", 0)) > 0
    tem_itens       = its.get("itens_cadastrados", 0) > 0
    tem_rem_atual   = its.get("itens_com_rem_atual", 0) > 0
    tem_rem_ant     = its.get("itens_com_rem_anterior", 0) > 0
    tem_aditivos    = len(diag.get("aditivos", [])) > 0
    equalizado      = eq.get("equalizado", False)

    # ── Classificação principal ──────────────────────────────────
    if tem_financeiro and tem_itens and tem_rem_atual and tem_rem_ant:
        classificacao = "Base completa"
    elif tem_financeiro and tem_rem_atual:
        classificacao = "Base híbrida"
    elif tem_financeiro and tem_itens:
        classificacao = "Base financeira com itens parciais"
    elif tem_financeiro:
        classificacao = "Base financeira"
    elif tem_itens and tem_rem_atual:
        classificacao = "Base itemizada"
    elif tem_itens:
        classificacao = "Base com ressalvas"
    else:
        classificacao = "Base insuficiente"

    # ── Confiança — ponderada pela Equalização ───────────────────
    pontos = 0
    if tem_financeiro:
        pontos += 2
    if tem_rem_atual:
        pontos += 1
    if tem_itens:
        pontos += 1
    if equalizado:
        pontos += 1
    if eq.get("financeiro_final") == "Sim":
        pontos += 1
    if eq.get("competencia_final") not in ("", "N/A", "Não sei"):
        pontos += 1
    if eq.get("corte_saldo") == "Sim" and eq.get("data_corte_saldo"):
        pontos += 1

    if pontos >= 6:
        confianca = "Alta"
    elif pontos >= 3:
        confianca = "Média"
    else:
        confianca = "Baixa"

    # ── Gaps ────────────────────────────────────────────────────
    gaps = []
    if not tem_financeiro:
        gaps.append("Sem financeiro histórico — retroativo financeiro definitivo não calculável.")
    if not tem_rem_atual:
        gaps.append("Sem remanescente atual — saldo remanescente não estimável.")
    if eq.get("competencia_final") in ("", "N/A", "Não sei"):
        gaps.append("Competência final do financeiro não informada.")
    if eq.get("corte_saldo") != "Sim" and tem_rem_atual:
        gaps.append("Data de corte do saldo remanescente não confirmada.")
    if eq.get("aditivos_no_saldo") == "Não sei" and tem_aditivos:
        gaps.append("Dúvida sobre aditivos incorporados no saldo — risco de dupla contagem.")
    if eq.get("historico_ciclos_form") in ("Não informado", "Não sei", "N/A"):
        gaps.append("Forma do histórico de ciclos anteriores não declarada.")
    if eq.get("preco_ciclo_cruzado") == "Sim":
        gaps.append("Valores formados num ciclo e executados em outro — verificar corte.")

    # ── Premissas assumidas ──────────────────────────────────────
    premissas = []
    if eq.get("financeiro_final") == "Sim":
        premissas.append("Valores financeiros confirmados como definitivos pela GCC.")
    comp = eq.get("competencia_final", "")
    if comp not in ("", "N/A", "Não sei"):
        premissas.append(f"Competência final considerada: {comp}.")
    data_corte = eq.get("data_corte_saldo", "")
    if data_corte:
        premissas.append(f"Corte do saldo remanescente em: {data_corte}.")
    if eq.get("aditivos_no_saldo") == "Sim":
        premissas.append("Saldo remanescente já considera aditivos/supressões.")
    hist = eq.get("historico_ciclos_form", "")
    if hist not in ("", "N/A", "Não sei", "Não informado"):
        premissas.append(f"Histórico de ciclos informado {hist.lower()}.")

    return {
        "classificacao": classificacao,
        "confianca":     confianca,
        "gaps":          gaps,
        "premissas":     premissas,
        "equalizado":    equalizado,
    }
