"""
_leitor_matriz_2_0.py
---------------------
Leitor da Matriz 2.0 (nite.xlsx / ColetaReajuste.xlsx).

Assinatura de reconhecimento: abas MAPA_CICLOS + EXECUCAO_FINANCEIRA.

Retorna dict compativel com o formato esperado por:
- _middle_layer_coleta.py (classificar_base_equalizacao)
- _equalizacao_base.py (render_equalizacao_base)
- 15_Analise_Contratual.py (processamento e documentos)

Chaves do retorno:
    ok                  bool
    erro                str
    versao              str  — "matriz_2_0"
    abas_encontradas    list
    parametros          dict
    ciclos              list
    financeiro          dict
    itens               dict
    aditivos            dict
    historico           list
    modo_detectado      str
    alertas             list
"""

import re
from io import BytesIO

import pandas as pd


ASSINATURA_M20 = {"MAPA_CICLOS", "EXECUCAO_FINANCEIRA"}
HEADER_ROW = 4   # linha 5 = indice 4 em todas as abas


# ── Utilitarios ───────────────────────────────────────────────────

def _txt(val, pad=""):
    if val is None: return pad
    try:
        if pd.isna(val): return pad
    except Exception: pass
    s = str(val).strip()
    return s if s.lower() not in ("nan","none","nat","<na>","") else pad

def _num(val, pad=0.0):
    if val is None: return pad
    try:
        if pd.isna(val): return pad
    except Exception: pass
    if isinstance(val, (int, float)): return float(val)
    s = str(val).strip().replace("R$","").replace("\xa0","").replace(" ","")
    if "," in s: s = s.replace(".","").replace(",",".")
    try: return float(s)
    except Exception: return pad

def _col(df, opcoes):
    """Localiza coluna pelo nome (busca parcial normalizada)."""
    def _norm(v):
        v = str(v).lower()
        v = re.sub(r"[^a-z0-9]", " ", v)
        return " ".join(v.split())
    mapa = {_norm(c): c for c in df.columns}
    for op in opcoes:
        k = _norm(op)
        if k in mapa: return mapa[k]
        for norm_c, orig_c in mapa.items():
            if k and k in norm_c: return orig_c
    return None

def _ler_aba(xls, nome, header=HEADER_ROW):
    try:
        df = pd.read_excel(xls, sheet_name=nome, header=header, dtype=object)
        df = df.dropna(how="all")
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()


# ── Leitores por aba ──────────────────────────────────────────────

def _ler_mapa_ciclos(xls, abas):
    """MAPA_CICLOS: tabela de lookup comp -> ciclo/fator."""
    resultado = {}
    if "MAPA_CICLOS" not in abas:
        return resultado
    try:
        df = pd.read_excel(xls, sheet_name="MAPA_CICLOS", header=0, dtype=object)
        df = df.dropna(how="all")
        for _, row in df.iterrows():
            comp  = _txt(row.iloc[0] if len(row) > 0 else "")
            ciclo = _txt(row.iloc[1] if len(row) > 1 else "")
            fat   = _num(row.iloc[2] if len(row) > 2 else 1.0, 1.0)
            if comp and ciclo:
                resultado[comp] = {"ciclo": ciclo, "fat": fat}
    except Exception:
        pass
    return resultado


def _ler_base(xls, abas):
    """BASE: parametros e ciclos."""
    params = {}
    ciclos = []
    if "BASE" not in abas:
        return params, ciclos
    df = _ler_aba(xls, "BASE")
    if df.empty:
        return params, ciclos

    col_id   = _col(df, ["ID Ciclo", "Ciclo"])
    col_per  = _col(df, ["Periodo"])
    col_ini  = _col(df, ["Inicio Financeiro"])
    col_fim  = _col(df, ["Fim Financeiro"])
    col_db   = _col(df, ["Data-Base"])
    col_pct  = _col(df, ["% Reajuste", "Percentual"])
    col_fat  = _col(df, ["Fator do Ciclo"])
    col_facu = _col(df, ["Fator Acumulado"])
    col_efeito = _col(df, ["Possui Efeito", "Efeito"])
    col_param  = _col(df, ["Parametro"])
    col_valor  = _col(df, ["Valor"])

    for _, row in df.iterrows():
        nome = _txt(row.get(col_id,"") if col_id else "").upper()
        if not re.match(r"^C\d+$", nome):
            continue
        ciclos.append({
            "ciclo":          nome,
            "periodo":        _txt(row.get(col_per,"")   if col_per   else ""),
            "ini_fin":        _txt(row.get(col_ini,"")   if col_ini   else ""),
            "fim_fin":        _txt(row.get(col_fim,"")   if col_fim   else ""),
            "data_base":      _txt(row.get(col_db,"")    if col_db    else ""),
            "percentual":     _num(row.get(col_pct, 0)   if col_pct   else 0),
            "fator_ciclo":    _num(row.get(col_fat, 1)   if col_fat   else 1, 1.0),
            "fator_acumulado":_num(row.get(col_facu, 1)  if col_facu  else 1, 1.0),
            "efeito":         _txt(row.get(col_efeito,"") if col_efeito else ""),
            "objeto_analise": _txt(row.get(col_efeito,"") if col_efeito else "").lower() == "sim",
        })
        # Parametros (coluna J/K na linha C0)
        if nome == "C0" and col_param and col_valor:
            p = _txt(row.get(col_param,""))
            v = _txt(row.get(col_valor,""))
            if p and v:
                params[p] = v

    return params, ciclos


def _ler_execucao(xls, abas, mapa_ciclos):
    """EXECUCAO_FINANCEIRA: financeiro mensal."""
    resultado = {
        "linhas": [], "linhas_mensais_preenchidas": 0,
        "total_nominal": 0.0, "total_com_efeito": 0.0,
        "total_reajustado": 0.0, "total_delta": 0.0,
        "por_ciclo": {},
    }
    if "EXECUCAO_FINANCEIRA" not in abas:
        return resultado

    df = _ler_aba(xls, "EXECUCAO_FINANCEIRA")
    if df.empty:
        return resultado

    col_comp   = _col(df, ["Competencia", "MM/AAAA"])
    col_valor  = _col(df, ["Valor Nominal", "Valor Consumido"])
    col_efeito = _col(df, ["Possui Efeito", "Efeito"])
    col_ciclo  = _col(df, ["Ciclo"])
    col_fat    = _col(df, ["Fator Aplicavel", "Fator"])
    col_base   = _col(df, ["Valor Base Considerado", "Base Considerada"])
    col_delta  = _col(df, ["Delta"])
    col_reaj   = _col(df, ["Valor Reajustado", "Reajustado"])

    linhas = []
    for _, row in df.iterrows():
        comp  = _txt(row.get(col_comp,"")  if col_comp  else "")
        valor = _num(row.get(col_valor, 0) if col_valor else 0)

        if not comp and abs(valor) < 0.004:
            continue

        efeito = _txt(row.get(col_efeito,"Sim") if col_efeito else "Sim")
        ciclo  = _txt(row.get(col_ciclo,"")     if col_ciclo  else "")

        # Ciclo via mapa se nao preenchido
        if not ciclo and comp and comp in mapa_ciclos:
            ciclo = mapa_ciclos[comp]["ciclo"]

        fat   = _num(row.get(col_fat, 1.0)  if col_fat  else 1.0, 1.0)
        base  = _num(row.get(col_base, 0)   if col_base else 0)
        delta = _num(row.get(col_delta, 0)  if col_delta else 0)
        reaj  = _num(row.get(col_reaj, 0)   if col_reaj  else 0)

        # Recalcular se base zerada mas valor preenchido
        if abs(base) < 0.004 and abs(valor) > 0.004 and efeito.lower() == "sim":
            base  = valor
            delta = round(base * fat - base, 2)
            reaj  = round(base * fat, 2)

        linhas.append({
            "competencia": comp,
            "ciclo":       ciclo,
            "valor_nominal": valor,
            "efeito":      efeito,
            "fator":       fat,
            "base":        base,
            "delta":       delta,
            "reajustado":  reaj,
        })

        if abs(valor) > 0.004:
            resultado["total_nominal"] += valor
            if efeito.lower() == "sim":
                resultado["total_com_efeito"] += base
                resultado["total_reajustado"] += reaj
                resultado["total_delta"]      += delta
            # Acumular por ciclo
            if ciclo:
                pc = resultado["por_ciclo"].setdefault(ciclo, {"nominal":0,"reajustado":0,"linhas":0})
                pc["nominal"]    += valor
                pc["reajustado"] += reaj
                pc["linhas"]     += 1

    resultado["linhas"] = linhas
    resultado["linhas_mensais_preenchidas"] = sum(1 for l in linhas if abs(l["valor_nominal"]) > 0.004)
    resultado["total_nominal"]    = round(resultado["total_nominal"], 2)
    resultado["total_com_efeito"] = round(resultado["total_com_efeito"], 2)
    resultado["total_reajustado"] = round(resultado["total_reajustado"], 2)
    resultado["total_delta"]      = round(resultado["total_delta"], 2)
    return resultado


def _ler_itens(xls, abas):
    """ITENS_CONTRATADOS: base fisica."""
    resultado = {
        "itens": [], "total_c0": 0.0,
        "total_rem_atual_original": 0.0, "total_rem_atual_atualizado": 0.0,
        "itens_cadastrados": 0, "itens_com_rem_atual": 0,
    }
    if "ITENS_CONTRATADOS" not in abas:
        return resultado

    df = _ler_aba(xls, "ITENS_CONTRATADOS")
    if df.empty:
        return resultado

    col_item  = _col(df, ["ID Item", "Item"])
    col_qtd   = _col(df, ["Qtd C0"])
    col_vu    = _col(df, ["VU C0"])
    col_vt    = _col(df, ["VT C0"])
    col_rem   = [_col(df, [f"Remanescente Inicio C{i}"]) for i in range(5)]
    col_ratu  = _col(df, ["Remanescente Atual"])
    col_vtr   = _col(df, ["VT Remanescente Atual Nominal"])
    col_fat   = _col(df, ["Fator Aplicavel", "Fator"])
    col_vtru  = _col(df, ["VT Remanescente Atualizado"])

    itens = []
    total_c0 = total_rn = total_ru = 0.0

    for _, row in df.iterrows():
        item = _txt(row.get(col_item,"") if col_item else "")
        if not item or item.upper() in ("TOTAL",""):
            continue

        qtd   = _num(row.get(col_qtd, 0) if col_qtd else 0)
        vu    = _num(row.get(col_vu,  0) if col_vu  else 0)
        vt_c0 = _num(row.get(col_vt,  0) if col_vt  else 0)
        if vt_c0 == 0 and qtd > 0 and vu > 0:
            vt_c0 = round(qtd * vu, 2)

        rems = [_num(row.get(c, 0) if c else 0) for c in col_rem]
        ratu  = _num(row.get(col_ratu, 0) if col_ratu else 0)
        vtr_n = _num(row.get(col_vtr,  0) if col_vtr  else 0)
        fat   = _num(row.get(col_fat, 1.0) if col_fat else 1.0, 1.0)
        vtr_u = _num(row.get(col_vtru, 0)  if col_vtru else 0)

        if vtr_n == 0 and ratu > 0 and vu > 0:
            vtr_n = round(ratu * vu, 2)
        if vtr_u == 0 and vtr_n > 0:
            vtr_u = round(vtr_n * fat, 2)

        total_c0 += vt_c0
        total_rn += vtr_n
        total_ru += vtr_u

        itens.append({
            "item": item, "qtd_c0": qtd, "vu_c0": vu, "vt_c0": vt_c0,
            "rem_c0": rems[0], "rem_c1": rems[1], "rem_c2": rems[2],
            "rem_c3": rems[3], "rem_c4": rems[4],
            "rem_atual": ratu, "vt_rem_orig": vtr_n,
            "fator": fat, "vt_rem_atu": vtr_u,
            "_tem_rem_atual": ratu > 0.001,
        })

    resultado.update({
        "itens": itens, "total_c0": round(total_c0, 2),
        "total_rem_atual_original":   round(total_rn, 2),
        "total_rem_atual_atualizado": round(total_ru, 2),
        "itens_cadastrados":  len(itens),
        "itens_com_rem_atual": sum(1 for i in itens if i["_tem_rem_atual"]),
    })
    return resultado


def _ler_aditivos(xls, abas):
    """ADITIVOS: acrescimos e supressoes."""
    resultado = {"linhas": [], "total_computavel": 0.0, "qtd_computaveis": 0}
    if "ADITIVOS" not in abas:
        return resultado

    df = _ler_aba(xls, "ADITIVOS")
    if df.empty:
        return resultado

    col_item  = _col(df, ["Item"])
    col_data  = _col(df, ["Data do Aditivo", "Data"])
    col_ciclo = _col(df, ["Ciclo / Marco", "Ciclo"])
    col_tipo  = _col(df, ["Tipo de Alteracao", "Tipo"])
    col_qtd   = _col(df, ["Qtd Acrescida", "Quantidade"])
    col_vu    = _col(df, ["Valor Unitario Original", "VU"])
    col_trat  = _col(df, ["Tratamento do Aditivo", "Tratamento"])
    col_comp  = _col(df, ["Valor Computavel", "Computavel"])

    linhas = []
    total_comp = 0.0

    for _, row in df.iterrows():
        comp_val = _num(row.get(col_comp, 0) if col_comp else 0)
        qtd      = _num(row.get(col_qtd,  0) if col_qtd  else 0)
        if abs(comp_val) < 0.004 and abs(qtd) < 0.001:
            continue
        tipo     = _txt(row.get(col_tipo, "") if col_tipo else "")
        trat     = _txt(row.get(col_trat, "") if col_trat else "")
        computar = "computar" in trat.lower() or "vta" in trat.lower()

        linhas.append({
            "item":        _txt(row.get(col_item,  "") if col_item  else ""),
            "data":        _txt(row.get(col_data,  "") if col_data  else ""),
            "ciclo":       _txt(row.get(col_ciclo, "") if col_ciclo else ""),
            "tipo":        tipo,
            "quantidade":  qtd,
            "vu_original": _num(row.get(col_vu, 0) if col_vu else 0),
            "tratamento":  trat,
            "valor_computavel": comp_val,
            "_computar":   computar,
        })
        if computar:
            total_comp += comp_val

    resultado.update({
        "linhas":           linhas,
        "total_computavel": round(total_comp, 2),
        "qtd_computaveis":  sum(1 for l in linhas if l["_computar"]),
    })
    return resultado


def _ler_historico(xls, abas):
    """HISTORICO_CICLOS: historico por ciclo."""
    linhas = []
    if "HISTORICO_CICLOS" not in abas:
        return linhas
    df = _ler_aba(xls, "HISTORICO_CICLOS")
    if df.empty:
        return linhas

    col_id   = _col(df, ["ID Ciclo", "Ciclo"])
    col_sit  = _col(df, ["Situacao"])
    col_reaj = _col(df, ["Houve Reajuste"])
    col_pct  = _col(df, ["% Concedido"])
    col_pago = _col(df, ["Valor Pago/Informado"])
    col_obs  = _col(df, ["Observacao"])

    for _, row in df.iterrows():
        nome = _txt(row.get(col_id,"") if col_id else "").upper()
        if not re.match(r"^C\d+$", nome):
            continue
        linhas.append({
            "ciclo":    nome,
            "situacao": _txt(row.get(col_sit,  "") if col_sit  else ""),
            "reajuste": _txt(row.get(col_reaj, "") if col_reaj else ""),
            "pct":      _num(row.get(col_pct,  0)  if col_pct  else 0),
            "pago":     _num(row.get(col_pago, 0)  if col_pago else 0),
            "obs":      _txt(row.get(col_obs,  "") if col_obs  else ""),
        })
    return linhas


def _detectar_modo(financeiro, itens):
    fin = financeiro.get("linhas_mensais_preenchidas", 0) > 0
    its = itens.get("itens_cadastrados", 0) > 0
    rem = itens.get("itens_com_rem_atual", 0) > 0
    if fin and its and rem: return "Completo: Financeiro + Remanescentes"
    if fin and rem:         return "Financeiro + Remanescente atual"
    if fin:                 return "Financeiro historico"
    if its and rem:         return "Itemizado com remanescente"
    if its:                 return "Itemizado parcial"
    return "Base insuficiente"


def _gerar_alertas(params, ciclos, financeiro, itens, aditivos):
    alertas = []
    if not ciclos:
        alertas.append("Nenhum ciclo identificado na aba BASE.")
    if financeiro.get("linhas_mensais_preenchidas", 0) == 0:
        alertas.append("Sem execucao financeira preenchida.")
    if itens.get("itens_cadastrados", 0) == 0:
        alertas.append("Sem itens contratados preenchidos.")
    if params.get("Ciclo atual/corte"):
        pass  # ok, ciclo de corte identificado
    return alertas




# ── Memória do VTA - Matriz 2.0 ───────────────────────────────────

def _m20_ciclo_num(ciclo):
    m = re.search(r"C\s*(\d+)", str(ciclo or "").upper())
    return int(m.group(1)) if m else None


def _m20_chave_num(d, chaves, pad=0.0):
    if not isinstance(d, dict):
        return pad
    for k in chaves:
        if k in d:
            return _num(d.get(k), pad)
    # busca tolerante por nome normalizado
    alvo = [re.sub(r"[^a-z0-9]", "", str(k).lower()) for k in chaves]
    for k, v in d.items():
        kn = re.sub(r"[^a-z0-9]", "", str(k).lower())
        if any(a and a in kn for a in alvo):
            return _num(v, pad)
    return pad


def _m20_chave_txt(d, chaves, pad=""):
    if not isinstance(d, dict):
        return pad
    for k in chaves:
        if k in d:
            return _txt(d.get(k), pad)
    alvo = [re.sub(r"[^a-z0-9]", "", str(k).lower()) for k in chaves]
    for k, v in d.items():
        kn = re.sub(r"[^a-z0-9]", "", str(k).lower())
        if any(a and a in kn for a in alvo):
            return _txt(v, pad)
    return pad


def _m20_fator_ciclo(ciclos, ciclo_nome):
    for c in ciclos or []:
        if not isinstance(c, dict):
            continue
        cid = (
            _m20_chave_txt(c, ["id", "ID Ciclo", "Ciclo", "ciclo"])
            or _m20_chave_txt(c, ["ID"])
        )
        if str(cid).strip().upper() == str(ciclo_nome).strip().upper():
            return _m20_chave_num(
                c,
                ["fator_acumulado", "Fator Acumulado Aplicável", "Fator Acumulado", "fat_acum", "fator"],
                1.0,
            )
    return 1.0


def _m20_financeiro_ciclo(financeiro, ciclo_nome):
    por_ciclo = financeiro.get("por_ciclo", {}) if isinstance(financeiro, dict) else {}
    if not isinstance(por_ciclo, dict):
        return 0.0

    candidatos = [
        ciclo_nome,
        str(ciclo_nome).upper(),
        str(ciclo_nome).lower(),
        str(ciclo_nome).replace(" ", ""),
    ]

    dados = None
    for c in candidatos:
        if c in por_ciclo:
            dados = por_ciclo[c]
            break

    if dados is None:
        # busca tolerante
        cn = str(ciclo_nome).replace(" ", "").upper()
        for k, v in por_ciclo.items():
            if str(k).replace(" ", "").upper() == cn:
                dados = v
                break

    if dados is None:
        return 0.0

    if isinstance(dados, (int, float)):
        return float(dados or 0)

    if isinstance(dados, dict):
        return _m20_chave_num(
            dados,
            [
                "total_reajustado",
                "valor_reajustado",
                "valor_atualizado",
                "atualizado",
                "reajustado",
                "valor",
                "total",
                "nominal",
            ],
            0.0,
        )

    return _num(dados, 0.0)


def _m20_rem_item(item, n):
    if not isinstance(item, dict):
        return None

    # dicionários internos de remanescentes
    for k in ["remanescentes", "remanescente", "rems", "rem"]:
        v = item.get(k)
        if isinstance(v, dict):
            for kk in [f"C{n}", f"c{n}", n, str(n)]:
                if kk in v:
                    return _num(v.get(kk), None)

    # chaves diretas
    candidatos = [
        f"rem_c{n}",
        f"reman_c{n}",
        f"rem_ini_c{n}",
        f"remanescente_c{n}",
        f"remanescente_inicio_c{n}",
        f"Remanescente Inicio C{n}",
        f"Remanescente Início C{n}",
        f"Remanescente C{n}",
    ]

    for k in candidatos:
        if k in item:
            return _num(item.get(k), None)

    alvo = [re.sub(r"[^a-z0-9]", "", c.lower()) for c in candidatos]
    for k, v in item.items():
        kn = re.sub(r"[^a-z0-9]", "", str(k).lower())
        if any(a and a in kn for a in alvo):
            return _num(v, None)

    return None


def _m20_rem_atual_item(item):
    if not isinstance(item, dict):
        return None
    return _m20_chave_num(
        item,
        [
            "rem_atual",
            "remanescente_atual",
            "Remanescente Atual",
            "qtd_remanescente_atual",
            "saldo_atual",
        ],
        None,
    )


def _m20_consumo_itemizado_ciclo(itens, ciclo_num, fator):
    """Fallback por itens quando não houver financeiro do ciclo."""
    total = 0.0
    encontrou_base = False

    for item in (itens.get("itens", []) if isinstance(itens, dict) else []):
        if not isinstance(item, dict):
            continue

        vu = _m20_chave_num(item, ["vu_c0", "VU C0", "valor_unitario", "valor unitario", "vu"], 0.0)
        if not vu:
            continue

        rem_ini = _m20_rem_item(item, ciclo_num)

        if ciclo_num >= 4:
            rem_fim = _m20_rem_atual_item(item)
        else:
            rem_fim = _m20_rem_item(item, ciclo_num + 1)

        if rem_ini is None or rem_fim is None:
            continue

        encontrou_base = True
        consumo = max(0.0, float(rem_ini) - float(rem_fim))
        total += consumo * vu * fator

    return encontrou_base, round(total, 2)


def _m20_ciclo_corte(params, ciclos, financeiro):
    # 1. tenta parâmetro explícito
    if isinstance(params, dict):
        for k, v in params.items():
            kn = str(k).lower()
            if "ciclo" in kn and ("corte" in kn or "atual" in kn):
                n = _m20_ciclo_num(v)
                if n is not None:
                    return n

    # 2. tenta maior ciclo com financeiro preenchido
    nums = []
    por_ciclo = financeiro.get("por_ciclo", {}) if isinstance(financeiro, dict) else {}
    if isinstance(por_ciclo, dict):
        for k, v in por_ciclo.items():
            n = _m20_ciclo_num(k)
            if n is not None and _m20_financeiro_ciclo(financeiro, f"C{n}") > 0:
                nums.append(n)

    # 3. tenta maior ciclo da BASE
    for c in ciclos or []:
        if isinstance(c, dict):
            n = _m20_ciclo_num(
                _m20_chave_txt(c, ["id", "ID Ciclo", "Ciclo", "ciclo"], "")
            )
            if n is not None:
                nums.append(n)

    return max(nums) if nums else 0


def _montar_memoria_vta_m20(params, ciclos, financeiro, itens, aditivos):
    """
    Regra Matriz 2.0:
    - C0 sempre compõe o VTA.
    - 'Possui Efeito Financeiro?' não exclui ciclo do VTA.
    - Fonte preferencial: financeiro.
    - Na ausência de financeiro: tenta itens.
    - Ciclo em execução: financeiro parcial + remanescente atual por itens.
    """
    corte = _m20_ciclo_corte(params, ciclos, financeiro)

    linhas = []
    total_vta = 0.0

    for n in range(0, corte + 1):
        ciclo = f"C{n}"
        fator = _m20_fator_ciclo(ciclos, ciclo)
        valor_fin = round(_m20_financeiro_ciclo(financeiro, ciclo), 2)

        if valor_fin > 0:
            fonte = "Financeiro"
            criterio = "Execução financeira do ciclo incluída no VTA"
            valor = valor_fin
            status = "Incluído"
            obs = "Fonte padrão do VTA."
        else:
            tem_itens, valor_itens = _m20_consumo_itemizado_ciclo(itens, n, fator)
            if tem_itens and valor_itens > 0:
                fonte = "Itens"
                criterio = "Financeiro ausente; execução estimada por consumo de itens"
                valor = valor_itens
                status = "Incluído"
                obs = "Fallback itemizado."
            else:
                fonte = "Ausente"
                criterio = "Sem financeiro e sem dados itemizados suficientes"
                valor = 0.0
                status = "Não incluído por falta de dados"
                obs = "Requer complementação da base."

        total_vta += valor
        linhas.append({
            "Componente": ciclo,
            "Origem": fonte,
            "Critério": criterio,
            "Valor": round(valor, 2),
            "Status": status,
            "Observação": obs,
        })

    # Remanescente atual: entra sempre que informado, especialmente em ciclo em execução.
    rem_atual = round(_num(itens.get("total_rem_atual_atualizado", 0) if isinstance(itens, dict) else 0, 0), 2)
    if rem_atual > 0:
        total_vta += rem_atual
        linhas.append({
            "Componente": "Remanescente atual",
            "Origem": "Itens",
            "Critério": "Saldo remanescente atual/parcial atualizado",
            "Valor": rem_atual,
            "Status": "Incluído",
            "Observação": "Usado para ciclo em execução ou saldo a executar.",
        })

    adit = round(_num(aditivos.get("total_computavel", 0) if isinstance(aditivos, dict) else 0, 0), 2)
    if adit != 0:
        total_vta += adit
        linhas.append({
            "Componente": "Aditivos/supressões computáveis",
            "Origem": "ADITIVOS",
            "Critério": "Tratamento = computar nesta análise",
            "Valor": adit,
            "Status": "Incluído",
            "Observação": "Supressões devem entrar com sinal negativo.",
        })

    total_vta = round(total_vta, 2)

    linhas.append({
        "Componente": "TOTAL",
        "Origem": "Consolidação Matriz 2.0",
        "Critério": "Soma dos componentes incluídos",
        "Valor": total_vta,
        "Status": "VTA",
        "Observação": "Valor Total Atualizado do Contrato.",
    })

    df = pd.DataFrame(linhas, columns=["Componente", "Origem", "Critério", "Valor", "Status", "Observação"])
    return linhas, df, total_vta


# ── Funcao publica ────────────────────────────────────────────────

def ler_matriz_2_0(bytes_xlsx):
    """
    Le a Matriz 2.0 (nite.xlsx / ColetaReajuste.xlsx).
    Retorna dict compativel com o processamento do cl8us.
    """
    resultado = {
        "ok": False, "erro": "", "versao": "matriz_2_0",
        "abas_encontradas": [], "parametros": {},
        "ciclos": [], "financeiro": {}, "itens": {},
        "aditivos": {}, "historico": [],
        "modo_detectado": "Base insuficiente", "alertas": [],
        # Aliases para compatibilidade com _equalizacao_base
        "linhas_preenchidas": 0,
    }

    try:
        xls = pd.ExcelFile(BytesIO(bytes_xlsx))
    except Exception as exc:
        resultado["erro"] = str(exc)
        return resultado

    abas = set(xls.sheet_names)
    resultado["abas_encontradas"] = sorted(abas)

    mapa_ciclos = _ler_mapa_ciclos(xls, abas)
    params, ciclos = _ler_base(xls, abas)
    financeiro = _ler_execucao(xls, abas, mapa_ciclos)
    itens      = _ler_itens(xls, abas)
    aditivos   = _ler_aditivos(xls, abas)
    historico  = _ler_historico(xls, abas)

    memoria_vta_m20, df_composicao_valor_total, total_vta_m20 = _montar_memoria_vta_m20(params, ciclos, financeiro, itens, aditivos)
    modo    = _detectar_modo(financeiro, itens)
    alertas = _gerar_alertas(params, ciclos, financeiro, itens, aditivos)

    resultado.update({
        "ok": True,
        "parametros":   params,
        "ciclos":       ciclos,
        "financeiro":   financeiro,
        "itens":        itens,
        "aditivos":     aditivos,
        "historico":    historico,
        "modo_detectado": modo,
        "alertas":      alertas,
        "memoria_vta_m20": memoria_vta_m20,
        "df_composicao_valor_total": df_composicao_valor_total,
        "valor_atualizado_contrato": total_vta_m20,
        # Aliases de compatibilidade
        "linhas_preenchidas": financeiro.get("linhas_mensais_preenchidas", 0),
    })
    return resultado


def eh_matriz_2_0(abas_set):
    """Retorna True se o XLSX tem assinatura da Matriz 2.0."""
    return bool(ASSINATURA_M20 & abas_set)
