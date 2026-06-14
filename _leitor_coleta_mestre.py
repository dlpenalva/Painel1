"""
_leitor_coleta_mestre.py
------------------------
Leitor da Coleta Mestre v10 (cl8us v2.0).

Estrutura esperada (6 abas):
    PARAMETROS_CONTRATO  — lida por posição de linha (A=label, C=valor)
    CICLOS               — header linha 5 com tags [CALC]/[GCC]
    FINANCEIRO           — linhas 5-6 consolidadas C0/C1 + mensais 8-207
    ITENS                — matricial: linha por item, colunas por ciclo
    ADITIVOS             — header linha 5, dados 6-55
    DIAGNOSTICO          — automático (não lido — gerado pelo próprio leitor)

Retorno de ler_coleta_mestre(bytes_xlsx) → dict:
    ok                bool
    erro              str
    abas_encontradas  list
    abas_ausentes     list
    parametros        dict   — campos por posição de linha
    ciclos            list   — um dict por ciclo
    financeiro        dict   — consolidados + mensais + totais
    itens             dict   — base C0 + remanescentes por ciclo
    aditivos          dict   — linhas + total computável
    modo_declarado    str    — o que GCC marcou em PARAMETROS linha 21
    modo_detectado    str    — detectado automaticamente
    alertas           list
    ajustes_gcc       int    — células com nota técnica preenchida
"""

import re
import unicodedata
from io import BytesIO

import pandas as pd


ABAS_ESPERADAS = [
    "PARAMETROS_CONTRATO",
    "CICLOS",
    "FINANCEIRO",
    "ITENS",
    "ADITIVOS",
    "DIAGNOSTICO",
]

# Mapa de posições fixas em PARAMETROS_CONTRATO (linha → chave)
PARAMETROS_POSICAO = {
    4:  "vigencia_inicial",
    5:  "vigencia_final",
    7:  "indice_contratual",
    8:  "data_base_original",
    9:  "valor_original_contrato",
    10: "ultimo_ciclo_formalizado",
    11: "ciclo_inicial_analise",
    12: "ciclo_atual_corte",
    13: "competencia_corte",
    14: "variacao_acumulada",
    15: "fator_acumulado_efetivo",
    17: "ha_aditivos",
    18: "ha_corte_operacional",
    21: "modo_declarado",
    23: "ha_ciclo_em_execucao",
    24: "ciclo_execucao_fiscal_informou",
    25: "competencia_corte_execucao",
    26: "c0_executado_manual",
    27: "saldo_rem_original_corte",
    28: "saldo_rem_atualizado_corte",
    29: "saldo_inclui_aditivos",
}


# ─────────────────────────────────────────────────────────────────
# Utilitários
# ─────────────────────────────────────────────────────────────────

def _flt(valor, padrao=0.0):
    if valor is None:
        return padrao
    try:
        if pd.isna(valor):
            return padrao
    except Exception:
        pass
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = (str(valor).strip()
             .replace("R$","").replace("\xa0","").replace(" ",""))
    if "," in texto and "." in texto:
        texto = texto.replace(".","").replace(",",".")
    elif "," in texto:
        texto = texto.replace(",",".")
    try:
        return float(texto)
    except Exception:
        return padrao


def _txt(valor, padrao=""):
    if valor is None:
        return padrao
    try:
        if pd.isna(valor):
            return padrao
    except Exception:
        pass
    t = str(valor).strip()
    return t if t.lower() not in ("nan","none","nat","<na>","") else padrao


def _strip_tag(texto):
    """Remove tags [CALC], [GCC], [AUTO], [FISCAL], [AJUSTE], [TRAT]."""
    return re.sub(r"\[.*?\]\s*\n?", "", str(texto)).strip()


def _col(df, opcoes):
    """Localiza coluna por nome normalizado (com ou sem tag)."""
    def norm(s):
        s = _strip_tag(str(s))
        s = unicodedata.normalize("NFKD", s)
        s = "".join(c for c in s if not unicodedata.combining(c))
        return re.sub(r"[^a-z0-9]+","_", s.lower()).strip("_")

    mapa = {norm(c): c for c in df.columns}
    for op in opcoes:
        k = norm(op)
        if k in mapa:
            return mapa[k]
    # Parcial
    for op in opcoes:
        k = norm(op)
        for col_n, col_o in mapa.items():
            if k and k in col_n:
                return col_o
    return None


# ─────────────────────────────────────────────────────────────────
# Leitores por aba
# ─────────────────────────────────────────────────────────────────

def _ler_parametros(xls, abas):
    """Lê PARAMETROS_CONTRATO por posição fixa de linha (col A=label, C=valor)."""
    if "PARAMETROS_CONTRATO" not in abas:
        return {}
    try:
        df = pd.read_excel(xls, sheet_name="PARAMETROS_CONTRATO",
                           header=None, dtype=str)
    except Exception:
        return {}

    params = {}
    for linha, chave in PARAMETROS_POSICAO.items():
        idx = linha - 1  # 0-indexed
        if idx >= len(df):
            continue
        # Valor sempre na coluna C (índice 2)
        val = df.iloc[idx, 2] if df.shape[1] > 2 else None
        params[chave] = _txt(val)

    # Converter numéricos
    for campo in ["valor_original_contrato","variacao_acumulada",
                  "fator_acumulado_efetivo","c0_executado_manual",
                  "saldo_rem_original_corte","saldo_rem_atualizado_corte"]:
        if campo in params:
            params[campo] = _flt(params[campo])

    return params


def _ler_ciclos(xls, abas):
    """Lê CICLOS — header linha 5 (0-indexed=4), dados linhas 6+."""
    if "CICLOS" not in abas:
        return []
    try:
        df = pd.read_excel(xls, sheet_name="CICLOS", header=4)
    except Exception:
        return []

    df.columns = [_strip_tag(c) for c in df.columns]
    df = df.dropna(how="all")

    col_ciclo    = _col(df, ["Ciclo"])
    col_base     = _col(df, ["Data-base","Data base"])
    col_ini      = _col(df, ["Início financeiro","Inicio financeiro","Início fin","Inicio fin"])
    col_fim      = _col(df, ["Fim financeiro","Fim fin"])
    col_pct      = _col(df, ["Percentual aplicado","Percentual"])
    col_fator    = _col(df, ["Fator do ciclo","Fator"])
    col_fat_acum = _col(df, ["Fator acumulado efetivo","Fator acumulado","Fator acum"])
    col_sit      = _col(df, ["Situação"])
    col_objeto   = _col(df, ["É objeto","objeto analise"])
    col_efeito   = _col(df, ["Tem efeito financeiro","efeito financeiro"])
    col_entra    = _col(df, ["Entra no Valor Total","Valor Total"])
    col_obs      = _col(df, ["Observação","Nota técnica"])

    ciclos = []
    for _, row in df.iterrows():
        nome = _txt(row.get(col_ciclo,"") if col_ciclo else "").upper()
        if not nome or not re.match(r"^C\d+$", nome):
            continue
        ciclos.append({
            "ciclo":             nome,
            "data_base":         _txt(row.get(col_base,"")     if col_base     else ""),
            "inicio_financeiro": _txt(row.get(col_ini,"")      if col_ini      else ""),
            "fim_financeiro":    _txt(row.get(col_fim,"")      if col_fim      else ""),
            "percentual":        _flt(row.get(col_pct,0)       if col_pct      else 0),
            "fator_ciclo":       max(_flt(row.get(col_fator,1) if col_fator    else 1, 1.0), 0.0001),
            "fator_acumulado":   max(_flt(row.get(col_fat_acum,1) if col_fat_acum else 1, 1.0), 0.0001),
            "situacao":          _txt(row.get(col_sit,"")      if col_sit      else ""),
            "objeto_analise":    _txt(row.get(col_objeto,"")   if col_objeto   else ""),
            "tem_efeito_fin":    _txt(row.get(col_efeito,"")   if col_efeito   else "Sim"),
            "entra_valor_total": _txt(row.get(col_entra,"Sim") if col_entra    else "Sim"),
            "observacao":        _txt(row.get(col_obs,"")      if col_obs      else ""),
        })
    return ciclos


def _ler_financeiro(xls, abas):
    """
    Lê FINANCEIRO.
    Linhas 5-6 (idx 4-5): C0 e C1 consolidados (col A=ciclo, C=valor, D=efeito).
    Linha 7 (idx 6): separador — ignorar.
    Linhas 8-207 (idx 7-206): mensais.
    Header na linha 4 (idx 3).
    """
    resultado = {
        "consolidados": [],           # C0 e C1 fixos
        "mensais": [],                # linhas mensais
        "total_com_efeito": 0.0,
        "total_sem_efeito": 0.0,
        "total_geral": 0.0,
        "linhas_mensais_preenchidas": 0,
        "ajustes_gcc": 0,
    }
    if "FINANCEIRO" not in abas:
        return resultado

    try:
        df_raw = pd.read_excel(xls, sheet_name="FINANCEIRO",
                               header=None, dtype=object)
    except Exception:
        return resultado

    # Linhas consolidadas — C0 fixo na linha 5; C1+ apenas se realmente existir
    import re as _re_cons
    _ciclo_pattern = _re_cons.compile(r'^C\d+$', _re_cons.IGNORECASE)
    for idx in range(4, min(4 + 6, len(df_raw))):
        row = df_raw.iloc[idx]
        ciclo_raw = _txt(row.iloc[0] if len(row)>0 else None)
        # Parar se não for um ciclo real (C0, C1, C2...)
        if not _ciclo_pattern.match(ciclo_raw):
            break
        valor  = _flt(row.iloc[2] if len(row)>2 else None)
        efeito = _txt(row.iloc[3] if len(row)>3 else None, "Não")
        ajuste = _txt(row.iloc[4] if len(row)>4 else None)
        resultado["consolidados"].append({
            "ciclo":     ciclo_raw.upper(),
            "valor":     valor,
            "efeito":    efeito,
            "ajuste":    ajuste,
            "_ajustado": bool(ajuste),
        })

    # Header na linha 4 (idx 3); dados mensais a partir da linha 8 (idx 7)
    try:
        df = pd.read_excel(xls, sheet_name="FINANCEIRO", header=3, dtype=object)
    except Exception:
        return resultado

    df.columns = [_strip_tag(c) for c in df.columns]
    df = df.dropna(how="all")

    col_ciclo  = _col(df, ["Ciclo"])
    col_comp   = _col(df, ["Competência","Competencia"])
    col_valor  = _col(df, ["Valor Informado","Valor informado","Valor"])
    col_efeito = _col(df, ["Tem efeito financeiro","efeito financeiro"])
    col_ajuste = _col(df, ["Nota técnica","Ajuste"])

    if not col_valor:
        return resultado

    mensais = []
    # Calcular skip: pular até após o separador ▼ (linha que não é dado mensal)
    skip = 0
    for _i, (_, _row) in enumerate(df.iterrows()):
        _v = str(_row.iloc[0]).strip() if len(_row) > 0 else ""
        # Separador encontrado — pular até aqui inclusive
        if "▼" in _v or "Linhas mensais" in _v:
            skip = _i + 1
            break
        # Linha consolidada (C0, C1 sem data) — pular
        import re as _re_sk
        _eh_ciclo = bool(_re_sk.match(r'^C\d+$', _v, _re_sk.IGNORECASE))
        if _eh_ciclo:
            skip = _i + 1
        else:
            break  # primeiro dado não-conso e não-separador = início dos mensais
    for i, (_, row) in enumerate(df.iterrows()):
        if i < skip:
            continue
        valor = _flt(row.get(col_valor, None))
        if abs(valor) < 0.004:
            continue
        efeito = _txt(row.get(col_efeito, "Sim") if col_efeito else "Sim")
        ajuste = _txt(row.get(col_ajuste, "") if col_ajuste else "")
        mensais.append({
            "ciclo":      _txt(row.get(col_ciclo, "") if col_ciclo else ""),
            "competencia":_txt(row.get(col_comp,  "") if col_comp  else ""),
            "valor":      valor,
            "efeito":     efeito,
            "ajuste":     ajuste,
            "_ajustado":  bool(ajuste),
        })

    resultado["mensais"] = mensais
    resultado["linhas_mensais_preenchidas"] = len(mensais)

    # Totais
    tot_com = sum(m["valor"] for m in mensais if m["efeito"].lower()=="sim")
    tot_sem = sum(m["valor"] for m in mensais if m["efeito"].lower()!="sim")
    for c in resultado["consolidados"]:
        if c["efeito"].lower() == "sim":
            tot_com += c["valor"]
        else:
            tot_sem += c["valor"]

    resultado["total_com_efeito"] = round(tot_com, 2)
    resultado["total_sem_efeito"] = round(tot_sem, 2)
    resultado["total_geral"]      = round(tot_com + tot_sem, 2)
    resultado["ajustes_gcc"]      = sum(1 for m in mensais if m["_ajustado"])
    return resultado


def _ler_itens(xls, abas):
    """
    Lê ITENS — estrutura matricial v10.
    Header linha 7 (idx 6).
    Linha 5 (idx 4): datas início de cada ciclo — usada como metadado.
    Dados linhas 8-207 (idx 7-206).
    Colunas: A=Item, B=Qtd C0, C=VU C0, D=VT C0(auto),
             E=Rem C1, F=Rem C2, G=Rem C3, H=Rem C4, I=Rem C5,
             J=Rem atual, K=VT Rem atual(auto), L=Fator, M=VT atualiz(auto),
             N=Ajuste.
    """
    resultado = {
        "itens": [],
        "total_c0": 0.0,
        "total_rem_atual_original": 0.0,
        "total_rem_atual_atualizado": 0.0,
        "itens_cadastrados": 0,
        "itens_com_rem_c1": 0,
        "itens_com_rem_atual": 0,
        "ajustes_gcc": 0,
    }
    if "ITENS" not in abas:
        return resultado

    try:
        df = pd.read_excel(xls, sheet_name="ITENS", header=6, dtype=object)
    except Exception:
        return resultado

    df.columns = [_strip_tag(c) for c in df.columns]
    df = df.dropna(how="all")

    col_item   = _col(df, ["Item"])
    col_qtd_c0 = _col(df, ["Qtd C0","Qtd"])
    col_vu_c0  = _col(df, ["VU C0","Valor unitário"])
    col_vt_c0  = _col(df, ["VT C0","Valor total"])
    col_rem_c1 = _col(df, ["Rem. C1","Rem C1","Remanescente C1"])
    col_rem_c2 = _col(df, ["Rem. C2","Rem C2","Remanescente C2"])
    col_rem_c3 = _col(df, ["Rem. C3","Rem C3","Remanescente C3"])
    col_rem_c4 = _col(df, ["Rem. C4","Rem C4","Remanescente C4"])
    col_rem_c5 = _col(df, ["Rem. C5","Rem C5","Remanescente C5"])
    col_rem_at = _col(df, ["Rem. atual","Rem atual","Remanescente atual"])
    col_vt_rem = _col(df, ["VT Rem. atual","VT Rem atual"])
    col_fator  = _col(df, ["Fator"])
    col_vt_atu = _col(df, ["VT Rem. atualiz","VT Rem atualiz","Valor atualiz"])
    col_ajuste = _col(df, ["Nota técnica","Ajuste"])

    itens = []
    total_c0 = 0.0
    total_rem_orig = 0.0
    total_rem_atu  = 0.0

    for _, row in df.iterrows():
        item = _txt(row.get(col_item,"") if col_item else "")
        if not item or item.upper() in ("TOTAL","TOTAIS","[GCC]"):
            continue

        qtd_c0 = _flt(row.get(col_qtd_c0,0) if col_qtd_c0 else 0)
        vu_c0  = _flt(row.get(col_vu_c0, 0) if col_vu_c0  else 0)
        vt_c0  = _flt(row.get(col_vt_c0, 0) if col_vt_c0  else 0)
        if vt_c0 == 0 and qtd_c0 > 0 and vu_c0 > 0:
            vt_c0 = round(qtd_c0 * vu_c0, 2)

        rem_c1 = _flt(row.get(col_rem_c1,0) if col_rem_c1 else 0)
        rem_c2 = _flt(row.get(col_rem_c2,0) if col_rem_c2 else 0)
        rem_c3 = _flt(row.get(col_rem_c3,0) if col_rem_c3 else 0)
        rem_c4 = _flt(row.get(col_rem_c4,0) if col_rem_c4 else 0)
        rem_c5 = _flt(row.get(col_rem_c5,0) if col_rem_c5 else 0)
        rem_at = _flt(row.get(col_rem_at,0) if col_rem_at else 0)
        vt_rem = _flt(row.get(col_vt_rem,0) if col_vt_rem else 0)
        fator  = max(_flt(row.get(col_fator,1.0) if col_fator else 1.0, 1.0), 0.0001)
        vt_atu = _flt(row.get(col_vt_atu,0) if col_vt_atu else 0)
        ajuste = _txt(row.get(col_ajuste,"") if col_ajuste else "")

        # Calcular VT Rem se não disponível (dados_only sem fórmula)
        if vt_rem == 0 and rem_at > 0 and vu_c0 > 0:
            vt_rem = round(rem_at * vu_c0, 2)
        if vt_atu == 0 and vt_rem > 0:
            vt_atu = round(vt_rem * fator, 2)

        total_c0      += vt_c0
        total_rem_orig += vt_rem
        total_rem_atu  += vt_atu

        itens.append({
            "item":      item,
            "qtd_c0":    qtd_c0,
            "vu_c0":     vu_c0,
            "vt_c0":     vt_c0,
            "rem_c1":    rem_c1,
            "rem_c2":    rem_c2,
            "rem_c3":    rem_c3,
            "rem_c4":    rem_c4,
            "rem_c5":    rem_c5,
            "rem_atual": rem_at,
            "vt_rem_atual_orig": vt_rem,
            "fator":     fator,
            "vt_rem_atual_atu":  vt_atu,
            "ajuste":    ajuste,
            "_tem_rem_c1":   rem_c1 > 0.001,
            "_tem_rem_atual": rem_at > 0.001,
            "_ajustado":  bool(ajuste),
        })

    resultado.update({
        "itens":                    itens,
        "total_c0":                 round(total_c0,      2),
        "total_rem_atual_original": round(total_rem_orig, 2),
        "total_rem_atual_atualizado": round(total_rem_atu, 2),
        "itens_cadastrados":        len(itens),
        "itens_com_rem_c1":         sum(1 for i in itens if i["_tem_rem_c1"]),
        "itens_com_rem_atual":      sum(1 for i in itens if i["_tem_rem_atual"]),
        "ajustes_gcc":              sum(1 for i in itens if i["_ajustado"]),
    })
    return resultado


def _ler_aditivos(xls, abas):
    """Lê ADITIVOS — header linha 5 (idx 4), dados 6-55."""
    resultado = {"linhas":[], "total_computavel":0.0, "qtd_computaveis":0}
    if "ADITIVOS" not in abas:
        return resultado

    try:
        df = pd.read_excel(xls, sheet_name="ADITIVOS", header=4, dtype=object)
    except Exception:
        return resultado

    df.columns = [_strip_tag(c) for c in df.columns]
    df = df.dropna(how="all")

    col_item   = _col(df, ["Item"])
    col_data   = _col(df, ["Data do aditivo","Data"])
    col_ciclo  = _col(df, ["Ciclo","Marco"])
    col_tipo   = _col(df, ["Tipo de alteração","Tipo"])
    col_qtd    = _col(df, ["Qtd acrescida","Qtd suprimida","Quantidade"])
    col_vu     = _col(df, ["Valor unitário original","Valor unitario"])
    col_vt_ori = _col(df, ["Valor original da alteração","Valor original"])
    col_aplic  = _col(df, ["Aplicar reajuste","reajuste acumulado"])
    col_fator  = _col(df, ["Fator acumulado aplicável","Fator"])
    col_vt_atu = _col(df, ["Valor atualizado da alteração","Valor atualizado"])
    col_trat   = _col(df, ["Tratamento do aditivo","Tratamento"])

    linhas = []
    total_comp = 0.0

    for _, row in df.iterrows():
        tipo    = _txt(row.get(col_tipo,"")    if col_tipo    else "")
        vt_ori  = _flt(row.get(col_vt_ori,0)  if col_vt_ori  else 0)
        vt_atu  = _flt(row.get(col_vt_atu,0)  if col_vt_atu  else 0)

        # Linha vazia real
        item_v = _txt(row.get(col_item,"") if col_item else "")
        data_v = _txt(row.get(col_data,"") if col_data else "")
        qtd_v  = _flt(row.get(col_qtd,0)  if col_qtd  else 0)
        if not item_v and not data_v and abs(qtd_v)<0.001 and abs(vt_ori)<0.004:
            continue
        if tipo.startswith("["):
            continue

        trat    = _txt(row.get(col_trat,"Computar nesta análise") if col_trat else "")
        computar = "computar" in trat.lower()

        linhas.append({
            "item":        item_v,
            "data":        data_v,
            "ciclo":       _txt(row.get(col_ciclo,"")  if col_ciclo  else ""),
            "tipo":        tipo,
            "quantidade":  qtd_v,
            "vu_original": _flt(row.get(col_vu,0)     if col_vu     else 0),
            "vt_original": vt_ori,
            "aplicar_reaj":_txt(row.get(col_aplic,"Sim") if col_aplic else "Sim"),
            "fator":       max(_flt(row.get(col_fator,1.0) if col_fator else 1.0, 1.0), 0.0001),
            "vt_atualizado": vt_atu,
            "tratamento":  trat,
            "_computar":   computar,
        })
        if computar:
            total_comp += vt_atu

    resultado.update({
        "linhas":           linhas,
        "total_computavel": round(total_comp, 2),
        "qtd_computaveis":  sum(1 for l in linhas if l["_computar"]),
    })
    return resultado


# ─────────────────────────────────────────────────────────────────


def _ler_roteiro_info_fiscais(xls, abas):
    """Lê a aba ROTEIRO_INFO_FISCAIS, quando existente.

    A aba é informativa e classificatória: não altera cálculo. O leitor apenas
    transforma o conteúdo em dicionário para diagnóstico, exibição e uso futuro
    pela Middle-Layer.
    """
    resultado = {
        "ok": False,
        "presente": False,
        "dados": {},
        "perguntas": {},
        "resumo": {},
        "modelo_sugerido": "",
        "base_esperada": "",
        "observacao_qualidade": "",
    }
    if "ROTEIRO_INFO_FISCAIS" not in abas:
        return resultado

    resultado["presente"] = True

    try:
        df = pd.read_excel(xls, sheet_name="ROTEIRO_INFO_FISCAIS", header=None, dtype=object)
    except Exception as exc:
        resultado["erro"] = str(exc)
        return resultado

    def _limpar(valor):
        try:
            if pd.isna(valor):
                return ""
        except Exception:
            pass
        return str(valor).strip()

    def _norm(texto):
        texto = str(texto or "").lower().strip()
        texto = texto.replace("ç", "c").replace("ã", "a").replace("á", "a").replace("à", "a")
        texto = texto.replace("â", "a").replace("é", "e").replace("ê", "e").replace("í", "i")
        texto = texto.replace("ó", "o").replace("ô", "o").replace("õ", "o").replace("ú", "u")
        return re.sub(r"\s+", " ", texto)

    dados = {}
    perguntas = {}

    for _, row in df.iterrows():
        vals = [_limpar(v) for v in row.tolist()]
        vals = [v for v in vals if v]
        if len(vals) < 2:
            continue

        chave = vals[0]
        valor = vals[1]
        chave_norm = _norm(chave)
        dados[chave] = valor

        if re.match(r"^\s*\d+[\.)\-]", chave) or "pergunta" in chave_norm:
            perguntas[chave] = valor

        if "modelo" in chave_norm and "sugerido" in chave_norm:
            resultado["modelo_sugerido"] = valor
        elif "base" in chave_norm and ("esperada" in chave_norm or "esperado" in chave_norm):
            resultado["base_esperada"] = valor
        elif ("observacao" in chave_norm or "qualidade" in chave_norm) and not resultado["observacao_qualidade"]:
            resultado["observacao_qualidade"] = valor

    resultado["dados"] = dados
    resultado["perguntas"] = perguntas
    resultado["resumo"] = {
        "modelo_sugerido": resultado.get("modelo_sugerido", ""),
        "base_esperada": resultado.get("base_esperada", ""),
        "observacao_qualidade": resultado.get("observacao_qualidade", ""),
    }
    resultado["ok"] = True
    return resultado


# Diagnóstico automático
# ─────────────────────────────────────────────────────────────────

def _detectar_modo(financeiro, itens, params):
    tem_fin      = financeiro["linhas_mensais_preenchidas"] > 0
    tem_c0_c1    = any(c["valor"] > 0 for c in financeiro["consolidados"])
    tem_rem_c1   = itens["itens_com_rem_c1"] > 0
    tem_rem_atual = itens["itens_com_rem_atual"] > 0
    tem_itens    = itens["itens_cadastrados"] > 0

    if tem_fin and tem_rem_c1 and tem_rem_atual:
        return "Completo: Financeiro + Remanescentes"
    elif tem_fin and tem_rem_atual:
        return "Financeiro + Remanescente atual"
    elif tem_fin and tem_rem_c1:
        return "Financeiro + Remanescentes históricos"
    elif tem_fin:
        return "Apenas Financeiro Histórico"
    elif tem_rem_atual and tem_rem_c1:
        return "Apenas Remanescentes (todos os ciclos)"
    elif tem_rem_atual:
        return "Apenas Saldo/Remanescente Atual"
    elif tem_itens:
        return "Itens parciais sem remanescentes"
    else:
        return "Base insuficiente"


def _gerar_alertas(params, financeiro, itens, aditivos, modo_declarado):
    alertas = []

    if not modo_declarado:
        alertas.append("⚠️ Modo não declarado em PARAMETROS — preencha antes do upload.")

    tem_fin = financeiro["linhas_mensais_preenchidas"] > 0
    c0_manual = params.get("c0_executado_manual", 0)
    c0_fin = next((c["valor"] for c in financeiro["consolidados"] if c["ciclo"]=="C0"), 0)
    if tem_fin and c0_manual == 0 and c0_fin == 0:
        alertas.append("⚠️ C0 não informado — sem C0 o Valor Total Atualizado ficará incompleto.")

    saldo_inclui = params.get("saldo_inclui_aditivos","Não")
    if saldo_inclui.lower() == "sim" and aditivos["qtd_computaveis"] > 0:
        alertas.append("⚠️ Saldo inclui aditivos E há aditivos para computar — risco de dupla contagem.")

    if (financeiro["linhas_mensais_preenchidas"] == 0 and
        itens["itens_com_rem_atual"] == 0 and
        all(c["valor"] == 0 for c in financeiro["consolidados"])):
        alertas.append("❌ Nenhuma base de execução ou saldo identificada.")

    total_ajustes = (financeiro["ajustes_gcc"] + itens["ajustes_gcc"])
    if total_ajustes > 0:
        alertas.append(f"ℹ️ {total_ajustes} ajuste(s) manual(is) GCC registrado(s) — verifique rastreabilidade.")

    modo_det = _detectar_modo(financeiro, itens, params)
    if modo_declarado and modo_declarado and "insuficiente" not in modo_det:
        pass  # consistente

    return alertas



# ─────────────────────────────────────────────────────────────────
# Leitores v11 — estrutura vertical (FINANCEIRO_COMP, ITENS_CICLO)
# ─────────────────────────────────────────────────────────────────

ABAS_V11 = {"FINANCEIRO_COMP", "ITENS_CICLO", "PARAMETROS_CICLOS"}

def _eh_v11(abas):
    """Retorna True se o XLSX tem assinatura da ColetaMestre v11."""
    return bool(ABAS_V11 & abas)


def _ler_financeiro_comp(xls, abas):
    """Lê FINANCEIRO_COMP (v11) — uma linha por competência."""
    resultado = {
        "consolidados": [],
        "mensais": [],
        "total_com_efeito": 0.0,
        "total_sem_efeito": 0.0,
        "total_geral": 0.0,
        "linhas_mensais_preenchidas": 0,
        "ajustes_gcc": 0,
    }
    if "FINANCEIRO_COMP" not in abas:
        return resultado
    try:
        df = pd.read_excel(xls, sheet_name="FINANCEIRO_COMP", header=3, dtype=object)
    except Exception:
        return resultado

    df.columns = [_strip_tag(c) for c in df.columns]
    df = df.dropna(how="all")

    col_ciclo  = _col(df, ["Ciclo"])
    col_comp   = _col(df, ["Competencia", "Competência"])
    col_valor  = _col(df, ["Valor reconhecido", "Valor Informado", "Valor"])
    col_efeito = _col(df, ["efeito financeiro", "Tem efeito"])
    col_obs    = _col(df, ["Observacao", "Observação"])

    if not col_valor:
        return resultado

    mensais = []
    for _, row in df.iterrows():
        valor = _flt(row.get(col_valor, None))
        if abs(valor) < 0.004:
            continue
        ciclo  = _txt(row.get(col_ciclo, "") if col_ciclo else "")
        efeito = _txt(row.get(col_efeito, "Sim") if col_efeito else "Sim")
        obs    = _txt(row.get(col_obs, "") if col_obs else "")
        mensais.append({
            "ciclo":       ciclo,
            "competencia": _txt(row.get(col_comp, "") if col_comp else ""),
            "valor":       valor,
            "efeito":      efeito,
            "ajuste":      obs,
            "_ajustado":   bool(obs),
        })

    resultado["mensais"] = mensais
    resultado["linhas_mensais_preenchidas"] = len(mensais)
    tot_com = sum(m["valor"] for m in mensais if m["efeito"].lower() == "sim")
    tot_sem = sum(m["valor"] for m in mensais if m["efeito"].lower() != "sim")
    resultado["total_com_efeito"] = round(tot_com, 2)
    resultado["total_sem_efeito"] = round(tot_sem, 2)
    resultado["total_geral"]      = round(tot_com + tot_sem, 2)
    resultado["ajustes_gcc"]      = sum(1 for m in mensais if m["_ajustado"])
    return resultado


def _ler_itens_ciclo(xls, abas):
    """Le ITENS_CICLO (v11) — estrutura matricial: uma linha por item."""
    resultado = {
        "itens": [], "total_c0": 0.0,
        "total_rem_atual_original": 0.0, "total_rem_atual_atualizado": 0.0,
        "itens_cadastrados": 0, "itens_com_rem_c1": 0,
        "itens_com_rem_atual": 0, "ajustes_gcc": 0,
    }
    if "ITENS_CICLO" not in abas:
        return resultado
    try:
        df = pd.read_excel(xls, sheet_name="ITENS_CICLO", header=4, dtype=object)
    except Exception:
        return resultado

    df.columns = [_strip_tag(c) for c in df.columns]
    df = df.dropna(how="all")

    col_item  = _col(df, ["Item"])
    col_desc  = _col(df, ["Descricao", "Descricao resumida"])
    col_un    = _col(df, ["Unidade"])
    col_qtd   = _col(df, ["Qtd C0", "Qtd"])
    col_vu    = _col(df, ["VU C0", "Valor unit"])
    col_vt_c0 = _col(df, ["VT C0", "Valor total"])
    col_fator = _col(df, ["Fator"])
    col_vt_rem_orig = _col(df, ["VT Rem original", "VT Rem original", "VT Rem original R$"])
    col_vt_rem_atu  = _col(df, ["VT Rem Atualiz", "VT Rem Atualiz", "VT Rem Atualiz R$"])
    col_obs   = _col(df, ["Observacao"])

    # Colunas de remanescente: todas que contem "Rem C" ou "Rem corte"
    rem_cols = {}
    for c in df.columns:
        cn = str(c).strip()
        m = re.search(r'Rem\s+(C\d+)', cn, re.IGNORECASE)
        if m:
            rem_cols[m.group(1).upper()] = c

    itens = []
    total_c0 = total_rem_orig = total_rem_atu = 0.0

    for _, row in df.iterrows():
        item = _txt(row.get(col_item, "") if col_item else "")
        if not item or item.upper() in ("TOTAL", "TOTAIS"):
            continue

        qtd_c0 = _flt(row.get(col_qtd, 0) if col_qtd else 0)
        vu_c0  = _flt(row.get(col_vu,  0) if col_vu  else 0)
        vt_c0  = _flt(row.get(col_vt_c0, 0) if col_vt_c0 else 0)
        if vt_c0 == 0 and qtd_c0 > 0 and vu_c0 > 0:
            vt_c0 = round(qtd_c0 * vu_c0, 2)

        fator      = max(_flt(row.get(col_fator, 1.0) if col_fator else 1.0, 1.0), 0.0001)
        vt_rem_ori = _flt(row.get(col_vt_rem_orig, 0) if col_vt_rem_orig else 0)
        vt_rem_atu = _flt(row.get(col_vt_rem_atu,  0) if col_vt_rem_atu  else 0)
        obs        = _txt(row.get(col_obs, "") if col_obs else "")

        # Remanescentes por ciclo
        rem_c1 = rem_c2 = rem_c3 = rem_c4 = rem_c5 = rem_atual = 0.0
        for nome_c, col_c in rem_cols.items():
            v = _flt(row.get(col_c, 0))
            if nome_c == "C1": rem_c1 = v
            elif nome_c == "C2": rem_c2 = v
            elif nome_c == "C3": rem_c3 = v
            elif nome_c == "C4": rem_c4 = v
            elif nome_c == "C5": rem_c5 = v
        # Ultimo remanescente disponivel = rem_atual (corte)
        for v in [rem_c5, rem_c4, rem_c3, rem_c2, rem_c1]:
            if v > 0.001:
                rem_atual = v
                break

        # Calcular VT Rem se formula nao resolvida
        if vt_rem_ori == 0 and rem_atual > 0 and vu_c0 > 0:
            vt_rem_ori = round(rem_atual * vu_c0, 2)
        if vt_rem_atu == 0 and vt_rem_ori > 0:
            vt_rem_atu = round(vt_rem_ori * fator, 2)

        total_c0      += vt_c0
        total_rem_orig += vt_rem_ori
        total_rem_atu  += vt_rem_atu

        itens.append({
            "item": item,
            "qtd_c0": qtd_c0, "vu_c0": vu_c0, "vt_c0": vt_c0,
            "rem_c1": rem_c1, "rem_c2": rem_c2, "rem_c3": rem_c3,
            "rem_c4": rem_c4, "rem_c5": rem_c5,
            "rem_atual": rem_atual,
            "vt_rem_atual_orig": vt_rem_ori,
            "fator": fator,
            "vt_rem_atual_atu": vt_rem_atu,
            "ajuste": obs,
            "_tem_rem_c1":    rem_c1 > 0.001,
            "_tem_rem_atual": rem_atual > 0.001,
            "_ajustado":      bool(obs),
        })

    resultado.update({
        "itens": itens,
        "total_c0":                   round(total_c0, 2),
        "total_rem_atual_original":   round(total_rem_orig, 2),
        "total_rem_atual_atualizado": round(total_rem_atu, 2),
        "itens_cadastrados":          len(itens),
        "itens_com_rem_c1":           sum(1 for i in itens if i["_tem_rem_c1"]),
        "itens_com_rem_atual":        sum(1 for i in itens if i["_tem_rem_atual"]),
        "ajustes_gcc":                sum(1 for i in itens if i["_ajustado"]),
    })
    return resultado


def _ler_aditivos_sup(xls, abas):
    """Lê ADITIVOS_SUP (v11)."""
    resultado = {"linhas": [], "total_computavel": 0.0, "qtd_computaveis": 0}
    if "ADITIVOS_SUP" not in abas:
        return resultado
    try:
        df = pd.read_excel(xls, sheet_name="ADITIVOS_SUP", header=3, dtype=object)
    except Exception:
        return resultado

    df.columns = [_strip_tag(c) for c in df.columns]
    df = df.dropna(how="all")

    col_data   = _col(df, ["Data"])
    col_ciclo  = _col(df, ["Ciclo"])
    col_tipo   = _col(df, ["Tipo"])
    col_item   = _col(df, ["Item"])
    col_qtd    = _col(df, ["Quantidade"])
    col_vu     = _col(df, ["Valor unit", "VU"])
    col_vt     = _col(df, ["Valor original"])
    col_trat   = _col(df, ["Tratamento"])

    linhas = []
    total_comp = 0.0
    for _, row in df.iterrows():
        vt  = _flt(row.get(col_vt, 0) if col_vt else 0)
        qtd = _flt(row.get(col_qtd, 0) if col_qtd else 0)
        if abs(vt) < 0.004 and abs(qtd) < 0.001:
            continue
        trat     = _txt(row.get(col_trat, "Nao sei") if col_trat else "Nao sei")
        computar = "computar" in trat.lower() and "parte" in trat.lower()
        linhas.append({
            "data":        _txt(row.get(col_data, "") if col_data else ""),
            "ciclo":       _txt(row.get(col_ciclo, "") if col_ciclo else ""),
            "tipo":        _txt(row.get(col_tipo, "") if col_tipo else ""),
            "item":        _txt(row.get(col_item, "") if col_item else ""),
            "quantidade":  qtd,
            "vu_original": _flt(row.get(col_vu, 0) if col_vu else 0),
            "vt_original": vt,
            "tratamento":  trat,
            "_computar":   computar,
        })
        if computar:
            total_comp += vt
    resultado.update({
        "linhas":           linhas,
        "total_computavel": round(total_comp, 2),
        "qtd_computaveis":  sum(1 for l in linhas if l["_computar"]),
    })
    return resultado


def _ler_parametros_ciclos(xls, abas):
    """Lê PARAMETROS_CICLOS (v11) — aba técnica oculta."""
    params = {}
    ciclos = []
    if "PARAMETROS_CICLOS" not in abas:
        return params, ciclos
    try:
        df = pd.read_excel(xls, sheet_name="PARAMETROS_CICLOS", header=None, dtype=object)
    except Exception:
        return params, ciclos

    # Parâmetros: linhas 3-7 (idx 2-6), col A=label, col B=valor
    param_map = {
        2: "indice_contratual",
        3: "data_base_original",
        4: "valor_original_contrato",
        5: "vigencia_inicial",
        6: "vigencia_final",
    }
    for idx, chave in param_map.items():
        if idx >= len(df):
            continue
        val = df.iloc[idx, 1] if df.shape[1] > 1 else None
        params[chave] = _txt(val)
    for campo in ["valor_original_contrato"]:
        if campo in params:
            params[campo] = _flt(params[campo])

    # Ciclos: header linha 9 (idx 8), dados 10+ (idx 9+)
    try:
        df_c = pd.read_excel(xls, sheet_name="PARAMETROS_CICLOS", header=8, dtype=object)
        df_c.columns = [_strip_tag(c) for c in df_c.columns]
        df_c = df_c.dropna(how="all")
        col_ciclo    = _col(df_c, ["Ciclo"])
        col_base     = _col(df_c, ["Data-base", "Data base"])
        col_pct      = _col(df_c, ["Percentual"])
        col_fator    = _col(df_c, ["Fator ciclo", "Fator do ciclo"])
        col_fat_acum = _col(df_c, ["Fator acumulado"])
        col_sit      = _col(df_c, ["Situacao", "Situação"])
        col_ini      = _col(df_c, ["Inicio fin", "Início fin"])
        col_fim      = _col(df_c, ["Fim fin"])
        for _, row in df_c.iterrows():
            nome = _txt(row.get(col_ciclo, "") if col_ciclo else "").upper()
            if not nome or not re.match(r"^C\d+$", nome):
                continue
            ciclos.append({
                "ciclo":            nome,
                "data_base":        _txt(row.get(col_base, "") if col_base else ""),
                "inicio_financeiro":_txt(row.get(col_ini, "")  if col_ini  else ""),
                "fim_financeiro":   _txt(row.get(col_fim, "")  if col_fim  else ""),
                "percentual":       _flt(row.get(col_pct, 0)   if col_pct  else 0),
                "fator_ciclo":      max(_flt(row.get(col_fator, 1) if col_fator else 1, 1.0), 0.0001),
                "fator_acumulado":  max(_flt(row.get(col_fat_acum, 1) if col_fat_acum else 1, 1.0), 0.0001),
                "situacao":         _txt(row.get(col_sit, "")  if col_sit  else ""),
                "objeto_analise":   "Sim",
                "tem_efeito_fin":   "Sim",
                "entra_valor_total":"Sim",
                "observacao":       "",
            })
    except Exception:
        pass
    return params, ciclos


def _ler_v11(bytes_xlsx):
    """Lê ColetaMestre v11 e retorna dict no mesmo formato do ler_coleta_mestre (v10)."""
    resultado = {
        "ok": False, "erro": "", "versao": "v11",
        "abas_encontradas": [], "abas_ausentes": [],
        "parametros": {}, "ciclos": [],
        "financeiro": {}, "itens": {}, "aditivos": {},
        "modo_declarado": "", "modo_detectado": "Base insuficiente",
        "alertas": [], "ajustes_gcc": 0,
        "roteiro_info_fiscais": {"ok": False, "presente": False},
    }
    try:
        xls = pd.ExcelFile(BytesIO(bytes_xlsx))
    except Exception as exc:
        resultado["erro"] = str(exc)
        return resultado

    abas = set(xls.sheet_names)
    resultado["abas_encontradas"] = sorted(abas)

    params, ciclos = _ler_parametros_ciclos(xls, abas)
    financeiro     = _ler_financeiro_comp(xls, abas)
    itens          = _ler_itens_ciclo(xls, abas)
    aditivos       = _ler_aditivos_sup(xls, abas)
    modo_detectado = _detectar_modo(financeiro, itens, params)
    alertas        = _gerar_alertas(params, financeiro, itens, aditivos, "")
    ajustes        = financeiro.get("ajustes_gcc", 0) + itens.get("ajustes_gcc", 0)

    resultado.update({
        "ok":             True,
        "parametros":     params,
        "ciclos":         ciclos,
        "financeiro":     financeiro,
        "itens":          itens,
        "aditivos":       aditivos,
        "modo_detectado": modo_detectado,
        "alertas":        alertas,
        "ajustes_gcc":    ajustes,
    })
    return resultado

# ─────────────────────────────────────────────────────────────────
# Função pública
# ─────────────────────────────────────────────────────────────────

def ler_coleta_mestre(bytes_xlsx):
    """
    Lê ColetaMestre v10 e retorna dicionário estruturado.

    Parâmetro:
        bytes_xlsx: bytes — conteúdo binário do XLSX.
    """
    resultado = {
        "ok":               False,
        "erro":             "",
        "abas_encontradas": [],
        "abas_ausentes":    [],
        "parametros":       {},
        "ciclos":           [],
        "financeiro":       {},
        "itens":            {},
        "aditivos":         {},
        "modo_declarado":   "",
        "modo_detectado":   "Base insuficiente",
        "alertas":          [],
        "ajustes_gcc":      0,
    }

    try:
        xls = pd.ExcelFile(BytesIO(bytes_xlsx))
    except Exception as exc:
        resultado["erro"] = f"Não foi possível abrir o XLSX: {exc}"
        return resultado

    abas = set(xls.sheet_names)

    # Detecção automática: se for v11, redirecionar para leitor próprio
    if _eh_v11(abas):
        return _ler_v11(bytes_xlsx)

    resultado["abas_encontradas"] = sorted(abas)
    resultado["abas_ausentes"]    = [a for a in ABAS_ESPERADAS if a not in abas]

    if resultado["abas_ausentes"]:
        resultado["alertas"].append(
            "Abas ausentes: " + ", ".join(resultado["abas_ausentes"])
        )

    params     = _ler_parametros(xls, abas)
    ciclos     = _ler_ciclos(xls, abas)
    financeiro = _ler_financeiro(xls, abas)
    itens      = _ler_itens(xls, abas)
    aditivos   = _ler_aditivos(xls, abas)
    roteiro_info = _ler_roteiro_info_fiscais(xls, abas)

    modo_declarado = params.get("modo_declarado", "")
    modo_detectado = _detectar_modo(financeiro, itens, params)
    alertas        = _gerar_alertas(params, financeiro, itens, aditivos, modo_declarado)
    alertas        += resultado["alertas"]  # adicionar alertas de abas ausentes
    if roteiro_info.get("ok"):
        alertas.append("Roteiro das Informações dos Fiscais identificado no XLS.")

    ajustes = financeiro.get("ajustes_gcc",0) + itens.get("ajustes_gcc",0)

    resultado.update({
        "ok":             True,
        "parametros":     params,
        "ciclos":         ciclos,
        "financeiro":     financeiro,
        "itens":          itens,
        "aditivos":       aditivos,
        "modo_declarado": modo_declarado,
        "roteiro_info_fiscais": roteiro_info,
        "modo_detectado": modo_detectado,
        "alertas":        alertas,
        "ajustes_gcc":    ajustes,
    })
    return resultado
