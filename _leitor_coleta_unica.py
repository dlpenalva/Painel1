"""
_leitor_coleta_unica.py
-----------------------
Leitor isolado da Coleta Única Inteligente (cl8us v2.0).

Responsabilidade:
    Ler o XLSX da Coleta Única preenchida e retornar um dicionário
    estruturado com todos os dados. NÃO altera cálculo nem o módulo Valores.

Uso:
    from _leitor_coleta_unica import ler_coleta_unica

    with open("Coleta_Unica.xlsx", "rb") as f:
        resultado = ler_coleta_unica(f.read())

    if resultado["ok"]:
        print(resultado["parametros"])
        print(resultado["ciclos"])
        print(resultado["financeiro"])
        print(resultado["itens"])
        print(resultado["modo_preliminar"])

Retorno (dict):
    ok                  bool   — False se falhou na abertura
    erro                str    — mensagem de erro, se houver
    abas_encontradas    list   — abas reconhecidas no XLSX
    abas_ausentes       list   — abas esperadas mas não encontradas
    parametros          dict   — dados de PARAMETROS_CONTRATO
    ciclos              list   — lista de dicts, um por ciclo
    financeiro          dict   — total e linhas de FINANCEIRO_HISTORICO
    itens               dict   — totais e linhas de ITENS_CICLOS
    aditivos            list   — lista de dicts de ADITIVOS
    ciclo_em_execucao   dict   — dados de CICLO_EM_EXECUCAO
    validacoes          dict   — respostas de VALIDACOES_FISCAIS
    modo_preliminar     str    — classificação da base
    ressalvas           list   — lista de ressalvas técnicas
"""

import re
import unicodedata
from io import BytesIO

import pandas as pd


# ─────────────────────────────────────────────
# Abas esperadas na Coleta Única
# ─────────────────────────────────────────────
# Abas do formato legado (AgoraClau)
ABAS_LEGADO = [
    "INICIO", "PARAMETROS_CONTRATO", "CICLOS",
    "FINANCEIRO_HISTORICO", "ITENS_CICLOS", "ADITIVOS",
    "CICLO_EM_EXECUCAO", "VALIDACOES_FISCAIS", "DIAGNOSTICO_BASE",
]

# Abas do formato novo (Planilha Master Reajuste)
ABAS_MASTER = [
    "PARAMETROS_CONTRATO", "CICLOS", "FINANCEIRO",
    "ITENS", "ADITIVOS", "DIAGNOSTICO",
]

# Usamos o conjunto de abas obrigatórias mínimas para validação
ABAS_ESPERADAS = ABAS_LEGADO  # mantido por compatibilidade

def _detectar_formato(abas_existentes):
    """Retorna 'master' se for a Planilha Master Reajuste, senão 'legado'."""
    if "FINANCEIRO" in abas_existentes and "ITENS" in abas_existentes:
        return "master"
    return "legado"


# ─────────────────────────────────────────────
# Utilitários internos
# ─────────────────────────────────────────────

def _normalizar(valor):
    """Normaliza texto para busca de colunas: minúsculo, sem acento, sem espaço."""
    if valor is None:
        return ""
    texto = str(valor).strip().lower()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    return texto.strip("_")


def _numero(valor):
    """Converte célula para float, retorna 0.0 se inválido."""
    if valor is None:
        return 0.0
    try:
        if pd.isna(valor):
            return 0.0
    except Exception:
        pass
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("R$", "").replace("\xa0", "").replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except Exception:
        return 0.0


def _texto(valor, padrao=""):
    """Retorna texto limpo ou padrão se vazio/nan."""
    if valor is None:
        return padrao
    try:
        if pd.isna(valor):
            return padrao
    except Exception:
        pass
    t = str(valor).strip()
    return t if t.lower() not in ("nan", "none", "nat", "<na>", "") else padrao


def _ler_aba(xls, nome_aba, linha_header):
    """Lê aba e retorna DataFrame limpo. Retorna DataFrame vazio se falhar."""
    try:
        df = pd.read_excel(xls, sheet_name=nome_aba, header=linha_header)
    except Exception:
        return pd.DataFrame()
    df = df.dropna(how="all").copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~pd.Series(df.columns).astype(str).str.startswith("Unnamed").values]
    return df


def _coluna(df, opcoes):
    """Localiza coluna pelo nome normalizado. Retorna None se não encontrar."""
    mapa = {_normalizar(c): c for c in df.columns}
    for opcao in opcoes:
        chave = _normalizar(opcao)
        if chave in mapa:
            return mapa[chave]
    # Busca parcial como fallback
    for opcao in opcoes:
        chave = _normalizar(opcao)
        for col_norm, col_original in mapa.items():
            if chave and chave in col_norm:
                return col_original
    return None


# ─────────────────────────────────────────────
# Leitores por aba
# ─────────────────────────────────────────────

def _ler_parametros(xls, abas_existentes):
    """Lê PARAMETROS_CONTRATO → dict {campo_normalizado: valor}."""
    if "PARAMETROS_CONTRATO" not in abas_existentes:
        return {}
    # Formato master: campo em col A, valor em col C (sem header de col nomeada)
    _fmt = _detectar_formato(abas_existentes)
    if _fmt == "master":
        # Ler sem header e mapear A→campo, C→valor
        import pandas as _pd_pm
        df_raw = _pd_pm.read_excel(xls, sheet_name="PARAMETROS_CONTRATO", header=None)
        params_master = {}
        for _, row in df_raw.iterrows():
            campo = str(row.iloc[0]).strip() if not _pd_pm.isna(row.iloc[0]) else ""
            valor = str(row.iloc[2]).strip() if len(row) > 2 and not _pd_pm.isna(row.iloc[2]) else ""
            if campo and valor and campo not in ("nan","None") and valor not in ("nan","None","0") and not campo.startswith("cl8us") and not campo.startswith(" ") and "←" not in campo and "■" not in campo:
                params_master[_normalizar(campo)] = valor
        return params_master
    # Formato legado
    df = _ler_aba(xls, "PARAMETROS_CONTRATO", 2)
    col_c = _coluna(df, ["Campo"])
    if col_c is None:
        df = _ler_aba(xls, "PARAMETROS_CONTRATO", 1)
    col_campo = _coluna(df, ["Campo"])
    col_valor = _coluna(df, ["Valor"])
    if not col_campo or not col_valor:
        return {}
    params = {}
    for _, row in df.iterrows():
        campo = _normalizar(row.get(col_campo, ""))
        if campo:
            params[campo] = _texto(row.get(col_valor, ""))
    return params


def _ler_ciclos(xls, abas_existentes):
    """Lê CICLOS → lista de dicts, um por ciclo."""
    if "CICLOS" not in abas_existentes:
        return []
    _fmt_c = _detectar_formato(abas_existentes)
    if _fmt_c == "master":
        df = _ler_aba(xls, "CICLOS", 3)  # header na linha 4 (índice 3)
        # Filtrar linha de sub-headers [CALC]
        if not df.empty:
            col_ciclo_raw = _coluna(df, ["Ciclo"])
            if col_ciclo_raw:
                df = df[~df[col_ciclo_raw].astype(str).str.contains(r"\[CALC\]", na=False)]
    else:
        df = _ler_aba(xls, "CICLOS", 2)
        if df.empty or _coluna(df, ["Ciclo"]) is None:
            df = _ler_aba(xls, "CICLOS", 3)
    col_ciclo     = _coluna(df, ["Ciclo"])
    col_base      = _coluna(df, ["Data-base"])
    col_pedido    = _coluna(df, ["Data do pedido"])
    col_ini_fin   = _coluna(df, ["Início financeiro"])
    col_fim_fin   = _coluna(df, ["Fim financeiro"])
    col_situacao  = _coluna(df, ["Situação"])
    col_percent   = _coluna(df, ["Percentual aplicado"])
    col_fator     = _coluna(df, ["Fator do ciclo"])
    col_fat_acum  = _coluna(df, ["Fator acumulado"])
    col_objeto    = _coluna(df, ["Objeto da análise atual?", "Objeto da analise atual"])
    col_formal    = _coluna(df, ["Ciclo já formalizado?", "Ciclo ja formalizado"])
    col_obs       = _coluna(df, ["Observação"])

    ciclos = []
    for _, row in df.iterrows():
        ciclo = _texto(row.get(col_ciclo, "") if col_ciclo else "")
        if not ciclo:
            continue
        ciclos.append({
            "ciclo":              ciclo,
            "data_base":          _texto(row.get(col_base, "")     if col_base     else ""),
            "data_pedido":        _texto(row.get(col_pedido, "")   if col_pedido   else ""),
            "inicio_financeiro":  _texto(row.get(col_ini_fin, "")  if col_ini_fin  else ""),
            "fim_financeiro":     _texto(row.get(col_fim_fin, "")  if col_fim_fin  else ""),
            "situacao":           _texto(row.get(col_situacao, "")  if col_situacao else ""),
            "percentual_aplicado":_numero(row.get(col_percent, 0)  if col_percent  else 0),
            "fator_ciclo":        _numero(row.get(col_fator, 1)    if col_fator    else 1),
            "fator_acumulado":    _numero(row.get(col_fat_acum, 1) if col_fat_acum else 1),
            "objeto_analise":     _texto(row.get(col_objeto, "")   if col_objeto   else ""),
            "ja_formalizado":     _texto(row.get(col_formal, "")   if col_formal   else ""),
            "observacao":         _texto(row.get(col_obs, "")      if col_obs      else ""),
        })
    return ciclos


def _ler_financeiro(xls, abas_existentes):
    """Lê FINANCEIRO_HISTORICO → dict com linhas e total."""
    resultado = {"linhas": [], "total": 0.0, "linhas_preenchidas": 0}
    _aba_fin = "FINANCEIRO_HISTORICO" if "FINANCEIRO_HISTORICO" in abas_existentes else (
               "FINANCEIRO" if "FINANCEIRO" in abas_existentes else None)
    if not _aba_fin:
        return resultado
    df = _ler_aba(xls, _aba_fin, 3)
    # Filtrar linhas de separador/cabeçalho que não são dados reais
    _col1 = df.columns[0] if not df.empty else None
    if _col1:
        df = df[~df[_col1].astype(str).str.contains(
            r"\[AUTO\]|\[FISCAL\]|\[GCC\]|▼|Linhas mensais|Ciclo sugerido|nan", na=False
        )]
    col_ciclo = _coluna(df, ["Ciclo sugerido", "[AUTO]\nCiclo", "Ciclo", "AUTO\nCiclo"])
    col_comp  = _coluna(df, ["Competência", "[FISCAL]\nCompetência", "Competencia"])
    col_valor = _coluna(df, [
        "Valor informado pelo fiscal",
        "Valor Informado",
        "[FISCAL]\nValor Informado",
        "Valor líquido considerado",
        "Valor medido/aprovado",
        "Valor pago",
    ])
    if not col_valor:
        return resultado
    total = 0.0
    linhas = []
    for _, row in df.iterrows():
        valor = _numero(row.get(col_valor, 0))
        if abs(valor) < 0.004:
            continue
        _comp_txt = _texto(row.get(col_comp, "") if col_comp else "")
        # Ignorar linhas consolidadas (competência contém 'Executado consolidado' ou 'Executado')
        if "executado" in _comp_txt.lower() or "consolidado" in _comp_txt.lower():
            continue
        linhas.append({
            "ciclo":       _texto(row.get(col_ciclo, "") if col_ciclo else ""),
            "competencia": _comp_txt,
            "valor":       valor,
        })
        total += valor
    resultado["linhas"] = linhas
    resultado["total"] = round(total, 2)
    resultado["linhas_preenchidas"] = len(linhas)
    return resultado


def _ler_itens(xls, abas_existentes):
    """Lê ITENS_CICLOS → dict com linhas e totais."""
    resultado = {
        "linhas": [],
        "total_c0": 0.0,
        "itens_cadastrados": 0,
        "itens_com_c0": 0,
        "itens_com_rem_anterior": 0,
        "itens_com_rem_atual": 0,
    }
    _aba_it = "ITENS_CICLOS" if "ITENS_CICLOS" in abas_existentes else (
               "ITENS" if "ITENS" in abas_existentes else None)
    if not _aba_it:
        return resultado
    # Master tem header em linha 7 (índice 6)
    _fmt_it = _detectar_formato(abas_existentes)
    df = _ler_aba(xls, _aba_it, 6) if _fmt_it == "master" else _ler_aba(xls, _aba_it, 3)
    # Filtrar linhas de cabeçalho/grupo
    _col_it = df.columns[0] if not df.empty else None
    if _col_it:
        df = df[~df[_col_it].astype(str).str.contains(
            r"\[GCC\]|\[FISCAL\]|\[AUTO\]|Base Contratual|Remanescentes|preenche|■|nan", na=False
        )]
    col_item      = _coluna(df, ["Item"])
    col_desc      = _coluna(df, ["Descrição resumida", "Descricao resumida"])
    col_un        = _coluna(df, ["Unidade"])
    col_qtd_c0    = _coluna(df, ["Quantidade contratada C0"])
    col_vu_c0     = _coluna(df, ["Valor unitário original C0", "Valor unitario original C0"])
    col_vt_c0     = _coluna(df, ["Valor total original C0"])
    col_rem_c1    = _coluna(df, ["Remanescente C1"])
    col_rem_c2    = _coluna(df, ["Remanescente C2"])
    col_rem_c3    = _coluna(df, ["Remanescente C3"])
    col_rem_c4    = _coluna(df, ["Remanescente C4"])
    col_rem_atual = _coluna(df, ["Remanescente ciclo atual/corte"])
    col_consumo   = _coluna(df, ["Consumo informado por ciclo?"])

    linhas = []
    total_c0 = 0.0

    for _, row in df.iterrows():
        item = _texto(row.get(col_item, "") if col_item else "")
        if not item or item.upper() == "TOTAL":
            continue
        qtd_c0 = _numero(row.get(col_qtd_c0, 0) if col_qtd_c0 else 0)
        vu_c0  = _numero(row.get(col_vu_c0,  0) if col_vu_c0  else 0)
        vt_c0  = _numero(row.get(col_vt_c0,  0) if col_vt_c0  else 0)
        if vt_c0 == 0 and qtd_c0 > 0 and vu_c0 > 0:
            vt_c0 = round(qtd_c0 * vu_c0, 2)

        rem_c1    = _numero(row.get(col_rem_c1,    0) if col_rem_c1    else 0)
        rem_c2    = _numero(row.get(col_rem_c2,    0) if col_rem_c2    else 0)
        rem_c3    = _numero(row.get(col_rem_c3,    0) if col_rem_c3    else 0)
        rem_c4    = _numero(row.get(col_rem_c4,    0) if col_rem_c4    else 0)
        rem_atual = _numero(row.get(col_rem_atual, 0) if col_rem_atual else 0)

        tem_c0            = qtd_c0 > 0.004 and vu_c0 > 0.004
        tem_rem_anterior  = any(abs(v) > 0.004 for v in [rem_c1, rem_c2, rem_c3, rem_c4])
        tem_rem_atual     = abs(rem_atual) > 0.004

        linhas.append({
            "item":             item,
            "descricao":        _texto(row.get(col_desc, "")    if col_desc    else ""),
            "unidade":          _texto(row.get(col_un, "")      if col_un      else ""),
            "qtd_c0":           qtd_c0,
            "vu_c0":            vu_c0,
            "vt_c0":            vt_c0,
            "rem_c1":           rem_c1,
            "rem_c2":           rem_c2,
            "rem_c3":           rem_c3,
            "rem_c4":           rem_c4,
            "rem_atual":        rem_atual,
            "consumo_por_ciclo":_texto(row.get(col_consumo, "") if col_consumo else ""),
            "_tem_c0":          tem_c0,
            "_tem_rem_anterior":tem_rem_anterior,
            "_tem_rem_atual":   tem_rem_atual,
        })
        total_c0 += vt_c0

    resultado["linhas"]                = linhas
    resultado["total_c0"]              = round(total_c0, 2)
    resultado["itens_cadastrados"]     = len(linhas)
    resultado["itens_com_c0"]          = sum(1 for l in linhas if l["_tem_c0"])
    resultado["itens_com_rem_anterior"]= sum(1 for l in linhas if l["_tem_rem_anterior"])
    resultado["itens_com_rem_atual"]   = sum(1 for l in linhas if l["_tem_rem_atual"])
    return resultado


def _ler_aditivos(xls, abas_existentes):
    """Lê ADITIVOS → lista de dicts."""
    if "ADITIVOS" not in abas_existentes:
        return []
    df = _ler_aba(xls, "ADITIVOS", 2)
    col_id     = _coluna(df, ["Identificação", "Identificacao"])
    col_tipo   = _coluna(df, ["Tipo"])
    col_item   = _coluna(df, ["Item"])
    col_qtd    = _coluna(df, ["Quantidade"])
    col_vu     = _coluna(df, ["Valor unitário", "Valor unitario"])
    col_vt     = _coluna(df, ["Valor original"])
    col_formal = _coluna(df, ["Já formalizado?", "Ja formalizado"])
    col_incorp = _coluna(df, ["Incorporar no Valor Total?"])

    aditivos = []
    for _, row in df.iterrows():
        ident = _texto(row.get(col_id, "") if col_id else "")
        vt    = _numero(row.get(col_vt, 0) if col_vt else 0)
        if not ident and abs(vt) < 0.004:
            continue
        aditivos.append({
            "identificacao":  ident,
            "tipo":           _texto(row.get(col_tipo,   "") if col_tipo   else ""),
            "item":           _texto(row.get(col_item,   "") if col_item   else ""),
            "quantidade":     _numero(row.get(col_qtd,   0)  if col_qtd   else 0),
            "valor_unitario": _numero(row.get(col_vu,    0)  if col_vu    else 0),
            "valor_original": vt,
            "ja_formalizado": _texto(row.get(col_formal, "") if col_formal else ""),
            "incorporar":     _texto(row.get(col_incorp, "") if col_incorp else ""),
        })
    return aditivos


def _ler_ciclo_em_execucao(xls, abas_existentes):
    """Lê CICLO_EM_EXECUCAO → dict com os campos principais."""
    padrao = {
        "aplicar_corte": "Não",
        "ciclo_corte":   "",
        "competencia_corte": "",
        "fonte": "",
        "rem_original": 0.0,
        "rem_atualizado": 0.0,
        "observacao": "",
    }
    if "CICLO_EM_EXECUCAO" not in abas_existentes:
        return padrao
    df = _ler_aba(xls, "CICLO_EM_EXECUCAO", 3)
    col_campo = _coluna(df, ["Campo"])
    col_valor = _coluna(df, ["Valor"])
    if not col_campo or not col_valor:
        return padrao

    dados = padrao.copy()
    for _, row in df.iterrows():
        campo = _normalizar(row.get(col_campo, ""))
        valor = _texto(row.get(col_valor, ""))
        if "aplicar_corte" in campo:
            dados["aplicar_corte"] = valor or "Não"
        elif "ciclo_em_execucao" in campo or "ciclo_em_execu" in campo:
            dados["ciclo_corte"] = valor
        elif "competencia_de_corte" in campo or "competencia_corte" in campo:
            dados["competencia_corte"] = valor
        elif "fonte_da_execucao" in campo or "fonte" in campo:
            dados["fonte"] = valor
        elif "remanescente_original" in campo:
            dados["rem_original"] = _numero(row.get(col_valor, 0))
        elif "remanescente_atualizado" in campo:
            dados["rem_atualizado"] = _numero(row.get(col_valor, 0))
        elif "observacao" in campo or "premissa" in campo:
            dados["observacao"] = valor
    return dados


def _ler_validacoes(xls, abas_existentes):
    """Lê VALIDACOES_FISCAIS → dict {pergunta_normalizada: resposta}."""
    if "VALIDACOES_FISCAIS" not in abas_existentes:
        return {}
    df = _ler_aba(xls, "VALIDACOES_FISCAIS", 2)
    col_pergunta = _coluna(df, ["Pergunta"])
    col_resposta = _coluna(df, ["Resposta"])
    if not col_pergunta or not col_resposta:
        return {}
    validacoes = {}
    for _, row in df.iterrows():
        pergunta = _normalizar(row.get(col_pergunta, ""))
        resposta = _texto(row.get(col_resposta, ""))
        if pergunta:
            validacoes[pergunta] = resposta
    return validacoes


def _classificar_modo(financeiro, itens, validacoes):
    """Classifica o modo preliminar de apuração a partir dos dados lidos."""
    tem_financeiro     = financeiro["linhas_preenchidas"] > 0 and abs(financeiro["total"]) > 0.004
    tem_itens          = itens["itens_cadastrados"] > 0
    tem_rem_atual      = itens["itens_com_rem_atual"] > 0
    tem_rem_anterior   = itens["itens_com_rem_anterior"] > 0

    if tem_financeiro and tem_itens and tem_rem_atual and tem_rem_anterior:
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

    ressalvas = []
    if tem_financeiro and not tem_rem_anterior:
        ressalvas.append("Sem memória itemizada dos ciclos anteriores; usar financeiro histórico para execução anterior, se tecnicamente suficiente.")
    if not tem_financeiro:
        ressalvas.append("Sem financeiro histórico preenchido; não calcular retroativo financeiro definitivo.")
    if tem_itens and not tem_rem_atual:
        ressalvas.append("Itens cadastrados sem remanescente atual/corte; o saldo remanescente pode ficar limitado.")

    fin_completo = next(
        (resp for perg, resp in validacoes.items() if "historico_financeiro" in perg and "completo" in perg),
        "",
    )
    if fin_completo and fin_completo.lower() not in ("sim", ""):
        ressalvas.append(f"Fiscal declarou financeiro histórico como: {fin_completo}.")

    return modo, ressalvas


# ─────────────────────────────────────────────
# Função pública principal
# ─────────────────────────────────────────────

def ler_coleta_unica(bytes_xlsx):
    """
    Lê o XLSX da Coleta Única e retorna dicionário estruturado.

    Parâmetro:
        bytes_xlsx: bytes — conteúdo binário do XLSX.

    Retorno: dict conforme docstring do módulo.
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
        "aditivos":         [],
        "ciclo_em_execucao":{},
        "validacoes":       {},
        "modo_preliminar":  "Base insuficiente",
        "ressalvas":        [],
    }

    try:
        xls = pd.ExcelFile(BytesIO(bytes_xlsx))
    except Exception as exc:
        resultado["erro"] = f"Não foi possível abrir o XLSX: {exc}"
        return resultado

    abas_existentes = set(xls.sheet_names)
    resultado["abas_encontradas"] = list(abas_existentes)
    _fmt = _detectar_formato(abas_existentes)
    _abas_ref = ABAS_MASTER if _fmt == "master" else ABAS_LEGADO
    resultado["abas_ausentes"] = [a for a in _abas_ref if a not in abas_existentes]
    resultado["formato_detectado"] = _fmt

    if resultado["abas_ausentes"]:
        resultado["ressalvas"].append("Abas ausentes: " + ", ".join(resultado["abas_ausentes"]))

    resultado["parametros"]        = _ler_parametros(xls, abas_existentes)
    resultado["ciclos"]            = _ler_ciclos(xls, abas_existentes)
    resultado["financeiro"]        = _ler_financeiro(xls, abas_existentes)
    resultado["itens"]             = _ler_itens(xls, abas_existentes)
    resultado["aditivos"]          = _ler_aditivos(xls, abas_existentes)
    resultado["ciclo_em_execucao"] = _ler_ciclo_em_execucao(xls, abas_existentes)
    resultado["validacoes"]        = _ler_validacoes(xls, abas_existentes)

    modo, ressalvas_modo = _classificar_modo(
        resultado["financeiro"],
        resultado["itens"],
        resultado["validacoes"],
    )
    resultado["modo_preliminar"] = modo
    resultado["ressalvas"].extend(ressalvas_modo)
    resultado["ok"] = True
    return resultado
