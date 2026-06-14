"""
_ponte_coleta_valores.py
------------------------
Ponte entre a Coleta Mestre v10 e o motor de cálculo (03_Valor_Global.py).

Estratégia:
    1. Detecta automaticamente se o arquivo é ColetaMestre v10 (abas novas)
       ou arquivo legado (estrutura antiga).
    2. Para v10: usa _leitor_coleta_mestre.py para extrair os dados e os
       converte para os DataFrames e dicts que o motor já conhece.
    3. Para legado: delega diretamente a processar_arquivo_coleta() do motor.

Ponto de entrada único:
    from _ponte_coleta_valores import processar_coleta
    resultado = processar_coleta(bytes_arquivo, modo_confirmado=None)

Retorno:
    dict — mesmo formato de resultado que processar_arquivo_coleta() retorna,
    com campos adicionais:
        origem_coleta   "v10" | "legado"
        modo_apuracao   str
        alertas_leitor  list
"""

from io import BytesIO
import pandas as pd
import numpy as np


# ─────────────────────────────────────────────────────────────────
# Detecção de versão
# ─────────────────────────────────────────────────────────────────

_ABAS_V10 = {"PARAMETROS_CONTRATO","CICLOS","FINANCEIRO","ITENS","ADITIVOS","DIAGNOSTICO"}
_ABAS_LEGADO = {"BASE_EXECUCAO_MENSAL","FINANCEIRO_HISTORICO","FINANCEIRO_MENSAL"}


def _detectar_versao(xls) -> str:
    """Retorna 'v10' se reconhece a ColetaMestre v10, senão 'legado'."""
    abas = set(xls.sheet_names)
    if "FINANCEIRO" in abas and "ITENS" in abas and "PARAMETROS_CONTRATO" in abas:
        # Confirmar que não é legado com aba FINANCEIRO_HISTORICO
        if not abas.intersection(_ABAS_LEGADO):
            return "v10"
    return "legado"


# ─────────────────────────────────────────────────────────────────
# Conversores v10 → formato motor
# ─────────────────────────────────────────────────────────────────

def _converter_ciclos_v10(ciclos_v10: list) -> pd.DataFrame:
    """
    Converte a lista de ciclos do leitor v10 para o DataFrame que
    padronizar_ciclos() do motor retornaria.
    Inclui C0 (sem reajuste) e ciclos de análise.
    """
    linhas = []
    fat_acum = 1.0

    for ciclo in ciclos_v10:
        nome = ciclo["ciclo"]
        pct  = float(ciclo.get("percentual", 0) or 0)
        fat  = float(ciclo.get("fator_ciclo", 1) or 1)
        fat_ac = float(ciclo.get("fator_acumulado", 1) or 1)
        sit  = ciclo.get("situacao", "")
        trat = ciclo.get("objeto_analise", "")
        tem_ef = ciclo.get("tem_efeito_fin", "Sim")
        entra  = ciclo.get("entra_valor_total", "Sim")

        # Derivar fator acumulado efetivo
        precluso = (tem_ef.lower() == "não") or ("preclu" in sit.lower())
        if precluso:
            fat_ef     = 1.0
            fat_acum_ef = fat_acum
        else:
            fat_acum   = fat_acum * fat
            fat_ef     = fat
            fat_acum_ef = fat_acum

        linhas.append({
            "Ciclo":                    nome,
            "Data-base":                ciclo.get("data_base",""),
            "Intervalo do índice":      "",
            "Janela de admissibilidade":"",
            "Data do pedido":           "",
            "Início financeiro":        ciclo.get("inicio_financeiro",""),
            "Fim financeiro":           ciclo.get("fim_financeiro",""),
            "Situação":                 sit,
            "Situação automática":      sit,
            "Acordo negocial":          "Não",
            "Situação aplicada":        sit,
            "Percentual apurado pelo índice": pct,
            "Percentual aplicado":      pct,
            "Ciclo negativo":           "Não",
            "Tratamento ciclo negativo":"",
            "Justificativa negocial":   "",
            "Referência documental":    "",
            "Tratamento financeiro do ciclo": "A apurar" if not precluso else "Sem efeito financeiro",
            "Variação":                 pct,
            "Fator":                    fat,
            "Fator acumulado":          fat_ac,
            "Fator acumulado efetivo":  fat_acum_ef,
            "Fator ciclo efetivo":      fat_ef,
            "_tem_efeito_financeiro":   not precluso,
            "_entra_valor_total":       entra.lower() == "sim",
        })

    df = pd.DataFrame(linhas)
    return df


def _converter_financeiro_v10(financeiro_v10: dict, ciclos_df: pd.DataFrame) -> pd.DataFrame:
    """
    Converte o financeiro do leitor v10 para o DataFrame que ler_financeiro() retornaria.
    Inclui linhas consolidadas C0/C1 e linhas mensais.
    Nome da coluna principal: 'Valor pago/faturado' (padrão do motor).
    """
    linhas = []

    # Linhas consolidadas (C0 e C1)
    for cons in financeiro_v10.get("consolidados", []):
        if cons["valor"] > 0:
            linhas.append({
                "Ciclo":              cons["ciclo"],
                "Competência":        cons["ciclo"] + " (consolidado)",
                "Valor pago/faturado": cons["valor"],
                "Tem efeito financeiro de reajuste?": cons["efeito"],
                "_consolidado":       True,
            })

    # Linhas mensais
    for m in financeiro_v10.get("mensais", []):
        linhas.append({
            "Ciclo":              m["ciclo"],
            "Competência":        m["competencia"],
            "Valor pago/faturado": m["valor"],
            "Tem efeito financeiro de reajuste?": m["efeito"],
            "_consolidado":       False,
        })

    if not linhas:
        return pd.DataFrame(columns=["Ciclo","Competência","Valor pago/faturado"])

    df = pd.DataFrame(linhas)

    # Normalizar Ciclo
    df["Ciclo"] = df["Ciclo"].fillna("").str.upper().str.strip()

    # Inferir ciclo pelas datas se estiver vazio (linhas mensais sem ciclo pré-preenchido)
    if not ciclos_df.empty and "Ciclo" in ciclos_df.columns:
        mapa_ciclo = {}
        for _, row in ciclos_df.iterrows():
            mapa_ciclo[row["Ciclo"]] = {
                "ini": str(row.get("Início financeiro","") or ""),
                "fim": str(row.get("Fim financeiro","")    or ""),
            }

    df["Valor pago/faturado"] = pd.to_numeric(df["Valor pago/faturado"], errors="coerce").fillna(0.0)
    df = df[df["Valor pago/faturado"].abs() > 0.003].copy()

    return df.reset_index(drop=True)


def _converter_itens_v10(itens_v10: dict) -> tuple:
    """
    Converte itens do leitor v10 para (df_itens, colunas_remanescentes).
    O motor espera colunas: Item, Qtd C0, VU C0, VT C0, Rem C1..Cn, Rem Atual.
    """
    itens_lista = itens_v10.get("itens", [])
    if not itens_lista:
        return pd.DataFrame(), []

    linhas = []
    for item in itens_lista:
        linhas.append({
            "Item":                    item["item"],
            "Quantidade contratada C0": item.get("qtd_c0", 0),
            "Valor unitário original C0": item.get("vu_c0", 0),
            "Valor total original C0":   item.get("vt_c0", 0),
            "Remanescente C1":           item.get("rem_c1", 0),
            "Remanescente C2":           item.get("rem_c2", 0),
            "Remanescente C3":           item.get("rem_c3", 0),
            "Remanescente C4":           item.get("rem_c4", 0),
            "Remanescente C5":           item.get("rem_c5", 0),
            "Remanescente ciclo atual/corte": item.get("rem_atual", 0),
        })

    df = pd.DataFrame(linhas)
    colunas_rem = [c for c in df.columns if "Remanescente" in c and df[c].sum() > 0]

    return df, colunas_rem


def _converter_params_v10(params_v10: dict) -> dict:
    """Converte params do leitor v10 para o dict que o motor usa."""
    return {
        "indice":               params_v10.get("indice_contratual",""),
        "data_base_original":   params_v10.get("data_base_original",""),
        "fator_acumulado":      params_v10.get("fator_acumulado_efetivo", 1.0),
        "fator_acumulado_total":params_v10.get("fator_acumulado_efetivo", 1.0),
        "fator_acumulado_final":params_v10.get("fator_acumulado_efetivo", 1.0),
        "valor_original_contrato": params_v10.get("valor_original_contrato", 0),
        "variacao_acumulada":   params_v10.get("variacao_acumulada", 0),
        "ciclo_inicial":        params_v10.get("ciclo_inicial_analise",""),
        "ciclo_atual":          params_v10.get("ciclo_atual_corte",""),
        "competencia_corte":    params_v10.get("competencia_corte",""),
        "ha_aditivos":          params_v10.get("ha_aditivos","Não"),
        "ha_corte":             params_v10.get("ha_corte_operacional","Não"),
        "c0_executado_manual":  params_v10.get("c0_executado_manual", 0),
        "rem_original_corte":   params_v10.get("saldo_rem_original_corte", 0),
        "rem_atualizado_corte": params_v10.get("saldo_rem_atualizado_corte", 0),
        "saldo_inclui_aditivos":params_v10.get("saldo_inclui_aditivos","Não"),
        "modo_declarado":       params_v10.get("modo_declarado",""),
    }


def _converter_aditivos_v10(aditivos_v10: dict) -> pd.DataFrame:
    """Converte aditivos do leitor v10 para DataFrame do motor."""
    linhas = aditivos_v10.get("linhas", [])
    if not linhas:
        return pd.DataFrame()

    rows = []
    for a in linhas:
        rows.append({
            "Identificação":          a.get("item",""),
            "Data de assinatura":     a.get("data",""),
            "Ciclo/Marco financeiro": a.get("ciclo",""),
            "Tipo":                   a.get("tipo","Acréscimo"),
            "Item":                   a.get("item",""),
            "Quantidade":             a.get("quantidade",0),
            "Valor unitário":         a.get("vu_original",0),
            "Valor original":         a.get("vt_original",0),
            "Fator aplicável":        a.get("fator",1.0),
            "Valor atualizado da alteração": a.get("vt_atualizado",0),
            "Tratamento do aditivo":  a.get("tratamento","Computar nesta análise"),
            "Incorporar no Valor Total?": "Sim" if a.get("_computar") else "Não",
        })

    return pd.DataFrame(rows)


def _determinar_modo(financeiro_v10, itens_v10, params_v10, modo_confirmado=None):
    """Determina o modo de apuração a partir dos dados v10."""
    if modo_confirmado:
        return modo_confirmado

    tem_fin     = financeiro_v10.get("linhas_mensais_preenchidas",0) > 0
    tem_c0_c1   = any(c["valor"]>0 for c in financeiro_v10.get("consolidados",[]))
    tem_rem     = itens_v10.get("itens_com_rem_c1",0) > 0
    tem_rem_atu = itens_v10.get("itens_com_rem_atual",0) > 0
    tem_itens   = itens_v10.get("itens_cadastrados",0) > 0

    if tem_fin and tem_rem and tem_rem_atu:
        return "Completo"
    elif tem_fin and tem_rem_atu:
        return "Financeiro/Base de execução mensal"
    elif tem_fin:
        return "Financeiro/Base de execução mensal"
    elif tem_rem_atu or tem_rem:
        return "Reduzido por Itens/Estoque"
    else:
        return "Base insuficiente"


# ─────────────────────────────────────────────────────────────────
# Chamada ao motor com dados convertidos
# ─────────────────────────────────────────────────────────────────

def _calcular_via_motor_v10(dados_v10: dict, modo_confirmado=None) -> dict:
    """
    Ponto de entrada interno para cálculo v10.
    Usa calcular_resultado_v10() que produz todos os DataFrames.
    Tenta também injetar no motor legado se disponível no contexto.
    """
    # Caminho principal: cálculo completo v10
    resultado = calcular_resultado_v10(dados_v10, modo_confirmado)

    # Enriquecimento opcional: tentar chamar funções do motor legado
    # para auditoria de consistência (não crítico — se falhar, ignora)
    try:
        import sys
        for mod_name, mod in sys.modules.items():
            if ("03_Valor_Global" in mod_name or "Valor_Global" in mod_name):
                if hasattr(mod, "montar_auditoria_consistencia"):
                    resultado["df_auditoria_consistencia"] = (
                        mod.montar_auditoria_consistencia(resultado)
                    )
                break
    except Exception:
        pass

    return resultado


# ─────────────────────────────────────────────────────────────────
# Ponto de entrada público
# ─────────────────────────────────────────────────────────────────

MODOS_COLETA = [
    "Completo",
    "Financeiro/Base de execução mensal",
    "Reduzido por Itens/Estoque",
    "Base insuficiente",
]


def processar_coleta(bytes_arquivo: bytes, modo_confirmado: str = None) -> dict:
    """
    Ponto de entrada único para processar qualquer Coleta.

    Detecta automaticamente:
    - ColetaMestre v10 → usa leitor v10 + ponte
    - Arquivo legado   → delega ao motor diretamente

    Parâmetros:
        bytes_arquivo:  bytes do XLSX
        modo_confirmado: str opcional — modo escolhido pelo GCC na UI

    Retorna:
        dict com os campos padrão de resultado do motor
    """
    try:
        xls = pd.ExcelFile(BytesIO(bytes_arquivo))
    except Exception as exc:
        return {"ok": False, "erro": f"Não foi possível abrir o XLSX: {exc}"}

    versao = _detectar_versao(xls)

    if versao == "v10":
        # Usar leitor v10
        try:
            from _leitor_coleta_mestre import ler_coleta_mestre
        except ImportError:
            return {"ok": False, "erro": "_leitor_coleta_mestre.py não encontrado na raiz."}

        dados_v10 = ler_coleta_mestre(bytes_arquivo)
        if not dados_v10.get("ok"):
            return {"ok": False, "erro": dados_v10.get("erro", "Erro na leitura do v10.")}

        resultado = _calcular_via_motor_v10(dados_v10, modo_confirmado)
        resultado["ok"] = True
        return resultado

    else:
        # Legado: delegar ao motor
        try:
            # Importar de dentro do contexto Streamlit (pages/)
            import sys
            for mod_name, mod in sys.modules.items():
                if ("03_Valor_Global" in mod_name or "Valor_Global" in mod_name):
                    if hasattr(mod, "processar_arquivo_coleta"):
                        resultado = mod.processar_arquivo_coleta(bytes_arquivo)
                        resultado["origem_coleta"] = "legado"
                        resultado["ok"] = True
                        if modo_confirmado:
                            resultado["modo_apuracao"] = modo_confirmado
                        return resultado

            # Se não encontrou o motor no sys.modules, tentar importar direto
            # (funciona em testes fora do Streamlit)
            from pages._03_Valor_Global import processar_arquivo_coleta
            resultado = processar_arquivo_coleta(bytes_arquivo)
            resultado["origem_coleta"] = "legado"
            resultado["ok"] = True
            return resultado

        except Exception as exc:
            return {
                "ok":            False,
                "erro":          f"Erro no processamento legado: {exc}",
                "origem_coleta": "legado",
            }


# ─────────────────────────────────────────────────────────────────
# Cálculo completo v10 — produz todos os DataFrames
# ─────────────────────────────────────────────────────────────────

def calcular_resultado_v10(dados_v10: dict, modo_confirmado: str = None) -> dict:
    """
    Calcula o resultado completo a partir dos dados do leitor v10.
    Produz todos os DataFrames que os documentos (Apostila, PDF) precisam.

    Lógica de cálculo:
        Valor Total = C0 + Σ(exec_ciclo × fator_acum_efetivo) + rem_atualizado + aditivos_computáveis

    Retroativo = Σ(exec_ciclo_com_efeito × fator_acum_efetivo) − Σ(exec_ciclo_com_efeito)
    """
    params     = dados_v10.get("parametros", {})
    ciclos_v10 = dados_v10.get("ciclos", [])
    fin_v10    = dados_v10.get("financeiro", {})
    itens_v10  = dados_v10.get("itens", {})
    adits_v10  = dados_v10.get("aditivos", {})

    # ── Converter ciclos → mapa nome→dict ──────────────────────────
    mapa_ciclo = {c["ciclo"]: c for c in ciclos_v10}

    def fator_efetivo(ciclo_nome: str) -> float:
        """Fator acumulado efetivo do ciclo — 1.0 se precluso."""
        c = mapa_ciclo.get(ciclo_nome, {})
        tem_ef = c.get("tem_efeito_fin", "Sim")
        if str(tem_ef).lower() in ("não","nao","false","0"):
            return 1.0
        return float(c.get("fator_acumulado", 1.0) or 1.0)

    def fator_ciclo_proprio(ciclo_nome: str) -> float:
        c = mapa_ciclo.get(ciclo_nome, {})
        tem_ef = c.get("tem_efeito_fin", "Sim")
        if str(tem_ef).lower() in ("não","nao","false","0"):
            return 1.0
        return float(c.get("fator_ciclo", 1.0) or 1.0)

    # ── C0 executado ───────────────────────────────────────────────
    c0_manual = float(params.get("c0_executado_manual", 0) or 0)
    c0_consol = next(
        (c["valor"] for c in fin_v10.get("consolidados", []) if c["ciclo"] == "C0"),
        0.0
    )
    c0_exec = c0_manual if c0_manual > 0 else c0_consol

    # ── Remanescente ───────────────────────────────────────────────
    rem_orig = float(params.get("saldo_rem_original_corte", 0) or 0)
    rem_atu  = float(params.get("saldo_rem_atualizado_corte", 0) or 0)
    fator_acum_global = float(params.get("fator_acumulado_efetivo", 1.0) or 1.0)
    if rem_atu == 0 and rem_orig > 0:
        rem_atu = round(rem_orig * fator_acum_global, 2)
    if rem_atu == 0:
        rem_atu = float(itens_v10.get("total_rem_atual_atualizado", 0) or 0)

    # ── Aditivos computáveis ───────────────────────────────────────
    adits_comp = float(adits_v10.get("total_computavel", 0) or 0)

    # ── Consolidar financeiro por ciclo ───────────────────────────
    # Soma valores por ciclo (consolidados + mensais)
    exec_por_ciclo: dict[str, float] = {}

    for cons in fin_v10.get("consolidados", []):
        ciclo = cons["ciclo"]
        val   = float(cons.get("valor", 0) or 0)
        if abs(val) > 0.001:
            exec_por_ciclo[ciclo] = exec_por_ciclo.get(ciclo, 0.0) + val

    for m in fin_v10.get("mensais", []):
        ciclo = m.get("ciclo", "")
        val   = float(m.get("valor", 0) or 0)
        if ciclo and abs(val) > 0.001:
            exec_por_ciclo[ciclo] = exec_por_ciclo.get(ciclo, 0.0) + val

    # ── Cálculo por ciclo ──────────────────────────────────────────
    linhas_por_ciclo = []
    soma_pago_com_ef  = 0.0
    soma_teorico      = 0.0
    soma_exec_atualiz = 0.0   # execução × fator acumulado efetivo

    for ciclo_nome, exec_orig in sorted(exec_por_ciclo.items()):
        fat_acum = fator_efetivo(ciclo_nome)       # fator acumulado efetivo do ciclo
        fat_pr   = fator_ciclo_proprio(ciclo_nome) # fator próprio (para label)
        c_info   = mapa_ciclo.get(ciclo_nome, {})
        tem_ef   = str(c_info.get("tem_efeito_fin","Sim")).lower() not in ("não","nao","false","0")

        # Usar fator ACUMULADO para ambos: retroativo e Valor Total
        exec_atu = round(exec_orig * fat_acum, 2)
        teorico  = exec_atu if tem_ef else exec_orig
        delta    = round(exec_atu - exec_orig, 2) if tem_ef else 0.0

        if tem_ef:
            soma_pago_com_ef  += exec_orig
            soma_teorico      += teorico

        soma_exec_atualiz += exec_atu

        linhas_por_ciclo.append({
            "Ciclo":                   ciclo_nome,
            "Valor pago efetivo":      exec_orig,
            "Fator próprio do ciclo":  fat_pr,
            "Fator acumulado efetivo": fat_acum,
            "Valor teórico calculado": teorico,
            "Delta do ciclo":          delta,
            "Tem efeito financeiro":   "Sim" if tem_ef else "Não",
            "Situação":                c_info.get("situacao",""),
        })

    # Adicionar C0 se não estava no financeiro
    if "C0" not in exec_por_ciclo and c0_exec > 0:
        linhas_por_ciclo.insert(0, {
            "Ciclo":                   "C0",
            "Valor pago efetivo":      c0_exec,
            "Fator próprio do ciclo":  1.0,
            "Fator acumulado efetivo": 1.0,
            "Valor teórico calculado": c0_exec,
            "Delta do ciclo":          0.0,
            "Tem efeito financeiro":   "Não",
            "Situação":                "Base sem reajuste",
        })
        soma_exec_atualiz += c0_exec

    df_fin_ciclo = pd.DataFrame(linhas_por_ciclo) if linhas_por_ciclo else pd.DataFrame()

    # ── Retroativo ─────────────────────────────────────────────────
    retroativo = round(soma_teorico - soma_pago_com_ef, 2)

    # ── Valor Total Atualizado ─────────────────────────────────────
    # soma_exec_atualiz já inclui C0 (adicionado no bloco acima)
    valor_global = round(
        soma_exec_atualiz
        + rem_atu
        + adits_comp,
        2
    )

    # ── df_composicao_valor_total ──────────────────────────────────
    linhas_comp = []
    if c0_exec > 0:
        linhas_comp.append({"Parcela": "C0 executado (sem reajuste)", "Valor": c0_exec})

    for ln in linhas_por_ciclo:
        if ln["Ciclo"] == "C0":
            continue
        fat_label = ln.get("Fator próprio do ciclo", ln.get("Fator efetivo", 1.0))
        label = f"{ln['Ciclo']} executado"
        if ln["Tem efeito financeiro"] == "Sim":
            label += f" atualizado (× {fat_label:.4f})"
        else:
            label += " sem reajuste (precluso)"
        linhas_comp.append({"Parcela": label, "Valor": ln["Valor teórico calculado"]})

    if rem_atu > 0:
        linhas_comp.append({"Parcela": "Saldo remanescente atualizado", "Valor": rem_atu})
    if adits_comp > 0:
        linhas_comp.append({"Parcela": "Aditivos/supressões computáveis atualizados", "Valor": adits_comp})
    linhas_comp.append({"Parcela": "VALOR TOTAL ATUALIZADO DO CONTRATO", "Valor": valor_global})

    df_composicao = pd.DataFrame(linhas_comp)

    # ── df_execucao_atualizada ─────────────────────────────────────
    cols_exec = ["Ciclo","Valor pago efetivo","Fator acumulado efetivo",
                 "Valor teórico calculado","Delta do ciclo"]
    cols_exec_disp = [c for c in cols_exec if c in df_fin_ciclo.columns]
    df_exec_atu = df_fin_ciclo[cols_exec_disp].copy() if not df_fin_ciclo.empty else pd.DataFrame()

    # ── df_aditivos_executivo ──────────────────────────────────────
    df_ad = _converter_aditivos_v10(adits_v10)

    # ── df_ciclos ──────────────────────────────────────────────────
    df_ciclos = _converter_ciclos_v10(ciclos_v10)

    # ── Montar resultado final ─────────────────────────────────────
    modo = modo_confirmado or dados_v10.get("modo_detectado", "Completo")
    fator_fmt = fator_acum_global

    resultado = {
        # Identificação
        "ok":                         True,
        "origem_coleta":              "v10",
        "modo_apuracao":              modo,
        "alertas_leitor":             dados_v10.get("alertas", []),
        # Dados calculadora
        "indice":                     params.get("indice_contratual", ""),
        "fator_acumulado":            fator_fmt,
        "variacao_acumulada":         float(params.get("variacao_acumulada", fator_fmt - 1) or (fator_fmt - 1)),
        # Valores principais
        "valor_pago_efetivo":         round(soma_pago_com_ef, 2),
        "total_pago_faturado":        round(soma_pago_com_ef, 2),
        "valor_teorico_calculado":    round(soma_teorico, 2),
        "valor_represado_a_pagar":    retroativo,
        "delta_acumulado":            retroativo,
        "remanescente_reajustado":    rem_atu,
        "total_aditivos_atualizados": adits_comp,
        "valor_atualizado_contrato":  valor_global,
        "valor_global_estoque":       valor_global,
        # DataFrames completos
        "df_ciclos":                  df_ciclos,
        "df_financeiro_mensal":       _converter_financeiro_v10(fin_v10, df_ciclos),
        "df_financeiro_por_ciclo":    df_fin_ciclo,
        "df_execucao_atualizada":     df_exec_atu,
        "df_composicao_valor_total":  df_composicao,
        "df_aditivos_executivo":      df_ad,
        "df_aditivos":                df_ad,
        "df_comparativo":             df_fin_ciclo,
        "df_itens":                   _converter_itens_v10(itens_v10)[0],
        # Params
        "params":                     _converter_params_v10(params),
        "params_v10":                 params,
    }

    return resultado
