"""View-model / adaptadores de UX da Adequacao Orcamentaria (UX V3).

Camada SEM Streamlit: formatacao, normalizacao e preparacao de dados para a
"planilha guiada" (abas Base / Historico / Projecao / Resultado). TODA a
matematica continua vivendo no motor de dominio (_adequacao_orcamentaria); este
modulo apenas estrutura entradas e apresenta saidas, delegando os calculos
(_round2, media_financeiro, media_pedidos_compra, janela_financeira_competencias,
valor_original_foi_informado, ...). Nao ha formula duplicada aqui.

Extraido de pages/12_Adequacao_Orcamentaria.py na Etapa UX V3 para reduzir a
complexidade da pagina (Secao 39). Os helpers sao os MESMOS ja homologados: as
assinaturas e o comportamento numerico permanecem identicos, preservando o
golden Financeiro/PC, projecao, programacao e XLSX.
"""
from __future__ import annotations

from datetime import datetime, date
from io import BytesIO
import re

import pandas as pd

from _seguranca_xlsx import opcoes_excel_writer_seguro

# A matematica vive no motor. A UI (esta camada) apenas delega.
from _adequacao_orcamentaria import (
    _round2,
    media_financeiro,
    valor_original_foi_informado,
    janela_financeira_competencias,
    CicloCadencia,
)


# ---------------------------------------------------------------- formatacao

def moeda(valor, com_prefixo=True):
    try:
        valor = round(float(valor or 0), 2)
    except Exception:
        valor = 0.0
    texto = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}" if com_prefixo else texto


def parse_moeda_br(valor):
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        try:
            if pd.isna(valor):
                return 0.0
        except Exception:
            pass
        return float(valor)
    texto = str(valor).strip()
    if not texto:
        return 0.0
    texto = texto.replace("R$", "").replace(" ", "").replace("\xa0", "").replace("%", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except Exception:
        return 0.0


def pct(valor):
    try:
        n = float(valor or 0)
    except Exception:
        n = 0.0
    if abs(n) < 1:
        n *= 100
    return f"{n:.2f}%".replace(".", ",")


def texto_seguro(valor, padrao="[campo a preencher]"):
    if valor is None:
        return padrao
    try:
        if pd.isna(valor):
            return padrao
    except Exception:
        pass
    texto = str(valor).strip()
    if not texto or texto.lower() in ["nan", "none", "null", "nat", "<na>"]:
        return padrao
    return texto


def normalizar_texto(valor):
    if valor is None:
        return ""
    texto = str(valor).strip().lower()
    mapa = str.maketrans("áàâãäéèêëíìîïóòôõöúùûüç", "aaaaaeeeeiiiiooooouuuuc")
    texto = texto.translate(mapa)
    texto = re.sub(r"[^a-z0-9]+", "_", texto)
    return texto.strip("_")


def localizar_coluna(df, opcoes):
    if not isinstance(df, pd.DataFrame) or df.empty:
        return None
    mapa = {normalizar_texto(c): c for c in df.columns}
    for opcao in opcoes:
        alvo = normalizar_texto(opcao)
        if alvo in mapa:
            return mapa[alvo]
    for col_norm, col_original in mapa.items():
        for opcao in opcoes:
            alvo = normalizar_texto(opcao)
            if alvo and alvo in col_norm:
                return col_original
    return None


MESES_PT = {
    "jan": 1, "janeiro": 1, "fev": 2, "fevereiro": 2, "mar": 3, "marco": 3, "março": 3,
    "abr": 4, "abril": 4, "mai": 5, "maio": 5, "jun": 6, "junho": 6,
    "jul": 7, "julho": 7, "ago": 8, "agosto": 8, "set": 9, "setembro": 9,
    "out": 10, "outubro": 10, "nov": 11, "novembro": 11, "dez": 12, "dezembro": 12,
}


def periodo_para_label(periodo):
    if periodo is None or pd.isna(periodo):
        return ""
    p = pd.Period(periodo, freq="M")
    nomes = ["jan", "fev", "mar", "abr", "mai", "jun", "jul", "ago", "set", "out", "nov", "dez"]
    return f"{nomes[p.month - 1]}/{str(p.year)[-2:]}"


def normalizar_competencia(valor):
    if valor is None:
        return None
    try:
        if pd.isna(valor):
            return None
    except Exception:
        pass
    if isinstance(valor, pd.Period):
        return valor.asfreq("M")
    if isinstance(valor, (datetime, date, pd.Timestamp)):
        data = pd.to_datetime(valor, errors="coerce")
        if pd.notna(data):
            return data.to_period("M")
    texto = str(valor).strip().lower()
    if not texto or texto in ["nan", "none", "nat", "total"]:
        return None
    texto = texto.replace(".", "/").replace("-", "/")
    m = re.search(r"([a-zçãéíóú]+)\s*/\s*(\d{2,4})", texto, flags=re.IGNORECASE)
    if m:
        mes_txt = normalizar_texto(m.group(1))
        ano = int(m.group(2))
        if ano < 100:
            ano += 2000
        mes = MESES_PT.get(mes_txt)
        if mes:
            return pd.Period(f"{ano}-{mes:02d}", freq="M")
    m = re.search(r"(\d{1,2})\s*/\s*(\d{2,4})", texto)
    if m:
        mes = int(m.group(1))
        ano = int(m.group(2))
        if 1 <= mes <= 12:
            if ano < 100:
                ano += 2000
            return pd.Period(f"{ano}-{mes:02d}", freq="M")
    data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
    if pd.notna(data):
        return data.to_period("M")
    return None


def periodo_de_data_final(valor):
    data = pd.to_datetime(valor, dayfirst=True, errors="coerce")
    if pd.isna(data):
        return None
    return data.to_period("M")


# ---------------------------------------------------------------- contexto/sessao

def extrair_contexto_valores(res):
    """Le o resultado_valor_global (dict) e devolve os valores importados da
    apuracao. Nunca inventa: so marca disponivel o que existe."""
    res = res or {}
    if not isinstance(res, dict):
        res = {}
    modo = res.get("modo_apuracao", "Completo")
    consolidado = res.get("resultado_consolidado")
    if isinstance(consolidado, dict) and "retroativo_reconhecido" in consolidado:
        # Fonte canonica ja consolidada pela apuracao (_resultado_consolidado):
        # o reconhecido nao e recalculado nem trocado por campo de semantica
        # diferente aqui (mesma regra ja aplicada ao potencial).
        valor_represado = parse_moeda_br(consolidado.get("retroativo_reconhecido"))
    elif modo == "Reduzido por Itens/Estoque":
        valor_represado = parse_moeda_br(res.get("valor_retroativo_estimado_itens_estoque", 0))
    else:
        valor_represado = parse_moeda_br(res.get("valor_represado_a_pagar", res.get("delta_total", 0)))
    variacao = parse_moeda_br(res.get("variacao_acumulada", res.get("fator_acumulado", 1) - 1 if res.get("fator_acumulado") else 0))
    indice = texto_seguro(res.get("indice", ""), "não informado")
    quantidade_ciclos = texto_seguro(res.get("quantidade_ciclos", ""), "[campo a preencher]")
    return {
        "disponivel": bool(res),
        "resultado": res,
        "valor_represado": valor_represado,
        "valor_retroativo_estimado_itens_estoque": parse_moeda_br(res.get("valor_retroativo_estimado_itens_estoque", 0)),
        "variacao": variacao,
        "indice": indice,
        "quantidade_ciclos": quantidade_ciclos,
    }


def carregar_itens_pc_da_sessao(resultado, diagnostico=None):
    """Reutiliza os Pedidos de Compra ja estruturados (itens_PC) lidos da Coleta,
    sem redigitacao de NUMERO_PC/DATA_PC/VALOR_PC. Procura os registros no
    resultado e no diagnostico da Coleta ja disponivel em sessao."""
    fontes = []
    if isinstance(resultado, dict):
        fontes.append(resultado)
    if isinstance(diagnostico, dict):
        fontes.append(diagnostico)
    for fonte in fontes:
        ipc = fonte.get("itens_pc_v10")
        if isinstance(ipc, dict) and isinstance(ipc.get("itens"), list) and ipc["itens"]:
            return ipc["itens"]
        if isinstance(fonte.get("itens_pc"), list) and fonte["itens_pc"]:
            return fonte["itens_pc"]
    return []


# ---------------------------------------------------------------- base financeira

_COLS_JANELA_FIN = ["_periodo", "Competência", "Valor pago/medido", "Situação", "valor"]


def financeiro_por_competencia(resultado):
    """Consolida o historico financeiro por competencia preservando ZERO x VAZIO.

    Cada competencia vira uma linha com:
      - valor: soma dos registros EFETIVAMENTE informados (None se nenhum foi);
      - Situação: "Informado", "Zero informado" ou "Sem informação".
    Multiplos registros na mesma competencia sao somados (apenas os informados).
    Nao elimina zeros e nao inventa meses ausentes.
    """
    cols = ["_periodo", "valor", "Situação"]
    if not isinstance(resultado, dict):
        return pd.DataFrame(columns=cols), "resultado_valor_global.df_financeiro_mensal indisponível"
    df = resultado.get("df_financeiro_mensal")
    origem = "resultado_valor_global.df_financeiro_mensal"
    if not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame(columns=cols), origem + " indisponível"
    col_comp = localizar_coluna(df, ["Competência", "Competencia", "Mês/Ano", "Mes/Ano", "Mês", "Mes"])
    col_valor = localizar_coluna(df, ["Valor bruto medido/aprovado por competência", "Valor bruto medido", "Valor medido/aprovado", "Valor pago/faturado", "Valor bruto faturado", "Valor faturado", "Valor pago", "Valor medido", "Valor"])
    if col_comp is None or col_valor is None:
        return pd.DataFrame(columns=cols), origem + " sem colunas reconhecidas"
    reg = {}
    for _, row in df.iterrows():
        periodo = normalizar_competencia(row.get(col_comp))
        if periodo is None:
            continue
        bruto = row.get(col_valor)
        d = reg.setdefault(periodo, {"soma": 0.0, "tem_informado": False})
        if valor_original_foi_informado(bruto):
            d["tem_informado"] = True
            d["soma"] += parse_moeda_br(bruto)
    linhas = []
    for periodo in sorted(reg):
        d = reg[periodo]
        if d["tem_informado"]:
            valor = round(d["soma"], 2)
            situacao = "Zero informado" if abs(valor) < 0.005 else "Informado"
        else:
            valor = None
            situacao = "Sem informação"
        linhas.append({"_periodo": periodo, "valor": valor, "Situação": situacao})
    return pd.DataFrame(linhas), origem


def janela_6_competencias(por_comp, n=6):
    """Adaptador de apresentacao: delega a janela ao motor
    (janela_financeira_competencias) e devolve um DataFrame pronto para a UI.

    A logica normativa (n competencias-calendario terminando na ultima informada,
    ZERO x VAZIO, sem puxar competencia anterior) vive e e testada no motor.
    """
    if not isinstance(por_comp, pd.DataFrame) or por_comp.empty:
        return pd.DataFrame(columns=_COLS_JANELA_FIN)
    pares = [
        (p.year, p.month, (None if pd.isna(v) else float(v)))
        for p, v in zip(por_comp["_periodo"], por_comp["valor"])
    ]
    janela = janela_financeira_competencias(pares, n)
    linhas = []
    for c in janela["competencias"]:
        periodo = pd.Period(year=c["ano"], month=c["mes"], freq="M")
        linhas.append({
            "_periodo": periodo,
            "Competência": periodo_para_label(periodo),
            "Valor pago/medido": c["valor"],
            "Situação": c["situacao"],
            "valor": c["valor"],
        })
    df = pd.DataFrame(linhas, columns=_COLS_JANELA_FIN)
    # Mantem "Sem informação" como None (nao NaN): evita que o mes vazio vire
    # NaN no float64 e contamine media/exports.
    for col in ("Valor pago/medido", "valor"):
        df[col] = df[col].astype(object).where(df[col].notna(), None)
    return df


def valores_informados_da_janela(ultimos):
    """Lista de valores INFORMADOS (inclui zero) da janela financeira, para a
    media (media_financeiro). Espelha o que a WEB usava inline: preserva
    ZERO x VAZIO (None fica fora do denominador)."""
    if not isinstance(ultimos, pd.DataFrame) or ultimos.empty:
        return []
    return [float(v) for v in ultimos["valor"].tolist() if pd.notna(v)]


def situacao_financeira_considerada(valor):
    """Situacao do valor EFETIVAMENTE considerado na adequacao (coerencia visual
    ZERO x VAZIO): vazio -> "Sem informação"; 0 -> "Zero informado";
    valor != 0 -> "Informado". Delega a distincao ao motor
    (valor_original_foi_informado); nunca converte vazio em zero."""
    if not valor_original_foi_informado(valor):
        return "Sem informação"
    return "Zero informado" if abs(parse_moeda_br(valor)) < 0.005 else "Informado"


def atualizar_exclusoes_manuais_pc(linhas, exclusoes_anteriores):
    """Estado V3 dos PCs: separa ELEGIBILIDADE TEMPORAL de EXCLUSAO MANUAL.

    `linhas`: iteravel de dict {"pc": str, "eligivel": bool, "usar": bool}.
      - eligivel = PC dentro da janela historica ATUAL (situacao "Considerado");
      - usar = estado do checkbox USAR na linha.
    Retorna o conjunto de exclusoes MANUAIS (voluntarias) atualizado:
      - PC elegivel e usar=False -> exclusao manual (add);
      - PC elegivel e usar=True  -> remove exclusao manual (discard);
      - PC fora da janela        -> PRESERVA o estado manual anterior. Um PC nao
        vira exclusao so por estar fora da janela, nem volta so por entrar nela.
    A janela permanece soberana; a media e calculada pelo motor com estas
    exclusoes voluntarias (pedidos_de_itens_pc(exclusoes=...)).
    """
    manual = set(str(e) for e in (exclusoes_anteriores or []))
    for linha in linhas:
        pc = str(linha.get("pc", ""))
        if not pc:
            continue
        if linha.get("eligivel"):
            if linha.get("usar"):
                manual.discard(pc)
            else:
                manual.add(pc)
        # fora da janela: preserva o estado anterior (nao mexe em `manual`)
    return manual


# ------------------------------------------------- cadencia (Etapa 51B)

def ciclos_para_cadencia(df_ciclos):
    """Converte resultado["df_ciclos"] no calendario de ciclos do motor.

    Usa a "Data-base" de cada ciclo como inicio; o fim e o mes anterior ao
    inicio do ciclo seguinte (ultimo ciclo: inicio + 11 meses). Ciclos
    preclusos ENTRAM (evidencia historica de execucao); a flag e apenas
    informativa. Linhas sem data aproveitam o ciclo anterior + 12 meses.
    """
    if not isinstance(df_ciclos, pd.DataFrame) or df_ciclos.empty:
        return []
    col_ciclo = localizar_coluna(df_ciclos, ["Ciclo"])
    col_base = localizar_coluna(df_ciclos, ["Data-base", "Data base"])
    col_sit = localizar_coluna(df_ciclos, ["Situação automática", "Situação"])
    if col_ciclo is None or col_base is None:
        return []
    inicios = []
    for _, row in df_ciclos.iterrows():
        nome = texto_seguro(row.get(col_ciclo), "")
        if not nome:
            continue
        data = pd.to_datetime(row.get(col_base), dayfirst=True, errors="coerce")
        situacao = normalizar_texto(row.get(col_sit)) if col_sit else ""
        inicios.append({
            "nome": nome,
            "inicio": (None if pd.isna(data) else data.to_period("M")),
            "precluso": "preclus" in situacao,
        })
    for i, item in enumerate(inicios):
        if item["inicio"] is None and i > 0 and inicios[i - 1]["inicio"] is not None:
            item["inicio"] = inicios[i - 1]["inicio"] + 12
    inicios = [i for i in inicios if i["inicio"] is not None]
    ciclos = []
    for i, item in enumerate(inicios):
        if i + 1 < len(inicios):
            fim = inicios[i + 1]["inicio"] - 1
        else:
            fim = item["inicio"] + 11
        ciclos.append(CicloCadencia(
            nome=item["nome"],
            inicio=item["inicio"].to_timestamp().date(),
            fim=fim.to_timestamp().date(),
            precluso=item["precluso"],
        ))
    return ciclos


def pares_de_financeiro(fin_por_comp):
    """Historico financeiro consolidado -> pares (ano, mes, valor) do motor."""
    if not isinstance(fin_por_comp, pd.DataFrame) or fin_por_comp.empty:
        return []
    pares = []
    for periodo, valor in zip(fin_por_comp["_periodo"], fin_por_comp["valor"]):
        if periodo is None or pd.isna(valor) or valor is None:
            continue
        pares.append((int(periodo.year), int(periodo.month), float(valor)))
    return pares


def pares_de_pedidos(pedidos):
    """Pedidos considerados -> pares (ano, mes, valor) do motor."""
    pares = []
    for p in (pedidos or []):
        if not getattr(p, "considerar", True) or getattr(p, "data", None) is None:
            continue
        pares.append((p.data.year, p.data.month, float(p.valor or 0.0)))
    return pares


def _mapa_base_por_periodo(base_por_competencia):
    """Normaliza {date/Period -> valor} em {(ano, mes) -> valor}."""
    mapa = {}
    for chave, valor in (base_por_competencia or {}).items():
        periodo = normalizar_competencia(chave)
        if periodo is not None:
            mapa[(int(periodo.year), int(periodo.month))] = float(valor or 0.0)
    return mapa


# ---------------------------------------------------------------- projecao

def gerar_periodos_projecao(ultima_competencia, data_final_vigencia):
    if ultima_competencia is None:
        return []
    fim = periodo_de_data_final(data_final_vigencia)
    if fim is None:
        return []
    inicio = pd.Period(ultima_competencia, freq="M") + 1
    if fim < inicio:
        return []
    return list(pd.period_range(inicio, fim, freq="M"))


def montar_base_editor(periodos, media_mensal, base_por_competencia=None, rotulo_base=None):
    """Editor da projecao. Sem `base_por_competencia` (legado): a media e a
    base sugerida de todos os meses. Com o mapa (cadencia/manual): cada mes
    exibe a SUA base (0 = sem ocorrencia prevista)."""
    if base_por_competencia is None:
        return pd.DataFrame([{
            "Competência": periodo_para_label(p),
            "Base automática pela média": moeda(media_mensal, com_prefixo=False),
            "Valor informado pelo fiscal": "",
            "Premissa do valor informado": "Valor sem reajuste",
            "Observação": "",
        } for p in periodos])
    mapa = _mapa_base_por_periodo(base_por_competencia)
    return pd.DataFrame([{
        "Competência": periodo_para_label(p),
        "Base automática pela média": moeda(
            mapa.get((int(p.year), int(p.month)), 0.0), com_prefixo=False),
        "Valor informado pelo fiscal": "",
        "Premissa do valor informado": "Valor sem reajuste",
        "Observação": "",
    } for p in periodos])


def calcular_projecao(df_editor, media_mensal, fator_reajuste,
                      base_por_competencia=None, origem_automatica=None):
    """Projecao mes a mes.

    Caminho legado (sem `base_por_competencia`): base automatica = media em
    todos os meses (comportamento historico, usado quando a premissa e
    MENSAL). Caminho por cadencia/manual: a base automatica de cada mes vem
    do mapa {competencia -> valor}; meses sem ocorrencia prevista tem base 0
    — a projecao deixa de presumir mensalidade (Etapa 51B). Overrides do
    fiscal (ZERO != VAZIO) prevalecem em ambos os caminhos.
    """
    linhas = []
    if not isinstance(df_editor, pd.DataFrame) or df_editor.empty:
        return pd.DataFrame(columns=["Competência", "Origem", "Premissa usada", "Valor base considerado", "Valor reajustado estimado", "Diferença futura a adequar", "Observação"])
    mapa = None if base_por_competencia is None else _mapa_base_por_periodo(base_por_competencia)
    for _, row in df_editor.iterrows():
        competencia = texto_seguro(row.get("Competência"), "")
        informado_raw = row.get("Valor informado pelo fiscal", "")
        informado = parse_moeda_br(informado_raw)
        premissa = texto_seguro(row.get("Premissa do valor informado"), "Valor sem reajuste")
        observacao = texto_seguro(row.get("Observação"), "")
        # Base G, valor reajustado H = ROUND(G*fator,2) e diferenca I = ROUND(H-G,2)
        # seguem o golden; o arredondamento e o do motor (_round2, paridade Excel).
        # ZERO != VAZIO: existencia do override e decidida pelo valor ORIGINAL
        # (0 informado significa execucao zero; vazio => projecao automatica).
        if valor_original_foi_informado(informado_raw):
            origem = "Valor informado pelo fiscal"
            if premissa == "Valor já reajustado":
                base_considerada = informado / fator_reajuste if fator_reajuste else informado
                premissa_usada = "Valor já reajustado"
            else:
                base_considerada = informado
                premissa_usada = "Valor sem reajuste"
        elif mapa is None:
            origem = "Média dos últimos 6 meses"
            base_considerada = media_mensal
            premissa_usada = "Média sem reajuste"
        else:
            periodo = normalizar_competencia(competencia)
            chave = (int(periodo.year), int(periodo.month)) if periodo is not None else None
            base_considerada = mapa.get(chave, 0.0)
            origem = origem_automatica or "Projeção pela cadência histórica"
            premissa_usada = ("Ocorrência prevista pela cadência"
                             if base_considerada > 0 else "Sem ocorrência prevista")
        valor_reajustado = _round2(base_considerada * fator_reajuste)
        diferenca = _round2(valor_reajustado - base_considerada)
        linhas.append({
            "Competência": competencia, "Origem": origem, "Premissa usada": premissa_usada,
            "Valor base considerado": _round2(base_considerada),
            "Valor reajustado estimado": valor_reajustado,
            "Diferença futura a adequar": diferenca,
            "Observação": observacao,
        })
    return pd.DataFrame(linhas)


def cronograma_por_exercicio(df_projecao, retroativo):
    # Programacao por exercicio (golden): diferencas futuras de cada ano + o
    # retroativo somente no PRIMEIRO exercicio da projecao (nao no ano corrente).
    linhas = {}
    anos = []
    if isinstance(df_projecao, pd.DataFrame) and not df_projecao.empty:
        for _, row in df_projecao.iterrows():
            periodo = normalizar_competencia(row.get("Competência"))
            if periodo is None:
                continue
            ano = int(periodo.year)
            anos.append(ano)
            linhas[ano] = linhas.get(ano, 0.0) + parse_moeda_br(row.get("Diferença futura a adequar", 0))
    if anos:
        primeiro_exercicio = min(anos)
        linhas[primeiro_exercicio] = linhas.get(primeiro_exercicio, 0.0) + float(retroativo or 0)
    return pd.DataFrame([
        {"Exercício": str(ano), "Valor": _round2(valor)}
        for ano, valor in sorted(linhas.items()) if abs(valor) > 0.004
    ])


# ---------------------------------------------------------------- XLSX

def gerar_xlsx_projecao(df_ultimos, df_projecao, resumo):
    output = BytesIO()
    with pd.ExcelWriter(
        output,
        engine="xlsxwriter",
        engine_kwargs=opcoes_excel_writer_seguro(),
    ) as writer:
        workbook = writer.book
        fmt_title = workbook.add_format({"bold": True, "font_size": 14, "font_color": "#0B1F3A"})
        fmt_header = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
        fmt_money = workbook.add_format({"num_format": 'R$ #,##0.00', "border": 1})
        fmt_text = workbook.add_format({"border": 1})
        fmt_note = workbook.add_format({"italic": True, "font_color": "#64748B"})
        fmt_total = workbook.add_format({"num_format": 'R$ #,##0.00', "border": 1, "bold": True, "bg_color": "#EAF2F8"})
        ws = workbook.add_worksheet("RESUMO")
        writer.sheets["RESUMO"] = ws
        ws.write(0, 0, "Adequação Orçamentária — Delta do Reajuste", fmt_title)
        ws.write(1, 0, "Complementação confirmada = retroativo reconhecido considerado + "
                       "diferença futura projetada. O retroativo potencial, quando existente, "
                       "é informado à parte e não integra a complementação.", fmt_note)
        ws.write(3, 0, "Indicador", fmt_header)
        ws.write(3, 1, "Valor", fmt_header)
        for r, (label, valor) in enumerate(resumo, start=4):
            ws.write(r, 0, label, fmt_text)
            if isinstance(valor, (int, float)):
                fmt = fmt_total if "Complementação" in label else fmt_money
                ws.write_number(r, 1, float(valor), fmt)
            else:
                ws.write(r, 1, valor, fmt_text)
        ws.set_column("A:A", 42)
        ws.set_column("B:B", 26)
        # MEDIA: as competencias de referencia com a mesma transparencia da WEB
        # (valor informado / zero informado / sem informacao) — WEB e XLSX contam
        # a mesma historia. "Sem informação" fica em branco (nao vira zero).
        _tem_situacao = isinstance(df_ultimos, pd.DataFrame) and "Situação" in getattr(df_ultimos, "columns", [])
        _cols_media = ["Competência", "Valor pago/medido"] + (["Situação"] if _tem_situacao else [])
        df_ult = df_ultimos[_cols_media].copy() if isinstance(df_ultimos, pd.DataFrame) and not df_ultimos.empty else pd.DataFrame(columns=_cols_media)
        # "Sem informação" => celula em branco (nunca zero); evita NaN no xlsxwriter.
        df_ult = df_ult.astype(object).where(pd.notna(df_ult), None)
        df_ult.to_excel(writer, sheet_name="MEDIA", index=False)
        ws_m = writer.sheets["MEDIA"]
        for c, col in enumerate(df_ult.columns):
            ws_m.write(0, c, col, fmt_header)
        ws_m.set_column("A:A", 18)
        ws_m.set_column("B:B", 22, fmt_money)
        if _tem_situacao:
            ws_m.set_column("C:C", 18)
        df_proj = df_projecao.copy() if isinstance(df_projecao, pd.DataFrame) else pd.DataFrame()
        df_proj.to_excel(writer, sheet_name="PROJECAO", index=False)
        ws_p = writer.sheets["PROJECAO"]
        for c, col in enumerate(df_proj.columns):
            ws_p.write(0, c, col, fmt_header)
        for c, col in enumerate(df_proj.columns):
            if col in ["Valor base considerado", "Valor reajustado estimado", "Diferença futura a adequar"]:
                ws_p.set_column(c, c, max(18, min(32, len(str(col)) + 4)), fmt_money)
                for r_idx, valor in enumerate(df_proj[col], start=1):
                    ws_p.write_number(r_idx, c, parse_moeda_br(valor), fmt_money)
            else:
                ws_p.set_column(c, c, max(16, min(32, len(str(col)) + 4)))
    output.seek(0)
    return output.getvalue()
