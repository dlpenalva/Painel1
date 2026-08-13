"""Motor da Calculadora de Garantia Contratual — fluxo único, 100% manual.

Etapa 49. A ferramenta compara duas fotografias informadas manualmente pelo
usuário:

    SITUAÇÃO ATUAL DO CONTRATO   x   GARANTIA ATUALMENTE CONSTITUÍDA

e responde objetivamente se falta dinheiro, se falta prazo, se faltam os dois ou
se não há nada a regularizar.

    GARANTIA NECESSÁRIA = ARREDONDAR(VALOR TOTAL ATUAL DO CONTRATO x PERCENTUAL; 2)
    COBERTURA ATUAL     = soma do valor total vigente das garantias INDEPENDENTES
    COMPLEMENTO         = max(GARANTIA NECESSÁRIA - COBERTURA ATUAL; 0)
    VALIDADE MÍNIMA     = TÉRMINO DA VIGÊNCIA CONTRATUAL + 90 DIAS CORRIDOS

Endossos sucessivos da MESMA garantia nunca se somam: o que compõe a cobertura é
o valor total vigente da garantia após o último endosso. Duas linhas com a mesma
identificação descrevem a mesma cadeia de garantia e são consolidadas na última
informada (ver ``consolidar_garantias``). Garantias independentes — identificações
distintas, ou linhas sem identificação — somam-se normalmente.

A suficiência financeira e a suficiência temporal são verificações INDEPENDENTES:
uma garantia pode ter valor bastante e prazo curto, e vice-versa.

Nenhuma função deste módulo lê o VTA, a Coleta, os RESULTADOS ou qualquer outra
fonte do sistema: todos os dados chegam por parâmetro, informados à mão.

Arredondamento financeiro: Decimal + ROUND_HALF_UP, duas casas. O float nunca é a
base do cálculo monetário — apenas a entrada digitada e a formatação final.

Módulo puro (sem Streamlit) para permitir testes focais.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation


PERCENTUAL_GARANTIA_PADRAO = Decimal("5.00")   # padrão editável pelo usuário
DIAS_VALIDADE_MINIMA = 90                      # dias corridos após o término da vigência
CENTAVO = Decimal("0.01")
CEM = Decimal("100")
TRACO = "—"

CLAUSULA_GARANTIA = "Cláusula 10 – Garantia Contratual"

# Nomes das colunas da grade de garantias vigentes (a página e o motor
# compartilham estas constantes para não divergirem).
COLUNA_IDENTIFICACAO = "Identificação da garantia"
COLUNA_VALOR = "Valor total atualmente garantido"
COLUNA_VALIDADE = "Validade"

# Diagnóstico: exatamente quatro resultados possíveis.
DIAGNOSTICO_REGULAR = "GARANTIA REGULAR"
DIAGNOSTICO_VALOR = "ATUALIZAR VALOR"
DIAGNOSTICO_VALIDADE = "ATUALIZAR VALIDADE"
DIAGNOSTICO_VALOR_E_VALIDADE = "ATUALIZAR VALOR E VALIDADE"

SEM_NECESSIDADE_DE_ATUALIZACAO = "Não foi identificada necessidade de atualização da garantia."


# ============================================================
# Conversão monetária e arredondamento financeiro
# ============================================================

def parse_moeda_br(valor):
    """Converte entrada monetária brasileira em ``Decimal``.

    Aceita: 2968866, 2968866,00, 2.968.866,00, R$ 2.968.866,00, 2968866.00 e
    valores numéricos (int/float/Decimal). Retorna ``None`` quando o valor não
    puder ser interpretado (para permitir distinção entre "zero" e "inválido").
    """
    if valor is None:
        return None
    if isinstance(valor, Decimal):
        return valor
    if isinstance(valor, bool):
        return None
    if isinstance(valor, int):
        return Decimal(valor)
    if isinstance(valor, float):
        if valor != valor:  # NaN
            return None
        return Decimal(str(valor))

    texto = str(valor).strip()
    if not texto:
        return None
    texto = (
        texto.replace("R$", "")
        .replace("r$", "")
        .replace(" ", "")
        .replace(" ", "")
    )
    if not texto:
        return None

    if "," in texto:
        # Formato BR: ponto é separador de milhar, vírgula é decimal.
        texto = texto.replace(".", "").replace(",", ".")
    elif texto.count(".") > 1:
        # Múltiplos pontos sem vírgula: pontos são separadores de milhar.
        texto = texto.replace(".", "")
    # Caso restante: inteiro ou decimal com ponto único.

    try:
        return Decimal(texto)
    except (InvalidOperation, ValueError):
        return None


def arredondar_financeiro(valor):
    """Arredonda para 2 casas com ROUND_HALF_UP (arredondamento financeiro)."""
    if not isinstance(valor, Decimal):
        valor = parse_moeda_br(valor)
        if valor is None:
            valor = Decimal("0")
    return valor.quantize(CENTAVO, rounding=ROUND_HALF_UP)


def moeda(valor, com_prefixo=True):
    """Formata como moeda brasileira: R$ 2.968.866,00."""
    dec = parse_moeda_br(valor)
    if dec is None:
        dec = Decimal("0")
    dec = arredondar_financeiro(dec)
    texto = f"{dec:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {texto}" if com_prefixo else texto


def formatar_brl(valor):
    """Função única de formatação monetária brasileira: R$ 2.968.866,00."""
    return moeda(valor, com_prefixo=True)


def parse_percentual(valor):
    """Converte o percentual informado em ``Decimal`` de pontos percentuais.

    5 / "5" / "5,00" / "5%" -> Decimal("5"). Retorna ``None`` quando inválido ou
    fora do intervalo (0, 100].
    """
    if isinstance(valor, str):
        valor = valor.replace("%", "")
    pct = parse_moeda_br(valor)
    if pct is None or pct <= 0 or pct > CEM:
        return None
    return pct


def formatar_percentual(percentual):
    """Formata o percentual sem zeros decimais inúteis: 5 -> "5"; 4,75 -> "4,75"."""
    pct = parse_percentual(percentual)
    if pct is None:
        return TRACO
    normalizado = pct.normalize()
    if normalizado == normalizado.to_integral_value():
        normalizado = normalizado.to_integral_value()
    return format(normalizado, "f").replace(".", ",")


# ============================================================
# Datas
# ============================================================

def parse_data_br(valor):
    """Converte a data informada em ``date``; ``None`` quando ausente/inválida.

    Aceita ``date``/``datetime`` (inclusive ``pandas.Timestamp``, que herda de
    ``datetime``), "31/12/2026" e "2026-12-31". ``NaT`` e vazio viram ``None``.
    """
    if valor is None:
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    if valor != valor:  # NaT/NaN
        return None

    texto = str(valor).strip()
    if not texto or texto.lower() in {"nat", "nan", "none"}:
        return None
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(texto[:10], formato).date()
        except ValueError:
            continue
    return None


def formatar_data_br(valor):
    """Formata a data como 31/12/2026; ``None`` vira travessão."""
    data = parse_data_br(valor)
    return data.strftime("%d/%m/%Y") if data else TRACO


def calcular_validade_minima(data_fim_vigencia):
    """Validade mínima exigida = término da vigência + 90 dias corridos."""
    data = parse_data_br(data_fim_vigencia)
    if data is None:
        return None
    return data + timedelta(days=DIAS_VALIDADE_MINIMA)


# ============================================================
# Garantias vigentes (entrada manual)
# ============================================================

def normalizar_garantias(registros):
    """Normaliza as linhas da grade de garantias vigentes.

    ``registros``: iterável de dicts com ``COLUNA_IDENTIFICACAO`` (opcional),
    ``COLUNA_VALOR`` e ``COLUNA_VALIDADE``. Retorna ``(linhas, avisos)``, onde
    cada linha é ``{"identificacao": str, "valor": Decimal, "validade": date|None}``.

    Linhas totalmente vazias são ignoradas em silêncio. A validade ausente não
    descarta a linha (o valor continua compondo a cobertura), mas é tratada como
    prazo NÃO comprovado na avaliação temporal — fail-closed.
    """
    linhas = []
    avisos = []
    vazios = {"", "nan", "none", "<na>", "nat"}
    for idx, reg in enumerate(registros, start=1):
        identificacao_bruta = reg.get(COLUNA_IDENTIFICACAO)
        identificacao = "" if identificacao_bruta is None else str(identificacao_bruta).strip()
        if identificacao.lower() in vazios:
            identificacao = ""

        valor_bruto = reg.get(COLUNA_VALOR)
        texto_valor = "" if valor_bruto is None else str(valor_bruto).strip()
        tem_texto_valor = texto_valor.lower() not in vazios
        validade = parse_data_br(reg.get(COLUNA_VALIDADE))

        if not identificacao and not tem_texto_valor and validade is None:
            continue  # linha vazia

        rotulo = identificacao or f"Garantia da linha {idx}"

        if not tem_texto_valor:
            avisos.append(f"{rotulo}: informe o valor total atualmente garantido.")
            continue

        valor = parse_moeda_br(valor_bruto)
        if valor is None:
            avisos.append(
                f'{rotulo}: o valor "{texto_valor}" não pôde ser interpretado. '
                "Use o formato R$ 50.000,00."
            )
            continue
        if valor < 0:
            avisos.append(f"{rotulo}: o valor total garantido não pode ser negativo.")
            continue

        if validade is None:
            avisos.append(f"{rotulo}: informe a data de validade da garantia.")

        linhas.append(
            {
                "identificacao": identificacao,
                "valor": arredondar_financeiro(valor),
                "validade": validade,
            }
        )

    return linhas, avisos


def consolidar_garantias(linhas):
    """Aplica a regra do endosso: cada cadeia de garantia entra uma única vez.

    Linhas que compartilham a MESMA identificação (comparação sem diferenciar
    maiúsculas/minúsculas nem espaços nas pontas) descrevem a mesma garantia em
    momentos diferentes — original e endossos sucessivos. Prevalece a ÚLTIMA
    linha informada, que traz o valor total vigente após o último endosso; as
    anteriores são descartadas para não duplicar a cobertura.

    Linhas sem identificação são sempre tratadas como garantias independentes.
    Retorna ``(consolidadas, avisos)`` preservando a ordem de entrada.
    """
    consolidadas = []
    avisos = []
    posicao_por_chave = {}
    for linha in linhas:
        chave = linha["identificacao"].strip().casefold()
        if chave and chave in posicao_por_chave:
            posicao = posicao_por_chave[chave]
            anterior = consolidadas[posicao]
            avisos.append(
                f"{linha['identificacao']}: informada mais de uma vez. Considerado apenas o valor "
                f"total vigente mais recente ({formatar_brl(linha['valor'])}); "
                f"{formatar_brl(anterior['valor'])} foi desconsiderado para não duplicar a cobertura."
            )
            consolidadas[posicao] = linha
            continue
        if chave:
            posicao_por_chave[chave] = len(consolidadas)
        consolidadas.append(linha)
    return consolidadas, avisos


def calcular_cobertura_atual(garantias):
    """Soma o valor total vigente das garantias independentes."""
    total = Decimal("0")
    for garantia in garantias:
        total += garantia["valor"]
    return arredondar_financeiro(total)


# ============================================================
# Cálculo financeiro e temporal
# ============================================================

def calcular_garantia_necessaria(valor_total_contrato, percentual=PERCENTUAL_GARANTIA_PADRAO):
    """Garantia necessária = ARREDONDAR(valor total atual x percentual; 2)."""
    base = parse_moeda_br(valor_total_contrato)
    if base is None:
        base = Decimal("0")
    pct = parse_percentual(percentual)
    if pct is None:
        pct = PERCENTUAL_GARANTIA_PADRAO
    return arredondar_financeiro(base * pct / CEM)


def calcular_complemento(garantia_necessaria, cobertura_atual):
    """Complemento = max(garantia necessária - cobertura atual; 0). Nunca negativo."""
    diferenca = arredondar_financeiro(garantia_necessaria) - arredondar_financeiro(cobertura_atual)
    if diferenca <= 0:
        return Decimal("0.00")
    return arredondar_financeiro(diferenca)


def diagnosticar(valor_suficiente, validade_suficiente):
    """Combina as duas verificações independentes nos quatro resultados possíveis."""
    if valor_suficiente and validade_suficiente:
        return DIAGNOSTICO_REGULAR
    if valor_suficiente:
        return DIAGNOSTICO_VALIDADE
    if validade_suficiente:
        return DIAGNOSTICO_VALOR
    return DIAGNOSTICO_VALOR_E_VALIDADE


def analisar_garantia(valor_total_contrato, percentual, data_fim_vigencia, garantias):
    """Análise completa da garantia a partir dos dados manuais.

    ``garantias`` são as linhas já normalizadas por ``normalizar_garantias``; a
    consolidação dos endossos é aplicada aqui. Exige valor total do contrato
    maior que zero e término da vigência informado — a página valida antes.

    A avaliação temporal considera insuficiente qualquer garantia que componha a
    cobertura e vença antes da validade mínima (ou cuja validade não tenha sido
    informada). Sem nenhuma garantia cadastrada não há prazo a regularizar: a
    pendência é apenas de valor.
    """
    valor_total = parse_moeda_br(valor_total_contrato)
    if valor_total is None or valor_total <= 0:
        raise ValueError("O valor total atual do contrato deve ser maior que zero.")
    pct = parse_percentual(percentual)
    if pct is None:
        raise ValueError("O percentual da garantia deve estar entre 0 e 100.")
    validade_minima = calcular_validade_minima(data_fim_vigencia)
    if validade_minima is None:
        raise ValueError("Informe o término da vigência contratual.")

    valor_total = arredondar_financeiro(valor_total)
    consolidadas, avisos = consolidar_garantias(list(garantias))

    garantia_necessaria = calcular_garantia_necessaria(valor_total, pct)
    cobertura_atual = calcular_cobertura_atual(consolidadas)
    complemento = calcular_complemento(garantia_necessaria, cobertura_atual)
    valor_suficiente = cobertura_atual >= garantia_necessaria

    detalhadas = []
    for garantia in consolidadas:
        validade = garantia["validade"]
        suficiente = validade is not None and validade >= validade_minima
        detalhadas.append({**garantia, "validade_suficiente": suficiente})
    validade_suficiente = all(item["validade_suficiente"] for item in detalhadas)

    return {
        "valor_total_contrato": valor_total,
        "percentual": pct,
        "garantia_necessaria": garantia_necessaria,
        "cobertura_atual": cobertura_atual,
        "complemento": complemento,
        "valor_suficiente": valor_suficiente,
        "data_fim_vigencia": parse_data_br(data_fim_vigencia),
        "validade_minima": validade_minima,
        "validade_suficiente": validade_suficiente,
        "garantias": detalhadas,
        "quantidade_garantias": len(detalhadas),
        "avisos_consolidacao": avisos,
        "diagnostico": diagnosticar(valor_suficiente, validade_suficiente),
    }


# ============================================================
# Texto para comunicação à contratada
# ============================================================

def _paragrafo_cobertura(analise):
    """Parágrafo da cobertura atual, com a concordância adequada."""
    quantidade = analise["quantidade_garantias"]
    complemento = moeda(analise["complemento"])
    if quantidade == 0:
        return (
            "Não há garantia contratual atualmente apresentada, sendo necessária, portanto, a "
            f"apresentação de garantia no valor de {complemento}."
        )
    cobertura = moeda(analise["cobertura_atual"])
    if quantidade > 1:
        return (
            f"As garantias atualmente apresentadas totalizam {cobertura}, sendo necessária, "
            f"portanto, a complementação no valor de {complemento}."
        )
    return (
        f"A garantia atualmente apresentada é de {cobertura}, sendo necessária, portanto, a "
        f"complementação no valor de {complemento}."
    )


def gerar_texto_comunicacao(analise):
    """Texto pronto para copiar e colar em e-mail, conforme o diagnóstico.

    Sem pendência, a comunicação apenas reflete a conclusão. Havendo apenas
    prazo a regularizar, nenhuma complementação financeira é mencionada.
    """
    diagnostico = analise["diagnostico"]
    if diagnostico == DIAGNOSTICO_REGULAR:
        return SEM_NECESSIDADE_DE_ATUALIZACAO

    data_minima = formatar_data_br(analise["validade_minima"])
    fecho = [
        "Gentileza encaminhar o respectivo endosso/comprovante após a regularização.",
        "",
        "Atenciosamente,",
    ]

    if diagnostico == DIAGNOSTICO_VALIDADE:
        return "\n".join(
            [
                "Prezados,",
                "",
                "Considerando a vigência atual do contrato, verificamos a necessidade de atualização "
                "da validade da garantia contratual.",
                "",
                f"Nos termos da {CLAUSULA_GARANTIA}, a garantia deverá possuir validade mínima até "
                f"{data_minima}, correspondente a {DIAS_VALIDADE_MINIMA} dias após o término da "
                "vigência contratual.",
                "",
                *fecho,
            ]
        )

    abertura = (
        "Considerando a atualização do valor total do contrato para "
        f"{moeda(analise['valor_total_contrato'])}, a garantia contratual correspondente a "
        f"{formatar_percentual(analise['percentual'])}% passa a ser de "
        f"{moeda(analise['garantia_necessaria'])}."
    )
    if diagnostico == DIAGNOSTICO_VALOR_E_VALIDADE:
        solicitacao = (
            f"Solicitamos a atualização da garantia, nos termos da {CLAUSULA_GARANTIA}, bem como a "
            f"atualização da sua validade, que deverá alcançar, no mínimo, {data_minima}, "
            f"correspondente a {DIAS_VALIDADE_MINIMA} dias após o término da vigência contratual."
        )
    else:
        solicitacao = (
            f"Solicitamos a atualização da garantia, nos termos da {CLAUSULA_GARANTIA}, observando "
            f"também a validade mínima necessária de {DIAS_VALIDADE_MINIMA} dias após o término da "
            "vigência contratual."
        )

    return "\n".join(
        [
            "Prezados,",
            "",
            abertura,
            "",
            _paragrafo_cobertura(analise),
            "",
            solicitacao,
            "",
            *fecho,
        ]
    )
