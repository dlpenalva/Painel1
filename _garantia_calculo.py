"""Motor da Calculadora de Garantia Contratual — linha do tempo contratual manual.

A ferramenta conta, de cima para baixo, a história do contrato e só então a
confronta com o que a contratada efetivamente apresentou:

    SITUAÇÃO ORIGINAL
        -> ALTERAÇÕES POSTERIORES DO CONTRATO
        -> SITUAÇÃO ATUAL DO CONTRATO
        -> GARANTIA ATUALMENTE APRESENTADA
        -> RESULTADO

Dois conceitos que NUNCA se misturam:

* HISTÓRIA DO CONTRATO — reajuste/apostila, repactuação, aditivo, prorrogação.
  Altera o valor, a vigência ou ambos, e portanto define a garantia EXIGIDA.
* GARANTIA APRESENTADA — apólice, endosso ou outra garantia admitida. Define a
  cobertura EXISTENTE. Um reajuste jamais entra nessa soma.

    GARANTIA EXIGIDA    = ARREDONDAR(VALOR TOTAL DO CONTRATO x PERCENTUAL; 2)
    COBERTURA ATUAL     = soma do valor total vigente das garantias INDEPENDENTES
    COMPLEMENTO         = max(GARANTIA EXIGIDA - COBERTURA ATUAL; 0)
    VALIDADE MÍNIMA     = TÉRMINO DA VIGÊNCIA + 90 DIAS CORRIDOS

Cada evento informa o VALOR TOTAL DO CONTRATO APÓS O EVENTO — nunca o acréscimo
isolado — e o motor deriva a variação contra a situação anterior. A prorrogação
pode não alterar o valor: sem valor informado, herda o valor contratual vigente.

Endossos sucessivos da MESMA garantia nunca se somam: o que compõe a cobertura é
o valor total vigente da garantia após o último endosso. Duas linhas com a mesma
referência descrevem a mesma cadeia e são consolidadas na última informada (ver
``consolidar_garantias``). Garantias independentes — referências distintas, ou
linhas sem referência — somam-se normalmente.

A suficiência financeira e a suficiência temporal são verificações INDEPENDENTES:
uma garantia pode ter valor bastante e prazo curto, e vice-versa.

Nenhuma função deste módulo lê o VTA, a Coleta, os RESULTADOS ou qualquer outra
fonte do sistema: todos os dados chegam por parâmetro, informados à mão.

Arredondamento financeiro: Decimal + ROUND_HALF_UP, duas casas. O float nunca é a
base do cálculo monetário — apenas a entrada digitada e a formatação final.

Datas: a ausência é normalizada UMA única vez, na fronteira de entrada, e vira
sempre ``None``. ``pandas.NaT`` é instância de ``datetime`` e sobrevive a um
``is not None``; por isso a presença de data é decidida por ``parse_data_br``, e
nunca por uma comparação direta com ``None`` — ver ``data_ausente``.

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

# Nomes das colunas da grade da garantia atualmente apresentada (a página e o
# motor compartilham estas constantes para não divergirem).
COLUNA_REFERENCIA = "Apólice / endosso / referência"
COLUNA_VALOR = "Valor garantido"
COLUNA_VALIDADE = "Validade"

# Nomes das colunas do quadro de alterações posteriores à assinatura.
COLUNA_EVENTO_TIPO = "Tipo"
COLUNA_EVENTO_DATA = "Data do instrumento/evento"
COLUNA_EVENTO_VALOR = "Valor total do contrato após o evento"
COLUNA_EVENTO_VIGENCIA = "Novo término da vigência"
COLUNA_EVENTO_OBSERVACAO = "Observação"

# Tipos de evento contratual. Nenhum deles é uma garantia apresentada.
TIPO_REAJUSTE = "Reajuste / Apostila"
TIPO_REPACTUACAO = "Repactuação"
TIPO_ADITIVO = "Aditivo"
TIPO_PRORROGACAO = "Prorrogação"
TIPO_OUTRO = "Outro"
TIPOS_EVENTO = (TIPO_REAJUSTE, TIPO_REPACTUACAO, TIPO_ADITIVO, TIPO_PRORROGACAO, TIPO_OUTRO)

# Eventos que existem justamente para mover o valor do contrato: sem o valor
# total após o evento não há o que registrar.
TIPOS_COM_VALOR_OBRIGATORIO = (TIPO_REAJUSTE, TIPO_REPACTUACAO, TIPO_ADITIVO)

# Diagnóstico consolidado: exatamente quatro resultados possíveis.
DIAGNOSTICO_REGULAR = "GARANTIA REGULAR"
DIAGNOSTICO_VALOR = "ATUALIZAR VALOR"
DIAGNOSTICO_VALIDADE = "ATUALIZAR VALIDADE"
DIAGNOSTICO_VALOR_E_VALIDADE = "ATUALIZAR VALOR E VALIDADE"

# Dimensões separadas: dinheiro e prazo nunca são condensados num único sim/não.
FINANCEIRO_COMPLEMENTAR = "COMPLEMENTAÇÃO NECESSÁRIA"
FINANCEIRO_SUFICIENTE = "COBERTURA FINANCEIRA SUFICIENTE"
FINANCEIRO_SUPERIOR = "COBERTURA FINANCEIRA SUPERIOR À EXIGIDA"

TEMPORAL_SUFICIENTE = "VALIDADE SUFICIENTE"
TEMPORAL_INSUFICIENTE = "PRORROGAÇÃO/ENDOSSO DE VALIDADE NECESSÁRIO"
TEMPORAL_NAO_INFORMADA = "VALIDADE NÃO INFORMADA"

SEM_NECESSIDADE_DE_ATUALIZACAO = (
    "Com os dados apresentados, não foi identificada necessidade de complementação "
    "ou de atualização da garantia contratual."
)


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
        .replace(" ", "")   # espaço não separável (copiar/colar de Excel e web)
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


def formatar_brl_opcional(valor):
    """Como ``formatar_brl``, mas a AUSÊNCIA continua ausência (travessão).

    Ausência não é zero: um evento sem valor informado não pode ser exibido como
    R$ 0,00.
    """
    if valor is None:
        return TRACO
    return formatar_brl(valor)


def formatar_variacao(valor):
    """Formata a variação contratual com o sinal explícito: + R$ 400.000,00."""
    if valor is None:
        return TRACO
    dec = arredondar_financeiro(valor)
    if dec == 0:
        return "Sem alteração"
    sinal = "+" if dec > 0 else "-"
    return f"{sinal} {formatar_brl(abs(dec))}"


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
# Datas — normalização canônica da ausência
# ============================================================

_TEXTOS_DATA_AUSENTE = {"", "nat", "nan", "none", "null", "<na>", TRACO}


def data_ausente(valor):
    """``True`` quando o valor representa AUSÊNCIA de data.

    Trata como ausência ``None``, ``pandas.NaT``, ``NaN``, string vazia e os
    textos que os widgets e DataFrames produzem para célula vazia. ``NaT`` exige
    cuidado especial: ele é instância de ``datetime`` e passa por um
    ``is not None``, mas é o único "datetime" que difere de si mesmo.
    """
    if valor is None:
        return True
    try:
        if valor != valor:   # NaN e NaT
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(valor, (date, datetime)):
        return False
    return str(valor).strip().lower() in _TEXTOS_DATA_AUSENTE


def parse_data_br(valor):
    """Converte a data informada em ``date``; ``None`` quando ausente/inválida.

    Aceita ``date``/``datetime`` (inclusive ``pandas.Timestamp`` válido, que
    herda de ``datetime``), "31/12/2026" e "2026-12-31". ``None``, ``NaT``,
    ``NaN``, vazio e texto inválido viram ``None`` — nunca um sentinela que
    depois exploda numa comparação relacional.
    """
    if data_ausente(valor):
        return None
    if isinstance(valor, datetime):
        convertida = valor.date()
        # Cinto e suspensório: ``NaT.date()`` devolve ``NaT``, não uma data.
        if data_ausente(convertida) or not isinstance(convertida, date):
            return None
        return convertida
    if isinstance(valor, date):
        return valor

    texto = str(valor).strip()
    if texto.lower() in _TEXTOS_DATA_AUSENTE:
        return None
    for formato in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(texto[:10], formato).date()
        except ValueError:
            continue
    return None


def formatar_data_br(valor):
    """Formata a data como 31/12/2026; ausência vira travessão."""
    data = parse_data_br(valor)
    return data.strftime("%d/%m/%Y") if data else TRACO


def calcular_validade_minima(data_fim_vigencia):
    """Validade mínima exigida = término da vigência + 90 dias corridos."""
    data = parse_data_br(data_fim_vigencia)
    if data is None:
        return None
    return data + timedelta(days=DIAS_VALIDADE_MINIMA)


# ============================================================
# Alterações posteriores à assinatura (história do contrato)
# ============================================================

def normalizar_eventos(registros):
    """Normaliza as linhas do quadro de alterações posteriores à assinatura.

    ``registros``: iterável de dicts com ``COLUNA_EVENTO_TIPO``,
    ``COLUNA_EVENTO_DATA``, ``COLUNA_EVENTO_VALOR``, ``COLUNA_EVENTO_VIGENCIA``
    e ``COLUNA_EVENTO_OBSERVACAO``.

    Retorna ``(eventos, avisos, pendencias)``. Cada evento é
    ``{"numero", "tipo", "data", "valor", "vigencia", "observacao"}``, com a
    numeração automática 1, 2, 3... na ORDEM DE INSERÇÃO — nunca reordenada pela
    data. Linhas totalmente vazias são ignoradas em silêncio; linhas incompletas
    geram aviso localizado e não entram na linha do tempo.

    ``pendencias`` são os avisos de linha DESCARTADA — entrada materialmente
    preenchida que não pôde ser considerada. Entrada incompleta não é o mesmo
    que dado inexistente: enquanto houver pendência a história do contrato está
    sabidamente incompleta e a página não conclui a análise (fail-closed).

    ``valor`` ausente é ``None`` (ausência não é zero): só a prorrogação e o
    evento "Outro" podem seguir sem valor, herdando o valor contratual vigente.
    """
    eventos = []
    avisos = []
    pendencias = []
    vazios = {"", "nan", "none", "<na>", "nat"}

    def descartar(mensagem):
        """Linha materialmente preenchida que não entra na linha do tempo."""
        avisos.append(mensagem)
        pendencias.append(mensagem)

    for posicao, registro in enumerate(registros, start=1):
        tipo_bruto = registro.get(COLUNA_EVENTO_TIPO)
        tipo = "" if tipo_bruto is None else str(tipo_bruto).strip()
        if tipo.lower() in vazios:
            tipo = ""

        valor_bruto = registro.get(COLUNA_EVENTO_VALOR)
        texto_valor = "" if valor_bruto is None else str(valor_bruto).strip()
        tem_texto_valor = texto_valor.lower() not in vazios

        data = parse_data_br(registro.get(COLUNA_EVENTO_DATA))
        vigencia = parse_data_br(registro.get(COLUNA_EVENTO_VIGENCIA))

        observacao_bruta = registro.get(COLUNA_EVENTO_OBSERVACAO)
        observacao = "" if observacao_bruta is None else str(observacao_bruta).strip()
        if observacao.lower() in vazios:
            observacao = ""

        if not tipo and not tem_texto_valor and data is None and vigencia is None and not observacao:
            continue  # linha vazia

        rotulo = f"Linha {posicao} das alterações"

        if not tipo:
            descartar(f"{rotulo}: selecione o tipo do evento.")
            continue
        if tipo not in TIPOS_EVENTO:
            descartar(
                f'{rotulo}: o tipo "{tipo}" não é reconhecido. '
                f"Use um destes: {', '.join(TIPOS_EVENTO)}."
            )
            continue

        valor = None
        if tem_texto_valor:
            valor = parse_moeda_br(valor_bruto)
            if valor is None:
                descartar(
                    f'{rotulo} ({tipo}): o valor "{texto_valor}" não pôde ser interpretado. '
                    "Use o formato R$ 1.000.000,00."
                )
                continue
            if valor < 0:
                descartar(f"{rotulo} ({tipo}): o valor total do contrato não pode ser negativo.")
                continue
            valor = arredondar_financeiro(valor)

        if tipo in TIPOS_COM_VALOR_OBRIGATORIO and valor is None:
            descartar(f"{rotulo} ({tipo}): informe o valor total do contrato após este evento.")
            continue
        if tipo == TIPO_PRORROGACAO and vigencia is None:
            descartar(
                f"{rotulo} ({tipo}): informe o novo término da vigência para registrar a prorrogação."
            )
            continue
        if tipo == TIPO_OUTRO and valor is None and vigencia is None and not observacao:
            descartar(
                f"{rotulo} ({tipo}): informe o valor total do contrato, o novo término da vigência "
                "ou uma observação."
            )
            continue

        eventos.append(
            {
                "numero": len(eventos) + 1,
                "tipo": tipo,
                "data": data,
                "valor": valor,
                "vigencia": vigencia,
                "observacao": observacao,
            }
        )

    return eventos, avisos, pendencias


def montar_linha_do_tempo(valor_original, percentual, fim_vigencia_original, eventos):
    """Encadeia os eventos sobre o marco original e devolve a evolução completa.

    Para cada evento, na ordem de inserção:

        VALOR ANTERIOR -> VALOR APÓS O EVENTO -> VARIAÇÃO
        -> GARANTIA EXIGIDA APÓS O EVENTO -> VIGÊNCIA APÓS O EVENTO
        -> VALIDADE MÍNIMA DA GARANTIA

    Evento sem valor informado (prorrogação pura, ou "Outro" meramente
    informativo) herda o valor contratual vigente; evento sem nova vigência
    herda a vigência vigente. O percentual é sempre o informado na situação
    original — esta ferramenta não trata percentual variável no meio da vigência.

    A cadeia é recalculada por inteiro a cada chamada: editar ou excluir um
    evento anterior reflete automaticamente em todos os posteriores.
    """
    pct = parse_percentual(percentual)
    if pct is None:
        pct = PERCENTUAL_GARANTIA_PADRAO

    base = parse_moeda_br(valor_original)
    valor_corrente = arredondar_financeiro(Decimal("0") if base is None else base)
    vigencia_corrente = parse_data_br(fim_vigencia_original)
    garantia_corrente = calcular_garantia_necessaria(valor_corrente, pct)

    etapas = []
    for evento in eventos:
        valor_anterior = valor_corrente
        garantia_anterior = garantia_corrente
        vigencia_anterior = vigencia_corrente

        valor_evento = evento.get("valor")
        valor_corrente = valor_anterior if valor_evento is None else arredondar_financeiro(valor_evento)

        vigencia_evento = parse_data_br(evento.get("vigencia"))
        vigencia_corrente = vigencia_anterior if vigencia_evento is None else vigencia_evento

        garantia_corrente = calcular_garantia_necessaria(valor_corrente, pct)

        etapas.append(
            {
                "numero": evento.get("numero", len(etapas) + 1),
                "tipo": evento.get("tipo", ""),
                "data": parse_data_br(evento.get("data")),
                "observacao": evento.get("observacao", ""),
                "valor_anterior": valor_anterior,
                "valor": valor_corrente,
                "variacao": arredondar_financeiro(valor_corrente - valor_anterior),
                "garantia_anterior": garantia_anterior,
                "garantia_exigida": garantia_corrente,
                "variacao_garantia": arredondar_financeiro(garantia_corrente - garantia_anterior),
                "vigencia_anterior": vigencia_anterior,
                "vigencia": vigencia_corrente,
                "validade_minima": calcular_validade_minima(vigencia_corrente),
                "valor_informado": valor_evento is not None,
                "vigencia_informada": vigencia_evento is not None,
            }
        )

    return etapas


def calcular_situacao_atual(valor_original, percentual, fim_vigencia_original, eventos=()):
    """Deriva a situação atual do contrato do marco original + eventos posteriores.

    O usuário nunca redigita a situação atual: ela é o resultado da história
    informada nesta própria página. Nenhum VTA, Valor Global ou upload participa.
    """
    valor_base = parse_moeda_br(valor_original)
    if valor_base is None or valor_base <= 0:
        raise ValueError("O valor original do contrato deve ser maior que zero.")
    pct = parse_percentual(percentual)
    if pct is None:
        raise ValueError("O percentual da garantia deve estar entre 0 e 100.")
    vigencia_original = parse_data_br(fim_vigencia_original)
    if vigencia_original is None:
        raise ValueError("Informe o término da vigência original do contrato.")

    valor_base = arredondar_financeiro(valor_base)
    linha_do_tempo = montar_linha_do_tempo(valor_base, pct, vigencia_original, list(eventos))

    if linha_do_tempo:
        valor_atual = linha_do_tempo[-1]["valor"]
        vigencia_atual = linha_do_tempo[-1]["vigencia"]
    else:
        valor_atual = valor_base
        vigencia_atual = vigencia_original

    return {
        "valor_original": valor_base,
        "percentual": pct,
        "garantia_original": calcular_garantia_necessaria(valor_base, pct),
        "vigencia_original": vigencia_original,
        "validade_minima_original": calcular_validade_minima(vigencia_original),
        "linha_do_tempo": linha_do_tempo,
        "quantidade_eventos": len(linha_do_tempo),
        "valor_atual": valor_atual,
        "variacao_acumulada": arredondar_financeiro(valor_atual - valor_base),
        "garantia_exigida": calcular_garantia_necessaria(valor_atual, pct),
        "vigencia_atual": vigencia_atual,
        "validade_minima": calcular_validade_minima(vigencia_atual),
    }


# ============================================================
# Garantia atualmente apresentada (entrada manual)
# ============================================================

def normalizar_garantias(registros):
    """Normaliza as linhas da grade da garantia atualmente apresentada.

    ``registros``: iterável de dicts com ``COLUNA_REFERENCIA`` (opcional),
    ``COLUNA_VALOR`` e ``COLUNA_VALIDADE``. Retorna
    ``(linhas, avisos, pendencias)``, onde cada linha é
    ``{"referencia": str, "valor": Decimal, "validade": date|None}``.

    Linhas totalmente vazias são ignoradas em silêncio. A validade ausente não
    descarta a linha nem interrompe o cálculo financeiro (o valor continua
    compondo a cobertura), mas é tratada como prazo NÃO comprovado na avaliação
    temporal — fail-closed; por isso ela avisa mas NÃO gera pendência.

    ``pendencias`` são os avisos de linha DESCARTADA — garantia materialmente
    preenchida cujo valor não pôde ser considerado. Entrada incompleta não é o
    mesmo que dado inexistente: a conclusão não pode fingir que a linha nunca
    existiu.
    """
    linhas = []
    avisos = []
    pendencias = []
    vazios = {"", "nan", "none", "<na>", "nat"}

    def descartar(mensagem):
        """Linha materialmente preenchida que não compõe a cobertura."""
        avisos.append(mensagem)
        pendencias.append(mensagem)
    for idx, reg in enumerate(registros, start=1):
        referencia_bruta = reg.get(COLUNA_REFERENCIA)
        referencia = "" if referencia_bruta is None else str(referencia_bruta).strip()
        if referencia.lower() in vazios:
            referencia = ""

        valor_bruto = reg.get(COLUNA_VALOR)
        texto_valor = "" if valor_bruto is None else str(valor_bruto).strip()
        tem_texto_valor = texto_valor.lower() not in vazios
        validade = parse_data_br(reg.get(COLUNA_VALIDADE))

        if not referencia and not tem_texto_valor and validade is None:
            continue  # linha vazia

        rotulo = referencia or f"Garantia da linha {idx}"

        if not tem_texto_valor:
            descartar(f"{rotulo}: informe o valor garantido.")
            continue

        valor = parse_moeda_br(valor_bruto)
        if valor is None:
            descartar(
                f'{rotulo}: o valor "{texto_valor}" não pôde ser interpretado. '
                "Use o formato R$ 50.000,00."
            )
            continue
        if valor < 0:
            descartar(f"{rotulo}: o valor garantido não pode ser negativo.")
            continue

        if validade is None:
            avisos.append(f"{rotulo}: validade não informada — o prazo não pôde ser verificado.")

        linhas.append(
            {
                "referencia": referencia,
                "valor": arredondar_financeiro(valor),
                "validade": validade,
            }
        )

    return linhas, avisos, pendencias


def consolidar_garantias(linhas):
    """Aplica a regra do endosso: cada cadeia de garantia entra uma única vez.

    Linhas que compartilham a MESMA referência (comparação sem diferenciar
    maiúsculas/minúsculas nem espaços nas pontas) descrevem a mesma garantia em
    momentos diferentes — original e endossos sucessivos. Prevalece a ÚLTIMA
    linha informada, que traz o valor total vigente após o último endosso; as
    anteriores são descartadas para não duplicar a cobertura.

    Linhas sem referência são sempre tratadas como garantias independentes.
    Retorna ``(consolidadas, avisos)`` preservando a ordem de entrada.
    """
    consolidadas = []
    avisos = []
    posicao_por_chave = {}
    for linha in linhas:
        chave = linha["referencia"].strip().casefold()
        if chave and chave in posicao_por_chave:
            posicao = posicao_por_chave[chave]
            anterior = consolidadas[posicao]
            avisos.append(
                f"{linha['referencia']}: informada mais de uma vez. Considerado apenas o valor "
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
    """Garantia exigida = ARREDONDAR(valor total do contrato x percentual; 2)."""
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


def classificar_financeiro(cobertura_atual, garantia_necessaria):
    """Três estados financeiros — cobertura superior NÃO é pendência.

    Cobertura acima da exigida não determina redução nem devolução: apenas
    registra que não há complemento a exigir. Eventual adequação depende da
    análise contratual, que este módulo não faz.
    """
    cobertura = arredondar_financeiro(cobertura_atual)
    necessaria = arredondar_financeiro(garantia_necessaria)
    if cobertura < necessaria:
        return FINANCEIRO_COMPLEMENTAR
    if cobertura == necessaria:
        return FINANCEIRO_SUFICIENTE
    return FINANCEIRO_SUPERIOR


def classificar_temporal(garantias_detalhadas):
    """Três estados temporais, independentes do resultado financeiro.

    Sem garantia apresentada não há validade a comprovar: o prazo é reportado
    como NÃO INFORMADO (fail-closed), e a pendência de dinheiro fica a cargo da
    dimensão financeira.
    """
    if not garantias_detalhadas:
        return TEMPORAL_NAO_INFORMADA
    if any(parse_data_br(item.get("validade")) is None for item in garantias_detalhadas):
        return TEMPORAL_NAO_INFORMADA
    if all(item["validade_suficiente"] for item in garantias_detalhadas):
        return TEMPORAL_SUFICIENTE
    return TEMPORAL_INSUFICIENTE


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
    """Confronta a SITUAÇÃO ATUAL apurada com a GARANTIA ATUALMENTE APRESENTADA.

    ``valor_total_contrato`` e ``data_fim_vigencia`` são os derivados da linha do
    tempo (ver ``calcular_situacao_atual``), nunca redigitados. ``garantias`` são
    as linhas já normalizadas por ``normalizar_garantias``; a consolidação dos
    endossos é aplicada aqui.

    A avaliação temporal considera insuficiente qualquer garantia que componha a
    cobertura e vença antes da validade mínima (ou cuja validade não tenha sido
    informada). Sem nenhuma garantia cadastrada não há prazo a regularizar: a
    pendência é apenas de valor.

    A validade de cada linha é reconvertida por ``parse_data_br`` imediatamente
    antes da comparação: é essa normalização — e não um ``is not None`` — que
    garante que nenhum ``NaT`` vindo do editor chegue a um operador relacional.
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
        validade = parse_data_br(garantia.get("validade"))
        suficiente = validade is not None and validade >= validade_minima
        detalhadas.append({**garantia, "validade": validade, "validade_suficiente": suficiente})
    validade_suficiente = all(item["validade_suficiente"] for item in detalhadas)

    return {
        "valor_total_contrato": valor_total,
        "percentual": pct,
        "garantia_necessaria": garantia_necessaria,
        "cobertura_atual": cobertura_atual,
        "complemento": complemento,
        "valor_suficiente": valor_suficiente,
        "situacao_financeira": classificar_financeiro(cobertura_atual, garantia_necessaria),
        "data_fim_vigencia": parse_data_br(data_fim_vigencia),
        "validade_minima": validade_minima,
        "validade_suficiente": validade_suficiente,
        "situacao_temporal": classificar_temporal(detalhadas),
        "garantias": detalhadas,
        "quantidade_garantias": len(detalhadas),
        "avisos_consolidacao": avisos,
        "diagnostico": diagnosticar(valor_suficiente, validade_suficiente),
    }


# ============================================================
# Texto para comunicação à contratada
# ============================================================

def _linha_referencia(numero_contrato="", contratada=""):
    """Cabeçalho "Ref.:" com o que tiver sido informado; vazio se nada houver."""
    partes = []
    numero = str(numero_contrato or "").strip()
    nome = str(contratada or "").strip()
    if numero:
        partes.append(f"Contrato nº {numero}")
    if nome:
        partes.append(nome)
    if not partes:
        return []
    return ["Ref.: " + " — ".join(partes), ""]


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


def gerar_texto_comunicacao(analise, numero_contrato="", contratada=""):
    """Texto pronto para copiar e colar em e-mail, conforme o diagnóstico.

    Quatro situações: sem garantia apresentada, garantia financeiramente
    insuficiente, garantia suficiente em valor mas com validade curta, e
    garantia suficiente nas duas dimensões. Havendo apenas prazo a regularizar,
    nenhuma complementação financeira é mencionada; havendo cobertura superior à
    exigida, nada é solicitado e nenhuma devolução é determinada.

    O percentual é sempre o efetivamente informado pelo usuário — jamais um
    "5%" fixo.
    """
    diagnostico = analise["diagnostico"]
    referencia = _linha_referencia(numero_contrato, contratada)

    if diagnostico == DIAGNOSTICO_REGULAR:
        return "\n".join([*referencia, SEM_NECESSIDADE_DE_ATUALIZACAO])

    data_minima = formatar_data_br(analise["validade_minima"])
    fim_vigencia = formatar_data_br(analise["data_fim_vigencia"])
    fecho = [
        "Gentileza encaminhar o respectivo endosso/comprovante após a regularização.",
        "",
        "Atenciosamente,",
    ]

    if diagnostico == DIAGNOSTICO_VALIDADE:
        return "\n".join(
            [
                *referencia,
                "Prezados,",
                "",
                f"Considerando a vigência atual do contrato, com término em {fim_vigencia}, "
                "verificamos a necessidade de atualização da validade da garantia contratual.",
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
            f"correspondente a {DIAS_VALIDADE_MINIMA} dias após o término da vigência contratual, "
            f"encerrada em {fim_vigencia}."
        )
    else:
        solicitacao = (
            f"Solicitamos a atualização da garantia, nos termos da {CLAUSULA_GARANTIA}, observando "
            f"também a validade mínima de {data_minima}, correspondente a {DIAS_VALIDADE_MINIMA} "
            f"dias após o término da vigência contratual, encerrada em {fim_vigencia}."
        )

    return "\n".join(
        [
            *referencia,
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
