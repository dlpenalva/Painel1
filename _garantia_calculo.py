"""Motor da Calculadora de Garantia Contratual — uma única linha do tempo.

A ferramenta conta a história do contrato em ordem cronológica, e a garantia
caminha junto com ela:

    SITUAÇÃO ORIGINAL
        -> ALTERAÇÕES POSTERIORES + GARANTIA APÓS CADA EVENTO
        -> SITUAÇÃO ATUAL
        -> RESULTADO

Não há quadro de garantias separado. Cada linha de alteração registra o que
aconteceu com o contrato E qual passou a ser a situação da garantia depois
daquele evento.

    GARANTIA EXIGIDA = ARREDONDAR(VALOR TOTAL DO CONTRATO x PERCENTUAL; 2)
    VALIDADE MÍNIMA  = TÉRMINO DA VIGÊNCIA + 90 DIAS CORRIDOS
    COMPLEMENTO      = max(GARANTIA EXIGIDA - GARANTIA APRESENTADA; 0)

FOTOGRAFIA, NÃO SOMA. A garantia informada numa linha é a fotografia da
garantia VIGENTE após aquele evento — nunca uma parcela a somar às anteriores.
Assinatura 50.000, depois 55.000, depois 55.000 significa garantia vigente de
55.000, jamais 160.000. Esta ferramenta não trabalha com cadeia de
apólice/endosso por identificação.

HERANÇA. Evento sem garantia e sem validade informadas significa "nenhuma nova
situação de garantia neste evento": a fotografia anterior permanece vigente.
Informar um novo valor de garantia cria uma nova fotografia, com a validade que
vier junto — e apenas ela; a validade anterior nunca é herdada por uma garantia
nova explicitamente informada.

Cada evento informa o VALOR TOTAL DO CONTRATO APÓS O EVENTO — nunca o acréscimo
isolado — e o motor deriva a variação contra a situação anterior. A prorrogação
pode não alterar o valor: sem valor informado, herda o valor contratual vigente.

A suficiência financeira e a suficiência temporal são verificações INDEPENDENTES:
uma garantia pode ter valor bastante e prazo curto, e vice-versa.

Nenhuma função deste módulo lê o VTA, a Coleta, os RESULTADOS ou qualquer outra
fonte do sistema: todos os dados chegam por parâmetro, informados à mão. A
ferramenta também não pede identificação de contrato, contratada ou apólice.

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

# Colunas da grade única de alterações posteriores à assinatura. O contrato e a
# garantia andam lado a lado na mesma linha.
COLUNA_EVENTO_TIPO = "Tipo"
COLUNA_EVENTO_DATA = "Data"
COLUNA_EVENTO_VALOR = "Valor do contrato após o evento"
COLUNA_EVENTO_VIGENCIA = "Novo término da vigência"
COLUNA_EVENTO_GARANTIA = "Garantia após o evento"
COLUNA_EVENTO_VALIDADE = "Validade da garantia"

# Tipos de evento contratual.
TIPO_REAJUSTE = "Reajuste"
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
        .replace(" ", "")   # espaço não separável (copiar/colar de Excel e web)
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

    Ausência não é zero: um evento sem garantia informada não pode ser exibido
    como R$ 0,00.
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
# Alterações posteriores à assinatura (contrato + garantia)
# ============================================================

_VAZIOS = {"", "nan", "none", "<na>", "nat"}


def _texto_preenchido(bruto):
    """A célula está materialmente preenchida? Devolve ``(tem, texto)``."""
    texto = "" if bruto is None else str(bruto).strip()
    return (texto.lower() not in _VAZIOS), texto


def normalizar_eventos(registros):
    """Normaliza as linhas da grade única de alterações posteriores.

    ``registros``: iterável de dicts com ``COLUNA_EVENTO_TIPO``,
    ``COLUNA_EVENTO_DATA``, ``COLUNA_EVENTO_VALOR``, ``COLUNA_EVENTO_VIGENCIA``,
    ``COLUNA_EVENTO_GARANTIA`` e ``COLUNA_EVENTO_VALIDADE``.

    Retorna ``(eventos, avisos, pendencias)``. Cada evento é
    ``{"numero", "tipo", "data", "valor", "vigencia", "garantia",
    "validade_garantia", "garantia_informada"}``, com a numeração automática
    1, 2, 3... na ORDEM DE INSERÇÃO — nunca reordenada pela data.

    Linhas totalmente vazias são ignoradas em silêncio. ``pendencias`` são os
    avisos de linha DESCARTADA — entrada materialmente preenchida que não pôde
    ser considerada. Entrada incompleta não é o mesmo que dado inexistente:
    enquanto houver pendência a história está sabidamente incompleta e a página
    não conclui a análise (fail-closed).

    ``garantia`` ausente é ``None`` (ausência não é zero): a fotografia anterior
    da garantia permanece vigente — ver ``montar_linha_do_tempo``.
    """
    eventos = []
    avisos = []
    pendencias = []

    def descartar(mensagem):
        """Linha materialmente preenchida que não entra na linha do tempo."""
        avisos.append(mensagem)
        pendencias.append(mensagem)

    for posicao, registro in enumerate(registros, start=1):
        tem_tipo, tipo = _texto_preenchido(registro.get(COLUNA_EVENTO_TIPO))
        if not tem_tipo:
            tipo = ""

        tem_valor, texto_valor = _texto_preenchido(registro.get(COLUNA_EVENTO_VALOR))
        tem_garantia, texto_garantia = _texto_preenchido(registro.get(COLUNA_EVENTO_GARANTIA))

        data = parse_data_br(registro.get(COLUNA_EVENTO_DATA))
        vigencia = parse_data_br(registro.get(COLUNA_EVENTO_VIGENCIA))
        validade = parse_data_br(registro.get(COLUNA_EVENTO_VALIDADE))

        if (
            not tipo and not tem_valor and not tem_garantia
            and data is None and vigencia is None and validade is None
        ):
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
        if tem_valor:
            valor = parse_moeda_br(registro.get(COLUNA_EVENTO_VALOR))
            if valor is None:
                descartar(
                    f'{rotulo} ({tipo}): o valor "{texto_valor}" não pôde ser interpretado. '
                    "Use o formato R$ 1.000.000,00."
                )
                continue
            if valor < 0:
                descartar(f"{rotulo} ({tipo}): o valor do contrato não pode ser negativo.")
                continue
            valor = arredondar_financeiro(valor)

        garantia = None
        if tem_garantia:
            garantia = parse_moeda_br(registro.get(COLUNA_EVENTO_GARANTIA))
            if garantia is None:
                descartar(
                    f'{rotulo} ({tipo}): a garantia "{texto_garantia}" não pôde ser interpretada. '
                    "Use o formato R$ 55.000,00."
                )
                continue
            if garantia < 0:
                descartar(f"{rotulo} ({tipo}): a garantia não pode ser negativa.")
                continue
            garantia = arredondar_financeiro(garantia)

        if tipo in TIPOS_COM_VALOR_OBRIGATORIO and valor is None:
            descartar(f"{rotulo} ({tipo}): informe o valor do contrato após este evento.")
            continue
        if tipo == TIPO_PRORROGACAO and vigencia is None:
            descartar(
                f"{rotulo} ({tipo}): informe o novo término da vigência para registrar a prorrogação."
            )
            continue
        if tipo == TIPO_OUTRO and valor is None and vigencia is None and garantia is None:
            descartar(
                f"{rotulo} ({tipo}): informe o valor do contrato, o novo término da vigência "
                "ou a garantia após o evento."
            )
            continue
        if validade is not None and garantia is None:
            # Validade sozinha não descreve uma fotografia: não dá para saber a
            # que garantia esse prazo pertence.
            descartar(
                f"{rotulo} ({tipo}): informe também o valor da garantia vigente após este evento."
            )
            continue

        eventos.append(
            {
                "numero": len(eventos) + 1,
                "tipo": tipo,
                "data": data,
                "valor": valor,
                "vigencia": vigencia,
                "garantia": garantia,
                "validade_garantia": validade,
                "garantia_informada": garantia is not None,
            }
        )

    return eventos, avisos, pendencias


def montar_linha_do_tempo(valor_original, percentual, fim_vigencia_original, eventos,
                          garantia_original=None, validade_garantia_original=None):
    """Encadeia os eventos sobre o marco original e devolve a evolução completa.

    Para cada evento, na ordem de inserção:

        VALOR ANTERIOR -> VALOR APÓS O EVENTO -> VARIAÇÃO
        -> GARANTIA EXIGIDA APÓS O EVENTO -> VIGÊNCIA -> VALIDADE MÍNIMA
        -> FOTOGRAFIA DA GARANTIA APRESENTADA APÓS O EVENTO

    Evento sem valor informado (prorrogação pura, ou "Outro" meramente temporal)
    herda o valor contratual vigente; evento sem nova vigência herda a vigência
    vigente.

    A garantia apresentada segue a regra da FOTOGRAFIA: sem garantia informada, a
    fotografia anterior permanece vigente; com garantia informada, ela substitui
    a anterior, acompanhada apenas da validade informada na própria linha. Nada
    se soma.

    Devolve ``(etapas, garantia_apresentada, validade_apresentada)`` — a última
    fotografia é a que vale hoje. O percentual é sempre o informado na situação
    original. A cadeia é recalculada por inteiro a cada chamada: editar ou
    excluir um evento anterior reflete automaticamente em todos os posteriores.
    """
    pct = parse_percentual(percentual)
    if pct is None:
        pct = PERCENTUAL_GARANTIA_PADRAO

    base = parse_moeda_br(valor_original)
    valor_corrente = arredondar_financeiro(Decimal("0") if base is None else base)
    vigencia_corrente = parse_data_br(fim_vigencia_original)
    exigida_corrente = calcular_garantia_necessaria(valor_corrente, pct)

    apresentada = parse_moeda_br(garantia_original)
    apresentada = None if apresentada is None else arredondar_financeiro(apresentada)
    # Sem garantia não há validade a carregar.
    validade_apresentada = parse_data_br(validade_garantia_original) if apresentada is not None else None

    etapas = []
    for evento in eventos:
        valor_anterior = valor_corrente
        exigida_anterior = exigida_corrente
        vigencia_anterior = vigencia_corrente
        apresentada_anterior = apresentada

        valor_evento = evento.get("valor")
        valor_corrente = valor_anterior if valor_evento is None else arredondar_financeiro(valor_evento)

        vigencia_evento = parse_data_br(evento.get("vigencia"))
        vigencia_corrente = vigencia_anterior if vigencia_evento is None else vigencia_evento

        exigida_corrente = calcular_garantia_necessaria(valor_corrente, pct)

        garantia_evento = evento.get("garantia")
        if garantia_evento is not None:
            # Nova fotografia: substitui a anterior e leva só a validade da
            # própria linha. Validade antiga JAMAIS é herdada por garantia nova.
            apresentada = arredondar_financeiro(garantia_evento)
            validade_apresentada = parse_data_br(evento.get("validade_garantia"))

        etapas.append(
            {
                "numero": evento.get("numero", len(etapas) + 1),
                "tipo": evento.get("tipo", ""),
                "data": parse_data_br(evento.get("data")),
                "valor_anterior": valor_anterior,
                "valor": valor_corrente,
                "variacao": arredondar_financeiro(valor_corrente - valor_anterior),
                "garantia_exigida_anterior": exigida_anterior,
                "garantia_exigida": exigida_corrente,
                "vigencia_anterior": vigencia_anterior,
                "vigencia": vigencia_corrente,
                "validade_minima": calcular_validade_minima(vigencia_corrente),
                "garantia_apresentada_anterior": apresentada_anterior,
                "garantia_apresentada": apresentada,
                "validade_apresentada": validade_apresentada,
                "garantia_informada": garantia_evento is not None,
                "valor_informado": valor_evento is not None,
                "vigencia_informada": vigencia_evento is not None,
            }
        )

    return etapas, apresentada, validade_apresentada


def calcular_situacao_atual(valor_original, percentual, fim_vigencia_original, eventos=(),
                            garantia_original=None, validade_garantia_original=None):
    """Deriva a situação atual do contrato E da garantia a partir da linha do tempo.

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
    garantia_base = parse_moeda_br(garantia_original)
    garantia_base = None if garantia_base is None else arredondar_financeiro(garantia_base)
    validade_base = parse_data_br(validade_garantia_original) if garantia_base is not None else None

    linha_do_tempo, apresentada, validade_apresentada = montar_linha_do_tempo(
        valor_base, pct, vigencia_original, list(eventos), garantia_base, validade_base
    )

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
        "garantia_apresentada_original": garantia_base,
        "validade_apresentada_original": validade_base,
        "vigencia_original": vigencia_original,
        "validade_minima_original": calcular_validade_minima(vigencia_original),
        "linha_do_tempo": linha_do_tempo,
        "quantidade_eventos": len(linha_do_tempo),
        "valor_atual": valor_atual,
        "variacao_acumulada": arredondar_financeiro(valor_atual - valor_base),
        "garantia_exigida": calcular_garantia_necessaria(valor_atual, pct),
        "vigencia_atual": vigencia_atual,
        "validade_minima": calcular_validade_minima(vigencia_atual),
        "garantia_apresentada": apresentada,
        "validade_apresentada": validade_apresentada,
    }


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


def diagnosticar(valor_suficiente, validade_suficiente):
    """Combina as duas verificações independentes nos quatro resultados possíveis."""
    if valor_suficiente and validade_suficiente:
        return DIAGNOSTICO_REGULAR
    if valor_suficiente:
        return DIAGNOSTICO_VALIDADE
    if validade_suficiente:
        return DIAGNOSTICO_VALOR
    return DIAGNOSTICO_VALOR_E_VALIDADE


def analisar_garantia(valor_total_contrato, percentual, data_fim_vigencia,
                      garantia_apresentada=None, validade_apresentada=None):
    """Confronta a SITUAÇÃO ATUAL apurada com a ÚLTIMA FOTOGRAFIA da garantia.

    Todos os argumentos são derivados da linha do tempo (ver
    ``calcular_situacao_atual``) — nada é redigitado.

    ``garantia_apresentada`` ausente (``None``) significa nenhuma garantia
    apresentada, e não R$ 0,00: sem garantia não há prazo a regularizar, e a
    pendência é apenas de valor.

    A validade é reconvertida por ``parse_data_br`` imediatamente antes da
    comparação: é essa normalização — e não um ``is not None`` — que garante que
    nenhum ``NaT`` vindo do editor chegue a um operador relacional.
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
    apresentada = parse_moeda_br(garantia_apresentada)
    tem_garantia = apresentada is not None
    apresentada = arredondar_financeiro(apresentada) if tem_garantia else None
    cobertura_atual = apresentada if tem_garantia else Decimal("0.00")
    validade = parse_data_br(validade_apresentada) if tem_garantia else None

    garantia_necessaria = calcular_garantia_necessaria(valor_total, pct)
    complemento = calcular_complemento(garantia_necessaria, cobertura_atual)
    valor_suficiente = cobertura_atual >= garantia_necessaria

    if not tem_garantia:
        # Sem garantia apresentada não há prazo a regularizar: a pendência é só
        # de valor, e o prazo é reportado como não informado (fail-closed).
        validade_suficiente = True
        situacao_temporal = TEMPORAL_NAO_INFORMADA
    elif validade is None:
        validade_suficiente = False
        situacao_temporal = TEMPORAL_NAO_INFORMADA
    elif validade >= validade_minima:
        validade_suficiente = True
        situacao_temporal = TEMPORAL_SUFICIENTE
    else:
        validade_suficiente = False
        situacao_temporal = TEMPORAL_INSUFICIENTE

    return {
        "valor_total_contrato": valor_total,
        "percentual": pct,
        "garantia_necessaria": garantia_necessaria,
        "garantia_apresentada": apresentada,
        "tem_garantia": tem_garantia,
        "cobertura_atual": cobertura_atual,
        "complemento": complemento,
        "valor_suficiente": valor_suficiente,
        "situacao_financeira": classificar_financeiro(cobertura_atual, garantia_necessaria),
        "data_fim_vigencia": parse_data_br(data_fim_vigencia),
        "validade_minima": validade_minima,
        "validade_apresentada": validade,
        "validade_suficiente": validade_suficiente,
        "situacao_temporal": situacao_temporal,
        "diagnostico": diagnosticar(valor_suficiente, validade_suficiente),
    }


# ============================================================
# Texto para comunicação à contratada
# ============================================================

def _paragrafo_cobertura(analise):
    """Parágrafo da garantia atualmente apresentada."""
    complemento = moeda(analise["complemento"])
    if not analise["tem_garantia"]:
        return (
            "Não há garantia contratual atualmente apresentada, sendo necessária, portanto, a "
            f"apresentação de garantia no valor de {complemento}."
        )
    return (
        f"A garantia atualmente apresentada é de {moeda(analise['cobertura_atual'])}, sendo "
        f"necessária, portanto, a complementação no valor de {complemento}."
    )


def gerar_texto_comunicacao(analise):
    """Texto pronto para copiar e colar em e-mail, conforme o diagnóstico.

    Quatro situações: sem garantia apresentada, garantia financeiramente
    insuficiente, garantia suficiente em valor mas com validade curta, e
    garantia suficiente nas duas dimensões. Havendo apenas prazo a regularizar,
    nenhuma complementação financeira é mencionada; havendo cobertura superior à
    exigida, nada é solicitado e nenhuma devolução é determinada.

    O percentual é sempre o efetivamente informado pelo usuário — jamais um
    "5%" fixo. O texto não carrega identificação de contrato ou de contratada.
    """
    diagnostico = analise["diagnostico"]
    if diagnostico == DIAGNOSTICO_REGULAR:
        return SEM_NECESSIDADE_DE_ATUALIZACAO

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
