import re
import unicodedata

# --- Quantidade de ciclos efetivamente considerados na apuracao -------------
# Regra unica de contagem e de redacao, compartilhada por todos os documentos.
# A frase do documento NAO pode contar linhas de quadro: um ciclo listado como
# "Fora da apuracao" aparece no quadro para rastreabilidade, mas nao foi
# considerado na analise.

_CHAVES_COMPUTAR = ("computar", "COMPUTAR", "computar_nesta_apuracao",
                    "COMPUTAR_NESTA_APURACAO")
_CHAVES_TRATAMENTO = ("Tratamento financeiro do ciclo", "Tratamento financeiro",
                      "tratamento_financeiro")
_CHAVES_CLASSIFICACAO = ("Situação", "Situacao", "situacao",
                         "Classificação", "Classificacao", "classificacao")

# Classificacoes que excluem o ciclo da apuracao de forma explicita.
_CLASSIFICACOES_EXCLUIDAS = ("fora da apuracao", "nao computado",
                             "nao aplicavel", "nao se aplica")


def _texto_normalizado(valor):
    texto = "" if valor is None else str(valor)
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return " ".join(texto.lower().split())


_AUSENTE = object()


def _valor_por_chave(ciclo, chaves):
    """Devolve o valor da primeira chave PRESENTE; _AUSENTE se nenhuma existir.

    A presenca da chave importa e e distinta do valor vazio: um ciclo que traz
    COMPUTAR em branco nao esta autorizado ao calculo e NAO deve cair na regra
    seguinte — vazio nunca vira "considerado".
    """
    for chave in chaves:
        if chave in ciclo:
            return ciclo.get(chave)
    return _AUSENTE


def ciclo_computado(ciclo):
    """True quando o ciclo integra efetivamente a apuracao.

    Ordem de precedencia (fonte canonica primeiro):
      1. marcacao COMPUTAR do arquivo (parametros!COMPUTAR_NESTA_APURACAO);
      2. tratamento financeiro consolidado ("Apurar");
      3. classificacao/situacao exibida no quadro, que so exclui quando diz
         explicitamente que o ciclo esta fora.
    """
    if not isinstance(ciclo, dict):
        return False

    marca = _valor_por_chave(ciclo, _CHAVES_COMPUTAR)
    if marca is not _AUSENTE:
        return _texto_normalizado(marca) == "sim"

    tratamento = _valor_por_chave(ciclo, _CHAVES_TRATAMENTO)
    if tratamento is not _AUSENTE:
        return "apurar" in _texto_normalizado(tratamento)

    classificacao = _valor_por_chave(ciclo, _CHAVES_CLASSIFICACAO)
    if classificacao is not _AUSENTE:
        texto = _texto_normalizado(classificacao)
        return not any(marcador in texto for marcador in _CLASSIFICACOES_EXCLUIDAS)

    return True


def contar_ciclos_computados(ciclos):
    """Quantidade de ciclos efetivamente considerados na apuracao."""
    return sum(1 for ciclo in ciclos or [] if ciclo_computado(ciclo))


def expressao_quantidade_ciclos(quantidade):
    """'1 ciclo' / 'N ciclos'. None quando nenhum ciclo foi computado.

    None sinaliza ao chamador que a frase de contagem nao deve ser afirmada:
    sem ciclo computado nao houve analise de ciclos a declarar.
    """
    try:
        total = int(quantidade)
    except (TypeError, ValueError):
        return None
    if total <= 0:
        return None
    return "1 ciclo" if total == 1 else f"{total} ciclos"


FRASE_SEM_CICLOS_COMPUTADOS = "Não foram identificados ciclos computados na apuração."


# Etapa 26H: padrao canonico de NOVO ITEM (cadastro fisico via aditivo).
# Etapa 26H.2: nomenclatura funcional aprovada = "N" + EXATAMENTE 3 algarismos
# (N001..N999). N1/N01/N0001/N1000 NAO sao reconhecidos automaticamente como
# novos itens; codigos existentes do contrato nunca casam com este padrao.
_PADRAO_NOVO_ITEM = re.compile(r"^[Nn]\d{3}$")


def eh_novo_item(item):
    """True somente para identificadores canonicos de novo item (Nxxx)."""
    if item is None:
        return False
    return bool(_PADRAO_NOVO_ITEM.match(str(item).strip()))


def qtd_base_efetiva(item, qtd_base):
    """QTD_BASE_ORIGINAL efetiva da regra 26H de base zero automatica.

    Novo item (Nxxx) com base vazia -> 0.0 (o fiscal nao digita o zero).
    Item normal com base vazia -> None (vazio NUNCA vira zero genericamente).
    Base informada (qualquer item) -> float(base); inconsistencias (Nxxx com
    base != 0) NAO sao corrigidas aqui — sao acusadas pelos CHECKs.
    """
    if qtd_base is None or (isinstance(qtd_base, str) and not qtd_base.strip()):
        return 0.0 if eh_novo_item(item) else None
    try:
        return float(qtd_base)
    except (TypeError, ValueError):
        return None


def _parse_moeda_br(valor):
    """Converte texto monetário brasileiro em float, preservando campos vazios como 0."""
    if valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip()
    if not texto:
        return 0.0
    texto = texto.replace("R$", "").replace("\xa0", "").replace(" ", "")
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    else:
        if texto.count(".") > 1:
            partes = texto.split(".")
            texto = "".join(partes[:-1]) + "." + partes[-1]
    try:
        return float(texto)
    except Exception:
        return 0.0


def _formatar_moeda_br(valor):
    try:
        valor = round(float(valor), 2)
    except Exception:
        valor = 0.0
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _formatar_moeda_br_md(valor):
    """Formata moeda escapando o $ para exibição correta no Markdown do Streamlit."""
    return _formatar_moeda_br(valor).replace("$", "\\$")


def _data_para_datetime(valor):
    if valor is None or valor == "":
        return None
    try:
        import pandas as pd
        return pd.to_datetime(valor, dayfirst=True).to_pydatetime()
    except Exception:
        return None


def _formatar_data(valor):
    """Formata datas para DD/MM/AAAA sem quebrar quando o valor estiver vazio."""
    if valor is None or valor == "":
        return ""
    try:
        import pandas as pd
        return pd.to_datetime(valor, dayfirst=True).strftime("%d/%m/%Y")
    except Exception:
        return str(valor)


def _competencias_mensais(data_inicio, data_fim):
    """Gera competências mensais inclusivas no formato MM/AAAA."""
    from dateutil.relativedelta import relativedelta

    inicio = _data_para_datetime(data_inicio)
    fim = _data_para_datetime(data_fim)
    if inicio is None or fim is None:
        return []

    atual = inicio.replace(day=1)
    limite = fim.replace(day=1)
    competencias = []
    while atual <= limite:
        competencias.append(atual.strftime("%m/%Y"))
        atual = atual + relativedelta(months=1)
    return competencias


def _percentual_formatado(valor):
    try:
        return f"{float(valor) * 100:,.2f}%".replace(".", ",")
    except Exception:
        return ""
