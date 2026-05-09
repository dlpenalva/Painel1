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
