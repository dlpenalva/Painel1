import re
import unicodedata

import streamlit as st

try:
    from _ui_utils import render_marca_topo
except Exception:
    def render_marca_topo():
        st.markdown("### TLB · cl8us")
        st.caption("apoio à gestão de contratos")



# ============================================================
# Utilitários simples — playground isolado
# ============================================================

def _normalizar(texto: str) -> str:
    texto = "" if texto is None else str(texto).lower().strip()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(ch for ch in texto if not unicodedata.combining(ch))
    return texto


def _tem(texto: str, termos) -> bool:
    base = _normalizar(texto)
    return any(_normalizar(t) in base for t in termos)


def _extrair_datas(texto: str):
    padrao = r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b|\b\d{1,2}[/-]\d{4}\b"
    return re.findall(padrao, texto or "")


def analisar_relato(relato: str):
    datas = _extrair_datas(relato)
    flags = {
        "multiplos_ciclos": _tem(relato, ["dois ciclos", "2 ciclos", "multiplos ciclos", "múltiplos ciclos", "c1", "c2", "c3", "segundo reajuste", "terceiro reajuste", "ciclos acumulados"]),
        "preclusao": _tem(relato, ["preclus", "precluiu", "perdeu prazo", "nao pediu", "não pediu", "sem pedido", "nunca pediu", "pulou", "fora do prazo"]),
        "pedido_tardio": _tem(relato, ["atras", "atraso", "tardio", "intempest", "oficio chegou depois", "ofício chegou depois", "solicitou depois", "pedido em dezembro", "protocolo"]),
        "efeito_financeiro": _tem(relato, ["efeito financeiro", "efeitos financeiros", "retroativo", "data do pedido", "dinheiro", "cofre", "meses", "comendo meses", "valor a pagar"]),
        "historico_anterior": _tem(relato, ["formalizado", "apostila anterior", "termo de apostila", "reajuste anterior", "ja concedido", "já concedido", "ja formalizado", "já formalizado"]),
        "aditivo": _tem(relato, ["aditivo", "termo aditivo", "acrescimo", "acréscimo", "novos servicos", "novos serviços", "escopo", "alteracao contratual", "alteração contratual"]),
        "supressao": _tem(relato, ["supress", "cortando", "corte", "reduzido", "reduziu", "25% do escopo"]),
        "garantia": _tem(relato, ["garantia", "endosso", "adensamento", "reforco", "reforço", "apolice", "apólice", "caucao", "caução", "valor garantido"]),
        "valor_global": _tem(relato, ["valor global", "valor total atualizado", "valor atualizado", "remanescente", "saldo", "itens", "quantidade remanescente", "estoque", "valor do contrato", "valor contratual"]),
        "duvida_direito": _tem(relato, ["tem direito", "empresa tem direito", "deferir", "indeferir", "salvar", "posso", "devo"]),
        "assinatura_contrato": _tem(relato, ["assinatura", "assinado", "data de assinatura", "contrato foi assinado"]),
        "minuta_apostila": _tem(relato, ["minuta", "apostila", "apostilamento", "termo de apostila"]),
        "indice_reajuste": _tem(relato, ["indice", "índice", "ist", "ipca", "igp-m", "igpm", "percentual", "reajuste"]),
    }
    return flags, datas


def tema_principal(flags):
    # A ideia aqui é evitar despejar todos os assuntos. O Orienta deve focar no que parece ser a dúvida.
    if flags["garantia"] and not (flags["preclusao"] or flags["pedido_tardio"] or flags["efeito_financeiro"] or flags["multiplos_ciclos"]):
        return "garantia"
    if flags["supressao"]:
        return "supressao"
    if flags["aditivo"] and not (flags["preclusao"] or flags["pedido_tardio"] or flags["efeito_financeiro"] or flags["multiplos_ciclos"]):
        return "aditivo"
    if flags["preclusao"] or flags["pedido_tardio"] or flags["efeito_financeiro"] or flags["multiplos_ciclos"] or flags["assinatura_contrato"]:
        return "reajuste"
    if flags["valor_global"]:
        return "valores"
    return "geral"


def _bullet(texto: str):
    st.markdown(f"- {texto}")


def _texto_corrente(texto: str):
    st.markdown(f"<div class='orienta-texto'>{texto}</div>", unsafe_allow_html=True)


def _alerta_informal():
    frase = "Essa é uma ajudinha preliminar. Lembre-se de checar no contrato, beleza? Apesar de ser uma máquina, eu posso errar também, né?!"
    st.markdown(
        f"<div class='orienta-alerta'>{frase}</div>",
        unsafe_allow_html=True,
    )


# ============================================================
# Respostas por tema
# ============================================================

def resposta_garantia(flags):
    _texto_corrente(
        "Pelo que você contou, eu olharia primeiro para a garantia. Se houve contrato original, depois um aditivo e depois um reajuste, o cuidado é conferir se a garantia que a empresa já apresentou ainda cobre o valor contratual atualizado. Aqui, a dúvida principal não parece ser se o reajuste é admissível, mas qual base usar para comparar a garantia exigida com a garantia já constituída."
    )
    st.markdown("### Eu começaria por aqui")
    itens = [
        "confirme o percentual de garantia previsto no contrato;",
        "veja qual é o valor da garantia já apresentada e se houve endossos anteriores;",
        "confirme qual é o valor contratual atualizado depois do aditivo e do reajuste;",
        "use o módulo Gestão da Garantia para calcular se há reforço/endosso;",
        "se houver diferença, registre a necessidade de atualização antes de fechar a instrução.",
    ]
    for item in itens:
        _bullet(item)

    st.markdown("### Cuidado")
    _texto_corrente(
        "Eu não calcularia a garantia só sobre o valor original se o contrato já mudou de valor. Também tomaria cuidado para não somar eventos em duplicidade: o valor-base da garantia deve ser o valor contratual atualizado que você vai adotar como referência, não uma soma solta de contrato original, aditivo e reajuste."
    )

    st.markdown("### No cl8us")
    for item in [
        "use Valores se precisar confirmar o Valor Total Atualizado do Contrato;",
        "use Gestão da Garantia para comparar garantia exigida e garantia constituída;",
        "use Infos Prévias para registrar garantia atual, validade, endossos e dados do aditivo;",
        "use Aditivos: 25% só se também precisar controlar acréscimo/supressão e limite percentual.",
    ]:
        _bullet(item)


def resposta_reajuste(flags, datas):
    _texto_corrente(
        "Pelo relato, eu trataria isso como uma dúvida de reajuste e admissibilidade. O ponto central é separar a linha do tempo do contrato da parte financeira: uma coisa é o ciclo existir; outra coisa é o pedido ser admissível; outra, ainda, é definir desde quando há efeito financeiro."
    )

    if datas:
        st.markdown("### Datas que apareceram no relato")
        _bullet(", ".join(datas))
        st.caption("Use essas datas só como pista. A conferência mesmo deve ser feita com as datas reais na Calculadora.")

    st.markdown("### Eu olharia para estes pontos")
    pontos = []
    if flags["assinatura_contrato"]:
        pontos.append("não use a assinatura como marco automático; normalmente, o dado mais importante para a anualidade é a data da proposta;")
    else:
        pontos.append("confirme a data da proposta, porque ela costuma ser o marco inicial da anualidade;")
    if flags["preclusao"]:
        pontos.append("se houve ausência de pedido ou pedido fora da janela, pode haver preclusão do ciclo;")
    if flags["pedido_tardio"] or flags["efeito_financeiro"]:
        pontos.append("se o pedido foi tardio, o índice pode ser apurado, mas os efeitos financeiros podem começar só na data do pedido;")
    if flags["multiplos_ciclos"]:
        pontos.append("em C2, C3 e seguintes, confira se o marco foi empurrado pelos efeitos financeiros do ciclo anterior ou pela data em que o ciclo poderia ter sido pleiteado;")
    if flags["historico_anterior"]:
        pontos.append("se já houve reajuste formalizado, registre isso no Contexto do Contrato antes de avançar;")
    if not pontos:
        pontos.append("confirme data da proposta, data do pedido, índice contratual e existência de ciclos anteriores.")

    for item in pontos:
        _bullet(item)

    st.markdown("### No cl8us")
    for item in [
        "comece pela Calculadora de Reajustes;",
        "se houver mais de um ciclo, use o fluxo de múltiplos ciclos;",
        "se houver histórico anterior, preencha o Contexto do Contrato;",
        "só depois gere o Arquivo de Coleta e avance para Valores, se houver retroativo ou saldo remanescente a apurar.",
    ]:
        _bullet(item)


def resposta_aditivo(flags):
    _texto_corrente(
        "Pelo relato, eu olharia primeiro para o aditivo e para a ordem dos eventos. A pergunta prática é: esse aditivo mudou o valor, o escopo ou as quantidades antes do reajuste ser consolidado? Se sim, ele pode afetar o valor-base, a garantia e o saldo remanescente que será usado depois."
    )
    st.markdown("### Eu faria esta checagem")
    for item in [
        "veja a data de assinatura do aditivo;",
        "confirme se ele foi de acréscimo, supressão ou apenas ajuste de prazo/forma;",
        "verifique se o valor do aditivo já está incorporado ao valor formalizado do contrato;",
        "se ele alterou itens ou quantidades, confira se isso aparece no saldo remanescente;",
        "se houver reajuste depois, tome cuidado para não somar o aditivo duas vezes.",
    ]:
        _bullet(item)

    st.markdown("### No cl8us")
    for item in [
        "use Aditivos: 25% para controle de acréscimos/supressões;",
        "use Infos Prévias para registrar o instrumento e a data;",
        "use Valores apenas quando o aditivo já estiver refletido nos dados de execução, itens ou saldo remanescente.",
    ]:
        _bullet(item)


def resposta_supressao(flags):
    _texto_corrente(
        "Pelo que você contou, o cuidado principal é a ordem dos eventos. Se a supressão já foi formalizada antes de consolidar o reajuste ou o saldo remanescente, a parcela suprimida não deve continuar aparecendo como saldo a ser atualizado."
    )
    st.markdown("### Eu cuidaria disso assim")
    for item in [
        "confirme a data da supressão;",
        "confirme quais itens, quantidades ou valores foram reduzidos;",
        "garanta que o Arquivo de Coleta já reflita o saldo remanescente depois da supressão;",
        "não aplique reajuste sobre parcela que já saiu do contrato;",
        "use Aditivos: 25% para registrar o evento e Valores para processar o saldo correto.",
    ]:
        _bullet(item)


def resposta_valores(flags):
    _texto_corrente(
        "Pelo relato, parece que sua dúvida está na formação dos valores. A ideia básica do cl8us é não tratar o contrato como um valor único parado no tempo. Ele separa o que já foi executado, por ciclo, do que ainda resta executar como saldo remanescente."
    )
    st.markdown("### O raciocínio é este")
    for item in [
        "o Valor Total Atualizado do Contrato é execução atualizada por ciclo mais saldo remanescente atualizado;",
        "o saldo remanescente depende das quantidades ainda existentes no início do ciclo de referência;",
        "aditivos e supressões não devem ser somados por fora se já estiverem refletidos nos itens, na execução ou no saldo;",
        "se o fiscal informar saldos errados, o valor final também ficará errado.",
    ]:
        _bullet(item)

    st.markdown("### No cl8us")
    for item in [
        "gere o Arquivo de Coleta após a Calculadora;",
        "peça ao fiscal os valores financeiros e os itens/remanescentes;",
        "importe o arquivo preenchido em Valores;",
        "confira Valor Represado a Pagar, execução atualizada e saldo remanescente atualizado.",
    ]:
        _bullet(item)


def resposta_geral(flags):
    _texto_corrente(
        "Pelo relato, ainda ficou um pouco aberto qual é o ponto principal. Eu usaria o Orienta só para organizar o caminho e, depois, colocaria os dados reais nos módulos certos do cl8us."
    )
    st.markdown("### Para começar")
    for item in [
        "se a dúvida for sobre direito ao reajuste, comece na Calculadora de Reajustes;",
        "se a dúvida for sobre retroativo, saldo ou valor atualizado, use Valores depois do Arquivo de Coleta;",
        "se a dúvida for sobre garantia, use Gestão da Garantia;",
        "se a dúvida for sobre aditivo ou supressão, use Aditivos: 25%;",
        "se faltarem dados básicos, preencha Infos Prévias.",
    ]:
        _bullet(item)


def render_resultado(relato):
    flags, datas = analisar_relato(relato)
    tema = tema_principal(flags)

    st.markdown("## Leitura preliminar do relato")
    _alerta_informal()

    if tema == "garantia":
        resposta_garantia(flags)
    elif tema == "supressao":
        resposta_supressao(flags)
    elif tema == "aditivo":
        resposta_aditivo(flags)
    elif tema == "reajuste":
        resposta_reajuste(flags, datas)
        if flags["garantia"]:
            st.markdown("### Garantia também entrou no radar")
            for item in [
                "se o reajuste aumentar o valor do contrato, confira se a garantia precisa de endosso;",
                "o módulo Gestão da Garantia ajuda a comparar garantia exigida e garantia constituída.",
            ]:
                _bullet(item)
    elif tema == "valores":
        resposta_valores(flags)
        if flags["garantia"]:
            st.markdown("### E não esqueça da garantia")
            _bullet("se o valor atualizado aumentou, confira se a garantia antiga ainda é suficiente.")
    else:
        resposta_geral(flags)


# ============================================================
# Interface
# ============================================================

render_marca_topo()

st.markdown(
    """
    <style>
    .orienta-texto {
        text-align: justify;
        line-height: 1.58;
        margin: 0.35rem 0 1rem 0;
        color: #334155;
        max-width: 980px;
    }
    .orienta-alerta {
        max-width: none;
        color: #6B4E00;
        font-size: .92rem;
        font-style: italic;
        line-height: 1.45;
        margin: .35rem 0 1.05rem 0;
    }
    div[data-testid="stMarkdownContainer"] ul {
        line-height: 1.48;
        max-width: 980px;
    }
    .stTextArea textarea {
        line-height: 1.45;
    }
    div.stButton > button:first-child {
        background: #EDE9FE;
        color: #4C1D95;
        border: 1px solid #C4B5FD;
        border-radius: 10px;
        font-weight: 600;
    }
    div.stButton > button:first-child:hover {
        background: #DDD6FE;
        color: #3B0764;
        border: 1px solid #A78BFA;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("💡 Cl8us Orienta")
st.caption("Orientador preliminar para dúvidas, relatos e condução de casos concretos. Não realiza cálculo e não gera documentos.")

relato = st.text_area(
    "Descreva o caso, dúvida ou ponto que deseja entender",
    value="",
    height=210,
    placeholder="Ex.: contrato original, teve um aditivo, depois veio um reajuste. Como faço a garantia contratual?",
)

if st.button("Me dê uma mão", use_container_width=True):
    if not relato.strip():
        st.error("Me conte o caso ou a dúvida primeiro. Sem isso eu fico no escuro.")
    else:
        render_resultado(relato)
