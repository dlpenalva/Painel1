from datetime import datetime, date
from io import BytesIO
from zoneinfo import ZoneInfo
from html import escape

import pandas as pd
import streamlit as st

from _ui_utils import render_marca_topo

st.set_page_config(page_title="TLB · cl8us - Adequação Orçamentária", layout="wide")


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
    texto = texto.replace("R$", "").replace(" ", "").replace("\xa0", "")
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


def data_hora_brasilia():
    return datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")


def extrair_contexto_valores():
    res = st.session_state.get("resultado_valor_global", {}) or {}
    contexto = res.get("contexto_contratual_anterior", {}) or {}
    valor_original = parse_moeda_br(res.get("valor_original_contrato", contexto.get("valor_original_contrato", 0)))
    valor_formalizado_anterior = parse_moeda_br(
        res.get("valor_formalizado_anterior", contexto.get("valor_formalizado_anterior", valor_original))
    )
    if valor_formalizado_anterior <= 0:
        valor_formalizado_anterior = valor_original
    valor_pago = parse_moeda_br(res.get("valor_pago_efetivo", res.get("total_pago_faturado", 0)))
    valor_teorico = parse_moeda_br(res.get("valor_teorico_calculado", res.get("total_devido_reajustado", 0)))
    valor_represado = parse_moeda_br(res.get("valor_represado_a_pagar", res.get("delta_total", 0)))
    saldo_original = parse_moeda_br(res.get("remanescente_original", 0))
    saldo_atualizado = parse_moeda_br(res.get("remanescente_reajustado", 0))
    valor_total_atualizado = parse_moeda_br(res.get("valor_atualizado_contrato", res.get("valor_global_estoque", 0)))
    indice = texto_seguro(res.get("indice", ""), "não informado")
    variacao = parse_moeda_br(res.get("variacao_acumulada", res.get("fator_acumulado", 1) - 1 if res.get("fator_acumulado") else 0))
    return {
        "disponivel": bool(res),
        "valor_original": valor_original,
        "valor_formalizado_anterior": valor_formalizado_anterior,
        "valor_pago": valor_pago,
        "valor_teorico": valor_teorico,
        "valor_represado": valor_represado,
        "saldo_original": saldo_original,
        "saldo_atualizado": saldo_atualizado,
        "valor_total_atualizado": valor_total_atualizado,
        "indice": indice,
        "variacao": variacao,
    }


def input_moeda(label, valor_padrao, key, help=None):
    txt = st.text_input(label, value=moeda(valor_padrao, com_prefixo=False), key=key, help=help)
    valor = parse_moeda_br(txt)
    st.caption(moeda(valor))
    return valor


def texto_didatico_adequacao(dados):
    """Monta explicação didática da adequação, sem presumir fato não comprovado."""
    base = float(dados.get("valor_reservado", 0) or 0)
    valor_total = float(dados.get("valor_total_atualizado", 0) or 0)
    valor_represado = float(dados.get("valor_represado", 0) or 0)
    saldo_atualizado = float(dados.get("saldo_atualizado", 0) or 0)
    necessidade = float(dados.get("necessidade_considerada", 0) or 0)
    saldo_disponivel = float(dados.get("saldo_disponivel", 0) or 0)
    adequacoes_anteriores = float(dados.get("adequacoes_anteriores", 0) or 0)
    saldo_total_considerado = float(dados.get("saldo_total_considerado", saldo_disponivel + adequacoes_anteriores) or 0)
    complementacao = float(dados.get("complementacao", 0) or 0)
    base_label = texto_seguro(dados.get("base_label"), "base de comparação selecionada")

    diferenca_total = valor_total - base
    pct_dif = (diferenca_total / base * 100) if abs(base) > 0.004 else 0.0

    if diferenca_total > 0.004:
        movimento = f"aumentou em {moeda(diferenca_total)} ({str(f'{pct_dif:.2f}%').replace('.', ',')})"
        leitura = "há aumento da referência financeira do contrato em relação à base adotada"
    elif diferenca_total < -0.004:
        movimento = f"diminuiu em {moeda(abs(diferenca_total))} ({str(f'{abs(pct_dif):.2f}%').replace('.', ',')})"
        leitura = "há redução da referência financeira do contrato em relação à base adotada"
    else:
        movimento = "não apresentou diferença material"
        leitura = "não há diferença material entre o valor atualizado e a base adotada"

    if complementacao > 0.004:
        fechamento = (
            f"Como o saldo orçamentário disponível informado é de {moeda(saldo_disponivel)} e as adequações orçamentárias anteriores consideradas somam {moeda(adequacoes_anteriores)}, "
            f"o total disponível/considerado para esta simulação é de {moeda(saldo_total_considerado)}. A complementação estimada é de {moeda(complementacao)}. Esse é o valor sugerido para reforçar a programação orçamentária, "
            "salvo conferência da área responsável."
        )
    else:
        fechamento = (
            f"Como o saldo orçamentário disponível informado é de {moeda(saldo_disponivel)} e as adequações orçamentárias anteriores consideradas somam {moeda(adequacoes_anteriores)}, "
            f"o total disponível/considerado para esta simulação é de {moeda(saldo_total_considerado)}. Não foi identificada complementação adicional nesta simulação. Se a alteração representar supressão ou redução de escopo, "
            "a área responsável pode avaliar readequação para menor valor."
        )

    return (
        f"Parte-se da premissa de que a programação orçamentária foi estruturada a partir da seguinte base: {base_label}, "
        f"no valor de {moeda(base)}. A análise atual indica que o Valor Total Atualizado do Contrato {movimento}; em termos práticos, {leitura}. "
        f"Para esta estimativa, a necessidade considerada foi composta pelo valor represado/retroativo ({moeda(valor_represado)}) "
        f"somado ao saldo remanescente atualizado ({moeda(saldo_atualizado)}), resultando em {moeda(necessidade)}. "
        f"{fechamento}"
    )


def render_box_didatico(texto):
    st.markdown(
        f"""
        <div style="
            background:#FFF7E6;
            border:1px solid #F6C35B;
            border-radius:14px;
            padding:16px 18px;
            margin:14px 0 18px 0;
            color:#334155;
            line-height:1.45rem;
        ">
            <div style="font-weight:800; color:#7A4A00; font-size:1.02rem; margin-bottom:6px;">Como o sistema chegou ao valor?</div>
            <div style="font-size:0.95rem;">{escape(str(texto))}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_box_adequacoes_anteriores(conteudo_callable):
    """Box suave para registrar adequações orçamentárias anteriores com rastreabilidade."""
    st.markdown(
        """
        <div style="
            background:#F8FAFC;
            border:1px solid #C7D2FE;
            border-radius:14px;
            padding:14px 16px 8px 16px;
            margin:12px 0 16px 0;
        ">
            <div style="font-weight:800; color:#334155; font-size:1.02rem; margin-bottom:4px;">Adequações orçamentárias anteriores</div>
            <div style="color:#64748B; font-size:0.90rem; line-height:1.35rem; margin-bottom:10px;">
                Informe aqui adequações, suplementações ou reprogramações orçamentárias anteriores relacionadas ao mesmo contrato, quando ainda devam ser consideradas nesta simulação.
                Use linhas separadas para preservar rastreabilidade. Não informe endossos de garantia.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    return conteudo_callable()


def render_card_valor(label, valor, nota=""):
    """Card compacto para evitar truncamento dos valores importados."""
    nota_html = f"<div style='color:#94A3B8; font-size:0.78rem; margin-top:4px;'>{nota}</div>" if nota else ""
    st.markdown(
        f"""
        <div style="
            background:#FFFFFF;
            border:1px solid #E5EAF0;
            border-radius:12px;
            padding:12px 14px;
            min-height:92px;
        ">
            <div style="color:#64748B; font-size:0.84rem; font-weight:600; margin-bottom:6px;">{label}</div>
            <div style="color:#0F172A; font-size:1.12rem; font-weight:800; line-height:1.2; word-break:break-word; overflow-wrap:anywhere;">{moeda(valor)}</div>
            {nota_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_box_complementacao(valor, tipo_alteracao):
    """Dá destaque ao principal resultado do módulo."""
    if valor > 0.004:
        titulo = "Complementação orçamentária estimada"
        texto = "Valor sugerido para reforçar a programação orçamentária, sujeito à validação da área responsável."
        bg = "#EAF2F8"
        border = "#9EC5E8"
        cor = "#123B63"
    else:
        titulo = "Sem complementação adicional estimada"
        texto = "A simulação não indica reforço adicional. Em caso de supressão/redução, avalie readequação para menor valor."
        bg = "#F0FDF4"
        border = "#BBF7D0"
        cor = "#166534"
    st.markdown(
        f"""
        <div style="
            background:{bg};
            border:1px solid {border};
            border-radius:16px;
            padding:18px 20px;
            margin:10px 0 18px 0;
        ">
            <div style="color:{cor}; font-weight:800; font-size:1.02rem; margin-bottom:4px;">{titulo}</div>
            <div style="color:#0F172A; font-weight:900; font-size:1.85rem; line-height:1.2;">{moeda(valor)}</div>
            <div style="color:#475569; font-size:0.90rem; margin-top:5px;">{texto}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )



def montar_memorando(dados):
    """Monta texto simples para edição/cola, com parágrafos numerados e memória ao final."""
    contrato = texto_seguro(dados.get("contrato"))
    tipo = texto_seguro(dados.get("tipo_alteracao"), "alteração contratual")
    cronograma = dados.get("cronograma", []) or []
    total_cronograma = sum(float(item.get("Valor", 0) or 0) for item in cronograma)
    obs = str(dados.get("observacao", "")).strip()
    adequacoes_linhas = dados.get("adequacoes_anteriores_linhas", []) or []
    if adequacoes_linhas:
        detalhamento_adequacoes = "; ".join(
            f"{texto_seguro(item.get('Referência'), '[campo a preencher]')} / {texto_seguro(item.get('Exercício'), '[campo a preencher]')}: {moeda(item.get('Valor', 0))}"
            for item in adequacoes_linhas
        )
    else:
        detalhamento_adequacoes = "não informadas"

    linhas = []
    linhas.append("MEMORANDO")
    linhas.append("")
    linhas.append(f"Assunto: Solicitação de adequação orçamentária — Contrato {contrato}")
    linhas.append("")
    linhas.append(
        f"1. Solicita-se adequação orçamentária para o Contrato {contrato}, em razão de {tipo}, "
        "conforme programação abaixo:"
    )
    linhas.append("")
    linhas.append("Exercício\tValor")
    for item in cronograma:
        exercicio = texto_seguro(item.get("Exercício"), "[campo a preencher]")
        valor = float(item.get("Valor", 0) or 0)
        if abs(valor) > 0.004:
            linhas.append(f"{exercicio}\t{moeda(valor)}")
    linhas.append(f"Total\t{moeda(total_cronograma)}")
    linhas.append("")
    if obs:
        linhas.append(f"2. {obs}")
        linhas.append("")
        num_ajuste = 3
        num_memoria = 4
    else:
        num_ajuste = 2
        num_memoria = 3
    linhas.append(
        f"{num_ajuste}. O ajuste é necessário para compatibilizar a programação orçamentária com o novo valor contratual estimado, "
        "considerando a alteração informada, a base de comparação adotada e os efeitos financeiros correspondentes."
    )
    linhas.append("")
    linhas.append(
        f"{num_memoria}. Memória de cálculo: base de comparação orçamentária adotada: "
        f"{texto_seguro(dados.get('base_label'), 'base informada')}, no valor de {moeda(dados.get('valor_reservado', 0))}; "
        f"valor já pago: {moeda(dados.get('valor_pago', 0))}; "
        f"saldo orçamentário disponível informado: {moeda(dados.get('saldo_disponivel', 0))}; "
        f"adequações orçamentárias anteriores consideradas: {moeda(dados.get('adequacoes_anteriores', 0))}; "
        f"detalhamento das adequações anteriores: {detalhamento_adequacoes}; "
        f"saldo total considerado: {moeda(dados.get('saldo_total_considerado', dados.get('saldo_disponivel', 0)))}; "
        f"valor represado/retroativo estimado: {moeda(dados.get('valor_represado', 0))}; "
        f"saldo remanescente atualizado: {moeda(dados.get('saldo_atualizado', 0))}; "
        f"necessidade considerada: {moeda(dados.get('necessidade_considerada', 0))}; "
        f"complementação orçamentária estimada: {moeda(dados.get('complementacao', 0))}."
    )
    linhas.append("")
    linhas.append(f"Gerado em: {data_hora_brasilia()}.")
    return "\n".join(linhas)


def _adicionar_texto_com_placeholder(paragraph, conteudo):
    """Adiciona texto ao parágrafo destacando [campo a preencher] em amarelo."""
    try:
        from docx.enum.text import WD_COLOR_INDEX
    except Exception:
        paragraph.add_run(str(conteudo))
        return
    import re as _re
    partes = _re.split(r"(\[campo a preencher[^\]]*\])", str(conteudo), flags=_re.IGNORECASE)
    for parte in partes:
        if not parte:
            continue
        run = paragraph.add_run(parte)
        if _re.fullmatch(r"\[campo a preencher[^\]]*\]", parte, flags=_re.IGNORECASE):
            run.bold = True
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW


def _add_paragrafo_numerado(document, numero, texto):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run(f"{numero}. ").bold = True
    _adicionar_texto_com_placeholder(p, texto)
    return p


def gerar_docx_memorando_estruturado(dados):
    """Gera DOCX com parágrafos numerados, tabela de cronograma com borda e memória no último parágrafo."""
    try:
        from docx import Document
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
        from docx.shared import Inches, Pt
    except Exception as exc:
        raise RuntimeError("A biblioteca python-docx não está instalada.") from exc

    document = Document()
    section = document.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)
    styles = document.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10.5)

    contrato = texto_seguro(dados.get("contrato"))
    tipo = texto_seguro(dados.get("tipo_alteracao"), "alteração contratual")
    cronograma = dados.get("cronograma", []) or []
    cronograma_validos = [item for item in cronograma if abs(float(item.get("Valor", 0) or 0)) > 0.004]
    total_cronograma = sum(float(item.get("Valor", 0) or 0) for item in cronograma_validos)
    obs = str(dados.get("observacao", "")).strip()

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = title.add_run("MEMORANDO")
    r.bold = True
    r.font.size = Pt(13)

    p_assunto = document.add_paragraph()
    p_assunto.add_run("Assunto: ").bold = True
    _adicionar_texto_com_placeholder(p_assunto, f"Solicitação de adequação orçamentária — Contrato {contrato}")

    _add_paragrafo_numerado(
        document,
        1,
        f"Solicita-se adequação orçamentária para o Contrato {contrato}, em razão de {tipo}, conforme programação abaixo:",
    )

    tabela = document.add_table(rows=1, cols=2)
    tabela.alignment = WD_TABLE_ALIGNMENT.CENTER
    tabela.style = "Table Grid"
    hdr = tabela.rows[0].cells
    hdr[0].text = "Exercício"
    hdr[1].text = "Valor"
    for cell in hdr:
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        for par in cell.paragraphs:
            for run in par.runs:
                run.bold = True

    for item in cronograma_validos:
        row = tabela.add_row().cells
        row[0].text = texto_seguro(item.get("Exercício"), "[campo a preencher]")
        row[1].text = moeda(float(item.get("Valor", 0) or 0))
    row = tabela.add_row().cells
    row[0].text = "Total"
    row[1].text = moeda(total_cronograma)
    for cell in row:
        for par in cell.paragraphs:
            for run in par.runs:
                run.bold = True

    numero = 2
    if obs:
        _add_paragrafo_numerado(document, numero, obs)
        numero += 1

    _add_paragrafo_numerado(
        document,
        numero,
        "O ajuste é necessário para compatibilizar a programação orçamentária com o novo valor contratual estimado, considerando a alteração informada, a base de comparação adotada e os efeitos financeiros correspondentes.",
    )
    numero += 1

    adequacoes_linhas = dados.get("adequacoes_anteriores_linhas", []) or []
    if adequacoes_linhas:
        detalhamento_adequacoes = "; ".join(
            f"{texto_seguro(item.get('Referência'), '[campo a preencher]')} / {texto_seguro(item.get('Exercício'), '[campo a preencher]')}: {moeda(item.get('Valor', 0))}"
            for item in adequacoes_linhas
        )
    else:
        detalhamento_adequacoes = "não informadas"

    memoria = (
        f"Memória de cálculo: base de comparação orçamentária adotada: {texto_seguro(dados.get('base_label'), 'base informada')}, "
        f"no valor de {moeda(dados.get('valor_reservado', 0))}; "
        f"valor já pago: {moeda(dados.get('valor_pago', 0))}; "
        f"saldo orçamentário disponível informado: {moeda(dados.get('saldo_disponivel', 0))}; "
        f"adequações orçamentárias anteriores consideradas: {moeda(dados.get('adequacoes_anteriores', 0))}; "
        f"detalhamento das adequações anteriores: {detalhamento_adequacoes}; "
        f"saldo total considerado: {moeda(dados.get('saldo_total_considerado', dados.get('saldo_disponivel', 0)))}; "
        f"valor represado/retroativo estimado: {moeda(dados.get('valor_represado', 0))}; "
        f"saldo remanescente atualizado: {moeda(dados.get('saldo_atualizado', 0))}; "
        f"necessidade considerada: {moeda(dados.get('necessidade_considerada', 0))}; "
        f"complementação orçamentária estimada: {moeda(dados.get('complementacao', 0))}."
    )
    _add_paragrafo_numerado(document, numero, memoria)

    document.add_paragraph(f"Gerado em: {data_hora_brasilia()}.")

    buffer = BytesIO()
    document.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


render_marca_topo()
st.title("Adequação Orçamentária")
st.caption("Estimativa para solicitação de adequação/readequação orçamentária decorrente de reajuste, repactuação, aditivo ou outra alteração contratual.")

ctx = extrair_contexto_valores()

with st.expander("Dados importados da análise atual", expanded=True):
    if ctx["disponivel"]:
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            render_card_valor("Valor original", ctx["valor_original"])
        with c2:
            render_card_valor("Valor represado", ctx["valor_represado"])
        with c3:
            render_card_valor("Saldo remanescente atualizado", ctx["saldo_atualizado"])
        with c4:
            render_card_valor("Valor Total Atualizado", ctx["valor_total_atualizado"])
        st.caption(f"Índice: {ctx['indice']} · Variação acumulada: {pct(ctx['variacao'])}")
    else:
        st.info(
            "Não há resultado do módulo Valores na sessão. O módulo também funciona de forma independente: "
            "preencha manualmente contrato, base de comparação, saldo disponível, valores de retroativo/remanescente e cronograma."
        )

st.subheader("Parâmetros da solicitação")
col_a, col_b, col_c = st.columns([1.1, 1.1, 1])
with col_a:
    contrato = st.text_input("Contrato", value=st.session_state.get("adequacao_contrato", ""), placeholder="Ex.: TLB-CTR-2023/00037")
with col_b:
    tipo_alteracao = st.selectbox(
        "Origem da alteração",
        ["Reajuste", "Repactuação", "Aditivo de acréscimo", "Aditivo de supressão/readequação", "Reequilíbrio", "Outro"],
        index=0,
    )
with col_c:
    competencia_corte = st.text_input("Competência de corte", value=date.today().strftime("%m/%Y"), placeholder="Ex.: 05/2026")

st.subheader("Dados orçamentários")
st.caption("Escolha a base de comparação. Se o contrato já tiver aditivos, reajustes ou repactuações formalizados, use a base importada/formalizada ou informe manualmente.")

base_opcao = st.radio(
    "Base de comparação orçamentária",
    [
        "Valor original do contrato",
        "Valor formalizado antes desta análise",
        "Valor informado manualmente",
    ],
    index=0,
    horizontal=True,
    help=(
        "Use valor original quando a reserva inicial foi feita na assinatura. "
        "Use valor formalizado quando já houver aditivos, reajustes ou repactuações incorporados. "
        "Use valor manual quando a área orçamentária já tiver informado outra base."
    ),
)

with st.expander("Como escolher a base de comparação", expanded=False):
    st.write(
        "Use **Valor original do contrato** quando a leitura desejada for comparar a programação inicial com a necessidade atual. "
        "Use **Valor formalizado antes desta análise** quando o módulo Valores tiver recebido essa fotografia administrativa anterior — por exemplo, contrato já alterado por aditivos, reajustes ou repactuações formalizados antes da análise atual. "
        "Esse campo vem do resultado do módulo Valores, especialmente de `valor_formalizado_anterior` ou do Contexto do Contrato. Quando essa base não tiver sido informada, o sistema usa o valor original como fallback. "
        "Use **Valor informado manualmente** quando a área orçamentária já tiver passado a reserva, dotação ou base de referência correta. "
        "O sistema não tenta desfazer automaticamente reajustes/aditivos já incorporados, porque isso poderia gerar leitura artificial; para uma análise fria, informe manualmente a base que deseja comparar."
    )

if base_opcao == "Valor original do contrato":
    base_padrao = ctx["valor_original"]
    base_label = "valor original do contrato"
elif base_opcao == "Valor formalizado antes desta análise":
    base_padrao = ctx.get("valor_formalizado_anterior", ctx["valor_original"])
    base_label = "valor formalizado antes desta análise"
else:
    base_padrao = ctx.get("valor_formalizado_anterior", ctx["valor_original"])
    base_label = "valor informado manualmente"

col1, col2, col3 = st.columns(3)
with col1:
    valor_reservado = input_moeda(
        "Valor de referência orçamentária",
        base_padrao,
        "adequacao_valor_reservado",
        help="Base usada para comparar a necessidade orçamentária atual. Pode ser o valor original, o valor formalizado anterior ou valor informado manualmente.",
    )
with col2:
    valor_pago = input_moeda("Valor já pago", ctx["valor_pago"], "adequacao_valor_pago")
with col3:
    saldo_padrao = max(valor_reservado - valor_pago, 0.0)
    saldo_disponivel = input_moeda("Saldo orçamentário disponível", saldo_padrao, "adequacao_saldo_disponivel")

def _capturar_adequacoes_anteriores():
    df_padrao_adequacoes = pd.DataFrame([
        {"Referência": "", "Exercício": "", "Valor": "0,00", "Observação": ""},
    ])
    df_adequacoes = st.data_editor(
        df_padrao_adequacoes,
        hide_index=True,
        use_container_width=True,
        num_rows="dynamic",
        key="adequacao_anteriores_editor",
        column_config={
            "Referência": st.column_config.TextColumn("Referência", help="Ex.: TLB-DES-2026/xxxxx, RC, despacho, orçamento anterior."),
            "Exercício": st.column_config.TextColumn("Exercício", help="Ex.: 2026"),
            "Valor": st.column_config.TextColumn("Valor", help="Informe no padrão brasileiro. Ex.: 41.349,35"),
            "Observação": st.column_config.TextColumn("Observação", help="Comentário opcional sobre a adequação anterior."),
        },
    )
    linhas = []
    for _, row in df_adequacoes.iterrows():
        referencia = texto_seguro(row.get("Referência", ""), "")
        exercicio = texto_seguro(row.get("Exercício", ""), "")
        valor = parse_moeda_br(row.get("Valor", 0))
        observacao_linha = texto_seguro(row.get("Observação", ""), "")
        if referencia or exercicio or observacao_linha or abs(valor) > 0.004:
            linhas.append({
                "Referência": referencia or "[campo a preencher]",
                "Exercício": exercicio or "[campo a preencher]",
                "Valor": valor,
                "Observação": observacao_linha,
            })
    total = round(sum(item["Valor"] for item in linhas), 2)
    st.caption(f"Total de adequações anteriores consideradas: {moeda(total)}")
    return linhas, total

adequacoes_anteriores_linhas, adequacoes_anteriores = render_box_adequacoes_anteriores(_capturar_adequacoes_anteriores)

st.subheader("Memória de cálculo")
col4, col5, col6 = st.columns(3)
with col4:
    valor_represado = input_moeda("Valor represado/retroativo", ctx["valor_represado"], "adequacao_valor_represado")
with col5:
    saldo_atualizado = input_moeda("Saldo remanescente atualizado", ctx["saldo_atualizado"], "adequacao_saldo_atualizado")
with col6:
    valor_total_atualizado = input_moeda("Valor Total Atualizado do Contrato", ctx["valor_total_atualizado"], "adequacao_vtac")

necessidade_considerada = round(valor_represado + saldo_atualizado, 2)
saldo_total_considerado = round(saldo_disponivel + adequacoes_anteriores, 2)
complementacao_estimada = round(max(necessidade_considerada - saldo_total_considerado, 0.0), 2)

colr1, colr2, colr3 = st.columns(3)
with colr1:
    render_card_valor("Necessidade considerada", necessidade_considerada)
with colr2:
    render_card_valor("Saldo disponível informado", saldo_disponivel)
with colr3:
    render_card_valor("Adequações anteriores consideradas", adequacoes_anteriores)
render_box_complementacao(complementacao_estimada, tipo_alteracao)

dados_didaticos = {
    "valor_reservado": valor_reservado,
    "base_label": base_label,
    "valor_total_atualizado": valor_total_atualizado,
    "valor_represado": valor_represado,
    "saldo_atualizado": saldo_atualizado,
    "necessidade_considerada": necessidade_considerada,
    "saldo_disponivel": saldo_disponivel,
    "adequacoes_anteriores": adequacoes_anteriores,
    "adequacoes_anteriores_linhas": adequacoes_anteriores_linhas,
    "saldo_total_considerado": saldo_total_considerado,
    "complementacao": complementacao_estimada,
}
render_box_didatico(texto_didatico_adequacao(dados_didaticos))

st.subheader("Cronograma orçamentário proposto")
st.caption("Inclua quantos exercícios forem necessários. O total do cronograma será usado no memorando.")
ano_atual = date.today().year
cronograma_padrao = pd.DataFrame([
    {"Exercício": str(ano_atual), "Valor": moeda(complementacao_estimada, com_prefixo=False)},
    {"Exercício": str(ano_atual + 1), "Valor": moeda(0.0, com_prefixo=False)},
])
cronograma_editado = st.data_editor(
    cronograma_padrao,
    hide_index=True,
    use_container_width=True,
    num_rows="dynamic",
    key="adequacao_cronograma_editor",
    column_config={
        "Exercício": st.column_config.TextColumn("Exercício", help="Informe o ano/exercício. Ex.: 2026"),
        "Valor": st.column_config.TextColumn("Valor", help="Informe o valor no padrão brasileiro. Ex.: 41.349,35"),
    },
)
cronograma_linhas = []
for _, row in cronograma_editado.iterrows():
    exercicio = texto_seguro(row.get("Exercício", ""), "")
    valor = parse_moeda_br(row.get("Valor", 0))
    if exercicio or abs(valor) > 0.004:
        cronograma_linhas.append({"Exercício": exercicio or "[campo a preencher]", "Valor": valor})
total_cronograma = round(sum(item["Valor"] for item in cronograma_linhas), 2)
st.metric("Total do cronograma", moeda(total_cronograma))


observacao = st.text_area(
    "Observação complementar",
    value="",
    placeholder="Ex.: adequação necessária para compatibilizar o cronograma com o novo valor do termo aditivo/reajuste.",
    height=90,
)

dados = {
    "contrato": contrato,
    "tipo_alteracao": tipo_alteracao,
    "competencia_corte": competencia_corte,
    "base_label": base_label,
    "valor_reservado": valor_reservado,
    "valor_pago": valor_pago,
    "saldo_disponivel": saldo_disponivel,
    "adequacoes_anteriores": adequacoes_anteriores,
    "adequacoes_anteriores_linhas": adequacoes_anteriores_linhas,
    "saldo_total_considerado": saldo_total_considerado,
    "valor_represado": valor_represado,
    "saldo_atualizado": saldo_atualizado,
    "valor_total_atualizado": valor_total_atualizado,
    "necessidade_considerada": necessidade_considerada,
    "complementacao": complementacao_estimada,
    "cronograma": cronograma_linhas,
    "observacao": observacao,
}

st.subheader("Memorando/Despacho")
texto = montar_memorando(dados)
texto_editado = st.text_area("Texto gerado", value=texto, height=360)

col_down1, col_down2 = st.columns([1, 1])
with col_down1:
    st.download_button(
        "Baixar memorando em TXT",
        data=texto_editado.encode("utf-8"),
        file_name="memorando_adequacao_orcamentaria.txt",
        mime="text/plain",
        type="primary",
    )
with col_down2:
    try:
        docx_bytes = gerar_docx_memorando_estruturado(dados)
        st.download_button(
            "Baixar memorando em DOCX",
            data=docx_bytes,
            file_name="memorando_adequacao_orcamentaria.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as exc:
        st.warning(f"Não foi possível gerar DOCX: {exc}")
