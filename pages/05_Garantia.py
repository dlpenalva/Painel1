"""Garantia Contratual — método único (Histórico dos valores totais do contrato).

Regra contratual absoluta (Cláusula Décima): a garantia corresponde a 5% do
valor total do contrato após cada instrumento. O endosso/adequação de cada
instrumento é a diferença entre a sua garantia e a garantia do instrumento
imediatamente anterior.

Toda a matemática vive no motor puro ``_garantia_calculo`` (Decimal +
ROUND_HALF_UP), permitindo testes focais. Esta página cuida apenas da interface:
histórico manual dos valores totais, VTA da análise atual fornecido pelo Claus,
memória de cálculo, resultado, comunicação (TXT) e relatório (PDF).
"""
import pandas as pd
import streamlit as st

from _ui_utils import render_cabecalho_pagina
from _garantia_calculo import (
    REPORTLAB_OK,
    moeda,
    extrair_vta,
    normalizar_linhas_historico,
    construir_linha_do_tempo,
    formatar_endosso,
    resumo_resultado,
    gerar_texto_comunicacao,
    montar_txt_bytes,
    gerar_pdf_garantia,
)

st.set_page_config(page_icon="assets/cl8us_favicon_512.png", page_title="Análises de Reajustes - Garantia", layout="wide")


def css():
    st.markdown(
        """
        <style>
        .garantia-card {
            background: #F6F8FA;
            border: 1px solid #E1E6EB;
            border-radius: 14px;
            padding: 18px 20px;
            margin: 6px 0 14px 0;
        }
        .garantia-card-destaque {
            background: #EAF2F8;
            border: 1px solid #C8D9E8;
            border-radius: 14px;
            padding: 20px 22px;
            margin: 8px 0 16px 0;
        }
        .garantia-label { color: #475569; font-size: 0.92rem; margin-bottom: 4px; }
        .garantia-valor { color: #1F2937; font-size: 1.55rem; font-weight: 700; line-height: 1.2; }
        .garantia-valor-destaque { color: #123B63; font-size: 2rem; font-weight: 800; line-height: 1.2; }
        .garantia-nota { color: #64748B; font-size: 0.88rem; margin-top: 6px; }
        .garantia-tabela-wrap { width: 100%; overflow-x: auto; margin: 10px 0 18px 0; }
        table.garantia-tabela { width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 0.92rem; }
        table.garantia-tabela th { background: #E6F0F7; color: #173B5D; border: 1px solid #C5D6E2; padding: 9px 10px; text-align: left; font-weight: 700; }
        table.garantia-tabela td { border: 1px solid #E5EAF0; padding: 9px 10px; vertical-align: top; white-space: normal; overflow-wrap: anywhere; word-break: normal; line-height: 1.35; }
        table.garantia-tabela td.valor { white-space: nowrap; overflow-wrap: normal; text-align: right; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def card(label, valor, nota=None, destaque=False):
    classe = "garantia-card-destaque" if destaque else "garantia-card"
    valor_classe = "garantia-valor-destaque" if destaque else "garantia-valor"
    nota_html = f'<div class="garantia-nota">{nota}</div>' if nota else ""
    st.markdown(
        f"""
        <div class="{classe}">
            <div class="garantia-label">{label}</div>
            <div class="{valor_classe}">{valor}</div>
            {nota_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_memoria_html(linha_do_tempo):
    """Renderiza a memória de cálculo completa como tabela HTML."""
    from html import escape

    linhas = []
    for linha in linha_do_tempo:
        instrumento = escape(str(linha.get("instrumento", "")))
        valor_total = escape(moeda(linha.get("valor_total", 0)))
        garantia = escape(moeda(linha.get("garantia", 0)))
        endosso = escape(formatar_endosso(linha.get("endosso"), linha.get("tipo_endosso", "inicial")))
        linhas.append(
            f"<tr><td>{instrumento}</td><td class='valor'>{valor_total}</td>"
            f"<td class='valor'>{garantia}</td><td class='valor'>{endosso}</td></tr>"
        )
    html = """
    <div class="garantia-tabela-wrap">
      <table class="garantia-tabela">
        <colgroup>
          <col style="width: 34%;">
          <col style="width: 24%;">
          <col style="width: 21%;">
          <col style="width: 21%;">
        </colgroup>
        <thead>
          <tr><th>Instrumento</th><th>Total do contrato</th><th>Garantia — 5%</th><th>Endosso / Adequação</th></tr>
        </thead>
        <tbody>
          {linhas}
        </tbody>
      </table>
    </div>
    """.format(linhas="\n".join(linhas))
    st.markdown(html, unsafe_allow_html=True)


# ============================================================
# Interface
# ============================================================

css()
render_cabecalho_pagina("Garantia Contratual", "")
if st.button("← Voltar para Central", key="voltar_central_garantia"):
    st.switch_page("pages/03_Valor_Global.py")
st.caption("Calcule a garantia de 5% e a adequação decorrente das alterações do valor total do contrato.")

# ------------------------------------------------------------
# 1) Histórico dos valores totais do contrato (entrada manual)
# ------------------------------------------------------------
st.subheader("Histórico dos valores totais do contrato")
st.caption(
    "Informe, na ordem cronológica, cada instrumento que alterou o valor total do contrato e o "
    "respectivo valor total consolidado após aquele instrumento. A primeira linha é normalmente o "
    "contrato original. Identifique o instrumento livremente (Contrato, Apostila 1, Aditivo 1...)."
)

historico_padrao = pd.DataFrame(
    [
        {"Instrumento": "Contrato", "Valor total do contrato": ""},
        {"Instrumento": "Apostila 1", "Valor total do contrato": ""},
    ]
)
historico_editado = st.data_editor(
    historico_padrao,
    hide_index=True,
    use_container_width=True,
    num_rows="dynamic",
    key="garantia_historico_editor",
    column_config={
        "Instrumento": st.column_config.TextColumn(
            "Instrumento",
            help="Texto livre: Contrato, Apostila 1, Apostila 2, Aditivo 1, Termo Aditivo 1...",
            width="medium",
        ),
        "Valor total do contrato": st.column_config.TextColumn(
            "Valor total do contrato",
            help="Valor total consolidado do contrato após o instrumento. Ex.: 3.013.124,45 ou R$ 3.013.124,45",
            width="medium",
        ),
    },
)

registros_historico = historico_editado.to_dict("records") if isinstance(historico_editado, pd.DataFrame) else []
linhas_historico, avisos_historico = normalizar_linhas_historico(registros_historico)
for aviso in avisos_historico:
    st.warning(aviso)

# ------------------------------------------------------------
# 2) Análise atual (instrumento manual + VTA fornecido pelo Claus)
# ------------------------------------------------------------
st.subheader("Análise atual")
resultado_valor_global = st.session_state.get("resultado_valor_global", {}) or {}
vta_atual = extrair_vta(resultado_valor_global)

col_atual_1, col_atual_2 = st.columns(2)
with col_atual_1:
    instrumento_atual = st.text_input(
        "Instrumento da análise atual",
        value="",
        placeholder="Ex.: Apostila 2, Aditivo 2...",
        help="Identifique o instrumento em análise. O valor total correspondente é o VTA fornecido pelo Claus.",
    ).strip()
with col_atual_2:
    if vta_atual is not None:
        st.metric("VTA atual fornecido pelo Claus", moeda(vta_atual))
    else:
        st.metric("VTA atual fornecido pelo Claus", "—")

st.caption(
    "O VTA (Valor Total Atualizado do Contrato) da análise atual é fornecido automaticamente pelo "
    "módulo Valor Global. Todos os valores anteriores são informados manualmente no histórico acima."
)

# ------------------------------------------------------------
# Validações de bloqueio (sem cálculo silencioso com valor zero)
# ------------------------------------------------------------
if vta_atual is None:
    st.error(
        "VTA atual indisponível. Conclua ou reabra o resultado do módulo Valor Global "
        "(página 03 · Valor Global) para que o VTA da análise atual seja fornecido ao Claus. "
        "Nenhum cálculo, PDF ou TXT é gerado com VTA inexistente."
    )
    st.stop()

if not instrumento_atual:
    st.info("Informe o **instrumento da análise atual** para gerar a memória de cálculo, o resultado, o TXT e o PDF.")
    st.stop()

# ------------------------------------------------------------
# 3) Memória de cálculo
# ------------------------------------------------------------
itens = list(linhas_historico) + [{"instrumento": instrumento_atual, "valor_total": vta_atual}]
linha_do_tempo = construir_linha_do_tempo(itens)
_, garantia_total, endosso_atual, tipo_endosso = resumo_resultado(linha_do_tempo)

st.subheader("Memória de cálculo")
st.caption(
    "A garantia de cada linha é 5% do valor total daquela linha. O endosso/adequação de cada linha é a "
    "diferença entre a sua garantia e a garantia da linha imediatamente anterior."
)
render_memoria_html(linha_do_tempo)

# ------------------------------------------------------------
# 4) Resultado principal
# ------------------------------------------------------------
st.subheader("Resultado")
col_r1, col_r2, col_r3 = st.columns(3)
with col_r1:
    card("VTA atual", moeda(vta_atual))
with col_r2:
    card("Garantia total exigida", moeda(garantia_total), "5% do valor total do contrato.")
with col_r3:
    if tipo_endosso == "reducao":
        card("Adequação apurada", f"Redução de {moeda(abs(endosso_atual))}", "Diferença negativa frente à garantia anterior.", destaque=True)
    elif tipo_endosso == "inicial":
        card("Garantia inicial", moeda(garantia_total), "Não há instrumento anterior.", destaque=True)
    else:
        card("Endosso necessário", moeda(endosso_atual), "Diferença frente à garantia anterior.", destaque=True)

if tipo_endosso == "reducao":
    st.warning(
        f"O instrumento atual ({instrumento_atual}) apura redução de {moeda(abs(endosso_atual))} "
        "em relação à garantia anterior. A adequação não é automática: submeta à análise e à aceitação da Telebras."
    )
elif tipo_endosso == "sem_alteracao":
    st.info("Não há alteração de garantia decorrente do instrumento atual (endosso de R$ 0,00). Verifique eventual adequação do prazo de validade.")
elif tipo_endosso == "inicial":
    st.success(f"Garantia inicial de {moeda(garantia_total)} (5% do VTA atual).")
else:
    st.success(f"Endosso complementar necessário de {moeda(endosso_atual)} para o instrumento atual ({instrumento_atual}).")

# ------------------------------------------------------------
# 5) Comunicação à contratada
# ------------------------------------------------------------
st.subheader("Comunicação à contratada")
col_c1, col_c2 = st.columns(2)
with col_c1:
    numero_contrato = st.text_input("Número do contrato", value="", placeholder="Ex.: 001/2024").strip()
with col_c2:
    contratada = st.text_input("Nome da contratada (opcional)", value="").strip()

texto_comunicacao = gerar_texto_comunicacao(
    numero_contrato=numero_contrato,
    vta_atual=vta_atual,
    garantia_total=garantia_total,
    endosso=endosso_atual,
    tipo_endosso=tipo_endosso,
    contratada=contratada,
)
st.text_area("Texto da comunicação", value=texto_comunicacao, height=340)
st.download_button(
    "Baixar comunicação em TXT",
    data=montar_txt_bytes(texto_comunicacao),
    file_name="comunicacao_garantia_contratual.txt",
    mime="text/plain",
    use_container_width=False,
)

# ------------------------------------------------------------
# 6) Relatório PDF (memória completa)
# ------------------------------------------------------------
st.subheader("Relatório")
dados_pdf = {
    "numero_contrato": numero_contrato,
    "contratada": contratada,
    "vta_atual": vta_atual,
    "garantia_total": garantia_total,
    "endosso": endosso_atual,
    "tipo_endosso": tipo_endosso,
    "linha_do_tempo": linha_do_tempo,
    "texto_comunicacao": texto_comunicacao,
}
if REPORTLAB_OK:
    pdf_bytes = gerar_pdf_garantia(dados_pdf)
    st.session_state["arquivo_garantia_pdf"] = pdf_bytes
    st.download_button(
        "Baixar memória de cálculo (PDF)",
        data=pdf_bytes,
        file_name="relatorio_garantia_contratual.pdf",
        mime="application/pdf",
        type="primary",
        use_container_width=False,
    )
else:
    st.info("Instale reportlab para gerar o PDF: pip install reportlab")

# ------------------------------------------------------------
# 7) Resultado consolidado em session_state (integração)
# ------------------------------------------------------------
st.session_state["resultado_garantia"] = {
    "vta_atual": float(vta_atual),
    "percentual_garantia": 0.05,
    "garantia_total_exigida": float(garantia_total),
    "endosso_ou_reducao": float(endosso_atual) if endosso_atual is not None else 0.0,
    "instrumento_atual": instrumento_atual,
    "tipo_endosso": tipo_endosso,
    "historico": [
        {
            "instrumento": linha["instrumento"],
            "valor_total": float(linha["valor_total"]),
            "garantia": float(linha["garantia"]),
            "endosso": float(linha["endosso"]) if linha["endosso"] is not None else None,
            "tipo_endosso": linha["tipo_endosso"],
        }
        for linha in linha_do_tempo
    ],
    # Chave legada preservada por compatibilidade (Saneador só checa dict não vazio).
    "metodo_garantia": "Histórico dos Valores Totais do Contrato",
    "endosso_necessario": float(endosso_atual) if (endosso_atual is not None and endosso_atual > 0) else 0.0,
}
