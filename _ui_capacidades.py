"""Componentes visuais do processamento progressivo da apuracao."""

from __future__ import annotations

import html
from typing import Any, Iterable

import streamlit as st


_CLASSES = {
    "completo": ("apuracao-ok", "Completo"),
    "parcial": ("apuracao-parcial", "Parcial"),
    "nao_informado": ("apuracao-neutro", "Não informado"),
    "bloqueado": ("apuracao-bloqueado", "Bloqueado"),
}


def _css() -> None:
    st.markdown(
        """
        <style>
        .apuracao-shell { background:#FFFFFF; border:1px solid #D8E2EC; border-radius:14px; padding:17px 18px; margin:12px 0 16px; box-shadow:0 5px 18px rgba(31,78,121,.06); }
        .apuracao-head { display:flex; justify-content:space-between; gap:16px; align-items:flex-start; margin-bottom:13px; }
        .apuracao-title { color:#123B63; font-size:1.08rem; font-weight:800; margin:0; }
        .apuracao-subtitle { color:#52667A; font-size:.86rem; line-height:1.4; margin-top:3px; }
        .apuracao-progress { color:#1F4E79; background:#EDF5FB; border:1px solid #C6D9E8; border-radius:999px; padding:5px 10px; font-size:.76rem; font-weight:700; white-space:nowrap; }
        .apuracao-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:9px; }
        .apuracao-card { background:#F9FBFD; border:1px solid #DCE6EF; border-radius:10px; padding:10px 11px; min-height:91px; }
        .apuracao-card-top { display:flex; justify-content:space-between; gap:7px; align-items:center; }
        .apuracao-name { color:#243B53; font-size:.87rem; font-weight:750; }
        .apuracao-badge { display:inline-flex; align-items:center; gap:5px; border-radius:999px; padding:3px 7px; font-size:.70rem; font-weight:750; line-height:1.15; }
        .apuracao-badge::before { content:""; width:7px; height:7px; border-radius:50%; background:currentColor; }
        .apuracao-ok { color:#17623A; background:#EAF7EF; border:1px solid #B8DEC7; }
        .apuracao-parcial { color:#8A5A00; background:#FFF8DB; border:1px solid #EAD38A; }
        .apuracao-neutro { color:#5B6875; background:#F1F4F7; border:1px solid #D7DEE5; }
        .apuracao-bloqueado { color:#9A3F16; background:#FFF0E8; border:1px solid #F0C1A9; }
        .apuracao-detail { color:#617386; font-size:.76rem; line-height:1.35; margin-top:8px; }
        .documentos-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(245px,1fr)); gap:9px; }
        .documento-card { background:#FFFFFF; border:1px solid #DCE6EF; border-left:4px solid #AFC7DA; border-radius:10px; padding:10px 11px; min-height:96px; }
        .documento-card.apuracao-ok-card { border-left-color:#54A472; }
        .documento-card.apuracao-parcial-card { border-left-color:#D2A62A; }
        .documento-card.apuracao-bloqueado-card { border-left-color:#D48155; }
        @media (max-width:700px) { .apuracao-head { flex-direction:column; } .apuracao-progress { white-space:normal; } }
        </style>
        """,
        unsafe_allow_html=True,
    )


def _card(item: dict[str, Any]) -> str:
    estado = str(item.get("estado") or "nao_informado")
    classe, _ = _CLASSES.get(estado, _CLASSES["nao_informado"])
    nome = html.escape(str(item.get("nome") or "Bloco"))
    rotulo = html.escape(str(item.get("rotulo") or "Não informado"))
    detalhe = html.escape(str(item.get("detalhe") or ""))
    return (
        '<article class="apuracao-card">'
        '<div class="apuracao-card-top">'
        f'<span class="apuracao-name">{nome}</span>'
        f'<span class="apuracao-badge {classe}">{rotulo}</span>'
        '</div>'
        f'<div class="apuracao-detail">{detalhe}</div>'
        '</article>'
    )


def render_status_apuracao(capacidades: dict[str, Any]) -> None:
    _css()
    resumo = capacidades.get("resumo") or {}
    itens = list((capacidades.get("blocos") or {}).values()) + list((capacidades.get("calculos") or {}).values())
    completos = int(resumo.get("completos") or 0)
    total = len(itens)
    cards = "".join(_card(item) for item in itens)
    st.markdown(
        '<section class="apuracao-shell" aria-label="Status da Apuração">'
        '<div class="apuracao-head"><div>'
        '<h2 class="apuracao-title">Status da Apuração</h2>'
        '<div class="apuracao-subtitle">Cada bloco avança de forma independente. Ausências afetam somente os resultados que dependem delas.</div>'
        '</div>'
        f'<div class="apuracao-progress">{completos} de {total} blocos concluídos</div></div>'
        f'<div class="apuracao-grid">{cards}</div>'
        '</section>',
        unsafe_allow_html=True,
    )


def render_status_documentos(
    capacidades: dict[str, Any],
    chaves: Iterable[str] | None = None,
) -> None:
    _css()
    documentos = capacidades.get("documentos") or {}
    ordem = list(chaves or documentos.keys())
    cards = []
    for chave in ordem:
        item = documentos.get(chave)
        if not item:
            continue
        estado = str(item.get("estado") or "nao_informado")
        classe, _ = _CLASSES.get(estado, _CLASSES["nao_informado"])
        nome = html.escape(str(item.get("nome") or "Documento"))
        rotulo = html.escape(str(item.get("rotulo") or "Não informado"))
        motivo = html.escape(str(item.get("motivo") or ""))
        cards.append(
            f'<article class="documento-card {classe}-card">'
            '<div class="apuracao-card-top">'
            f'<span class="apuracao-name">{nome}</span>'
            f'<span class="apuracao-badge {classe}">{rotulo}</span>'
            '</div>'
            f'<div class="apuracao-detail">{motivo}</div>'
            '</article>'
        )
    st.markdown(
        '<section class="apuracao-shell" aria-label="Documentos da apuração">'
        '<div class="apuracao-head"><div>'
        '<h2 class="apuracao-title">Documentos da Apuração</h2>'
        '<div class="apuracao-subtitle">Todos permanecem visíveis; o estado informa o que já pode ser utilizado e o que ainda depende de complementação.</div>'
        '</div></div>'
        f'<div class="documentos-grid">{"".join(cards)}</div>'
        '</section>',
        unsafe_allow_html=True,
    )


def render_resultados_progressivos(resultado: dict[str, Any]) -> None:
    """Apresenta apenas valores explicitamente liberados pelo motor."""

    resultados = resultado.get("resultados_progressivos") or {}
    disponiveis = [item for item in resultados.values() if isinstance(item, dict) and item.get("disponivel")]
    if not disponiveis:
        st.info("Os blocos preenchidos foram reconhecidos. Recalcule e salve o XLS no Excel para gravar os valores dependentes das fórmulas.")
        return

    st.subheader("Resultados disponíveis")
    colunas = st.columns(min(3, len(disponiveis)))
    indice = 0
    for item in disponiveis:
        valor = item.get("valor")
        if valor is None:
            continue
        try:
            valor_formatado = f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except (TypeError, ValueError):
            valor_formatado = str(valor)
        with colunas[indice % len(colunas)]:
            st.metric(item.get("nome") or "Resultado", valor_formatado)
            st.caption(f"{item.get('rotulo')}. Origem: {item.get('origem') or 'XLS'}.")
        indice += 1
