"""Hotfix global do cabecalho (render_cabecalho_pagina).

Testa o HTML REAL emitido pelo helper — nao apenas a presenca/ausencia de
strings no codigo-fonte. Prova o contrato global:

  cabecalho = marca Cl8us + titulo + aviso de privacidade
  (sem qualquer frase descritiva; descricao recebida e ignorada)

E prova que o bloco HTML e integro: o <div> de privacidade permanece DENTRO
da <section>, sem linha em branco intermediaria capaz de fazer o CommonMark do
Streamlit fechar o bloco e renderizar tags como bloco de codigo (bug pos-reboot).
"""
from __future__ import annotations


def _captura_markdown(func, *args, **kwargs):
    import _ui_utils
    capturado = []
    orig = _ui_utils.st.markdown
    _ui_utils.st.markdown = lambda *a, **k: capturado.append(a[0] if a else "")
    try:
        func(*args, **kwargs)
    finally:
        _ui_utils.st.markdown = orig
    return "\n".join(str(x) for x in capturado)


def _assere_estrutura_base(html):
    assert "cl8us-page-header" in html
    assert "cl8us-brand-img" in html
    assert "<h1>Teste</h1>" in html
    assert "cl8us-page-privacy" in html
    assert "Use apenas para documentos não sigilosos e de livre acesso." in html
    # sem paragrafo de descricao em hipotese alguma
    assert "<p>" not in html and "<p></p>" not in html


def _assere_bloco_integro(html):
    """O privacy DIV nao pode sair do bloco: sem linha em branco entre o inicio
    da <section> e </section>, e sem indentacao de 4+ espacos (bloco de codigo)."""
    inicio = html.index("<section")
    fim = html.index("</section>")
    miolo = html[inicio:fim]
    # nenhuma linha vazia dentro da section (nem so-espacos)
    for linha in miolo.splitlines()[1:]:
        assert linha.strip() != "", "linha em branco dentro do bloco HTML do cabecalho"
    # o privacy DIV esta contido dentro da section
    assert html.index("cl8us-page-privacy") < fim
    # nenhuma linha do bloco inicia com 4+ espacos (evita indented code block)
    for linha in miolo.splitlines():
        assert not linha.startswith("    "), "indentacao capaz de virar bloco de codigo"


def test_descricao_vazia_produz_bloco_integro():
    import _ui_utils
    html = _captura_markdown(_ui_utils.render_cabecalho_pagina, "Teste", "")
    _assere_estrutura_base(html)
    _assere_bloco_integro(html)


def test_descricao_nao_vazia_nao_aparece():
    import _ui_utils
    html = _captura_markdown(
        _ui_utils.render_cabecalho_pagina, "Teste", "DESCRICAO QUE NAO DEVE APARECER"
    )
    _assere_estrutura_base(html)
    _assere_bloco_integro(html)
    assert "DESCRICAO QUE NAO DEVE APARECER" not in html


def test_default_sem_descricao():
    import _ui_utils
    html = _captura_markdown(_ui_utils.render_cabecalho_pagina, "Teste")
    _assere_estrutura_base(html)
    _assere_bloco_integro(html)
