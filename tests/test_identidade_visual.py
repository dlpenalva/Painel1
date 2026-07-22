"""Hotfix da identidade visual: prova que o percurso REAL da homepage
(render_cabecalho_pagina) exibe a marca em imagem e nao a marca textual antiga,
e que o favicon esta configurado nas paginas.
"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


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


def test_assets_existem():
    assert (ROOT / "assets" / "cl8us_header_proporcional.png").is_file()
    assert (ROOT / "assets" / "cl8us_favicon_512.png").is_file()


def test_render_cabecalho_pagina_usa_imagem_sem_marca_textual():
    import _ui_utils
    html = _captura_markdown(_ui_utils.render_cabecalho_pagina, "Titulo Funcional", "Descricao Funcional")
    assert "data:image/png;base64," in html      # header oficial no percurso da homepage
    assert "cl8us-brand-img" in html
    assert "min(420px" in html and "100%)" in html   # dimensionamento responsivo
    assert "<strong>TLB</strong>" not in html    # marca textual antiga removida
    assert "Titulo Funcional" in html and "Descricao Funcional" in html  # funcional preservado


def test_render_marca_topo_usa_imagem():
    import _ui_utils
    html = _captura_markdown(_ui_utils.render_marca_topo)
    assert "data:image/png;base64," in html
    assert '<span class="tlb-cl8us-name">' not in html   # marca textual antiga removida


def test_homepage_percorre_render_cabecalho_pagina():
    src = (ROOT / "pages" / "00_Calculadora_Reajustes.py").read_text(encoding="utf-8")
    assert "render_cabecalho_pagina(" in src   # a homepage usa a funcao corrigida


def test_favicon_configurado_no_entrypoint_e_paginas():
    app = (ROOT / "app.py").read_text(encoding="utf-8")
    assert "cl8us_favicon_512.png" in app and "page_icon=" in app
    for pg in (ROOT / "pages").glob("*.py"):
        s = pg.read_text(encoding="utf-8")
        if "set_page_config" in s:
            assert "cl8us_favicon_512.png" in s, f"{pg.name} sem favicon"


def test_data_uri_do_header_nao_vazio():
    import _ui_utils
    uri = _ui_utils._header_data_uri()
    assert uri.startswith("data:image/png;base64,") and len(uri) > 1000
