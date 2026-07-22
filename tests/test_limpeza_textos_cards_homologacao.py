"""Limpeza de textos e cards da homologacao (itens 4/5/6/7).

Prova, sem tocar em matematica/dados:
  * as quatro descricoes de cabecalho nao sao mais renderizadas (helper omite o
    <p> quando a descricao e vazia; call-sites passam "");
  * titulos funcionais e a marca Cl8us permanecem;
  * o box "O Arquivo Coleta Oficial reune..." foi removido de pages/03;
  * os seis cards documentais continuam presentes, com titulos e acoes, porem
    sem a legenda descritiva interna (motivo) e com regra de altura uniforme;
  * a deteccao de calendario invalido (flag/logica) continua existindo;
  * a Adequacao (page 12) nao teve alteracao funcional alem do cabecalho.
"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

INICIO = (ROOT / "pages" / "00_Calculadora_Reajustes.py").read_text(encoding="utf-8")
SIMPLES = (ROOT / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
DOCUMENTOS = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")
ADEQUACAO = (ROOT / "pages" / "12_Adequacao_Orcamentaria.py").read_text(encoding="utf-8")
UI = (ROOT / "_ui_utils.py").read_text(encoding="utf-8")


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


# ---------------------------------------------------------------- item 4 + C/N
def test_helper_omite_paragrafo_quando_descricao_vazia():
    import _ui_utils
    html = _captura_markdown(_ui_utils.render_cabecalho_pagina, "So Titulo", "")
    assert "<p>" not in html and "<p></p>" not in html   # sem paragrafo/vazio
    assert "So Titulo" in html                            # titulo preservado
    assert "cl8us-page-header" in html                    # estrutura preservada
    assert "cl8us-brand-img" in html                      # marca Cl8us intacta
    assert "data:image/png;base64," in html               # header oficial
    assert "documentos" in html and "acesso" in html      # aviso de privacidade


def test_helper_ainda_renderiza_paragrafo_quando_ha_descricao():
    import _ui_utils
    html = _captura_markdown(_ui_utils.render_cabecalho_pagina, "T", "Descricao real")
    assert "<p>Descricao real</p>" in html                # retrocompativel


def test_helper_trata_espacos_como_vazio():
    import _ui_utils
    html = _captura_markdown(_ui_utils.render_cabecalho_pagina, "T", "   ")
    assert "<p>" not in html


# ------------------------------------------------------------ item 4 (A) + (B)
def test_quatro_descricoes_removidas_dos_call_sites():
    assert "Baixe o XLS Coleta, registre as informações do processo" not in INICIO
    assert "Ferramenta para análise contratual de um único ciclo" not in SIMPLES
    assert "Envie o Arquivo Coleta Oficial preenchido para validar cada bloco" not in DOCUMENTOS
    assert "Estimativa simplificada do delta orçamentário" not in ADEQUACAO


def test_titulos_funcionais_permanecem():
    assert '"Reajustes contratuais"' in INICIO
    assert '"Calculadora 1 ciclo"' in SIMPLES
    assert '"Painel da Apuração Contratual"' in DOCUMENTOS
    assert '"Adequação Orçamentária"' in ADEQUACAO


def test_call_sites_passam_descricao_vazia():
    for src in (INICIO, SIMPLES, DOCUMENTOS, ADEQUACAO):
        assert "render_cabecalho_pagina(" in src
    # helper com default retrocompativel
    assert 'def render_cabecalho_pagina(titulo, descricao="")' in UI


# ------------------------------------------------------------------- item 5 (D)
def test_box_coleta_removido():
    assert "cl8us-docs-note" not in DOCUMENTOS
    assert "O Arquivo Coleta Oficial reúne os dados da apuração" not in DOCUMENTOS


# --------------------------------------------------------- item 6/7 (E-I, L, H)
def test_seis_cards_continuam_presentes():
    from _capacidades_apuracao import SEIS_DOCUMENTOS_CANONICOS
    assert len(SEIS_DOCUMENTOS_CANONICOS) == 6
    assert "def render_documentos_funcionais_upload(resultado):" in DOCUMENTOS


def test_titulos_e_acoes_dos_seis_cards_permanecem():
    from _capacidades_apuracao import SEIS_DOCUMENTOS_CANONICOS
    # Os titulos vem da tupla canonica (fonte da verdade) e sao renderizados
    # dinamicamente em pages/03 via `#### {titulo}` — nao como literais na pagina.
    titulos = {titulo for _chave, titulo in SEIS_DOCUMENTOS_CANONICOS}
    assert titulos == {
        "Sumário Executivo", "Adequação Orçamentária", "Despacho Saneador",
        "Termo de Apostila", "Garantia Contratual", "DOU",
    }
    assert 'st.markdown(f"#### {titulo}")' in DOCUMENTOS   # titulo renderizado no card
    for marcador in (
        "upload_docs_sumario_executivo",
        "Abrir Adequação Orçamentária",
        "Despacho_Saneador_Instrucao_Processual.docx",
        "Termo_de_Apostila_Reajuste_Contratual.docx",
        "Abrir Garantia Contratual",
        "Abrir DOU",
    ):
        assert marcador in DOCUMENTOS


def test_legenda_de_motivo_removida_dos_cards():
    # item 6: sem a frase descritiva interna; bloqueio funcional (botao) preservado.
    assert 'st.caption(documento.get("motivo")' not in DOCUMENTOS
    assert "Complete os dados necessários para liberar este documento." not in DOCUMENTOS
    assert 'key=f"upload_docs_{chave}_pendencia"' in DOCUMENTOS   # botao desabilitado mantido


def test_seis_cards_regra_de_altura_uniforme():
    # item 7: marcador emitido em cada card + regra de min-height uniforme + acao na base.
    assert '<span class="upload-doc-card"></span>' in DOCUMENTOS
    assert "min-height:7rem" in DOCUMENTOS
    assert "margin-top:auto" in DOCUMENTOS
    assert ".upload-doc-card { display:none; }" in DOCUMENTOS


# -------------------------------------------------------------- item 8 (K) flag
def test_deteccao_calendario_invalido_preservada():
    painel = (ROOT / "_painel_executivo.py").read_text(encoding="utf-8")
    assistente = (ROOT / "_assistente_fiscal.py").read_text(encoding="utf-8")
    assert "LINHA_TEMPORAL_INVALIDA" in painel
    assert "LINHA_TEMPORAL_INVALIDA" in assistente


# ---------------------------------------------------------------- item 10/M/N
def test_adequacao_sem_alteracao_funcional():
    # page 12 mantem o motor normativo e a saida XLSX; so o cabecalho mudou.
    assert "from _adequacao_orcamentaria import" in ADEQUACAO
    assert "gerar_xlsx_projecao(" in ADEQUACAO
    assert '"Baixar XLSX"' in ADEQUACAO


def test_favicon_e_marca_intactos():
    assert "cl8us_favicon_512.png" in DOCUMENTOS and "cl8us_favicon_512.png" in ADEQUACAO
    assert (ROOT / "assets" / "cl8us_header_proporcional.png").is_file()
    assert (ROOT / "assets" / "cl8us_favicon_512.png").is_file()
