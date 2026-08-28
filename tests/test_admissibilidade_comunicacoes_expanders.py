"""ADM-UX-COM2: as duas comunicacoes da admissibilidade como expanders irmaos.

Cobre a unificacao visual (box amarelo -> expander) e blinda o que NAO pode
mudar: assunto, corpo e bytes do TXT da Contratada permanecem os mesmos.
"""

from __future__ import annotations

import hashlib
from contextlib import contextmanager
from pathlib import Path

import pytest

import _email_contratada as mod_email
import _solicitacao_fiscal as mod_fiscal
from _email_contratada import (
    ASSUNTO_EMAIL_CONTRATADA,
    COMANDO_EMAIL_CONTRATADA,
    gerar_rascunho_email_contratada,
    montar_texto_comunicacao,
    montar_txt_download,
)
from _solicitacao_fiscal import (
    COMANDO_SOLICITACAO_FISCAL,
    TEXTO_SOLICITACAO_FISCAL,
    sincronizar_texto_solicitacao_fiscal,
)


ROOT = Path(__file__).resolve().parents[1]

TITULO_EXPANDER_1 = "Comunicar à contratada o resultado da admissibilidade"
TITULO_EXPANDER_2 = "Solicitar ao fiscal o preenchimento da planilha Coleta"

CICLO_PADRAO = [
    {
        "ciclo": "C1",
        "situacao_aplicada": "Tempestivo",
        "variacao_formatada": "3,27%",
        "financeiro_inicio": "13/02/2026",
    }
]


class _StFake:
    """Coletor minimo do Streamlit usado pelos dois renderizadores."""

    def __init__(self) -> None:
        self.chamadas: list[tuple] = []
        self.session_state: dict = {}

    @contextmanager
    def expander(self, label, expanded=True):
        self.chamadas.append(("expander", label, expanded))
        yield self

    def caption(self, texto, **kwargs):
        self.chamadas.append(("caption", texto, kwargs))

    def markdown(self, texto, **kwargs):
        self.chamadas.append(("markdown", texto, kwargs))

    def error(self, texto, **kwargs):
        self.chamadas.append(("error", texto, kwargs))

    def download_button(self, label, **kwargs):
        self.chamadas.append(("download_button", label, kwargs))

    def text_area(self, label, **kwargs):
        self.chamadas.append(("text_area", label, kwargs))

    def _tipos(self) -> list[str]:
        return [c[0] for c in self.chamadas]

    def _primeira(self, tipo: str) -> tuple:
        for chamada in self.chamadas:
            if chamada[0] == tipo:
                return chamada
        raise AssertionError(f"nenhuma chamada de {tipo}")


@pytest.fixture()
def st_email(monkeypatch) -> _StFake:
    fake = _StFake()
    monkeypatch.setattr(mod_email, "st", fake)
    return fake


@pytest.fixture()
def st_fiscal(monkeypatch) -> _StFake:
    fake = _StFake()
    monkeypatch.setattr(mod_fiscal, "st", fake)
    return fake


# 1/3: expander 1 com titulo exato, comecando recolhido.
def test_expander_1_titulo_exato_e_recolhido(st_email):
    mod_email.render_email_contratada(CICLO_PADRAO, key="k1")
    _tipo, label, expanded = st_email._primeira("expander")
    assert label == TITULO_EXPANDER_1
    assert expanded is False
    assert COMANDO_EMAIL_CONTRATADA == TITULO_EXPANDER_1


# 2/3: expander 2 com titulo exato, comecando recolhido.
def test_expander_2_titulo_exato_e_recolhido(st_fiscal):
    mod_fiscal.render_solicitacao_fiscal_coleta(("simples", "assinatura"))
    _tipo, label, expanded = st_fiscal._primeira("expander")
    assert label == TITULO_EXPANDER_2
    assert expanded is False
    assert COMANDO_SOLICITACAO_FISCAL == TITULO_EXPANDER_2


# 4: o box amarelo nao e mais renderizado.
def test_box_amarelo_nao_e_mais_renderizado(st_email):
    mod_email.render_email_contratada(CICLO_PADRAO, key="k1")
    visuais = [c for c in st_email.chamadas if c[0] in ("markdown", "caption")]
    assert visuais
    for _tipo, conteudo, _kwargs in visuais:
        assert "background" not in conteudo
        assert "border" not in conteudo
        assert "COMUNICAÇÃO À CONTRATADA" not in conteudo
    fonte = (ROOT / "_email_contratada.py").read_text(encoding="utf-8")
    assert "FFF9E8" not in fonte
    assert "E8B923" not in fonte


# 5/6/7: assunto e corpo intactos, agora num textarea editavel unico.
def test_assunto_e_corpo_no_textarea_editavel(st_email):
    assunto_ref, corpo_ref = gerar_rascunho_email_contratada(
        CICLO_PADRAO, "CT-99/2026", "ICTI (Ipeadata)", fator_acumulado=1.0327
    )
    mod_email.render_email_contratada(
        CICLO_PADRAO,
        numero_contrato="CT-99/2026",
        indice="ICTI (Ipeadata)",
        fator_acumulado=1.0327,
        assinatura_analise=("simples", "a"),
        key="txt_x",
    )
    assert assunto_ref == ASSUNTO_EMAIL_CONTRATADA
    assert st_email._tipos() == ["markdown", "expander", "caption", "text_area"]
    _tipo, _label, kwargs = st_email._primeira("text_area")
    assert kwargs["key"] == "txt_x"
    assert kwargs.get("disabled") in (None, False)
    texto = st_email.session_state["txt_x"]
    assert texto == montar_texto_comunicacao(assunto_ref, corpo_ref)
    assert texto.startswith(f"Assunto: {ASSUNTO_EMAIL_CONTRATADA}\n\nPrezados,")
    assert corpo_ref in texto
    # Os bytes do TXT continuam derivaveis do mesmo assunto/corpo.
    assert montar_txt_download(assunto_ref, corpo_ref).decode(
        "utf-8-sig"
    ) == f"ASSUNTO: {assunto_ref}\n\n{corpo_ref}"


# 2 (ajuste de homologacao): o botao de download saiu da interface.
def test_download_txt_nao_e_mais_renderizado(st_email):
    mod_email.render_email_contratada(CICLO_PADRAO, key="k1")
    assert "download_button" not in st_email._tipos()
    fonte = (ROOT / "_email_contratada.py").read_text(encoding="utf-8")
    assert "st.download_button" not in fonte
    assert "Baixar rascunho" not in fonte
    # O helper permanece disponivel (nao houve limpeza de codigo).
    assert callable(montar_txt_download)


# 1 (ajuste de homologacao): respiro antes da area das comunicacoes.
def test_espaco_antes_das_comunicacoes(st_email):
    mod_email.render_email_contratada(CICLO_PADRAO, key="k1")
    primeira = st_email.chamadas[0]
    assert primeira[0] == "markdown"
    assert primeira[1] == mod_email._ESPACO_ANTES_DAS_COMUNICACOES
    assert primeira[2].get("unsafe_allow_html") is True
    # Espacador puro: sem cor, sem borda, sem texto.
    assert primeira[1] == '<div style="height:0"></div>'


def test_caption_do_expander_1_e_curta(st_email):
    mod_email.render_email_contratada(CICLO_PADRAO, key="k1")
    _tipo, texto, _kwargs = st_email._primeira("caption")
    assert texto == "Revise o texto abaixo e copie-o para a comunicação à contratada."
    assert texto == mod_email.CAPTION_EMAIL_CONTRATADA


# 3/7 (ajuste de homologacao): os dois expanders seguem o mesmo padrao.
def test_os_dois_expanders_usam_o_mesmo_padrao(st_email, st_fiscal):
    mod_email.render_email_contratada(
        CICLO_PADRAO, assinatura_analise=("simples", "a"), key="txt_x"
    )
    mod_fiscal.render_solicitacao_fiscal_coleta(("simples", "a"))
    assert st_email._tipos()[1:] == st_fiscal._tipos()
    for fake in (st_email, st_fiscal):
        assert fake._primeira("expander")[2] is False
        assert fake._primeira("text_area")[2]["height"] == 460
        assert fake._primeira("text_area")[2]["label_visibility"] == "collapsed"


def test_edicao_da_contratada_persiste_e_reseta_em_nova_analise(st_email):
    mod_email.render_email_contratada(
        CICLO_PADRAO, assinatura_analise=("simples", "a"), key="txt_x"
    )
    base = st_email.session_state["txt_x"]
    st_email.session_state["txt_x"] = "Edição da contratada"
    mod_email.render_email_contratada(
        CICLO_PADRAO, assinatura_analise=("simples", "a"), key="txt_x"
    )
    assert st_email.session_state["txt_x"] == "Edição da contratada"
    mod_email.render_email_contratada(
        CICLO_PADRAO, assinatura_analise=("simples", "b"), key="txt_x"
    )
    assert st_email.session_state["txt_x"] == base


# 17: trava de regressao dos bytes do TXT da main homologada 44aa7be.
def test_bytes_do_txt_permanecem_os_da_main_homologada():
    _assunto, corpo = gerar_rascunho_email_contratada([], None, None)
    assert (
        hashlib.sha256(montar_txt_download(_assunto, corpo)).hexdigest()
        == "95f609a3ed0987117619499ec18122e957d0c1fd8ed66dc9fb1947f7040674af"
    )
    _assunto, corpo = gerar_rascunho_email_contratada(
        CICLO_PADRAO, "CT-99/2026", "ICTI (Ipeadata)"
    )
    assert (
        hashlib.sha256(montar_txt_download(_assunto, corpo)).hexdigest()
        == "1a4e6b0c424f741a87a67650bd7c3069a996f49819731e3dd502a05eca7d826d"
    )


# O fail-closed nao pode ficar escondido dentro de um expander recolhido.
def test_fail_closed_nao_esconde_o_erro_dentro_do_expander(st_email):
    mod_email.render_email_contratada(
        [
            {
                "ciclo": "C1",
                "situacao_aplicada": "Tempestivo",
                "variacao_formatada": "3,27%",
                "financeiro_inicio": "13/02/2026",
                "memoria_calculo": [{"tipo": "MES", "ordem": 1}],
            }
        ],
        key="k1",
    )
    assert st_email._tipos() == ["error"]


# 8/9/10: solicitacao ao fiscal editavel, persistente e resetavel.
def test_solicitacao_fiscal_continua_editavel(st_fiscal):
    mod_fiscal.render_solicitacao_fiscal_coleta(("simples", "a"))
    assert st_fiscal._tipos() == ["expander", "caption", "text_area"]
    _tipo, _label, kwargs = st_fiscal._primeira("text_area")
    assert kwargs["key"] == "solicitacao_fiscal_coleta_texto"
    assert kwargs.get("disabled") in (None, False)
    assert (
        st_fiscal.session_state["solicitacao_fiscal_coleta_texto"]
        == TEXTO_SOLICITACAO_FISCAL
    )


def test_session_state_preserva_edicao_e_reseta_em_nova_analise(st_fiscal):
    mod_fiscal.render_solicitacao_fiscal_coleta(("simples", "a"))
    st_fiscal.session_state["solicitacao_fiscal_coleta_texto"] = "Edição do fiscal"
    mod_fiscal.render_solicitacao_fiscal_coleta(("simples", "a"))
    assert (
        st_fiscal.session_state["solicitacao_fiscal_coleta_texto"]
        == "Edição do fiscal"
    )
    mod_fiscal.render_solicitacao_fiscal_coleta(("simples", "b"))
    assert (
        st_fiscal.session_state["solicitacao_fiscal_coleta_texto"]
        == TEXTO_SOLICITACAO_FISCAL
    )
    assert sincronizar_texto_solicitacao_fiscal({}, ("x", 1)) == (
        TEXTO_SOLICITACAO_FISCAL
    )


# 11: placeholder deliberado do contrato, so dentro do texto editavel.
def test_placeholder_do_contrato_exato_e_unico():
    assert TEXTO_SOLICITACAO_FISCAL.count("[CONTRATO: TLB-CTR-/]") == 1
    assert (
        "Recebemos o pedido de reajuste apresentado pela contratada no âmbito "
        "do [CONTRATO: TLB-CTR-/] e concluímos a etapa inicial de "
        "admissibilidade e conferência dos percentuais aplicáveis."
    ) in TEXTO_SOLICITACAO_FISCAL


# 12 a 15: trechos exatos da versao aprovada.
@pytest.mark.parametrize(
    "trecho",
    [
        "na data correspondente",
        "Também precisamos da posição atual do contrato",
        "Nessa aba, devem ser informados:",
        "- a data da posição atual; e",
        "- a quantidade que ainda resta de cada item nessa data.",
        "reunindo o que já foi executado",
        "Se tiver dúvidas, fico à disposição.",
    ],
)
def test_trechos_da_versao_aprovada(trecho):
    assert trecho in TEXTO_SOLICITACAO_FISCAL


@pytest.mark.parametrize(
    "removido",
    [
        "na posição correspondente",
        "uma fotografia atual do contrato",
        "o preenchimento é basicamente composto por",
        "reunindo de forma consistente",
    ],
)
def test_redacoes_antigas_removidas(removido):
    assert removido not in TEXTO_SOLICITACAO_FISCAL


def test_texto_termina_na_despedida_aprovada():
    assert TEXTO_SOLICITACAO_FISCAL.endswith(
        "Se tiver dúvidas, fico à disposição.\n\nAtenciosamente,"
    )


# 16/17: nenhuma acao virou gate e nenhum formulario novo foi criado.
def test_nenhum_formulario_de_identificacao_foi_criado():
    for arquivo in ("_solicitacao_fiscal.py", "_email_contratada.py"):
        fonte = (ROOT / arquivo).read_text(encoding="utf-8")
        for proibido in (
            "st.text_input",
            "st.selectbox",
            "st.form(",
            "st.number_input",
            "st.date_input",
            "st.stop()",
        ):
            assert proibido not in fonte


def test_paginas_mantem_a_ordem_e_nao_ganharam_logica_nova():
    for pagina in ("01_Calculo_Simples.py", "02_Calculo_Represados.py"):
        fonte = (ROOT / "pages" / pagina).read_text(encoding="utf-8")
        assert fonte.rindex("render_email_contratada(") < fonte.rindex(
            "render_solicitacao_fiscal_coleta("
        )
        cauda = fonte[fonte.rindex("render_email_contratada(") :]
        assert "st.expander" not in cauda
        assert "assinatura_analise=(" in cauda


def test_paginas_reusam_a_assinatura_de_analise_ja_existente():
    """A contratada reseta na MESMA troca de analise que o fiscal."""
    simples = (ROOT / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
    multiplos = (ROOT / "pages" / "02_Calculo_Represados.py").read_text(
        encoding="utf-8"
    )
    assert '"chave_analise_simples_processada"' in simples
    assert simples.count('"chave_analise_simples_processada"') >= 2
    assert '"processar_reajustes_multiplos_key"' in multiplos
    assert multiplos.count('"processar_reajustes_multiplos_key"') >= 2
    assert 'key="txt_comunicacao_contratada_simples_v1"' in simples
    assert 'key="txt_comunicacao_contratada_multiciclo"' in multiplos
