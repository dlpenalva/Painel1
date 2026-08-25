"""Protecao do redesign clean da pagina Inicio.

Cobre os pontos de validacao da entrega: a HOME abre sem excecao, usa o logo
oficial existente, perdeu os cinco boxes e a secao introdutoria, mostra as tres
etapas com pictogramas vetoriais, traz o bloco da Coleta com o MESMO download de
antes e o convite ao menu lateral — sem repetir a navegacao da sidebar.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

INICIO_PATH = ROOT / "pages" / "00_Calculadora_Reajustes.py"
INICIO = INICIO_PATH.read_text(encoding="utf-8")
APP = (ROOT / "app.py").read_text(encoding="utf-8")
UI = (ROOT / "_ui_utils.py").read_text(encoding="utf-8")


def _home():
    from streamlit.testing.v1 import AppTest

    at = AppTest.from_file(str(INICIO_PATH), default_timeout=180)
    at.run()
    return at


def _blob(at) -> str:
    return "\n".join(str(m.value) for m in at.markdown)


# ---------------------------------------------------------------------------
# 1 / 9 — a pagina abre e o download continua sendo o mesmo
# ---------------------------------------------------------------------------

def test_home_abre_sem_excecao():
    at = _home()
    assert not at.exception, at.exception
    assert not list(at.error)


def test_download_da_coleta_permanece_a_mesma_origem():
    at = _home()
    botoes = list(at.get("download_button"))
    assert len(botoes) == 1
    assert botoes[0].label == "Baixar Arquivo Coleta Oficial"
    # Nenhuma geracao nova de XLSX: a HOME segue chamando a fonte unica.
    assert "gerar_coleta_oficial_preenchida" in INICIO
    assert "assinatura_template_coleta" in INICIO
    assert "file_name=nome_download_coleta()" in INICIO
    assert 'key="download_coleta_inicio"' in INICIO
    assert "TEMPLATE_COLETA_OFICIAL.exists()" in INICIO


# ---------------------------------------------------------------------------
# 2 / 3 — logo oficial existente, sem recriacao e sem segundo branding
# ---------------------------------------------------------------------------

def test_hero_usa_o_asset_oficial_do_projeto():
    assert "_header_data_uri" in INICIO
    assert "cl8us_header_proporcional.png" in UI
    at = _home()
    assert "data:image/png;base64," in _blob(at)


def test_hero_preserva_titulo_e_aviso():
    at = _home()
    blob = _blob(at)
    assert "Reajustes contratuais" in blob
    assert "Use apenas para documentos não sigilosos e de livre acesso." in blob
    assert "cl8us-page-privacy" in blob


def test_sidebar_e_seu_branding_permanecem_intactos():
    # A HOME nao desenha branding proprio na lateral nem toca na sidebar.
    assert "cl8us-side-brand" in APP
    assert "cl8us-side-brand" not in INICIO
    assert "st.sidebar" not in INICIO
    assert "render_versao_sidebar" in APP


# ---------------------------------------------------------------------------
# 4 / 5 — o que saiu
# ---------------------------------------------------------------------------

def test_os_cinco_boxes_antigos_sairam():
    for numero in range(1, 6):
        assert f'"{numero} ·' not in INICIO
    assert "home-card" not in INICIO
    assert "_conteudo_card" not in INICIO
    assert "st.switch_page(" not in INICIO


def test_secao_introdutoria_saiu():
    for texto in (
        "Como funciona",
        "Fluxo operacional",
        "Preparar os marcos da coleta",
        "O Arquivo Coleta Oficial é o produto principal",
    ):
        assert texto not in INICIO
    at = _home()
    assert not list(at.subheader)


# ---------------------------------------------------------------------------
# 6 / 7 — as tres etapas, os pictogramas e a ilustracao
# ---------------------------------------------------------------------------

def test_tres_etapas_com_textos_aprovados():
    blob = _blob(_home())
    assert "1. PREPARE" in blob
    assert "Defina as datas, os ciclos e gere o Arquivo Coleta Oficial." in blob
    assert "2. PREENCHA" in blob
    assert "complementar as informações necessárias." in blob
    assert "3. APURE" in blob
    assert "Envie o arquivo completo e obtenha resultados e documentos." in blob


def test_fluxo_e_puramente_explicativo():
    at = _home()
    # Sem botoes e sem links entre as etapas: navegacao so na sidebar.
    assert not list(at.button)
    assert "home-fluxo" in _blob(at)


def test_pictogramas_sao_svg_inline_sem_dependencia_nem_emoji():
    blob = _blob(_home())
    # 3 etapas + 2 setas + lampada + ilustracao da Coleta.
    assert blob.count("<svg") == 7
    assert blob.count('class="home-seta"') == 2
    assert "home-ico" in blob
    # Nada externo e nada de emoji como substituto visual.
    assert "http://" not in blob and "https://" not in blob
    # Sem recurso externo e sem biblioteca nova declarada na propria pagina.
    assert "http" not in INICIO
    assert "import" not in INICIO.split("MIME_XLSX")[1]
    assert not any(ord(c) > 0x2100 for c in INICIO)


def test_pictogramas_seguem_a_identidade_visual():
    # Traco azul da marca sobre disco azul muito claro.
    assert "#1F5F8B" in INICIO
    assert "#E9F2F9" in INICIO
    assert "border-radius:50%" in INICIO


def test_ilustracao_da_coleta_tem_pilha_folha_e_escudo():
    assert "ILUSTRACAO_COLETA" in INICIO
    blob = _blob(_home())
    assert blob.count("<rect") == 2          # pilha de documentos
    assert 'viewBox="0 0 122 92"' in blob    # folha principal + escudo
    assert "#E9F2F9" in blob                 # preenchimento do escudo


# ---------------------------------------------------------------------------
# Bloco da Coleta e convite ao menu lateral
# ---------------------------------------------------------------------------

def test_bloco_da_coleta_e_o_convite_ao_menu():
    blob = _blob(_home())
    assert "Arquivo Coleta Oficial" in blob
    assert (
        "É o arquivo central da apuração e acompanha o processo da "
        "preparação até os resultados." in blob
    )
    assert "Use o menu lateral para acessar as funcionalidades do cl8us." in blob
    assert "home-dica-corpo" in blob


def test_botao_da_coleta_usa_a_cor_de_acao_da_marca():
    assert 'type="primary"' in INICIO
    assert "--cl8us-action: #7A1733;" in APP


# ---------------------------------------------------------------------------
# 8 — responsividade e escopo do CSS
# ---------------------------------------------------------------------------

def test_css_e_responsivo_e_escopado_na_home():
    assert "@media (max-width:900px)" in INICIO
    assert ".home-fluxo{flex-direction:column;}" in INICIO
    # O botao real de download vive DENTRO do cartao da Coleta: a coluna e o cartao.
    assert '[data-testid="stColumn"]:has(.home-coleta-corpo)' in INICIO
    # Todo seletor do CSS injetado pela HOME e ancorado em `home-`: mesmo o
    # reuso da pilula global de privacidade entra escopado (.home-hero .cl8us-...).
    blob = _blob(_home())
    estilo = blob[blob.index("<style>") + len("<style>"):blob.index("</style>")]
    seletores = [s.strip() for s in re.findall(r"([^{};]+)\{", estilo)]
    seletores = [s for s in seletores if s and not s.startswith("@media")]
    assert seletores, "nenhum seletor encontrado"
    assert all("home-" in s for s in seletores), seletores


def test_home_nao_altera_css_global_do_app():
    # O redesign nao acrescentou regra da home ao CSS compartilhado do app.py.
    assert "home-hero" not in APP
    assert "home-etapa" not in APP
    assert "home-coleta-corpo" not in APP
    assert "home-dica-corpo" not in APP
