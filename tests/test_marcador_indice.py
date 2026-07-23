"""Marcadores decorativos dos dois titulos laterais padronizados com 🔹.

- "Data-base/âncora inicial da análise atual: 🔹"
- "Índice do contrato 🔹"

O antigo circulo CSS (.cl8us-index-title::after) foi removido. Controles nativos
(radio/selectbox/BaseWeb) permanecem intactos.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / "app.py").read_text(encoding="utf-8")
UI = (ROOT / "_ui_utils.py").read_text(encoding="utf-8")
MULTI = (ROOT / "pages" / "02_Calculo_Represados.py").read_text(encoding="utf-8")

MARCADOR = "\U0001F539"  # 🔹 small blue diamond


def test_data_base_tem_marcador_no_final():
    assert f"Data-base/âncora inicial da análise atual: {MARCADOR}" in MULTI


def test_indice_do_contrato_tem_marcador_no_final():
    assert f'<div class="cl8us-index-title">Índice do contrato {MARCADOR}</div>' in UI


def test_marcadores_usam_o_mesmo_caractere():
    # Ambos os titulos usam exatamente o mesmo caractere 🔹.
    assert MULTI.count(MARCADOR) == 1
    assert UI.count(MARCADOR) == 1


def test_circulo_css_antigo_removido():
    assert ".cl8us-index-title::after" not in APP
    # Nenhum residuo do marcador circular preto do indice.
    assert "#1A1A1A" not in APP


def test_indice_title_sem_marcador_circular():
    # O bloco .cl8us-index-title (sem ::after) nao deve desenhar circulo.
    ini = APP.index(".cl8us-index-title {")
    bloco = APP[ini: APP.index("}", ini)]
    assert "border-radius" not in bloco
    assert "width" not in bloco
    assert "height" not in bloco


def test_marker_tecnico_invisivel_preservado():
    assert ".cl8us-index-marker { display: none; }" in APP
    assert ":has(.cl8us-index-marker)" in APP


def test_nao_usa_outros_simbolos_proibidos():
    for simbolo in ("●", "•", "\U0001F4CC", "\U0001F4CD", "\U0001F538",
                    "▫️", "▪️"):
        assert f"Índice do contrato {simbolo}" not in UI
        assert f"análise atual: {simbolo}" not in MULTI


def test_controles_nativos_nao_alterados():
    # A estilizacao existente do radio lateral (apenas borda/realce do label)
    # permanece; nenhuma regra escondendo o marcador nativo foi adicionada.
    assert '[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked)' in APP
    assert '[role="radio"]' not in APP           # nao mexemos no marcador nativo do radio
    assert "stRadio\"] [role=\"radio\"]" not in APP


def _bloco_date_input_data_base():
    # Extrai o corpo da chamada st.date_input do campo Data-base, do rotulo ate
    # o primeiro parentese de fechamento.
    rotulo = f"Data-base/âncora inicial da análise atual: {MARCADOR}"
    ini = MULTI.index(rotulo)
    fim = MULTI.index(")", ini)
    return MULTI[ini:fim]


def test_date_input_data_base_sem_help():
    # O icone de ajuda nativo (parametro help=) foi removido do date_input da
    # Data-base; a causa raiz era esse help=, nao CSS/pseudo-elemento.
    bloco = _bloco_date_input_data_base()
    assert "help=" not in bloco
    assert "Informe a data-base original do contrato" not in MULTI


def test_selectbox_indice_sem_help():
    # O selectbox do Indice do contrato nao possui mais o parametro help=.
    assert "help=" not in UI
    assert "Confirme o índice contratual" not in UI


def test_nenhuma_regra_css_para_ocultar_icones():
    # A correcao foi pela causa (remocao do help=), nao por CSS escondendo o
    # icone nativo de ajuda/tooltip do Streamlit.
    for fonte in (APP, UI, MULTI):
        assert "stTooltip" not in fonte
        assert "TooltipIcon" not in fonte
        assert "TooltipHoverTarget" not in fonte
        assert "data-baseweb=\"tooltip\"" not in fonte
