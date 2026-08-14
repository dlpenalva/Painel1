# -*- coding: utf-8 -*-
"""Etapa 51D — capitulo 5 do Sumario Executivo: Item | C0..C(ciclo vigente).

Protege o novo contrato do capitulo "5. Itens e valores atualizados":
uma unica coluna monetaria por ciclo (VALOR UNITARIO de historico_VU),
sem quantidade e sem total; limite = CICLO VIGENTE em execucao (preclusos
continuam na historia; ciclo futuro nunca aparece); item de aditivo mostra
"—" antes do nascimento e ausencia de VU nunca vira R$ 0,00. Tambem prova
a exclusao da frase de consolidacao do cabecalho.
"""
from datetime import date

import pytest

from _sumario_executivo import (
    NAO_INFORMADO,
    _montar_secao_historico_vu,
    gerar_sumario_executivo,
    gerar_sumario_executivo_pdf,
)
from test_sumario_executivo import (
    _ciclo_param,
    _texto_pdf,
    leitura_simples_financeiro,
)

FRASE_PROIBIDA = (
    "Documento executivo consolidado a partir do objeto canônico do "
    "processo e do XLS processado. Nenhum valor é recalculado."
)


def _ciclos_sec(analisados):
    return [{"ciclo": f"C{i}", "computar": ("Sim" if i in analisados else "Não")}
            for i in range(5)]


def _hist(itens):
    return {"itens": itens}


def _dados_pdf(hist_secao, itens_legado=None):
    """Dados minimos para o renderer (capitulo 5 novo)."""
    return {
        "disponivel": True,
        "identificacao": {"indice": "IST", "metodo": "PCs", "ciclo_vigente": "C3",
                          "data_corte": "30/09/2026", "gerado_em": "13/08/2026"},
        "sintese": {"metodo_vta": "PCs", "vta": 1.0,
                    "variacao_acumulada": 0.01, "retroativo_total": 1.0},
        "ciclos": [], "financeiro": {},
        "itens": itens_legado or [],
        "historico_vu": hist_secao,
        "memoria_calculo": [], "aditivos": {}, "observacoes": [],
        "campos_nao_confiaveis": [], "referencias_vta": {},
    }


# ------------------------------------------------- 1. frase excluida

def test_frase_do_cabecalho_excluida_do_pdf():
    pdf = gerar_sumario_executivo(leitura_simples_financeiro())
    texto = _texto_pdf(pdf)
    assert "Documento executivo consolidado" not in texto
    assert FRASE_PROIBIDA.split(".")[0] not in texto
    assert "Sumário Executivo — Reajuste Contratual" in texto
    assert "1. Identificação" in texto


# ------------------------------------------------- 2-6. limite por ciclo vigente

@pytest.mark.parametrize("vigente,esperado", [
    ("C1", ["C0", "C1"]),
    ("C2", ["C0", "C1", "C2"]),
    ("C3", ["C0", "C1", "C2", "C3"]),
    ("C4", ["C0", "C1", "C2", "C3", "C4"]),
])
def test_colunas_ate_o_ciclo_vigente(vigente, esperado):
    secao = _montar_secao_historico_vu(
        _hist([{"item": "1", "vu_ciclos": {"VU_C0": 10.0}}]),
        _ciclos_sec(analisados=set()),
        vigente,
    )
    assert secao["ciclos"] == esperado
    assert secao["ultimo_ciclo"] == vigente


def test_ciclo_futuro_nunca_aparece_mesmo_com_vu_registrado():
    # VU_C4 fisicamente presente na aba, mas ciclo vigente C3.
    secao = _montar_secao_historico_vu(
        _hist([{"item": "1", "vu_ciclos": {
            "VU_C0": 10.0, "VU_C1": 11.0, "VU_C2": 12.0,
            "VU_C3": 13.0, "VU_C4": 14.0}}]),
        _ciclos_sec(analisados={1, 2, 3, 4}),
        "C3",
    )
    assert secao["ciclos"] == ["C0", "C1", "C2", "C3"]
    assert "C4" not in secao["itens"][0]["vus"]


def test_preclusos_permanecem_como_historia():
    # C1/C2 preclusos (computar=Não) e vigente C3: a historia exibe C0..C3.
    secao = _montar_secao_historico_vu(
        _hist([{"item": "1", "vu_ciclos": {
            "VU_C0": 100.0, "VU_C1": 101.0, "VU_C2": 102.0, "VU_C3": 102.89}}]),
        _ciclos_sec(analisados={3}),
        "C3",
    )
    assert secao["ciclos"] == ["C0", "C1", "C2", "C3"]
    assert secao["itens"][0]["vus"]["C1"] == pytest.approx(101.0)


def test_fallback_sem_ciclo_vigente_usa_ultimo_analisado():
    secao = _montar_secao_historico_vu(
        _hist([{"item": "1", "vu_ciclos": {"VU_C0": 10.0}}]),
        _ciclos_sec(analisados={1, 2}),
        None,
    )
    assert secao["ciclos"] == ["C0", "C1", "C2"]


# ------------------------------------------------- 7-12. renderer do capitulo 5

def _pdf_c3():
    secao = {
        "disponivel": True, "ultimo_ciclo": "C3",
        "ciclos": ["C0", "C1", "C2", "C3"],
        "itens": [
            {"item": 1, "descricao": "Enlace principal",
             "vus": {"C0": 853961.0, "C1": 853961.0, "C2": 853961.0,
                     "C3": 878639.92}},
            {"item": "N001", "descricao": "Item de aditivo (nasce em C2)",
             "vus": {"C0": None, "C1": None, "C2": 100.0, "C3": 102.89}},
            {"item": 2, "descricao": "VU ausente apos nascimento",
             "vus": {"C0": 50.0, "C1": None, "C2": 55.0, "C3": None}},
        ],
    }
    # dados["itens"] legado presente de proposito: o capitulo 5 NAO pode
    # voltar a consumi-lo (Qtd/VU/Total sumiram do PDF).
    legado = [{"item": 1, "descricao": "Enlace principal", "qtd_c0": 2.0,
               "vu_c0": 853961.0, "total_c0": 1707922.0,
               "qtd_ciclos": {"C1": 2.0}, "vu_ciclos": {"C1": 853961.0},
               "total_ciclos": {"C1": 1707922.0}}]
    return _texto_pdf(gerar_sumario_executivo_pdf(_dados_pdf(secao, legado)))


def test_capitulo5_sem_quantidade_e_sem_total():
    texto = _pdf_c3()
    assert "5. Itens e valores atualizados" in texto
    for proibido in ("Qtd C0", "Qtd C1", "Qtd C2", "Qtd C3",
                     "Total C0", "Total C1", "Total C2", "Total C3",
                     "VU C0", "VU C1"):
        assert proibido not in texto, proibido


def test_capitulo5_vu_exato_e_item_original():
    texto = _pdf_c3()
    assert "878.639,92" in texto      # VU registrado em historico_VU (C3)
    assert "853.961,00" in texto      # cadeia completa do item original
    assert "1.707.922,00" not in texto  # total do contrato legado NAO aparece


def test_capitulo5_item_de_aditivo_pre_nascimento():
    texto = _pdf_c3()
    assert "N001" in texto
    assert "—" in texto               # ciclos anteriores ao nascimento
    assert "R$ 100,00" in texto
    assert "R$ 102,89" in texto
    assert "R$ 0,00" not in texto     # ausencia jamais vira zero


def test_capitulo5_ausencia_pos_nascimento_usa_convencao():
    texto = _pdf_c3()
    assert NAO_INFORMADO in texto     # item 2: VU ausente em C1/C3


def test_capitulo5_multiitem_paginacao_e_continuacao():
    fitz = pytest.importorskip("fitz")
    secao = {
        "disponivel": True, "ultimo_ciclo": "C4",
        "ciclos": ["C0", "C1", "C2", "C3", "C4"],
        "itens": [
            {"item": n, "descricao": f"Item de teste {n}",
             "vus": {"C0": 10.0 + n, "C1": 11.0 + n, "C2": 12.0 + n,
                     "C3": 13.0 + n, "C4": 14.0 + n}}
            for n in range(1, 36)
        ],
    }
    pdf = gerar_sumario_executivo_pdf(_dados_pdf(secao))
    assert pdf.startswith(b"%PDF-")
    texto = _texto_pdf(pdf)
    assert "5. Itens e valores atualizados (continuação)" in texto
    doc = fitz.open(stream=pdf, filetype="pdf")
    assert doc.page_count >= 2
    assert "R$ 45,00" in texto        # item 35 em C0 renderizado


def test_capitulo5_sem_historico_vu_informa_ausencia():
    texto = _texto_pdf(gerar_sumario_executivo_pdf(
        _dados_pdf({"disponivel": False, "ciclos": [], "itens": []})))
    assert "5. Itens e valores atualizados" in texto
    assert "historico_VU" in texto


# ------------------------------------------------- 17. demais capitulos intactos

def test_demais_capitulos_preservados():
    texto = _texto_pdf(gerar_sumario_executivo(leitura_simples_financeiro()))
    for secao in ("1. Identificação", "2. Síntese da apuração",
                  "3. Ciclos e efeitos financeiros", "4. Valores financeiros",
                  "5. Itens e valores atualizados", "6. Memória de cálculo",
                  "7. Alterações contratuais consideradas"):
        assert secao in texto, secao
