from __future__ import annotations

import io
import re

import pytest
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

import _ui_utils
from _coleta_oficial import (
    TEMPLATE_COLETA_OFICIAL,
    gerar_coleta_oficial_preenchida,
    obter_coleta_oficial_bytes,
    registrar_versoes_coleta,
)
from _leitor_masterfile_v10 import ler_masterfile_v10
from _versao import (
    CL8US_VERSION,
    COLETA_COMPATIBILIDADE_ANTERIORES,
    COLETA_VERSION,
    COLETA_VERSOES_ACEITAS,
)


def _dados_calculadora() -> dict:
    return {
        "origem": "Reajuste Simples",
        "indice": "IST",
        "data_base_original": "01/01/2023",
        "ciclos": [{
            "ciclo": "C1",
            "data_base": "01/01/2023",
            "data_pedido": "01/01/2024",
            "situacao": "TEMPESTIVO",
            "percentual_aplicado": 0.10,
            "financeiro_inicio": "01/01/2024",
        }],
    }


def _formulas(wb) -> dict[tuple[str, str], str]:
    return {
        (ws.title, cell.coordinate): cell.value
        for ws in wb.worksheets
        for row in ws.iter_rows()
        for cell in row
        if isinstance(cell.value, str) and cell.value.startswith("=")
    }


def _conteudo_exceto_marcadores(wb) -> dict[tuple[str, str], tuple]:
    ignorar = {("CONTROLE", ref) for ref in ("A24", "B24", "A25", "B25")}
    return {
        (ws.title, cell.coordinate): (
            cell.value,
            cell.data_type,
            cell.style_id,
            cell.number_format,
            cell.protection.locked,
        )
        for ws in wb.worksheets
        for row in ws.iter_rows()
        for cell in row
        if (ws.title, cell.coordinate) not in ignorar
    }


def test_versoes_publicas_e_politica_preparada() -> None:
    # Quadro "Execucao sem efeito financeiro" na RESULTADOS: mudanca ESTRUTURAL do
    # XLS (Cl8us 11.4, Modelo de Coleta 11.1); as duas versoes sao independentes e
    # a Coleta 11.0 continua aceita (ver COLETA_VERSOES_ACEITAS).
    # RESULTADOS executiva + RESULTADOS_DETALHE: mudanca ESTRUTURAL do XLS
    # (Cl8us 11.5, Modelo de Coleta 11.2); 11.0 e 11.1 seguem aceitas.
    # PR #174 = Cl8us 11.6 (so app); PR #175 = Cl8us 11.7 / Coleta 11.3 (UX
    # da Coleta). Bumps aplicados pelo hotfix de versionamento obrigatorio.
    # Cl8us 11.8 / Coleta 11.4: fator dos aditivos, ciclo em execucao pela
    # data de corte e "Acrescimo - novo item" (formulas do XLS).
    assert CL8US_VERSION == "11.8"
    assert COLETA_VERSION == "11.4"
    assert COLETA_VERSOES_ACEITAS == ("11.0", "11.1", "11.2", "11.3", "11.4")
    assert CL8US_VERSION != COLETA_VERSION
    assert COLETA_COMPATIBILIDADE_ANTERIORES == 2
    assert re.fullmatch(r"\d{2}\.\d", CL8US_VERSION)
    assert re.fullmatch(r"\d{2}\.\d", COLETA_VERSION)


def test_sidebar_exibe_as_duas_versoes_da_fonte_unica(monkeypatch) -> None:
    legendas: list[str] = []
    monkeypatch.setattr(_ui_utils.st, "markdown", lambda *_a, **_k: None)
    monkeypatch.setattr(_ui_utils.st, "caption", legendas.append)
    monkeypatch.setattr(_ui_utils, "atualizado_em", lambda: "30/09/2026 12:00")

    _ui_utils.render_versao_sidebar()

    assert legendas == [
        "Última atualização publicada em 30/09/2026 12:00",
        "Cl8us 11.8",
        "Modelo de Coleta 11.4",
    ]


def test_area_de_versao_estava_livre_no_template_oficial() -> None:
    wb = load_workbook(TEMPLATE_COLETA_OFICIAL, data_only=False)
    ws = wb["CONTROLE"]
    assert all(ws[coordenada].value is None for coordenada in ("A24", "B24", "A25", "B25"))
    assert not any(
        faixa.min_row <= linha <= faixa.max_row
        and faixa.min_col <= coluna <= faixa.max_col
        for faixa in ws.merged_cells.ranges
        for linha in (24, 25)
        for coluna in (1, 2)
    )
    assert not ws.tables
    assert not ws._images
    assert not ws._charts
    for nome in wb.defined_names.values():
        try:
            destinos = list(nome.destinations)
        except Exception:
            continue
        for aba, referencia in destinos:
            if aba != "CONTROLE":
                continue
            min_col, min_row, max_col, max_row = range_boundaries(
                referencia.replace("$", "")
            )
            assert max_row < 24 or min_row > 25 or max_col < 1 or min_col > 2
    assert not any(
        faixa.min_row <= linha <= faixa.max_row
        and faixa.min_col <= coluna <= faixa.max_col
        for validacao in ws.data_validations.dataValidation
        for faixa in validacao.ranges.ranges
        for linha in (24, 25)
        for coluna in (1, 2)
    )


def test_registro_de_versao_nao_altera_formulas() -> None:
    wb = load_workbook(TEMPLATE_COLETA_OFICIAL, data_only=False)
    formulas_antes = _formulas(wb)
    conteudo_antes = _conteudo_exceto_marcadores(wb)
    nomes_antes = {
        nome: definido.attr_text for nome, definido in wb.defined_names.items()
    }

    registrar_versoes_coleta(wb)

    assert _formulas(wb) == formulas_antes
    assert _conteudo_exceto_marcadores(wb) == conteudo_antes
    assert {
        nome: definido.attr_text for nome, definido in wb.defined_names.items()
    } == nomes_antes
    ws = wb["CONTROLE"]
    assert ws["A24"].value == "Modelo de Coleta"
    assert ws["B24"].value == COLETA_VERSION
    assert ws["A25"].value == "Gerado pelo Cl8us"
    assert ws["B25"].value == CL8US_VERSION


def test_registro_de_versao_nunca_sobrescreve_conteudo() -> None:
    wb = load_workbook(TEMPLATE_COLETA_OFICIAL, data_only=False)
    wb["CONTROLE"]["A24"] = "Conteudo preexistente"

    with pytest.raises(ValueError, match="A24"):
        registrar_versoes_coleta(wb)

    assert wb["CONTROLE"]["A24"].value == "Conteudo preexistente"


def test_coleta_nova_em_branco_e_preenchida_recebem_os_marcadores() -> None:
    for conteudo in (
        obter_coleta_oficial_bytes(),
        gerar_coleta_oficial_preenchida(_dados_calculadora()),
    ):
        ws = load_workbook(io.BytesIO(conteudo), data_only=False)["CONTROLE"]
        assert ws["A24"].value == "Modelo de Coleta"
        assert ws["B24"].value == COLETA_VERSION
        assert ws["A25"].value == "Gerado pelo Cl8us"
        assert ws["B25"].value == CL8US_VERSION


def test_leitor_recupera_versao_explicita_da_coleta_nova() -> None:
    leitura = ler_masterfile_v10(obter_coleta_oficial_bytes())
    assert leitura["coleta_version"] == COLETA_VERSION
    assert leitura["versao_detectada"] == "v10-rc"


def test_arquivo_anterior_sem_marcador_mantem_comportamento_atual() -> None:
    leitura = ler_masterfile_v10(TEMPLATE_COLETA_OFICIAL.read_bytes())
    assert leitura["ok"] is True
    assert leitura["erro"] == ""
    assert leitura["versao_detectada"] == "v10-rc"
    assert "coleta_version" not in leitura


@pytest.mark.parametrize("marcador", ["11.2", "11.3"])
def test_coleta_11_2_e_11_3_com_detalhe_sao_aceitas(marcador) -> None:
    """11.3 (PR #175) so muda apresentacao: arquivo 11.2 segue aceito e ambos
    exigem RESULTADOS_DETALHE (fail-closed desde a 11.2)."""
    from _compatibilidade_coleta import (
        LINHAGEM_COLETA_11,
        detectar_linhagem_coleta,
        exige_resultados_detalhe,
    )

    wb = load_workbook(io.BytesIO(obter_coleta_oficial_bytes()), data_only=False)
    wb["CONTROLE"]["B24"] = marcador
    deteccao = detectar_linhagem_coleta(wb)
    assert deteccao["codigo"] == LINHAGEM_COLETA_11
    assert deteccao["marcador_publico"] == marcador
    assert exige_resultados_detalhe(wb)
