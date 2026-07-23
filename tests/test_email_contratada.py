"""Comunicacao a Contratada — etapa PRE-apuracao financeira (§10.1).

Cobre: 1 ciclo; varios ciclos; precluso sem efeito; tempestivo com data;
data dd/mm/aaaa; indice amigavel; ausencia de valores financeiros/retroativo/
VTA; ausencia de "planilha anexa" e de "valores validados".
"""
from pathlib import Path

import re

from _email_contratada import (
    ASSUNTO_EMAIL_CONTRATADA,
    gerar_rascunho_email_contratada,
)
from _sanitizacao_documental import contem_emoji

ROOT = Path(__file__).resolve().parents[1]


def _sem_valores_financeiros(corpo: str) -> None:
    low = corpo.lower()
    assert "R$" not in corpo
    assert "vta" not in low
    assert "planilha" not in low
    assert "valores validados" not in low
    assert "valor retroativo" not in low
    assert "remanescente" not in low
    assert "unitári" not in low
    # Nenhum valor monetario com centavos
    assert not re.search(r"R?\$?\s*\d{1,3}(\.\d{3})*,\d{2}(?!%)", corpo)


def test_um_ciclo_tempestivo_com_data():
    assunto, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C1", "situacao_aplicada": "Tempestivo",
          "variacao_formatada": "3,27%", "financeiro_inicio": "13/02/2026"}],
        numero_contrato="CT-99/2026",
        indice="ICTI (Ipeadata)",
    )
    assert assunto == ASSUNTO_EMAIL_CONTRATADA
    assert "Contrato CT-99/2026" in corpo
    assert corpo.count("• Ciclo") == 1
    assert "• Ciclo 1: 3,27% – tempestivo, com efeitos financeiros a partir de 13/02/2026." in corpo
    assert re.search(r"\b13/02/2026\b", corpo)  # dd/mm/aaaa
    assert "ICTI (Ipeadata)" in corpo
    _sem_valores_financeiros(corpo)
    assert not contem_emoji(corpo)


def test_varios_ciclos_precluso_e_tempestivo():
    _, corpo = gerar_rascunho_email_contratada(
        [
            {"ciclo": "C0", "situacao_aplicada": "Base"},
            {"ciclo": "C1", "situacao_aplicada": "Precluso",
             "variacao_formatada": "0,00%", "financeiro_inicio": ""},
            {"ciclo": "C2", "situacao_aplicada": "Precluso",
             "percentual_aplicado": 0.0431,
             "financeiro_inicio": "Sem efeitos financeiros automáticos"},
            {"ciclo": "C3", "situacao_aplicada": "Tempestivo",
             "variacao_formatada": "3,27%", "financeiro_inicio": "13/02/2026"},
        ],
        numero_contrato="CT-10/2026",
        indice="IPCA (433)",
    )
    assert corpo.count("• Ciclo") == 3          # C0 ignorado
    assert "• Ciclo 1: 0,00% – precluso, sem efeitos financeiros;" in corpo
    assert "• Ciclo 2: 4,31% – precluso, sem efeitos financeiros;" in corpo
    assert "• Ciclo 3: 3,27% – tempestivo, com efeitos financeiros a partir de 13/02/2026." in corpo
    assert "IPCA" in corpo
    assert "433" not in corpo
    _sem_valores_financeiros(corpo)


def test_tempestivo_sem_data_segura_fica_pendente():
    _, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C1", "situacao_aplicada": "Tempestivo",
          "variacao_formatada": "3,27%", "financeiro_inicio": ""}],
        numero_contrato="CT-1/2026",
        indice="IST (Série Local)",
    )
    assert "efeitos financeiros pendentes de definição" in corpo
    assert "IST (Série Local)" in corpo
    _sem_valores_financeiros(corpo)


def test_natureza_pre_apuracao_sem_apostila_pronta():
    _, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C1", "situacao_aplicada": "Tempestivo",
          "variacao_formatada": "3,27%", "financeiro_inicio": "13/02/2026"}],
        numero_contrato="CT-1/2026", indice="IGP-M (189)",
    )
    low = corpo.lower()
    assert "manifestação de concordância" in low
    assert "apuração dos valores financeiros correspondentes" in low
    assert "IGP-M" in corpo
    assert "189" not in corpo
    assert "apostilamento" not in low


def test_contrato_e_indice_ausentes_usam_marcadores():
    _, corpo = gerar_rascunho_email_contratada([], numero_contrato=None, indice=None)
    assert "[CONTRATO]" in corpo
    assert "[ÍNDICE]" in corpo


def test_integrado_nas_duas_calculadoras():
    simples = (ROOT / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
    multiplo = (ROOT / "pages" / "02_Calculo_Represados.py").read_text(encoding="utf-8")
    assert "render_email_contratada(" in simples
    assert "render_email_contratada(" in multiplo
    # Nao ha mais versao textual antiga duplicada na pagina 01
    assert "_gerar_email_fornecedor" not in simples
