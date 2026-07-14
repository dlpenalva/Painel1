from pathlib import Path

from _email_contratada import (
    ASSUNTO_EMAIL_CONTRATADA,
    gerar_rascunho_email_contratada,
)


ROOT = Path(__file__).resolve().parents[1]


def test_rascunho_preenche_percentuais_sem_inventar_retroativo():
    assunto, corpo = gerar_rascunho_email_contratada(
        [
            {"ciclo": "C1", "percentual_aplicado": 0.05, "fator_acumulado": 1.05},
            {"ciclo": "C2", "percentual_aplicado": 0.04, "fator_acumulado": 1.092},
        ]
    )

    assert assunto == ASSUNTO_EMAIL_CONTRATADA
    assert "C1\t5,00%\t5,00%" in corpo
    assert "C2\t4,00%\t9,20%" in corpo
    assert "[NÚMERO DO CONTRATO]" in corpo
    assert "C1\tR$ [VALOR VALIDADO]" in corpo
    assert "Total\tR$ [VALOR TOTAL VALIDADO]" in corpo


def test_rascunho_usa_contrato_informado_e_ignora_c0():
    _, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C0"}, {"ciclo": "C1", "variacao": 0.031}],
        "CT-123/2026",
    )

    assert "Contrato nº CT-123/2026" in corpo
    assert "C0\t" not in corpo
    assert "C1\t3,10%" in corpo


def test_opcao_esta_integrada_apos_processamento_nas_duas_calculadoras():
    simples = (ROOT / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
    multiplo = (ROOT / "pages" / "02_Calculo_Represados.py").read_text(encoding="utf-8")

    assert "gerar_rascunho_email_contratada" in simples
    assert "render_email_contratada(" in multiplo
    assert multiplo.index("render_email_contratada(") > multiplo.index("if historico_coleta:")
