"""Comunicacao a Contratada — etapa PRE-apuracao financeira (§10.1).

Cobre: 1 ciclo; varios ciclos; precluso sem efeito; tempestivo com data;
data dd/mm/aaaa; indice amigavel; ausencia de valores financeiros/retroativo/
VTA; ausencia de "planilha anexa" e de "valores validados"; secao MEMORIA DE
CALCULO reapresentando a memoria do indice do payload (mensal e IST/INDICE),
uma secao por ciclo com memoria, nada para ciclo sem memoria.
"""
from pathlib import Path

import re

import pytest

from _email_contratada import (
    ASSUNTO_EMAIL_CONTRATADA,
    _competencia_mm_aaaa,
    _situacao_efeito,
    gerar_rascunho_email_contratada,
    montar_txt_download,
)
from _sanitizacao_documental import contem_emoji

ROOT = Path(__file__).resolve().parents[1]


@pytest.mark.parametrize(
    ("entrada", "esperado"),
    [
        ("2026-02-13", "02/2026"),
        ("2026-02-13T00:00:00", "02/2026"),
        ("13/02/2026", "02/2026"),
        ("02/2026", "02/2026"),
        ("31/02/2026", None),
        ("2026-02-31", None),
        ("99/02/2026", None),
        ("13/99/2026", None),
    ],
)
def test_competencia_valida_calendario_real(entrada, esperado):
    assert _competencia_mm_aaaa(entrada) == esperado


def test_data_invalida_mantem_efeitos_pendentes_de_definicao():
    assert _situacao_efeito({
        "situacao_aplicada": "Tempestivo",
        "financeiro_inicio": "31/02/2026",
    }) == "tempestivo, efeitos financeiros pendentes de definição"


def _sem_valores_financeiros(corpo: str) -> None:
    low = corpo.lower()
    assert "R$" not in corpo
    assert "vta" not in low
    assert "planilha" not in low
    assert "valores validados" not in low
    assert "valor retroativo" not in low
    assert "remanescente" not in low
    assert "unitári" not in low
    # Nenhum valor monetario com centavos. O lookahead exclui percentuais
    # ("4,30%"), fator apurado (6 casas: "1,043031") e numero-indice (4 casas:
    # "104,5600") da memoria de calculo — que nao sao valores financeiros.
    assert not re.search(r"R?\$?\s*\d{1,3}(\.\d{3})*,\d{2}(?![\d%])", corpo)


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
    assert "• Ciclo 1: 3,27% – tempestivo, com efeitos financeiros a partir de 02/2026." in corpo
    assert re.search(r"\b02/2026\b", corpo)
    assert "13/02/2026" not in corpo
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
    assert "• Ciclo 3: 3,27% – tempestivo, com efeitos financeiros a partir de 02/2026." in corpo
    assert "IPCA" in corpo
    assert "433" not in corpo
    _sem_valores_financeiros(corpo)


def test_tempestivo_sem_data_segura_fica_pendente():
    _, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C1", "situacao_aplicada": "Tempestivo",
          "variacao_formatada": "3,27%", "financeiro_inicio": ""}],
        numero_contrato="CT-1/2026",
        indice="IST (Série Local)",  # rotulo legado ainda aceito na entrada
    )
    assert "efeitos financeiros pendentes de definição" in corpo
    # user-facing renomeado: IST (Anatel), preservando reconhecimento do legado
    assert "IST (Anatel)" in corpo
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


MEMORIA_MENSAL_C3 = [
    {"tipo": "MES", "ordem": 1, "competencia": "2025-02-01",
     "valor_indice": 0.0148},
    {"tipo": "MES", "ordem": 2, "competencia": "2025-03-01",
     "valor_indice": 0.0051},
    {"tipo": "MES", "ordem": 3, "competencia": "2026-01-01",
     "valor_indice": 0.0039},
    {"tipo": "RESULTADO", "ordem": 4, "fator_acumulado": 1.043031,
     "variacao_final": 0.0430,
     "metodo_fonte": "Produtorio de taxas mensais (SGS/BCB)"},
]

MEMORIA_IST_C2 = [
    {"tipo": "INDICE", "ordem": 1, "competencia": "2025-02-01",
     "valor_indice": 104.56},
    {"tipo": "INDICE", "ordem": 2, "competencia": "2026-01-01",
     "valor_indice": 108.91},
    {"tipo": "RESULTADO", "ordem": 3, "fator_acumulado": 1.0416,
     "variacao_final": 0.0416,
     "metodo_fonte": "Divisao de Numero-Indice (IST/Anatel)"},
]


def test_memoria_mensal_no_ciclo_unico():
    _, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C3", "situacao_aplicada": "Tempestivo",
          "variacao_formatada": "4,30%", "financeiro_inicio": "01/02/2026",
          "memoria_calculo": MEMORIA_MENSAL_C3}],
        numero_contrato="CT-99/2026", indice="INPC",
    )
    # A secao entra DEPOIS da frase final, com duas quebras de linha.
    assert "informações acima.\n\nMEMÓRIA DE CÁLCULO\n\nCiclo 3\n\n" in corpo
    assert "02/2025: 1,48%" in corpo
    assert "03/2025: 0,51%" in corpo
    assert "01/2026: 0,39%" in corpo
    assert "Fator apurado: 1,043031" in corpo
    assert "Variação apurada: 4,30%" in corpo
    assert "Método/Fonte: Produtorio de taxas mensais (SGS/BCB)" in corpo
    _sem_valores_financeiros(corpo)
    assert not contem_emoji(corpo)


def test_memoria_ist_apresenta_numeros_indice_sem_meses_ficticios():
    _, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C2", "situacao_aplicada": "Tempestivo",
          "variacao_formatada": "4,16%", "financeiro_inicio": "01/03/2026",
          "memoria_calculo": MEMORIA_IST_C2}],
        numero_contrato="CT-99/2026", indice="IST (Anatel)",
    )
    assert "MEMÓRIA DE CÁLCULO" in corpo
    assert "Número-índice inicial (02/2025): 104,5600" in corpo
    assert "Número-índice final (01/2026): 108,9100" in corpo
    assert "Fator apurado: 1,041600" in corpo
    assert "Variação apurada: 4,16%" in corpo
    assert "Método/Fonte: Divisao de Numero-Indice (IST/Anatel)" in corpo
    # IST nao vira memoria mensal ficticia: nenhuma linha "mm/aaaa: x%".
    assert not re.search(r"^\d{2}/\d{4}: ", corpo, flags=re.MULTILINE)
    _sem_valores_financeiros(corpo)


def test_multiciclo_uma_secao_por_ciclo_com_memoria():
    _, corpo = gerar_rascunho_email_contratada(
        [
            {"ciclo": "C1", "situacao_aplicada": "Tempestivo",
             "variacao_formatada": "1,99%", "financeiro_inicio": "01/01/2025",
             "memoria_calculo": [
                 {"tipo": "MES", "ordem": 1, "competencia": "2024-01-01",
                  "valor_indice": 0.0199},
                 {"tipo": "RESULTADO", "ordem": 2, "fator_acumulado": 1.0199,
                  "variacao_final": 0.0199, "metodo_fonte": "SGS/BCB"},
             ]},
            {"ciclo": "C2", "situacao_aplicada": "Precluso",
             "variacao_formatada": "0,00%", "financeiro_inicio": ""},
            {"ciclo": "C3", "situacao_aplicada": "Tempestivo",
             "variacao_formatada": "4,30%", "financeiro_inicio": "01/02/2026",
             "memoria_calculo": MEMORIA_MENSAL_C3},
        ],
        numero_contrato="CT-10/2026", indice="INPC",
    )
    assert corpo.count("MEMÓRIA DE CÁLCULO") == 1
    assert "\nCiclo 1\n" in corpo
    assert "\nCiclo 3\n" in corpo
    # C2 nao tem memoria: nenhum bloco proprio (o marcador "• Ciclo 2" da
    # lista de percentuais permanece, mas sem secao de memoria).
    assert "\nCiclo 2\n" not in corpo
    assert "Fator apurado: 1,019900" in corpo
    assert "Fator apurado: 1,043031" in corpo
    _sem_valores_financeiros(corpo)


def test_ciclo_sem_memoria_nao_gera_secao():
    _, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C1", "situacao_aplicada": "Tempestivo",
          "variacao_formatada": "3,27%", "financeiro_inicio": "13/02/2026"}],
        numero_contrato="CT-1/2026", indice="IPCA",
    )
    assert "MEMÓRIA DE CÁLCULO" not in corpo
    assert corpo.rstrip().endswith("informações acima.")
    _sem_valores_financeiros(corpo)


def test_fail_closed_memoria_inconsistente_gera_erro_controlado():
    """Memoria existe no payload mas nao pode virar secao -> ValueError.

    Nunca mais "memoria sumiu -> TXT silenciosamente incompleto".
    """
    memoria_invalida = [{"tipo": "MES", "ordem": 1}]  # sem competencia/valor
    with pytest.raises(ValueError, match="MEMÓRIA DE CÁLCULO|bloco"):
        gerar_rascunho_email_contratada(
            [{"ciclo": "C1", "situacao_aplicada": "Tempestivo",
              "variacao_formatada": "3,27%", "financeiro_inicio": "13/02/2026",
              "memoria_calculo": memoria_invalida}],
            numero_contrato="CT-1/2026", indice="IPCA",
        )


def test_fail_closed_nao_dispara_com_memoria_valida_nem_sem_memoria():
    # com memoria valida: secao presente, sem erro
    _, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C3", "situacao_aplicada": "Tempestivo",
          "variacao_formatada": "4,30%", "financeiro_inicio": "01/02/2026",
          "memoria_calculo": MEMORIA_MENSAL_C3}],
        numero_contrato="CT-1/2026", indice="INPC",
    )
    assert "MEMÓRIA DE CÁLCULO" in corpo
    # sem memoria: comportamento normal (sem secao, sem erro)
    _, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C1", "situacao_aplicada": "Tempestivo",
          "variacao_formatada": "3,27%", "financeiro_inicio": "13/02/2026"}],
        numero_contrato="CT-1/2026", indice="IPCA",
    )
    assert "MEMÓRIA DE CÁLCULO" not in corpo


def test_montar_txt_download_sao_os_bytes_do_botao():
    """Os bytes finais do download: BOM utf-8-sig + ASSUNTO + corpo."""
    dados = montar_txt_download("Assunto X", "Corpo Y")
    assert dados == "ASSUNTO: Assunto X\n\nCorpo Y".encode("utf-8-sig")
    assert dados.decode("utf-8-sig").startswith("ASSUNTO: Assunto X")


def test_integrado_nas_duas_calculadoras():
    simples = (ROOT / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
    multiplo = (ROOT / "pages" / "02_Calculo_Represados.py").read_text(encoding="utf-8")
    assert "render_email_contratada(" in simples
    assert "render_email_contratada(" in multiplo
    # Nao ha mais versao textual antiga duplicada na pagina 01
    assert "_gerar_email_fornecedor" not in simples
    # As duas paginas passam o acumulado CANONICO explicitamente (a linha
    # "Variação Acumulada Final" nao depende do campo dentro do payload).
    assert 'fator_acumulado=_adm_email.get("fator_acumulado")' in simples
    assert "fator_acumulado=float(fator_acum)" in multiplo


def test_variacao_acumulada_final_fonte_canonica_explicita():
    """O acumulado canonico da analise chega por parametro: a linha aparece
    MESMO quando o payload dos ciclos nao traz fator_acumulado (defeito real
    reproduzido no teste funcional pos-PR #61)."""
    _, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C1", "situacao_aplicada": "Tempestivo",
          "variacao_formatada": "3,27%", "financeiro_inicio": "13/02/2026"}],
        fator_acumulado=1.1474,
    )
    assert "Variação Acumulada Final: 14,74%." in corpo
    # Posicao: depois da lista dos ciclos e antes do paragrafo dos preclusos.
    pos_ciclo = corpo.rindex("• Ciclo")
    pos_acum = corpo.index("Variação Acumulada Final")
    pos_preclusos = corpo.index("Eventuais ciclos preclusos")
    assert pos_ciclo < pos_acum < pos_preclusos
    _sem_valores_financeiros(corpo)


def test_variacao_acumulada_final_fallback_payload():
    """Sem parametro explicito, mantem o fallback: ultimo fator_acumulado
    numerico do payload dos ciclos (nunca soma de percentuais)."""
    _, corpo = gerar_rascunho_email_contratada(
        [
            {"ciclo": "C1", "variacao_formatada": "3,27%",
             "fator_acumulado": 1.0327},
            {"ciclo": "C2", "variacao_formatada": "4,30%",
             "fator_acumulado": 1.0771},
        ],
    )
    assert "Variação Acumulada Final: 7,71%." in corpo


def test_variacao_acumulada_final_omitida_sem_fonte():
    """Sem acumulado canonico e sem fator no payload, nada e inventado."""
    _, corpo = gerar_rascunho_email_contratada(
        [{"ciclo": "C1", "variacao_formatada": "3,27%"}],
    )
    assert "Variação Acumulada Final" not in corpo
