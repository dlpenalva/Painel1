"""Testes focais da Garantia Contratual (método único).

Cobrem as funções puras do motor ``_garantia_calculo`` e os 8 cenários de
aceitação do enunciado: conversão monetária, cálculo de 5% com ROUND_HALF_UP,
diferença entre garantias consecutivas, normalização do histórico, linha do
tempo, e os quatro textos de comunicação (inicial, endosso, redução, sem
alteração). Também valida que a página não mantém resíduos do modelo antigo
(dois métodos, tipo de evento, garantia constituída).
"""
import unittest
from decimal import Decimal
from pathlib import Path

from _garantia_calculo import (
    formatar_brl,
    parse_moeda_br,
    arredondar_financeiro,
    calcular_garantia,
    moeda,
    extrair_vta,
    normalizar_linhas_historico,
    construir_linha_do_tempo,
    formatar_endosso,
    resumo_resultado,
    gerar_texto_comunicacao,
    montar_txt_bytes,
    gerar_pdf_garantia,
    REPORTLAB_OK,
)

ROOT = Path(__file__).resolve().parents[1]
GARANTIA = (ROOT / "pages" / "05_Garantia.py").read_text(encoding="utf-8")


def _tl(*pares):
    """Atalho: constrói a linha do tempo a partir de (instrumento, valor_str)."""
    itens = [{"instrumento": n, "valor_total": parse_moeda_br(v)} for n, v in pares]
    return construir_linha_do_tempo(itens)


class ConversaoMonetariaTests(unittest.TestCase):
    def test_formatos_aceitos(self):
        self.assertEqual(parse_moeda_br("2968866"), Decimal("2968866"))
        self.assertEqual(parse_moeda_br("2968866,00"), Decimal("2968866.00"))
        self.assertEqual(parse_moeda_br("2.968.866,00"), Decimal("2968866.00"))
        self.assertEqual(parse_moeda_br("R$ 2.968.866,00"), Decimal("2968866.00"))
        self.assertEqual(parse_moeda_br("2.968.866"), Decimal("2968866"))
        self.assertEqual(parse_moeda_br("3013124.45"), Decimal("3013124.45"))
        self.assertEqual(parse_moeda_br(3013124.45), Decimal("3013124.45"))
        self.assertEqual(parse_moeda_br(2968866), Decimal("2968866"))

    def test_valores_nao_interpretaveis_retornam_none(self):
        self.assertIsNone(parse_moeda_br(""))
        self.assertIsNone(parse_moeda_br("   "))
        self.assertIsNone(parse_moeda_br("abc"))
        self.assertIsNone(parse_moeda_br(None))
        self.assertIsNone(parse_moeda_br(True))

    def test_formatacao_saida(self):
        self.assertEqual(moeda("2968866"), "R$ 2.968.866,00")
        self.assertEqual(moeda(Decimal("155862.62")), "R$ 155.862,62")
        self.assertEqual(moeda(Decimal("155862.62"), com_prefixo=False), "155.862,62")


class ArredondamentoTests(unittest.TestCase):
    def test_round_half_up(self):
        # 0,125 arredonda para cima (0,13) com ROUND_HALF_UP (float daria 0,12).
        self.assertEqual(arredondar_financeiro(Decimal("0.125")), Decimal("0.13"))
        self.assertEqual(arredondar_financeiro(Decimal("0.135")), Decimal("0.14"))
        # 3.013.124,45 x 5% = 150.656,2225 -> 150.656,22
        self.assertEqual(calcular_garantia(Decimal("3013124.45")), Decimal("150656.22"))
        # 3.117.252,48 x 5% = 155.862,624 -> 155.862,62
        self.assertEqual(calcular_garantia(Decimal("3117252.48")), Decimal("155862.62"))


class CenarioPrincipalTests(unittest.TestCase):
    """Cenário 1 do enunciado."""

    def test_garantias_e_endossos(self):
        tl = _tl(
            ("Contrato", "2.968.866,00"),
            ("Apostila 1", "3.013.124,45"),
            ("Apostila 2", "3.117.252,48"),
        )
        self.assertEqual(tl[0]["garantia"], Decimal("148443.30"))
        self.assertIsNone(tl[0]["endosso"])
        self.assertEqual(tl[0]["tipo_endosso"], "inicial")

        self.assertEqual(tl[1]["garantia"], Decimal("150656.22"))
        self.assertEqual(tl[1]["endosso"], Decimal("2212.92"))
        self.assertEqual(tl[1]["tipo_endosso"], "aumento")

        self.assertEqual(tl[2]["garantia"], Decimal("155862.62"))
        self.assertEqual(tl[2]["endosso"], Decimal("5206.40"))

        vta, garantia_total, endosso, tipo = resumo_resultado(tl)
        self.assertEqual(vta, Decimal("3117252.48"))
        self.assertEqual(garantia_total, Decimal("155862.62"))
        self.assertEqual(endosso, Decimal("5206.40"))
        self.assertEqual(tipo, "aumento")


class CenarioIndependenteDoNomeTests(unittest.TestCase):
    """Cenário 2/3: o cálculo não depende do nome/tipo e cada endosso é sempre
    relativo à linha imediatamente anterior."""

    def test_aditivo_e_apostila_misturados(self):
        tl = _tl(
            ("Contrato", "2.968.866,00"),
            ("Aditivo 1", "3.013.124,45"),
            ("Apostila 1", "3.117.252,48"),
        )
        # Mesmos valores do cenário 1, apenas nomes diferentes -> mesmos endossos.
        self.assertEqual(tl[1]["endosso"], Decimal("2212.92"))
        self.assertEqual(tl[2]["endosso"], Decimal("5206.40"))

    def test_varios_instrumentos_sempre_relativo_ao_anterior(self):
        tl = _tl(
            ("Contrato", "1.000.000,00"),
            ("Apostila 1", "1.100.000,00"),
            ("Apostila 2", "1.150.000,00"),
            ("Aditivo 1", "1.400.000,00"),
            ("Aditivo 2", "1.450.000,00"),
            ("Apostila 3", "1.500.000,00"),
        )
        garantias = [l["garantia"] for l in tl]
        for i in range(1, len(tl)):
            self.assertEqual(tl[i]["endosso"], arredondar_financeiro(garantias[i] - garantias[i - 1]))


class CenarioSemAlteracaoTests(unittest.TestCase):
    """Cenário 4: valor anterior igual ao VTA atual -> endosso R$ 0,00."""

    def test_endosso_zero(self):
        tl = _tl(
            ("Contrato", "2.968.866,00"),
            ("Apostila 1", "2.968.866,00"),
        )
        self.assertEqual(tl[1]["endosso"], Decimal("0.00"))
        self.assertEqual(tl[1]["tipo_endosso"], "sem_alteracao")
        self.assertEqual(formatar_endosso(tl[1]["endosso"], tl[1]["tipo_endosso"]), "R$ 0,00")


class CenarioReducaoTests(unittest.TestCase):
    """Cenário 5: VTA atual inferior ao valor total anterior -> redução."""

    def test_reducao_nao_vira_zero(self):
        tl = _tl(
            ("Contrato", "3.000.000,00"),
            ("Supressão 1", "2.800.000,00"),
        )
        self.assertEqual(tl[1]["endosso"], Decimal("-10000.00"))
        self.assertEqual(tl[1]["tipo_endosso"], "reducao")
        self.assertEqual(formatar_endosso(tl[1]["endosso"], tl[1]["tipo_endosso"]), "Redução de R$ 10.000,00")


class CenarioGarantiaInicialTests(unittest.TestCase):
    """Cenário 6: nenhum histórico anterior -> garantia inicial = 5% do VTA."""

    def test_apenas_analise_atual(self):
        tl = _tl(("Apostila 3", "3.117.252,48"))
        self.assertEqual(len(tl), 1)
        self.assertEqual(tl[0]["garantia"], Decimal("155862.62"))
        self.assertIsNone(tl[0]["endosso"])
        self.assertEqual(tl[0]["tipo_endosso"], "inicial")
        self.assertEqual(formatar_endosso(tl[0]["endosso"], tl[0]["tipo_endosso"]), "—")


class NormalizacaoHistoricoTests(unittest.TestCase):
    def test_linha_vazia_ignorada(self):
        linhas, avisos = normalizar_linhas_historico([{"Instrumento": "", "Valor total do contrato": ""}])
        self.assertEqual(linhas, [])
        self.assertEqual(avisos, [])

    def test_instrumento_sem_valor(self):
        linhas, avisos = normalizar_linhas_historico([{"Instrumento": "Contrato", "Valor total do contrato": ""}])
        self.assertEqual(linhas, [])
        self.assertEqual(len(avisos), 1)

    def test_valor_sem_instrumento(self):
        linhas, avisos = normalizar_linhas_historico([{"Instrumento": "", "Valor total do contrato": "1000"}])
        self.assertEqual(linhas, [])
        self.assertEqual(len(avisos), 1)

    def test_valor_negativo_ou_zero(self):
        linhas, avisos = normalizar_linhas_historico([
            {"Instrumento": "A", "Valor total do contrato": "0"},
            {"Instrumento": "B", "Valor total do contrato": "-5"},
        ])
        self.assertEqual(linhas, [])
        self.assertEqual(len(avisos), 2)

    def test_valor_nao_interpretavel(self):
        linhas, avisos = normalizar_linhas_historico([{"Instrumento": "A", "Valor total do contrato": "abc"}])
        self.assertEqual(linhas, [])
        self.assertEqual(len(avisos), 1)

    def test_linhas_validas_preservam_ordem(self):
        linhas, avisos = normalizar_linhas_historico([
            {"Instrumento": "Contrato", "Valor total do contrato": "2.968.866,00"},
            {"Instrumento": "Apostila 1", "Valor total do contrato": "R$ 3.013.124,45"},
        ])
        self.assertEqual(avisos, [])
        self.assertEqual([l["instrumento"] for l in linhas], ["Contrato", "Apostila 1"])
        self.assertEqual(linhas[1]["valor_total"], Decimal("3013124.45"))


class VtaTests(unittest.TestCase):
    def test_chave_canonica(self):
        self.assertEqual(extrair_vta({"valor_atualizado_contrato": 3117252.48}), Decimal("3117252.48"))

    def test_fallback_valor_global_financeiro(self):
        self.assertEqual(extrair_vta({"valor_global_financeiro": 3117252.48}), Decimal("3117252.48"))

    def test_ausente_ou_invalido_retorna_none(self):
        self.assertIsNone(extrair_vta({}))
        self.assertIsNone(extrair_vta(None))
        self.assertIsNone(extrair_vta({"valor_atualizado_contrato": 0}))
        self.assertIsNone(extrair_vta({"valor_atualizado_contrato": -1}))


class TextoComunicacaoTests(unittest.TestCase):
    def test_texto_endosso_positivo(self):
        texto = gerar_texto_comunicacao("123/2024", Decimal("3117252.48"), Decimal("155862.62"), Decimal("5206.40"), "aumento")
        self.assertIn("Assunto: Garantia contratual", texto)
        self.assertIn("Contrato nº 123/2024", texto)
        self.assertIn("passa a corresponder a R$ 3.117.252,48", texto)
        self.assertIn("totalizando R$ 155.862,62", texto)
        self.assertIn("endosso complementar no valor de R$ 5.206,40", texto)
        self.assertIn("prazo de até 5 dias úteis", texto)
        self.assertIn("TELEBRAS", texto)

    def test_texto_garantia_inicial(self):
        texto = gerar_texto_comunicacao("123/2024", Decimal("3117252.48"), Decimal("155862.62"), None, "inicial")
        self.assertIn("Deverá ser apresentada garantia contratual no valor total de R$ 155.862,62", texto)
        self.assertNotIn("endosso complementar", texto)

    def test_texto_reducao(self):
        texto = gerar_texto_comunicacao("123/2024", Decimal("2800000.00"), Decimal("140000.00"), Decimal("-10000.00"), "reducao")
        self.assertIn("representando redução de R$ 10.000,00", texto)
        self.assertIn("submetida à análise e à aceitação da Telebras", texto)
        self.assertNotIn("prazo de até 5 dias úteis", texto)

    def test_texto_sem_alteracao(self):
        texto = gerar_texto_comunicacao("123/2024", Decimal("2968866.00"), Decimal("148443.30"), Decimal("0.00"), "sem_alteracao")
        self.assertIn("não havendo complementação financeira", texto)
        self.assertNotIn("prazo de até 5 dias úteis", texto)

    def test_contrato_ausente_usa_placeholder(self):
        texto = gerar_texto_comunicacao("", Decimal("1"), Decimal("0.05"), None, "inicial")
        self.assertIn("[número do contrato]", texto)

    def test_txt_utf8_com_bom(self):
        dados = montar_txt_bytes("Teste ção")
        self.assertTrue(dados.startswith(b"\xef\xbb\xbf"))


@unittest.skipUnless(REPORTLAB_OK, "reportlab ausente")
class PdfTests(unittest.TestCase):
    def test_geracao_real_do_pdf(self):
        tl = _tl(("Contrato", "2.968.866,00"), ("Apostila 1", "3.013.124,45"), ("Apostila 2", "3.117.252,48"))
        vta, garantia_total, endosso, tipo = resumo_resultado(tl)
        texto = gerar_texto_comunicacao("123/2024", vta, garantia_total, endosso, tipo)
        pdf = gerar_pdf_garantia({
            "numero_contrato": "123/2024",
            "contratada": "Fornecedora X",
            "vta_atual": vta,
            "garantia_total": garantia_total,
            "endosso": endosso,
            "tipo_endosso": tipo,
            "linha_do_tempo": tl,
            "texto_comunicacao": texto,
        })
        self.assertIsInstance(pdf, (bytes, bytearray))
        self.assertTrue(pdf.startswith(b"%PDF"))
        self.assertGreater(len(pdf), 1000)


class FormatarBrlTests(unittest.TestCase):
    """Função única de formatação monetária brasileira."""

    def test_padrao_brasileiro(self):
        self.assertEqual(formatar_brl(Decimal("3117252.48")), "R$ 3.117.252,48")
        self.assertEqual(formatar_brl(Decimal("2968866")), "R$ 2.968.866,00")
        self.assertEqual(formatar_brl(Decimal("155862.62")), "R$ 155.862,62")
        self.assertEqual(formatar_brl(Decimal("0")), "R$ 0,00")

    def test_nao_produz_formatos_proibidos(self):
        saida = formatar_brl(Decimal("3117252.48"))
        self.assertNotIn("3117252.48", saida)
        self.assertNotIn("3,117,252.48", saida)


class VtaOpcionalTests(unittest.TestCase):
    """Correção: VTA é atalho opcional; ausência não bloqueia a página."""

    def test_valor_manual_funciona_sem_vta(self):
        # Cenário adicional 1: tudo manual, sem VTA — cálculo normal.
        self.assertIsNone(extrair_vta({}))  # sem VTA no Claus
        valor_manual = parse_moeda_br("3.117.252,48")
        tl = construir_linha_do_tempo(
            [{"instrumento": "Contrato", "valor_total": parse_moeda_br("2.968.866,00")},
             {"instrumento": "Apostila 2", "valor_total": valor_manual}]
        )
        _, garantia_total, endosso, tipo = resumo_resultado(tl)
        self.assertEqual(garantia_total, Decimal("155862.62"))
        self.assertEqual(tipo, "aumento")

    def test_pagina_tem_checkbox_e_campo_manual(self):
        self.assertIn("Usar o VTA disponível no Claus", GARANTIA)
        self.assertIn("Valor total atual do contrato", GARANTIA)
        self.assertIn("st.checkbox", GARANTIA)
        # Aviso de substituição explícita do valor manual pelo VTA.
        self.assertIn("substituirá o valor informado manualmente", GARANTIA)

    def test_pagina_nao_bloqueia_por_vta_ausente(self):
        # O antigo bloqueio por VTA inexistente foi removido.
        self.assertNotIn("Nenhum cálculo, PDF ou TXT é gerado com VTA inexistente", GARANTIA)
        self.assertNotIn("VTA atual indisponível", GARANTIA)
        # Checkbox desmarcada por padrão e sem seleção de métodos.
        self.assertIn("value=False", GARANTIA)
        self.assertNotIn("st.radio", GARANTIA)

    def test_pagina_usa_formatacao_unica(self):
        self.assertIn("formatar_brl", GARANTIA)
        self.assertNotIn("moeda(", GARANTIA)


class PaginaSemResiduosDoModeloAntigoTests(unittest.TestCase):
    """A página deve ter um único método e não pode manter resíduos do antigo."""

    def test_metodo_unico_e_integracoes(self):
        # Integrações preservadas.
        self.assertIn('st.session_state["arquivo_garantia_pdf"] = pdf_bytes', GARANTIA)
        self.assertIn('st.session_state["resultado_garantia"]', GARANTIA)
        self.assertIn("← Voltar para Central", GARANTIA)
        self.assertIn('st.switch_page("pages/03_Valor_Global.py")', GARANTIA)
        self.assertIn("table.garantia-tabela th { background: #E6F0F7", GARANTIA)
        # Resíduos do modelo antigo removidos.
        self.assertNotIn("Método 1 — Delta da Garantia", GARANTIA)
        self.assertNotIn("Método 2 — Linha do Tempo Completa", GARANTIA)
        self.assertNotIn("Tipo do evento", GARANTIA)
        self.assertNotIn("garantia_constituida", GARANTIA)
        self.assertNotIn("Garantia/endossos já apresentados", GARANTIA)
        self.assertNotIn("st.radio", GARANTIA)


if __name__ == "__main__":
    unittest.main()
