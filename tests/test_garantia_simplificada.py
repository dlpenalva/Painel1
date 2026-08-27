"""Testes focais da Calculadora de Garantia Contratual (Etapa 49).

Fluxo único e 100% manual: a página compara a situação atual do contrato com a
garantia atualmente constituída e classifica o resultado em exatamente quatro
diagnósticos. Os quinze casos obrigatórios do enunciado estão cobertos aqui:
percentual padrão e manual, suficiência/insuficiência financeira e temporal (e
suas combinações), múltiplas garantias independentes, ausência de dupla contagem
de endossos, os +90 dias, os textos de comunicação, o isolamento absoluto de
qualquer dado externo em session_state e o card da Central como Calculadora.
"""
import unittest
from datetime import date
from decimal import Decimal
from pathlib import Path

from _garantia_calculo import (
    CLAUSULA_GARANTIA,
    DIAGNOSTICO_REGULAR,
    DIAGNOSTICO_VALIDADE,
    DIAGNOSTICO_VALOR,
    DIAGNOSTICO_VALOR_E_VALIDADE,
    DIAS_VALIDADE_MINIMA,
    PERCENTUAL_GARANTIA_PADRAO,
    SEM_NECESSIDADE_DE_ATUALIZACAO,
    analisar_garantia,
    arredondar_financeiro,
    calcular_cobertura_atual,
    calcular_complemento,
    calcular_garantia_necessaria,
    calcular_validade_minima,
    consolidar_garantias,
    formatar_brl,
    formatar_data_br,
    formatar_percentual,
    gerar_texto_comunicacao,
    moeda,
    normalizar_garantias,
    parse_data_br,
    parse_moeda_br,
)
from _garantia_calculo import (
    COLUNA_EVENTO_DATA,
    COLUNA_EVENTO_OBSERVACAO,
    COLUNA_EVENTO_TIPO,
    COLUNA_EVENTO_VALOR,
    COLUNA_EVENTO_VIGENCIA,
    COLUNA_REFERENCIA,
    COLUNA_VALIDADE,
    COLUNA_VALOR,
    FINANCEIRO_COMPLEMENTAR,
    FINANCEIRO_SUFICIENTE,
    FINANCEIRO_SUPERIOR,
    TEMPORAL_INSUFICIENTE,
    TEMPORAL_NAO_INFORMADA,
    TEMPORAL_SUFICIENTE,
    TIPO_ADITIVO,
    TIPO_PRORROGACAO,
    TIPO_REAJUSTE,
    calcular_situacao_atual,
    data_ausente,
    montar_linha_do_tempo,
    normalizar_eventos,
)

ROOT = Path(__file__).resolve().parents[1]
GARANTIA = (ROOT / "pages" / "05_Garantia.py").read_text(encoding="utf-8")
# Código da página sem o docstring do módulo: é ele que descreve, em português,
# justamente as fontes que a página NÃO consulta (VTA, Coleta, RESULTADOS).
CORPO_GARANTIA = GARANTIA.split('"""', 2)[2]
CENTRAL = (ROOT / "pages" / "06_Central_Arquivos.py").read_text(encoding="utf-8")

FIM_VIGENCIA = date(2026, 12, 31)
VALIDADE_MINIMA = date(2027, 3, 31)   # 31/12/2026 + 90 dias corridos


def _garantia(valor, validade, referencia=""):
    """Linha já normalizada, como a devolvida por ``normalizar_garantias``."""
    return {
        "referencia": referencia,
        "valor": arredondar_financeiro(parse_moeda_br(valor)),
        "validade": validade,
    }


def _analise(valor_contrato="1.000.000,00", percentual=PERCENTUAL_GARANTIA_PADRAO,
             fim_vigencia=FIM_VIGENCIA, garantias=()):
    return analisar_garantia(
        valor_total_contrato=valor_contrato,
        percentual=percentual,
        data_fim_vigencia=fim_vigencia,
        garantias=list(garantias),
    )


# ============================================================
# Conversão monetária, percentual e datas
# ============================================================

class ConversaoTests(unittest.TestCase):
    def test_formatos_monetarios_aceitos(self):
        for entrada in ("1000000", "1000000,00", "1.000.000,00", "R$ 1.000.000,00", 1000000):
            self.assertEqual(parse_moeda_br(entrada), Decimal("1000000"), entrada)

    def test_espaco_nao_separavel_e_aceito(self):
        # Valores copiados de Excel ou de páginas web trazem NBSP depois do R$.
        self.assertEqual(parse_moeda_br("R$\u00a01.000,00"), Decimal("1000.00"))
        self.assertEqual(parse_moeda_br("1.000,00\u00a0"), Decimal("1000.00"))

    def test_valores_nao_interpretaveis(self):
        for entrada in ("", "abc", None, "R$"):
            self.assertIsNone(parse_moeda_br(entrada))

    def test_formatacao_brasileira(self):
        self.assertEqual(formatar_brl(Decimal("50000")), "R$ 50.000,00")
        self.assertEqual(moeda(Decimal("1234.5")), "R$ 1.234,50")

    def test_arredondamento_half_up(self):
        self.assertEqual(arredondar_financeiro(Decimal("0.005")), Decimal("0.01"))
        self.assertEqual(arredondar_financeiro(Decimal("2.345")), Decimal("2.35"))

    def test_percentual_formatado_sem_zeros_inuteis(self):
        self.assertEqual(formatar_percentual(Decimal("5.00")), "5")
        self.assertEqual(formatar_percentual(Decimal("4.75")), "4,75")
        self.assertEqual(formatar_percentual(3.5), "3,5")

    def test_datas_aceitas_e_formatadas(self):
        self.assertEqual(parse_data_br("31/12/2026"), FIM_VIGENCIA)
        self.assertEqual(parse_data_br("2026-12-31"), FIM_VIGENCIA)
        self.assertIsNone(parse_data_br(""))
        self.assertEqual(formatar_data_br(FIM_VIGENCIA), "31/12/2026")


# ============================================================
# Casos 1 e 2 — percentual padrão e percentual manual
# ============================================================

class PercentualTests(unittest.TestCase):
    def test_caso1_percentual_padrao_e_cinco_por_cento(self):
        self.assertEqual(PERCENTUAL_GARANTIA_PADRAO, Decimal("5.00"))
        self.assertIn("PERCENTUAL_GARANTIA_PADRAO", GARANTIA)
        self.assertEqual(calcular_garantia_necessaria("1.000.000,00"), Decimal("50000.00"))
        analise = _analise()
        self.assertEqual(analise["percentual"], Decimal("5.00"))
        self.assertEqual(analise["garantia_necessaria"], Decimal("50000.00"))

    def test_caso2_percentual_manual_diferente_de_cinco(self):
        self.assertEqual(
            calcular_garantia_necessaria("1.000.000,00", Decimal("3")), Decimal("30000.00")
        )
        analise = _analise(percentual=Decimal("2.5"))
        self.assertEqual(analise["garantia_necessaria"], Decimal("25000.00"))
        self.assertEqual(formatar_percentual(analise["percentual"]), "2,5")

    def test_percentual_arredonda_meio_centavo_para_cima(self):
        # 1.000.000,10 x 5% = 50.000,005 -> 50.000,01 (ROUND_HALF_UP)
        self.assertEqual(
            calcular_garantia_necessaria("1000000,10", Decimal("5")), Decimal("50000.01")
        )


# ============================================================
# Casos 3, 4 e 13 — suficiência financeira
# ============================================================

class SuficienciaFinanceiraTests(unittest.TestCase):
    def test_caso3_cobertura_insuficiente_gera_complemento(self):
        analise = _analise(garantias=[_garantia("40.000,00", VALIDADE_MINIMA)])
        self.assertEqual(analise["garantia_necessaria"], Decimal("50000.00"))
        self.assertEqual(analise["cobertura_atual"], Decimal("40000.00"))
        self.assertEqual(analise["complemento"], Decimal("10000.00"))
        self.assertFalse(analise["valor_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALOR)

    def test_caso4_cobertura_exatamente_suficiente(self):
        analise = _analise(garantias=[_garantia("50.000,00", VALIDADE_MINIMA)])
        self.assertEqual(analise["complemento"], Decimal("0.00"))
        self.assertTrue(analise["valor_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_cobertura_superior_nao_gera_complemento_negativo(self):
        analise = _analise(garantias=[_garantia("80.000,00", VALIDADE_MINIMA)])
        self.assertEqual(analise["complemento"], Decimal("0.00"))
        self.assertGreaterEqual(analise["complemento"], 0)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_complemento_nunca_negativo(self):
        self.assertEqual(calcular_complemento(Decimal("50000"), Decimal("90000")), Decimal("0.00"))
        self.assertEqual(calcular_complemento(Decimal("50000"), Decimal("50000")), Decimal("0.00"))
        self.assertEqual(calcular_complemento(Decimal("50000"), Decimal("40000")), Decimal("10000.00"))

    def test_caso13_garantia_regular_valor_e_validade_suficientes(self):
        analise = _analise(garantias=[_garantia("50.000,00", date(2027, 6, 30))])
        self.assertTrue(analise["valor_suficiente"])
        self.assertTrue(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_sem_garantia_cadastrada_a_pendencia_e_apenas_de_valor(self):
        analise = _analise(garantias=[])
        self.assertEqual(analise["cobertura_atual"], Decimal("0.00"))
        self.assertEqual(analise["complemento"], Decimal("50000.00"))
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALOR)


# ============================================================
# Casos 5, 6 e 9 — validade mínima e suficiência temporal
# ============================================================

class ValidadeTests(unittest.TestCase):
    def test_caso9_validade_minima_soma_noventa_dias_corridos(self):
        self.assertEqual(DIAS_VALIDADE_MINIMA, 90)
        self.assertEqual(calcular_validade_minima(FIM_VIGENCIA), VALIDADE_MINIMA)
        self.assertEqual((VALIDADE_MINIMA - FIM_VIGENCIA).days, 90)
        # Dias corridos, não "tres meses": 31/12/2027 + 90 dias cai em 30/03/2028.
        self.assertEqual(calcular_validade_minima(date(2027, 12, 31)), date(2028, 3, 30))

    def test_caso5_validade_insuficiente_com_valor_suficiente(self):
        analise = _analise(garantias=[_garantia("50.000,00", date(2027, 1, 31))])
        self.assertTrue(analise["valor_suficiente"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)
        self.assertEqual(analise["complemento"], Decimal("0.00"))

    def test_validade_exatamente_na_data_minima_e_suficiente(self):
        analise = _analise(garantias=[_garantia("50.000,00", VALIDADE_MINIMA)])
        self.assertTrue(analise["validade_suficiente"])

    def test_caso6_valor_e_validade_insuficientes(self):
        analise = _analise(garantias=[_garantia("40.000,00", date(2027, 1, 31))])
        self.assertFalse(analise["valor_suficiente"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALOR_E_VALIDADE)
        self.assertEqual(analise["complemento"], Decimal("10000.00"))

    def test_uma_garantia_curta_entre_varias_torna_a_validade_insuficiente(self):
        analise = _analise(
            garantias=[
                _garantia("30.000,00", date(2027, 6, 30), "Apolice A"),
                _garantia("20.000,00", date(2027, 1, 15), "Apolice B"),
            ]
        )
        self.assertTrue(analise["valor_suficiente"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)

    def test_validade_nao_informada_e_tratada_como_insuficiente(self):
        analise = _analise(garantias=[_garantia("50.000,00", None)])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)

    def test_verificacoes_financeira_e_temporal_sao_independentes(self):
        curta, longa = date(2027, 1, 31), date(2027, 6, 30)
        combinacoes = {
            ("50.000,00", longa): DIAGNOSTICO_REGULAR,
            ("40.000,00", longa): DIAGNOSTICO_VALOR,
            ("50.000,00", curta): DIAGNOSTICO_VALIDADE,
            ("40.000,00", curta): DIAGNOSTICO_VALOR_E_VALIDADE,
        }
        for (valor, validade), esperado in combinacoes.items():
            with self.subTest(valor=valor, validade=validade):
                analise = _analise(garantias=[_garantia(valor, validade)])
                self.assertEqual(analise["diagnostico"], esperado)


# ============================================================
# Casos 7 e 8 — garantias independentes e endossos
# ============================================================

class GarantiasIndependentesTests(unittest.TestCase):
    def test_caso7_multiplas_garantias_independentes_somam(self):
        analise = _analise(
            garantias=[
                _garantia("50.000,00", date(2027, 6, 30), "Apolice A"),
                _garantia("20.000,00", date(2027, 6, 30), "Carta de fianca B"),
            ]
        )
        self.assertEqual(analise["cobertura_atual"], Decimal("70000.00"))
        self.assertEqual(analise["quantidade_garantias"], 2)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_garantias_sem_identificacao_sao_sempre_independentes(self):
        analise = _analise(
            garantias=[
                _garantia("30.000,00", date(2027, 6, 30)),
                _garantia("20.000,00", date(2027, 6, 30)),
            ]
        )
        self.assertEqual(analise["cobertura_atual"], Decimal("50000.00"))
        self.assertEqual(analise["quantidade_garantias"], 2)

    def test_caso8_endossos_da_mesma_garantia_nao_sao_somados(self):
        # Garantia original de 30.000 endossada para 40.000: cobertura = 40.000.
        analise = _analise(
            garantias=[
                _garantia("30.000,00", date(2027, 6, 30), "Apolice 123"),
                _garantia("40.000,00", date(2027, 6, 30), "Apolice 123"),
            ]
        )
        self.assertEqual(analise["cobertura_atual"], Decimal("40000.00"))
        self.assertNotEqual(analise["cobertura_atual"], Decimal("70000.00"))
        self.assertEqual(analise["quantidade_garantias"], 1)
        self.assertEqual(analise["complemento"], Decimal("10000.00"))
        self.assertTrue(analise["avisos_consolidacao"])

    def test_consolidacao_ignora_caixa_e_espacos_na_identificacao(self):
        consolidadas, avisos = consolidar_garantias(
            [
                _garantia("30.000,00", date(2027, 6, 30), " Apolice 123 "),
                _garantia("40.000,00", date(2027, 6, 30), "APOLICE 123"),
            ]
        )
        self.assertEqual(len(consolidadas), 1)
        self.assertEqual(consolidadas[0]["valor"], Decimal("40000.00"))
        self.assertEqual(len(avisos), 1)

    def test_endosso_consolidado_mantem_a_validade_mais_recente(self):
        consolidadas, _ = consolidar_garantias(
            [
                _garantia("30.000,00", date(2027, 6, 30), "Apolice 123"),
                _garantia("40.000,00", date(2027, 1, 31), "Apolice 123"),
            ]
        )
        self.assertEqual(consolidadas[0]["validade"], date(2027, 1, 31))

    def test_cobertura_de_lista_vazia_e_zero(self):
        self.assertEqual(calcular_cobertura_atual([]), Decimal("0.00"))


class NormalizacaoGarantiasTests(unittest.TestCase):
    def test_linha_vazia_ignorada_em_silencio(self):
        linhas, avisos, _pend = normalizar_garantias(
            [{"Apólice / endosso / referência": "", "Valor garantido": "", "Validade": None}]
        )
        self.assertEqual(linhas, [])
        self.assertEqual(avisos, [])

    def test_linha_valida_normalizada(self):
        linhas, avisos, _pend = normalizar_garantias(
            [
                {
                    "Apólice / endosso / referência": " Apolice 123 ",
                    "Valor garantido": "R$ 40.000,00",
                    "Validade": "31/03/2027",
                }
            ]
        )
        self.assertEqual(avisos, [])
        self.assertEqual(linhas[0]["referencia"], "Apolice 123")
        self.assertEqual(linhas[0]["valor"], Decimal("40000.00"))
        self.assertEqual(linhas[0]["validade"], date(2027, 3, 31))

    def test_valor_ausente_ou_invalido_gera_aviso_e_descarta(self):
        linhas, avisos, _pend = normalizar_garantias(
            [
                {"Apólice / endosso / referência": "A", "Valor garantido": "", "Validade": None},
                {"Apólice / endosso / referência": "B", "Valor garantido": "abc", "Validade": None},
                {"Apólice / endosso / referência": "C", "Valor garantido": "-10", "Validade": None},
            ]
        )
        self.assertEqual(linhas, [])
        self.assertEqual(len(avisos), 3)

    def test_validade_ausente_mantem_a_linha_e_avisa(self):
        linhas, avisos, _pend = normalizar_garantias(
            [{"Apólice / endosso / referência": "A", "Valor garantido": "10.000,00", "Validade": None}]
        )
        self.assertEqual(len(linhas), 1)
        self.assertIsNone(linhas[0]["validade"])
        self.assertTrue(any("validade" in aviso.lower() for aviso in avisos))


# ============================================================
# Casos 10, 11 e 12 — textos para comunicação
# ============================================================

class TextoComunicacaoTests(unittest.TestCase):
    def test_caso10_texto_somente_de_valor(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[_garantia("40.000,00", date(2027, 6, 30))]))
        self.assertIn("Prezados,", texto)
        self.assertIn("R$ 1.000.000,00", texto)
        self.assertIn("5%", texto)
        self.assertIn("R$ 50.000,00", texto)
        self.assertIn("A garantia atualmente apresentada é de R$ 40.000,00", texto)
        self.assertIn("complementação no valor de R$ 10.000,00", texto)
        self.assertIn(CLAUSULA_GARANTIA, texto)
        self.assertIn("90 dias após o término da vigência contratual", texto)
        self.assertIn("Gentileza encaminhar o respectivo endosso/comprovante", texto)
        self.assertTrue(texto.rstrip().endswith("Atenciosamente,"))

    def test_caso11_texto_somente_de_validade_nao_menciona_complemento(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[_garantia("50.000,00", date(2027, 1, 31))]))
        self.assertIn("necessidade de atualização da validade", texto)
        self.assertIn("validade mínima até 31/03/2027", texto)
        self.assertIn(CLAUSULA_GARANTIA, texto)
        self.assertNotIn("complementação", texto)
        self.assertNotIn("R$", texto)

    def test_caso12_texto_de_valor_e_validade_menciona_ambos(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[_garantia("40.000,00", date(2027, 1, 31))]))
        self.assertIn("complementação no valor de R$ 10.000,00", texto)
        self.assertIn("atualização da sua validade", texto)
        self.assertIn("31/03/2027", texto)
        self.assertIn(CLAUSULA_GARANTIA, texto)

    def test_texto_regular_nao_pede_nada(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[_garantia("50.000,00", date(2027, 6, 30))]))
        self.assertEqual(texto, SEM_NECESSIDADE_DE_ATUALIZACAO)
        self.assertNotIn("Solicitamos", texto)

    def test_pluralizacao_com_multiplas_garantias(self):
        texto = gerar_texto_comunicacao(
            _analise(
                garantias=[
                    _garantia("20.000,00", date(2027, 6, 30), "A"),
                    _garantia("15.000,00", date(2027, 6, 30), "B"),
                ]
            )
        )
        self.assertIn("As garantias atualmente apresentadas totalizam R$ 35.000,00", texto)
        self.assertNotIn("A garantia atualmente apresentada é", texto)

    def test_texto_sem_garantia_pede_apresentacao(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[]))
        self.assertIn("Não há garantia contratual atualmente apresentada", texto)
        self.assertIn("R$ 50.000,00", texto)

    def test_percentual_manual_aparece_no_texto(self):
        texto = gerar_texto_comunicacao(
            _analise(percentual=Decimal("3"), garantias=[_garantia("10.000,00", date(2027, 6, 30))])
        )
        self.assertIn("correspondente a 3% passa a ser de R$ 30.000,00", texto)


# ============================================================
# Página: fluxo único, manual, sem PDF
# ============================================================

class PaginaFluxoUnicoTests(unittest.TestCase):
    def test_sete_blocos_na_ordem_cronologica(self):
        posicoes = [
            GARANTIA.index('st.subheader("Identificação")'),
            GARANTIA.index('st.subheader("Situação original do contrato")'),
            GARANTIA.index('st.subheader("Alterações posteriores à assinatura")'),
            GARANTIA.index('st.subheader("Situação atual do contrato")'),
            GARANTIA.index('st.subheader("Garantia atualmente apresentada")'),
            GARANTIA.index('st.subheader("Resultado da análise")'),
            GARANTIA.index('st.subheader("Texto para a contratada")'),
        ]
        self.assertEqual(posicoes, sorted(posicoes))

    def test_nao_existe_escolha_de_modo_nem_pagina_com_abas(self):
        for residuo in (
            "st.radio",
            "st.tabs",
            "Valores totais",
            "Acréscimos ou reduções",
            "Memória de cálculo",
            "garantia_historico_valores_totais",
            "garantia_historico_alteracoes",
            "Identificação da garantia",
            "Garantia atualmente constituída",
            "Garantia vigente",
        ):
            self.assertNotIn(residuo, GARANTIA, f"resíduo do modelo antigo: {residuo}")

    def test_alteracoes_do_contrato_nao_sao_garantias_apresentadas(self):
        # Um único editor para os eventos e um único editor para a garantia
        # apresentada: nenhuma fonte visual concorrente para a mesma informação.
        self.assertEqual(GARANTIA.count("st.data_editor("), 2)
        self.assertIn("garantia_eventos_contrato", GARANTIA)
        self.assertIn("garantia_vigente_linhas", GARANTIA)
        # A situação atual é derivada, nunca redigitada.
        self.assertIn("calcular_situacao_atual(", GARANTIA)
        self.assertIn('valor_total_contrato=situacao["valor_atual"]', GARANTIA)
        self.assertIn('data_fim_vigencia=situacao["vigencia_atual"]', GARANTIA)

    def test_pagina_nao_gera_pdf_txt_nem_download(self):
        for residuo in (
            "gerar_pdf_garantia",
            "download_button",
            "arquivo_garantia_pdf",
            "montar_txt_bytes",
            "reportlab",
            "REPORTLAB_OK",
            "application/pdf",
        ):
            self.assertNotIn(residuo, GARANTIA, f"geração de arquivo remanescente: {residuo}")
        self.assertIn("st.text_area(", GARANTIA)

    def test_motor_nao_expoe_mais_pdf_nem_vta(self):
        import _garantia_calculo as motor

        for removido in ("gerar_pdf_garantia", "extrair_vta", "montar_txt_bytes", "REPORTLAB_OK"):
            self.assertFalse(hasattr(motor, removido), f"{removido} deveria ter sido removido do motor")

    def test_campos_manuais_presentes(self):
        from _garantia_calculo import COLUNA_REFERENCIA, COLUNA_VALIDADE, COLUNA_VALOR

        self.assertIn("Valor original do contrato", GARANTIA)
        self.assertIn("Percentual da garantia (%)", GARANTIA)
        self.assertIn("Término da vigência original", GARANTIA)
        self.assertEqual(COLUNA_VALOR, "Valor garantido")
        self.assertEqual(COLUNA_REFERENCIA, "Apólice / endosso / referência")
        self.assertEqual(COLUNA_VALIDADE, "Validade")
        for constante in ("COLUNA_REFERENCIA", "COLUNA_VALOR", "COLUNA_VALIDADE"):
            self.assertIn(constante, CORPO_GARANTIA)
        self.assertIn('num_rows="dynamic"', GARANTIA)

    def test_colunas_do_quadro_de_eventos_presentes(self):
        from _garantia_calculo import (
            COLUNA_EVENTO_DATA,
            COLUNA_EVENTO_OBSERVACAO,
            COLUNA_EVENTO_TIPO,
            COLUNA_EVENTO_VALOR,
            COLUNA_EVENTO_VIGENCIA,
        )

        self.assertEqual(COLUNA_EVENTO_VALOR, "Valor total do contrato após o evento")
        self.assertEqual(COLUNA_EVENTO_VIGENCIA, "Novo término da vigência")
        self.assertEqual(COLUNA_EVENTO_DATA, "Data do instrumento/evento")
        for constante in (
            "COLUNA_EVENTO_TIPO",
            "COLUNA_EVENTO_DATA",
            "COLUNA_EVENTO_VALOR",
            "COLUNA_EVENTO_VIGENCIA",
            "COLUNA_EVENTO_OBSERVACAO",
            "TIPOS_EVENTO",
        ):
            self.assertIn(constante, CORPO_GARANTIA)
        self.assertEqual(COLUNA_EVENTO_TIPO, "Tipo")
        self.assertEqual(COLUNA_EVENTO_OBSERVACAO, "Observação")
        self.assertIn("Evolução do contrato", GARANTIA)

    def test_navegacao_de_retorno_preservada(self):
        self.assertIn("← Voltar para Central", GARANTIA)
        self.assertIn("st.switch_page(_destino_voltar_garantia)", GARANTIA)
        self.assertIn('st.session_state.pop("origem_navegacao_garantia", None)', GARANTIA)

    def test_pagina_publica_apenas_resultado_proprio(self):
        self.assertIn('st.session_state["resultado_garantia"]', GARANTIA)

    def test_texto_da_contratada_nao_congela_na_primeira_apuracao(self):
        # ARMADILHA: st.text_area com key fixa só honra value= no primeiro
        # render; sem a ressincronia por session_state o texto congela na
        # primeira apuração e passa a mentir sobre o resultado exibido acima.
        bloco = GARANTIA[GARANTIA.index("Texto para a contratada") :]
        self.assertIn(
            'if st.session_state.get("garantia_texto_comunicacao") != texto_comunicacao:', bloco
        )
        self.assertIn('st.session_state["garantia_texto_comunicacao"] = texto_comunicacao', bloco)
        inicio = bloco.index("st.text_area(")
        area = bloco[inicio : inicio + 260]
        self.assertNotIn("value=", area, "value= com key fixa congela o texto")


# ============================================================
# Caso 14 — isolamento absoluto de dados externos
# ============================================================

SESSAO_EXTERNA = {
    "resultado_valor_global": {
        "valor_atualizado_contrato": 999_999_999.00,
        "valor_global_financeiro": 999_999_999.00,
        "valor_executado_atualizado": 888_888_888.00,
        "remanescente_reajustado": 777_777_777.00,
    },
    "diagnostico_coleta_v2": {"capacidades": {}, "metadados": {"indice": "IPCA"}},
    "dados_admissibilidade": {"data_base": "01/01/2020"},
    "resultado_adequacao_orcamentaria": {"total": 123_456.78},
    "input_ciclos": [{"ciclo": "C1", "data": "01/08/2024"}],
    "assinatura_processada_upload_docs": "sig-externa",
}


class IsolamentoSessionStateTests(unittest.TestCase):
    def test_pagina_nao_referencia_nenhuma_chave_externa(self):
        for chave in (
            "resultado_valor_global",
            "valor_atualizado_contrato",
            "valor_global_financeiro",
            "extrair_vta",
            "vta_claus",
            "VTA",
            "diagnostico_coleta_v2",
            "input_ciclos",
            "dados_admissibilidade",
            "Coleta",
            "RESULTADOS",
        ):
            self.assertNotIn(chave, CORPO_GARANTIA, f"a Garantia não pode consultar {chave}")

    def test_motor_nao_referencia_nenhuma_chave_externa(self):
        fonte = (ROOT / "_garantia_calculo.py").read_text(encoding="utf-8")
        for chave in ("resultado_valor_global", "valor_atualizado_contrato", "session_state", "streamlit"):
            self.assertNotIn(chave, fonte, f"o motor não pode depender de {chave}")

    def test_sessao_externa_nao_preenche_a_garantia(self):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(str(ROOT / "pages" / "05_Garantia.py"), default_timeout=60)
        for chave, valor in SESSAO_EXTERNA.items():
            at.session_state[chave] = valor
        at.run()
        self.assertFalse(at.exception)

        # Os campos manuais continuam vazios / com o default próprio.
        self.assertEqual(at.text_input(key="garantia_valor_original").value, "")
        self.assertEqual(at.text_input(key="garantia_numero_contrato").value, "")
        self.assertEqual(at.text_input(key="garantia_contratada").value, "")
        self.assertEqual(at.number_input(key="garantia_percentual").value, 5.0)
        self.assertIsNone(at.date_input(key="garantia_fim_vigencia").value)

        # Nenhum valor externo vazou para a tela e a página segue pedindo os
        # dados manualmente (não há resultado calculado a partir da sessão).
        textos = " ".join(
            [elemento.value for elemento in at.markdown]
            + [elemento.value for elemento in at.info]
            + [elemento.value for elemento in at.warning]
        )
        for vazado in ("999.999.999", "888.888.888", "777.777.777", "123.456,78"):
            self.assertNotIn(vazado, textos)
        self.assertIn("valor original do contrato", " ".join(e.value for e in at.info))
        self.assertNotIn("resultado_garantia", at.session_state)

    def test_pagina_calcula_somente_com_entrada_manual(self):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(str(ROOT / "pages" / "05_Garantia.py"), default_timeout=60)
        for chave, valor in SESSAO_EXTERNA.items():
            at.session_state[chave] = valor
        at.run()
        at.text_input(key="garantia_valor_original").set_value("1.000.000,00").run()
        at.date_input(key="garantia_fim_vigencia").set_value(FIM_VIGENCIA).run()
        self.assertFalse(at.exception)

        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["valor_total_contrato"], Decimal("1000000.00"))
        self.assertEqual(resultado["garantia_necessaria"], Decimal("50000.00"))
        self.assertEqual(resultado["validade_minima"], VALIDADE_MINIMA)
        self.assertEqual(resultado["diagnostico"], DIAGNOSTICO_VALOR)


# ============================================================
# Caso 15 — card da Central de Arquivos como Calculadora
# ============================================================

class CardCentralTests(unittest.TestCase):
    def test_card_permanece_na_mesma_posicao_do_catalogo(self):
        import re

        catalogo = CENTRAL[: CENTRAL.index("def aplicar_css_central")]
        nomes = re.findall(r'"nome":\s*"([^"]+)"', catalogo)
        self.assertEqual(
            nomes,
            [
                "Sumário Executivo",
                "Adequação Orçamentária",
                "Despacho Saneador",
                "Termo de Apostila",
                "Garantia Contratual",
                "DOU",
            ],
        )

    def test_card_da_garantia_e_uma_calculadora(self):
        catalogo = CENTRAL[: CENTRAL.index("def aplicar_css_central")]
        inicio = catalogo.index('"nome": "Garantia Contratual"')
        bloco = catalogo[inicio : catalogo.index('"nome": "DOU"')]
        self.assertIn('"formato": "Calculadora"', bloco)
        self.assertIn("Calcule e confira o valor e a validade da garantia contratual.", bloco)
        self.assertIn('"pagina": "pages/05_Garantia.py"', bloco)
        self.assertIn('"ferramenta": True', bloco)
        self.assertIn('"sempre_acessivel": True', bloco)
        self.assertNotIn("arquivo_garantia_pdf", bloco)
        self.assertNotIn("application/pdf", bloco)
        self.assertIn("Abrir calculadora", CENTRAL)

    def test_demais_cards_continuam_entregando_arquivo(self):
        for chave in (
            "arquivo_sumario_executivo_pdf",
            "arquivo_adequacao_orcamentaria_xlsx",
            "arquivo_despacho_saneador_docx",
            "arquivo_termo_apostila_docx",
            "arquivo_dou_docx",
        ):
            self.assertIn(chave, CENTRAL)
        self.assertEqual(CENTRAL.count('"ferramenta": True'), 1)

    def test_card_de_ferramenta_nunca_consulta_a_sessao(self):
        # Um PDF antigo remanescente em session_state não pode transformar o
        # card em download: a flag "ferramenta" zera a consulta à sessão antes
        # de qualquer decisão de renderização.
        render = CENTRAL[CENTRAL.index("def render_documento") : CENTRAL.index("render_marca_topo()")]
        self.assertIn('ferramenta = bool(documento.get("ferramenta"))', render)
        self.assertIn(
            'arquivo = None if ferramenta else st.session_state.get(documento["session_key"])', render
        )
        self.assertIn('label="Abrir calculadora" if ferramenta else "Gerar e baixar"', render)


# ============================================================
# GAR-UX1 — casos A a K: linha do tempo contratual e ausência de data
# ============================================================

VIGENCIA_PRORROGADA = date(2027, 12, 31)
VALIDADE_MINIMA_PRORROGADA = date(2028, 3, 30)   # 31/12/2027 + 90 dias corridos


def _evento(tipo, valor=None, vigencia=None, data=None, observacao=""):
    """Linha crua do quadro de alterações, como o ``st.data_editor`` a devolve."""
    return {
        COLUNA_EVENTO_TIPO: tipo,
        COLUNA_EVENTO_DATA: data,
        COLUNA_EVENTO_VALOR: valor,
        COLUNA_EVENTO_VIGENCIA: vigencia,
        COLUNA_EVENTO_OBSERVACAO: observacao,
    }


def _situacao(valor_original="1.000.000,00", percentual=PERCENTUAL_GARANTIA_PADRAO,
              vigencia=FIM_VIGENCIA, registros=()):
    eventos, avisos, _pend = normalizar_eventos(list(registros))
    situacao = calcular_situacao_atual(
        valor_original=valor_original,
        percentual=percentual,
        fim_vigencia_original=vigencia,
        eventos=eventos,
    )
    return situacao, avisos


def _analise_da_situacao(situacao, garantias=()):
    return analisar_garantia(
        valor_total_contrato=situacao["valor_atual"],
        percentual=situacao["percentual"],
        data_fim_vigencia=situacao["vigencia_atual"],
        garantias=list(garantias),
    )


class CasoASomenteSituacaoOriginalTests(unittest.TestCase):
    def test_sem_eventos_a_situacao_atual_e_a_original(self):
        situacao, avisos = _situacao()
        self.assertEqual(avisos, [])
        self.assertEqual(situacao["quantidade_eventos"], 0)
        self.assertEqual(situacao["valor_original"], Decimal("1000000.00"))
        self.assertEqual(situacao["valor_atual"], Decimal("1000000.00"))
        self.assertEqual(situacao["variacao_acumulada"], Decimal("0.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("50000.00"))
        self.assertEqual(situacao["vigencia_atual"], FIM_VIGENCIA)
        self.assertEqual(situacao["validade_minima"], VALIDADE_MINIMA)

    def test_linhas_vazias_do_editor_sao_ignoradas_em_silencio(self):
        situacao, avisos = _situacao(
            registros=[_evento(None), _evento(""), _evento(None, valor="")]
        )
        self.assertEqual(avisos, [])
        self.assertEqual(situacao["quantidade_eventos"], 0)


class CasoBReajusteTests(unittest.TestCase):
    def test_reajuste_move_valor_variacao_e_garantia(self):
        situacao, avisos = _situacao(
            registros=[_evento(TIPO_REAJUSTE, valor="1.100.000,00")]
        )
        self.assertEqual(avisos, [])
        etapa = situacao["linha_do_tempo"][0]
        self.assertEqual(etapa["numero"], 1)
        self.assertEqual(etapa["valor_anterior"], Decimal("1000000.00"))
        self.assertEqual(etapa["valor"], Decimal("1100000.00"))
        self.assertEqual(etapa["variacao"], Decimal("100000.00"))
        self.assertEqual(etapa["garantia_exigida"], Decimal("55000.00"))
        self.assertEqual(etapa["variacao_garantia"], Decimal("5000.00"))
        self.assertEqual(situacao["valor_atual"], Decimal("1100000.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("55000.00"))
        # O reajuste NÃO é garantia apresentada: a cobertura continua vindo só
        # da grade de garantias.
        analise = _analise_da_situacao(situacao, [_garantia("50.000,00", VALIDADE_MINIMA)])
        self.assertEqual(analise["cobertura_atual"], Decimal("50000.00"))
        self.assertEqual(analise["complemento"], Decimal("5000.00"))
        self.assertEqual(analise["situacao_financeira"], FINANCEIRO_COMPLEMENTAR)

    def test_reajuste_sem_valor_nao_entra_e_avisa(self):
        situacao, avisos = _situacao(registros=[_evento(TIPO_REAJUSTE)])
        self.assertEqual(situacao["quantidade_eventos"], 0)
        self.assertEqual(len(avisos), 1)
        self.assertIn("valor total do contrato após este evento", avisos[0])


class CasoCProrrogacaoPuraTests(unittest.TestCase):
    def test_prorrogacao_sem_valor_preserva_valor_e_recalcula_validade(self):
        situacao, avisos = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, valor="1.100.000,00"),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA),
            ]
        )
        self.assertEqual(avisos, [])
        prorrogacao = situacao["linha_do_tempo"][1]
        self.assertFalse(prorrogacao["valor_informado"])
        self.assertEqual(prorrogacao["valor"], Decimal("1100000.00"))
        self.assertEqual(prorrogacao["variacao"], Decimal("0.00"))
        self.assertEqual(prorrogacao["garantia_exigida"], Decimal("55000.00"))
        self.assertEqual(situacao["valor_atual"], Decimal("1100000.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("55000.00"))
        self.assertEqual(situacao["vigencia_atual"], VIGENCIA_PRORROGADA)
        self.assertEqual(situacao["validade_minima"], VALIDADE_MINIMA_PRORROGADA)

    def test_cobertura_cheia_com_validade_curta_nao_pede_dinheiro(self):
        situacao, _ = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, valor="1.100.000,00"),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA),
            ]
        )
        analise = _analise_da_situacao(situacao, [_garantia("55.000,00", VALIDADE_MINIMA)])
        self.assertTrue(analise["valor_suficiente"])
        self.assertEqual(analise["complemento"], Decimal("0.00"))
        self.assertEqual(analise["situacao_financeira"], FINANCEIRO_SUFICIENTE)
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_INSUFICIENTE)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)
        texto = gerar_texto_comunicacao(analise)
        self.assertNotIn("complementação", texto)
        self.assertNotIn("R$", texto)

    def test_prorrogacao_pode_alterar_o_valor_quando_informado(self):
        situacao, avisos = _situacao(
            registros=[_evento(TIPO_PRORROGACAO, valor="1.200.000,00", vigencia=VIGENCIA_PRORROGADA)]
        )
        self.assertEqual(avisos, [])
        self.assertEqual(situacao["valor_atual"], Decimal("1200000.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("60000.00"))
        self.assertEqual(situacao["vigencia_atual"], VIGENCIA_PRORROGADA)

    def test_prorrogacao_sem_nova_vigencia_nao_entra_e_avisa(self):
        situacao, avisos = _situacao(registros=[_evento(TIPO_PRORROGACAO)])
        self.assertEqual(situacao["quantidade_eventos"], 0)
        self.assertEqual(len(avisos), 1)
        self.assertIn("novo término da vigência", avisos[0])


class CasoDReducaoDeValorTests(unittest.TestCase):
    def test_reducao_nao_pede_complemento_nem_determina_devolucao(self):
        situacao, _ = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, valor="1.100.000,00"),
                _evento(TIPO_ADITIVO, valor="900.000,00"),
            ]
        )
        reducao = situacao["linha_do_tempo"][1]
        self.assertEqual(reducao["variacao"], Decimal("-200000.00"))
        self.assertEqual(situacao["valor_atual"], Decimal("900000.00"))
        self.assertEqual(situacao["variacao_acumulada"], Decimal("-100000.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("45000.00"))

        analise = _analise_da_situacao(situacao, [_garantia("55.000,00", VALIDADE_MINIMA)])
        self.assertEqual(analise["complemento"], Decimal("0.00"))
        self.assertEqual(analise["situacao_financeira"], FINANCEIRO_SUPERIOR)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)
        texto = gerar_texto_comunicacao(analise)
        self.assertEqual(texto, SEM_NECESSIDADE_DE_ATUALIZACAO)
        for proibido in ("devolução", "devolver", "redução da garantia", "Solicitamos"):
            self.assertNotIn(proibido, texto)


class CasoEMaisDeUmEventoTests(unittest.TestCase):
    def test_sequencia_automatica_na_ordem_de_insercao(self):
        situacao, avisos = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, valor="1.100.000,00", data=date(2025, 3, 1)),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA, data=date(2026, 11, 20)),
                _evento(TIPO_ADITIVO, valor="1.300.000,00", data=date(2027, 2, 10)),
            ]
        )
        self.assertEqual(avisos, [])
        self.assertEqual([e["numero"] for e in situacao["linha_do_tempo"]], [1, 2, 3])
        self.assertEqual(
            [e["tipo"] for e in situacao["linha_do_tempo"]],
            [TIPO_REAJUSTE, TIPO_PRORROGACAO, TIPO_ADITIVO],
        )
        self.assertEqual(situacao["valor_atual"], Decimal("1300000.00"))
        self.assertEqual(situacao["vigencia_atual"], VIGENCIA_PRORROGADA)
        self.assertEqual(situacao["garantia_exigida"], Decimal("65000.00"))
        self.assertEqual(situacao["validade_minima"], VALIDADE_MINIMA_PRORROGADA)

    def test_ordem_de_insercao_nunca_e_reordenada_pela_data(self):
        # Datas fora de ordem não reordenam nada: manda a ordem das linhas.
        situacao, _ = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, valor="1.100.000,00", data=date(2027, 5, 1)),
                _evento(TIPO_ADITIVO, valor="1.300.000,00", data=date(2025, 1, 1)),
            ]
        )
        self.assertEqual([e["numero"] for e in situacao["linha_do_tempo"]], [1, 2])
        self.assertEqual(situacao["linha_do_tempo"][0]["tipo"], TIPO_REAJUSTE)
        self.assertEqual(situacao["valor_atual"], Decimal("1300000.00"))

    def test_excluir_evento_anterior_recalcula_toda_a_cadeia(self):
        completo = [
            _evento(TIPO_REAJUSTE, valor="1.100.000,00"),
            _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA),
            _evento(TIPO_ADITIVO, valor="1.300.000,00"),
        ]
        sem_prorrogacao = [completo[0], completo[2]]
        situacao, _ = _situacao(registros=sem_prorrogacao)
        self.assertEqual([e["numero"] for e in situacao["linha_do_tempo"]], [1, 2])
        self.assertEqual(situacao["valor_atual"], Decimal("1300000.00"))
        self.assertEqual(situacao["vigencia_atual"], FIM_VIGENCIA)   # volta à vigência original
        self.assertEqual(situacao["validade_minima"], VALIDADE_MINIMA)

    def test_alterar_evento_anterior_reflete_nos_posteriores(self):
        base = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, valor="1.100.000,00"),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA),
            ]
        )[0]
        editado = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, valor="1.500.000,00"),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA),
            ]
        )[0]
        self.assertEqual(base["valor_atual"], Decimal("1100000.00"))
        self.assertEqual(editado["valor_atual"], Decimal("1500000.00"))
        self.assertEqual(editado["linha_do_tempo"][1]["garantia_exigida"], Decimal("75000.00"))


# ------------------------------------------------------------
# Casos F, G, H, I e J — ausência de data e NaT
# ------------------------------------------------------------

class CasoFGHDataAusenteTests(unittest.TestCase):
    def test_caso_f_nat_do_editor_nao_derruba_a_pagina(self):
        import pandas as pd

        # Exatamente o que o st.data_editor devolve numa célula de data vazia.
        grade = pd.DataFrame(
            {
                COLUNA_REFERENCIA: pd.Series(["Apolice A"], dtype="object"),
                COLUNA_VALOR: pd.Series(["50.000,00"], dtype="object"),
                COLUNA_VALIDADE: pd.Series([pd.NaT], dtype="datetime64[ns]"),
            }
        )
        registros = grade.to_dict("records")
        self.assertTrue(data_ausente(registros[0][COLUNA_VALIDADE]))

        linhas, avisos, _pend = normalizar_garantias(registros)
        self.assertEqual(len(linhas), 1)
        self.assertIsNone(linhas[0]["validade"])
        self.assertTrue(any("validade" in aviso.lower() for aviso in avisos))

        analise = _analise(garantias=linhas)   # não pode levantar TypeError
        self.assertEqual(analise["cobertura_atual"], Decimal("50000.00"))
        self.assertTrue(analise["valor_suficiente"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_NAO_INFORMADA)

    def test_caso_f_nat_cru_chegando_ao_motor_nao_explode(self):
        import pandas as pd

        # Blindagem da causa raiz: NaT é instância de datetime e sobrevive a um
        # "is not None"; o motor precisa reconhecê-lo como ausência.
        self.assertTrue(data_ausente(pd.NaT))
        self.assertIsNone(parse_data_br(pd.NaT))
        analise = _analise(garantias=[_garantia("50.000,00", pd.NaT)])
        self.assertIsNone(analise["garantias"][0]["validade"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)

    def test_caso_f_nat_tambem_e_seguro_no_quadro_de_eventos(self):
        import pandas as pd

        situacao, avisos = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, valor="1.100.000,00", data=pd.NaT, vigencia=pd.NaT),
                _evento(TIPO_PRORROGACAO, data=pd.NaT, vigencia=pd.Timestamp("2027-12-31")),
            ]
        )
        self.assertEqual(avisos, [])
        self.assertIsNone(situacao["linha_do_tempo"][0]["data"])
        self.assertEqual(situacao["linha_do_tempo"][0]["vigencia"], FIM_VIGENCIA)
        self.assertEqual(situacao["vigencia_atual"], VIGENCIA_PRORROGADA)

    def test_caso_g_none_tem_o_mesmo_comportamento_seguro(self):
        self.assertTrue(data_ausente(None))
        self.assertIsNone(parse_data_br(None))
        analise = _analise(garantias=[_garantia("50.000,00", None)])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_NAO_INFORMADA)

    def test_caso_g_outras_formas_de_ausencia(self):
        for ausente in (float("nan"), "", "   ", "NaT", "nan", "None", "<NA>"):
            with self.subTest(ausente=ausente):
                self.assertTrue(data_ausente(ausente))
                self.assertIsNone(parse_data_br(ausente))

    def test_caso_h_timestamp_valido_compara_normalmente(self):
        import pandas as pd

        self.assertFalse(data_ausente(pd.Timestamp("2027-06-30")))
        self.assertEqual(parse_data_br(pd.Timestamp("2027-06-30")), date(2027, 6, 30))
        analise = _analise(garantias=[_garantia("50.000,00", pd.Timestamp("2027-06-30"))])
        self.assertTrue(analise["validade_suficiente"])
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_SUFICIENTE)
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_REGULAR)

    def test_datas_validas_nunca_sao_confundidas_com_ausencia(self):
        from datetime import datetime as _datetime

        for presente in (FIM_VIGENCIA, _datetime(2026, 12, 31, 23, 59), "31/12/2026"):
            with self.subTest(presente=presente):
                self.assertFalse(data_ausente(presente))
                self.assertEqual(parse_data_br(presente), FIM_VIGENCIA)


class CasoIValidadeInsuficienteTests(unittest.TestCase):
    def test_validade_curta_classifica_sem_erro(self):
        analise = _analise(garantias=[_garantia("50.000,00", date(2027, 1, 31))])
        self.assertTrue(analise["valor_suficiente"])
        self.assertEqual(analise["situacao_financeira"], FINANCEIRO_SUFICIENTE)
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_INSUFICIENTE)
        self.assertEqual(analise["complemento"], Decimal("0.00"))
        self.assertEqual(analise["diagnostico"], DIAGNOSTICO_VALIDADE)


class CasoJMultiplasGarantiasTests(unittest.TestCase):
    def test_uma_data_preenchida_e_outra_vazia_nao_provoca_erro(self):
        import pandas as pd

        grade = pd.DataFrame(
            {
                COLUNA_REFERENCIA: pd.Series(["Apolice A", "Carta de fianca B"], dtype="object"),
                COLUNA_VALOR: pd.Series(["30.000,00", "20.000,00"], dtype="object"),
                COLUNA_VALIDADE: pd.Series(
                    [pd.Timestamp("2027-06-30"), pd.NaT], dtype="datetime64[ns]"
                ),
            }
        )
        linhas, avisos, _pend = normalizar_garantias(grade.to_dict("records"))
        self.assertEqual(len(linhas), 2)
        self.assertEqual(linhas[0]["validade"], date(2027, 6, 30))
        self.assertIsNone(linhas[1]["validade"])
        self.assertEqual(len(avisos), 1)

        analise = _analise(garantias=linhas)
        # Semântica econômica preservada: garantias independentes somam.
        self.assertEqual(analise["cobertura_atual"], Decimal("50000.00"))
        self.assertEqual(analise["quantidade_garantias"], 2)
        self.assertTrue(analise["valor_suficiente"])
        self.assertFalse(analise["validade_suficiente"])
        self.assertEqual(analise["situacao_temporal"], TEMPORAL_NAO_INFORMADA)

    def test_endosso_continua_sem_duplicar_a_cobertura(self):
        analise = _analise(
            garantias=[
                _garantia("30.000,00", date(2027, 6, 30), "Apolice 123"),
                _garantia("50.000,00", date(2027, 6, 30), "Apolice 123"),
            ]
        )
        self.assertEqual(analise["cobertura_atual"], Decimal("50000.00"))
        self.assertEqual(analise["quantidade_garantias"], 1)


class CasoKPercentualDiferenteTests(unittest.TestCase):
    def test_percentual_de_tres_por_cento_atravessa_toda_a_pagina(self):
        situacao, _ = _situacao(
            percentual=Decimal("3"),
            registros=[_evento(TIPO_REAJUSTE, valor="2.000.000,00")],
        )
        self.assertEqual(situacao["garantia_original"], Decimal("30000.00"))
        self.assertEqual(situacao["linha_do_tempo"][0]["garantia_exigida"], Decimal("60000.00"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("60000.00"))

        analise = _analise_da_situacao(situacao, [_garantia("40.000,00", VALIDADE_MINIMA)])
        self.assertEqual(analise["percentual"], Decimal("3"))
        self.assertEqual(analise["garantia_necessaria"], Decimal("60000.00"))
        self.assertEqual(analise["complemento"], Decimal("20000.00"))

        texto = gerar_texto_comunicacao(analise)
        self.assertIn("correspondente a 3% passa a ser de R$ 60.000,00", texto)
        self.assertNotIn("5%", texto)

    def test_percentual_quebrado_nao_vira_string_fixa(self):
        situacao, _ = _situacao(percentual=Decimal("4.75"))
        self.assertEqual(situacao["garantia_exigida"], Decimal("47500.00"))
        self.assertEqual(formatar_percentual(situacao["percentual"]), "4,75")
        analise = _analise_da_situacao(situacao, [])
        self.assertIn("correspondente a 4,75%", gerar_texto_comunicacao(analise))


# ------------------------------------------------------------
# Texto à contratada: as quatro situações do enunciado
# ------------------------------------------------------------

class TextoQuatroSituacoesTests(unittest.TestCase):
    def test_a_sem_garantia_apresentada_indica_a_garantia_total(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[]))
        self.assertIn("Não há garantia contratual atualmente apresentada", texto)
        self.assertIn("R$ 50.000,00", texto)

    def test_b_insuficiencia_financeira_indica_o_complemento(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[_garantia("40.000,00", date(2027, 6, 30))]))
        self.assertIn("complementação no valor de R$ 10.000,00", texto)

    def test_c_apenas_validade_nao_inventa_complemento(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[_garantia("50.000,00", date(2027, 1, 31))]))
        self.assertNotIn("R$ 0,00", texto)
        self.assertNotIn("complementação", texto)
        self.assertIn("validade mínima até 31/03/2027", texto)

    def test_d_tudo_suficiente_nao_declara_aceitacao_juridica(self):
        texto = gerar_texto_comunicacao(_analise(garantias=[_garantia("50.000,00", date(2027, 6, 30))]))
        self.assertIn("Com os dados apresentados", texto)
        self.assertIn("não foi identificada necessidade de complementação", texto)
        for proibido in ("aceita", "aprovada", "homologada", "Solicitamos"):
            self.assertNotIn(proibido, texto)

    def test_identificacao_entra_na_referencia_do_texto(self):
        analise = _analise(garantias=[_garantia("40.000,00", date(2027, 6, 30))])
        texto = gerar_texto_comunicacao(
            analise, numero_contrato="123/2024", contratada="Empresa Exemplo Ltda."
        )
        self.assertTrue(texto.startswith("Ref.: Contrato nº 123/2024 — Empresa Exemplo Ltda."))
        self.assertIn("Prezados,", texto)

    def test_identificacao_ausente_nao_deixa_referencia_vazia(self):
        analise = _analise(garantias=[_garantia("40.000,00", date(2027, 6, 30))])
        self.assertTrue(gerar_texto_comunicacao(analise).startswith("Prezados,"))
        self.assertTrue(
            gerar_texto_comunicacao(analise, contratada="Empresa Exemplo Ltda.").startswith(
                "Ref.: Empresa Exemplo Ltda."
            )
        )

    def test_texto_menciona_a_vigencia_atual_apurada(self):
        situacao, _ = _situacao(
            registros=[
                _evento(TIPO_REAJUSTE, valor="1.100.000,00"),
                _evento(TIPO_PRORROGACAO, vigencia=VIGENCIA_PRORROGADA),
            ]
        )
        analise = _analise_da_situacao(situacao, [_garantia("40.000,00", VALIDADE_MINIMA_PRORROGADA)])
        texto = gerar_texto_comunicacao(analise)
        self.assertIn("R$ 1.100.000,00", texto)
        self.assertIn("encerrada em 31/12/2027", texto)


# ------------------------------------------------------------
# Página: preenchimento parcial não derruba a aplicação
# ------------------------------------------------------------

class PaginaPreenchimentoParcialTests(unittest.TestCase):
    def _abrir(self):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(str(ROOT / "pages" / "05_Garantia.py"), default_timeout=60)
        at.run()
        return at

    def test_pagina_abre_vazia_sem_excecao(self):
        at = self._abrir()
        self.assertFalse(at.exception)
        self.assertIn("valor original do contrato", " ".join(e.value for e in at.info))

    def test_somente_situacao_original_ja_produz_a_situacao_atual(self):
        at = self._abrir()
        at.text_input(key="garantia_valor_original").set_value("1.000.000,00").run()
        at.date_input(key="garantia_fim_vigencia").set_value(FIM_VIGENCIA).run()
        self.assertFalse(at.exception)

        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["valor_original"], Decimal("1000000.00"))
        self.assertEqual(resultado["valor_total_contrato"], Decimal("1000000.00"))
        self.assertEqual(resultado["garantia_necessaria"], Decimal("50000.00"))
        self.assertEqual(resultado["quantidade_eventos"], 0)
        self.assertEqual(resultado["validade_minima"], VALIDADE_MINIMA)
        self.assertEqual(resultado["situacao_temporal"], TEMPORAL_NAO_INFORMADA)

    def test_percentual_manual_chega_ao_texto_da_pagina(self):
        at = self._abrir()
        at.text_input(key="garantia_valor_original").set_value("1.000.000,00").run()
        at.date_input(key="garantia_fim_vigencia").set_value(FIM_VIGENCIA).run()
        at.number_input(key="garantia_percentual").set_value(3.0).run()
        self.assertFalse(at.exception)

        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["garantia_necessaria"], Decimal("30000.00"))
        self.assertIn("correspondente a 3% passa a ser de R$ 30.000,00", resultado["texto_comunicacao"])
        self.assertNotIn("5%", resultado["texto_comunicacao"])

    def test_identificacao_da_pagina_chega_a_referencia_do_texto(self):
        at = self._abrir()
        at.text_input(key="garantia_numero_contrato").set_value("123/2024").run()
        at.text_input(key="garantia_contratada").set_value("Empresa Exemplo Ltda.").run()
        at.text_input(key="garantia_valor_original").set_value("1.000.000,00").run()
        at.date_input(key="garantia_fim_vigencia").set_value(FIM_VIGENCIA).run()
        self.assertFalse(at.exception)
        self.assertTrue(
            at.session_state["resultado_garantia"]["texto_comunicacao"].startswith(
                "Ref.: Contrato nº 123/2024 — Empresa Exemplo Ltda."
            )
        )


# ------------------------------------------------------------
# GAR-UX1 fail-closed — entrada materialmente incompleta não conclui
# ------------------------------------------------------------

class PendenciasNoMotorTests(unittest.TestCase):
    def test_a_reajuste_sem_valor_vira_pendencia(self):
        eventos, avisos, pendencias = normalizar_eventos([_evento(TIPO_REAJUSTE)])
        self.assertEqual(eventos, [])
        self.assertEqual(len(pendencias), 1)
        self.assertEqual(pendencias, avisos)
        self.assertIn("valor total do contrato após este evento", pendencias[0])

    def test_b_prorrogacao_sem_vigencia_vira_pendencia(self):
        eventos, avisos, pendencias = normalizar_eventos([_evento(TIPO_PRORROGACAO)])
        self.assertEqual(eventos, [])
        self.assertEqual(len(pendencias), 1)
        self.assertIn("novo término da vigência", pendencias[0])

    def test_c_linha_totalmente_vazia_nao_gera_pendencia(self):
        eventos, avisos, pendencias = normalizar_eventos([_evento(None), _evento("")])
        self.assertEqual(eventos, [])
        self.assertEqual(avisos, [])
        self.assertEqual(pendencias, [])

    def test_d_evento_valido_nao_gera_pendencia(self):
        eventos, avisos, pendencias = normalizar_eventos(
            [_evento(TIPO_REAJUSTE, "1.100.000,00")]
        )
        self.assertEqual(len(eventos), 1)
        self.assertEqual(avisos, [])
        self.assertEqual(pendencias, [])

    def test_e_garantia_descartada_vira_pendencia(self):
        linhas, avisos, pendencias = normalizar_garantias(
            [{COLUNA_REFERENCIA: "Apolice 123", COLUNA_VALOR: "", COLUNA_VALIDADE: None}]
        )
        self.assertEqual(linhas, [])
        self.assertEqual(len(pendencias), 1)
        self.assertIn("informe o valor garantido", pendencias[0])

    def test_e_validade_ausente_avisa_mas_nao_e_pendencia(self):
        # Validade em branco é um estado legítimo (VALIDADE NÃO INFORMADA): a
        # linha compõe a cobertura e a análise financeira segue.
        linhas, avisos, pendencias = normalizar_garantias(
            [{COLUNA_REFERENCIA: "Apolice 123", COLUNA_VALOR: "50.000,00", COLUNA_VALIDADE: None}]
        )
        self.assertEqual(len(linhas), 1)
        self.assertEqual(len(avisos), 1)
        self.assertEqual(pendencias, [])

    def test_garantia_totalmente_vazia_nao_gera_pendencia(self):
        linhas, avisos, pendencias = normalizar_garantias(
            [{COLUNA_REFERENCIA: "", COLUNA_VALOR: "", COLUNA_VALIDADE: None}]
        )
        self.assertEqual((linhas, avisos, pendencias), ([], [], []))


class PaginaFailClosedTests(unittest.TestCase):
    """A página não conclui a análise com história contratual incompleta."""

    def _rodar(self, eventos=(), garantias=()):
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(str(ROOT / "pages" / "05_Garantia.py"), default_timeout=90)
        at.run()
        at.text_input(key="garantia_valor_original").set_value("1.000.000,00").run()
        at.date_input(key="garantia_fim_vigencia").set_value(FIM_VIGENCIA).run()
        if eventos:
            at.session_state["garantia_eventos_contrato"] = {
                "edited_rows": {}, "added_rows": list(eventos), "deleted_rows": [0, 1, 2],
            }
        if garantias:
            at.session_state["garantia_vigente_linhas"] = {
                "edited_rows": {}, "added_rows": list(garantias), "deleted_rows": [0, 1],
            }
        at.run()
        self.assertFalse(at.exception)
        return at

    def _linha_editor(self, tipo=None, valor=None, vigencia=None):
        return {
            COLUNA_EVENTO_TIPO: tipo,
            COLUNA_EVENTO_DATA: None,
            COLUNA_EVENTO_VALOR: valor,
            COLUNA_EVENTO_VIGENCIA: None if vigencia is None else vigencia.isoformat(),
            COLUNA_EVENTO_OBSERVACAO: "",
        }

    def _linha_garantia(self, valor=None, validade=None, referencia=""):
        return {
            COLUNA_REFERENCIA: referencia,
            COLUNA_VALOR: valor,
            COLUNA_VALIDADE: None if validade is None else validade.isoformat(),
        }

    def _suspenso(self, at):
        """Conclusão suspensa: nada publicado e o convite a completar a linha."""
        self.assertNotIn("resultado_garantia", at.session_state)
        infos = " ".join(e.value for e in at.info)
        self.assertIn("Complete ou remova a linha para concluir a análise", infos)
        return infos

    def test_a_reajuste_incompleto_nao_publica_conclusao(self):
        at = self._rodar(eventos=[self._linha_editor(TIPO_REAJUSTE)])
        avisos = " ".join(e.value for e in at.warning)
        self.assertIn("informe o valor total do contrato após este evento", avisos)
        infos = self._suspenso(at)
        self.assertIn("Há alteração contratual com dados pendentes", infos)
        # Nem diagnóstico nem texto conclusivo chegam à tela.
        corpo = " ".join([e.value for e in at.markdown] + [e.value for e in at.warning])
        self.assertNotIn(DIAGNOSTICO_REGULAR, corpo)
        self.assertNotIn("Texto para a contratada", corpo)

    def test_b_prorrogacao_incompleta_nao_publica_conclusao(self):
        at = self._rodar(eventos=[self._linha_editor(TIPO_PRORROGACAO)])
        avisos = " ".join(e.value for e in at.warning)
        self.assertIn("novo término da vigência", avisos)
        self._suspenso(at)

    def test_c_linha_vazia_continua_ignorada_e_nao_bloqueia(self):
        at = self._rodar(eventos=[self._linha_editor()])
        validacoes = [e.value for e in at.warning if e.value.startswith("Linha ")]
        self.assertEqual(validacoes, [])
        self.assertEqual([e.value for e in at.info], [])
        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["quantidade_eventos"], 0)
        self.assertEqual(resultado["valor_total_contrato"], Decimal("1000000.00"))

    def test_d_evento_valido_continua_calculando(self):
        at = self._rodar(eventos=[self._linha_editor(TIPO_REAJUSTE, "1.100.000,00")])
        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["quantidade_eventos"], 1)
        self.assertEqual(resultado["valor_total_contrato"], Decimal("1100000.00"))
        self.assertEqual(resultado["garantia_necessaria"], Decimal("55000.00"))
        self.assertTrue(resultado["texto_comunicacao"])

    def test_e_garantia_apresentada_incompleta_nao_publica_conclusao(self):
        at = self._rodar(garantias=[self._linha_garantia(referencia="Apolice 123")])
        avisos = " ".join(e.value for e in at.warning)
        self.assertIn("informe o valor garantido", avisos)
        infos = self._suspenso(at)
        self.assertIn("Há garantia apresentada com dados pendentes", infos)

    def test_e_garantia_sem_validade_conclui_normalmente(self):
        # Validade em branco não é pendência: a conclusão sai, com a situação
        # temporal reportada como não informada.
        at = self._rodar(garantias=[self._linha_garantia("50.000,00", None, "Apolice 123")])
        resultado = at.session_state["resultado_garantia"]
        self.assertEqual(resultado["cobertura_atual"], Decimal("50000.00"))
        self.assertEqual(resultado["situacao_temporal"], TEMPORAL_NAO_INFORMADA)

    def test_pendencia_no_evento_nao_esconde_a_evolucao_ja_valida(self):
        # A evolução válida até ali continua visível e os editores permanecem.
        at = self._rodar(
            eventos=[
                self._linha_editor(TIPO_REAJUSTE, "1.100.000,00"),
                self._linha_editor(TIPO_PRORROGACAO),
            ]
        )
        self._suspenso(at)
        corpo = " ".join(e.value for e in at.markdown)
        self.assertIn("Evolução do contrato", corpo)
        self.assertIn("R$ 1.100.000,00", corpo)

    def test_conclusao_anterior_nao_sobrevive_a_pendencia(self):
        # session_state persiste entre reruns: sem retirar a publicação, o
        # Saneador continuaria lendo um resultado que a tela já não sustenta.
        from streamlit.testing.v1 import AppTest

        at = AppTest.from_file(str(ROOT / "pages" / "05_Garantia.py"), default_timeout=90)
        at.run()
        at.text_input(key="garantia_valor_original").set_value("1.000.000,00").run()
        at.date_input(key="garantia_fim_vigencia").set_value(FIM_VIGENCIA).run()
        self.assertIn("resultado_garantia", at.session_state)   # conclusão publicada

        at.session_state["garantia_eventos_contrato"] = {
            "edited_rows": {}, "added_rows": [self._linha_editor(TIPO_REAJUSTE)],
            "deleted_rows": [0, 1, 2],
        }
        at.run()
        self.assertFalse(at.exception)
        self.assertNotIn("resultado_garantia", at.session_state)

    def test_ordem_da_cadeia_e_a_ordem_das_linhas_e_nao_a_data(self):
        # Verificação exigida no checkpoint: datas fora de ordem não reordenam.
        eventos, _, _ = normalizar_eventos(
            [
                _evento(TIPO_REAJUSTE, "1.100.000,00", data=date(2027, 5, 1)),
                _evento(TIPO_ADITIVO, "1.300.000,00", data=date(2025, 1, 1)),
            ]
        )
        self.assertEqual([e["numero"] for e in eventos], [1, 2])
        self.assertEqual([e["tipo"] for e in eventos], [TIPO_REAJUSTE, TIPO_ADITIVO])
        # E a UI segue orientando o lançamento em ordem cronológica.
        self.assertIn("em ordem cronológica", GARANTIA)
        self.assertIn("eventos mais antigos", GARANTIA)


if __name__ == "__main__":
    unittest.main()
