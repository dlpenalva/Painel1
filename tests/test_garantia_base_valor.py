"""A base do cálculo da Garantia: valor atualizado, nunca zero presumido.

`pages/05_Garantia.py` é uma página Streamlit e não pode ser importada sem
subir a tela. As funções sob teste são extraídas do próprio arquivo por AST e
executadas isoladas — o código exercitado é o de produção, linha por linha, sem
cópia paralela que possa divergir dele.
"""

import ast
import unittest
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
GARANTIA = (ROOT / "pages" / "05_Garantia.py").read_text(encoding="utf-8")

FUNCOES_SOB_TESTE = (
    "parse_moeda_br",
    "numero_para_input",
    "limpar_texto",
    "obter_contexto_valor_global",
)


def carregar_funcoes_da_pagina() -> dict:
    """Executa apenas as funções de interesse, sem tocar no Streamlit."""

    arvore = ast.parse(GARANTIA)
    escolhidas = [
        no
        for no in arvore.body
        if isinstance(no, ast.FunctionDef) and no.name in FUNCOES_SOB_TESTE
    ]
    faltando = set(FUNCOES_SOB_TESTE) - {no.name for no in escolhidas}
    if faltando:
        raise AssertionError(f"Funções ausentes em 05_Garantia.py: {sorted(faltando)}")

    espaco = {"pd": pd}
    exec(compile(ast.Module(body=escolhidas, type_ignores=[]), "05_Garantia.py", "exec"), espaco)
    return espaco


ESPACO = carregar_funcoes_da_pagina()
obter_contexto_valor_global = ESPACO["obter_contexto_valor_global"]


def resultado_oficial(**sobrescritas) -> dict:
    """Recorte do contrato de 68 chaves que a Garantia consome."""

    base = {
        "valor_original_contrato": 1_000_000.00,
        "valor_atualizado_contrato": 1_045_000.00,
        "valor_formalizado_anterior": 1_000_000.00,
        "valor_executado_atualizado": 300_000.00,
        "remanescente_reajustado": 745_000.00,
        "variacao_acumulada": 0.045,
        "quantidade_aditivos_total": 0,
        "modo_apuracao": "Financeiro",
        "valor_retroativo_estimado_itens_estoque": 0.0,
        "contexto_contratual_anterior": {},
    }
    base.update(sobrescritas)
    return base


class BaseDoValorDaGarantiaTests(unittest.TestCase):
    def test_base_da_garantia_e_o_valor_atualizado_do_contrato(self):
        contexto = obter_contexto_valor_global(resultado_oficial())
        self.assertEqual(contexto["valor_total_atualizado"], 1_045_000.00)

    def test_valor_atualizado_nao_e_confundido_com_o_valor_original(self):
        contexto = obter_contexto_valor_global(resultado_oficial())
        self.assertNotEqual(contexto["valor_total_atualizado"], contexto["valor_original"])
        self.assertEqual(contexto["valor_original"], 1_000_000.00)

    def test_contrato_oficial_nao_produz_zero_silencioso(self):
        """O defeito temido: base zerada apesar de o contrato trazer o valor."""

        contexto = obter_contexto_valor_global(resultado_oficial())
        self.assertNotEqual(contexto["valor_total_atualizado"], 0.0)
        self.assertGreater(contexto["valor_total_atualizado"], 0.0)

    def test_chaves_inexistentes_no_contrato_nao_sao_consultadas(self):
        """Se a leitura dependesse delas, o valor viria zerado — e não vem."""

        resultado = resultado_oficial()
        resultado.pop("valor_global_contrato", None)
        resultado.pop("valor_global_estoque", None)
        contexto = obter_contexto_valor_global(resultado)
        self.assertEqual(contexto["valor_total_atualizado"], 1_045_000.00)

    def test_valor_obrigatorio_ausente_nao_e_mascarado_por_outra_base(self):
        """Sem o campo oficial, a base é 0,0 — e nunca o valor original.

        Zero aqui é o comportamento explícito da tela (que exibe a base ao
        usuário e permite corrigi-la), não um desvio silencioso para outra
        grandeza contratual.
        """

        resultado = resultado_oficial()
        del resultado["valor_atualizado_contrato"]
        contexto = obter_contexto_valor_global(resultado)
        self.assertEqual(contexto["valor_total_atualizado"], 0.0)
        self.assertNotEqual(contexto["valor_total_atualizado"], contexto["valor_original"])

    def test_resultado_ausente_nao_levanta_e_nao_inventa_base(self):
        for vazio in (None, {}):
            contexto = obter_contexto_valor_global(vazio)
            self.assertEqual(contexto["valor_total_atualizado"], 0.0)
            self.assertEqual(contexto["valor_original"], 0.0)


class CalculoDaGarantiaTests(unittest.TestCase):
    def test_percentual_e_formula_da_garantia_permanecem_intactos(self):
        # Regra homologada do módulo: base × percentual, arredondada a centavos.
        self.assertIn(
            "garantia_exigida_total = round(valor_total_atualizado * percentual_garantia, 2)",
            GARANTIA,
        )
        self.assertIn("valor_garantia_original = round(valor_original * percentual_garantia, 2)", GARANTIA)
        self.assertIn("percentual_garantia = percentual_garantia_pct / 100", GARANTIA)

    def test_garantia_exigida_acompanha_a_base_atualizada(self):
        contexto = obter_contexto_valor_global(resultado_oficial())
        percentual_garantia = 5.0 / 100
        garantia_exigida_total = round(contexto["valor_total_atualizado"] * percentual_garantia, 2)
        self.assertEqual(garantia_exigida_total, 52_250.00)

    def test_rotulo_exibido_corresponde_a_base_utilizada(self):
        self.assertIn('"Valor Total Atualizado do Contrato"', GARANTIA)
        self.assertIn('"valor_total_atualizado_contrato": valor_total_atualizado', GARANTIA)


if __name__ == "__main__":
    unittest.main()
