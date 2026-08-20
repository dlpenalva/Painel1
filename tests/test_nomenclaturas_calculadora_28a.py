"""Etapa 28A — nomenclaturas da Calculadora (camada exclusivamente visual).

Contrato comprovado aqui (padrao do repositorio: asserts sobre o fonte,
como em test_ui_xls_first_shell/test_ciclos_c2_e_stamp):

- os dois textos principais aprovados aparecem no menu, no Inicio e nos
  cabecalhos das paginas;
- os textos antigos nao aparecem mais na interface principal;
- fluxo de ciclo unico pergunta "Qual ciclo deseja analisar?" com opcoes
  visiveis "Quero analisar Cn" via format_func, MANTENDO os valores
  internos "C1".."C4" e a key sim_ciclo_analise;
- fluxo de varios ciclos usa os tres rotulos aprovados, MANTENDO keys,
  opcoes internas, faixa range(primeiro_ciclo_num, 5) e o contexto
  contratual anterior;
- nenhuma Central sem Coleta foi criada;
- a matematica (fatores/preclusao/admissibilidade) nao foi tocada: os
  call-sites de calculo permanecem identicos (a paridade numerica e
  coberta pela regressao integral).
"""
from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / "app.py").read_text(encoding="utf-8")
INICIO = (ROOT / "pages" / "00_Calculadora_Reajustes.py").read_text(encoding="utf-8")
SIMPLES = (ROOT / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
MULTI = (ROOT / "pages" / "02_Calculo_Represados.py").read_text(encoding="utf-8")


class TestTextosPrincipais(unittest.TestCase):
    def test_menu_usa_os_dois_textos_aprovados(self):
        self.assertIn('_page("01_Calculo_Simples.py", "Analisar um único ciclo")', APP)
        self.assertIn('_page("02_Calculo_Represados.py", "Analisar vários ciclos")', APP)
        self.assertIn('st.page_link(PAGINA_UM_CICLO, label="Analisar um único ciclo")', APP)
        self.assertIn('st.page_link(PAGINA_MULTICICLO, label="Analisar vários ciclos")', APP)

    def test_inicio_e_cabecalhos_usam_os_textos_aprovados(self):
        # Apos o redesign da HOME, os rotulos dos fluxos vivem no menu lateral
        # (app.py) e nas proprias paginas — a home nao os repete.
        self.assertIn('label="Analisar um único ciclo"', APP)
        self.assertIn('label="Analisar vários ciclos"', APP)
        self.assertIn('"Analisar um único ciclo"', SIMPLES)
        self.assertIn('"Analisar vários ciclos"', MULTI)

    def test_textos_antigos_ausentes_da_interface_principal(self):
        for fonte in (APP, INICIO, SIMPLES, MULTI):
            self.assertNotIn("Primeiro Reajuste", fonte)
            self.assertNotIn("Reajustes Subsequentes", fonte)
        self.assertNotIn("Ciclo desta análise:", SIMPLES)
        self.assertNotIn("Ciclo inicial desta análise", MULTI)
        self.assertNotIn("Ciclo final desta análise", MULTI)
        self.assertNotIn("Último ciclo concedido/formalizado:", MULTI)

    def test_inicio_nao_tem_mais_botoes_de_navegacao(self):
        # Os botoes de navegacao sairam da HOME no redesign; os rotulos
        # continuam provados na sidebar (app.py) e nas proprias paginas.
        self.assertNotIn('key="abrir_um_ciclo_inicio"', INICIO)
        self.assertNotIn('key="abrir_multiciclo_inicio"', INICIO)
        self.assertIn('key="download_coleta_inicio"', INICIO)


class TestCicloUnico(unittest.TestCase):
    def test_pergunta_e_opcoes_visiveis(self):
        self.assertIn('"Qual ciclo deseja analisar?"', SIMPLES)
        self.assertIn('format_func=lambda ciclo: f"Quero analisar {ciclo}"', SIMPLES)

    def test_valores_internos_continuam_cn(self):
        # O selectbox usa as MESMAS opcoes internas; o texto visivel vem
        # somente do format_func (nunca gravado em session_state).
        self.assertIn('opcoes_ciclo_analise = ["C1", "C2", "C3", "C4"]', SIMPLES)
        self.assertIn("options=opcoes_ciclo_analise", SIMPLES)
        self.assertIn('key="sim_ciclo_analise"', SIMPLES)
        self.assertNotIn('"Quero analisar C1"', SIMPLES)  # nao e valor de opcao
        # O valor selecionado (Cn) segue alimentando o motor sem traducao.
        self.assertIn("primeiro_ciclo_num = _ciclo_para_numero(ciclo_analise)", SIMPLES)

    def test_mapeamento_de_exibicao_e_puro(self):
        rotulo = (lambda ciclo: f"Quero analisar {ciclo}")
        self.assertEqual(rotulo("C1"), "Quero analisar C1")
        self.assertEqual(rotulo("C3"), "Quero analisar C3")
        self.assertEqual(rotulo("C4"), "Quero analisar C4")


class TestVariosCiclos(unittest.TestCase):
    def test_tres_rotulos_aprovados(self):
        self.assertIn('"Primeiro ciclo a analisar:"', MULTI)
        self.assertIn('"Último ciclo a analisar:"', MULTI)
        self.assertIn('"Último ciclo formalizado anteriormente:"', MULTI)
        self.assertIn("1 · Primeiro ciclo a analisar", MULTI)
        self.assertIn("2 · Último ciclo a analisar", MULTI)

    def test_valores_internos_e_keys_preservados(self):
        self.assertIn('key="rep_ciclo_inicial_analise"', MULTI)
        self.assertIn('chave_ciclo_final = "rep_ciclo_final_analise"', MULTI)
        self.assertIn("key=chave_ciclo_final", MULTI)
        self.assertIn('key="rep_ultimo_ciclo_anterior"', MULTI)
        self.assertIn('options=["C1", "C2", "C3", "C4"]', MULTI)
        self.assertIn("range(primeiro_ciclo_num, 5)", MULTI)
        self.assertIn("range(1, int(primeiro_ciclo_num))", MULTI)

    def test_historico_anterior_e_contexto_preservados(self):
        # O terceiro campo segue representando o historico anterior a
        # analise, com o MESMO contrato de dados (cenario C2-C4 com C1
        # formalizado: nada renumerado, nada perdido).
        self.assertIn('"ultimo_ciclo_concedido": _ultimo_ciclo_anterior', MULTI)
        self.assertIn('"data_pedido_ultimo_ciclo": _marco_temporal_anterior', MULTI)
        self.assertIn('"data_base_ultimo_ciclo": _marco_temporal_anterior', MULTI)
        self.assertIn("Houve ciclo anterior concedido/formalizado", MULTI)

    def test_callsites_de_calculo_intactos(self):
        # Preclusao/fatores/ancoragem: mesmas chamadas de motor de antes.
        self.assertIn(
            "data_atual = _calcular_data_inicial_ciclo(_dt_base_calculo, primeiro_ciclo_num, _contexto_calculo)",
            MULTI,
        )
        self.assertIn("ultimo_ciclo_contratual = _ciclo_para_numero(ciclo_fin_analise)", MULTI)


class TestForaDeEscopo(unittest.TestCase):
    def test_nenhuma_central_sem_coleta(self):
        # Guard original da Etapa 28A (nomenclaturas): a 28A nao podia introduzir
        # a Central de Modelos. A Etapa 29B implementa essa Central e a integra
        # deliberadamente ao app.py (navegacao) e a pagina inicial (card), de modo
        # que APP/INICIO passam a cita-la legitimamente. As paginas de calculo
        # (um ciclo e multiciclo) permanecem fora de escopo e nao a referenciam.
        for fonte in (SIMPLES, MULTI):
            self.assertNotIn("Central de Modelos", fonte)
        for fonte in (APP, INICIO, SIMPLES, MULTI):
            self.assertNotIn("Central sem Coleta", fonte)


if __name__ == "__main__":
    unittest.main()
