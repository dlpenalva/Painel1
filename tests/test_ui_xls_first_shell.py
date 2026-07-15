from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]
APP = (ROOT / "app.py").read_text(encoding="utf-8")
INICIO = (ROOT / "pages" / "00_Calculadora_Reajustes.py").read_text(encoding="utf-8")
SIMPLES = (ROOT / "pages" / "01_Calculo_Simples.py").read_text(encoding="utf-8")
MULTI = (ROOT / "pages" / "02_Calculo_Represados.py").read_text(encoding="utf-8")
DOCUMENTOS = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")
UI = (ROOT / "_ui_utils.py").read_text(encoding="utf-8")
UI_CAPACIDADES = (ROOT / "_ui_capacidades.py").read_text(encoding="utf-8")
CAPACIDADES = (ROOT / "_capacidades_apuracao.py").read_text(encoding="utf-8")
CENTRAL = (ROOT / "pages" / "06_Central_Arquivos.py").read_text(encoding="utf-8")
GARANTIA = (ROOT / "pages" / "05_Garantia.py").read_text(encoding="utf-8")
SANEADOR = (ROOT / "pages" / "10_Saneador.py").read_text(encoding="utf-8")
PREVISAO = (ROOT / "pages" / "12_Adequacao_Orcamentaria.py").read_text(encoding="utf-8")


class TestCascaXlsFirst(unittest.TestCase):
    def test_menu_principal_tem_somente_as_quatro_rotas_operacionais(self):
        self.assertIn('st.page_link(PAGINA_INICIO, label="Início")', APP)
        self.assertIn('st.page_link(PAGINA_UM_CICLO, label="Calculadora 1 ciclo")', APP)
        self.assertIn('st.page_link(PAGINA_MULTICICLO, label="Calculadora multiciclo")', APP)
        self.assertIn('st.page_link(PAGINA_UPLOAD, label="Upload e docs")', APP)
        self.assertIn('position="hidden"', APP)

    def test_menu_replica_rotulos_e_densidade_do_modelo_3(self):
        self.assertIn('>Piloto controlado</div>', APP)
        self.assertIn('>Documentos</div>', APP)
        self.assertNotIn('>XLS preenchido</div>', APP)
        self.assertIn('font-weight: 760;', APP)
        self.assertIn('margin-block: -.6rem;', APP)

    def test_modulos_legados_permanecem_registrados_mas_fora_do_menu(self):
        self.assertNotIn('st.expander("Ferramentas complementares", expanded=False)', APP)
        self.assertNotIn('for pagina in PAGINAS_AUXILIARES:', APP)
        self.assertIn('("04_Relatorio_Global.py", "Relatórios")', APP)
        self.assertIn('("13_DOU.py", "DOU")', APP)

    def test_area_principal_usa_fundo_neutro_e_sidebar_mantem_azul(self):
        self.assertIn('--cl8us-sidebar: #C6D9E8;', APP)
        self.assertIn('--cl8us-main-start: #FBF8F1;', APP)
        self.assertIn('--cl8us-main-end: #F2ECE1;', APP)
        self.assertNotIn('--cl8us-bg-start:', APP)

    def test_inicio_expoe_quatro_boxes_e_os_destinos_corretos(self):
        for numero in range(1, 5):
            self.assertEqual(INICIO.count(f'"{numero} ·'), 1)
        self.assertIn('st.switch_page("pages/03_Valor_Global.py")', INICIO)
        self.assertIn('st.switch_page("pages/01_Calculo_Simples.py")', INICIO)
        self.assertIn('st.switch_page("pages/02_Calculo_Represados.py")', INICIO)

    def test_modelo_xls_e_a_fonte_do_download_inicial(self):
        self.assertIn("CAMINHO_MODELO_COLETA", INICIO)
        self.assertIn("NOME_ARQUIVO_COLETA", INICIO)
        self.assertIn('file_name=NOME_ARQUIVO_COLETA', INICIO)
        self.assertIn('"Baixar Coleta_Reajuste.xlsx"', INICIO)

    def test_inicio_nao_reintroduz_seletor_intermediario_da_versao_antiga(self):
        self.assertNotIn("executar_motor", INICIO)
        self.assertNotIn("fluxo_query", INICIO)
        self.assertNotIn("A análise envolve mais de um ciclo", INICIO)
        self.assertNotIn("runpy", INICIO)

    def test_inicio_nao_exibe_aviso_excluido(self):
        self.assertNotIn("Se as informações forem parciais", INICIO)
        self.assertNotIn("A ausência de base segura", INICIO)

    def test_cabecalho_em_box_e_usado_nas_quatro_rotas(self):
        self.assertIn("def render_cabecalho_pagina", UI)
        self.assertIn("cl8us-page-header", APP)
        for pagina in (INICIO, SIMPLES, MULTI, DOCUMENTOS):
            self.assertIn("render_cabecalho_pagina(", pagina)

    def test_documentos_fica_enxuto_antes_do_upload_e_preserva_o_motor(self):
        self.assertNotIn("Mesa GCC", DOCUMENTOS)
        self.assertIn('"Painel da Apuração Contratual"', DOCUMENTOS)
        self.assertIn("1 · Baixar arquivo de trabalho", DOCUMENTOS)
        self.assertIn("2 · Enviar Coleta_Reajuste.xlsx preenchido", DOCUMENTOS)
        self.assertIn("CAMINHO_MODELO_COLETA.read_bytes()", DOCUMENTOS)
        self.assertIn('key="upload_coleta_documentos"', DOCUMENTOS)
        self.assertIn("if arquivo is None:", DOCUMENTOS)
        self.assertIn("st.stop()", DOCUMENTOS)
        self.assertIn('if st.button("Validar Coleta Preenchida"', DOCUMENTOS)
        self.assertLess(DOCUMENTOS.index("if arquivo is None:"), DOCUMENTOS.index("if arquivo is not None:"))

    def test_modo_um_usa_botoes_de_download_na_cor_padrao(self):
        coleta = SIMPLES[SIMPLES.index('label="Baixar Coleta_Reajuste.xlsx"'):]
        coleta = coleta[:coleta.index(")")]
        rascunho = SIMPLES[SIMPLES.index('label="Baixar rascunho (.txt)"'):]
        rascunho = rascunho[:rascunho.index(")")]
        self.assertNotIn('type="primary"', coleta)
        self.assertNotIn('type="primary"', rascunho)

    def test_upload_nao_exibe_textos_excluidos(self):
        self.assertNotIn("render_aviso_privacidade", DOCUMENTOS)
        self.assertNotIn("A coleta possui base para a próxima etapa", DOCUMENTOS)
        self.assertNotIn("O XLS consolidou os quatro eixos", DOCUMENTOS)

    def test_upload_religa_processamento_e_expoe_os_oito_arquivos(self):
        self.assertIn("adaptar_coleta_reajuste_para_documentos(conteudo)", DOCUMENTOS)
        self.assertNotIn('elif diagnostico.get("valido"):', DOCUMENTOS)
        self.assertNotIn("_resultado_processado_pela_web", DOCUMENTOS)
        self.assertIn("render_status_apuracao", DOCUMENTOS)
        self.assertIn("render_status_documentos", DOCUMENTOS)
        self.assertIn("Ações sobre os documentos", DOCUMENTOS)
        for arquivo in (
            "Planilha Executiva",
            "Valores Unitários e Totais por Ciclo",
            "Mapa dos Marcos",
            "Relatório Executivo",
            "Minuta de Apostilamento",
            "Checklist Processual",
            "Garantia Contratual",
            "Saneador",
        ):
            self.assertIn(arquivo, DOCUMENTOS + UI_CAPACIDADES + CAPACIDADES + CENTRAL)

    def test_referencias_antigas_foram_removidas_da_interface(self):
        self.assertNotIn("Mesa GCC", DOCUMENTOS)
        self.assertNotIn("Calcule os marcos na web", INICIO)
        self.assertNotIn("RESULTADOS:", DOCUMENTOS)

    def test_linha_do_tempo_aceita_colunas_opcionais_ausentes(self):
        self.assertIn("def valores_aditivo(nome_coluna, padrao):", DOCUMENTOS)
        self.assertIn("return [padrao] * len(aditivos_temp.index)", DOCUMENTOS)
        self.assertIn("if isinstance(coluna, pd.DataFrame):", DOCUMENTOS)
        self.assertNotIn('aditivos_temp.get("Tratamento do aditivo", "").apply', DOCUMENTOS)

    def test_central_e_hub_compacto_dos_oito_documentos_oficiais(self):
        documentos = (
            "Relatório Executivo",
            "Minuta de Apostilamento",
            "Despacho Saneador",
            "Previsão Orçamentária",
            "Extrato para Publicação (DOU)",
            "Sumário do Reajuste",
            "Itens por Ciclo",
            "Garantia Contratual",
        )
        catalogo = CENTRAL[: CENTRAL.index("def aplicar_css_central")]
        self.assertEqual(catalogo.count('"nome":'), 8)
        for documento in documentos:
            self.assertIn(documento, CENTRAL)
        self.assertIn("render_status_entradas(CAPACIDADES)", CENTRAL)
        self.assertIn('ordem = ("financeiro", "itens", "pcs", "consumidos", "remanescentes")', CENTRAL)
        self.assertIn("st.container(border=True)", CENTRAL)
        self.assertIn('label="Gerar e baixar"', CENTRAL)
        self.assertIn('"Baixar documento"', CENTRAL)
        self.assertNotIn("render_status_documentos", CENTRAL)
        self.assertNotIn("central-header", CENTRAL)
        self.assertNotIn("Pendente de dados", CENTRAL)
        self.assertNotIn("Disponível com ressalvas", CENTRAL)

    def test_central_bloqueia_somente_estrutura_critica_e_reaproveita_geradores(self):
        self.assertIn('CAPACIDADES.get("estruturalmente_valido", True)', CENTRAL)
        self.assertIn('"Corrigir XLS para continuar"', CENTRAL)
        self.assertIn('st.session_state["arquivo_garantia_pdf"] = pdf_bytes', GARANTIA)
        self.assertIn('st.session_state["arquivo_previsao_orcamentaria_docx"] = docx_bytes', PREVISAO)
        self.assertIn('st.session_state["arquivo_saneador_docx"] = docx_bytes', SANEADOR)
        self.assertIn("[campo a preencher]", SANEADOR)


if __name__ == "__main__":
    unittest.main()
