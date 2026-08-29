from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PAGINA = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")
CONSOLIDADO = (ROOT / "_resultado_consolidado.py").read_text(encoding="utf-8")


def _bloco_status_css() -> str:
    inicio = PAGINA.index(".resultado-status-bloqueado")
    return PAGINA[inicio:PAGINA.index("}", inicio) + 1]


def test_estado_interno_bloqueado_permanece_inalterado():
    """STATUS-CANON-1: o bloqueio nao sumiu — mudou de eixo.

    Ele deixou de ser status DA APURACAO e passou a viver integralmente na
    FORMALIZACAO, sem perder nenhuma das fontes que o originam.
    """
    assert 'STATUS_BLOQUEADO = "BLOQUEADO"' in CONSOLIDADO
    assert '"bloqueada": bloqueado' in CONSOLIDADO
    assert '"BLOQUEADA"' in CONSOLIDADO
    assert 'resultado.get("formalizacao_bloqueada")' in CONSOLIDADO
    assert 'resultado.get("bloqueios_formalizacao")' in CONSOLIDADO
    # A apuracao nunca mais e carimbada de BLOQUEADO pela politica.
    assert 'status = STATUS_BLOQUEADO' not in CONSOLIDADO


def test_selo_principal_bloqueado_vira_em_conferencia_so_na_ui():
    assert 'status_em_conferencia = bool(formalizacao.get("bloqueada"))' in PAGINA
    assert 'formalizacao_exibicao = "EM CONFERÊNCIA"' in PAGINA
    assert 'colunas_segunda_linha.append(("Formalização", formalizacao_exibicao))' in PAGINA


def test_box_usa_rotulo_e_mensagem_solicitados():
    """O box principal passou a nomear o que ele de fato e: o status DA APURACAO."""
    assert (
        '"Pode haver divergência entre a planilha e a apuração do retroativo "'
        in PAGINA
    )
    assert '"por Pedidos de Compra."' in PAGINA
    assert "Status da apuração: {html.escape(str(status_exibicao))}" in PAGINA
    assert "Status de confiabilidade:" not in PAGINA


def test_formalizacao_e_ressalva_tem_linhas_proprias():
    """STATUS-CANON-1: apuracao, formalizacao e ressalva sao eixos distintos."""
    assert 'if formalizacao_exibicao != "SEM BLOQUEIO":' in PAGINA
    assert "Formalização: {html.escape(str(formalizacao_exibicao))}" in PAGINA
    assert 'rotulo_ressalva = (' in PAGINA
    assert '"Ressalva" if len(ressalvas) == 1 else f"Ressalvas ({len(ressalvas)})"' in PAGINA


def test_status_principal_espelha_o_vocabulario_da_aba_resultados():
    """O painel nao inventa vocabulario proprio para a conclusao do XLS."""
    assert 'STATUS_CONFIAVEL = "VALIDADO"' in CONSOLIDADO
    assert '"VALIDADO": STATUS_CONFIAVEL' in CONSOLIDADO
    assert '"REVISE": STATUS_PENDENTE' in CONSOLIDADO
    assert '"ESTIMADO": STATUS_ESTIMADO' in CONSOLIDADO
    for rotulo in ('"VALIDADO": "confiavel"', '"ESTIMADO": "ressalvas"'):
        assert rotulo in PAGINA


def test_estado_bloqueado_usa_ambar_e_nao_vermelho():
    css = _bloco_status_css()
    assert "background:#FFF8E6" in css
    assert "border-color:#D69E00" in css
    assert "color:#713F12" in css
    for vermelho in ("#FFF1F2", "#C62828", "#7F1D1D"):
        assert vermelho not in css


def test_detalhe_tecnico_permanece_recolhido_no_expander():
    # Bloqueios, ressalvas e informacoes somam no expander: nenhum fundamento e
    # descartado por causa da separacao dos eixos.
    assert 'list(consolidado.get("bloqueios") or [])' in PAGINA
    assert "+ list(ressalvas)" in PAGINA
    assert "+ list(informacoes)" in PAGINA
    assert 'with st.expander("Ver fundamentos do status")' in PAGINA


def test_informacao_tem_linha_neutra_propria():
    """STATUS-CANON-1.1: execucao zero e informacao, nao ressalva."""
    assert 'informacoes = consolidado.get("informacoes") or []' in PAGINA
    assert "Informações da apuração" in PAGINA
    assert "resultado-status-informativo" in PAGINA
    # Tom neutro: nem verde de validado, nem ambar de pendencia.
    inicio = PAGINA.index(".resultado-status-informativo")
    css = PAGINA[inicio:PAGINA.index("}", inicio) + 1]
    assert "background:#F1F5F9" in css
    for cor_de_alerta in ("#FFF8E6", "#D69E00", "#ECFDF3", "#16803A"):
        assert cor_de_alerta not in css
