from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
PAGINA = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")
CONSOLIDADO = (ROOT / "_resultado_consolidado.py").read_text(encoding="utf-8")


def _bloco_status_css() -> str:
    inicio = PAGINA.index(".resultado-status-bloqueado")
    return PAGINA[inicio:PAGINA.index("}", inicio) + 1]


def test_estado_interno_bloqueado_permanece_inalterado():
    assert 'STATUS_BLOQUEADO = "BLOQUEADO"' in CONSOLIDADO
    assert 'status = STATUS_BLOQUEADO' in CONSOLIDADO
    assert '"bloqueada": bloqueado' in CONSOLIDADO
    assert '"BLOQUEADA"' in CONSOLIDADO
    assert 'resultado.get("formalizacao_bloqueada")' in CONSOLIDADO
    assert 'resultado.get("bloqueios_formalizacao")' in CONSOLIDADO


def test_selo_principal_bloqueado_vira_em_conferencia_so_na_ui():
    assert 'status_em_conferencia = status == "BLOQUEADO"' in PAGINA
    assert 'formalizacao_exibicao = "EM CONFERÊNCIA"' in PAGINA
    assert 'colunas_segunda_linha.append(("Formalização", formalizacao_exibicao))' in PAGINA


def test_box_usa_rotulo_e_mensagem_solicitados():
    assert '"PENDENTE DE CONFERÊNCIA" if status_em_conferencia else status' in PAGINA
    assert (
        '"Pode haver divergência entre a planilha e a apuração do retroativo "'
        in PAGINA
    )
    assert '"por Pedidos de Compra."' in PAGINA
    assert "Status de confiabilidade: {html.escape(str(status_exibicao))}" in PAGINA


def test_estado_bloqueado_usa_ambar_e_nao_vermelho():
    css = _bloco_status_css()
    assert "background:#FFF8E6" in css
    assert "border-color:#D69E00" in css
    assert "color:#713F12" in css
    for vermelho in ("#FFF1F2", "#C62828", "#7F1D1D"):
        assert vermelho not in css


def test_detalhe_tecnico_permanece_recolhido_no_expander():
    assert 'detalhes = consolidado.get("bloqueios")' in PAGINA
    assert 'with st.expander("Ver fundamentos do status")' in PAGINA
