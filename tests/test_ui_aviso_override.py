import ast
from copy import deepcopy
from pathlib import Path
from unittest.mock import patch

from _ui_utils import render_avisos_override_efeito_financeiro


ROOT = Path(__file__).resolve().parents[1]
PAGINA = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")
PREFIXO = "Marcacao de efeito financeiro ajustada manualmente:"


def _diagnostico(*avisos):
    return {"avisos": list(avisos), "valido": True, "pronto_para_consolidar": True}


def test_um_override_aparece_uma_vez_e_nao_interrompe_o_fluxo():
    aviso = f"{PREFIXO} C1 - 04/2024."
    with patch("_ui_utils.st.warning") as warning, patch("_ui_utils.st.stop") as stop:
        exibidos = render_avisos_override_efeito_financeiro(_diagnostico(aviso))

    assert exibidos == (aviso,)
    warning.assert_called_once_with(aviso)
    stop.assert_not_called()


def test_dois_overrides_distintos_aparecem_sem_duplicacao():
    primeiro = f"{PREFIXO} C1 - 04/2024."
    segundo = f"{PREFIXO} C2 - 04/2025."
    with patch("_ui_utils.st.warning") as warning:
        exibidos = render_avisos_override_efeito_financeiro(
            _diagnostico(primeiro, primeiro, segundo)
        )

    assert exibidos == (primeiro, segundo)
    assert [chamada.args[0] for chamada in warning.call_args_list] == [primeiro, segundo]


def test_sem_override_nao_exibe_aviso():
    with patch("_ui_utils.st.warning") as warning:
        exibidos = render_avisos_override_efeito_financeiro(
            _diagnostico("Ainda nao ha valores mensais preenchidos pelo fiscal.")
        )

    assert exibidos == ()
    warning.assert_not_called()


def test_formato_legado_nao_produz_falso_aviso():
    with patch("_ui_utils.st.warning") as warning:
        exibidos = render_avisos_override_efeito_financeiro(
            {"avisos": ["Arquivo legado aceito sem metadado de inicio do efeito."]}
        )

    assert exibidos == ()
    warning.assert_not_called()


def test_aviso_e_renderizado_antes_dos_cards_e_do_bloqueio_independente():
    arvore = ast.parse(PAGINA)
    chamadas = [
        no
        for no in ast.walk(arvore)
        if isinstance(no, ast.Call) and isinstance(no.func, ast.Name)
    ]
    linha_aviso = next(
        no.lineno for no in chamadas if no.func.id == "render_avisos_override_efeito_financeiro"
    )
    linha_cards = next(
        no.lineno for no in chamadas if no.func.id == "render_documentos_funcionais_upload"
    )
    linha_erro = next(
        no.lineno
        for no in ast.walk(arvore)
        if isinstance(no, ast.Call)
        and isinstance(no.func, ast.Attribute)
        and isinstance(no.func.value, ast.Name)
        and no.func.value.id == "st"
        and no.func.attr == "error"
        and any(
            isinstance(argumento, ast.Constant)
            and "A coleta não pôde ser processada com segurança." in str(argumento.value)
            for argumento in no.args
        )
    )

    assert linha_aviso < linha_cards < linha_erro


def test_aviso_nao_altera_resultado_valido_nem_gate_documental():
    aviso = f"{PREFIXO} C1 - 04/2024."
    resultado = {
        "retroativo_total": 125.50,
        "vta": 1500.00,
        "capacidades": {
            "documentos": {
                "planilha_executiva": {"habilitado": True},
                "relatorio_executivo": {"habilitado": True},
            }
        },
    }
    antes = deepcopy(resultado)

    with patch("_ui_utils.st.warning"), patch("_ui_utils.st.stop") as stop:
        render_avisos_override_efeito_financeiro(_diagnostico(aviso))

    assert resultado == antes
    assert all(item["habilitado"] for item in resultado["capacidades"]["documentos"].values())
    stop.assert_not_called()


def test_g_vazio_continua_fora_da_camada_de_aviso():
    with patch("_ui_utils.st.warning") as warning:
        exibidos = render_avisos_override_efeito_financeiro({"avisos": [], "valido": False})

    assert exibidos == ()
    warning.assert_not_called()
