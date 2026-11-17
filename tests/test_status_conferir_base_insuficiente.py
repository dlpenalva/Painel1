"""Regressao Cl8us 12.16 / Coleta 12.8 — status CONFERIR e BASE INSUFICIENTE.

Defeito 1: RESULTADOS_DETALHE!B3 (Coleta 12.7, PR #199) passou a emitir
CONFERIR e BASE INSUFICIENTE, mas o mapa de status do Python so conhecia
VALIDADO / VALIDADO COM RESSALVAS / ESTIMADO / REVISE. Um arquivo recalculado
no Excel aparecia como "status oficial indisponivel... recalcule e reenvie" e a
formalizacao como AGUARDA CONFIRMACAO (caso real 1-luc.xlsx).

Defeito 2: com o mesmo status, RESULTADOS!C13 continuava dizendo "Nenhuma
pendencia registrada pelas validacoes desta apuracao".

Os valores (VTA, retroativo reconhecido/potencial/em analise) nao dependem do
status: os testes comparam cada metodo com o mesmo payload em VALIDADO. A
igualdade exata da formula de C13 com o template e com a Coleta gerada fica com
tests/test_resultados_retro_reconhecido_potencial.py (descritor).
"""

from __future__ import annotations

import re
from pathlib import Path

import pytest

from _conferencias_execucao import FORMULAS_MENSAGENS_RESULTADOS
from _resultado_consolidado import (
    ORIGEM_STATUS_RESULTADOS,
    STATUS_BASE_INSUFICIENTE,
    STATUS_CONFERIR,
    STATUS_CONFIAVEL,
    STATUS_PENDENTE,
    montar_resultado_consolidado,
)

RAIZ = Path(__file__).resolve().parents[1]
STATUS_CONFERENCIA = (STATUS_CONFERIR, STATUS_BASE_INSUFICIENTE)
METODOS = ("financeiro", "pc", "consumidos")
_MODO_CONTROLE = {"financeiro": "financeiro", "pc": "pc", "consumidos": "d"}


def _payload(metodo: str, status: str, **extra) -> dict:
    resultado = {
        "diagnostico_coleta": {
            "metadados": {
                "status_resultados": {
                    "geral": status,
                    "valores": {"retroativo_oficial": 125.0},
                },
                "formula_status_resultados_presente": True,
            }
        },
        "valor_atualizado_contrato": 1_000.0,
        "valor_represado_a_pagar": 125.0 if metodo != "consumidos" else None,
        "controle": {"modo": _MODO_CONTROLE[metodo], "ciclo_vigente": "C1"},
        "memoria_por_ciclo": {"vta": {"metodo": metodo}},
        "totais_canonicos_pc": {
            "ate_o_corte": {
                "retroativo": 125.0,
                "valor_atualizado_em_analise": 40.0,
                "delta_potencial": 0.0,
            },
            "posterior_ao_corte": {"quantidade": 0, "valor_pc": 0.0},
        },
        "politica_entrega_segura": {
            "status": "PRONTO_PARA_VALIDACAO_FISCAL",
            "pode_confirmar": True,
            "pendencias": [],
            "bloqueios": [],
        },
        "reconciliacao_xls_python": {"status_geral": "CONCILIADO"},
        "formalizacao_bloqueada": False,
        "bloqueios_formalizacao": [],
    }
    resultado.update(extra)
    return resultado


def _valores(consolidado: dict) -> tuple:
    return (
        consolidado["vta"],
        consolidado["retroativo_reconhecido"],
        consolidado["retroativo_potencial"],
        consolidado["valor_atualizado_em_analise"],
    )


@pytest.mark.parametrize("metodo", METODOS)
@pytest.mark.parametrize("status", STATUS_CONFERENCIA)
def test_status_de_conferencia_e_reconhecido_sem_virar_outro(status, metodo):
    c = montar_resultado_consolidado(_payload(metodo, status))
    apuracao = c["status_apuracao"]
    assert apuracao["codigo"] == status
    assert apuracao["rotulo"] == c["status_confiabilidade"] == status
    assert apuracao["disponivel"] is True and apuracao["conclusivo"] is True
    assert apuracao["origem"] == ORIGEM_STATUS_RESULTADOS
    # Nao e VALIDADO, nao e REVISE (PENDENTE) e nao e indisponibilidade.
    assert c["status_confiabilidade"] not in (STATUS_CONFIAVEL, STATUS_PENDENTE)
    assert "não está disponível" not in c["mensagem_status"]
    assert "recalcule" not in c["mensagem_status"]
    assert status in c["mensagem_status"]
    assert any("não altera o VTA nem bloqueia" in r for r in c["ressalvas"])
    assert c["formalizacao"] == {
        "bloqueada": False,
        "status": "SEM BLOQUEIO",
        "mensagem": "Sem bloqueio explícito à formalização.",
    }


@pytest.mark.parametrize("metodo", METODOS)
@pytest.mark.parametrize("status", STATUS_CONFERENCIA)
def test_valores_por_metodo_nao_dependem_do_status(status, metodo):
    validado = montar_resultado_consolidado(_payload(metodo, "VALIDADO"))
    conferencia = montar_resultado_consolidado(_payload(metodo, status))
    assert _valores(conferencia) == _valores(validado)


def test_retroativo_reconhecido_por_metodo_preservado():
    por_metodo = {
        m: montar_resultado_consolidado(_payload(m, STATUS_CONFERIR))
        for m in METODOS
    }
    assert por_metodo["financeiro"]["retroativo_reconhecido"] == 125.0
    assert por_metodo["pc"]["retroativo_reconhecido"] == 125.0
    assert por_metodo["pc"]["valor_atualizado_em_analise"] == 40.0
    # VTA-C2: Consumidos continua sem fonte independente de pagamento.
    assert por_metodo["consumidos"]["retroativo_reconhecido"] is None
    assert por_metodo["financeiro"]["retroativo_potencial"] is None


@pytest.mark.parametrize("status", STATUS_CONFERENCIA)
def test_zero_legitimo_continua_zero(status):
    payload = _payload("financeiro", status, valor_represado_a_pagar=0.0)
    c = montar_resultado_consolidado(payload)
    assert c["retroativo_reconhecido"] == 0.0
    assert c["retroativo_reconhecido"] is not None


@pytest.mark.parametrize("status", STATUS_CONFERENCIA)
def test_ausencia_continua_ausencia(status):
    payload = _payload("financeiro", status, valor_represado_a_pagar=None)
    payload["diagnostico_coleta"]["metadados"]["status_resultados"]["valores"] = {}
    c = montar_resultado_consolidado(payload)
    assert c["retroativo_reconhecido"] is None
    # Status reconhecido nao fabrica VTA: sem VTA o nucleo fica indisponivel.
    sem_vta = montar_resultado_consolidado(
        _payload("financeiro", status, valor_atualizado_contrato=None)
    )
    assert sem_vta["status_confiabilidade"] == STATUS_PENDENTE
    assert sem_vta["status_apuracao"]["conclusivo"] is False


@pytest.mark.parametrize("metodo", METODOS)
@pytest.mark.parametrize("status", STATUS_CONFERENCIA)
def test_bloqueio_economico_preservado(status, metodo):
    motivo = "PC sem início de efeito financeiro: formalização bloqueada."
    payload = _payload(
        metodo, status,
        formalizacao_bloqueada=True, bloqueios_formalizacao=[motivo],
    )
    c = montar_resultado_consolidado(payload)
    assert c["formalizacao"]["bloqueada"] is True
    assert c["formalizacao"]["status"] == "BLOQUEADA"
    assert c["formalizacao"]["mensagem"] == motivo
    assert c["bloqueios"] == [motivo]
    # O eixo da apuracao continua sendo a conclusao do XLS.
    assert c["status_confiabilidade"] == status


def test_revise_e_texto_desconhecido_continuam_fail_closed():
    revise = montar_resultado_consolidado(_payload("financeiro", "REVISE"))
    assert revise["status_confiabilidade"] == STATUS_PENDENTE
    assert revise["status_apuracao"]["codigo"] == "REVISE"
    for texto in ("CONFERIR PARCIAL", "BASE", "EM CONFERÊNCIA"):
        c = montar_resultado_consolidado(_payload("financeiro", texto))
        assert c["status_apuracao"]["disponivel"] is False, texto
        assert c["status_confiabilidade"] == STATUS_PENDENTE, texto


def test_painel_tem_classe_visual_de_ressalva_para_os_dois_status():
    pagina = (RAIZ / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")
    assert '"CONFERIR": "ressalvas"' in pagina
    assert '"BASE INSUFICIENTE": "ressalvas"' in pagina


def test_c13_so_troca_o_texto_quando_o_status_pede_conferencia():
    formula = FORMULAS_MENSAGENS_RESULTADOS["C13"]
    # Gate: apuracao sem pendencias (regra de A7) E status de conferencia.
    assert formula.startswith(
        '=IF(AND(LEFT(RESULTADOS_DETALHE!$A$7,4)="APUR",'
        'OR(STATUS_RESULTADOS="CONFERIR",STATUS_RESULTADOS="BASE INSUFICIENTE")),'
    )
    # Fora do gate, o texto anterior (T67) sai intacto.
    assert formula.endswith(",MEMORIA_RESULTADOS!$T$67)")
    # O texto de potencial (T67 diferente de B7) continua anexado.
    assert ('&IF(MEMORIA_RESULTADOS!$T$67=RESULTADOS_DETALHE!$B$7,"",'
            '" "&MEMORIA_RESULTADOS!$T$67)') in formula
    assert "Nenhuma pendência registrada" not in formula
    literais = re.findall(r'"((?:[^"]|"")*)"', formula)
    assert all(len(t) <= 255 for t in literais)  # limite do Excel por literal
    assert sum("não altera o VTA nem bloqueia a formalização" in t
               for t in literais) == 2
    sem_literais = re.sub(r'"(?:[^"]|"")*"', '""', formula)
    assert sem_literais.isascii()
    assert sem_literais.count("(") == sem_literais.count(")")


# =========================================================================== #
# Pendencias finais do PR #200
# (1) Sumario Executivo: "Em analise" fora do metodo PCs -> "Nao aplicavel".
# (2) RESULTADOS!B12:G13: CONFERIR/BASE INSUFICIENTE fora do verde de VALIDADO.
# =========================================================================== #
import copy  # noqa: E402
import io  # noqa: E402
import shutil  # noqa: E402

from openpyxl import load_workbook  # noqa: E402

from _sumario_executivo import (  # noqa: E402
    NAO_APLICAVEL,
    _em_analise_nao_aplicavel,
    gerar_sumario_executivo,
    montar_dados_sumario_executivo,
)
from test_sumario_executivo import (  # noqa: E402
    leitura_multiciclo_pc,
    leitura_simples_financeiro,
)

TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
REGRA_RESSALVA = (
    'AND(LEFT($C$12,4)="APUR",'
    'OR(STATUS_RESULTADOS="CONFERIR",STATUS_RESULTADOS="BASE INSUFICIENTE"))'
)


def _com_metodo(leitura: dict, metodo) -> dict:
    leitura = copy.deepcopy(leitura)
    leitura["resultados_xls"] = {
        "disponivel": True, "valores": {"METODO_RETROATIVO": metodo},
    }
    return leitura


def _fin(em_analise=0.0, pcs=()):
    return {
        "pc_por_ciclo": list(pcs),
        "retroativo_em_analise": em_analise,
        "retroativo_por_estado": {
            "em análise": em_analise, "reconhecido": 3958.54, "pago": 0.0,
        },
    }


@pytest.mark.parametrize("metodo,fin,esperado", [
    ("Financeiro", _fin(), True),
    ("Itens", _fin(), True),
    ("PCs", _fin(), False),                      # zero apurado: R$ 0,00
    (None, _fin(), False),                       # metodo desconhecido
    ("SELECIONE O METODO EM CONTROLE!B1", _fin(), False),
    ("Financeiro", _fin(em_analise=12.34), False),  # valor nunca e ocultado
    ("Financeiro", _fin(pcs=[{"ciclo": "C1"}]), False),
])
def test_em_analise_nao_aplicavel_so_fora_de_pcs_sem_valor(metodo, fin, esperado):
    assert _em_analise_nao_aplicavel(fin, metodo) is esperado


def _quadro_retroativo(pdf: bytes) -> dict:
    fitz = pytest.importorskip("fitz")
    quadro = {}
    for page in fitz.open(stream=pdf, filetype="pdf"):
        linhas: dict = {}
        for bloco in page.get_text("dict")["blocks"]:
            for linha in bloco.get("lines", []):
                y = round((linha["bbox"][1] + linha["bbox"][3]) / 2)
                texto = "".join(s["text"] for s in linha["spans"]).strip()
                if texto:
                    linhas.setdefault(y, []).append((linha["bbox"][0], texto))
        for y in sorted(linhas):
            partes = [t for _, t in sorted(linhas[y])]
            if partes[0] in ("Em análise", "Reconhecido", "Pago") and len(partes) == 2:
                quadro[partes[0]] = partes[1]
    return quadro


def test_pdf_financeiro_mostra_nao_aplicavel_e_preserva_valores():
    sem = montar_dados_sumario_executivo(leitura_simples_financeiro())
    com = montar_dados_sumario_executivo(
        _com_metodo(leitura_simples_financeiro(), "Financeiro"))
    assert com["financeiro"]["em_analise_nao_aplicavel"] is True
    # Somente apresentacao: valores, estados e sintese identicos.
    fin_sem = {k: v for k, v in sem["financeiro"].items()
               if k != "em_analise_nao_aplicavel"}
    fin_com = {k: v for k, v in com["financeiro"].items()
               if k != "em_analise_nao_aplicavel"}
    assert fin_com == fin_sem
    assert com["sintese"] == sem["sintese"]
    q_sem = _quadro_retroativo(gerar_sumario_executivo(leitura_simples_financeiro()))
    q_com = _quadro_retroativo(gerar_sumario_executivo(
        _com_metodo(leitura_simples_financeiro(), "Financeiro")))
    assert q_com["Em análise"] == NAO_APLICAVEL
    assert q_sem["Em análise"].startswith("R$")  # sem metodo: comportamento anterior
    assert q_com["Reconhecido"] == q_sem["Reconhecido"]
    assert q_com["Pago"] == q_sem["Pago"]


def test_pdf_metodo_pc_preserva_valor_em_analise():
    sem = _quadro_retroativo(gerar_sumario_executivo(leitura_multiciclo_pc()))
    com = _quadro_retroativo(gerar_sumario_executivo(
        _com_metodo(leitura_multiciclo_pc(), "PCs")))
    dados = montar_dados_sumario_executivo(_com_metodo(leitura_multiciclo_pc(), "PCs"))
    assert dados["financeiro"]["em_analise_nao_aplicavel"] is False
    assert com == sem
    assert com["Em análise"].startswith("R$")


def test_pdf_pcs_com_metodo_financeiro_nao_oculta_valor():
    # Evidencia de PC presente: mesmo com metodo != PCs, o valor e exibido.
    dados = montar_dados_sumario_executivo(
        _com_metodo(leitura_multiciclo_pc(), "Financeiro"))
    assert dados["financeiro"]["pc_por_ciclo"]
    assert dados["financeiro"]["em_analise_nao_aplicavel"] is False


def _regras_faixa(wb):
    ws = wb["RESULTADOS"]
    faixa = next(f for f in ws.conditional_formatting if str(f.sqref) == "B12:G13")
    return sorted(faixa.rules, key=lambda r: r.priority)


def _gerada():
    from _coleta_oficial import obter_coleta_oficial_bytes
    return load_workbook(io.BytesIO(obter_coleta_oficial_bytes()))


@pytest.mark.parametrize("origem", ["template", "coleta_gerada"])
def test_faixa_situacao_conferir_em_ambar_antes_do_verde(origem):
    wb = load_workbook(TEMPLATE) if origem == "template" else _gerada()
    regras = _regras_faixa(wb)
    assert [r.formula[0] for r in regras] == [
        'LEFT($C$12,4)="REVI"',
        REGRA_RESSALVA,
        'AND(LEFT($C$12,4)="APUR",N($D$9)>0)',
        'AND(LEFT($C$12,4)="APUR",N($D$9)<=0)',
    ]
    ressalva = regras[1]
    assert not ressalva.stopIfTrue  # convencao da aba executiva
    assert ressalva.dxf.fill.bgColor.rgb == "FFFFEB9C"
    assert ressalva.dxf.font.color.rgb == "FF9C5700"
    # Verde de VALIDADO e demais estados preservados.
    assert regras[0].dxf.fill.bgColor.rgb == "FFFFC7CE"
    assert regras[2].dxf.fill.bgColor.rgb == "FFEEF3CF"
    assert regras[3].dxf.fill.bgColor.rgb == "FFC6EFCE"
    # Prioridades unicas na aba (nenhuma colisao apos a renumeracao).
    prioridades = [r.priority for f in wb["RESULTADOS"].conditional_formatting
                   for r in f.rules]
    assert len(prioridades) == len(set(prioridades))


def test_ferramenta_da_cor_e_idempotente(tmp_path):
    from tools.aplicar_cor_status_resultados import aplicar
    copia = tmp_path / "coleta.xlsx"
    shutil.copyfile(TEMPLATE, copia)
    antes = copia.read_bytes()
    assert aplicar(copia) is False
    assert copia.read_bytes() == antes


def test_ferramenta_da_cor_aceita_ordem_e_dxf_do_excel():
    # Arquivo salvo pelo Excel (caso real 1-luc, Coleta 12.6): <cfRule> fora de
    # ordem e dxfId renumerados. Precedencia = priority; dxf pela formula.
    from tools.aplicar_cor_status_resultados import NS, patch_planilha
    xml = (
        f'<worksheet xmlns="{NS}"><sheetData/>'
        '<conditionalFormatting sqref="G6">'
        '<cfRule type="expression" dxfId="2" priority="2"><formula>'
        'ISNUMBER(SEARCH("ESTIMAD",$G$6))</formula></cfRule>'
        '<cfRule type="expression" dxfId="0" priority="1"><formula>'
        'OR($G$6="VALIDADO",$G$6="RECONCILIADO")</formula></cfRule>'
        '</conditionalFormatting>'
        '<conditionalFormatting sqref="B12:G13">'
        '<cfRule type="expression" dxfId="32" priority="38"><formula>'
        'AND(LEFT($C$12,4)="APUR",N($D$9)&lt;=0)</formula></cfRule>'
        '<cfRule type="expression" dxfId="31" priority="36"><formula>'
        'LEFT($C$12,4)="REVI"</formula></cfRule>'
        '<cfRule type="expression" dxfId="30" priority="37"><formula>'
        'AND(LEFT($C$12,4)="APUR",N($D$9)&gt;0)</formula></cfRule>'
        '</conditionalFormatting></worksheet>'
    ).encode()
    novo = patch_planilha(xml)
    texto = novo.decode()
    regras = re.findall(r'dxfId="(\d+)" priority="(\d+)">'
                        r'<formula>([^<]+)</formula>', texto)
    por_formula = {f: (d, int(p)) for d, p, f in regras}
    assert por_formula[REGRA_RESSALVA] == ("2", 37)          # dxf do ESTIMADO
    assert por_formula['LEFT($C$12,4)="REVI"'] == ("31", 36)
    assert por_formula['AND(LEFT($C$12,4)="APUR",N($D$9)&gt;0)'] == ("30", 38)
    assert por_formula['AND(LEFT($C$12,4)="APUR",N($D$9)&lt;=0)'] == ("32", 39)
    assert patch_planilha(novo) is None                      # idempotente
