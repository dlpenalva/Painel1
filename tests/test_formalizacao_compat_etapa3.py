"""ETAPA 3/03 — formalizacao segura de Coletas compatibilizadas.

A Etapa 2 reconhece a linhagem, adapta em memoria e recompoe os derivados pela
regra vigente, mas deixou os gates documentais fechados. Aqui se prova que a
formalizacao so e liberada por EVIDENCIA TECNICA produzida pelo proprio
mecanismo de compatibilidade (nunca por nome, data, versao ou tamanho da
diferenca), e que cada cenario nao compativel permanece fail-closed.

Matriz coberta (letras do enunciado):
 A COLETA_11 intacta | B L1 Financeiro | C L1 PC | D L1 Itens Consumidos
 E L2 Financeiro | F L2 + aditivo | G L2 + PC | H L2 Itens Consumidos
 I estrutura nao homologada | J bloco nao reproduzivel | K divergencia nao
 explicada.

Cada decisao nasce no motor (`_formalizacao_compatibilidade`); a UI apenas a
consome e o fluxo dos documentos/expanders e o mesmo de uma Coleta 11.0.
"""
from __future__ import annotations

import hashlib
import io
import json
import warnings
from pathlib import Path

import pytest
from openpyxl import load_workbook

import _compatibilidade_coleta as cc
import _compatibilidade_valores as cv
import _formalizacao_compatibilidade as fc
import _leitor_masterfile_v10 as leitor
from _compatibilidade_valores import ATRIBUTO_AUDITORIA, aplicar_compatibilidade_valores, valores_legados
from _politica_entrega_segura import (
    MENSAGEM_COLETA_PRECISAO_ANTERIOR,
    avaliar_entrega_segura,
    mensagens_bloqueio_documental_duro,
)
from _versao import CL8US_VERSION, COLETA_VERSION

warnings.filterwarnings("ignore")

ROOT = Path(__file__).resolve().parents[1]
PASTA = Path(__file__).resolve().parent / "fixtures" / "coletas_compat"
GOLDEN_26 = Path(r"C:\Users\danie\Downloads\anthropic-skills") / (
    "Coleta_Reajuste_C1_C2_C3_ICTI_26-08-2026.xlsx"
)
TRES_DOCUMENTOS = ("sumario_executivo", "despacho_saneador", "termo_apostila")

_CACHE: dict = {}


def _bytes(nome: str) -> bytes:
    return (PASTA / f"{nome}.xlsx").read_bytes()


def _runtime(nome: str, conteudo: bytes | None = None):
    if nome not in _CACHE:
        from _coleta_reajuste_documentos import processar_coleta_oficial_runtime

        _CACHE[nome] = processar_coleta_oficial_runtime(conteudo or _bytes(nome))
    return _CACHE[nome]


def _documentos_habilitados(resultado) -> dict[str, bool]:
    documentos = resultado["capacidades"]["documentos"]
    return {chave: bool(documentos[chave]["habilitado"]) for chave in TRES_DOCUMENTOS}


# --------------------------------------------------------------------------- #
# A. COLETA_11 — comportamento atual intacto.
# --------------------------------------------------------------------------- #
@pytest.mark.parametrize("nome", ["coleta_11_financeiro", "coleta_11_pc",
                                  "coleta_11_consumidos_item_unico"])
def test_a_coleta_11_continua_normal_sem_decisao_de_compatibilidade(nome):
    resultado, _ = _runtime(nome)
    assert resultado["compatibilidade_formalizacao"] == {}
    assert resultado["compatibilidade_auditoria"]["compatibilidade_aplicada"] is False
    assert resultado["bloqueios_formalizacao"] == []
    assert resultado["bloqueios_documentais_duros"] == []
    assert resultado["formalizacao_bloqueada"] is False
    assert all(_documentos_habilitados(resultado).values())
    assert resultado["reconciliacao_xls_python"]["status_geral"] == "CONCILIADO"


# --------------------------------------------------------------------------- #
# B, C, D, H. Cenarios que DEVEM poder formalizar.
# --------------------------------------------------------------------------- #
LIBERADOS = [
    ("B_L1_financeiro", "pre11_l1_financeiro", cc.LINHAGEM_PRE_11_L1),
    ("C_L1_pc", "pre11_l1_pc", cc.LINHAGEM_PRE_11_L1),
    ("D_L1_consumidos", "pre11_l1_consumidos_item_unico", cc.LINHAGEM_PRE_11_L1),
    ("H_L2_consumidos", "pre11_l2_consumidos", cc.LINHAGEM_PRE_11_L2),
]


@pytest.mark.parametrize("rotulo,nome,linhagem", LIBERADOS, ids=[r[0] for r in LIBERADOS])
def test_cenarios_compativeis_sao_formalizaveis(rotulo, nome, linhagem):
    resultado, _ = _runtime(nome)
    decisao = resultado["compatibilidade_formalizacao"]
    assert decisao["linhagem"] == linhagem
    assert decisao["status"] == fc.STATUS_FORMALIZAVEL
    assert decisao["elegivel"] is True and decisao["formalizacao_liberada"] is True
    assert decisao["blocos_nao_reproduziveis"] == {}
    assert decisao["restricoes"] == []
    assert decisao["divergencias_nao_explicadas"] == []
    assert decisao["divergencias_esperadas"], "o caso precisa ter divergencia explicada"
    assert decisao["ciclos_precisao_bruta"] == ["C1"]
    # Nada bloqueia, nem a mensagem de precisao anterior nem divergencia.
    assert resultado["bloqueios_formalizacao"] == []
    assert resultado["bloqueios_documentais_duros"] == []
    assert resultado["formalizacao_bloqueada"] is False
    # A auditoria XLS x Python permanece, reclassificada pela CAUSA.
    reconciliacao = resultado["reconciliacao_xls_python"]
    assert reconciliacao["status_geral"] == fc.STATUS_DIVERGENCIA_COMPATIBILIZADA
    assert reconciliacao["divergencias_relevantes"] == []
    assert reconciliacao["divergencias_compatibilizadas"]
    # O aviso historico nao some: a Coleta continua registrada como anterior.
    assert resultado["compatibilidade_auditoria"]["compatibilidade_aplicada"] is True
    # Mensagem discreta e nao alarmista, via "informacoes" do motor.
    politica = resultado["politica_entrega_segura"]
    assert fc.MENSAGEM_COMPATIBILIZADA in politica["informacoes"]
    assert politica["status"] != "BLOQUEADO_PARA_FORMALIZACAO"


@pytest.mark.parametrize("rotulo,nome,linhagem", LIBERADOS, ids=[r[0] for r in LIBERADOS])
def test_cenarios_compativeis_habilitam_os_tres_documentos_e_geram_os_arquivos(
    rotulo, nome, linhagem
):
    from docx import Document

    from _sumario_executivo import gerar_sumario_executivo
    from _templates_documentos import gerar_despacho_saneador, gerar_termo_apostila

    resultado, _ = _runtime(nome)
    assert all(_documentos_habilitados(resultado).values())
    pdf = gerar_sumario_executivo(resultado)
    saneador = gerar_despacho_saneador(resultado)
    termo = gerar_termo_apostila(resultado)          # => termo_disponivel
    assert pdf[:5] == b"%PDF-" and len(pdf) > 2000
    assert saneador[:2] == b"PK" and termo[:2] == b"PK"      # DOCX (zip) validos
    for arquivo in (saneador, termo):
        assert Document(io.BytesIO(arquivo)).paragraphs        # abre e tem conteudo


# --------------------------------------------------------------------------- #
# Caso real obrigatorio (golden externo; pulado quando ausente).
# --------------------------------------------------------------------------- #
@pytest.mark.skipif(not GOLDEN_26.exists(), reason="golden externo ausente")
def test_caso_real_obrigatorio_formaliza_com_o_valor_canonico_atual():
    resultado, _ = _runtime("golden_26", GOLDEN_26.read_bytes())
    decisao = resultado["compatibilidade_formalizacao"]
    assert decisao["linhagem"] == cc.LINHAGEM_PRE_11_L1
    assert decisao["status"] == fc.STATUS_FORMALIZAVEL and decisao["formalizacao_liberada"]
    assert decisao["ciclos_precisao_bruta"] == ["C3"]
    esperadas = {d["campo"]: d for d in decisao["divergencias_esperadas"]}
    assert set(esperadas) == {"RETRO_FIN", "RETRO_OFICIAL", "REM_ATUALIZADO_OFICIAL", "VTA_FINAL"}
    vta = esperadas["VTA_FINAL"]
    # XLS legado x valor canonico atual: R$ 1,43, so por causa da precisao de C3.
    assert vta["xls"] == 8_713_820.26 and vta["python_vigente"] == 8_713_821.69
    assert vta["python_legado_reproduzido"] == 8_713_820.26 and vta["diferenca"] == 1.43
    assert vta["causa"] == fc.CAUSA_PRECISAO
    assert resultado["valor_atualizado_contrato"] == 8_713_821.69
    assert resultado["bloqueios_formalizacao"] == [] and not resultado["formalizacao_bloqueada"]
    assert all(_documentos_habilitados(resultado).values())
    # C3 = 2,89% (nao 2,8899...%) no resultado canonico; o bruto fica no historico.
    c3 = resultado["objeto_processo"]["dados_operacionais"]["parametros_v10"]["por_ciclo"]["C3"]
    assert c3["percentual_reajuste"] == 0.0289 and c3["percentual_reajuste_bruto"] != 0.0289


@pytest.mark.skipif(not GOLDEN_26.exists(), reason="golden externo ausente")
def test_caso_real_documentos_nao_vazam_o_valor_legado():
    from _baseline_fotografia import _texto_docx
    from _sumario_executivo import montar_dados_sumario_executivo
    from _templates_documentos import gerar_despacho_saneador, gerar_termo_apostila

    resultado, _ = _runtime("golden_26", GOLDEN_26.read_bytes())
    termo = " ".join(_texto_docx(gerar_termo_apostila(resultado)))
    saneador = " ".join(_texto_docx(gerar_despacho_saneador(resultado)))
    sintese = montar_dados_sumario_executivo(resultado)["sintese"]
    legados = ("8.713.820,26", "24.678,92", "1.388.251,07")
    for texto in (termo, saneador, json.dumps(sintese, ensure_ascii=False, default=str)):
        assert not [v for v in legados if v in texto], "valor legado vazou"
    for texto in (termo, saneador):
        assert "8.713.821,69" in texto and "24.679,47" in texto
    # Sumario, Saneador, Termo e painel convergem no mesmo VTA canonico.
    assert sintese["vta"] == resultado["valor_atualizado_contrato"] == 8_713_821.69
    assert sintese["retroativo_total"] == 24_679.47


# --------------------------------------------------------------------------- #
# Nao vazamento do XLS legado (cenario sintetico, sempre disponivel).
# --------------------------------------------------------------------------- #
def test_documentos_usam_somente_valores_canonicos_e_o_painel_converge():
    from _baseline_fotografia import _texto_docx
    from _sumario_executivo import montar_dados_sumario_executivo
    from _templates_documentos import gerar_despacho_saneador, gerar_termo_apostila

    resultado, _ = _runtime("pre11_l1_financeiro")
    xls = {c["campo"]: c["xls"] for c in resultado["reconciliacao_xls_python"]["campos"]}
    assert xls["VTA_FINAL"] == 1_341_003.94           # o XLS legado diz isto...
    assert resultado["valor_atualizado_contrato"] == 1_340_973.60   # ...o canonico e isto
    termo = " ".join(_texto_docx(gerar_termo_apostila(resultado)))
    saneador = " ".join(_texto_docx(gerar_despacho_saneador(resultado)))
    sintese = montar_dados_sumario_executivo(resultado)["sintese"]
    for texto in (termo, saneador, json.dumps(sintese, ensure_ascii=False, default=str)):
        assert "1.341.003,94" not in texto and "1341003.94" not in texto
        assert "26.131,44" not in texto and "294.872,5" not in texto
    for texto in (termo, saneador):
        assert "1.340.973,60" in texto and "26.112,00" in texto
    # Status do XLS = REVISE neste cenario sintetico: o Sumario publica a PREVIA.
    assert (sintese["vta"] or sintese["vta_previa"]) == 1_340_973.60
    assert sintese["retroativo_total"] == 26_112.00
    # Fonte canonica unica: o painel (consolidado) publica o mesmo VTA.
    assert resultado["resultado_consolidado"]["vta"] == 1_340_973.60


def test_auditoria_do_xls_legado_permanece_disponivel():
    resultado, _ = _runtime("pre11_l1_financeiro")
    reconciliacao = resultado["reconciliacao_xls_python"]
    registro = {c["campo"]: c for c in reconciliacao["divergencias_compatibilizadas"]}
    vta = registro["VTA_FINAL"]
    assert vta["xls"] == 1_341_003.94 and vta["python"] == 1_340_973.60
    assert vta["python_legado_reproduzido"] == 1_341_003.94
    assert vta["causa"] == fc.CAUSA_PRECISAO
    assert vta["status"] == fc.STATUS_DIVERGENCIA_COMPATIBILIZADA
    # Os valores do XLS seguem expostos na visao do arquivo (nunca apagados).
    assert resultado["resultados_xls"]["valores"]["VTA_FINAL"] == 1_341_003.94
    auditoria = resultado["compatibilidade_auditoria"]["valores"]
    assert auditoria["aplicada"] is True and auditoria["celulas_alteradas"] > 50
    assert auditoria["blocos_nao_reproduziveis"] == {}


# --------------------------------------------------------------------------- #
# Cenarios que DEVEM permanecer bloqueados.
# --------------------------------------------------------------------------- #
def test_f_l2_com_aditivo_permanece_bloqueado_com_motivo_especifico():
    resultado, _ = _runtime("pre11_l2_financeiro_regerada")
    assert resultado["formalizacao_bloqueada"] is True
    duros = resultado["bloqueios_documentais_duros"]
    assert duros and "aditivo em ciclo de reajuste" in duros[0]
    # O motivo especifico substitui a orientacao generica de regerar a Coleta.
    assert MENSAGEM_COLETA_PRECISAO_ANTERIOR not in resultado["bloqueios_formalizacao"]
    assert not any(_documentos_habilitados(resultado).values())
    documento = resultado["capacidades"]["documentos"]["termo_apostila"]
    assert documento["motivo"] == duros[0]


def test_g_l2_pc_permanece_bloqueado_sem_inventar_data():
    resultado, _ = _runtime("pre11_l2_pc")
    assert resultado["formalizacao_bloqueada"] is True
    texto = " ".join(resultado["bloqueios_formalizacao"])
    assert "inicio do efeito financeiro ausente ou inconsistente" in texto
    decisao = resultado["compatibilidade_formalizacao"]
    assert decisao["formalizacao_liberada"] is False
    assert decisao["elegivel"] is False                  # sem conferencia XLS x Python
    # Nenhuma data de efeito foi inventada para o L2.
    por_ciclo = resultado["objeto_processo"]["dados_operacionais"]["parametros_v10"]["por_ciclo"]
    assert all(not por_ciclo[c].get("inicio_efeito_financeiro") for c in ("C1", "C2", "C3"))


def test_e_l2_financeiro_sem_aditivo_segue_bloqueado_pela_divergencia_nao_explicada():
    """O VTA do XLS L2 segue a formula anterior (nao inclui o executado): nao e
    precisao, o replay nao o reproduz, logo o mecanismo NAO a explica."""
    resultado, _ = _runtime("pre11_l2_financeiro_sem_extras")
    decisao = resultado["compatibilidade_formalizacao"]
    assert decisao["status"] == fc.STATUS_NAO_COMPATIBILIZADA and not decisao["elegivel"]
    assert [d["campo"] for d in decisao["divergencias_nao_explicadas"]] == ["VTA_FINAL"]
    esperadas = {d["campo"] for d in decisao["divergencias_esperadas"]}
    assert {"RETRO_FIN", "REM_ATUALIZADO_OFICIAL"} <= esperadas   # a precisao e explicada
    assert resultado["formalizacao_bloqueada"] is True
    assert not any(_documentos_habilitados(resultado).values())


def test_k_divergencia_nao_explicada_continua_bloqueada():
    """L1 Itens Consumidos com varios itens: o QTD_REM do XLS enxerga so o primeiro
    item (defeito do template, presente ate no 11.0). Nao e precisao: bloqueia."""
    resultado, _ = _runtime("pre11_l1_consumidos")
    decisao = resultado["compatibilidade_formalizacao"]
    assert decisao["status"] == fc.STATUS_NAO_COMPATIBILIZADA
    nao_explicadas = {d["campo"] for d in decisao["divergencias_nao_explicadas"]}
    assert "QTD_REM_OFICIAL" in nao_explicadas
    assert resultado["formalizacao_bloqueada"] is True
    assert resultado["reconciliacao_xls_python"]["divergencias_relevantes"]
    assert not resultado["reconciliacao_xls_python"].get("divergencias_compatibilizadas")


def test_i_estrutura_nao_homologada_continua_rejeitada():
    from _coleta_reajuste_documentos import processar_coleta_oficial_runtime

    wb = load_workbook(PASTA / "pre11_l2_financeiro.xlsx")
    wb["itens_PC"]["A1"] = "PC"
    saida = io.BytesIO()
    wb.save(saida)
    with pytest.raises(ValueError, match="não homologado"):
        processar_coleta_oficial_runtime(saida.getvalue())


def test_j_bloco_nao_reproduzivel_bloqueia_com_motivo_especifico(monkeypatch):
    from _coleta_reajuste_documentos import processar_coleta_oficial_runtime

    monkeypatch.setattr(
        cv, "_reproduz_cache",
        lambda wb, bloco: (False, [("financeiro", 3, 5, 1.0, 2.0)]),
    )
    resultado, _ = processar_coleta_oficial_runtime(_bytes("pre11_l1_financeiro"))
    auditoria = resultado["compatibilidade_auditoria"]["valores"]
    assert auditoria["blocos_nao_reproduziveis"] and not auditoria["blocos_adaptados"]
    assert auditoria["aplicada"] is False
    assert resultado["formalizacao_bloqueada"] is True


# --------------------------------------------------------------------------- #
# Politica de precisao: cirurgica, nunca generica.
# --------------------------------------------------------------------------- #
def _leitura_sintetica(**extras):
    base = {
        "ok": True,
        "controle": {"modo": "Principal"},
        "parametros_v10": {"ciclos_precisao_bruta": ["C3"], "por_ciclo": {}},
        "objeto_processo": {},
        "compatibilidade_formalizacao": {},
        "compatibilidade_restricoes": [],
    }
    base.update(extras)
    return base


def test_precisao_anterior_nao_compatibilizada_bloqueia():
    leitura = _leitura_sintetica()
    assert mensagens_bloqueio_documental_duro(leitura) == [MENSAGEM_COLETA_PRECISAO_ANTERIOR]
    assert MENSAGEM_COLETA_PRECISAO_ANTERIOR in avaliar_entrega_segura(leitura)["bloqueios"]


def test_precisao_anterior_compatibilizada_nao_bloqueia_e_informa():
    leitura = _leitura_sintetica(compatibilidade_formalizacao={
        "elegivel": True, "mensagem": fc.MENSAGEM_COMPATIBILIZADA,
    })
    assert mensagens_bloqueio_documental_duro(leitura) == []
    politica = avaliar_entrega_segura(leitura)
    assert MENSAGEM_COLETA_PRECISAO_ANTERIOR not in politica["bloqueios"]
    assert fc.MENSAGEM_COMPATIBILIZADA in politica["informacoes"]


def test_adaptacao_incompleta_bloqueia_com_motivo_especifico():
    leitura = _leitura_sintetica(compatibilidade_formalizacao={
        "elegivel": False, "blocos_nao_reproduziveis": {"itens_PC": [{"celula": "H2"}]},
    })
    duros = mensagens_bloqueio_documental_duro(leitura)
    assert len(duros) == 1 and "itens_PC" in duros[0]
    assert MENSAGEM_COLETA_PRECISAO_ANTERIOR not in duros      # motivo real, nao generico
    assert duros[0] in avaliar_entrega_segura(leitura)["bloqueios"]


def test_restricao_de_modelo_bloqueia_com_a_propria_mensagem():
    leitura = _leitura_sintetica(compatibilidade_restricoes=[
        {"codigo": "X", "mensagem": "Restricao especifica do modelo."},
    ])
    assert mensagens_bloqueio_documental_duro(leitura) == ["Restricao especifica do modelo."]


def test_metodo_pc_mantem_a_regra_existente_sem_mensagem_de_precisao():
    leitura = _leitura_sintetica(controle={"modo": "PC"})
    assert mensagens_bloqueio_documental_duro(leitura) == []


# --------------------------------------------------------------------------- #
# A decisao: causa comprovada, nunca apenas tolerancia.
# --------------------------------------------------------------------------- #
def _recon(campos, relevantes=None):
    return {"disponivel": True, "sem_cache": False, "campos": campos,
            "divergencias_relevantes": relevantes or [], "status_geral": None}


def _campo(nome, xls, py, status="DIVERGENCIA_RELEVANTE"):
    return {"campo": nome, "rotulo": nome, "xls": xls, "python": py, "status": status}


def _leitura_decisao(atual, **extras):
    base = {
        "coleta_linhagem": {"codigo": cc.LINHAGEM_PRE_11_L1},
        "compatibilidade_aplicada": True,
        "compatibilidade_valores": {
            "aplicada": True, "ciclos_precisao_bruta": ["C1"],
            "blocos_adaptados": ["financeiro"], "blocos_nao_reproduziveis": {},
        },
        "compatibilidade_restricoes": [],
        "reconciliacao_xls_python": atual,
        "composicao_vta": {},
    }
    base.update(extras)
    return base


def test_divergencia_explicada_pela_causa_independe_do_tamanho():
    atual_rel = _campo("VTA_FINAL", 100.00, 1_000_000.00)           # diferenca ENORME
    legado = {"reconciliacao_xls_python": _recon([_campo("VTA_FINAL", 100.00, 100.00, "CONCILIADO")]),
              "ok": True}
    leitura = _leitura_decisao(_recon([atual_rel], [atual_rel]))
    decisao = fc.decidir_formalizacao_compatibilidade(leitura, legado)
    assert decisao["elegivel"] is True                  # a CAUSA decide, nao o tamanho
    assert decisao["divergencias_esperadas"][0]["causa"] == fc.CAUSA_PRECISAO


def test_um_centavo_sem_reproducao_do_legado_continua_relevante():
    atual_rel = _campo("VTA_FINAL", 100.00, 100.01)
    legado = {"reconciliacao_xls_python": _recon([_campo("VTA_FINAL", 100.00, 100.01)]),
              "ok": True}                              # o replay NAO muda o resultado
    decisao = fc.decidir_formalizacao_compatibilidade(
        _leitura_decisao(_recon([atual_rel], [atual_rel])), legado)
    assert decisao["elegivel"] is False
    assert [d["campo"] for d in decisao["divergencias_nao_explicadas"]] == ["VTA_FINAL"]


@pytest.mark.parametrize("py_legado,esperado", [(100.00, True), (100.01, True), (100.02, False)])
def test_reproducao_do_calculo_antigo_admite_so_arredondamento_de_1_centavo(py_legado, esperado):
    atual_rel = _campo("VTA_FINAL", 100.00, 150.00)
    legado = {"reconciliacao_xls_python": _recon([_campo("VTA_FINAL", 100.00, py_legado)]),
              "ok": True}
    decisao = fc.decidir_formalizacao_compatibilidade(
        _leitura_decisao(_recon([atual_rel], [atual_rel])), legado)
    assert decisao["elegivel"] is esperado


def test_sem_causa_de_precisao_nao_ha_compatibilizacao():
    atual_rel = _campo("VTA_FINAL", 100.00, 150.00)
    legado = {"reconciliacao_xls_python": _recon([_campo("VTA_FINAL", 100.00, 100.00)]), "ok": True}
    leitura = _leitura_decisao(_recon([atual_rel], [atual_rel]))
    leitura["compatibilidade_valores"]["ciclos_precisao_bruta"] = []
    assert fc.decidir_formalizacao_compatibilidade(leitura, legado)["elegivel"] is False


def test_replay_indisponivel_e_fail_closed():
    atual_rel = _campo("VTA_FINAL", 100.00, 150.00)
    decisao = fc.decidir_formalizacao_compatibilidade(
        _leitura_decisao(_recon([atual_rel], [atual_rel])), None)
    assert decisao["elegivel"] is False and "reproduzir" in decisao["motivo"]


def test_reclassificacao_so_ocorre_quando_a_decisao_e_elegivel():
    campo = _campo("VTA_FINAL", 100.00, 150.00)
    reconciliacao = _recon([campo], [campo])
    fc.aplicar_reclassificacao(reconciliacao, {"elegivel": False,
                                               "divergencias_esperadas": [{"campo": "VTA_FINAL"}]})
    assert reconciliacao["divergencias_relevantes"] and campo["status"] == "DIVERGENCIA_RELEVANTE"


# --------------------------------------------------------------------------- #
# Replay legado: isolado, reversivel e sem tocar o arquivo/entradas.
# --------------------------------------------------------------------------- #
def test_valores_legados_restaura_o_cache_original_e_depois_o_recomposto():
    wb = load_workbook(PASTA / "pre11_l2_financeiro.xlsx", data_only=True)
    auditoria = aplicar_compatibilidade_valores(wb)
    assert auditoria["aplicada"]
    # primeira celula realmente alterada (a auditoria guarda o valor do XLS)
    alterada = auditoria["amostra"][0]
    aba, linha, coluna = alterada["aba"], alterada["linha"], alterada["coluna"]
    original, recomposto = alterada["xls"], alterada["recomposto"]
    assert recomposto != original
    assert wb[aba].cell(linha, coluna).value == recomposto
    with valores_legados(wb):
        assert wb[aba].cell(linha, coluna).value == original
    assert wb[aba].cell(linha, coluna).value == recomposto
    with pytest.raises(RuntimeError):
        with valores_legados(wb):
            raise RuntimeError("falha dentro do bloco")
    assert wb[aba].cell(linha, coluna).value == recomposto     # restaurado mesmo em erro


def test_replay_nao_vaza_estado_nem_altera_arquivo_ou_entradas():
    from _contexto_coleta import ContextoColeta

    conteudo = _bytes("pre11_l2_consumidos")
    antes = hashlib.sha256(conteudo).hexdigest()
    original = load_workbook(io.BytesIO(conteudo), data_only=True)
    with ContextoColeta(conteudo) as contexto:
        leitura = leitor.ler_masterfile_v10(conteudo, exigir_modelo_oficial=True, contexto=contexto)
        wb = contexto.workbook_valores
        assert leitura["compatibilidade_formalizacao"]["elegivel"] is True
        # entradas do fiscal intactas depois do replay
        for aba, celula in (("parametros", "E3"), ("itens_Remanesc", "C2"), ("itens_Consumidos", "G2")):
            assert wb[aba][celula].value == original[aba][celula].value
        assert getattr(wb, ATRIBUTO_AUDITORIA)["aplicada"] is True
        assert wb["parametros"]["F3"].value == 1.0512          # recomposto depois do replay
    assert hashlib.sha256(conteudo).hexdigest() == antes
    assert leitor._REPLAY_LEGADO.get() is False


# --------------------------------------------------------------------------- #
# Versionamento: Cl8us 11.4 gera Coleta no Modelo 11.1; a Coleta 11.0 (fixture)
# continua aceita como COLETA_11, sem adaptacao.
# --------------------------------------------------------------------------- #
def test_cl8us_11_4_e_modelo_de_coleta_11_1_sao_independentes():
    # Cl8us 11.5 / Coleta 11.2 (RESULTADOS-EXECUTIVO-V3); a regra testada —
    # versao do gerador independente da linhagem — e a mesma.
    # Cl8us 11.7 / Coleta 11.3 (versionamento obrigatorio pos-#175).
    assert CL8US_VERSION == "12.4" and COLETA_VERSION == "11.7"
    wb = load_workbook(PASTA / "coleta_11_financeiro.xlsx")
    assert cc.detectar_linhagem_coleta(wb)["codigo"] == cc.LINHAGEM_COLETA_11
    wb["CONTROLE"]["B25"] = CL8US_VERSION            # quem GEROU a Coleta: Cl8us 11.4
    # A fixture e uma Coleta 11.0 (gerada antes do quadro "sem efeito"): o
    # marcador dela NAO e reescrito e ela segue sendo COLETA_11.
    assert wb["CONTROLE"]["B24"].value == "11.0" != COLETA_VERSION
    # O modelo continua COLETA_11: a versao do gerador nao define a linhagem.
    deteccao = cc.detectar_linhagem_coleta(wb)
    assert deteccao["codigo"] == cc.LINHAGEM_COLETA_11
    assert deteccao["marcador_publico"] == "11.0"
    assert deteccao["suportada"] is True
    assert deteccao["compatibilidade_aplicada"] is False


def test_coleta_11_1_e_coleta_11_0_sao_a_mesma_linhagem_sem_adaptacao():
    for marcador in ("11.0", "11.1"):
        wb = load_workbook(PASTA / "coleta_11_financeiro.xlsx")
        wb["CONTROLE"]["B24"] = marcador
        deteccao = cc.detectar_linhagem_coleta(wb)
        assert deteccao["codigo"] == cc.LINHAGEM_COLETA_11, marcador
        assert deteccao["suportada"] is True and not deteccao["compatibilidade_aplicada"]
        assert deteccao["modelo_canonico"] == COLETA_VERSION
    # Marcador fora da familia 11.x continua rejeitado (nada e inferido).
    wb = load_workbook(PASTA / "coleta_11_financeiro.xlsx")
    wb["CONTROLE"]["B24"] = "11.9"
    assert cc.detectar_linhagem_coleta(wb)["codigo"] == cc.LINHAGEM_NAO_HOMOLOGADA


# --------------------------------------------------------------------------- #
# Fluxo UNICO da pagina: cards e expanders pelo mesmo caminho do 11.0.
# --------------------------------------------------------------------------- #
PAGINA = (ROOT / "pages" / "03_Valor_Global.py").read_text(encoding="utf-8")


def _trecho(inicio: str, fim: str) -> str:
    a = PAGINA.index(inicio)
    return PAGINA[a: PAGINA.index(fim, a + 1)]


def test_os_expanders_dependem_so_do_termo_gerado_sem_excecao_para_legado():
    fluxo = _trecho("    termo_disponivel = render_documentos_funcionais_upload(resultado)",
                    "documentos_cap = ")
    assert "if termo_disponivel:" in fluxo
    assert "_render_comunicados_pos_documentos()" in fluxo
    for proibido in ("compatibilidade", "PRE_11", "legado", "linhagem"):
        assert proibido not in fluxo
    corpo = _trecho("def _render_comunicados_pos_documentos", "def render_documentos_funcionais_upload")
    assert corpo.count("st.expander(") == 3
    for rotulo in ("E-mail ao fiscal", "E-mail solicitando RC ao fiscal", "E-mail à contratada"):
        assert rotulo in corpo


def test_a_pagina_consome_a_lista_de_bloqueios_documentais_do_motor():
    render = _trecho("def render_documentos_funcionais_upload", "def _invalidar_caso_antes")
    assert 'resultado.get("bloqueios_documentais_duros")' in render
    assert "st.warning(mensagem_bloqueio)" in render
    # a UI nao decide compatibilidade: nenhum "if PRE_11_*" na pagina
    assert "PRE_11_L1" not in PAGINA and "PRE_11_L2" not in PAGINA


def test_a_decisao_vive_no_motor_nao_na_camada_streamlit():
    fonte_decisao = (ROOT / "_formalizacao_compatibilidade.py").read_text(encoding="utf-8")
    assert "streamlit" not in fonte_decisao
    assert "import streamlit" not in (ROOT / "_politica_entrega_segura.py").read_text(encoding="utf-8")


def test_card_do_topo_nao_exibe_o_retroativo_do_xls_antes_do_canonico():
    """Regressao real vista na tela: o card "Retroativo reconhecido" mostrava
    R$ 24.678,92 (XLS legado) ao lado de R$ 24.679,47 (canonico)."""
    a = PAGINA.index("_retro_val = (_tc_pc.get(")
    trecho = PAGINA[a: PAGINA.index("_retro_str = (", a)]
    posicao_canonica = trecho.index('"retroativo_reconhecido"')
    posicao_xls = trecho.index('"retroativo_oficial"')
    assert posicao_canonica < posicao_xls, "o XLS so pode ser a ULTIMA queda"
