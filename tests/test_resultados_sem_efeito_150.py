# -*- coding: utf-8 -*-
"""RESULTADOS-SEM-EFEITO-150 — "Execucao sem efeito financeiro" na RESULTADOS.

O QUE ESTA SUITE PROVA
----------------------
Camada ESTRUTURAL (sempre roda; openpyxl sobre o template):
  * o template carrega exatamente as formulas do aplicador, em ASCII, em
    ingles e com parenteses balanceados (regra ZERO CORRUPCAO XLSX);
  * os textos acentuados vivem em celulas constantes de MEMORIA_RESULTADOS;
  * o bloco antigo do PC (T28:T30 + A23:E23) segue intacto — nao ha segundo
    calculo concorrente;
  * o quadro e so explicativo: nenhuma formula dele toca VTA/retroativo e
    nenhuma outra formula do arquivo o consome.

Camada de VALORES (Excel real, opt-in `RUN_EXCEL_INTEGRATION=1`; o CI nao tem
Excel). Cenarios gerados sobre o template oficial e recalculados no Excel:
  * PC: PC anterior ao inicio do efeito ENTRA; posterior NAO; C0 e ciclo nao
    computado (ambos com EFEITO_FINANCEIRO_PC = Nao) NAO entram; sem PC no
    periodo => R$ 0,00 conhecido; sem atraso => "sem execucao anterior";
    a soma por ciclo reconcilia com o helper antigo T28:T30;
  * Financeiro: duas competencias antes do efeito; competencia com G = Sim
    fica fora; sem pagamento no periodo; multiciclo com fator encadeado;
  * Consumidos: sem data de consumo => "nao mensuravel", sem estimativa;
    consumo zero => 0,00 conhecido; sem consumo informado => quadro oculto;
  * SEGURANCA: comparacao A/B, celula a celula, do workbook recalculado com o
    template BASE (PR #171) e com o template novo — VTA, retroativo
    reconhecido/potencial, remanescente e tudo mais ficam identicos fora da
    allowlist (RESULTADOS_DETALHE!E15:H21 + D17:D20 de borda e MEMORIA S78:W120).
"""
from __future__ import annotations

from _resultados_abas import aba_resultados_tecnica as _aba_tecnica_resultados
import gc
import io
import os
import re
import subprocess
import time
from datetime import date
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

import _baseline_cenarios as cen
from tools import aplicar_resultados_sem_efeito_150 as ap

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"
# Commit da main imediatamente ANTES desta frente (merge do PR #171).
COMMIT_BASE = "b947ccb8eed2f7a8daf99522f9fdeb736debcfad"

com = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para Excel COM",
)

# Grandezas economicas que o quadro NUNCA pode referenciar.
NOMES_ECONOMICOS = (
    "RETRO_OFICIAL", "RETRO_PC", "RETRO_FIN", "RETRO_ITENS", "VTA_FINAL",
    "VTA_CALCULADO", "VTA_MANUAL_OFICIAL", "RETROATIVO_POTENCIAL_PC",
    "RETROATIVO_POTENCIAL_VTA", "RETROATIVO_POTENCIAL_APURADO",
    "EXECUTADO_APURADO", "REM_BASE_OFICIAL", "REM_ATUALIZADO_OFICIAL",
    "SALDO_REMANESCENTE_ATUAL", "AJUSTES_DEVIDOS",
)


# =========================================================================== #
# 1. ESTRUTURA (sempre roda)
# =========================================================================== #
@pytest.fixture(scope="module")
def wb_formulas():
    wb = load_workbook(TEMPLATE, data_only=False)
    yield wb
    wb.close()


def _normalizar(formula: str) -> str:
    return re.sub(r"\s+", "", str(formula or "").replace("_xlfn.", ""))


def _equilibrada(formula: str) -> bool:
    nivel, em_texto = 0, False
    for ch in formula:
        if ch == '"':
            em_texto = not em_texto
        elif not em_texto:
            nivel += (ch == "(") - (ch == ")")
            if nivel < 0:
                return False
    return nivel == 0 and not em_texto


def test_template_carrega_exatamente_as_formulas_do_aplicador(wb_formulas):
    memoria, res = wb_formulas[ap.ABA_MEMORIA], wb_formulas[ap.ABA_RESULTADOS]
    divergentes = []
    for endereco, esperada in ap.formulas_helpers().items():
        if _normalizar(memoria[endereco].value) != _normalizar(esperada):
            divergentes.append(f"{ap.ABA_MEMORIA}!{endereco}")
    for endereco, esperada in ap.formulas_resultados().items():
        if _normalizar(res[endereco].value) != _normalizar(esperada):
            divergentes.append(f"{ap.ABA_RESULTADOS}!{endereco}")
    assert not divergentes, divergentes
    assert len(ap.formulas_helpers()) == 76 and len(ap.formulas_resultados()) == 22


def test_formulas_novas_sao_ascii_em_ingles_e_balanceadas():
    todas = {**ap.formulas_helpers(), **ap.formulas_resultados()}
    for endereco, formula in todas.items():
        assert formula.startswith("="), endereco
        assert formula.isascii(), f"{endereco}: caractere nao-ASCII na formula"
        assert _equilibrada(formula), f"{endereco}: parenteses desbalanceados"
        # TEXT() depende do idioma do Excel ("dd/mm/yyyy" vs "dd/mm/aaaa"): datas
        # sao montadas por DAY/MONTH/YEAR.
        assert "TEXT(" not in formula, endereco
        # Nada de funcao em portugues.
        assert not re.search(r"\b(SE|SEERRO|PROCV|OU|CONT\.SE|SOMASE)\(", formula)


def test_textos_acentuados_ficam_em_celulas_constantes(wb_formulas):
    memoria = wb_formulas[ap.ABA_MEMORIA]
    textos = {chave: memoria[f"T{linha}"].value for chave, linha in ap.LINHA_TEXTO.items()}
    assert textos["T_TITULO"] == "EXECUÇÃO SEM EFEITO FINANCEIRO"
    assert textos["T_HDR_TERIAM"] == "Valor que teriam com o reajuste"
    assert "Valor atualizado" not in " ".join(textos.values())
    assert textos["T_NOTA"] == (
        "A diferença é apenas informativa. Não constitui retroativo reconhecido "
        "ou potencial e não representa valor a pagar."
    )
    assert textos["T_CASO_C"] == (
        "Execução sem efeito financeiro identificada, mas o impacto monetário "
        "não pode ser mensurado com segurança com os dados disponíveis."
    )
    assert textos["T_CASO_D"].startswith("Não foi identificada execução anterior")
    for valor in textos.values():
        assert isinstance(valor, str) and valor.strip()
    # Constantes, nunca formulas.
    for linha in ap.LINHA_TEXTO.values():
        assert not str(memoria[f"T{linha}"].value).startswith("=")


def test_geometria_do_quadro_sem_deslocar_linhas(wb_formulas):
    res = wb_formulas[ap.ABA_RESULTADOS]
    mesclas = {str(m) for m in res.merged_cells.ranges}
    assert {"E16:H16", "E21:H21", "E22:F22", "G22:H22", "E23:H23"} <= mesclas
    assert "E16:H21" not in mesclas
    # Nenhuma linha inserida: os pinos tecnicos seguem nos mesmos enderecos.
    assert res["A22"].value == "TOTAL" and res["D22"].value == '=IFERROR(RETRO_OFICIAL,"")'
    assert res["A24"].value.startswith('=IF($B$5="PCs","3. REMANESCENTE')
    assert res["A25"].value == "Ciclo"
    # Linhas do quadro alinhadas a C1..C4 da tabela 2.
    assert [res[f"A{linha}"].value for linha in (17, 18, 19, 20)] == ["C1", "C2", "C3", "C4"]
    # Destaque ambar (nunca vermelho) na diferenca.
    regras = [cf.rules for cf in res.conditional_formatting if str(cf.sqref) == "H17:H20"]
    assert len(regras) == 1 and len(regras[0]) == 1
    regra = regras[0][0]
    assert regra.formula == ['$H17<>""']
    assert str(regra.dxf.fill.bgColor.rgb).upper().endswith("FFF2CC")


def test_bloco_antigo_do_pc_segue_intacto_sem_segundo_calculo(wb_formulas):
    memoria, res = wb_formulas[ap.ABA_MEMORIA], wb_formulas[ap.ABA_RESULTADOS]
    for endereco in ("T28", "T29", "T30"):
        formula = memoria[endereco].value
        assert 'itens_PC!$L$2:$L$5001,"Nao"' in formula, endereco
        assert 'itens_PC!$D$2:$D$5001,">0"' in formula, endereco
        for ciclo in ("C1", "C2", "C3", "C4"):
            assert f'itens_PC!$C$2:$C$5001,"{ciclo}"' in formula, (endereco, ciclo)
        # C0 nunca entrou: o universo ja era C1..C4 com parametros!A = "Sim".
        assert '"C0"' not in formula
        assert 'parametros!$A$3="Sim"' in formula
    assert res["A23"].value.startswith('=IF(OR(MEMORIA_RESULTADOS!$B$4<>"PCs"')
    assert "MEMORIA_RESULTADOS!$T$28" in res["A23"].value
    assert res["A23"].number_format == ";;;"          # linha de apoio oculta


def test_quadro_e_so_explicativo_nao_toca_grandeza_economica():
    todas = {**ap.formulas_helpers(), **ap.formulas_resultados()}
    for endereco, formula in todas.items():
        for nome in NOMES_ECONOMICOS:
            assert nome not in formula, f"{endereco} referencia {nome}"
        # Nem as celulas-fonte das grandezas (retroativo/VTA/remanescente).
        assert not re.search(
            r"MEMORIA_RESULTADOS!\$?[BCDE]\$?(1[56]|2[3-6]|3[3-5])\b", formula
        ), endereco
        assert not re.search(r"RESULTADOS_DETALHE!\$?[BCD]\$?(2[0-2]|3[6-8]|8[3-7])\b", formula), endereco


def test_nenhuma_outra_formula_consome_o_quadro(wb_formulas):
    area_memoria = re.compile(r"(?<![A-Za-z$!])\$?[S-W]\$?(7[89]|8\d|9\d|10\d|11\d|120)\b")
    ref_memoria_ext = re.compile(
        r"MEMORIA_RESULTADOS!\$?[S-W]\$?(7[89]|8\d|9\d|10\d|11\d|120)\b"
    )
    ref_quadro = re.compile(r"[^_A-Z]RESULTADOS_DETALHE!\$?[E-H]\$?(1[5-9]|2[01])\b")
    proprios = set(ap.formulas_helpers()) | set(ap.formulas_resultados())
    consumidores = []
    for ws in wb_formulas.worksheets:
        for linha in ws.iter_rows():
            for celula in linha:
                valor = celula.value
                if not (isinstance(valor, str) and valor.startswith("=")):
                    continue
                chave = celula.coordinate
                if ws.title in (ap.ABA_MEMORIA, ap.ABA_RESULTADOS) and chave in proprios:
                    continue
                # Coleta 11.2: a RESULTADOS executiva EXIBE o quadro (espelho
                # puro, sem soma — contrato em test_resultados_executivo_v3).
                if ws.title == "RESULTADOS":
                    continue
                if ref_memoria_ext.search(valor) or ref_quadro.search(" " + valor):
                    consumidores.append(f"{ws.title}!{chave}")
                elif ws.title == ap.ABA_MEMORIA and area_memoria.search(valor):
                    consumidores.append(f"{ws.title}!{chave}")
    assert not consumidores, consumidores[:10]


# =========================================================================== #
# 1b. FLUXO REAL: Coleta gerada pelo app e compatibilidade com a Coleta 11.0
# =========================================================================== #
def _dados_calculadora() -> dict:
    return {
        "origem": "Reajuste Simples", "indice": "IST",
        "data_base_original": "01/01/2023",
        "ciclos": [{
            "ciclo": "C1", "data_base": "01/01/2023", "data_pedido": "01/01/2024",
            "situacao": "TEMPESTIVO", "percentual_aplicado": 0.10,
            "financeiro_inicio": "01/01/2024",
        }],
    }


def test_coleta_gerada_pelo_app_carrega_o_quadro_e_o_marcador_11_1():
    from _coleta_oficial import gerar_coleta_oficial_preenchida, obter_coleta_oficial_bytes
    from _versao import CL8US_VERSION, COLETA_VERSION

    # Coleta 11.2 / Cl8us 11.5 (RESULTADOS-EXECUTIVO-V3): o quadro do PR #172
    # segue identico, agora na camada tecnica RESULTADOS_DETALHE.
    # Cl8us 11.7 / Coleta 11.3 (versionamento obrigatorio pos-#175).
    assert (CL8US_VERSION, COLETA_VERSION) == ("12.0", "11.4")
    esperadas = ap.formulas_resultados()
    for conteudo in (obter_coleta_oficial_bytes(),
                     gerar_coleta_oficial_preenchida(_dados_calculadora())):
        wb = load_workbook(io.BytesIO(conteudo), data_only=False)
        assert wb["CONTROLE"]["B24"].value == "11.4"
        assert wb["CONTROLE"]["B25"].value == "12.0"
        res = wb[ap.ABA_RESULTADOS]
        for endereco, formula in esperadas.items():
            assert _normalizar(res[endereco].value) == _normalizar(formula), endereco
        mesclas = {str(m) for m in res.merged_cells.ranges}
        assert {"E16:H16", "E21:H21"} <= mesclas and "E16:H21" not in mesclas
        memoria = wb[ap.ABA_MEMORIA]
        assert memoria[f"T{ap.L_ESTADO}"].value == ap.formulas_helpers()[f"T{ap.L_ESTADO}"]


@pytest.mark.parametrize(
    "fixture", ["coleta_11_pc", "coleta_11_financeiro", "coleta_11_consumidos_item_unico"]
)
def test_coleta_11_0_existente_continua_aceita_sem_adaptacao(fixture):
    import _compatibilidade_coleta as cc
    from _leitor_masterfile_v10 import ler_masterfile_v10

    caminho = RAIZ / "tests" / "fixtures" / "coletas_compat" / f"{fixture}.xlsx"
    wb = load_workbook(caminho, data_only=False)
    assert wb["CONTROLE"]["B24"].value == "11.0"             # marcador antigo
    assert f"T{ap.L_ESTADO}" not in {
        c.coordinate for c in wb[ap.ABA_MEMORIA][ap.L_ESTADO] if c.value not in (None, "")
    }                                                         # sem o quadro novo
    leitura = ler_masterfile_v10(caminho.read_bytes(), exigir_modelo_oficial=True)
    assert leitura["ok"] is True, leitura.get("erro")
    assert leitura["coleta_linhagem"]["codigo"] == cc.LINHAGEM_COLETA_11
    assert leitura["coleta_linhagem"]["suportada"] is True
    assert leitura["compatibilidade_aplicada"] is False
    assert leitura["derivados_recalculados"] == []


# =========================================================================== #
# 2. VALORES (Excel real, opt-in)
# =========================================================================== #
def _bytes_template_base() -> bytes | None:
    try:
        return subprocess.check_output(
            ["git", "-C", str(RAIZ), "show", f"{COMMIT_BASE}:templates/COLETA_REAJUSTE_OFICIAL.xlsx"],
            stderr=subprocess.DEVNULL, timeout=60,
        )
    except Exception:
        return None


def _montar(template, ajustar) -> bytes:
    wb = load_workbook(template, data_only=False)
    ajustar(wb)
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


_CACHE_VALORES: dict = {}


def _recalcular(conteudo: bytes, chave: str, tmp_path_factory):
    """Recalcula no Excel real e devolve o workbook com os valores em cache."""
    if chave in _CACHE_VALORES:
        return _CACHE_VALORES[chave]
    import pythoncom
    import pywintypes
    import win32com.client as win32

    pasta = tmp_path_factory.mktemp("sem_efeito_150")
    origem, destino = pasta / f"{chave}.xlsx", pasta / f"{chave}.recalc.xlsx"
    origem.write_bytes(conteudo)
    for tentativa in range(1, 5):
        pythoncom.CoInitialize()
        excel = wb = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb = excel.Workbooks.Open(str(origem), UpdateLinks=0)
            excel.CalculateFull()
            wb.SaveAs(str(destino), FileFormat=51)
            wb.Close(SaveChanges=False)
            break
        except pywintypes.com_error as erro:
            if erro.args[0] != -2147418111 or tentativa == 4:
                raise
            time.sleep(8.0)
        finally:
            if excel is not None:
                for _ in range(10):
                    try:
                        excel.Quit()
                        break
                    except Exception:
                        time.sleep(1.0)
            # Soltar os proxies COM ANTES de desinicializar: CoUninitialize com
            # referencias vivas a um servidor ja encerrado derruba o Python
            # (0x80010108, objeto desconectado).
            wb = excel = None
            gc.collect()
            pythoncom.CoUninitialize()
    resultado = load_workbook(destino, data_only=True)
    _CACHE_VALORES[chave] = resultado
    return resultado


# --- construtores de cenario (entradas) ------------------------------------ #
def _ciclos(inicios: dict[int, date], ate: int = 1):
    ciclos = cen._ate(ate)
    for indice, inicio in inicios.items():
        ciclos[indice]["inicio_efeito"] = inicio
    return ciclos


def _cen_pc(pedidos, inicio_efeito=date(2024, 7, 1)):
    def ajustar(wb):
        cen._base(wb, ciclos=_ciclos({1: inicio_efeito}),
                  metodo="PC (Pedidos de Compra)", ciclo_vigente="C1",
                  data_corte=date(2024, 12, 31))
        cen._itens_pc(wb, pedidos)
        cen._ciclo_em_execucao(wb, data=date(2024, 12, 31), linhas=cen.EXECUCAO_FISICA_PADRAO)
        cen._cobertura(wb, pcs_ate=date(2024, 12, 31))
    return ajustar


PEDIDOS_MISTOS = [
    {"numero": "PC-001", "data": date(2024, 2, 15), "valor": 180_000.00, "pago": "Sim"},
    {"numero": "PC-002", "data": date(2024, 5, 20), "valor": 96_500.00, "pago": "Sim"},
    {"numero": "PC-003", "data": date(2024, 9, 30), "valor": 145_250.00, "pago": "Nao"},  # posterior
    {"numero": "PC-C0", "data": date(2023, 6, 10), "valor": 50_000.00, "pago": "Sim"},      # C0
    {"numero": "PC-C2", "data": date(2025, 3, 10), "valor": 70_000.00, "pago": "Nao"},      # C2 nao computado
]


def _cen_fin(*, inicio_efeito=date(2024, 3, 1), g_nao=(date(2024, 1, 1), date(2024, 2, 1)),
             sem_pagamento=()):
    def ajustar(wb):
        cen._base(wb, ciclos=_ciclos({1: inicio_efeito}),
                  metodo="Financeiro (Mensalidade)", ciclo_vigente="C1",
                  data_corte=date(2024, 12, 31))
        ws = wb["financeiro"]
        for i, (competencia, valor) in enumerate(cen._competencias(date(2023, 1, 1), 24, 42_500.00)):
            ws.cell(i + 2, 1).value = competencia
            if competencia not in sem_pagamento:
                ws.cell(i + 2, 3).value = valor
            ws.cell(i + 2, 7).value = "Nao" if competencia in g_nao else "Sim"
        cen._ciclo_em_execucao(wb, data=date(2024, 12, 31), linhas=cen.EXECUCAO_FISICA_PADRAO)
        cen._cobertura(wb, financeiro_ate=date(2024, 12, 31))
    return ajustar


def _cen_fin_multiciclo(wb):
    ciclos = _ciclos({1: date(2024, 3, 1), 2: date(2025, 1, 1), 3: date(2026, 2, 1)}, ate=3)
    cen._base(wb, ciclos=ciclos, metodo="Financeiro (Mensalidade)",
              ciclo_vigente="C3", data_corte=date(2026, 12, 31))
    nao = {date(2024, 1, 1), date(2024, 2, 1), date(2026, 1, 1)}
    ws = wb["financeiro"]
    for i, (competencia, valor) in enumerate(cen._competencias(date(2023, 1, 1), 48, 42_500.00)):
        ws.cell(i + 2, 1).value = competencia
        ws.cell(i + 2, 3).value = valor
        ws.cell(i + 2, 7).value = "Nao" if competencia in nao else "Sim"
    cen._ciclo_em_execucao(wb, data=date(2026, 12, 31), linhas=cen.EXECUCAO_FISICA_PADRAO)
    cen._cobertura(wb, financeiro_ate=date(2026, 12, 31))


def _cen_fin_sem_linhas(wb):
    """Ciclo C1 comeca em jan/2024, efeitos em mar/2024 e o `financeiro` NAO tem
    linha alguma em jan/fev: o periodo sem efeito existe so pelas datas."""
    cen._base(wb, ciclos=_ciclos({1: date(2024, 3, 1)}),
              metodo="Financeiro (Mensalidade)", ciclo_vigente="C1",
              data_corte=date(2024, 12, 31))
    ws = wb["financeiro"]
    ausentes = {date(2024, 1, 1), date(2024, 2, 1)}
    linhas = [(c, v) for c, v in cen._competencias(date(2023, 1, 1), 24, 42_500.00)
              if c not in ausentes]
    for i, (competencia, valor) in enumerate(linhas):
        ws.cell(i + 2, 1).value = competencia
        ws.cell(i + 2, 3).value = valor
        ws.cell(i + 2, 7).value = "Sim"
    cen._ciclo_em_execucao(wb, data=date(2024, 12, 31), linhas=cen.EXECUCAO_FISICA_PADRAO)
    cen._cobertura(wb, financeiro_ate=date(2024, 12, 31))


def _cen_itens(consumo, inicio_efeito=date(2024, 3, 1)):
    """`consumo`: quantidades QTD_CONS_C1 por item (None = nao informado)."""
    def ajustar(wb):
        cen._base(wb, ciclos=_ciclos({1: inicio_efeito}), metodo="Itens Consumidos",
                  ciclo_vigente="C1", data_corte=date(2024, 12, 31))
        cen._itens_consumidos(wb, [
            {"codigo": "ITEM-001", "quantidade": 75.0, "valor_unitario": 250.00},
            {"codigo": "ITEM-002", "quantidade": 50.0, "valor_unitario": 1_500.00},
            {"codigo": "ITEM-003", "quantidade": 28.0, "valor_unitario": 3_200.00},
        ])
        ws = wb["itens_Consumidos"]
        for linha, quantidade in zip((2, 3, 4), consumo):
            if quantidade is not None:
                ws.cell(linha, 7).value = quantidade        # QTD_CONS_C1
        cen._ciclo_em_execucao(wb, data=date(2024, 12, 31), linhas=cen.EXECUCAO_FISICA_PADRAO)
    return ajustar


# --- leitura do quadro ------------------------------------------------------ #
def _quadro(vals):
    res, mem = vals[ap.ABA_RESULTADOS], vals[ap.ABA_MEMORIA]
    return {
        "titulo": res["E15"].value, "hdr": [res[f"{c}15"].value for c in "FGH"],
        "nota": res["E16"].value, "caso": res["E21"].value,
        "linhas": {n: [res[f"{c}{linha}"].value for c in "EFGH"]
                   for n, (_c, _p, _cc, linha) in ap.CICLOS.items()},
        "estado": {n: mem[f"{c}{ap.L_ESTADO}"].value for n, (c, *_r) in ap.CICLOS.items()},
        "qtd": {n: mem[f"{c}{ap.L_QTD}"].value for n, (c, *_r) in ap.CICLOS.items()},
        "mem": mem,
    }


@pytest.fixture(scope="module")
def tpf(tmp_path_factory):
    return tmp_path_factory


def _calc(nome, ajustar, tpf, template=TEMPLATE):
    return _recalcular(_montar(template, ajustar), nome, tpf)


def _arred(valor: Decimal) -> float:
    return float(valor.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


# ---- PC --------------------------------------------------------------------- #
@com
def test_pc_anterior_entra_posterior_c0_e_ciclo_nao_computado_nao(tpf):
    vals = _calc("pc_misto", _cen_pc(PEDIDOS_MISTOS), tpf)
    q = _quadro(vals)
    assert q["estado"][1] == "A" and q["estado"][2] is None
    # So PC-001 e PC-002 (CICLO_PC = C1 e DATA_PC < 01/07/2024).
    assert q["qtd"][1] == 2
    mem = q["mem"]
    assert mem[f"T{ap.L_ORIGINAL}"].value == 276_500.00
    assert mem[f"T{ap.L_TERIAM}"].value == 290_656.80          # 180.000x1,0512 + 96.500x1,0512
    assert mem[f"T{ap.L_DIFERENCA}"].value == 14_156.80
    assert q["linhas"][1][1:] == [276_500.00, 290_656.80, 14_156.80]
    assert q["linhas"][1][0] == (
        "C1 — efeitos a partir de 01/07/2024\n2 PC(s) · 01/01/2024 a 30/06/2024"
    )
    assert q["titulo"] == "EXECUÇÃO SEM EFEITO FINANCEIRO"
    assert q["hdr"] == ["Valor original", "Valor que teriam com o reajuste",
                        "Diferença sem efeito financeiro"]
    assert q["caso"] in (None, "")
    # C0 e C2 existem como PCs "Nao", mas o legado ja os excluia: o filtro novo
    # reconcilia com T28:T30 (nenhum segundo calculo concorrente).
    assert mem["T28"].value == 2
    assert mem["T29"].value == 276_500.00
    assert mem["T30"].value == 14_156.80


@com
def test_pc_sem_pedido_no_periodo_e_zero_conhecido(tpf):
    posteriores = [p for p in PEDIDOS_MISTOS if p["numero"] == "PC-003"]
    q = _quadro(_calc("pc_sem_periodo", _cen_pc(posteriores), tpf))
    assert q["estado"][1] == "B" and q["qtd"][1] == 0
    assert q["linhas"][1] == [
        "C1 — efeitos a partir de 01/07/2024\nSem execução · 01/01/2024 a 30/06/2024",
        0, 0, 0,
    ]


@com
def test_pc_sem_atraso_dos_efeitos_informa_discretamente(tpf):
    q = _quadro(_calc("pc_sem_atraso", _cen_pc(PEDIDOS_MISTOS, date(2024, 1, 1)), tpf))
    assert q["estado"][1] == "D"
    # Sem duplicacao: a linha do ciclo so identifica o ciclo; a frase explicativa
    # aparece uma unica vez, no rodape do quadro.
    assert q["linhas"][1] == ["C1 — efeitos desde o início do ciclo", None, None, None]
    assert q["caso"] == (
        "Não foi identificada execução anterior ao início dos efeitos financeiros."
    )
    assert q["hdr"] == [None, None, None]


# ---- Financeiro ------------------------------------------------------------- #
@com
def test_financeiro_duas_competencias_antes_do_efeito(tpf):
    q = _quadro(_calc("fin_duas", _cen_fin(), tpf))
    assert q["estado"][1] == "A" and q["qtd"][1] == 2
    assert q["linhas"][1] == [
        "C1 — efeitos a partir de 03/2024\n2 competência(s) · 01/2024 a 02/2024",
        85_000.00, 89_352.00, 4_352.00,                       # 2 x 42.500 x 1,0512
    ]
    assert q["hdr"][0] == "Valor pago"


@com
def test_financeiro_competencia_com_g_sim_fica_fora(tpf):
    q = _quadro(_calc("fin_g_sim", _cen_fin(g_nao=(date(2024, 1, 1),)), tpf))
    assert q["qtd"][1] == 1
    assert q["linhas"][1][1:] == [42_500.00, 44_676.00, 2_176.00]
    assert q["linhas"][1][0].endswith("01/2024")


@com
def test_financeiro_sem_pagamento_no_periodo_e_zero_conhecido(tpf):
    sem = (date(2024, 1, 1), date(2024, 2, 1))
    q = _quadro(_calc("fin_sem_pagto", _cen_fin(sem_pagamento=sem), tpf))
    assert q["estado"][1] == "B" and q["qtd"][1] == 0
    assert q["linhas"][1][1:] == [0, 0, 0]
    assert q["linhas"][1][0].splitlines()[1] == "Sem execução · 01/2024 a 02/2024"


@com
def test_financeiro_multiciclo_usa_o_fator_encadeado_do_motor(tpf):
    q = _quadro(_calc("fin_multi", _cen_fin_multiciclo, tpf))
    assert q["estado"] == {1: "A", 2: "D", 3: "A", 4: None}
    assert q["linhas"][1][1:] == [85_000.00, 89_352.00, 4_352.00]
    assert q["linhas"][2][1:] == [None, None, None]
    fator_c3 = Decimal("1.0512") * Decimal("1.0374") * Decimal("1.0289")
    esperado = _arred(Decimal("42500") * fator_c3)
    assert q["linhas"][3][1] == 42_500.00
    assert q["linhas"][3][2] == pytest.approx(esperado, abs=0.011)
    assert q["linhas"][3][3] == pytest.approx(esperado - 42_500.00, abs=0.011)
    descricao = q["linhas"][3][0].splitlines()
    assert descricao[0] == "C3 — efeitos a partir de 02/2026"
    assert descricao[1] == "1 competência(s) · 01/2026"


# ---- Itens consumidos ------------------------------------------------------- #
@com
def test_consumidos_sem_granularidade_temporal_nao_estima(tpf):
    q = _quadro(_calc("itens_nm", _cen_itens((75.0, 50.0, 28.0)), tpf))
    assert q["estado"][1] == "C"
    linha = q["linhas"][1]
    # Nenhum valor monetario: nada de rateio, proporcao ou R$ 0,00 inventado.
    assert linha[1] is None and linha[2] is None
    assert linha[3] == "Não mensurável"
    assert linha[0] == (
        "C1 — efeitos a partir de 01/03/2024\nSem efeito: 01/01/2024 a 29/02/2024"
    )
    assert q["caso"].startswith("Execução sem efeito financeiro identificada, mas o impacto")
    assert q["hdr"] == [None, None, None]
    mem = q["mem"]
    for linha_mem in (ap.L_ORIGINAL, ap.L_TERIAM, ap.L_DIFERENCA):
        assert mem[f"T{linha_mem}"].value in (None, "")


@com
def test_consumidos_com_consumo_zero_e_zero_conhecido(tpf):
    q = _quadro(_calc("itens_zero", _cen_itens((0.0, 0.0, 0.0)), tpf))
    assert q["estado"][1] == "B"
    assert q["linhas"][1][1:] == [0, 0, 0]


@com
def test_consumidos_sem_consumo_informado_nao_mostra_quadro(tpf):
    q = _quadro(_calc("itens_vazio", _cen_itens((None, None, None)), tpf))
    assert q["estado"][1] is None and q["titulo"] in (None, "")
    assert q["linhas"][1] == [None, None, None, None]


@com
def test_consumidos_sem_atraso_dos_efeitos(tpf):
    q = _quadro(_calc("itens_d", _cen_itens((75.0, 50.0, 28.0), date(2024, 1, 1)), tpf))
    assert q["estado"][1] == "D" and q["caso"].startswith("Não foi identificada execução")


# ---- Seguranca: VTA e retroativos nao mudam (A/B) --------------------------- #
CENARIOS_AB = {
    "pc": _cen_pc(PEDIDOS_MISTOS),
    "financeiro": _cen_fin(),
    "itens": _cen_itens((75.0, 50.0, 28.0)),
}
# =NOW() em CONTROLE!B14 e o espelho dele em cobertura_temporal!B4: volateis por
# natureza (segundos de diferenca entre as duas execucoes).
IGNORAR_AB = {("CONTROLE", "B14"), ("cobertura_temporal", "B4")}


def _fora_da_allowlist(aba: str, linha: int, coluna: int) -> bool:
    if aba == ap.ABA_MEMORIA and 78 <= linha <= 120 and 19 <= coluna <= 23:
        return False
    if aba == ap.ABA_RESULTADOS and 15 <= linha <= 21 and 5 <= coluna <= 8:
        return False
    return True


def _comparar_ab(ajustar, chave, tpf):
    """Recalcula o MESMO cenario com o template base e o novo e compara."""
    base = _bytes_template_base()
    if base is None:
        pytest.skip(f"template do commit {COMMIT_BASE[:7]} indisponivel (git)")
    antes = _recalcular(_montar(io.BytesIO(base), ajustar), f"ab_antes_{chave}", tpf)
    depois = _recalcular(_montar(TEMPLATE, ajustar), f"ab_depois_{chave}", tpf)

    # Grandezas de negocio, explicitamente.
    for aba, celula in (
        ("RESULTADOS_DETALHE", "B3"), ("RESULTADOS_DETALHE", "C5"), ("RESULTADOS_DETALHE", "D5"),
        ("RESULTADOS_DETALHE", "B22"), ("RESULTADOS_DETALHE", "C22"), ("RESULTADOS_DETALHE", "D22"),
        ("RESULTADOS_DETALHE", "G22"), ("RESULTADOS_DETALHE", "H5"), ("RESULTADOS_DETALHE", "B36"),
        ("RESULTADOS_DETALHE", "B38"), ("MEMORIA_RESULTADOS", "B26"),
        ("MEMORIA_RESULTADOS", "B16"), ("MEMORIA_RESULTADOS", "T38"),
        ("MEMORIA_RESULTADOS", "T39"), ("MEMORIA_RESULTADOS", "D35"),
    ):
        assert antes[aba][celula].value == depois[aba][celula].value, (aba, celula)
    assert depois[_aba_tecnica_resultados(depois)]["C5"].value not in (None, "")     # VTA calculado

    # Todo o resto, celula a celula.
    assert antes.sheetnames == depois.sheetnames
    diferencas = []
    for aba in antes.sheetnames:
        a, d = antes[aba], depois[aba]
        for linha in range(1, max(a.max_row, d.max_row) + 1):
            for coluna in range(1, max(a.max_column, d.max_column) + 1):
                if not _fora_da_allowlist(aba, linha, coluna):
                    continue
                coordenada = a.cell(linha, coluna).coordinate
                if (aba, coordenada) in IGNORAR_AB:
                    continue
                if a.cell(linha, coluna).value != d.cell(linha, coluna).value:
                    diferencas.append(f"{aba}!{coordenada}")
    assert not diferencas, diferencas[:20]
    return depois


@com
@pytest.mark.parametrize("metodo", sorted(CENARIOS_AB))
def test_vta_e_retroativos_identicos_antes_e_depois(metodo, tpf):
    _comparar_ab(CENARIOS_AB[metodo], metodo, tpf)


# ---- Financeiro: periodo sem efeito existe so pelas datas (revisao PR #172) -- #
@com
def test_financeiro_periodo_sem_efeito_sem_linhas_e_zero_conhecido(tpf):
    depois = _comparar_ab(_cen_fin_sem_linhas, "fin_sem_linhas", tpf)   # VTA/retro iguais
    q = _quadro(depois)
    assert q["estado"][1] == "B" and q["qtd"][1] == 0
    descricao, original, com_reajuste, diferenca = q["linhas"][1]
    assert descricao == (
        "C1 — efeitos a partir de 03/2024\nSem execução · 01/2024 a 02/2024"
    )
    assert "desde o início do ciclo" not in descricao          # nao e o estado D
    assert (original, com_reajuste, diferenca) == (0, 0, 0)
    assert q["caso"] in (None, "")
