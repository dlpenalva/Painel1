"""Coleta 11.7 no Excel real — Nxxx valorizado pela base economica do VU.

N001 nasce em C3 por R$ 10,00 com C1 = 3,81% e 100 unidades remanescentes:

* A (base C0): aditivos, historico_VU!F, CICLO_EM_EXECUCAO!E = 10,38 e o VTA
  sobe exatamente 38,00 (1.622,86 -> 1.660,86); C0/C1/C2 seguem vazios;
* B (base C1): fator 1,0000, VU 10,00 e VTA da Coleta 11.6;
* item existente com simples acrescimo: historico herdado, nada muda;
* ciclo em execucao C4 sem fator proprio: o fallback da CICLO_EM_EXECUCAO
  usa a mesma base;
* N001 criado por duas linhas "novo item" (mesma data, H=Nao): alerta nas
  duas linhas e VU vazio (fail-closed), sem cair no nascimento.
"""
from __future__ import annotations

import gc
import os
import time
import zipfile
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import pytest

from _coleta_oficial import gerar_coleta_oficial_preenchida


pytestmark = pytest.mark.skipif(
    os.environ.get("RUN_EXCEL_INTEGRATION") != "1",
    reason="defina RUN_EXCEL_INTEGRATION=1 para executar o Excel COM",
)

NOMES_VTA = (
    "VTA_FINAL",
    "VTA_SEM_POTENCIAL",
    "RETRO_OFICIAL",
    "RETROATIVO_POTENCIAL_APURADO",
    "RETROATIVO_POTENCIAL_VTA",
)
# Coleta 11.6 (main b4a3287), cenario A: N001 valorizado a 10,00.
VTA_11_6 = {
    "VTA_FINAL": 1622.86,
    "VTA_SEM_POTENCIAL": 1500.0,
    "RETRO_OFICIAL": 22.86,
    "RETROATIVO_POTENCIAL_APURADO": 0.0,
    "RETROATIVO_POTENCIAL_VTA": 0.0,
}


SUPERFICIES_VTA = (
    ("RESULTADOS", "B9"),
    ("RESULTADOS", "C18"),
    ("RESULTADOS_DETALHE", "B86"),
    ("RESULTADOS_DETALHE", "C86"),
    ("comparativo_VTA", "B207"),
    ("MEMORIA_RESULTADOS", "B28"),
)
ERRO_EXCEL = -2146826000  # valores de erro (#VALUE!, #N/A...) chegam como int < isso


def _sem_erro(valores) -> bool:
    return not any(isinstance(v, int) and v < ERRO_EXCEL for v in valores)


def _dados() -> dict:
    return {
        "origem": "Validacao Coleta 11.7",
        "indice": "IST",
        "data_base_original": "01/01/2023",
        "data_corte": date(2026, 8, 31),
        "ciclos": [{
            "ciclo": "C1",
            "data_inicio": date(2024, 1, 1),
            "data_fim": date(2024, 12, 31),
            "data_pedido": date(2024, 1, 1),
            "financeiro_inicio": date(2024, 1, 1),
            "percentual_aplicado": 0.0381,
            "objeto_analise_atual": True,
        }],
    }


def _tentar(pythoncom, acao):
    ultimo = None
    for _ in range(30):
        try:
            return acao()
        except Exception as exc:  # pragma: no cover - somente Excel COM
            ultimo = exc
            codigo = getattr(exc, "hresult", None)
            if codigo is None and getattr(exc, "args", ()):
                codigo = exc.args[0]
            if codigo != -2147418111:
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(0.2)
    raise ultimo


def _preencher(
    livro, base_vu: str, *, acrescimo_existente: bool, corte=None,
    segunda_inclusao: str | None = None,
) -> None:
    from pywintypes import Time

    livro.Worksheets("CONTROLE").Range("B1").Value = "Financeiro (Mensalidade)"
    if corte is not None:
        livro.Worksheets("CONTROLE").Range("B3").Value = Time(corte)
    livro.Worksheets("financeiro").Range("C14").Value = 600
    itens = livro.Worksheets("itens_Remanesc")
    for linha, item, base, c1, c3 in (
        (2, "1.1", 1000, 950, 850),
        (3, "N001", None, None, None),
    ):
        itens.Range(f"A{linha}").Value = item
        if base is not None:
            itens.Range(f"B{linha}").Value = base
        itens.Range(f"C{linha}").Value = 10
        if c1 is not None:
            itens.Range(f"E{linha}").Value = c1
        if c3 is not None:
            itens.Range(f"I{linha}").Value = c3
    eventos = [("N001", "Acréscimo - novo item", base_vu)]
    if segunda_inclusao:
        # Mesma data e H=Nao: o unico caso que a 11.6 deixava sem alerta.
        eventos.append(("N001", "Acréscimo - novo item", segunda_inclusao))
    if acrescimo_existente:
        eventos.append(("1.1", "Acrescimo", None))
    aditivos = livro.Worksheets("aditivos")
    for linha, (item, tipo, base) in enumerate(eventos, start=2):
        aditivos.Range(f"A{linha}").Value = item
        aditivos.Range(f"B{linha}").Value = Time(datetime(2026, 2, 1))
        aditivos.Range(f"D{linha}").Value = tipo
        aditivos.Range(f"E{linha}").Value = 50 if segunda_inclusao else 100
        aditivos.Range(f"H{linha}").Value = "Nao" if segunda_inclusao else "Sim"
        aditivos.Range(f"K{linha}").Value = "Sim"
        if base:
            aditivos.Range(f"N{linha}").Value = base
    ciclo = livro.Worksheets("CICLO_EM_EXECUCAO")
    ciclo.Range("D5").Value2 = livro.Worksheets("CONTROLE").Range("B3").Value2
    livro.Application.CalculateFullRebuild()
    ciclo.Range("C13").Value = 900 if acrescimo_existente else 800
    ciclo.Range("C14").Value = 100


def _ler(livro) -> dict:
    aditivos = livro.Worksheets("aditivos")
    historico = livro.Worksheets("historico_VU")
    ciclo = livro.Worksheets("CICLO_EM_EXECUCAO")
    valor, quantidade = aditivos.Range("J2").Value, aditivos.Range("L2").Value
    return {
        "aditivo_I": aditivos.Range("I2").Value,
        # Base invalida deixa J vazio: sem VU atualizado do aditivo.
        "aditivo_VU": valor / quantidade if isinstance(valor, float) and quantidade else None,
        "aditivo_M": aditivos.Range("M2").Value,
        "aditivo_M3": aditivos.Range("M3").Value,
        "hist_N001": [historico.Range(f"{c}3").Value for c in "CDEFG"],
        "hist_11": [historico.Range(f"{c}2").Value for c in "CDEFG"],
        "cee_ciclo": ciclo.Range("C3").Value,
        "cee_E": (ciclo.Range("E13").Value, ciclo.Range("E14").Value),
        "cee_G": (ciclo.Range("G13").Value, ciclo.Range("G14").Value),
        "vta": {nome: livro.Names(nome).RefersToRange.Value2 for nome in NOMES_VTA},
        "gate": livro.Worksheets("MEMORIA_RESULTADOS").Range("T48").Value,
        # Consumidores do VTA: herdam o vazio, sem erro de formula.
        "superficies": [
            livro.Worksheets(aba).Range(celula).Value
            for aba, celula in SUPERFICIES_VTA
        ],
    }


def _rodar(tmp_path: Path, nome: str, base_vu, *, corrigir=None, **kwargs) -> dict:
    import pythoncom
    import win32com.client

    caminho = tmp_path / f"Coleta_11_7_{nome}.xlsx"
    caminho.write_bytes(gerar_coleta_oficial_preenchida(_dados()))
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    livro = None
    try:
        abrir = lambda: excel.Workbooks.Open(  # noqa: E731
            str(caminho.resolve()), UpdateLinks=0, ReadOnly=False, CorruptLoad=0
        )
        livro = _tentar(pythoncom, abrir)
        time.sleep(1.0)
        _tentar(pythoncom, lambda: _preencher(livro, base_vu, **kwargs))
        _tentar(pythoncom, excel.CalculateFullRebuild)
        resultado = _tentar(pythoncom, lambda: _ler(livro))
        if corrigir is not None:
            _tentar(pythoncom, lambda: corrigir(livro))
            _tentar(pythoncom, excel.CalculateFullRebuild)
            resultado["corrigido"] = _tentar(pythoncom, lambda: _ler(livro))
        _tentar(pythoncom, livro.Save)
        _tentar(pythoncom, lambda: livro.Close(SaveChanges=False))
        livro = _tentar(pythoncom, abrir)
        _tentar(pythoncom, excel.CalculateFullRebuild)
        resultado["reaberto"] = _tentar(pythoncom, lambda: _ler(livro))
        _tentar(pythoncom, lambda: livro.Close(SaveChanges=False))
        livro = None
    finally:
        if livro is not None:
            try:
                livro.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        excel = None
        gc.collect()
        pythoncom.CoUninitialize()
    with zipfile.ZipFile(BytesIO(caminho.read_bytes())) as pacote:
        assert not any(b"repairLoad" in pacote.read(n) for n in pacote.namelist())
        assert any(
            b"x14:conditionalFormattings" in pacote.read(n)
            for n in pacote.namelist()
            if n.startswith("xl/worksheets/")
        )
    return resultado


def test_cenario_a_base_c0_propaga_10_38_ate_o_vta(tmp_path: Path):
    r = _rodar(tmp_path, "A", "C0", acrescimo_existente=False)
    assert r["aditivo_I"] == pytest.approx(1.0381)
    assert r["aditivo_VU"] == pytest.approx(10.38)
    assert r["aditivo_M"] == "OK"
    assert r["hist_N001"][:3] == ["", "", ""]  # C0/C1/C2: item ainda nao existe
    assert r["hist_N001"][3] == pytest.approx(10.38)
    assert r["cee_ciclo"] == "C3"
    assert r["cee_E"] == (pytest.approx(10.38), pytest.approx(10.38))
    assert r["cee_G"][1] == pytest.approx(1038.0)
    assert r["vta"] == {
        **VTA_11_6,
        "VTA_FINAL": pytest.approx(1660.86),
        "VTA_SEM_POTENCIAL": pytest.approx(1538.0),
        "RETRO_OFICIAL": pytest.approx(22.86),
    }
    assert r["vta"]["VTA_FINAL"] - VTA_11_6["VTA_FINAL"] == pytest.approx(38.0)
    assert r["reaberto"] == {k: v for k, v in r.items() if k != "reaberto"}


def test_cenario_b_base_c1_mantem_vu_10(tmp_path: Path):
    r = _rodar(tmp_path, "B", "C1", acrescimo_existente=False)
    assert r["aditivo_I"] == pytest.approx(1.0)
    assert r["aditivo_VU"] == pytest.approx(10.0)
    assert r["hist_N001"] == ["", "", "", pytest.approx(10.0), ""]
    assert r["cee_E"][1] == pytest.approx(10.0)
    assert r["vta"] == pytest.approx(VTA_11_6)


def test_item_existente_com_acrescimo_herda_o_historico(tmp_path: Path):
    r = _rodar(tmp_path, "C", "C0", acrescimo_existente=True)
    assert r["hist_11"] == [10.0, pytest.approx(10.38), "", "", ""]
    assert r["cee_E"][0] == pytest.approx(10.38)
    assert r["cee_G"][0] == pytest.approx(9342.0)  # 900 x 10,38, como na 11.6
    assert r["hist_N001"][3] == pytest.approx(10.38)


@pytest.mark.parametrize("segunda", ("C1", "C0"), ids=("3c_bases_C0_C1", "3d_ambas_C0"))
def test_nxxx_com_duas_inclusoes_alerta_e_vu_vazio(tmp_path: Path, segunda):
    r = _rodar(tmp_path, f"dup_{segunda}", "C0", acrescimo_existente=False,
               segunda_inclusao=segunda)
    alerta = "ALERTA: NOVO_ITEM_COM_MAIS_DE_UMA_INCLUSAO"
    assert (r["aditivo_M"], r["aditivo_M3"]) == (alerta, alerta)
    assert r["aditivo_I"] in (None, "")
    # Fail-closed: nem a base C0 nem o nascimento (10,00) sao escolhidos.
    assert r["hist_N001"] == ["", "", "", "", ""]
    assert r["cee_E"] == (pytest.approx(10.38), "")
    assert r["cee_G"][1] in (None, "")
    assert r["hist_11"] == [10.0, pytest.approx(10.38), "", "", ""]


def _vta_indisponivel(r: dict, alerta: str) -> None:
    assert str(r["aditivo_M"]).startswith(alerta)
    assert r["gate"] >= 1
    assert r["vta"]["VTA_FINAL"] in (None, "")
    assert r["vta"]["VTA_SEM_POTENCIAL"] in (None, "")
    assert _sem_erro(r["superficies"])
    assert not any(isinstance(v, float) and v > 0 for v in r["superficies"][:4])


def test_gate_vta_nxxx_duplicado(tmp_path: Path):
    r = _rodar(tmp_path, "gate_dup", "C0", acrescimo_existente=False, segunda_inclusao="C1")
    _vta_indisponivel(r, "ALERTA: NOVO_ITEM_COM_MAIS_DE_UMA_INCLUSAO")


def test_gate_vta_base_economica_obrigatoria_ausente(tmp_path: Path):
    r = _rodar(tmp_path, "gate_sem_base", None, acrescimo_existente=False)
    _vta_indisponivel(r, "ALERTA: BASE_VU_OBRIGATORIA")


def test_gate_vta_base_economica_posterior(tmp_path: Path):
    # Base C3 com fator-alvo C1 (unico fator conhecido): posterior ao alvo.
    r = _rodar(tmp_path, "gate_base_posterior", "C3", acrescimo_existente=False)
    _vta_indisponivel(r, "ALERTA: BASE_VU_INVALIDA_OU_POSTERIOR_AO_FATOR_ALVO")


def test_gate_vta_corrigido_no_mesmo_xls_volta_ao_valor(tmp_path: Path):
    def corrigir(livro):
        livro.Worksheets("aditivos").Range("N2").Value = "C0"

    r = _rodar(tmp_path, "gate_corrigido", None, acrescimo_existente=False, corrigir=corrigir)
    _vta_indisponivel(r, "ALERTA: BASE_VU_OBRIGATORIA")
    c = r["corrigido"]
    assert c["aditivo_M"] == "OK"
    assert c["gate"] == 0
    assert c["vta"]["VTA_FINAL"] == pytest.approx(1660.86)
    assert c["vta"]["VTA_SEM_POTENCIAL"] == pytest.approx(1538.0)
    assert c["hist_N001"][3] == pytest.approx(10.38)


@pytest.mark.parametrize(("base_vu", "esperado"), (("C0", 10.38), ("C1", 10.0)))
def test_fallback_c4_sem_fator_usa_a_mesma_base(tmp_path: Path, base_vu, esperado):
    r = _rodar(
        tmp_path, f"D_{base_vu}", base_vu,
        acrescimo_existente=False, corte=datetime(2027, 3, 31),
    )
    assert r["cee_ciclo"] == "C4"
    assert r["hist_N001"][4] == ""  # C4 sem fator: historico nao inventa
    assert r["cee_E"] == (pytest.approx(10.38), pytest.approx(esperado))
