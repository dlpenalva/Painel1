# -*- coding: utf-8 -*-
"""CONSUMO-GLOSA-1 — prova em Excel real da cadeia completa ate o VTA.

Roda o caso focal em COPIAS TEMPORARIAS (o template versionado nunca e
tocado), sempre com o ciclo CalculateFullRebuild -> Save -> Close -> Reopen,
conferindo os valores DEPOIS da reabertura.

Caso focal (mesmos numeros do enunciado):
    item I1, QTD_CONTRATADA 1500, VU_ORIGINAL 100,00
    QTD_CONS_C1 1000  ->  valor calculado da execucao = 100.000,00
    percentual de C1 = 8%  ->  fator acumulado = 1,08

Cenarios:
    A  legado (checkpoint x novo, ambos sem ajuste)   -> tudo igual
    B  valor pago 90.000                              -> 97.200 / retro 7.200
    C  glosa 10.000                                   -> identico a B
    D  valor pago 0                                   -> glosa integral
    E  valor pago 105.000 (> calculado)               -> REVISAR, fail-closed
    F  glosa 110.000 (> calculado)                    -> REVISAR, fail-closed
"""
from __future__ import annotations

import argparse
import shutil
import tempfile
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client

XL_CALC_AUTOMATIC = -4105

# Celulas do checkpoint (existem nos DOIS arquivos) — base da comparacao A.
CELULAS_LEGADO = {
    "itens_Consumidos": ("H2", "O2", "P2", "Q2", "V2"),
    "MEMORIA_RESULTADOS": (
        "D10", "D11", "D12", "D13", "D14", "D15",
        "F20", "B21", "C33", "D33", "B35", "C35", "D35", "B26",
    ),
}
# Celulas que so existem apos esta frente.
# Cenario G (multiciclo). Separado em legado x novo: o checkpoint nao tem as
# celulas novas, entao a comparacao "checkpoint == novo" so pode olhar o legado.
CELULAS_MULTICICLO_LEGADO = {
    "itens_Consumidos": ("J2", "O2", "V2"),
    "MEMORIA_RESULTADOS": (
        "D12", "D15", "F20", "C33", "D33", "D35", "B26",
    ),
    "RESULTADOS": ("D22", "B83", "B86"),
}
CELULAS_MULTICICLO_NOVAS = {
    "itens_Consumidos": (
        "Y4", "Z4", "AA4", "AB4", "AC4", "AD4", "AE4", "AF4", "AG4",
    ),
    "MEMORIA_RESULTADOS": ("T70", "T71", "T72", "T73", "T74", "T75"),
}
CELULAS_NOVAS = {
    "itens_Consumidos": (
        "Y3", "Z3", "AA3", "AB3", "AC3", "AD3", "AE3", "AF3", "AG3",
    ),
    "MEMORIA_RESULTADOS": ("T70", "T71", "T72", "T73", "T74", "T75"),
}
# RESULTADOS nao ganhou celula nova; o que se prova aqui e que ela reflete o
# ajuste SOZINHA, pelas celulas que ja existiam no checkpoint.
CELULAS_RESULTADOS = {
    "RESULTADOS": ("B36", "D22", "B83", "B85", "B86", "C5"),
}


def _semear(wb, ajuste=None, multiciclo=False) -> None:
    """Entradas minimas do caso focal. `ajuste` = (tipo, valor) ou None.

    CONTROLE e as demais abas de entrada sao protegidas no template; como
    isto roda sempre em copia temporaria, basta desproteger para semear (o
    arquivo versionado nunca passa por aqui).

    multiciclo=True monta o cenario do gate corretivo: C1 JA FORMALIZADO
    (8%, fora da apuracao) e C2 em apuracao (5%). Ali o fator acumulado
    (1,134) e o fator novo (1,05) sao diferentes, entao F-1 = 0,134 nunca
    coincide com F-F/D = 0,054 — e reaplicar o acumulado sobre o valor pago
    ficaria visivel.
    """
    for nome in ("CONTROLE", "parametros", "itens_Consumidos"):
        try:
            wb.Worksheets(nome).Unprotect()
        except Exception:
            pass

    controle = wb.Worksheets("CONTROLE")
    controle.Range("B1").Value = "Itens Consumidos"
    parametros = wb.Worksheets("parametros")
    consumidos = wb.Worksheets("itens_Consumidos")

    consumidos.Range("A2").Value = "I1"
    consumidos.Range("B2").Value = 1500
    consumidos.Range("C2").Value = 100

    if multiciclo:
        controle.Range("B2").Value = "C2"
        parametros.Range("A2").Value = "Sim"
        parametros.Range("A3").Value = "Nao"   # C1 ja formalizado
        parametros.Range("A4").Value = "Sim"   # C2 em apuracao
        parametros.Range("E3").Value = 0.08
        parametros.Range("E4").Value = 0.05
        consumidos.Range("I2").Value = 1000    # QTD_CONS_C2
        linha_ajuste = 4                       # bloco lateral: C2
    else:
        controle.Range("B2").Value = "C1"
        parametros.Range("A2").Value = "Sim"
        parametros.Range("A3").Value = "Sim"
        parametros.Range("E3").Value = 0.08
        consumidos.Range("G2").Value = 1000    # QTD_CONS_C1
        linha_ajuste = 3                       # bloco lateral: C1

    if ajuste is not None:
        tipo, valor = ajuste
        consumidos.Range(f"Z{linha_ajuste}").Value = tipo
        consumidos.Range(f"AA{linha_ajuste}").Value = valor


def _ler(wb, mapas) -> dict[str, object]:
    abas = [ws.Name for ws in wb.Worksheets]
    valores: dict[str, object] = {}
    for mapa in mapas:
        for aba, enderecos in mapa.items():
            if aba not in abas:
                continue
            ws = wb.Worksheets(aba)
            for endereco in enderecos:
                try:
                    valores[f"{aba}!{endereco}"] = ws.Range(endereco).Value
                except Exception:
                    valores[f"{aba}!{endereco}"] = "<erro de leitura>"
    return valores


def _com(acao, tentativas: int = 12, espera: float = 1.0):
    """Repete uma chamada COM enquanto o Excel devolver RPC_E_CALL_REJECTED.

    Excel recusa chamadas enquanto esta ocupado (recalculo pesado, salvamento);
    a rejeicao e transitoria e nao indica erro de logica.
    """
    ultimo = None
    for tentativa in range(tentativas):
        try:
            return acao()
        except pywintypes.com_error as exc:  # noqa: PERF203
            if exc.hresult not in (-2147418111, -2147417846):
                raise
            ultimo = exc
            time.sleep(espera * (tentativa + 1))
    raise RuntimeError(f"Excel permaneceu ocupado apos {tentativas} tentativas: {ultimo}")


def _fechar(wb) -> None:
    """Fecha o workbook de forma IDEMPOTENTE.

    Se um Close for aceito e so depois o Excel devolver RPC_E_CALL_REJECTED,
    a retentativa cai sobre um workbook ja invalido e o pywin32 passa a
    resolver `Close` como bool em vez de metodo. Checar `callable` antes de
    cada tentativa cobre exatamente esse caso.
    """
    for tentativa in range(12):
        metodo = getattr(wb, "Close", None)
        if not callable(metodo):
            return
        try:
            metodo(SaveChanges=False)
            return
        except pywintypes.com_error as exc:
            if exc.hresult not in (-2147418111, -2147417846):
                return
            time.sleep(1.0 * (tentativa + 1))


def _rodar(
    excel, origem: Path, ajuste=None, mapas=None, multiciclo=False
) -> dict[str, object]:
    """Semeia, recalcula, salva, fecha e REABRE antes de ler."""
    mapas = mapas or (CELULAS_LEGADO, CELULAS_RESULTADOS, CELULAS_NOVAS)
    tmp_dir = Path(tempfile.mkdtemp(prefix="cl8us_verif_glosa_"))
    alvo = tmp_dir / origem.name
    shutil.copyfile(origem, alvo)
    wb = _com(lambda: excel.Workbooks.Open(
        str(alvo), UpdateLinks=0, ReadOnly=False, CorruptLoad=0
    ))
    try:
        _semear(wb, ajuste, multiciclo=multiciclo)
        excel.Calculation = XL_CALC_AUTOMATIC
        _com(excel.CalculateFullRebuild)
        _com(wb.Save)
    finally:
        _fechar(wb)

    wb = _com(lambda: excel.Workbooks.Open(
        str(alvo), UpdateLinks=0, ReadOnly=True, CorruptLoad=0
    ))
    try:
        return _ler(wb, mapas)
    finally:
        _fechar(wb)
        shutil.rmtree(tmp_dir, ignore_errors=True)


def _fmt(valor) -> str:
    if isinstance(valor, float):
        return f"{valor:,.2f}"
    return "(vazio)" if valor in (None, "") else str(valor)


def _tabela(titulo: str, valores: dict[str, object]) -> None:
    print(f"\n--- {titulo} ---")
    for chave, valor in valores.items():
        print(f"  {chave:<34} = {_fmt(valor)}")


def _iguais(a, b) -> bool:
    if isinstance(a, float) and isinstance(b, float):
        return round(a, 2) == round(b, 2)
    return a == b


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("checkpoint", type=Path)
    parser.add_argument("novo", type=Path)
    args = parser.parse_args()

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    falhas: list[str] = []
    try:
        base_check = _rodar(
            excel, args.checkpoint.resolve(),
            mapas=(CELULAS_LEGADO, CELULAS_RESULTADOS),
        )
        base_novo = _rodar(excel, args.novo.resolve())
        pago = _rodar(excel, args.novo.resolve(), ("Valor pago", 90000))
        glosa = _rodar(excel, args.novo.resolve(), ("Glosa", 10000))
        zero = _rodar(excel, args.novo.resolve(), ("Valor pago", 0))
        pago_maior = _rodar(excel, args.novo.resolve(), ("Valor pago", 105000))
        glosa_maior = _rodar(excel, args.novo.resolve(), ("Glosa", 110000))
        # G — multiciclo: C1 ja formalizado (8%), C2 em apuracao (5%).
        mapas_mc = (CELULAS_MULTICICLO_LEGADO, CELULAS_MULTICICLO_NOVAS)
        mc_check = _rodar(
            excel, args.checkpoint.resolve(),
            mapas=(CELULAS_MULTICICLO_LEGADO,), multiciclo=True,
        )
        mc_base = _rodar(
            excel, args.novo.resolve(), mapas=mapas_mc, multiciclo=True
        )
        mc_glosa = _rodar(
            excel, args.novo.resolve(), ("Glosa", 10000),
            mapas=mapas_mc, multiciclo=True,
        )
        mc_pago = _rodar(
            excel, args.novo.resolve(), ("Valor pago", 98000),
            mapas=mapas_mc, multiciclo=True,
        )
    finally:
        excel.Quit()
        del excel
        pythoncom.CoUninitialize()

    # ---------------- CENARIO A: legado identico -----------------------
    _tabela("A. CHECKPOINT sem ajuste", base_check)
    _tabela("A. NOVO sem ajuste", base_novo)
    difs = {
        chave: (base_check[chave], base_novo.get(chave))
        for chave in base_check
        if not _iguais(base_check[chave], base_novo.get(chave))
    }
    if difs:
        falhas.append(f"CENARIO A: divergencias no legado -> {difs}")
    else:
        print(
            "\n[OK] CENARIO A: checkpoint == novo, celula a celula, sem ajuste."
        )

    # ---------------- CENARIOS B/C/D -----------------------------------
    remanescente = base_novo["MEMORIA_RESULTADOS!D35"]
    for nome, valores, esperado in (
        ("B (valor pago 90.000)", pago, (90000.0, 10000.0, 97200.0, 7200.0)),
        ("C (glosa 10.000)", glosa, (90000.0, 10000.0, 97200.0, 7200.0)),
        ("D (valor pago 0)", zero, (0.0, 100000.0, 0.0, 0.0)),
    ):
        _tabela(nome, valores)
        considerado, glosa_esp, atualizado, retro = esperado
        checagens = {
            "itens_Consumidos!AB3": considerado,
            "itens_Consumidos!AC3": glosa_esp,
            "itens_Consumidos!AE3": atualizado,
            "itens_Consumidos!AF3": retro,
            "MEMORIA_RESULTADOS!F20": atualizado,
            "MEMORIA_RESULTADOS!D11": retro,
            # remanescente fisico intocado
            "MEMORIA_RESULTADOS!C33": base_novo["MEMORIA_RESULTADOS!C33"],
            "MEMORIA_RESULTADOS!D33": base_novo["MEMORIA_RESULTADOS!D33"],
            "MEMORIA_RESULTADOS!D35": remanescente,
            # quantidades intocadas
            "itens_Consumidos!O2": base_novo["itens_Consumidos!O2"],
            "itens_Consumidos!V2": base_novo["itens_Consumidos!V2"],
        }
        for chave, esperada in checagens.items():
            if not _iguais(valores.get(chave), esperada):
                falhas.append(
                    f"CENARIO {nome}: {chave} = {_fmt(valores.get(chave))}, "
                    f"esperado {_fmt(esperada)}"
                )
        vta_esperado = round(atualizado + (remanescente or 0.0), 2)
        if not _iguais(valores.get("MEMORIA_RESULTADOS!B26"), vta_esperado):
            falhas.append(
                f"CENARIO {nome}: VTA (B26) = "
                f"{_fmt(valores.get('MEMORIA_RESULTADOS!B26'))}, "
                f"esperado {_fmt(vta_esperado)}"
            )

    # ---------------- CENARIOS E/F: fail-closed ------------------------
    for nome, valores in (
        ("E (valor pago 105.000)", pago_maior),
        ("F (glosa 110.000)", glosa_maior),
    ):
        _tabela(nome, valores)
        status = str(valores.get("itens_Consumidos!AG3") or "")
        if not status.startswith("REVISAR"):
            falhas.append(f"CENARIO {nome}: AG3 = {status!r}, esperado REVISAR:*")
        for chave in (
            "MEMORIA_RESULTADOS!F20", "MEMORIA_RESULTADOS!B26",
            "itens_Consumidos!AB3", "itens_Consumidos!AC3",
        ):
            if valores.get(chave) not in (None, ""):
                falhas.append(
                    f"CENARIO {nome}: {chave} deveria ficar vazio "
                    f"(fail-closed), veio {_fmt(valores.get(chave))}"
                )

    # ---------------- CENARIO G: multiciclo ----------------------------
    # C1 formalizado 8% (fora da apuracao) + C2 em apuracao 5%:
    #   F  = 1,134   D = 1,05   F/D = 1,08
    #   F - 1   = 0,134   <-- reaplicaria o C1 ja formalizado
    #   F - F/D = 0,054   <-- correto, so o reajuste novo
    _tabela("G. MULTICICLO — checkpoint sem ajuste", mc_check)
    _tabela("G. MULTICICLO — novo sem ajuste", mc_base)
    difs_mc = {
        chave: (mc_check[chave], mc_base.get(chave))
        for chave in mc_check
        if not _iguais(mc_check[chave], mc_base.get(chave))
    }
    if difs_mc:
        falhas.append(f"CENARIO G: legado multiciclo divergiu -> {difs_mc}")
    else:
        print("\n[OK] CENARIO G: checkpoint == novo no multiciclo, sem ajuste.")

    # O cenario so prova algo se os dois fatores forem mesmo diferentes.
    if not _iguais(mc_base.get("itens_Consumidos!Y4"), 108000.0):
        falhas.append(
            "CENARIO G: valor calculado deveria estar na base vigente anterior "
            f"(108.000,00); veio {_fmt(mc_base.get('itens_Consumidos!Y4'))}"
        )
    if not _iguais(mc_base.get("itens_Consumidos!AD4"), 1.05):
        falhas.append(
            "CENARIO G: fator aplicado deveria ser o NOVO do ciclo (1,05); "
            f"veio {_fmt(mc_base.get('itens_Consumidos!AD4'))}"
        )

    remanescente_mc = mc_base["MEMORIA_RESULTADOS!D35"]
    for nome, valores in (
        ("G-glosa (10.000)", mc_glosa),
        ("G-valor pago (98.000)", mc_pago),
    ):
        _tabela(f"G. MULTICICLO — {nome}", valores)
        checagens = {
            "itens_Consumidos!Y4": 108000.0,   # base vigente anterior
            "itens_Consumidos!AB4": 98000.0,   # valor pago considerado
            "itens_Consumidos!AC4": 10000.0,   # glosa
            "itens_Consumidos!AD4": 1.05,      # SO o reajuste novo
            "itens_Consumidos!AE4": 102900.0,  # 98.000 x 1,05
            "itens_Consumidos!AF4": 4900.0,    # 98.000 x 0,05
            # Convergencia XLS: o bloco e a memoria dizem o MESMO retroativo.
            "MEMORIA_RESULTADOS!D12": 4900.0,
            "MEMORIA_RESULTADOS!D15": 4900.0,
            "MEMORIA_RESULTADOS!T74": 4900.0,
            "RESULTADOS!D22": 4900.0,
            "MEMORIA_RESULTADOS!F20": 102900.0,
            "RESULTADOS!B83": 102900.0,
            # remanescente e quantidades intactos
            "MEMORIA_RESULTADOS!C33": mc_base["MEMORIA_RESULTADOS!C33"],
            "MEMORIA_RESULTADOS!D33": mc_base["MEMORIA_RESULTADOS!D33"],
            "MEMORIA_RESULTADOS!D35": remanescente_mc,
        }
        for chave, esperada in checagens.items():
            if not _iguais(valores.get(chave), esperada):
                falhas.append(
                    f"CENARIO {nome}: {chave} = {_fmt(valores.get(chave))}, "
                    f"esperado {_fmt(esperada)}"
                )
        # O erro que este gate existe para impedir: 98.000 x 1,134 = 111.132,
        # retroativo 13.132 — reaplicaria os 8% de C1.
        if _iguais(valores.get("itens_Consumidos!AE4"), 111132.0):
            falhas.append(
                f"CENARIO {nome}: REAPLICOU O FATOR ACUMULADO sobre o valor pago"
            )
        vta_mc = round(102900.0 + (remanescente_mc or 0.0), 2)
        for chave in ("MEMORIA_RESULTADOS!B26", "RESULTADOS!B86"):
            if not _iguais(valores.get(chave), vta_mc):
                falhas.append(
                    f"CENARIO {nome}: {chave} = {_fmt(valores.get(chave))}, "
                    f"esperado {_fmt(vta_mc)}"
                )

    print("\n" + "=" * 70)
    if falhas:
        print("FALHAS:")
        for falha in falhas:
            print("  -", falha)
        raise SystemExit(1)
    print("TODOS OS CENARIOS PASSARAM EM EXCEL REAL (apos reabertura).")


if __name__ == "__main__":
    main()
