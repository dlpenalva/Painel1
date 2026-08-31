# -*- coding: utf-8 -*-
"""RETROATIVO_POTENCIAL_PC (PR 2, FASE 1-4).

Grandeza EXCLUSIVAMENTE informativa: expoe no XLS o mesmo numero que a web ja
publica como `retroativo_potencial`, sem entrar em nenhuma soma oficial.

Cadeia canonica que este name reproduz:

    itens_PC!J (DELTA_POTENCIAL, por PC)
      -> leitor: col_delta = _col(mapa, "DELTA_POTENCIAL", "RETROATIVO POTENCIAL")
      -> registro["delta_potencial"]
      -> _totais_canonicos_pc -> blocos["ate_o_corte"]["delta_potencial"]
      -> _resultado_consolidado -> consolidado["retroativo_potencial"]
      -> pages/03_Valor_Global.py e pages/12_Adequacao_Orcamentaria.py

Os outros dois filtros do Python sao INERTES na grade oficial e por isso nao
aparecem na formula: `itens_PC` nao possui coluna ENTRA_NO_CALCULO (o default
e "Sim") e DESCARTADO_DUPLICIDADE so e atribuido pela via fiscal
(STATUS_PAGAMENTO_PC / VALOR_EFETIVAMENTE_PAGO), colunas ausentes da grade.
"""
from __future__ import annotations

import datetime as dt
import re
import io
from pathlib import Path

import pytest
from openpyxl import load_workbook

from _leitor_masterfile_v10 import ler_masterfile_v10

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

NOME = "RETROATIVO_POTENCIAL_PC"
DESTINO = "MEMORIA_RESULTADOS!$T$38"
FORMULA = (
    '=ROUND(SUMIFS(itens_PC!$J$2:$J$5001,'
    'itens_PC!$B$2:$B$5001,"<="&$T$31),2)'
)

# Os 14 names publicados pelo PR #135 precisam sobreviver intactos.
NOMES_PR1 = {
    "EXECUTADO_APURADO": "RESULTADOS!$B$83",
    "AJUSTES_DEVIDOS": "RESULTADOS!$B$84",
    "CONFERENCIA_FORMACAO_VTA": "RESULTADOS!$B$87",
    "PC_TOTAL_CADASTRADO": "MEMORIA_RESULTADOS!$T$33",
    "PC_TOTAL_ATE_CORTE": "MEMORIA_RESULTADOS!$T$34",
    "PC_TOTAL_COM_EFEITO": "MEMORIA_RESULTADOS!$T$36",
    "PC_TOTAL_SEM_EFEITO": "MEMORIA_RESULTADOS!$T$37",
    "AUDITORIA_SITUACAO_ATUAL_CONTRATO": "MEMORIA_RESULTADOS!$W$50",
    "AUDITORIA_ULTIMA_REFERENCIA_ABERTURA": "MEMORIA_RESULTADOS!$W$48",
    "AUDITORIA_COMPARATIVO_INTEGRAL": "comparativo_VTA!$B$208",
    "AUDITORIA_DIFERENCA_REFERENCIAS": "MEMORIA_RESULTADOS!$W$51",
    "AUDITORIA_SITUACAO_ATUAL_STATUS": "RESULTADOS!$H$10",
    "AUDITORIA_ABERTURA_STATUS": "RESULTADOS!$H$11",
    "AUDITORIA_CONFERENCIA_STATUS": "MEMORIA_RESULTADOS!$W$52",
}


@pytest.fixture(scope="module")
def wb():
    return load_workbook(TEMPLATE)


# --------------------------------------------------------------- FASE 1
def test_name_publicado_no_destino_aprovado(wb):
    assert NOME in wb.defined_names
    assert wb.defined_names[NOME].value == DESTINO


def test_formula_de_t38_e_a_definicao_canonica(wb):
    assert wb["MEMORIA_RESULTADOS"]["T38"].value == FORMULA


def test_rotulo_de_apoio_em_s38(wb):
    rotulo = wb["MEMORIA_RESULTADOS"]["S38"].value
    assert isinstance(rotulo, str) and "potencial" in rotulo.lower()


def test_d45_permanece_com_a_semantica_anterior(wb):
    """D45 NAO pode ser alterada: continua somando a coluna inteira."""
    assert wb["MEMORIA_RESULTADOS"]["D45"].value == (
        "=ROUND(SUM(itens_PC!$J$2:$J$5001),2)"
    )


def test_names_do_pr1_sobrevivem_e_t35_segue_sem_nome(wb):
    for nome, destino in NOMES_PR1.items():
        assert wb.defined_names[nome].value == destino, nome
    assert not any("$T$35" in str(d.value) for d in wb.defined_names.values())
    assert not [n for n, d in wb.defined_names.items() if "[" in str(d.value)]


def test_vta_atualizacao_cheia_continua_fora_da_cadeia_oficial(wb):
    assert "VTA_ATUALIZACAO_CHEIA" in wb.defined_names


# --------------------------------------------------------------- FASE 3
# A camada humana do PR 2 vive a partir da linha 90 de RESULTADOS. Abaixo
# disso esta o motor tecnico, que nao pode enxergar o potencial.
PRIMEIRA_LINHA_APRESENTACAO = 90


def test_potencial_so_e_citado_pela_camada_de_apresentacao(wb):
    """O name pode ser EXIBIDO, nunca somado.

    Citacao permitida apenas em RESULTADOS a partir da linha 90 (a camada
    humana). Qualquer citacao no motor tecnico (linhas 1-87) ou em outra aba
    significaria que o potencial entrou numa cadeia de calculo.
    """
    indevidas = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cel in row:
                v = cel.value
                if not (isinstance(v, str) and v.startswith("=") and NOME in v):
                    continue
                apresentacao = (
                    ws.title == "RESULTADOS"
                    and cel.row >= PRIMEIRA_LINHA_APRESENTACAO
                )
                if not apresentacao:
                    indevidas.append(f"{ws.title}!{cel.coordinate}")
    assert indevidas == [], f"{NOME} virou insumo de calculo em {indevidas}"


def test_t38_nao_e_citada_por_nenhuma_formula(wb):
    """Ninguem pode somar T38: e saida terminal, nao insumo."""
    citam = []
    ws_mem = wb["MEMORIA_RESULTADOS"]
    for row in ws_mem.iter_rows():
        for cel in row:
            v = cel.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue
            if cel.coordinate == "T38":
                continue
            if "$T$38" in v or "T38" in v.replace("$", ""):
                citam.append(cel.coordinate)
    assert citam == [], f"T38 consumida por {citam}"


def test_celulas_canonicas_do_vta_e_do_retro_intactas(wb):
    """As formulas que produzem VTA e retroativo nao foram tocadas."""
    ws = wb["RESULTADOS"]
    assert ws["C5"].value == '=IF(VTA_FINAL="","",VTA_FINAL)'
    assert ws["D5"].value == "=$D$22"
    assert ws["D22"].value == '=IFERROR(RETRO_OFICIAL,"")'
    assert ws["B86"].value == '=IF(VTA_FINAL="","",VTA_FINAL)'
    mem = wb["MEMORIA_RESULTADOS"]
    assert mem["T31"].value == (
        '=IF(ISNUMBER(CONTROLE!$B$3),CONTROLE!$B$3,DATE(9999,12,31))'
    )


def test_ajustes_manuais_c43_g50_intactos(wb):
    ws = wb["RESULTADOS"]
    assert [ws[f"A{r}"].value for r in range(43, 51)] == [
        "Retroativo manual oficial", "Ajuste do VTA", "VTA manual substitutivo",
        "Complemento histórico", "Complemento histórico",
        "Complemento histórico", "Complemento histórico", "Complemento histórico",
    ]
    validacoes = {(dv.type, str(dv.sqref)) for dv in ws.data_validations.dataValidation}
    assert ("list", "G43:G50") in validacoes
    assert ("decimal", "C46:C50") in validacoes


# --------------------------------------------------------------- FASE 2 e 4
#
# T38 e um SUMIFS. O contrato que ele encerra e:
#
#     soma de itens_PC!J onde itens_PC!B <= T31 (data de corte)
#         ==
#     leitor -> totais_canonicos.ate_o_corte.delta_potencial
#
# As fixtures escrevem VALORES LITERAIS em itens_PC (inclusive na coluna J),
# porque um workbook montado por openpyxl nao carrega cache do Excel: sem
# literal, `J` seria formula sem valor e os dois lados leriam coisas
# diferentes por construcao, nao por divergencia semantica.

COL = {"A": 1, "B": 2, "C": 3, "D": 4, "F": 6, "G": 7, "J": 10, "L": 12}


def _monta(pcs, data_corte=dt.date(2024, 12, 31)):
    """Workbook derivado do template com PCs escritos como literais."""
    wb = load_workbook(TEMPLATE)
    wb["CONTROLE"]["B1"] = "PCs"
    wb["CONTROLE"]["B2"] = "C2"
    wb["CONTROLE"]["B3"] = data_corte
    ws = wb["itens_PC"]
    for i, pc in enumerate(pcs):
        r = 2 + i
        ws.cell(r, COL["A"]).value = pc["numero"]
        ws.cell(r, COL["B"]).value = pc["data"]
        ws.cell(r, COL["C"]).value = pc.get("ciclo", "C1")
        ws.cell(r, COL["D"]).value = pc["valor"]
        ws.cell(r, COL["F"]).value = pc.get("atualizado", pc["valor"])
        ws.cell(r, COL["G"]).value = pc["pago"]
        ws.cell(r, COL["J"]).value = pc["delta"]
        ws.cell(r, COL["L"]).value = pc.get("efeito", "Sim")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _t38(raw, data_corte):
    """Avalia a semantica do SUMIFS de T38 sobre o workbook."""
    ws = load_workbook(io.BytesIO(raw), data_only=True)["itens_PC"]
    limite = data_corte or dt.date(9999, 12, 31)
    total = 0.0
    for r in range(2, 5002):
        j = ws.cell(r, COL["J"]).value
        if not isinstance(j, (int, float)) or isinstance(j, bool):
            continue
        b = ws.cell(r, COL["B"]).value
        if isinstance(b, dt.datetime):
            b = b.date()
        if isinstance(b, dt.date) and b <= limite:
            total += j
    return round(total, 2)


def _web(raw):
    res = ler_masterfile_v10(io.BytesIO(raw))
    totais = (res.get("itens_pc") or {}).get("totais_canonicos") or {}
    bloco = totais.get("ate_o_corte") or {}
    return round(float(bloco.get("delta_potencial") or 0.0), 2)


CORTE = dt.date(2024, 12, 31)

CENARIOS_POTENCIAL = {
    # nome                        PCs                                            esperado
    "pc_com_potencial": ([
        dict(numero="PC-1", data=dt.date(2024, 3, 1), valor=100000.0,
             atualizado=107436.8, pago="Nao", delta=7436.8),
    ], 7436.8),
    "pc_sem_potencial": ([
        dict(numero="PC-1", data=dt.date(2024, 3, 1), valor=100000.0,
             atualizado=100000.0, pago="Nao", delta=0.0),
    ], 0.0),
    "pc_validada_paga": ([
        # PC pago: itens_PC!J e 0 por definicao — reconhecido, nao potencial.
        dict(numero="PC-1", data=dt.date(2024, 3, 1), valor=100000.0,
             atualizado=107436.8, pago="Sim", delta=0.0),
    ], 0.0),
    "pc_em_analise": ([
        dict(numero="PC-1", data=dt.date(2024, 4, 1), valor=80000.0,
             atualizado=83500.0, pago="", delta=3500.0),
    ], 3500.0),
    "pc_sem_efeito_financeiro": ([
        dict(numero="PC-1", data=dt.date(2024, 5, 1), valor=60000.0,
             atualizado=61200.0, pago="Nao", delta=0.0, efeito="Nao"),
    ], 0.0),
    "multiplas_pcs": ([
        dict(numero="PC-1", data=dt.date(2024, 2, 1), valor=50000.0,
             atualizado=51000.0, pago="Nao", delta=1000.0),
        dict(numero="PC-2", data=dt.date(2024, 6, 1), valor=70000.0,
             atualizado=72500.0, pago="Nao", delta=2500.0),
        dict(numero="PC-3", data=dt.date(2024, 9, 1), valor=30000.0,
             atualizado=30000.0, pago="Sim", delta=0.0),
    ], 3500.0),
    "pc_posterior_ao_corte_nao_conta": ([
        dict(numero="PC-1", data=dt.date(2024, 3, 1), valor=50000.0,
             atualizado=51000.0, pago="Nao", delta=1000.0),
        # depois do corte: entra em posterior_ao_corte, nunca no potencial
        dict(numero="PC-2", data=dt.date(2025, 3, 1), valor=90000.0,
             atualizado=99000.0, pago="Nao", delta=9000.0),
    ], 1000.0),
}


@pytest.mark.parametrize("nome", sorted(CENARIOS_POTENCIAL))
def test_paridade_t38_com_a_grandeza_canonica_da_web(nome):
    pcs, esperado = CENARIOS_POTENCIAL[nome]
    raw = _monta(pcs, CORTE)
    xls, web = _t38(raw, CORTE), _web(raw)
    assert xls == esperado, f"{nome}: T38 deu {xls}, esperado {esperado}"
    assert xls == web, f"{nome}: XLS {xls} != web canonico {web}"


def test_pc_posterior_ao_corte_discrimina_t38_de_d45():
    """Prova de que a formula filtrada e necessaria: D45 erraria aqui."""
    pcs, esperado = CENARIOS_POTENCIAL["pc_posterior_ao_corte_nao_conta"]
    raw = _monta(pcs, CORTE)
    ws = load_workbook(io.BytesIO(raw), data_only=True)["itens_PC"]
    d45 = round(sum(
        ws.cell(r, COL["J"]).value
        for r in range(2, 5002)
        if isinstance(ws.cell(r, COL["J"]).value, (int, float))
    ), 2)
    assert _t38(raw, CORTE) == esperado == _web(raw)
    assert d45 == 10000.0
    assert d45 != esperado, "cenario nao discrimina; revise a fixture"


def test_potencial_nao_contamina_retro_oficial_nem_vta(monkeypatch):
    """Mesmos dados com e sem potencial produzem o MESMO retro/VTA oficiais."""
    base = [dict(numero="PC-1", data=dt.date(2024, 3, 1), valor=100000.0,
                 atualizado=100000.0, pago="Sim", delta=0.0)]
    com_potencial = base + [
        dict(numero="PC-2", data=dt.date(2024, 7, 1), valor=80000.0,
             atualizado=88000.0, pago="Nao", delta=8000.0),
    ]
    r_sem = ler_masterfile_v10(io.BytesIO(_monta(base, CORTE)))
    r_com = ler_masterfile_v10(io.BytesIO(_monta(com_potencial, CORTE)))

    def oficiais(res):
        rx = res.get("resultados_xls") or {}
        valores = rx.get("valores") or {}
        return valores.get("RETRO_OFICIAL"), valores.get("VTA_FINAL")

    assert oficiais(r_sem) == oficiais(r_com), (
        "o potencial alterou RETRO_OFICIAL/VTA_FINAL"
    )
    pot_sem = ((r_sem.get("itens_pc") or {}).get("totais_canonicos") or {}
               ).get("ate_o_corte", {}).get("delta_potencial")
    pot_com = ((r_com.get("itens_pc") or {}).get("totais_canonicos") or {}
               ).get("ate_o_corte", {}).get("delta_potencial")
    assert round(float(pot_sem or 0), 2) == 0.0
    assert round(float(pot_com or 0), 2) == 8000.0


# ------------------------------------------------- FASE 5: novo leiaute
#
# A aba passa a ter duas camadas: o motor tecnico (linhas 1-87, oculto) e a
# apresentacao humana (linhas 90-166). Os testes abaixo protegem a fronteira.

PRIMEIRA_LINHA_TECNICA = 1
ULTIMA_LINHA_TECNICA = 87


@pytest.fixture(scope="module")
def ws(wb):
    return wb["RESULTADOS"]


def test_camada_tecnica_esta_oculta_e_a_humana_visivel(ws):
    ocultas = [r for r in range(1, 90) if ws.row_dimensions[r].hidden]
    assert len(ocultas) == 89, "a camada tecnica precisa sair da leitura"
    visiveis = [r for r in range(90, 167) if ws.row_dimensions[r].hidden]
    assert visiveis == [], f"linhas da apresentacao ocultas: {visiveis}"
    # A coluna I hospeda so as guias de formatacao condicional.
    assert ws.column_dimensions["I"].hidden


def test_impressao_em_duas_paginas_paisagem(ws):
    assert ws.print_area == "'RESULTADOS'!$A$90:$H$166"
    assert ws.page_setup.orientation == "landscape"
    assert int(ws.page_setup.scale) == 75
    # Quebra explicita entre a pagina 1 e a pagina 2.
    assert [b.id for b in ws.row_breaks.brk] == [116]


def test_pinos_tecnicos_seguem_nas_mesmas_coordenadas(ws):
    assert ws["B3"].value.startswith("=IF(OR($H$8=")
    assert ws["B22"].value == '=IF(COUNT(B16:B20)=0,"",ROUND(SUM(B16:B20),2))'
    assert ws["B37"].value == (
        '=IFERROR(INDEX($C$26:$C$30,MATCH(UPPER(CONTROLE!$B$2),$A$26:$A$30,0)),"")'
    )
    assert ws["B38"].value == '=IF(OR(B36="",B37=""),"",ROUND(B37-B36,2))'


def test_hero_values_leem_a_cadeia_canonica(ws):
    assert ws["A95"].value == '=IF(VTA_FINAL="","",VTA_FINAL)'
    assert ws["E95"].value == "=$D$22"          # RETRO_OFICIAL via D22
    assert ws["A94"].value == "VTA OFICIAL"
    assert ws["E94"].value == "RETROATIVO TOTAL A PAGAR"


def test_formacao_do_vta_usa_os_names_do_pr1(ws):
    assert "EXECUTADO_APURADO" in ws["B121"].value
    assert "AJUSTES_DEVIDOS" in ws["B122"].value
    assert ws["B123"].value == '=IF($B$85="","",$B$85)'
    assert "VTA_FINAL" in ws["B124"].value
    assert "CONFERENCIA_FORMACAO_VTA" in ws["B125"].value
    assert "DE ACORDO" in ws["C125"].value


def test_terminologia_nova_esta_aplicada(ws):
    assert ws["A117"].value == (
        "VALOR TOTAL ATUALIZADO DO CONTRATO — METODOLOGIA E FORMAÇÃO"
    )
    assert ws["D128"].value == "Diferença"          # nunca "Delta"
    assert ws["A144"].value == "SITUAÇÃO ATUAL DO CONTRATO"
    assert ws["A150"].value == (
        "REFERÊNCIAS PARA CONFERÊNCIA — NÃO SÃO O VTA OFICIAL"
    )
    assert ws["C151"].value == "FINALIDADE"          # nunca "Situação"
    assert ws["C152"].value == "REFERÊNCIA ATUAL"
    assert ws["C153"].value == "REFERÊNCIA DE ABERTURA"
    assert ws["C154"].value == "COMPARATIVO TEÓRICO"
    assert ws["A156"].value == "CONFERÊNCIA ENTRE REFERÊNCIAS DO CONTRATO"
    assert ws["A157"].value == "DIFERENÇA"


def test_coluna_do_efeito_financeiro_e_a_competencia_de_inicio(ws):
    """parametros!H guarda QUANDO o efeito comeca, nao SE ha efeito nem quanto.

    O rotulo antigo ("Efeito financeiro") era ambiguo: podia ser lido como
    Sim/Nao ou como valor monetario.
    """
    assert ws["F106"].value == "INÍCIO DO EFEITO FINANCEIRO"
    assert ws["F106"].alignment.wrap_text is True
    for i in range(5):
        r, p = 107 + i, 2 + i
        assert ws.cell(r, 6).value == (
            f'=IF(parametros!$H${p}="","",parametros!$H${p})'
        )
        assert ws.cell(r, 6).number_format == "mm/yyyy"


def test_jargao_antigo_nao_aparece_na_camada_humana(ws):
    proibidos = ("Delta", "posição física", "posicao fisica", "AUDITORIA INTERNA")
    achados = []
    for r in range(90, 167):
        for c in range(1, 9):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            for termo in proibidos:
                if termo.lower() in v.lower():
                    achados.append((ws.cell(r, c).coordinate, termo))
    assert achados == [], f"jargao antigo visivel: {achados}"


def test_aviso_de_revisao_esta_presente_sem_workflow(ws):
    aviso = ws["A103"].value
    assert isinstance(aviso, str)
    assert aviso.startswith("IMPORTANTE")
    assert "não substitui a revisão do responsável" in aviso
    # Informativo puro: nada de aceite, trava ou registro de concordancia.
    validacoes = {str(dv.sqref) for dv in ws.data_validations.dataValidation}
    assert not any("103" in s for s in validacoes)


def test_bloco_de_potencial_e_condicional_e_ambar(ws, wb):
    for ref in ("A100", "D100", "A101"):
        assert NOME in ws[ref].value
        assert 'MEMORIA_RESULTADOS!$B$4="PCs"' in ws[ref].value
    assert "EM ANÁLISE" in ws["A100"].value
    assert "Não compõe os valores oficiais" in ws["A101"].value
    # A guia da formatacao condicional concentra a logica (locale-proof).
    assert ws["I100"].value.startswith("=IF(AND(")


def test_conclusao_reflete_o_potencial_em_analise(ws):
    """§12: com potencial em analise nao se declara conclusao total."""
    assert "COM VALOR POTENCIAL EM ANÁLISE" in ws["A113"].value
    assert "retroativo potencial" in ws["A114"].value
    assert "não compõe o VTA nem o retroativo oficial" in ws["A114"].value


def test_fator_aparece_com_seis_casas_e_so_na_pagina_2(ws):
    for r in range(138, 143):
        assert ws.cell(r, 3).number_format == "0.000000"
    # §9: a pagina 1 nao exibe fatores.
    for r in range(107, 112):
        assert ws.cell(r, 3).number_format != "0.000000"


# ------------------------------------------- FASE 5.1: acabamento visual

def test_ausencia_de_potencial_nao_desenha_travessao(ws):
    """Sem potencial material, D100 fica visualmente vazio.

    O monetario homologado usa "—" na 4a secao (texto), o que fazia uma
    celula condicional que devolve "" desenhar um travessao ambar orfao.
    """
    assert ws["D100"].number_format.endswith(";")
    assert '"—"' not in ws["D100"].number_format


def test_conferencia_da_execucao_nao_deixa_linhas_de_travessao(ws):
    """Fora do Financeiro as cinco linhas ficam visualmente vazias."""
    for r in range(161, 166):
        for c in (2, 3, 4):
            fmt = ws.cell(r, c).number_format
            assert fmt.endswith(";"), f"{ws.cell(r, c).coordinate}: {fmt}"
            assert '"—"' not in fmt
    assert "Não aplicável ao método selecionado" in ws["F160"].value


def test_cabecalhos_longos_quebram_em_vez_de_truncar(ws):
    assert ws["C160"].value == "Execução estimada pelo quantitativo"
    for c in range(1, 6):
        assert ws.cell(160, c).alignment.wrap_text is True
    assert ws.row_dimensions[160].height >= 24


def test_camada_humana_nao_expoe_nada_tecnico(ws):
    """Nem literais nem formulas podem citar aba, endereco ou name tecnico."""
    proibido = re.compile(
        r"(MEMORIA(_RESULTADOS)?!|CONTROLE!|parametros!|comparativo_VTA!"
        r"|itens_PC!|historico_VU!|financeiro!|posicao_\w+!"
        r"|VTA_FINAL|RETRO_OFICIAL|apuracao|POSICAO|posi[cç][aã]o f[ií]sica)",
        re.I,
    )
    achados = []
    for r in range(90, 167):
        for c in range(1, 9):
            v = ws.cell(r, c).value
            if not isinstance(v, str):
                continue
            # Numa formula so o texto entre aspas chega ao usuario.
            alvos = re.findall(r'"([^"]*)"', v) if v.startswith("=") else [v]
            for alvo in alvos:
                if proibido.search(alvo):
                    achados.append((ws.cell(r, c).coordinate, alvo[:60]))
    assert achados == [], f"vazamento tecnico visivel: {achados}"


def test_coluna_origem_usa_texto_proprio(ws):
    assert ws["C121"].value == "Apuração do método selecionado"
    assert ws["C122"].value == "Ajustes reconhecidos na apuração"
    assert ws["C123"].value == "Saldo contratual atualizado"
    assert ws["C124"].value == "Resultado oficial da apuração"
