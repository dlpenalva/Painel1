"""Adaptacao em memoria dos derivados cacheados de Coletas anteriores ao 11.0.

PROBLEMA
--------
Uma Coleta das linhagens PRE_11_L1/PRE_11_L2 pode ter sido recalculada pelo
Excel com o percentual BRUTO do ciclo (0,05123816) em vez do oficial de duas
casas (0,0512). Todo valor DERIVADO gravado em cache carrega essa precisao:
fatores, valores unitarios, valor atualizado e delta financeiro, consumo,
remanescente, retroativo por PC. As ENTRADAS do fiscal (datas, valores pagos,
quantidades, PCs, itens) estao corretas.

O QUE ESTE MODULO FAZ
---------------------
Recompoe, somente no workbook de valores mantido em memoria, os derivados que
dependem da cadeia de fatores, aplicando a regra vigente (percentual fechado
em duas casas). O arquivo fisico nunca e alterado, nenhuma entrada e
tocada e os valores originais do XLS ficam registrados na auditoria.

PROTECAO CONTRA FORMULA DIFERENTE
---------------------------------
Cada bloco e calculado DUAS vezes: (1) com a cadeia BRUTA do proprio arquivo,
para provar que a formula reproduzida aqui e a que o Excel executou naquele
arquivo — o resultado precisa coincidir com o cache; (2) com a cadeia
OFICIAL, que e o que passa a valer. Bloco que nao reproduz o cache nao e
adaptado (fail-closed por bloco): o valor do XLS permanece e a auditoria
declara o bloco como nao reproduzivel.

O QUE NAO ENTRA
---------------
Abas de RESULTADOS (RESULTADOS, MEMORIA_RESULTADOS, comparativo_VTA) sao a
"visao do XLS": continuam como o Excel gravou, para que a reconciliacao
XLS x Python exponha a diferenca em vez de esconde-la.
"""

from __future__ import annotations

from datetime import date, datetime, timedelta
from decimal import ROUND_HALF_UP, Decimal
from typing import Any

from _compatibilidade_coleta import (
    LINHAGEM_PRE_11_L1,
    LINHAGEM_PRE_11_L2,
    detectar_linhagem_coleta,
)
from _reajuste_utils import (
    fator_oficial,
    tem_precisao_superior_a_oficial,
)

ATRIBUTO_AUDITORIA = "_cl8us_compat_valores"

_CICLOS = ("C0", "C1", "C2", "C3", "C4")
_TOLERANCIA_DINHEIRO = 0.005
_TOLERANCIA_FATOR = 1e-9
_AMOSTRA_AUDITORIA = 40

# Colunas por ciclo (1-based) das abas que carregam um bloco por ciclo.
_VU_HISTORICO = {0: 3, 1: 4, 2: 5, 3: 6, 4: 7}               # historico_VU C..G
_QTD_POSICAO = {1: 11, 2: 15, 3: 19, 4: 23}                  # posicao_contratual K,O,S,W
_VALOR_REMANESCENTE = {1: 6, 2: 8, 3: 10, 4: 12}             # itens_Remanesc F,H,J,L
_QTD_EXECUTADA = {1: 13, 2: 15, 3: 17, 4: 19}                # itens_Remanesc M,O,Q,S
_VALOR_EXECUTADO = {1: 14, 2: 16, 3: 18, 4: 20}              # itens_Remanesc N,P,R,T
_RC_VU = {0: 2, 1: 5, 2: 8, 3: 11, 4: 14}                    # itens_RC B,E,H,K,N
_RC_QTD = {0: 3, 1: 6, 2: 9, 3: 12, 4: 15}                   # itens_RC C,F,I,L,O
_RC_TOTAL = {0: 4, 1: 7, 2: 10, 3: 13, 4: 16}                # itens_RC D,G,J,M,P
_CONS_QTD = {0: 5, 1: 7, 2: 9, 3: 11, 4: 13}                 # itens_Consumidos E,G,I,K,M
_CONS_VALOR = {0: 6, 1: 8, 2: 10, 3: 12, 4: 14}              # itens_Consumidos F,H,J,L,N


# --------------------------------------------------------------------------- #
# Utilitarios com a semantica das celulas do Excel.
# --------------------------------------------------------------------------- #
def _num(valor: Any) -> float | None:
    if isinstance(valor, bool) or valor is None:
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    return None


def _vazio(valor: Any) -> bool:
    return valor is None or (isinstance(valor, str) and valor.strip() == "")


def _r2(valor: float) -> float:
    """ROUND(x, 2) do Excel: meio para cima sobre 15 digitos significativos."""
    decimal = Decimal(format(valor, ".15g"))
    return float(decimal.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def _r4(valor: float) -> float:
    decimal = Decimal(format(valor, ".15g"))
    return float(decimal.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP))


def _igual(antigo: Any, novo: Any, tolerancia: float) -> bool:
    if _vazio(antigo) and _vazio(novo):
        return True
    a, n = _num(antigo), _num(novo)
    if a is not None and n is not None:
        return abs(a - n) <= tolerancia
    return antigo == novo


def _como_data(valor: Any) -> date | None:
    """Data de uma celula (datetime ou serial do Excel); None se nao for data."""
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    numero = _num(valor)
    if numero is not None and 20000 < numero < 80000:
        return (datetime(1899, 12, 30) + timedelta(days=numero)).date()
    return None


def _celula(ws, linha: int, coluna: int) -> Any:
    return ws.cell(linha, coluna).value


def _soma(valores) -> float:
    return sum(v for v in (_num(x) for x in valores) if v is not None)


# --------------------------------------------------------------------------- #
# Cadeias de fatores (mesma regra de parametros!F e parametros!A11:E15).
# --------------------------------------------------------------------------- #
class _Fatores:
    """Cadeia historica integral (parametros!F) e cadeia da apuracao (E11:E15)."""

    __slots__ = ("historico", "apuracao", "computar", "percentuais")

    def __init__(self, percentuais: list[float | None], computar: list[bool]):
        self.percentuais = percentuais
        self.computar = computar
        historico: list[float | None] = [1.0]
        for indice in range(1, 5):
            anterior = historico[-1]
            percentual = percentuais[indice]
            if percentual is None or anterior is None:
                historico.append(None)
            else:
                historico.append(anterior * (1 + percentual))
        self.historico = historico
        apuracao: list[float | None] = [1.0]
        for indice in range(1, 5):
            anterior = apuracao[-1]
            if computar[indice]:
                percentual = percentuais[indice]
                if percentual is None or anterior is None:
                    apuracao.append(None)
                else:
                    apuracao.append(anterior * (1 + percentual))
            else:
                apuracao.append(anterior)
        self.apuracao = apuracao


def _percentuais_brutos(wb) -> tuple[list[float | None], list[bool]]:
    ws = wb["parametros"]
    percentuais: list[float | None] = []
    computar: list[bool] = []
    for linha in range(2, 7):
        percentuais.append(_num(_celula(ws, linha, 5)))
        computar.append(str(_celula(ws, linha, 1) or "").strip() == "Sim")
    return percentuais, computar


def _percentuais_oficiais(brutos: list[float | None]) -> list[float | None]:
    """Fecha cada percentual (decimal) em duas casas (regra vigente)."""
    return [None if p is None else fator_oficial(p) - 1.0 for p in brutos]


# --------------------------------------------------------------------------- #
# Calculo de cada bloco. Cada funcao devolve {(aba, linha, coluna): valor}.
# --------------------------------------------------------------------------- #
def _calc_parametros(f: _Fatores) -> dict:
    saida: dict = {}
    for indice in range(5):
        valor = f.historico[indice]
        saida[("parametros", 2 + indice, 6)] = valor if valor is not None else ""
    # memoria do fator aplicavel a apuracao (A11:E15, TOTAL em D16/E16/F16)
    for indice in range(1, 5):
        linha = 11 + indice
        if f.computar[indice]:
            percentual = f.percentuais[indice]
            saida[("parametros", linha, 3)] = percentual if percentual is not None else ""
            saida[("parametros", linha, 4)] = (
                1 + percentual if percentual is not None else ""
            )
        else:
            saida[("parametros", linha, 3)] = ""
            saida[("parametros", linha, 4)] = ""
        valor = f.apuracao[indice]
        saida[("parametros", linha, 5)] = valor if valor is not None else ""
    final = f.apuracao[4]
    saida[("parametros", 16, 4)] = final if final is not None else ""
    saida[("parametros", 16, 5)] = final if final is not None else ""
    saida[("parametros", 16, 6)] = final - 1 if final is not None else ""
    return saida


def _linhas_com_item(ws, primeira: int, ultima: int, coluna: int = 1):
    for linha in range(primeira, ultima + 1):
        if not _vazio(_celula(ws, linha, coluna)):
            yield linha


def _calc_financeiro(wb, f: _Fatores, l2: bool) -> dict:
    ws = wb["financeiro"]
    saida: dict = {}
    indice_ciclo = {c.lower(): i for i, c in enumerate(_CICLOS)}
    for linha in range(2, min(ws.max_row, 400) + 1):
        ciclo = str(_celula(ws, linha, 2) or "").strip().lower()
        pago = _num(_celula(ws, linha, 3))
        efeito = _celula(ws, linha, 7)
        if ciclo not in indice_ciclo:
            continue
        numero = indice_ciclo[ciclo]
        if numero == 0:
            fator: Any = 1.0
        elif f.computar[numero]:
            fator = f.apuracao[numero] if f.apuracao[numero] is not None else ""
        else:
            fator = ""
        saida[("financeiro", linha, 4)] = fator
        if pago is None:
            continue
        fator_num = _num(fator)
        if l2:
            atualizado: Any = _r2(pago * fator_num) if fator_num is not None else ""
            if atualizado == "":
                delta: Any = ""
            else:
                delta = 0.0 if str(efeito or "") != "Sim" else _r2(atualizado - pago)
        else:
            if _vazio(efeito):
                atualizado, delta = "", ""
            elif efeito == "Sim":
                atualizado = _r2(pago * fator_num) if fator_num is not None else ""
                delta = _r2(atualizado - pago) if atualizado != "" else ""
            else:
                atualizado, delta = pago, 0.0
        saida[("financeiro", linha, 5)] = atualizado
        saida[("financeiro", linha, 6)] = delta
    for linha in range(2, min(ws.max_row, 400) + 1):
        if str(_celula(ws, linha, 2) or "").strip().upper() == "TOTAL":
            for coluna in (5, 6):
                saida[("financeiro", linha, coluna)] = _soma(
                    saida.get(("financeiro", r, coluna)) for r in range(2, linha)
                )
            break
    return saida


def _calc_historico(wb, f: _Fatores, l2: bool) -> tuple[dict, dict]:
    """VU por item e ciclo (historico_VU C..G) na cadeia recebida."""
    hist = wb["historico_VU"]
    rem = wb["itens_Remanesc"]
    pos = wb["posicao_contratual"]
    valores: dict[int, dict[int, Any]] = {}
    saida: dict = {}
    for linha in _linhas_com_item(hist, 2, 210):
        vu0_rem = _num(_celula(rem, linha, 3))
        nascimento = _num(_celula(pos, linha, 25))         # posicao_contratual!Y
        vu0 = _num(_celula(hist, linha, 3))
        por_ciclo: dict[int, Any] = {0: vu0}
        for ciclo in range(1, 5):
            if l2:
                fator = f.historico[ciclo]
                por_ciclo[ciclo] = (
                    _r2(vu0 * fator) if vu0 is not None and fator is not None else ""
                )
            elif nascimento is None or nascimento > ciclo:
                por_ciclo[ciclo] = ""
            elif nascimento == ciclo:
                por_ciclo[ciclo] = vu0_rem if vu0_rem is not None else ""
            else:
                fator_ciclo = f.historico[ciclo]
                fator_origem = f.historico[int(nascimento)]
                por_ciclo[ciclo] = (
                    _r2(vu0_rem * fator_ciclo / fator_origem)
                    if None not in (vu0_rem, fator_ciclo, fator_origem)
                    else ""
                )
        for ciclo in range(1, 5):
            saida[("historico_VU", linha, _VU_HISTORICO[ciclo])] = por_ciclo[ciclo]
        valores[linha] = por_ciclo
        final = f.historico[4]
        saida[("historico_VU", linha, 8)] = (
            _r4(final - 1)
            if vu0 not in (None, 0.0) and final is not None
            else ""
        )
    for indice in range(5):
        valor = f.historico[indice]
        saida[("historico_VU", 2 + indice, 12)] = valor if valor is not None else ""
    return valores, saida


def _calc_remanescentes(wb, f: _Fatores, l2: bool, vu: dict) -> dict:
    rem = wb["itens_Remanesc"]
    pos = wb["posicao_contratual"]
    saida: dict = {}
    for indice in range(5):                               # Z2:Z6 (fator acumulado)
        valor = f.historico[indice]
        saida[("itens_Remanesc", 2 + indice, 26)] = valor if valor is not None else ""
    for linha in _linhas_com_item(rem, 2, 210):
        vu_item = vu.get(linha) or {}
        vu0 = _num(_celula(rem, linha, 3))
        nascimento_data = _num(_celula(pos, linha, 38))   # posicao_contratual!AL
        for ciclo in range(1, 5):
            qtd_pos = _celula(pos, linha, _QTD_POSICAO[ciclo])
            exec_qtd = _celula(rem, linha, _QTD_EXECUTADA[ciclo])
            if l2:
                fator = f.historico[ciclo]
                valor_rem: Any = (
                    ""
                    if _vazio(qtd_pos) or vu0 is None or fator is None
                    else _r2(_num(qtd_pos) * vu0 * fator)
                )
                valor_exec: Any = (
                    ""
                    if _vazio(exec_qtd) or vu0 is None or fator is None
                    else _r2(_num(exec_qtd) * vu0 * fator)
                )
            else:
                vu_ciclo = _num(vu_item.get(ciclo))
                fora = nascimento_data is not None and nascimento_data > ciclo
                valor_rem = (
                    ""
                    if fora or _vazio(qtd_pos) or vu_ciclo is None
                    else _r2(_num(qtd_pos) * vu_ciclo)
                )
                valor_exec = (
                    ""
                    if _vazio(exec_qtd) or vu_ciclo is None
                    else _r2(_num(exec_qtd) * vu_ciclo)
                )
            saida[("itens_Remanesc", linha, _VALOR_REMANESCENTE[ciclo])] = valor_rem
            saida[("itens_Remanesc", linha, _VALOR_EXECUTADO[ciclo])] = valor_exec
        exec_c0 = _celula(rem, linha, 28)                 # AB
        if l2:
            fator = f.historico[0]
            valor = (
                "" if _vazio(exec_c0) or vu0 is None or fator is None
                else _r2(_num(exec_c0) * vu0 * fator)
            )
        else:
            vu_c0 = _num(vu_item.get(0))
            valor = "" if _vazio(exec_c0) or vu_c0 is None else _r2(_num(exec_c0) * vu_c0)
        saida[("itens_Remanesc", linha, 29)] = valor      # AC
    itens = list(_linhas_com_item(rem, 2, 210))
    if itens and itens[-1] + 1 <= 210 and _vazio(_celula(rem, itens[-1] + 1, 1)):
        total = itens[-1] + 1
        colunas = (*_VALOR_REMANESCENTE.values(), *_VALOR_EXECUTADO.values(), 29)
        for coluna in colunas:
            if _vazio(_celula(rem, total, coluna)):
                continue                      # o XLS nao gravou total nesta coluna
            valores = [saida.get(("itens_Remanesc", r, coluna)) for r in itens]
            saida[("itens_Remanesc", total, coluna)] = _r2(_soma(valores))
    return saida


def _calc_rc(wb, f: _Fatores, l2: bool, vu: dict, ciclo_em_execucao: dict) -> dict:
    rc = wb["itens_RC"]
    rem = wb["itens_Remanesc"]
    saida: dict = {}
    primeira = 3
    for linha_rc in range(primeira, 204):
        item = _celula(rc, linha_rc, 1)
        origem = linha_rc - 1                             # linha em itens_Remanesc/historico_VU
        if str(item or "").strip().upper() == "TOTAL":
            for ciclo in range(5):
                coluna = _RC_TOTAL[ciclo]
                saida[("itens_RC", linha_rc, coluna)] = _r2(
                    _soma(saida.get(("itens_RC", r, coluna)) for r in range(primeira, linha_rc))
                )
            break
        if _vazio(item):
            continue
        vu_item = vu.get(origem) or {}
        vu0_rem = _num(_celula(rem, origem, 3))
        for ciclo in range(5):
            if l2:
                fator = f.historico[ciclo]
                vu_ciclo: Any = (
                    _r2(vu0_rem * fator)
                    if vu0_rem is not None and fator is not None else ""
                )
            else:
                vu_ciclo = vu_item.get(ciclo, "")
                vu_ciclo = "" if _vazio(vu_ciclo) else vu_ciclo
            saida[("itens_RC", linha_rc, _RC_VU[ciclo])] = vu_ciclo
            qtd = _celula(rc, linha_rc, _RC_QTD[ciclo])
            saida[("itens_RC", linha_rc, _RC_TOTAL[ciclo])] = (
                "" if _vazio(vu_ciclo) or _vazio(qtd)
                else _r2(_num(vu_ciclo) * _num(qtd))
            )
        if ciclo_em_execucao:
            # V, W, X = CICLO_EM_EXECUCAO!E, F, G (linha + 10).
            for coluna_rc, coluna_ciclo in ((22, 5), (23, 6), (24, 7)):
                valor = ciclo_em_execucao.get(
                    ("CICLO_EM_EXECUCAO", linha_rc + 10, coluna_ciclo), ""
                )
                saida[("itens_RC", linha_rc, coluna_rc)] = "" if _vazio(valor) else valor
    return saida


def _calc_consumidos(wb, f: _Fatores) -> dict:
    ws = wb["itens_Consumidos"]
    saida: dict = {}
    for indice in range(5):                               # U2:U6
        valor = f.historico[indice]
        saida[("itens_Consumidos", 2 + indice, 21)] = valor if valor is not None else ""
    for linha in _linhas_com_item(ws, 2, 200):
        valores = []
        vu = _num(_celula(ws, linha, 3))
        for ciclo in range(5):
            qtd = _celula(ws, linha, _CONS_QTD[ciclo])
            fator = f.historico[ciclo]
            valor: Any = (
                "" if _vazio(qtd) or vu is None or fator is None
                else _r2(_num(qtd) * vu * fator)
            )
            saida[("itens_Consumidos", linha, _CONS_VALOR[ciclo])] = valor
            valores.append(valor)
        saida[("itens_Consumidos", linha, 16)] = _soma(valores)        # P
    return saida


def _calc_pc(wb, f: _Fatores, l2: bool) -> dict:
    ws = wb["itens_PC"]
    saida: dict = {}
    limite = min(ws.max_row, 5001)
    colunas_soma = (6, 8, 9, 10)                          # F, H, I, J
    itens_quadro: list[tuple] = []        # (ciclo, data_pc, pago, F, H, I, J)
    somas = {coluna: {c: 0.0 for c in _CICLOS} for coluna in colunas_soma}
    fora = {coluna: 0.0 for coluna in colunas_soma}       # ciclo fora de C0..C4
    for linha in range(2, limite + 1):
        ciclo = _celula(ws, linha, 3)
        valor = _num(_celula(ws, linha, 4))
        pago = _celula(ws, linha, 7)
        if ciclo not in _CICLOS:
            if not _vazio(ciclo):
                for coluna in colunas_soma:
                    numero = _num(_celula(ws, linha, coluna))
                    if numero is not None:
                        fora[coluna] += numero
            continue
        fator = f.apuracao[_CICLOS.index(ciclo)]
        saida[("itens_PC", linha, 5)] = fator if fator is not None else ""
        atualizado: Any = (
            _r2(valor * fator) if valor is not None and fator is not None else ""
        )
        saida[("itens_PC", linha, 6)] = atualizado
        incremento = (
            _r2(atualizado - valor)
            if atualizado != "" and valor is not None else ""
        )
        if l2:
            retro: Any = incremento if pago == "Sim" else 0.0
            em_analise: Any = atualizado if pago == "Nao" else 0.0
            delta: Any = incremento if pago == "Nao" else 0.0
        else:
            efeito = _celula(ws, linha, 12)              # L: EFEITO_FINANCEIRO_PC
            if pago == "Sim":
                if atualizado == "" or valor is None or _vazio(efeito):
                    retro = ""
                else:
                    retro = incremento if efeito == "Sim" else 0.0
                em_analise, delta = 0.0, 0.0
            else:
                retro = 0.0
                em_analise = "" if atualizado == "" or _vazio(efeito) else atualizado
                if atualizado == "" or valor is None or _vazio(efeito):
                    delta = ""
                else:
                    delta = incremento if efeito == "Sim" else 0.0
        saida[("itens_PC", linha, 8)] = retro
        saida[("itens_PC", linha, 9)] = em_analise
        saida[("itens_PC", linha, 10)] = delta
        if not l2:
            # U = VALOR_CONSIDERADO: valor original + retroativo reconhecido.
            saida[("itens_PC", linha, 21)] = (
                "" if valor is None or _vazio(retro) else _r2(valor + _num(retro))
            )
        for coluna, dado in ((6, atualizado), (8, retro), (9, em_analise), (10, delta)):
            numero = _num(dado)
            if numero is not None:
                somas[coluna][ciclo] += numero
        itens_quadro.append((
            ciclo, _como_data(_celula(ws, linha, 2)), pago,
            saida.get(("itens_PC", linha, 21)), retro, em_analise, delta,
        ))
    # Quadro-resumo por ciclo (M:S). O leiaute difere por linhagem: localizamos
    # as linhas pelos rotulos da coluna M (primeiro quadro, linhas 2..10).
    linhas_ciclo: dict[str, int] = {}
    outras = total = None
    for linha in range(2, 11):
        rotulo = str(_celula(ws, linha, 13) or "").strip()
        if rotulo in _CICLOS and rotulo not in linhas_ciclo:
            linhas_ciclo[rotulo] = linha
        elif rotulo.upper().startswith("OUTRAS") and outras is None:
            outras = linha
        elif rotulo.upper() == "TOTAL":
            total = linha
            break
    destino = {6: 16, 8: 17, 9: 18, 10: 19}               # P, Q, R, S
    for coluna, coluna_resumo in destino.items():
        acumulado = 0.0
        for ciclo, linha in linhas_ciclo.items():
            saida[("itens_PC", linha, coluna_resumo)] = somas[coluna][ciclo]
            acumulado += somas[coluna][ciclo]
        if outras is not None:
            saida[("itens_PC", outras, coluna_resumo)] = fora[coluna]
            acumulado += fora[coluna]
        if total is not None:
            saida[("itens_PC", total, coluna_resumo)] = acumulado
    if not l2:
        _calc_pc_quadro_corte(wb, ws, itens_quadro, saida, fora)
    return saida


def _calc_pc_quadro_corte(wb, ws, itens_quadro: list, saida: dict, fora: dict) -> None:
    """Segundo quadro de itens_PC (PCs considerados ate a data de corte).

    Colunas P (valor considerado U), Q (retroativo H), R (em analise I) e S
    (potencial J), filtradas pelo corte de MEMORIA_RESULTADOS!T31. Cada linha
    de ciclo soma por CICLO_PC; a linha "Outras" e o total do quadro menos as
    linhas de ciclo; TOTAL fecha o quadro. N, O e T nao dependem do fator.
    """
    if "MEMORIA_RESULTADOS" not in wb.sheetnames:
        return
    corte_bruto = _celula(wb["MEMORIA_RESULTADOS"], 31, 20)            # T31
    corte = _como_data(corte_bruto)
    if corte is None:
        return                                   # sem corte reconhecivel: nao adaptar
    linhas_ciclo: dict[str, int] = {}
    outras = total = None
    for linha in range(11, 30):
        rotulo = str(_celula(ws, linha, 13) or "").strip()
        if rotulo in _CICLOS and rotulo not in linhas_ciclo:
            linhas_ciclo[rotulo] = linha
        elif rotulo.upper().startswith("OUTRAS") and linhas_ciclo and outras is None:
            outras = linha
        elif rotulo.upper() == "TOTAL" and linhas_ciclo:
            total = linha
            break
    if len(linhas_ciclo) != 5 or total is None:
        return

    def dentro(data_pc):
        return data_pc is not None and data_pc <= corte

    # (coluna do quadro, posicao do valor em itens_quadro, exige G = Sim)
    regras = ((16, 3, True), (17, 4, True), (18, 5, False), (19, 6, False))
    for coluna, posicao, exige_pago in regras:
        por_ciclo = {c: 0.0 for c in _CICLOS}
        total_quadro = 0.0
        for ciclo, data_pc, pago, *valores in itens_quadro:
            if not dentro(data_pc) or (exige_pago and pago != "Sim"):
                continue
            numero = _num(valores[posicao - 3])
            if numero is None:
                continue
            por_ciclo[ciclo] += numero
            total_quadro += numero
        # fora dos ciclos: itens com ciclo nao reconhecido (mesma regra de G)
        acumulado = 0.0
        for ciclo, linha in linhas_ciclo.items():
            saida[("itens_PC", linha, coluna)] = por_ciclo[ciclo]
            acumulado += por_ciclo[ciclo]
        if outras is not None:
            saida[("itens_PC", outras, coluna)] = total_quadro - acumulado
            acumulado = total_quadro
        saida[("itens_PC", total, coluna)] = acumulado


def _calc_aditivos(wb, f: _Fatores, l2: bool, vu: dict) -> dict:
    ws = wb["aditivos"]
    saida: dict = {}
    hist = wb["historico_VU"]
    linhas_item = {
        str(_celula(hist, r, 1)): r for r in _linhas_com_item(hist, 2, 210)
    }
    for linha in _linhas_com_item(ws, 2, 200):
        ciclo = _celula(ws, linha, 3)
        if ciclo not in _CICLOS:
            fator: Any = ""
        else:
            valor = f.historico[_CICLOS.index(ciclo)]
            fator = valor if valor is not None else ""
        saida[("aditivos", linha, 9)] = fator
        delta = _celula(ws, linha, 12)
        vu_original = _celula(ws, linha, 6)
        aplicar = str(_celula(ws, linha, 8) or "").upper() == "SIM"
        if _vazio(delta) or _vazio(vu_original):
            saida[("aditivos", linha, 10)] = ""
            continue
        if l2:
            multiplicador = _num(fator) if aplicar and _num(fator) is not None else 1.0
            saida[("aditivos", linha, 10)] = _r2(
                _num(delta) * _num(vu_original) * multiplicador
            )
        else:
            vu_ciclo = None
            r_item = linhas_item.get(str(_celula(ws, linha, 1)))
            if r_item is not None and ciclo in _CICLOS:
                vu_ciclo = _num((vu.get(r_item) or {}).get(_CICLOS.index(ciclo)))
            base = vu_ciclo if aplicar and vu_ciclo is not None else _num(vu_original)
            saida[("aditivos", linha, 10)] = _r2(_num(delta) * base)
    return saida


def _calc_ciclo_em_execucao(wb, vu: dict) -> dict:
    if "CICLO_EM_EXECUCAO" not in wb.sheetnames:
        return {}
    ws = wb["CICLO_EM_EXECUCAO"]
    vigente = str(_celula(ws, 3, 3) or "").strip().upper()
    if vigente not in _CICLOS:
        return {}
    ciclo = _CICLOS.index(vigente)
    saida: dict = {}
    total = 0.0
    qualquer = False
    for linha in range(13, 212):
        item = _celula(ws, linha, 1)
        if _vazio(item):
            continue
        vu_ciclo = (vu.get(linha - 11) or {}).get(ciclo)
        vu_ciclo = "" if _vazio(vu_ciclo) else vu_ciclo
        saida[("CICLO_EM_EXECUCAO", linha, 5)] = vu_ciclo
        consumida = _celula(ws, linha, 4)
        restante = _celula(ws, linha, 3)
        saida[("CICLO_EM_EXECUCAO", linha, 6)] = (
            "" if _vazio(consumida) or vu_ciclo == "" else _r2(_num(consumida) * _num(vu_ciclo))
        )
        valor_rem: Any = (
            "" if _vazio(restante) or vu_ciclo == "" else _r2(_num(restante) * _num(vu_ciclo))
        )
        saida[("CICLO_EM_EXECUCAO", linha, 7)] = valor_rem
        if _num(valor_rem) is not None:
            total += _num(valor_rem)
        qualquer = True
    if qualquer and not _vazio(_celula(ws, 9, 1)):
        saida[("CICLO_EM_EXECUCAO", 9, 1)] = _r2(total)
    return saida


_BLOCOS = (
    "parametros", "financeiro", "historico_VU", "itens_Remanesc", "itens_RC",
    "itens_Consumidos", "itens_PC", "aditivos", "CICLO_EM_EXECUCAO",
)


def _calcular_tudo(wb, f: _Fatores, l2: bool) -> dict[str, dict]:
    """Todos os blocos, em ordem de dependencia, na cadeia `f`."""
    blocos: dict[str, dict] = {}
    blocos["parametros"] = _calc_parametros(f)
    blocos["financeiro"] = _calc_financeiro(wb, f, l2)
    vu, blocos["historico_VU"] = _calc_historico(wb, f, l2)
    blocos["itens_Remanesc"] = _calc_remanescentes(wb, f, l2, vu)
    blocos["CICLO_EM_EXECUCAO"] = _calc_ciclo_em_execucao(wb, vu)
    blocos["itens_RC"] = _calc_rc(wb, f, l2, vu, blocos["CICLO_EM_EXECUCAO"])
    blocos["itens_Consumidos"] = _calc_consumidos(wb, f)
    blocos["itens_PC"] = _calc_pc(wb, f, l2)
    blocos["aditivos"] = _calc_aditivos(wb, f, l2, vu)
    return blocos


def _tolerancia(aba: str, coluna: int) -> float:
    fator = (
        aba == "parametros"
        or (aba == "historico_VU" and coluna == 12)
        or (aba == "itens_Remanesc" and coluna == 26)
        or (aba == "itens_Consumidos" and coluna == 21)
        or (aba == "financeiro" and coluna == 4)
        or (aba == "itens_PC" and coluna == 5)
        or (aba == "aditivos" and coluna == 9)
    )
    return _TOLERANCIA_FATOR if fator else _TOLERANCIA_DINHEIRO


def _reproduz_cache(wb, bloco: dict) -> tuple[bool, list]:
    """True quando o bloco, calculado na cadeia bruta, coincide com o cache."""
    divergentes = []
    for (aba, linha, coluna), esperado in bloco.items():
        atual = _celula(wb[aba], linha, coluna)
        if not _igual(atual, esperado, _tolerancia(aba, coluna)):
            divergentes.append((aba, linha, coluna, atual, esperado))
            if len(divergentes) >= 5:
                break
    return not divergentes, divergentes


MENSAGEM_L2_ADITIVO = (
    "Esta Coleta é de um modelo anterior ao 11.0 (PRE_11_L2) e possui aditivo em "
    "ciclo de reajuste. Esse modelo não incorporava o efeito do aditivo na "
    "quantidade remanescente ajustada do ciclo; o valor remanescente e o VTA "
    "apurados a partir dela podem estar incompletos. Regere a Coleta no modelo "
    "11.0 e faça novo upload antes da formalização."
)


def restricoes_compatibilidade(wb, deteccao: dict | None = None) -> list[dict[str, str]]:
    """Situacoes em que a linhagem NAO tem a regra vigente e o dado existe no arquivo.

    Nao e precisao de percentual: e regra de modelo. O L2 calcula
    QTD_REM_AJUSTADA_Cn = QTD_REM_BASE_Cn (o aditivo do ciclo nao entra); o 11.0
    calcula BASE + DELTA_Cn. Como o aditivo esta no arquivo mas a posicao dele
    nao, o resultado seria silenciosamente incompleto: fica fail-closed, com o
    motivo tecnico declarado.
    """
    deteccao = deteccao or detectar_linhagem_coleta(wb)
    if deteccao.get("codigo") != LINHAGEM_PRE_11_L2 or "aditivos" not in wb.sheetnames:
        return []
    ws = wb["aditivos"]
    for linha in _linhas_com_item(ws, 2, 200):
        ciclo = _celula(ws, linha, 3)
        delta = _num(_celula(ws, linha, 12))
        if ciclo in _CICLOS[1:] and delta not in (None, 0.0):
            return [{"codigo": "L2_ADITIVO_SEM_EFEITO_NA_POSICAO",
                     "mensagem": MENSAGEM_L2_ADITIVO}]
    return []


def _auditoria_vazia(motivo: str, deteccao: dict | None = None) -> dict[str, Any]:
    return {
        "aplicada": False,
        "motivo": motivo,
        "linhagem": (deteccao or {}).get("codigo"),
        "ciclos_precisao_bruta": [],
        "blocos_adaptados": [],
        "blocos_nao_reproduziveis": {},
        "celulas_alteradas": 0,
        "amostra": [],
    }


def aplicar_compatibilidade_valores(wb, deteccao: dict | None = None) -> dict[str, Any]:
    """Recompoe derivados de uma Coleta PRE-11 com precisao bruta (em memoria).

    Idempotente por workbook. So deve receber ``workbook_valores``
    (data_only=True): o workbook de formulas nunca e alterado.
    """
    existente = getattr(wb, ATRIBUTO_AUDITORIA, None)
    if existente is not None:
        return existente

    def _registrar(auditoria: dict[str, Any]) -> dict[str, Any]:
        setattr(wb, ATRIBUTO_AUDITORIA, auditoria)
        return auditoria

    deteccao = deteccao or detectar_linhagem_coleta(wb)
    codigo = deteccao.get("codigo")
    if codigo not in (LINHAGEM_PRE_11_L1, LINHAGEM_PRE_11_L2):
        return _registrar(_auditoria_vazia("linhagem sem adaptacao de valores", deteccao))
    try:
        percentuais, computar = _percentuais_brutos(wb)
    except Exception:
        return _registrar(_auditoria_vazia("parametros ilegivel", deteccao))
    if any(p is not None and abs(p) > 1 for p in percentuais):
        return _registrar(_auditoria_vazia("percentual fora da unidade decimal", deteccao))
    com_precisao = [
        _CICLOS[i]
        for i, p in enumerate(percentuais)
        if i > 0 and p is not None and tem_precisao_superior_a_oficial(p)
    ]
    if not com_precisao:
        return _registrar(_auditoria_vazia("percentuais ja oficiais; nada a recompor", deteccao))

    l2 = codigo == LINHAGEM_PRE_11_L2
    bloco_bruto = _calcular_tudo(wb, _Fatores(percentuais, computar), l2)
    bloco_oficial = _calcular_tudo(
        wb, _Fatores(_percentuais_oficiais(percentuais), computar), l2
    )

    adaptados: list[str] = []
    nao_reproduziveis: dict[str, list] = {}
    alteracoes: list[tuple] = []
    for nome in _BLOCOS:
        calculo_bruto = bloco_bruto.get(nome) or {}
        if not calculo_bruto:
            continue
        ok, divergentes = _reproduz_cache(wb, calculo_bruto)
        if not ok:
            nao_reproduziveis[nome] = [
                {"aba": aba, "linha": linha, "coluna": coluna,
                 "xls": atual, "recomposto": esperado}
                for (aba, linha, coluna, atual, esperado) in divergentes
            ]
            continue
        ws = wb[nome]
        for (aba, linha, coluna), novo in (bloco_oficial.get(nome) or {}).items():
            antigo = _celula(ws, linha, coluna)
            if _igual(antigo, novo, _tolerancia(aba, coluna) / 50):
                continue
            ws.cell(linha, coluna).value = novo
            alteracoes.append((aba, linha, coluna, antigo, novo))
        adaptados.append(nome)

    return _registrar({
        "aplicada": bool(adaptados),
        "motivo": "precisao de reajuste anterior a regra vigente",
        "linhagem": codigo,
        "ciclos_precisao_bruta": com_precisao,
        "blocos_adaptados": adaptados,
        "blocos_nao_reproduziveis": nao_reproduziveis,
        "celulas_alteradas": len(alteracoes),
        "amostra": [
            {"aba": a, "linha": l, "coluna": c, "xls": antigo, "recomposto": novo}
            for (a, l, c, antigo, novo) in alteracoes[:_AMOSTRA_AUDITORIA]
        ],
    })
