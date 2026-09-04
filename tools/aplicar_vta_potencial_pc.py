# -*- coding: utf-8 -*-
"""VTA-POT-1: incorpora o retroativo POTENCIAL ao VTA no metodo PC.

Regra de negocio (somente metodo Pedidos de Compra):

    VTA = C0 executado + ciclos encerrados + remanescente atualizado
        + retroativo POTENCIAL dos MESMOS PCs (criterio prudencial)

A parcela potencial permanece SEMPRE identificada como POTENCIAL: ela nao vira
retroativo reconhecido, valor a pagar nem PC pago. Financeiro e Itens Consumidos
ficam byte a byte identicos.

O aplicador usa Excel COM para preservar recursos x14 que o openpyxl remove.
Nao altera ``itens_PC!A:L``, nao move celulas, nao repontua nenhum named range
existente e nao toca nas ancoras EXECUTADO_APURADO (B83), AJUSTES_DEVIDOS (B84)
nem CONFERENCIA_FORMACAO_VTA (B87).

Uso: <python> tools/aplicar_vta_potencial_pc.py [caminho.xlsx]
"""
from __future__ import annotations

import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

XL_LEFT, XL_RIGHT, XL_CENTER = -4131, -4152, -4108
XL_VCENTER = -4108
XL_CONTINUOUS, XL_THIN = 1, 2
XL_NONE = -4142
XL_EXPRESSION = 2

AZUL_ESCURO = "1F4E78"
AZUL_MUITO_CLARO = "F2F7FC"
BRANCO = "FFFFFF"
CINZA_BORDA = "BFBFBF"
TEXTO_PADRAO = "1F3864"

# VTA-POT-1 — constante central de apresentacao da parcela potencial.
# Amarelo-palha muito claro: distinguivel, nunca alarmista. Sem vermelho e sem
# amarelo/laranja forte. O mesmo par e reproduzido na web e nos documentos.
POTENCIAL_BG = "FFF4CC"
POTENCIAL_TEXTO = "7F6000"

MOEDA_LOCAL = '"R$" #.##0,00;-"R$" #.##0,00;"R$" 0,00;"—"'
MOEDA_INVARIANTE = '"R$" #,##0.00;-"R$" #,##0.00;"R$" 0.00;"—"'

METODO = 'MEMORIA_RESULTADOS!$B$4'
E_PC = f'{METODO}="PCs"'
# Formatacao condicional nao aceita referencia direta a outra aba nesta versao
# do Excel (E_INVALIDARG). O named range existente METODO_RETROATIVO aponta
# para a MESMA celula (MEMORIA_RESULTADOS!$B$4) e e aceito.
# Alem disso, FormatConditions.Add so aceita FUNCOES no idioma da instalacao
# (AND falha no Excel pt-BR). A conjuncao vai pelo produto de condicoes, que
# independe de locale e nao usa separador de argumentos.
E_PC_CF = 'METODO_RETROATIVO="PCs"'


def _cf_e(*condicoes: str) -> str:
    """Conjuncao de condicoes sem funcao nem separador (independe de locale)."""
    return "=" + "*".join(f"({c})" for c in condicoes)

# Espelho exato da regra Python (_motor_composicao_vta._composicao_vta_pc):
# potencial dos PCs que JA compoem a execucao do VTA — quadro itens_PC!S12:S16,
# ja filtrado pela data de corte, restrito aos ciclos C1..C(vigente-1).
# Construcao identica a T22 (execucao dos mesmos ciclos), so que sobre S.
F_T39 = (
    '=IF($T$20="",0,ROUND(SUMPRODUCT('
    '(ROW(itens_PC!$S$12:$S$16)-12>=1)*'
    '(ROW(itens_PC!$S$12:$S$16)-12<$T$20)*'
    '(itens_PC!$S$12:$S$16)),2))'
)
F_T40 = (
    '=IF(OR($T$24=0,$T$20="",$T$26>0,AND(NOT(itens_PC!$P$3>0),$T$27>0)),'
    '"",ROUND($T$21+$T$22+$T$23,2))'
)
F_T25 = (
    '=IF(OR($T$24=0,$T$20="",$T$26>0,AND(NOT(itens_PC!$P$3>0),$T$27>0)),'
    '"CALCULO MANUAL REQUERIDO",ROUND($T$21+$T$22+$T$23+$T$39,2))'
)

# RESULTADOS!84 — a parcela "(+)" do quadro passa a ser dependente do metodo.
# Financeiro mantem "Ajustes ainda devidos" (B21). PCs passa a exibir a parcela
# POTENCIAL. Itens segue em 0. B87 (conferencia) ja soma N($B$84): fecha sozinho.
F_A84 = (
    f'=IF({E_PC},"(+) Retroativo potencial — POTENCIAL",'
    '"(+) Ajustes ainda devidos")'
)
F_B84 = (
    f'=IF({METODO}="Financeiro",'
    'IF(MEMORIA_RESULTADOS!$B$21="","",MEMORIA_RESULTADOS!$B$21),'
    f'IF({E_PC},'
    'IF(MEMORIA_RESULTADOS!$T$25="CALCULO MANUAL REQUERIDO","",'
    'MEMORIA_RESULTADOS!$T$39),'
    f'IF({METODO}="Itens",0,"")))'
)
F_C84 = (
    f'=IF({METODO}="Financeiro",'
    '"Reajuste ja reconhecido e ainda nao contido no valor pago.",'
    f'IF({E_PC},'
    '"Retroativo POTENCIAL dos PCs ainda em analise pela area gestora, '
    'incorporado ao VTA por criterio prudencial. Nao e retroativo reconhecido '
    'a pagar.",'
    f'IF({METODO}="Itens",'
    '"Nao aplicavel: o reajuste ja esta dentro da execucao atualizada.","")))'
)

# Bloco 10 — demonstracao explicita exigida pela regra de negocio.
F_A89 = (
    f'=IF({E_PC},"10. COMPOSIÇÃO PRUDENCIAL DO VTA — PARCELA POTENCIAL","")'
)
F_B90 = f'=IF({E_PC},IF(MEMORIA_RESULTADOS!$T$40="","",MEMORIA_RESULTADOS!$T$40),"")'
F_B91 = (
    f'=IF({E_PC},'
    'IF(MEMORIA_RESULTADOS!$T$25="CALCULO MANUAL REQUERIDO","",'
    'MEMORIA_RESULTADOS!$T$39),"")'
)
F_B92 = f'=IF({E_PC},IF(VTA_FINAL="","",VTA_FINAL),"")'
F_C92 = (
    f'=IF({E_PC},IF(OR($B$90="",$B$92=""),"Aguardando base para conferir.",'
    'IF(ABS(ROUND($B$92-($B$90+N($B$91)),2))<=MEMORIA_RESULTADOS!$D$4,'
    '"VTA = valor sem a parcela potencial + parcela potencial",'
    '"VERIFICAR DIFERENÇA")),"")'
)
F_A93 = (
    f'=IF(AND({E_PC},N($B$91)<>0),'
    '"O VTA inclui R$ "&TEXT($B$91,"#.##0,00")&" de retroativo potencial, '
    'incorporado por critério prudencial. Essa parcela permanece sujeita à '
    'confirmação pela área gestora e não representa, nesta data, retroativo '
    'reconhecido a pagar.","")'
)


def _bgr(hex_rgb: str) -> int:
    return int(hex_rgb[4:6] + hex_rgb[2:4] + hex_rgb[0:2], 16)


def _fechar(wb) -> None:
    try:
        wb.Close(SaveChanges=False)
    except Exception:
        pass


def _moeda(rng) -> None:
    try:
        rng.NumberFormatLocal = MOEDA_LOCAL
    except Exception:
        rng.NumberFormat = MOEDA_INVARIANTE


def _bordas(rng) -> None:
    for indice in (7, 8, 9, 10, 11, 12):
        borda = rng.Borders(indice)
        borda.LineStyle = XL_CONTINUOUS
        borda.Weight = XL_THIN
        borda.Color = _bgr(CINZA_BORDA)


def _remover_cf_propria(rng, formula: str) -> None:
    """Remove APENAS a regra deste aplicador (idempotencia).

    NUNCA usar ``FormatConditions.Delete()`` na faixa inteira: as abas oficiais
    carregam regras x14 homologadas que seriam destruidas, e o proprio Excel
    passa a recusar novos ``Add`` depois disso.
    """
    condicoes = rng.FormatConditions
    for indice in range(condicoes.Count, 0, -1):
        try:
            atual = condicoes.Item(indice)
            if str(atual.Formula1 or "").replace(" ", "") == formula.replace(" ", ""):
                atual.Delete()
        except Exception:
            continue


def _cf(rng, formula: str, fundo: str, texto: str | None = None) -> None:
    """Formatacao condicional: pinta so quando a condicao vale.

    Usada para que Financeiro e Itens Consumidos nao herdem nenhum destaque
    (nem o bloco 10, que so existe no metodo PC).
    """
    _remover_cf_propria(rng, formula)
    # A assinatura por palavra-chave falha no Excel pt-BR via COM; a posicional
    # (Type, Operator, Formula1) e a unica aceita para xlExpression.
    cond = rng.FormatConditions.Add(XL_EXPRESSION, 0, formula)
    cond.Interior.Color = _bgr(fundo)
    if texto:
        cond.Font.Color = _bgr(texto)


def _aplicar_memoria(mem) -> None:
    """Auxiliares T39/T40 e a nova formula do VTA-PC (T25)."""
    mem.Range("S39").Value = (
        "Retroativo potencial incorporado ao VTA (PCs ate o corte, ciclos "
        "anteriores ao vigente)"
    )
    mem.Range("T39").Formula = F_T39
    mem.Range("S40").Value = "VTA-PC antes da parcela potencial"
    mem.Range("T40").Formula = F_T40
    mem.Range("T25").Formula = F_T25
    _moeda(mem.Range("T39:T40"))


def _aplicar_itens_pc(ws) -> None:
    """Destaque suave da coluna RETROATIVO POTENCIAL nos dois quadros.

    Somente formatacao: nenhuma formula, largura, merge ou ancora e tocada.
    A:L permanece integralmente preservada (verificado antes e depois).
    """
    principal_antes = ws.Range("A1:L5001").Formula

    # Cabecalhos: destaque permanente e discreto.
    for endereco in ("S2", "S11"):
        rng = ws.Range(endereco)
        rng.Interior.Color = _bgr(POTENCIAL_BG)
        rng.Font.Color = _bgr(POTENCIAL_TEXTO)

    # Valores: so pinta quando ha potencial (zero permanece neutro, leitura
    # melhor). Regra por celula, sem tocar na linha inteira do PC.
    _cf(ws.Range("S3:S9"), "=N($S3)<>0", POTENCIAL_BG, POTENCIAL_TEXTO)
    _cf(ws.Range("S12:S18"), "=N($S12)<>0", POTENCIAL_BG, POTENCIAL_TEXTO)

    if ws.Range("A1:L5001").Formula != principal_antes:
        raise RuntimeError("A faixa operacional itens_PC!A:L foi alterada.")


def _titulo(ws, endereco: str, formula: str, tamanho: int = 11) -> None:
    rng = ws.Range(endereco)
    if not rng.MergeCells:
        rng.Merge()
    ws.Range(endereco.split(":")[0]).Formula = formula
    rng.Font.Color = _bgr(BRANCO)
    rng.Font.Bold = True
    rng.Font.Size = tamanho
    rng.HorizontalAlignment = XL_LEFT
    rng.VerticalAlignment = XL_VCENTER


def _aplicar_resultados(ws) -> None:
    # --- Quadro do VTA (bloco 9): a linha "(+)" vira a parcela do metodo. ---
    ws.Range("A84").Formula = F_A84
    ws.Range("B84").Formula = F_B84
    ws.Range("C84").Formula = F_C84
    # Destaque suave somente quando ha parcela potencial de fato.
    _cf(
        ws.Range("A84:C84"),
        _cf_e(E_PC_CF, "N($B$84)<>0"),
        POTENCIAL_BG,
        POTENCIAL_TEXTO,
    )
    ws.Range("A81").Value = (
        "Para conferir: VTA oficial = execução apurada + parcela adicional do "
        "método + saldo remanescente atualizado."
    )

    # --- Bloco 10: demonstracao explicita da composicao prudencial (PC). ---
    for linha in (89, 93):
        rng = ws.Range(f"A{linha}:H{linha}")
        if not rng.MergeCells:
            rng.Merge()
    _titulo(ws, "A89:H89", F_A89, tamanho=12)
    _cf(ws.Range("A89:H89"), f"={E_PC_CF}", AZUL_ESCURO, BRANCO)
    ws.Rows("89:89").RowHeight = 24

    rotulos = {
        90: ("VTA antes da parcela potencial",
             "Execução apurada + saldo remanescente atualizado, sem a parcela potencial."),
        91: ("(+) Retroativo potencial — POTENCIAL",
             "Sujeito à confirmação pela área gestora. Não é retroativo reconhecido a pagar."),
        92: ("(=) VALOR TOTAL ATUALIZADO — VTA", ""),
    }
    for linha, (rotulo, explicacao) in rotulos.items():
        ws.Range(f"A{linha}").Formula = f'=IF({E_PC},"{rotulo}","")'
        if explicacao:
            ws.Range(f"C{linha}").Formula = f'=IF({E_PC},"{explicacao}","")'
    ws.Range("B90").Formula = F_B90
    ws.Range("B91").Formula = F_B91
    ws.Range("B92").Formula = F_B92
    ws.Range("C92").Formula = F_C92
    _moeda(ws.Range("B90:B92"))
    ws.Range("B90:B92").HorizontalAlignment = XL_RIGHT
    ws.Range("A92:C92").Font.Bold = True
    # Fundo neutro do bloco; a linha POTENCIAL recebe o amarelo-palha por
    # regra propria, aplicada depois (prevalece na ordem das condicoes).
    _cf(ws.Range("A90:C92"), f"={E_PC_CF}", AZUL_MUITO_CLARO, TEXTO_PADRAO)
    _cf(ws.Range("A91:C91"), f"={E_PC_CF}", POTENCIAL_BG, POTENCIAL_TEXTO)
    _bordas(ws.Range("A90:C92"))

    ws.Range("A93").Formula = F_A93
    ws.Range("A93:H93").WrapText = True
    ws.Range("A93:H93").VerticalAlignment = XL_VCENTER
    ws.Range("A93:H93").Font.Size = 9
    _cf(ws.Range("A93:H93"), _cf_e(E_PC_CF, "N($B$91)<>0"),
        POTENCIAL_BG, POTENCIAL_TEXTO)
    ws.Rows("93:93").RowHeight = 30


def _aplicar_nomes(wb) -> None:
    nomes = {
        "RETROATIVO_POTENCIAL_VTA": "=MEMORIA_RESULTADOS!$T$39",
        "VTA_SEM_POTENCIAL": "=MEMORIA_RESULTADOS!$T$40",
    }
    existentes = {n.Name for n in wb.Names}
    for nome, refere in nomes.items():
        if nome in existentes:
            wb.Names(nome).RefersTo = refere
        else:
            wb.Names.Add(Name=nome, RefersTo=refere)


def aplicar(caminho: Path) -> None:
    import win32com.client as com

    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        # Regras x14 homologadas das abas oficiais: contadas antes e conferidas
        # depois. O aplicador so ACRESCENTA regras — nunca remove as existentes.
        cf_antes = {
            nome: wb.Worksheets(nome).Cells.FormatConditions.Count
            for nome in ("RESULTADOS", "itens_PC", "MEMORIA_RESULTADOS")
        }
        _aplicar_memoria(wb.Worksheets("MEMORIA_RESULTADOS"))
        _aplicar_itens_pc(wb.Worksheets("itens_PC"))
        _aplicar_resultados(wb.Worksheets("RESULTADOS"))
        _aplicar_nomes(wb)
        for nome, antes in cf_antes.items():
            depois = wb.Worksheets(nome).Cells.FormatConditions.Count
            if depois < antes:
                raise RuntimeError(
                    f"{nome}: formatacao condicional perdida "
                    f"({antes} -> {depois})."
                )
        excel.CalculateFullRebuild()
        wb.Worksheets("CONTROLE").Activate()
        wb.Save()
        _fechar(wb)
        wb = None
    finally:
        if wb is not None:
            _fechar(wb)
        excel.Quit()


def main() -> int:
    caminho = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else TEMPLATE
    if not caminho.exists():
        print(f"ERRO: arquivo não encontrado: {caminho}")
        return 1
    aplicar(caminho)
    print(f"OK: VTA-POT-1 aplicado em {caminho}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
