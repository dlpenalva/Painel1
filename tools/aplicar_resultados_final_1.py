# -*- coding: utf-8 -*-
"""RESULTADOS-FINAL-1: acabamento visual da aba RESULTADOS restaurada.

Continuacao direta do RESULTADOS-ROLLBACK-1 (PR #141). A aba voltou a
apresentacao anterior e ficou funcional; o que resta e ACABAMENTO. Esta
frente e conservadora e localizada: nao cria camada nova de apresentacao,
nao volta as linhas 90:166 da UX2, nao move formula nenhuma de lugar e nao
toca em metodologia ou matematica.

O QUE MUDA (somente apresentacao + rotulos)
  1. linhas 41:50 ("5. AJUSTES MANUAIS") ficam OCULTAS. Nada e apagado,
     limpo ou movido: C43:G50 permanece integralmente nas mesmas
     coordenadas, com as mesmas formulas e validacoes.
  2. tabela 6 ganha o titulo "6. TOTAIS E INDICADORES DE CONFERENCIA" e o
     padrao visual das tabelas 1-4 (faixa azul escuro, cabecalho azul
     claro, bordas finas, rotulo a esquerda / valor a direita, moeda
     brasileira, VTA Oficial e Status em negrito). A coluna "Composicao
     auditavel" vira "REFERENCIA PARA CONFERENCIA" e deixa de citar abas,
     celulas e nomes definidos: a ORIGEM DOS DADOS NAO MUDA, so o texto.
  3. tabela 7 vira um bloco explicativo simples (faixa azul escuro no
     titulo, fundo claro, sem excesso de bordas) com a hierarquia
     "COMPOSICAO ADOTADA" / "METODO APLICADO". "acertos ainda devidos"
     passa a "ajustes ainda devidos" — rotulo, nao contrato.
  4. tabela 8 vira tabela secundaria padronizada; a nota inferior fica.
  5. tabela 9 ganha destaque superior as 6-8; "(=) VTA OFICIAL" recebe
     fundo azul escuro com texto branco (NAO verde: verde e reservado a
     semantica de validacao) e a Conferencia cai de hierarquia.
  6. RETROATIVO POTENCIAL — EM ANALISE passa a ser exibido em E22:H22,
     na propria linha do TOTAL do retroativo reconhecido.

FONTE DO RETROATIVO POTENCIAL — NENHUM CALCULO NOVO
  G22 le exclusivamente o nome definido RETROATIVO_POTENCIAL_PC
  (= MEMORIA_RESULTADOS!$T$38), que ja existia e foi preservado pelo
  PR #141. A grandeza nao e recalculada, redefinida nem duplicada aqui.

POR QUE E22:H22 E POR QUE NAO DESLOCAR NADA
  E22:H22 e a unica faixa vizinha ao retroativo reconhecido que esta
  simultaneamente vazia, fora de qualquer merge (E16:H21 termina na 21 e
  E23:H23 comeca na 23), fora de formatacao condicional, sem validacao de
  dados e sem nome definido apontando para ela. Usa-la nao desloca uma
  linha sequer: os pinos tecnicos B3, B22, B36:B38, C43:G50 e B83:B87
  permanecem exatamente onde estavam.

REGRA ZERO CORRUPCAO XLSX: aplicacao por Excel COM — openpyxl remove a
formatacao condicional x14 que a aba usa nas linhas 1:87.

Uso:  python tools/aplicar_resultados_final_1.py [caminho.xlsx]
      (exige pywin32 + Excel real)
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
TEMPLATE = RAIZ / "templates" / "COLETA_REAJUSTE_OFICIAL.xlsx"

# ----------------------------------------------------------------- constantes
XL_LEFT, XL_RIGHT, XL_CENTER = -4131, -4152, -4108
XL_VCENTER = -4108
XL_CONTINUOUS, XL_THIN = 1, 2
XL_EXPRESSION = 2
XL_NONE = -4142
BORDAS_EXTERNAS = (7, 8, 9, 10)      # left, top, bottom, right
BORDAS_INTERNAS = (11, 12)           # inside vertical, inside horizontal

# Paleta ja usada pelas tabelas 1-4 desta mesma aba.
AZUL_ESCURO = "1F4E78"
AZUL_CLARO = "D9EAF7"
AZUL_MUITO_CLARO = "F2F7FC"
BRANCO = "FFFFFF"
CINZA_CLARO = "F2F2F2"
CINZA_BORDA = "BFBFBF"
CINZA_TEXTO = "595959"
AMBAR_FUNDO = "FFF2CC"
AMBAR_TEXTO = "7F6000"
TEXTO_PADRAO = "1F3864"

MOEDA_LOCAL = '"R$" #.##0,00;-"R$" #.##0,00;"R$" 0,00;"—"'
MOEDA_INVARIANTE = '"R$" #,##0.00;-"R$" #,##0.00;"R$" 0.00;"—"'

PRIMEIRA_LINHA_AJUSTES = 41
ULTIMA_LINHA_AJUSTES = 50

# Retroativo potencial: rotulo/observacao em E22:F22, valor em G22:H22.
CELULA_POTENCIAL_ROTULO = "E22"
CELULA_POTENCIAL_VALOR = "G22"
FAIXA_POTENCIAL = "E22:H22"
ALTURA_LINHA_POTENCIAL = 42.0

FORMULA_POTENCIAL_VALOR = '=IF($B$5<>"PCs","",IFERROR(RETROATIVO_POTENCIAL_PC,""))'
FORMULA_POTENCIAL_ROTULO = (
    '=IF($G$22="","","Retroativo potencial — em análise"&CHAR(10)&'
    '"Valor potencial relacionado a itens ainda em análise. '
    'Não integra o retroativo reconhecido enquanto não houver '
    'confirmação.")'
)

TITULO_TABELA_6 = "6. TOTAIS E INDICADORES DE CONFERÊNCIA"
CABECALHO_TABELA_6 = ("Medida", "Valor", "REFERÊNCIA PARA CONFERÊNCIA")

# Rotulos acentuados e referencias em linguagem de negocio. A ORIGEM DOS
# DADOS (as formulas da coluna B) nao e tocada em nenhuma destas linhas.
LINHAS_TABELA_6 = {
    55: ("1. Total cadastrado de pedidos de compra",
         "Inventário integral dos pedidos, sem corte de data."),
    56: ("2. Total considerado até a data de corte",
         "Exclui os pedidos com data posterior à data de corte."),
    57: ("3. Total enquadrado nos ciclos",
         "Pedidos distribuídos por ciclo, já sem os posteriores ao corte."),
    58: ("4. Total com efeito financeiro",
         "Pedidos com efeito financeiro reconhecido até a data de corte."),
    59: ("5. Total sem efeito financeiro",
         "Enquadrado menos o total com efeito: competências anteriores ao "
         "início dos efeitos."),
    60: ("6. Retroativo reconhecido",
         "Retroativo oficial apurado — o mesmo total do quadro 2."),
    61: ("7. Execução física do ciclo atual",
         "Execução informada pela fiscalização quando a posição "
         "física está completa."),
    62: ("8. Remanescente físico atual",
         "Remanescente na abertura do ciclo menos a execução física."),
    63: ("9. VTA oficial",
         "Resultado final do método de apuração selecionado."),
    64: ("10. Referências auditáveis",
         "Posição atual, última abertura e contrato integralmente "
         "reajustado."),
    65: ("11. Reconciliação",
         "Diferença entre as duas formas de cálculo; 0,00 reconcilia."),
    66: ("12. Status",
         "Conclusão global desta aba."),
}
# Linhas monetarias da tabela 6 (64 e texto, 66 e o status textual).
LINHAS_MOEDA_TABELA_6 = (55, 56, 57, 58, 59, 60, 61, 62, 63, 65)
B64_SEM_REFERENCIA_TECNICA = '="Disponíveis na memória de cálculo"'

TEXTO_COMPOSICAO_VTA = (
    "COMPOSIÇÃO ADOTADA — VTA = execução já realizada + "
    "ajustes ainda devidos + saldo remanescente atualizado."
)
FORMULA_METODO_APLICADO = (
    '="MÉTODO APLICADO — "&'
    'IF(MEMORIA_RESULTADOS!$B$4="Financeiro",'
    '"Método Financeiro: execução já realizada = valores '
    'efetivamente desembolsados informados no financeiro.",'
    'IF(MEMORIA_RESULTADOS!$B$4="PCs",'
    '"Método PCs: execução já realizada = valores apurados pelos '
    'Pedidos de Compra conforme as regras do método.",'
    'IF(MEMORIA_RESULTADOS!$B$4="Itens",'
    '"Método Consumido: execução já realizada = quantidades '
    'consumidas x valores unitários aplicáveis.",'
    '"Selecione o método na aba CONTROLE.")))'
)

C85_SEM_REFERENCIA_TECNICA = "Saldo que ainda falta executar, já atualizado."
C86_SEM_REFERENCIA_TECNICA = (
    "Resultado final do método de apuração selecionado."
)


def _fechar(wb, tentativas: int = 10) -> None:
    """Excel recusa Close logo apos um Save pesado (RPC_E_CALL_REJECTED)."""
    for _ in range(tentativas):
        try:
            wb.Close(SaveChanges=False)
            return
        except Exception:
            time.sleep(1.0)


def _bgr(rgb: str) -> int:
    """COM recebe a cor em BGR; o template guarda em RGB."""
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return b * 65536 + g * 256 + r


def _moeda(rng) -> None:
    """pt-BR rejeita alguns codigos invariantes; tenta o local primeiro."""
    try:
        rng.NumberFormatLocal = MOEDA_LOCAL
    except Exception:
        rng.NumberFormat = MOEDA_INVARIANTE


def _faixa_titulo(ws, endereco: str, tamanho: int = 11) -> None:
    rng = ws.Range(endereco)
    rng.Interior.Color = _bgr(AZUL_ESCURO)
    rng.Font.Color = _bgr(BRANCO)
    rng.Font.Bold = True
    rng.Font.Size = tamanho
    rng.HorizontalAlignment = XL_LEFT
    rng.VerticalAlignment = XL_VCENTER


def _cabecalho(ws, endereco: str) -> None:
    rng = ws.Range(endereco)
    rng.Interior.Color = _bgr(AZUL_CLARO)
    rng.Font.Color = _bgr(TEXTO_PADRAO)
    rng.Font.Bold = True
    rng.Font.Size = 10
    rng.VerticalAlignment = XL_VCENTER


def _bordas_finas(ws, endereco: str) -> None:
    rng = ws.Range(endereco)
    for indice in BORDAS_EXTERNAS + BORDAS_INTERNAS:
        borda = rng.Borders(indice)
        borda.LineStyle = XL_CONTINUOUS
        borda.Weight = XL_THIN
        borda.Color = _bgr(CINZA_BORDA)


def _aplicar_tabela_6(ws) -> None:
    ws.Range("A53").Value = TITULO_TABELA_6
    _faixa_titulo(ws, "A53:C53")
    for coluna, rotulo in zip("ABC", CABECALHO_TABELA_6):
        ws.Range(coluna + "54").Value = rotulo
    _cabecalho(ws, "A54:C54")
    ws.Range("B54").HorizontalAlignment = XL_RIGHT

    for linha, (rotulo, referencia) in LINHAS_TABELA_6.items():
        ws.Range("A%d" % linha).Value = rotulo
        ws.Range("C%d" % linha).Value = referencia
    ws.Range("B64").Formula = B64_SEM_REFERENCIA_TECNICA

    corpo = ws.Range("A55:C66")
    corpo.Font.Size = 10
    corpo.Font.Color = _bgr(TEXTO_PADRAO)
    corpo.Interior.Color = _bgr(BRANCO)
    ws.Range("A55:A66").HorizontalAlignment = XL_LEFT
    ws.Range("B55:B66").HorizontalAlignment = XL_RIGHT
    ws.Range("C55:C66").HorizontalAlignment = XL_LEFT
    ws.Range("C55:C66").Font.Size = 9
    ws.Range("C55:C66").Font.Color = _bgr(CINZA_TEXTO)
    ws.Range("C55:C66").WrapText = True

    for linha in LINHAS_MOEDA_TABELA_6:
        _moeda(ws.Range("B%d" % linha))
    # B64 e B66 sao textuais: nunca recebem mascara monetaria.
    ws.Range("B63").Font.Bold = True      # VTA Oficial
    ws.Range("B66").Font.Bold = True      # Status
    _bordas_finas(ws, "A54:C66")


def _aplicar_tabela_7(ws) -> None:
    _faixa_titulo(ws, "A68:C68")
    ws.Range("A69").Value = TEXTO_COMPOSICAO_VTA
    ws.Range("A70").Formula = FORMULA_METODO_APLICADO
    bloco = ws.Range("A69:H70")
    bloco.Interior.Color = _bgr(AZUL_MUITO_CLARO)
    bloco.Font.Size = 10
    bloco.Font.Color = _bgr(TEXTO_PADRAO)
    bloco.HorizontalAlignment = XL_LEFT
    bloco.VerticalAlignment = XL_VCENTER
    # Bloco explicativo: sem grade interna, so respiro.
    for indice in BORDAS_EXTERNAS + BORDAS_INTERNAS:
        bloco.Borders(indice).LineStyle = XL_NONE


def _aplicar_tabela_8(ws) -> None:
    _faixa_titulo(ws, "A71:E71")
    _cabecalho(ws, "A72:E72")
    ws.Range("A72:E72").WrapText = True
    corpo = ws.Range("A73:E77")
    corpo.Font.Size = 10
    corpo.Font.Color = _bgr(TEXTO_PADRAO)
    corpo.Interior.Color = _bgr(BRANCO)
    ws.Range("A73:A77").HorizontalAlignment = XL_CENTER
    _moeda(ws.Range("B73:D77"))
    ws.Range("B73:D77").HorizontalAlignment = XL_RIGHT
    conferencia = ws.Range("E73:E77")
    conferencia.HorizontalAlignment = XL_CENTER
    conferencia.Font.Size = 9
    conferencia.Font.Color = _bgr(CINZA_TEXTO)
    conferencia.WrapText = True
    _bordas_finas(ws, "A72:E77")
    nota = ws.Range("A78")
    nota.Font.Size = 9
    nota.Font.Italic = True
    nota.Font.Color = _bgr(CINZA_TEXTO)


def _aplicar_tabela_9(ws) -> None:
    # Destaque superior as tabelas 6-8: mesma faixa, um corpo maior.
    _faixa_titulo(ws, "A79:C79", tamanho=12)
    notas = ws.Range("A80:C81")
    notas.Interior.Color = _bgr(AZUL_MUITO_CLARO)
    notas.Font.Size = 10
    notas.Font.Color = _bgr(TEXTO_PADRAO)
    notas.HorizontalAlignment = XL_LEFT
    _cabecalho(ws, "A82:C82")
    ws.Range("B82").HorizontalAlignment = XL_RIGHT

    ws.Range("C85").Value = C85_SEM_REFERENCIA_TECNICA
    ws.Range("C86").Value = C86_SEM_REFERENCIA_TECNICA

    corpo = ws.Range("A83:C85")
    corpo.Font.Size = 10
    corpo.Font.Color = _bgr(TEXTO_PADRAO)
    corpo.Interior.Color = _bgr(BRANCO)
    corpo.Font.Bold = False
    ws.Range("A83:A85").HorizontalAlignment = XL_LEFT
    _moeda(ws.Range("B83:B87"))
    ws.Range("B83:B87").HorizontalAlignment = XL_RIGHT
    ws.Range("C83:C87").HorizontalAlignment = XL_LEFT
    ws.Range("C83:C87").Font.Size = 9
    ws.Range("C83:C87").WrapText = True

    # (=) VTA OFICIAL: maior hierarquia da aba, em azul escuro.
    total = ws.Range("A86:C86")
    total.Interior.Color = _bgr(AZUL_ESCURO)
    total.Font.Color = _bgr(BRANCO)
    total.Font.Bold = True
    total.Font.Size = 11
    ws.Range("C86").Font.Size = 9

    # Conferencia: hierarquia inferior, cinza claro.
    conferencia = ws.Range("A87:C87")
    conferencia.Interior.Color = _bgr(CINZA_CLARO)
    conferencia.Font.Color = _bgr(CINZA_TEXTO)
    conferencia.Font.Bold = False
    conferencia.Font.Size = 9
    _bordas_finas(ws, "A82:C87")


def _aplicar_retroativo_potencial(ws) -> None:
    """Le so RETROATIVO_POTENCIAL_PC; nao cria calculo nem desloca linha."""
    ws.Range(FAIXA_POTENCIAL).UnMerge()
    ws.Range(CELULA_POTENCIAL_VALOR).Formula = FORMULA_POTENCIAL_VALOR
    ws.Range(CELULA_POTENCIAL_ROTULO).Formula = FORMULA_POTENCIAL_ROTULO
    ws.Range("E22:F22").Merge()
    ws.Range("G22:H22").Merge()

    rotulo = ws.Range(CELULA_POTENCIAL_ROTULO)
    rotulo.Font.Size = 9
    rotulo.Font.Bold = False
    rotulo.HorizontalAlignment = XL_LEFT
    rotulo.VerticalAlignment = XL_VCENTER
    rotulo.WrapText = True

    valor = ws.Range(CELULA_POTENCIAL_VALOR)
    _moeda(valor)
    valor.Font.Size = 11
    valor.Font.Bold = True
    valor.HorizontalAlignment = XL_RIGHT
    valor.VerticalAlignment = XL_VCENTER

    faixa = ws.Range(FAIXA_POTENCIAL)
    faixa.Font.Color = _bgr(AMBAR_TEXTO)
    # Ambar somente quando ha valor: fora do metodo PCs a faixa some.
    faixa.FormatConditions.Delete()
    # Assinatura COM: Add(Type, Operator, Formula1). Com xlExpression o
    # operador nao se aplica e a formula tem de ir na 3a posicao.
    regra = faixa.FormatConditions.Add(XL_EXPRESSION, None, '=$G$22<>""')
    regra.Interior.Color = _bgr(AMBAR_FUNDO)
    regra.Font.Color = _bgr(AMBAR_TEXTO)
    ws.Rows("22:22").RowHeight = ALTURA_LINHA_POTENCIAL


def aplicar(caminho: Path) -> None:
    import win32com.client as com

    excel = com.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(caminho))
        ws = wb.Worksheets("RESULTADOS")

        # 1. AJUSTES MANUAIS: ocultar sem apagar, limpar ou mover.
        faixa_ajustes = "%d:%d" % (PRIMEIRA_LINHA_AJUSTES, ULTIMA_LINHA_AJUSTES)
        ws.Rows(faixa_ajustes).Hidden = True

        _aplicar_tabela_6(ws)
        _aplicar_tabela_7(ws)
        _aplicar_tabela_8(ws)
        _aplicar_tabela_9(ws)
        _aplicar_retroativo_potencial(ws)

        ws.Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
        ws.Range("D5").Select()
        wb.Worksheets("CONTROLE").Activate()
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1

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
        print("ERRO: arquivo nao encontrado: %s" % caminho)
        return 1
    aplicar(caminho)
    print("OK: acabamento RESULTADOS-FINAL-1 aplicado em %s" % caminho)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
