from io import BytesIO
from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st

from _ui_utils import render_marca_topo, render_aviso_privacidade


st.set_page_config(page_title="TLB · cl8us - Coleta Única", layout="wide")


def gerar_coleta_unica_inteligente():
    output = BytesIO()

    import pandas as pd

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        workbook = writer.book

        # Paleta geral
        azul = "#1F4E79"
        azul_claro = "#DDEBF7"
        verde = "#0F766E"
        verde_claro = "#CCFBF1"
        amarelo = "#FFF2CC"
        laranja = "#FCE4D6"
        lilas = "#F3E8FF"
        cinza = "#E7E6E6"
        bege = "#F6F3EE"
        vermelho = "#FCE4D6"

        fmt_titulo = workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": azul,
            "font_size": 14, "align": "left", "valign": "vcenter"
        })
        fmt_subtitulo = workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": verde,
            "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True
        })
        fmt_header = workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": azul,
            "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True
        })
        fmt_header_verde = workbook.add_format({
            "bold": True, "font_color": "white", "bg_color": verde,
            "align": "center", "valign": "vcenter", "border": 1, "text_wrap": True
        })
        fmt_texto = workbook.add_format({"border": 1, "valign": "top", "text_wrap": True})
        fmt_input = workbook.add_format({"border": 1, "bg_color": amarelo, "valign": "top", "text_wrap": True})
        fmt_input_money = workbook.add_format({"border": 1, "bg_color": amarelo, "num_format": 'R$ #,##0.00'})
        fmt_input_num = workbook.add_format({"border": 1, "bg_color": amarelo, "num_format": '#,##0.00'})
        fmt_input_date = workbook.add_format({"border": 1, "bg_color": amarelo, "num_format": "dd/mm/yyyy"})
        fmt_auto = workbook.add_format({"border": 1, "bg_color": cinza, "text_wrap": True})
        fmt_auto_money = workbook.add_format({"border": 1, "bg_color": cinza, "num_format": 'R$ #,##0.00'})
        fmt_auto_num = workbook.add_format({"border": 1, "bg_color": cinza, "num_format": '#,##0.00'})
        fmt_obs = workbook.add_format({"border": 1, "bg_color": bege, "text_wrap": True, "valign": "top"})
        fmt_total = workbook.add_format({"bold": True, "border": 1, "bg_color": "#D9EAD3"})
        fmt_total_money = workbook.add_format({"bold": True, "border": 1, "bg_color": "#D9EAD3", "num_format": 'R$ #,##0.00'})
        fmt_ciclo_anterior = workbook.add_format({"border": 1, "bg_color": "#F4E7D3", "num_format": '#,##0.00'})
        fmt_ciclo_atual = workbook.add_format({"border": 1, "bg_color": verde_claro, "num_format": '#,##0.00'})
        fmt_estimativo = workbook.add_format({"border": 1, "bg_color": lilas, "num_format": '#,##0.00'})
        fmt_alerta = workbook.add_format({"border": 1, "bg_color": vermelho, "text_wrap": True, "font_color": "#9C0006"})
        fmt_ok = workbook.add_format({"border": 1, "bg_color": "#E2F0D9", "text_wrap": True, "font_color": "#274E13"})
        fmt_percent = workbook.add_format({"border": 1, "bg_color": amarelo, "num_format": "0.00%"})
        fmt_factor = workbook.add_format({"border": 1, "bg_color": cinza, "num_format": "0.0000"})

        data_geracao = datetime.now(ZoneInfo("America/Sao_Paulo")).strftime("%d/%m/%Y %H:%M")

        # =====================================================
        # INICIO
        # =====================================================
        ws = workbook.add_worksheet("INICIO")
        writer.sheets["INICIO"] = ws
        ws.hide_gridlines(2)
        ws.set_column("A:A", 4)
        ws.set_column("B:B", 36)
        ws.set_column("C:C", 96)
        ws.merge_range("B2:C2", "Coleta Única Inteligente — Experimental", fmt_titulo)
        ws.merge_range(
            "B3:C5",
            "Modelo experimental para capturar, em uma única base, dados contratuais, ciclos, histórico financeiro, itens por ciclo, aditivos, glosas, corte operacional e validações fiscais. Esta planilha ainda não substitui o Arquivo de Coleta atual.",
            fmt_obs,
        )

        orientacoes = [
            ("1. Finalidade", "Permitir que o fiscal informe o que possui: financeiro completo, itens completos, itens parciais, saldo atual, aditivos, glosas ou ressalvas."),
            ("2. Regra central", "O sistema deve identificar o modo possível de apuração a partir da qualidade da base preenchida, e não exigir que o fiscal escolha o modo."),
            ("3. Financeiro histórico", "Base preferencial para cálculo financeiro e retroativo, quando preenchida até o mês atual ou competência de corte."),
            ("4. Itens por ciclo", "Base física/itemizada. Pode estar completa desde C0 ou apenas parcial, conforme disponibilidade da fiscalização."),
            ("5. Maleabilidade", "A ausência de uma aba não deve quebrar o sistema. Deve gerar diagnóstico, ressalva e modo de cálculo compatível."),
            ("6. Atenção", "Este modelo é experimental. Use para desenho e validação de fluxo antes de integrar ao módulo Valores."),
        ]

        row = 7
        for titulo, texto in orientacoes:
            ws.write(row, 1, titulo, fmt_subtitulo)
            ws.write(row, 2, texto, fmt_texto)
            ws.set_row(row, 42)
            row += 1

        ws.write("B15", "Data de geração", fmt_subtitulo)
        ws.write("C15", data_geracao, fmt_texto)

        # =====================================================
        # PARAMETROS_CONTRATO
        # =====================================================
        ws = workbook.add_worksheet("PARAMETROS_CONTRATO")
        writer.sheets["PARAMETROS_CONTRATO"] = ws
        ws.hide_gridlines(2)
        ws.set_column("A:A", 42)
        ws.set_column("B:B", 44)
        ws.set_column("C:C", 90)
        ws.merge_range("A1:C1", "Parâmetros do Contrato", fmt_titulo)
        headers = ["Campo", "Valor", "Orientação"]
        for col, h in enumerate(headers):
            ws.write(2, col, h, fmt_header)

        linhas = [
            ("Contrato", "", "Número do contrato."),
            ("Processo", "", "Número do processo administrativo."),
            ("Contratada", "", "Nome da empresa contratada."),
            ("Índice contratual", "", "Ex.: IST, IPCA, IGP-M, ICTI."),
            ("Data-base original", "", "Data da proposta ou marco inicial contratual."),
            ("Vigência inicial", "", "Data de início da vigência."),
            ("Vigência final", "", "Data final da vigência atual."),
            ("Valor original do contrato", "", "Valor inicial do contrato em C0."),
            ("Valor formalizado antes desta análise", "", "Valor já formalizado antes da análise atual, se houver."),
            ("Último ciclo já formalizado", "C0 / Nenhum", "Ex.: C1, C2. Use C0 / Nenhum se não houver ciclo anterior formalizado."),
            ("Data do pedido do último ciclo formalizado", "", "Preencher quando houver ciclo anterior concedido/formalizado."),
            ("Ciclo inicial da análise atual", "", "Ex.: C1, C2, C3. Pode ser derivado do histórico."),
            ("Ciclo atual / ciclo de corte", "", "Último ciclo que deve aparecer na coleta."),
            ("Competência atual / competência de corte", "", "Mês mais recente para o financeiro histórico. Ex.: 04/2026."),
            ("Observação geral", "", "Campo livre para contexto da fiscalização."),
        ]

        for r, (campo, valor, orientacao) in enumerate(linhas, start=3):
            ws.write(r, 0, campo, fmt_texto)
            if "Valor" in campo:
                ws.write(r, 1, valor, fmt_input_money)
            elif "Data" in campo or "Vigência" in campo:
                ws.write(r, 1, valor, fmt_input_date)
            else:
                ws.write(r, 1, valor, fmt_input)
            ws.write(r, 2, orientacao, fmt_obs)

        # =====================================================
        # CICLOS
        # =====================================================
        ws = workbook.add_worksheet("CICLOS")
        writer.sheets["CICLOS"] = ws
        ws.hide_gridlines(2)
        headers = [
            "Ciclo", "Data-base", "Data do pedido", "Início financeiro", "Fim financeiro",
            "Situação", "Percentual aplicado", "Fator do ciclo", "Fator acumulado",
            "Objeto da análise atual?", "Ciclo já formalizado?", "Observação"
        ]
        ws.merge_range(0, 0, 0, len(headers)-1, "Ciclos contratuais — C0 até ciclo atual/corte", fmt_titulo)
        for c, h in enumerate(headers):
            ws.write(2, c, h, fmt_header)

        ciclos_base = ["C0", "C1", "C2", "C3", "C4", "C5"]
        for r, ciclo in enumerate(ciclos_base, start=3):
            ws.write(r, 0, ciclo, fmt_input)
            ws.write(r, 1, "", fmt_input_date)
            ws.write(r, 2, "Não se aplica" if ciclo == "C0" else "", fmt_input)
            ws.write(r, 3, "Não se aplica" if ciclo == "C0" else "", fmt_input)
            ws.write(r, 4, "Não se aplica" if ciclo == "C0" else "", fmt_input)
            ws.write(r, 5, "Base sem reajuste" if ciclo == "C0" else "", fmt_input)
            ws.write_number(r, 6, 0 if ciclo == "C0" else 0, fmt_percent)
            ws.write_number(r, 7, 1 if ciclo == "C0" else 1, fmt_factor)
            ws.write_number(r, 8, 1 if ciclo == "C0" else 1, fmt_factor)
            ws.write(r, 9, "Não" if ciclo == "C0" else "", fmt_input)
            ws.write(r, 10, "Sim" if ciclo == "C0" else "", fmt_input)
            ws.write(r, 11, "C0 é ciclo-base e não reajusta." if ciclo == "C0" else "", fmt_obs)

        ws.set_column("A:A", 12)
        ws.set_column("B:E", 20)
        ws.set_column("F:F", 30)
        ws.set_column("G:I", 18)
        ws.set_column("J:K", 24)
        ws.set_column("L:L", 70)
        ws.freeze_panes(3, 0)

        # =====================================================
        # FINANCEIRO_HISTORICO
        # =====================================================
        ws = workbook.add_worksheet("FINANCEIRO_HISTORICO")
        writer.sheets["FINANCEIRO_HISTORICO"] = ws
        ws.hide_gridlines(2)
        headers = [
            "Ciclo sugerido", "Competência", "Valor medido/aprovado",
            "Valor faturado", "Valor pago", "Glosa/retenção/desconto",
            "Valor líquido considerado", "Fonte", "Observação"
        ]
        ws.merge_range(0, 0, 0, len(headers)-1, "Histórico financeiro mensal até a competência atual/corte", fmt_titulo)
        ws.merge_range(
            1, 0, 1, len(headers)-1,
            "Base preferencial para cálculo financeiro. Preencher até o mês atual ou até a competência de corte. Se houver glosa/retenção/desconto, registrar para permitir leitura correta.",
            fmt_obs,
        )
        for c, h in enumerate(headers):
            ws.write(3, c, h, fmt_header)
        for r in range(4, 124):
            ws.write(r, 0, "", fmt_input)
            ws.write(r, 1, "", fmt_input)
            ws.write(r, 2, "", fmt_input_money)
            ws.write(r, 3, "", fmt_input_money)
            ws.write(r, 4, "", fmt_input_money)
            ws.write(r, 5, "", fmt_input_money)
            ws.write_formula(r, 6, f'=IF(C{r+1}="","",ROUND(C{r+1}-F{r+1},2))', fmt_auto_money)
            ws.write(r, 7, "", fmt_input)
            ws.write(r, 8, "", fmt_obs)
        total_row = 124
        ws.write(total_row, 0, "TOTAL", fmt_total)
        for c in range(1, 6):
            ws.write(total_row, c, "", fmt_total)
        ws.write_formula(total_row, 6, f"=ROUND(SUM(G5:G124),2)", fmt_total_money)
        ws.set_column("A:A", 16)
        ws.set_column("B:B", 16)
        ws.set_column("C:G", 24)
        ws.set_column("H:H", 26)
        ws.set_column("I:I", 70)
        ws.freeze_panes(4, 0)

        # =====================================================
        # ITENS_CICLOS
        # =====================================================
        ws = workbook.add_worksheet("ITENS_CICLOS")
        writer.sheets["ITENS_CICLOS"] = ws
        ws.hide_gridlines(2)
        headers = [
            "Item", "Descrição resumida", "Unidade", "Quantidade contratada C0",
            "Valor unitário original C0", "Valor total original C0",
            "Remanescente C1", "Remanescente C2", "Remanescente C3",
            "Remanescente C4", "Remanescente ciclo atual/corte",
            "Consumo informado por ciclo?", "Observação fiscal"
        ]
        ws.merge_range(0, 0, 0, len(headers)-1, "Itens por ciclo — preenchimento flexível", fmt_titulo)
        ws.merge_range(
            1, 0, 1, len(headers)-1,
            "Itens de ciclos anteriores podem ficar em branco se a fiscalização possuir apenas financeiro histórico. Campos de ciclos anteriores usam cor de memória; ciclo atual/corte usa cor específica.",
            fmt_obs,
        )
        for c, h in enumerate(headers):
            ws.write(3, c, h, fmt_header if c < 10 else fmt_header_verde)

        for r in range(4, 184):
            ws.write(r, 0, "", fmt_input)
            ws.write(r, 1, "", fmt_input)
            ws.write(r, 2, "", fmt_input)
            ws.write(r, 3, "", fmt_input_num)
            ws.write(r, 4, "", fmt_input_money)
            ws.write_formula(r, 5, f'=IF(OR(D{r+1}="",E{r+1}=""),"",ROUND(D{r+1}*E{r+1},2))', fmt_auto_money)
            ws.write(r, 6, "", fmt_ciclo_anterior)
            ws.write(r, 7, "", fmt_ciclo_anterior)
            ws.write(r, 8, "", fmt_ciclo_anterior)
            ws.write(r, 9, "", fmt_ciclo_anterior)
            ws.write(r, 10, "", fmt_ciclo_atual)
            ws.write(r, 11, "", fmt_estimativo)
            ws.write(r, 12, "", fmt_obs)
        total_row = 184
        ws.write(total_row, 0, "TOTAL", fmt_total)
        for c in range(1, 5):
            ws.write(total_row, c, "", fmt_total)
        ws.write_formula(total_row, 5, "=ROUND(SUM(F5:F184),2)", fmt_total_money)
        for c in range(6, 11):
            col_letter = chr(ord("A") + c)
            ws.write_formula(total_row, c, f"=ROUND(SUM({col_letter}5:{col_letter}184),2)", fmt_total)
        ws.set_column("A:A", 12)
        ws.set_column("B:B", 42)
        ws.set_column("C:C", 14)
        ws.set_column("D:K", 22)
        ws.set_column("L:L", 24)
        ws.set_column("M:M", 70)
        ws.freeze_panes(4, 0)

        # =====================================================
        # ADITIVOS_SUPRESSOES
        # =====================================================
        ws = workbook.add_worksheet("ADITIVOS_SUPRESSOES")
        writer.sheets["ADITIVOS_SUPRESSOES"] = ws
        ws.hide_gridlines(2)
        headers = [
            "Identificação", "Data do instrumento", "Data de efeito", "Ciclo/marco financeiro",
            "Tipo", "Item", "Quantidade", "Valor unitário", "Valor original",
            "Já formalizado?", "Incorporar no Valor Total?", "Observação"
        ]
        ws.merge_range(0, 0, 0, len(headers)-1, "Aditivos e supressões", fmt_titulo)
        for c, h in enumerate(headers):
            ws.write(2, c, h, fmt_header)
        for r in range(3, 83):
            ws.write(r, 0, "", fmt_input)
            ws.write(r, 1, "", fmt_input_date)
            ws.write(r, 2, "", fmt_input_date)
            ws.write(r, 3, "", fmt_auto)
            ws.write(r, 4, "Acréscimo", fmt_input)
            ws.write(r, 5, "", fmt_input)
            ws.write(r, 6, "", fmt_input_num)
            ws.write(r, 7, "", fmt_input_money)
            ws.write_formula(r, 8, f'=IF(OR(G{r+1}="",H{r+1}=""),"",ROUND(G{r+1}*H{r+1},2))', fmt_auto_money)
            ws.write(r, 9, "", fmt_input)
            ws.write(r, 10, "", fmt_input)
            ws.write(r, 11, "", fmt_obs)
        ws.set_column("A:A", 22)
        ws.set_column("B:D", 20)
        ws.set_column("E:E", 18)
        ws.set_column("F:F", 16)
        ws.set_column("G:I", 20)
        ws.set_column("J:K", 24)
        ws.set_column("L:L", 70)
        ws.data_validation("E4:E83", {"validate": "list", "source": ["Acréscimo", "Supressão"]})
        ws.data_validation("J4:K83", {"validate": "list", "source": ["Sim", "Não"]})

        # =====================================================
        # GLOSAS_AJUSTES
        # =====================================================
        ws = workbook.add_worksheet("GLOSAS_AJUSTES")
        writer.sheets["GLOSAS_AJUSTES"] = ws
        ws.hide_gridlines(2)
        headers = ["Competência", "Ciclo", "Tipo de ajuste", "Valor", "Afeta retroativo?", "Referência documental", "Observação"]
        ws.merge_range(0, 0, 0, len(headers)-1, "Glosas, retenções, descontos e ajustes", fmt_titulo)
        for c, h in enumerate(headers):
            ws.write(2, c, h, fmt_header)
        for r in range(3, 83):
            ws.write(r, 0, "", fmt_input)
            ws.write(r, 1, "", fmt_input)
            ws.write(r, 2, "", fmt_input)
            ws.write(r, 3, "", fmt_input_money)
            ws.write(r, 4, "", fmt_input)
            ws.write(r, 5, "", fmt_input)
            ws.write(r, 6, "", fmt_obs)
        ws.data_validation("E4:E83", {"validate": "list", "source": ["Sim", "Não"]})
        ws.set_column("A:B", 16)
        ws.set_column("C:C", 24)
        ws.set_column("D:D", 20)
        ws.set_column("E:F", 24)
        ws.set_column("G:G", 70)

        # =====================================================
        # CICLO_EM_EXECUCAO
        # =====================================================
        ws = workbook.add_worksheet("CICLO_EM_EXECUCAO")
        writer.sheets["CICLO_EM_EXECUCAO"] = ws
        ws.hide_gridlines(2)
        ws.merge_range("A1:D1", "Corte operacional no ciclo em execução", fmt_titulo)
        ws.merge_range(
            "A2:D2",
            "Usar somente quando houver fotografia intermediária do ciclo atual. Se marcado como Não, o sistema deve manter o corte padrão no início dos ciclos.",
            fmt_obs,
        )
        headers = ["Campo", "Valor", "Orientação", "Uso pelo cl8us"]
        for c, h in enumerate(headers):
            ws.write(3, c, h, fmt_header_verde)
        linhas = [
            ("Aplicar corte operacional?", "Não", "Preencha Sim apenas se houver fotografia intermediária no ciclo atual.", "Chave principal."),
            ("Ciclo em execução", "", "Ex.: C3.", "Identifica o ciclo de corte."),
            ("Competência de corte operacional", "", "Ex.: 04/2026.", "Limita a execução financeira até essa competência."),
            ("Fonte da execução realizada", "Financeiro histórico", "Preferencialmente usar FINANCEIRO_HISTORICO.", "Define a origem da execução realizada."),
            ("Valor remanescente original no corte", "", "Preencher se o saldo estiver em valor original.", "O sistema poderá aplicar fator."),
            ("Valor remanescente atualizado no corte", "", "Preencher se o saldo já estiver atualizado.", "Usar diretamente, sem novo reajuste."),
            ("Observação fiscal", "", "Registrar premissa da fotografia.", "Memória da apuração."),
        ]
        for r, linha in enumerate(linhas, start=4):
            ws.write(r, 0, linha[0], fmt_texto)
            ws.write(r, 1, linha[1], fmt_input)
            ws.write(r, 2, linha[2], fmt_obs)
            ws.write(r, 3, linha[3], fmt_obs)
        ws.data_validation("B5", {"validate": "list", "source": ["Não", "Sim"]})
        ws.set_column("A:A", 42)
        ws.set_column("B:B", 34)
        ws.set_column("C:C", 80)
        ws.set_column("D:D", 52)

        # =====================================================
        # VALIDACOES_FISCAIS
        # =====================================================
        ws = workbook.add_worksheet("VALIDACOES_FISCAIS")
        writer.sheets["VALIDACOES_FISCAIS"] = ws
        ws.hide_gridlines(2)
        ws.merge_range("A1:D1", "Validações fiscais e premissas da base", fmt_titulo)
        headers = ["Pergunta", "Resposta", "Obrigatório?", "Observação"]
        for c, h in enumerate(headers):
            ws.write(2, c, h, fmt_header)
        perguntas = [
            ("O histórico financeiro está completo até a competência atual/corte?", "", "Sim", "Base preferencial para cálculo financeiro."),
            ("Os valores financeiros representam medição/aprovação/faturamento devido?", "", "Sim", "Indicar ressalvas se houver diferença entre pago, faturado e medido."),
            ("Há glosas, descontos, multas ou retenções relevantes?", "", "Sim", "Se sim, preencher GLOSAS_AJUSTES."),
            ("Os itens de ciclos anteriores foram informados?", "", "Não", "Se não, o sistema poderá usar apenas financeiro histórico."),
            ("Os itens do ciclo atual/corte foram informados?", "", "Sim", "Importante para saldo remanescente."),
            ("Há aditivos/supressões a considerar?", "", "Não", "Se sim, preencher ADITIVOS_SUPRESSOES."),
            ("A base de itens equivale à execução medida/aprovada?", "", "Condicional", "Obrigatório nos modos por itens."),
            ("Há ressalvas fiscais sobre a base informada?", "", "Sim", "Registrar no campo observação."),
        ]
        for r, p in enumerate(perguntas, start=3):
            ws.write(r, 0, p[0], fmt_texto)
            ws.write(r, 1, p[1], fmt_input)
            ws.write(r, 2, p[2], fmt_texto)
            ws.write(r, 3, p[3], fmt_obs)
        ws.data_validation("B4:B11", {"validate": "list", "source": ["Sim", "Não", "Parcial", "Não se aplica"]})
        ws.set_column("A:A", 58)
        ws.set_column("B:B", 18)
        ws.set_column("C:C", 18)
        ws.set_column("D:D", 82)

        # =====================================================
        # DIAGNOSTICO_BASE
        # =====================================================
        ws = workbook.add_worksheet("DIAGNOSTICO_BASE")
        writer.sheets["DIAGNOSTICO_BASE"] = ws
        ws.hide_gridlines(2)
        ws.merge_range("A1:D1", "Diagnóstico automático da base — futuro leitor do cl8us", fmt_titulo)
        ws.merge_range(
            "A2:D2",
            "Nesta versão experimental, o diagnóstico é orientativo. Na próxima etapa, o módulo Valores deverá ler essas condições e classificar automaticamente o modo possível de apuração.",
            fmt_obs,
        )
        headers = ["Critério", "Resultado esperado", "Uso futuro", "Ressalva"]
        for c, h in enumerate(headers):
            ws.write(3, c, h, fmt_header)

        diagnosticos = [
            ("Financeiro histórico completo", "Sim/Não/Parcial", "Define se o cálculo financeiro pode ser principal.", "Se Não, não calcular retroativo financeiro definitivo."),
            ("Itens por ciclo completos", "Sim/Não/Parcial", "Permite memória física completa por ciclo.", "Se parcial, gerar ressalva."),
            ("Itens atuais/remanescentes disponíveis", "Sim/Não", "Permite calcular saldo remanescente atualizado.", "Se ausente, limitar Valor Total."),
            ("Glosas/ajustes informados", "Sim/Não", "Permite tratar líquido considerado.", "Se houver glosa não informada, ressalvar."),
            ("Aditivos/supressões informados", "Sim/Não", "Permite governança do valor global.", "Evitar dupla contagem."),
            ("Corte operacional solicitado", "Sim/Não", "Define uso da aba CICLO_EM_EXECUCAO.", "Se Não, manter corte padrão."),
            ("Modo recomendado", "A definir pelo sistema", "Completo / Financeiro Histórico / Itens / Reduzido / Híbrido.", "Não escolher manualmente nesta etapa."),
        ]

        for r, d in enumerate(diagnosticos, start=4):
            ws.write(r, 0, d[0], fmt_texto)
            ws.write(r, 1, d[1], fmt_auto)
            ws.write(r, 2, d[2], fmt_obs)
            ws.write(r, 3, d[3], fmt_obs)

        ws.set_column("A:A", 36)
        ws.set_column("B:B", 26)
        ws.set_column("C:D", 78)

    output.seek(0)
    return output.getvalue()


render_marca_topo()
st.title("Coleta Única Inteligente")

render_aviso_privacidade(tem_download=True)

st.info(
    "Módulo experimental para gerar uma Coleta Única mais maleável. "
    "Nesta etapa, o XLSX é apenas modelo de preenchimento e diagnóstico. "
    "Ainda não altera o cálculo do módulo Valores nem substitui o Arquivo de Coleta atual."
)

st.markdown(
    """
### Objetivo

Criar uma base única capaz de receber diferentes níveis de informação fiscal:

- financeiro histórico completo;
- itens desde C0 até o ciclo atual;
- itens apenas do ciclo atual;
- financeiro sem itens anteriores;
- aditivos/supressões;
- glosas e ajustes;
- corte operacional no ciclo em execução;
- validações fiscais e ressalvas.
"""
)

xlsx = gerar_coleta_unica_inteligente()

st.download_button(
    label="Baixar Coleta Única Experimental em XLSX",
    data=xlsx,
    file_name="Coleta_Unica_Inteligente_Experimental.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)

st.warning(
    "Esta versão é experimental. Use para validar a estrutura de coleta. "
    "O leitor do módulo Valores será adaptado somente em etapa posterior."
)
