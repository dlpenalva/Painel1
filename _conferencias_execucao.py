"""Conferencias informativas; nenhuma formula oficial do VTA e alterada."""
from __future__ import annotations


def classificar_conferencia(financeiro, fisico, periodo_financeiro=None,
                            periodo_fisico=None, completo=False,
                            aplicavel=True, justificativa=None, tolerancia=.05):
    if not aplicavel:
        return 'NÃO APLICÁVEL'
    if (financeiro is None or fisico is None or not completo
            or not periodo_financeiro or not periodo_fisico):
        return 'BASE INSUFICIENTE'
    if periodo_financeiro != periodo_fisico:
        return 'COMPARAÇÃO NÃO APLICÁVEL'
    if abs(financeiro - fisico) <= tolerancia:
        return 'VALIDADO'
    return 'JUSTIFICADO' if str(justificativa or '').strip() else 'CONFERIR'


def aplicar_conferencias(wb):
    """Reutiliza C73:E77; periodos/completude ficam na memoria oculta.

    Fotografias delimitam o periodo fisico. Competencia vazia nao e zero.
    Sem fotografia seguinte nao se infere encerramento do ciclo vigente.
    """
    det = wb['RESULTADOS_DETALHE']
    mem = wb['MEMORIA_RESULTADOS']
    for n, col in enumerate(('AC', 'N', 'P', 'R', 'T')):
        r, p, h = 73 + n, 2 + n, 280 + n
        ids = 'itens_Remanesc!$A$2:$A$201'
        vals = f'itens_Remanesc!${col}$2:${col}$201'
        det[f'C{r}'] = (f'=IF(MEMORIA_RESULTADOS!$B$4<>"Financeiro","Nao aplicavel ao metodo selecionado",'
            f'IF(OR(COUNTIF({ids},"<>")=0,SUMPRODUCT(--({ids}<>""),'
            f'--(ISNUMBER({vals})=FALSE))>0),"Sem historico quantitativo suficiente",'
            f'ROUND(SUMIF({ids},"<>",{vals}),2)))')
        if n == 4:
            det[f'C{r}'] = ('=IF(MEMORIA_RESULTADOS!$B$4<>"Financeiro",'
                '"Nao aplicavel ao metodo selecionado","Sem historico quantitativo suficiente")')
        # S:Z e reservado nesta faixa, fora das tabelas existentes.
        mem[f'S{h}'] = f'C{n}'
        mem[f'T{h}'] = f'=parametros!C{p}'
        mem[f'U{h}'] = f'=parametros!D{p}+1'
        mem[f'V{h}'] = f'=parametros!I{p}'
        mem[f'W{h}'] = f'=parametros!I{p+1}' if n < 4 else '=""'
        # Uma competencia por mes, todas numericas; 12 linhas duplicadas nao
        # comprovam cobertura anual. Nao inferir fechamento por contagem apenas.
        checks = [f'COUNTIFS(financeiro!$B$2:$B$73,"c{n}",'
                  f'financeiro!$A$2:$A$73,DATE(YEAR(T{h}),MONTH(T{h})+{k},1))=1'
                  for k in range(12)]
        mem[f'X{h}'] = (f'=IF(OR(NOT(ISNUMBER(T{h})),NOT(ISNUMBER(U{h})),U{h}<=T{h}),0,'
            f'IF(AND(DAY(T{h})=1,U{h}=DATE(YEAR(T{h})+1,MONTH(T{h}),1),'
            f'COUNTIF(financeiro!$B$2:$B$73,"c{n}")=12,'
            + ','.join(checks) + ','
            f'SUMPRODUCT(--(financeiro!$B$2:$B$73="c{n}"),'
            f'--ISNUMBER(financeiro!$C$2:$C$73))=12),1,0))')
        m = 'MEMORIA_RESULTADOS!'
        det[f'E{r}'] = (f'=IF({m}$B$4<>"Financeiro","NÃO APLICÁVEL",'
            f'IF({n}>IFERROR(VALUE(RIGHT(CONTROLE!$B$2,1)),4),"NÃO APLICÁVEL",'
            f'IF(OR(NOT(ISNUMBER({m}V{h})),NOT(ISNUMBER({m}W{h})),{m}V{h}=0,{m}W{h}=0,'
            f'C{r}="",NOT(ISNUMBER(B{r})),NOT(ISNUMBER(C{r}))),"BASE INSUFICIENTE",'
            f'IF({m}X{h}<>1,"BASE INSUFICIENTE",'
            f'IF(OR({m}V{h}<>{m}T{h},{m}W{h}<>{m}U{h}),"COMPARAÇÃO NÃO APLICÁVEL",'
            f'IF(ABS(D{r})<={m}$D$4,"VALIDADO","CONFERIR"))))))')
    # Preserva a precedencia dos erros de calculo e do status ESTIMADO.
    old = str(det['B3'].value)
    if 'COUNTIF($E$73:$E$77' not in old:
        old = old.replace('"VALIDADO"))',
            'IF(OR(LEFT(MEMORIA_RESULTADOS!$W$52,6)="REVISE",'
            'COUNTIF($E$73:$E$77,"CONFERIR")>0),"CONFERIR",'
            'IF(OR(COUNTIF($E$73:$E$77,"BASE INSUFICIENTE")>0,'
            'COUNTIF($E$73:$E$77,"COMPARAÇÃO NÃO APLICÁVEL")>0),'
            '"BASE INSUFICIENTE","VALIDADO"))))')
        det['B3'] = old
    wb['RESULTADOS']['D27'] = (
        '=IF(AND(ISNUMBER(VTA_FINAL),ISNUMBER(MEMORIA_RESULTADOS!B20),'
        'VTA_FINAL<MEMORIA_RESULTADOS!B20),'
        '"VTA inferior ao contrato original. A composição pode resultar em redução; '
        'confira execução passada, saldo remanescente, datas e justificativas. '
        'Diferenças de conferência não alteram o VTA. Status: "&STATUS_RESULTADOS,'
        '"VTA oficial calculado pelo método selecionado; confira as ressalvas abaixo.")')
    wb['RESULTADOS']['D32'] = (
        '=MEMORIA_RESULTADOS!W52&"; C0: "&RESULTADOS_DETALHE!E73&'
        '"; C1: "&RESULTADOS_DETALHE!E74&"; C2: "&RESULTADOS_DETALHE!E75&'
        '"; C3: "&RESULTADOS_DETALHE!E76&"; C4: "&RESULTADOS_DETALHE!E77&'
        '". Conferência informativa; não altera o VTA."')


def conferir_workbook(wb, justificativas=None):
    """Le cached values do XLSX; nunca grava no arquivo do fiscal."""
    from datetime import datetime, date, timedelta
    justificativas = justificativas or {}
    def number(v):
        return isinstance(v,(int,float)) and not isinstance(v,bool)
    def data(v):
        return v.date() if isinstance(v,datetime) else v if isinstance(v,date) else None
    registros = []
    for n,col in enumerate(('AC','N','P','R','T')):
        ciclo, p = f'C{n}', n+2
        inicio = data(wb['parametros'][f'C{p}'].value)
        fim = data(wb['parametros'][f'D{p}'].value)
        pf = (inicio,fim+timedelta(days=1)) if inicio and fim else None
        a = data(wb['parametros'][f'I{p}'].value)
        b = data(wb['parametros'][f'I{p+1}'].value) if n<4 else None
        pp = (a,b) if a and b else None
        rows = [row for row in wb['financeiro'].iter_rows(min_row=2,max_row=73)
                if str(row[1].value or '').upper()==ciclo]
        pagos = [row[2].value for row in rows if number(row[2].value)]
        datas = [data(row[0].value) for row in rows if number(row[2].value)]
        completo = bool(pf and len(rows)==12 and len(datas)==12
                        and len({(d.year,d.month) for d in datas if d})==12
                        and inicio.day==1
                        and pf[1].year==inicio.year+1 and pf[1].month==inicio.month
                        and pf[1].day==1
                        and all(d and d.day==1 and inicio<=d<=fim for d in datas))
        itens = [r for r in range(2,202) if wb['itens_Remanesc'][f'A{r}'].value not in (None,'')]
        fisicos = [wb['itens_Remanesc'][f'{col}{r}'].value for r in itens]
        financeiro = round(sum(pagos),2) if pagos else None
        fisico = round(sum(fisicos),2) if n < 4 and fisicos and all(number(v) for v in fisicos) else None
        metodo = wb['MEMORIA_RESULTADOS']['B4'].value
        vigente = str(wb['CONTROLE']['B2'].value or '').upper()
        aplicavel = metodo=='Financeiro' and ciclo<=vigente
        status = classificar_conferencia(financeiro,fisico,pf,pp,completo,aplicavel,
                                         justificativas.get(ciclo))
        registros.append(dict(ciclo=ciclo,financeiro=financeiro,fisico=fisico,
            diferenca=round(fisico-financeiro,2) if fisico is not None and financeiro is not None else None,
            periodo_financeiro=pf,periodo_fisico=pp,completo=completo,
            status=status,bloqueia_formalizacao=False))
    return registros
