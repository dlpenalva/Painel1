<#
  RESULTADOS-UX2 / PR 2 - aplicador unico e reproduzivel.

  Parte do template do checkpoint 8bd98ce e produz o template final do PR 2:

    1. MEMORIA_RESULTADOS!S38/T38 + defined name RETROATIVO_POTENCIAL_PC;
    2. camada humana da aba RESULTADOS nas linhas 90-166 (espelhos);
    3. formatos, cores, tipografia e formatacao condicional;
    4. ocultacao da camada tecnica (linhas 1-89);
    5. print setup: paisagem, area A90:H166, quebra explicita antes da 117.

  O motor tecnico das linhas 1-87 NAO e tocado: nenhuma formula economica e
  movida, reancorada ou reescrita. C43:G50 permanece nas mesmas coordenadas.

  Todo o conteudo acentuado vive em tools/resultados_ux2_plano.json (UTF-8),
  gerado por tools/gerar_plano_resultados_ux2.py. Este script e ASCII puro
  de proposito: o PowerShell 5.1 desta maquina nao le acentos do proprio .ps1
  de forma confiavel.

  Requer apenas Excel instalado (COM). Nao introduz dependencia nova.

  Uso:
    powershell -File tools/aplicar_resultados_ux2_pr2.ps1 -Template <xlsx>
#>
param(
  [string]$Template = "templates/COLETA_REAJUSTE_OFICIAL.xlsx",
  [string]$Plano = "tools/resultados_ux2_plano.json"
)

$ErrorActionPreference = "Stop"

$alvo = (Resolve-Path $Template).Path
$planoPath = (Resolve-Path $Plano).Path
$p = Get-Content -Raw -Encoding UTF8 $planoPath | ConvertFrom-Json

# Constantes COM
$xlLandscape       = 2
$xlExpression      = 2
$xlPageBreakManual = -4135
$xlLeft            = -4131
$xlCenter          = -4108
$xlA4              = 9

function ConvertTo-BGR([string]$rgb) {
  # Excel COM usa BGR; o plano guarda RGB no formato hexadecimal usual.
  $r = [Convert]::ToInt32($rgb.Substring(0, 2), 16)
  $g = [Convert]::ToInt32($rgb.Substring(2, 2), 16)
  $b = [Convert]::ToInt32($rgb.Substring(4, 2), 16)
  return ($b * 65536) + ($g * 256) + $r
}

Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.AskToUpdateLinks = $false
$xl.Interactive = $false

$wb = $null
try {
  $wb = $xl.Workbooks.Open($alvo, 0, $false)
  Start-Sleep -Seconds 2

  $mem = $wb.Worksheets.Item("MEMORIA_RESULTADOS")
  $ws  = $wb.Worksheets.Item($p.aba)

  # ------------------------------------------------------------- ETAPA 1
  # Retroativo potencial PC (semantica ja aprovada nas FASES 1-4).
  if (-not $mem.Range("T38").Formula) {
    $mem.Range("S38").Value2 = $p.memoria.S38
    $mem.Range("T38").Formula = $p.memoria.T38
    Write-Output "  S38/T38 criados"
  } else {
    Write-Output "  S38/T38 ja presentes"
  }

  $temNome = $false
  foreach ($n in $wb.Names) {
    if ($n.Name -eq $p.name_novo.nome) { $temNome = $true }
  }
  if (-not $temNome) {
    $wb.Names.Add($p.name_novo.nome, $p.name_novo.refere) | Out-Null
    Write-Output "  name criado"
  } else {
    Write-Output "  name ja presente"
  }

  # ------------------------------------------------------------- ETAPA 2
  # O formato monetario e COPIADO de uma celula existente: reproduz o padrao
  # homologado sem depender do locale do Excel.
  $FMT_MOEDA = $ws.Range("C5").NumberFormat

  foreach ($m in $p.merges) {
    $r = $ws.Range($m)
    if (-not $r.MergeCells) { $r.Merge() }
  }
  Write-Output ("  merges aplicados: " + $p.merges.Count)

  foreach ($c in $p.celulas) {
    $rng = $ws.Range($c.ref)
    $campos = $c.PSObject.Properties.Name
    if ($campos -contains "formula") {
      $rng.Formula = $c.formula
    } elseif ($campos -contains "valor") {
      $rng.Value2 = $c.valor
    }
    # ARMADILHA: .NumberFormat no Excel pt-BR reinterpreta padroes em ingles
    # ("0.000000" -> "0,000,000"). Formato monetario e COPIADO de uma celula
    # homologada; os demais vao em pt-BR por .NumberFormatLocal.
    if ($campos -contains "fmt") {
      if ($c.fmt -eq "@COPIAR_DE:C5") { $rng.NumberFormat = $FMT_MOEDA }
    } elseif ($campos -contains "fmtl") {
      $rng.NumberFormatLocal = $c.fmtl
    }
    if ($campos -contains "tam") { $rng.Font.Size = $c.tam }
    if ($campos -contains "negrito") { $rng.Font.Bold = $true }
    if ($campos -contains "italico") { $rng.Font.Italic = $true }
    if ($campos -contains "cor") { $rng.Font.Color = (ConvertTo-BGR $c.cor) }
    if ($campos -contains "fundo") {
      $rng.Interior.Color = (ConvertTo-BGR $c.fundo)
    }
    if ($campos -contains "wrap") { $rng.WrapText = $true }
    $rng.HorizontalAlignment = $xlLeft
    $rng.VerticalAlignment = $xlCenter
  }
  Write-Output ("  celulas aplicadas: " + $p.celulas.Count)

  # (§25) Acertos de formato fora da camada nova: so number format, nunca
  # formula, valor ou validacao.
  # ARMADILHA: a aba CONTROLE e protegida; formatar nela sem Unprotect
  # devolve "nao e possivel definir a propriedade" (ou 0x800A03EC, que aborta
  # a sessao COM e perde TODAS as edicoes pendentes). Desprotege, formata e
  # reprotege, preservando o estado original de cada aba.
  foreach ($fx in $p.formatos_extra) {
    $aba = $wb.Worksheets.Item($fx.aba)
    $estavaProtegida = $aba.ProtectContents
    if ($estavaProtegida) { $aba.Unprotect() }
    $r = $aba.Range($fx.ref)
    if ($fx.PSObject.Properties.Name -contains "fmt") {
      $r.NumberFormat = $FMT_MOEDA
    } else {
      $r.NumberFormatLocal = $fx.fmtl
    }
    if ($estavaProtegida) { $aba.Protect() }
  }
  Write-Output ("  formatos extra: " + $p.formatos_extra.Count)

  $nAlturas = 0
  foreach ($prop in $p.alturas.PSObject.Properties) {
    $ws.Rows($prop.Name).RowHeight = [double]$prop.Value
    $nAlturas++
  }
  Write-Output ("  alturas aplicadas: " + $nAlturas)

  # Formatacao condicional: a cor deriva do status canonico, nunca fixa.
  # Para xlExpression o parametro Operator NAO se aplica: passar 0 faz o
  # Excel devolver "valor fora do intervalo esperado".
  foreach ($cf in $p.condicionais) {
    $rng = $ws.Range($cf.faixa)
    $fc = $rng.FormatConditions.Add($xlExpression, [Type]::Missing, $cf.expr)
    $campos = $cf.PSObject.Properties.Name
    if ($campos -contains "fundo") {
      $fc.Interior.Color = (ConvertTo-BGR $cf.fundo)
    }
    if ($campos -contains "cor") { $fc.Font.Color = (ConvertTo-BGR $cf.cor) }
  }
  Write-Output ("  condicionais aplicadas: " + $p.condicionais.Count)

  # ------------------------------------------------------------- ETAPA 3
  # Camada tecnica invisivel. C43:G50 continua nas mesmas coordenadas, com
  # formulas e validacoes intactas - apenas sai da leitura executiva.
  foreach ($faixa in $p.ocultar_linhas) {
    $ini = $faixa[0]
    $fim = $faixa[1]
    $ws.Rows("$ini`:$fim").Hidden = $true
  }
  foreach ($col in $p.ocultar_colunas) {
    $ws.Columns($col).Hidden = $true
  }
  Write-Output "  camada tecnica oculta"

  # ------------------------------------------------------------- ETAPA 4
  $ws.ResetAllPageBreaks()
  $ws.PageSetup.PrintArea = $p.impressao.print_area
  $ws.PageSetup.Orientation = $xlLandscape
  $ws.PageSetup.PaperSize = $xlA4
  # Escala explicita, nao FitToPages: com quebra manual o Excel calcula a
  # escala pela altura total e ignora o corte, devolvendo 3 paginas.
  $ws.PageSetup.FitToPagesWide = $false
  $ws.PageSetup.FitToPagesTall = $false
  $ws.PageSetup.Zoom = $p.impressao.zoom
  $ws.PageSetup.LeftMargin = $xl.InchesToPoints(0.3)
  $ws.PageSetup.RightMargin = $xl.InchesToPoints(0.3)
  $ws.PageSetup.TopMargin = $xl.InchesToPoints(0.4)
  $ws.PageSetup.BottomMargin = $xl.InchesToPoints(0.4)
  $linhaQuebra = [string]$p.impressao.quebra_antes_da_linha
  $ws.Rows($linhaQuebra).PageBreak = $xlPageBreakManual
  Write-Output "  print setup aplicado"

  # A leitura executiva comeca no topo da camada visivel.
  $ws.Activate()
  $xl.ActiveWindow.ScrollRow = 90
  $ws.Range("A90").Select() | Out-Null

  $xl.Application.CalculateFullRebuild()
  Start-Sleep -Seconds 3
  $wb.Save()
  $wb.Close($false)
  $wb = $null
  Write-Output "SALVO_OK"

  # Reabertura de prova: sem reparo, sem erro.
  $wb2 = $xl.Workbooks.Open($alvo, 0, $false)
  Start-Sleep -Seconds 2
  $paginas = $wb2.Worksheets.Item($p.aba).HPageBreaks.Count + 1
  Write-Output ("REABERTO_OK abas=" + $wb2.Worksheets.Count +
                " names=" + $wb2.Names.Count +
                " paginas=" + $paginas)
  $wb2.Close($false)
} catch {
  Write-Output ("ERRO: " + $_.Exception.Message)
  if ($wb) { try { $wb.Close($false) } catch {} }
  try { $xl.Quit() } catch {}
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  exit 1
} finally {
  try { $xl.Quit() } catch {}
  [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  Start-Sleep -Seconds 2
  Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
}
