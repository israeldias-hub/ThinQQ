# Suplemento do PowerPoint: roda a suite VBA do projeto de teste (.pptm) e testes de ponta a ponta (Excel proprio com o core de teste +
# PowerPoint proprio): colar vinculado, mudar o dado, PptTqq_Sync (a) com o Excel injetado como "do usuario" (nao recebe Quit) e (b) com um
# Excel novo injetado como "criado pelo suplemento": o suplemento abre o core da pasta de versoes (a mais nova que responde, somente leitura),
# sincroniza e da Quit nele. So roda sem POWERPNT aberto (instancia unica). O PowerPoint criado aqui e encerrado pelo PID no fim
# (Close-PptApp), inclusive depois de erro ou de watchdog. Sem checagem de compilacao pelo VBE (o VBE do PowerPoint nao lista o projeto
# aberto sem janela e mexer nele impede o Quit): erro de compilacao = dialogo modal no Run = watchdog ($TimeoutSec) e o teste falha.
param([string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path, [int]$TimeoutSec = 300)
$ErrorActionPreference = 'Stop'
. (Join-Path $RepoRoot 'build\lib\Build-Common.ps1')
$fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }
function Wait-PidExit([int]$id, [int]$sec) { for ($i = 0; $i -lt ($sec * 4) -and (Get-Process -Id $id -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 250 }; return (-not (Get-Process -Id $id -ErrorAction SilentlyContinue)) }
$pptm = Join-Path $DistDir 'ThinQQ.ppt.test.pptm'
$coreTest = Join-Path $DistDir 'ThinQQ.core.test.xlam'
if (-not (Test-Path $pptm) -or -not (Test-Path $coreTest)) { Write-Host 'ERRO: monte antes (build.ps1 -Target all -Configuration Test e -Target ppt -Configuration Test)'; exit 1 }
$tmp = Join-Path $env:TEMP ('thinqq_test_ppt_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $tmp | Out-Null
$tpName = ''; $pp = $null; $xl = $null; $xl2 = $null; $wd = $null; $wdx = $null; $wdx2 = $null; $xlPid = 0; $xl2Pid = 0
$tp = $null; $pres = $null; $slide = $null; $core = $null; $wb = $null; $ws = $null; $vals = $null; $fk = $null; $shp = $null
$pres3 = $null; $pres4 = $null; $pres5 = $null; $shp3 = $null; $shp4 = $null; $shp5 = $null; $s5 = $null; $dwb = $null
$marker = Join-Path ([IO.Path]::GetTempPath()) "thinqq_watchdog_pptaddin_$PID.flag"
Remove-Item -LiteralPath $marker -Force -ErrorAction SilentlyContinue
try {
    $pp = New-PptApp
    $wd = Start-PidWatchdog $script:PptPid[0] $TimeoutSec $marker
    $tp = $pp.Presentations.Open($pptm, -1, 0, 0)     # somente leitura, sem janela
    $tpName = [string]$tp.Name
    $out = Join-Path $RepoRoot 'tests\results\ppt.tap'
    New-Item -ItemType Directory -Force (Split-Path $out) | Out-Null
    Remove-Item $out -ErrorAction SilentlyContinue
    $failed = $pp.Run("$tpName!modPptTestRunner.RunAllPptTests", $out)
    Write-Host '=== ppt (VBA) ==='
    $tap = @(Get-Content $out)
    $tap | ForEach-Object { Write-Host $_ }
    $plan = $tap | Where-Object { $_ -match '^1\.\.(\d+)$' } | Select-Object -First 1
    Check ($null -ne $plan -and [int]($plan -replace '^1\.\.', '') -gt 0) "suite VBA do suplemento com testes ($plan)"
    Check ([int]$failed -eq 0) "suite VBA do suplemento ($failed falha(s))"

    Write-Host '=== ppt (ponta a ponta) ==='
    $xl = New-ExcelApp
    $xlPid = Get-ExcelPid $xl
    $wdx = Start-ExcelWatchdog $xl $TimeoutSec
    $xl.EnableEvents = $true
    $core = $xl.Workbooks.Open($coreTest)
    $cn = [string]$core.Name
    $wb = $xl.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1); $ws.Name = 'Dados'
    $ws.Range('A1:C1').Value2 = @('Ano', 'Sul', 'Norte')
    $ws.Range('A2').Value2 = 'T1'; $ws.Range('B2').Value2 = 10; $ws.Range('C2').Value2 = 5
    $ws.Range('A3').Value2 = 'T2'; $ws.Range('B3').Value2 = 20; $ws.Range('C3').Value2 = 6
    $wb.SaveAs((Join-Path $tmp 'e2e.xlsx'), 51)
    $ws.Range('A1:C3').Select() | Out-Null
    [void]$xl.Run("'$cn'!Core_Insert", 'stacked')
    [void]$xl.Run("'$cn'!Core_AddAnnotation", 'cagr', 1, 2)
    $pres = $pp.Presentations.Add(0)
    $slide = $pres.Slides.Add(1, 12)
    [void]$xl.Run("'$cn'!Ppt_SetAppForTest", $pp)
    [void]$xl.Run("'$cn'!Ppt_PinTargetForTest", $slide, $false)
    [void]$xl.Run("'$cn'!Core_CopyToPowerPoint", 'chart')
    Check ($slide.Shapes.Count -eq 1) 'grafico colado'

    # (a) Excel injetado como o do usuario (ja com o core): sincroniza e NAO recebe Quit
    $ws.Range('B3').Value2 = 40
    [void]$pp.Run("$tpName!PptTqq_ResetTestHooks")
    [void]$pp.Run("$tpName!PptTqq_SetExcelForTest", $xl, '', $false)
    [void]$pp.Run("$tpName!PptTqq_ClearLastMsg")
    [void]$pp.Run("$tpName!PptTqq_Sync", $pres, $true)
    $msg = [string]$pp.Run("$tpName!PptTqq_LastResult")
    Check ($msg -like 'Gráficos do ThinQQ atualizados: 1.*') "(a) resumo do core pelo suplemento: $msg"
    Check ([string]$pp.Run("$tpName!PptTqq_LastMsg") -eq '') '(a) sucesso puro: sem caixa (1.11.0)'
    # (a1) estados pelo core de verdade (API 6), como na abertura: so o Excel ja aberto, sem abrir pasta fechada
    $st = [string]$pp.Run("$tpName!PptTqq_LinkStates", $pres, '', $false)
    Check ($st -match '^\d+\|[^|]+\|[^|]*e2e\.xlsx\|chart:\w*\|manual\|ok$') "(a1) PptTqq_LinkStates pelo core (API 6): $st"
    $vals = @($slide.Shapes.Item(1).Chart.SeriesCollection(1).Values)
    Check ($vals -contains 40) "(a) valor novo no slide ($($vals -join ','))"
    $alive = $false; try { $alive = ([string]$xl.Workbooks.Item('e2e.xlsx').Name -eq 'e2e.xlsx') } catch { }   # Workbooks.Count nao conta o core (suplemento)
    Check ($alive -and [bool](Get-Process -Id $xlPid -ErrorAction SilentlyContinue)) '(a) Excel injetado que nao foi criado pelo suplemento continua aberto (sem Quit)'
    # (a2) "Atualizar uma vez" (1.6.0) com o core de verdade (API 2): o mesmo grafico colado de novo no slide, com o MESMO nome da primeira
    # forma (copiar e colar no PowerPoint); so a forma "selecionada" (a segunda) e atualizada, a primeira fica como estava
    $wb.Save()
    $ws.ChartObjects(1).Activate() | Out-Null
    [void]$xl.Run("'$cn'!Core_CopyToPowerPoint", 'chart')
    Check ($slide.Shapes.Count -eq 2) "(a2) segundo grafico colado ($($slide.Shapes.Count))"
    $shp = $slide.Shapes.Item(2)
    $shp.Name = [string]$slide.Shapes.Item(1).Name
    $ws.Range('B3').Value2 = 45
    # Os valores lidos de um grafico vinculado acompanham a pasta aberta (medido na C3: a forma nao atualizada tambem mostra 45); o que prova
    # QUAL forma o core atualizou e a impressao digital TQQ_FP (dados + geometria), regravada so na forma sincronizada
    $fp1 = [string]$slide.Shapes.Item(1).Tags.Item('TQQ_FP')
    $fp2 = [string]$shp.Tags.Item('TQQ_FP')
    [void]$pp.Run("$tpName!E2e_UpdateShape", $shp)
    $msg = [string]$pp.Run("$tpName!PptTqq_LastResult")
    Check ($msg -like 'Gráficos do ThinQQ atualizados: 1.*') "(a2) Atualizar uma vez pelo core (API 2): $msg"
    Check ([string]$shp.Name -ne [string]$slide.Shapes.Item(1).Name) "(a2) nomes unicos antes de montar as entradas ($($shp.Name))"
    $vals = @($shp.Chart.SeriesCollection(1).Values)
    Check ($vals -contains 45) "(a2) valor novo na forma selecionada ($($vals -join ','))"
    Check ($fp1.Length -gt 0 -and $fp2.Length -gt 0) "(a2) antes: as duas formas tem TQQ_FP (da sincronizacao (a) e da colagem)"
    Check ([string]$shp.Tags.Item('TQQ_FP') -ne $fp2) '(a2) a forma selecionada foi sincronizada (TQQ_FP mudou com B3 = 45)'
    Check ([string]$slide.Shapes.Item(1).Tags.Item('TQQ_FP') -eq $fp1) '(a2) a outra forma, de mesmo nome, nao foi sincronizada (TQQ_FP igual)'
    $shp.Delete(); $shp = $null
    [void]$pp.Run("$tpName!PptTqq_ResetTestHooks")

    # (b) Excel novo, sem o core, injetado como criado pelo suplemento; pasta de versoes com 9.9.11 parcial (sem o arquivo), 9.9.10 com um
    # suplemento que nao e o core (nao responde ao Core_Ping: cai para a versao anterior) e 9.9.9 com o core de teste
    $ver = Join-Path $tmp 'versions'
    $coreFile = 'ThinQQ.core.test.xlam'
    foreach ($v in '9.9.9', '9.9.10', '9.9.11') { New-Item -ItemType Directory -Force (Join-Path $ver $v) | Out-Null }
    Copy-Item -LiteralPath $coreTest -Destination (Join-Path $ver "9.9.9\$coreFile")
    $fk = $xl.Workbooks.Add()
    $fk.SaveAs((Join-Path $ver '9.9.10\falso.xlam'), 55)   # xlOpenXMLAddIn; renomeado depois (o Excel nao abre dois livros com o mesmo nome)
    $fk.Close($false); $fk = $null
    Rename-Item -LiteralPath (Join-Path $ver '9.9.10\falso.xlam') -NewName $coreFile
    $ws.Range('B3').Value2 = 50
    $wb.Save(); $wb.Close($false); $wb = $null; $ws = $null   # pasta fechada: o core do Excel novo abre somente leitura
    $xl2 = New-ExcelApp
    $xl2Pid = Get-ExcelPid $xl2
    $wdx2 = Start-ExcelWatchdog $xl2 $TimeoutSec
    $xl2.EnableEvents = $true                             # como um CreateObject("Excel.Application") do suplemento
    [void]$pp.Run("$tpName!PptTqq_SetExcelForTest", $xl2, $ver, $true)
    [void]$pp.Run("$tpName!PptTqq_Sync", $pres, $true)
    $msg = [string]$pp.Run("$tpName!PptTqq_LastResult")
    $info = [string]$pp.Run("$tpName!PptTqq_CoreInfoForTest")
    $log = [string]$pp.Run("$tpName!PptTqq_LastLog")
    Check ($info -like "*\9.9.9\$coreFile|ro") "(b) core aberto da versao mais nova que responde, somente leitura: $info"
    Check ($log -like '*9.9.10*') "(b) a versao 9.9.10 que nao responde foi pulada: $log"
    Check ($msg -like 'Gráficos do ThinQQ atualizados: 1.*') "(b) resumo do core aberto pelo suplemento: $msg"
    $vals = @($slide.Shapes.Item(1).Chart.SeriesCollection(1).Values)
    Check ($vals -contains 50) "(b) valor novo no slide ($($vals -join ','))"
    [void]$pp.Run("$tpName!PptTqq_ResetTestHooks")
    Stop-ExcelWatchdog $wdx2
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl2) } catch { }
    $xl2 = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Check (Wait-PidExit $xl2Pid 20) "(b) o Excel criado pelo suplemento recebeu Quit e saiu (PID $xl2Pid)"

    # (c) 1.8.0 (G4): "Inserir gráfico" pelo suplemento com o core de verdade (API 3) no Excel injetado (o do usuario, com o core): o grafico
    # nasce vinculado a pasta de dados ao lado da apresentacao e a planilha dos dados abre; "Editar dados" reabre; salvar a apresentacao salva a
    # pasta de dados; apresentacao MOVIDA e RENOMEADA (pasta de dados junto, com o nome antigo): PptTqq_RelinkData reaponta o vinculo
    $g4 = Join-Path $tmp 'g4'; New-Item -ItemType Directory -Force $g4 | Out-Null
    $pres3 = $pp.Presentations.Add(0); [void]$pres3.Slides.Add(1, 12)
    $pres3.SaveAs((Join-Path $g4 'e2e-g4.pptx'))
    $dataName = 'e2e-g4 - dados ThinQQ.xlsx'
    $dataPath = Join-Path $g4 $dataName
    [void]$pp.Run("$tpName!PptTqq_ResetTestHooks")
    [void]$pp.Run("$tpName!PptTqq_SetExcelForTest", $xl, '', $false)
    [void]$pp.Run("$tpName!PptTqq_InsertChartIn", $pres3, $pres3.Slides.Item(1), 'stacked')
    $hint = [string]$pp.Run("$tpName!PptTqq_LastHintForTest")
    $trail = [string]$pp.Run("$tpName!PptTqq_TrailForTest")
    Check ($hint -like "Gráfico criado. Clique em OK e digite os dados*'$dataName'*" -and $trail -eq 'hint;edit;') "(c) inserido pelo suplemento, aviso de primeira vez antes de abrir os dados ($trail): $hint"
    Check ($pres3.Slides.Item(1).Shapes.Count -eq 1) "(c) uma forma no slide ($($pres3.Slides.Item(1).Shapes.Count))"
    $shp3 = $pres3.Slides.Item(1).Shapes.Item(1)
    Check ([string]$shp3.Tags.Item('TQQ_DATA') -eq '1' -and [string]$shp3.Tags.Item('TQQ_AUTO') -eq '' -and ([string]$pres3.Tags.Item('TQQ_LINKIDX')).Contains([string]$shp3.Name)) '(c) forma marcada TQQ_DATA, nasce manual (1.11.0) e entra no TQQ_LINKIDX'
    Check ([string]$shp3.LinkFormat.SourceFullName -eq $dataPath -and (Test-Path -LiteralPath $dataPath)) "(c) vinculo com a pasta de dados ao lado da apresentacao ($([string]$shp3.LinkFormat.SourceFullName))"
    $dwb = $xl.Workbooks.Item($dataName)
    Check ([string]$xl.ActiveWorkbook.Name -eq $dataName -and [string]$xl.ActiveSheet.Name -eq 'Gráfico 1') "(c) a planilha dos dados abriu ($([string]$xl.ActiveSheet.Name))"
    $dwb.Worksheets.Item(1).Activate() | Out-Null
    [void]$pp.Run("$tpName!E2e_EditShape", $shp3)
    $msg = [string]$pp.Run("$tpName!PptTqq_LastMsg")
    Check ($msg -like 'edit:*' -and [string]$xl.ActiveSheet.Name -eq 'Gráfico 1') "(c) Editar dados reabre a planilha do grafico ($msg)"
    $dwb.Worksheets.Item('Gráfico 1').Range('B2').Value2 = 77
    Check (-not [bool]$dwb.Saved) '(c) pre-condicao: a pasta de dados ficou alterada'
    [void]$pp.Run("$tpName!PptTqq_OnSave", $pres3)
    Check ([bool]$dwb.Saved) '(c) salvar a apresentacao salvou a pasta de dados alterada (PptTqq_OnSave -> Core_PptDataSave)'
    # Salvar como B na mesma pasta + inserir: segunda pasta de dados ao lado da primeira; a marca da apresentacao passa a ser a de B
    $pres3.SaveAs((Join-Path $g4 'e2e-g4b.pptx'))
    $dataNameB = 'e2e-g4b - dados ThinQQ.xlsx'
    [void]$pp.Run("$tpName!PptTqq_InsertChartIn", $pres3, $pres3.Slides.Item(1), 'line')
    Check ($pres3.Slides.Item(1).Shapes.Count -eq 2) "(c) Salvar como + inserir: segunda forma ($($pres3.Slides.Item(1).Shapes.Count))"
    Check ([string]$pres3.Slides.Item(1).Shapes.Item(2).LinkFormat.SourceFullName -eq (Join-Path $g4 $dataNameB) -and [string]$pres3.Tags.Item('TQQ_DATABOOK') -eq (Join-Path $g4 $dataNameB)) '(c) a segunda forma usa a pasta de dados de B (marca de B)'
    # nascem manuais (1.11.0): o usuario liga "Automatico" nas duas, para o reapontamento abaixo provar que o indice das automaticas acompanha
    [void]$xl.Run("'$cn'!PptSync_SetAuto", $pres3.Slides.Item(1).Shapes.Item(1), $true)
    [void]$xl.Run("'$cn'!PptSync_SetAuto", $pres3.Slides.Item(1).Shapes.Item(2), $true)
    $pres3.Save()
    $dwb.Close($false); $dwb = $null; $shp3 = $null
    $xl.Workbooks.Item($dataNameB).Close($true)
    $pres3.Close(); $pres3 = $null
    # pasta inteira movida, mas a pasta de dados de A ainda nao chegou (guardada com outro nome): a forma de A NAO pode ir para a de B (revisao I1)
    $g4b = Join-Path $tmp 'g4b'; New-Item -ItemType Directory -Force $g4b | Out-Null
    Move-Item -LiteralPath (Join-Path $g4 'e2e-g4b.pptx') -Destination (Join-Path $g4b 'e2e-g4b.pptx')
    Move-Item -LiteralPath (Join-Path $g4 $dataNameB) -Destination (Join-Path $g4b $dataNameB)
    Move-Item -LiteralPath $dataPath -Destination (Join-Path $g4b 'guardada.xlsx')
    Remove-Item -LiteralPath (Join-Path $g4 'e2e-g4.pptx') -Force
    $newData = Join-Path $g4b $dataName; $newDataB = Join-Path $g4b $dataNameB
    $pres4 = $pp.Presentations.Open((Join-Path $g4b 'e2e-g4b.pptx'), 0, 0, 0)
    $shp4 = $pres4.Slides.Item(1).Shapes.Item(1); $shp5 = $pres4.Slides.Item(1).Shapes.Item(2)
    Check ([string]$shp4.LinkFormat.SourceFullName -eq $dataPath) '(c) depois de mover: o vinculo ainda aponta para o caminho antigo (PG4)'
    $n = [int]$pp.Run("$tpName!PptTqq_RelinkData", $pres4)
    Check ($n -eq 1) "(c) so a forma de B reapontada ($n)"
    Check ([string]$shp5.LinkFormat.SourceFullName -eq $newDataB -and [string]$shp5.Tags.Item('TQQ_BOOK') -eq $newDataB -and [string]$pres4.Tags.Item('TQQ_DATABOOK') -eq $newDataB) "(c) forma de B: vinculo, TQQ_BOOK e TQQ_DATABOOK no caminho novo ($([string]$shp5.LinkFormat.SourceFullName))"
    Check ([string]$shp4.LinkFormat.SourceFullName -eq $dataPath) "(c) forma de A NAO foi para a pasta de dados de B ($([string]$shp4.LinkFormat.SourceFullName))"
    Rename-Item -LiteralPath (Join-Path $g4b 'guardada.xlsx') -NewName $dataName
    $n = [int]$pp.Run("$tpName!PptTqq_RelinkData", $pres4)
    Check ($n -eq 1 -and [string]$shp4.LinkFormat.SourceFullName -eq $newData -and [string]$shp4.Tags.Item('TQQ_BOOK') -eq $newData) "(c) a pasta de dados de A chegou: forma de A reapontada para ela ($n, $([string]$shp4.LinkFormat.SourceFullName))"
    Check ([string]$pres4.Tags.Item('TQQ_DATABOOK') -eq $newDataB) '(c) a marca continua a de B'
    $idx = [string]$pres4.Tags.Item('TQQ_AUTOIDX')
    Check ($idx.Contains($newData) -and $idx.Contains($newDataB) -and -not $idx.Contains($dataPath) -and -not $idx.Contains((Join-Path $g4 $dataNameB))) '(c) indice das automaticas acompanha'
    $lidx = [string]$pres4.Tags.Item('TQQ_LINKIDX')
    Check ($lidx.Contains($newData) -and $lidx.Contains($newDataB) -and -not $lidx.Contains($dataPath) -and -not $lidx.Contains((Join-Path $g4 $dataNameB))) "(c) TQQ_LINKIDX (todos os vinculos, 1.11.0) acompanha o reapontamento ($lidx)"
    $n = [int]$pp.Run("$tpName!PptTqq_RelinkData", $pres4)
    Check ($n -eq 0) "(c) terceira chamada: nada a reapontar ($n)"
    $pres4.Save()
    $shp4 = $null; $shp5 = $null
    $pres4.Close(); $pres4 = $null
    # apresentacao e pasta de dados RENOMEADAS juntas: a forma da pasta marcada vai para a pasta de dados do nome atual
    $dataNameR = 'e2e-g4-renomeada - dados ThinQQ.xlsx'
    Rename-Item -LiteralPath (Join-Path $g4b 'e2e-g4b.pptx') -NewName 'e2e-g4-renomeada.pptx'
    Rename-Item -LiteralPath $newDataB -NewName $dataNameR
    $pres5 = $pp.Presentations.Open((Join-Path $g4b 'e2e-g4-renomeada.pptx'), 0, 0, 0)
    $n = [int]$pp.Run("$tpName!PptTqq_RelinkData", $pres5)
    $s5 = $pres5.Slides.Item(1).Shapes
    Check ($n -eq 1 -and [string]$s5.Item(2).LinkFormat.SourceFullName -eq (Join-Path $g4b $dataNameR) -and [string]$s5.Item(1).LinkFormat.SourceFullName -eq $newData -and [string]$pres5.Tags.Item('TQQ_DATABOOK') -eq (Join-Path $g4b $dataNameR)) "(c) renomeadas juntas: so a forma da pasta marcada reapontada ($n)"
    $s5 = $null
    $pres5.Saved = -1; $pres5.Close(); $pres5 = $null
    [void]$pp.Run("$tpName!PptTqq_ResetTestHooks")

    [void]$xl.Run("'$cn'!Ppt_ResetTestHooks")
    $pres.Saved = -1; $pres.Close()
} catch {
    Write-Host "not ok - excecao: $($_.Exception.Message)"; $fail++
} finally {
    try { if ($pp -and $tpName) { [void]$pp.Run("$tpName!PptTqq_ResetTestHooks") } } catch { }
    Stop-ExcelWatchdog $wdx2
    if ($xl2) { Close-ExcelApp $xl2; $xl2 = $null }
    if ($xl2Pid -gt 0 -and -not (Wait-PidExit $xl2Pid 10)) { Stop-Process -Id $xl2Pid -Force -ErrorAction SilentlyContinue }   # Excel criado por este script
    # referencias COM soltas antes do Quit (senao o PowerPoint nao fecha e cai no encerramento pelo PID)
    $owned = @($tp, $pres, $pres3, $pres4, $pres5)
    $tp = $null; $pres = $null; $slide = $null; $core = $null; $wb = $null; $ws = $null; $vals = $null; $fk = $null; $shp = $null
    $pres3 = $null; $pres4 = $null; $pres5 = $null; $shp3 = $null; $shp4 = $null; $shp5 = $null; $s5 = $null; $dwb = $null
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Close-PptApp $pp -Owned $owned   # watchdog do PowerPoint ainda ativo
    $owned = $null; $pp = $null
    if ($wd -and $wd.Proc -and -not $wd.Proc.HasExited) { Stop-Process -Id $wd.Proc.Id -Force -ErrorAction SilentlyContinue }
    Write-Host "# PowerPoint fechado em $('{0:N1}' -f $sw.Elapsed.TotalSeconds) s"
    if ($xl) { Close-ExcelApp $xl; $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
    Stop-ExcelWatchdog $wdx
    $sw = [Diagnostics.Stopwatch]::StartNew()
    if ($xlPid -gt 0) { [void](Wait-PidExit $xlPid 15) }
    Write-Host "# Excel do teste (PID $xlPid) vivo depois do Quit: $([bool]($xlPid -gt 0 -and (Get-Process -Id $xlPid -ErrorAction SilentlyContinue))) ($('{0:N1}' -f $sw.Elapsed.TotalSeconds) s)"
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
if (Test-Path -LiteralPath $marker) { Remove-Item -LiteralPath $marker -Force -ErrorAction SilentlyContinue; Write-Host 'not ok - o PowerPoint do teste não respondeu (encerrado pelo watchdog)'; $fail++ }
foreach ($w in @($wdx, $wdx2)) { if ($w -and (Test-ExcelWatchdogFired $w)) { Remove-Item -LiteralPath $w.Marker -Force -ErrorAction SilentlyContinue; Write-Host 'not ok - um Excel do teste não respondeu (encerrado pelo watchdog)'; $fail++ } }
$left = @(Get-Process POWERPNT -ErrorAction SilentlyContinue)
$desc = ($left | ForEach-Object { "PID $($_.Id) desde $($_.StartTime.ToString('HH:mm:ss'))" }) -join '; '
Check ($left.Count -eq 0) "nenhum POWERPNT sobrou ($($left.Count)) $desc"
if ($fail -gt 0) { exit 1 } else { exit 0 }
