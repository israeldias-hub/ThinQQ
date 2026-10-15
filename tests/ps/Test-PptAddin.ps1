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
    [void]$pp.Run("$tpName!PptTqq_Sync", $pres, $true)
    $msg = [string]$pp.Run("$tpName!PptTqq_LastMsg")
    Check ($msg -like 'Gráficos do ThinQQ atualizados: 1.*') "(a) resumo do core pelo suplemento: $msg"
    $vals = @($slide.Shapes.Item(1).Chart.SeriesCollection(1).Values)
    Check ($vals -contains 40) "(a) valor novo no slide ($($vals -join ','))"
    $alive = $false; try { $alive = ([string]$xl.Workbooks.Item('e2e.xlsx').Name -eq 'e2e.xlsx') } catch { }   # Workbooks.Count nao conta o core (suplemento)
    Check ($alive -and [bool](Get-Process -Id $xlPid -ErrorAction SilentlyContinue)) '(a) Excel injetado que nao foi criado pelo suplemento continua aberto (sem Quit)'
    # (a2) "Atualizar uma vez" (1.6.0) com o core de verdade (API 2): Core_PptSyncShapes so na forma listada
    $ws.Range('B3').Value2 = 45
    $shp = $slide.Shapes.Item(1)
    $entry = [string]$pp.Run("$tpName!PptTqq_IdxEntry", [int]$slide.SlideID, [string]$shp.Name, '')
    [void]$pp.Run("$tpName!PptTqq_SyncEntries", $pres, $entry, $true, $false)
    $msg = [string]$pp.Run("$tpName!PptTqq_LastMsg")
    Check ($msg -like 'Gráficos do ThinQQ atualizados: 1.*') "(a2) Atualizar uma vez pelo core (API 2): $msg"
    $vals = @($slide.Shapes.Item(1).Chart.SeriesCollection(1).Values)
    Check ($vals -contains 45) "(a2) valor novo no slide ($($vals -join ','))"
    $shp = $null
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
    $msg = [string]$pp.Run("$tpName!PptTqq_LastMsg")
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
    $owned = @($tp, $pres)
    $tp = $null; $pres = $null; $slide = $null; $core = $null; $wb = $null; $ws = $null; $vals = $null; $fk = $null; $shp = $null
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
