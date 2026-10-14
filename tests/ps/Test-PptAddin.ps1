# Suplemento do PowerPoint: roda a suite VBA do projeto de teste (.pptm) e um teste de ponta a ponta (Excel proprio com o core de teste +
# PowerPoint proprio): colar vinculado, mudar o dado, PptTqq_Sync com o Excel injetado. So roda sem POWERPNT aberto (instancia unica).
# O PowerPoint criado aqui e encerrado pelo PID no fim (Close-PptApp), inclusive depois de erro ou de watchdog. Sem checagem de compilacao
# pelo VBE (o VBE do PowerPoint nao lista o projeto aberto sem janela e mexer nele impede o Quit): erro de compilacao = dialogo modal no
# Run = watchdog ($TimeoutSec) encerra o PowerPoint pelo PID e o teste falha.
param([string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path, [int]$TimeoutSec = 300)
$ErrorActionPreference = 'Stop'
. (Join-Path $RepoRoot 'build\lib\Build-Common.ps1')
$fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }
$pptm = Join-Path $DistDir 'ThinQQ.ppt.test.pptm'
$coreTest = Join-Path $DistDir 'ThinQQ.core.test.xlam'
if (-not (Test-Path $pptm) -or -not (Test-Path $coreTest)) { Write-Host 'ERRO: monte antes (build.ps1 -Target all -Configuration Test e -Target ppt -Configuration Test)'; exit 1 }
$tmp = Join-Path $env:TEMP ('thinqq_test_ppt_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $tmp | Out-Null
$pp = $null; $xl = $null; $wd = $null; $wdx = $null; $xlPid = 0
$tp = $null; $pres = $null; $slide = $null; $core = $null; $wb = $null; $ws = $null; $vals = $null
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
    $ws.Range('B3').Value2 = 40
    [void]$pp.Run("$tpName!PptTqq_SetExcelForTest", $xl, '')
    [void]$pp.Run("$tpName!PptTqq_Sync", $pres, $true)
    $msg = [string]$pp.Run("$tpName!PptTqq_LastMsg")
    Check ($msg -like 'Gráficos do ThinQQ atualizados: 1.*') "resumo do core pelo suplemento: $msg"
    $vals = @($slide.Shapes.Item(1).Chart.SeriesCollection(1).Values)
    Check ($vals -contains 40) "valor novo no slide ($($vals -join ','))"
    [void]$pp.Run("$tpName!PptTqq_SetExcelForTest", $null, '')
    [void]$xl.Run("'$cn'!Ppt_ResetTestHooks")
    $pres.Saved = -1; $pres.Close()
    $wb.Close($false)
} catch {
    Write-Host "not ok - excecao: $($_.Exception.Message)"; $fail++
} finally {
    if ($wd -and $wd.Proc -and -not $wd.Proc.HasExited) { Stop-Process -Id $wd.Proc.Id -Force -ErrorAction SilentlyContinue }
    Stop-ExcelWatchdog $wdx
    # referencias COM soltas antes do Quit (senao o PowerPoint nao fecha e cai no encerramento pelo PID)
    $tp = $null; $pres = $null; $slide = $null; $core = $null; $wb = $null; $ws = $null; $vals = $null
    $sw = [Diagnostics.Stopwatch]::StartNew()
    Close-PptApp $pp
    $pp = $null
    Write-Host "# PowerPoint fechado em $('{0:N1}' -f $sw.Elapsed.TotalSeconds) s"
    if ($xl) { Close-ExcelApp $xl; $xl = $null; [GC]::Collect(); [GC]::WaitForPendingFinalizers() }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    if ($xlPid -gt 0) { for ($i = 0; $i -lt 60 -and (Get-Process -Id $xlPid -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 250 } }
    Write-Host "# Excel do teste (PID $xlPid) vivo depois do Quit: $([bool]($xlPid -gt 0 -and (Get-Process -Id $xlPid -ErrorAction SilentlyContinue))) ($('{0:N1}' -f $sw.Elapsed.TotalSeconds) s)"
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
if (Test-Path -LiteralPath $marker) { Remove-Item -LiteralPath $marker -Force -ErrorAction SilentlyContinue; Write-Host 'not ok - o PowerPoint do teste não respondeu (encerrado pelo watchdog)'; $fail++ }
if ($wdx -and (Test-ExcelWatchdogFired $wdx)) { Remove-Item -LiteralPath $wdx.Marker -Force -ErrorAction SilentlyContinue; Write-Host 'not ok - o Excel do teste não respondeu (encerrado pelo watchdog)'; $fail++ }
$left = @(Get-Process POWERPNT -ErrorAction SilentlyContinue)
$desc = ($left | ForEach-Object { "PID $($_.Id) desde $($_.StartTime.ToString('HH:mm:ss'))" }) -join '; '
Check ($left.Count -eq 0) "nenhum POWERPNT sobrou ($($left.Count)) $desc"
if ($fail -gt 0) { exit 1 } else { exit 0 }
