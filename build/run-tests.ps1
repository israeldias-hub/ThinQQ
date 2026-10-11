param(
    [ValidateSet('loader','core','all')] [string]$Target = 'all',
    [switch]$NoBuild,
    [switch]$LoaderSwap,     # tambem roda tests/ps/Test-LoaderSwap.ps1 (Excel real; monta Release em dist; ~30 s)
    [string]$Filter = '',    # so roda testes cujo 'Modulo.Teste' contem este texto (iteracao rapida; pula as suites PowerShell)
    [int]$TimeoutSec = 600   # watchdog da execução no Excel: mata o Excel iniciado aqui se não terminar (saída 2)
)
. "$PSScriptRoot\lib\Build-Common.ps1"
$swTotal = [System.Diagnostics.Stopwatch]::StartNew()

# Limpeza conservadora do %TEMP%: só pastas thinqq_test_* com mais de 1 dia e modBuildInfo_*.bas
$cut = (Get-Date).AddDays(-1)
$stale = @(Get-ChildItem -LiteralPath $env:TEMP -Directory -Filter 'thinqq_test_*' -ErrorAction SilentlyContinue |
    Where-Object { $_.LastWriteTime -lt $cut -and -not $_.Attributes.HasFlag([IO.FileAttributes]::ReparsePoint) })
$removed = 0
foreach ($d in $stale) { Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction SilentlyContinue; if (-not (Test-Path -LiteralPath $d.FullName)) { $removed++ } }
Get-ChildItem -Path (Join-Path $env:TEMP '*') -Include 'modBuildInfo_*.bas','modTestDispatch_*.bas' -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue
Write-Host "Limpeza: $removed pasta(s) thinqq_test_* antigas removidas de %TEMP%"

if (-not $NoBuild) {
    $global:LASTEXITCODE = 0
    & "$PSScriptRoot\build.ps1" -Target all -Configuration Test
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}

$resultsDir = Join-Path $RepoRoot 'tests\results'
New-Item -ItemType Directory -Force $resultsDir | Out-Null
Remove-Item (Join-Path $resultsDir 'http-base.txt') -ErrorAction SilentlyContinue
$targets = if ($Target -eq 'all') { @('core','loader') } else { @($Target) }

# Fixtures de release para o teste do updater (core de teste + manifesto apontando para o servidor local)
$relDir = Join-Path $RepoRoot 'tests\fixtures\release'
New-Item -ItemType Directory -Force $relDir | Out-Null
Copy-Item (Join-Path $DistDir 'ThinQQ.core.test.xlam') $relDir -Force
$sha = (Get-FileHash (Join-Path $relDir 'ThinQQ.core.test.xlam') -Algorithm SHA256).Hash.ToLower()
@{ schemaVersion = 1; coreVersion = '9.9.9'; coreUrl = 'http://127.0.0.1:18080/release/ThinQQ.core.test.xlam'
   coreSha256 = $sha; loaderVersion = '0.0.1'; loaderUrl = ''; loaderSha256 = ''; tutorialUrl = ''
   publishedAt = '2099-01-01'; notes = "• fixture de teste" } | ConvertTo-Json | Set-Content (Join-Path $relDir 'manifest.json') -Encoding utf8

$server = $null
$node = Get-Command node -ErrorAction SilentlyContinue
if ($node) {
    $server = Start-Process -FilePath $node.Source -ArgumentList @((Join-Path $RepoRoot 'tests\http\server.js'), '18080', (Join-Path $RepoRoot 'tests\fixtures')) -PassThru -WindowStyle Hidden
    Start-Sleep -Milliseconds 800
    Set-Content (Join-Path $resultsDir 'http-base.txt') 'http://127.0.0.1:18080/' -NoNewline
}


if ($Filter) { $env:TQQ_TEST_FILTER = $Filter } else { Remove-Item Env:\TQQ_TEST_FILTER -ErrorAction SilentlyContinue }
$totalFailed = 0; $noTests = $false; $timedOut = $false; $compileFailed = $false
$xl = New-ExcelApp
$wd = Start-ExcelWatchdog $xl $TimeoutSec
try {
    $xl.EnableEvents = $true
    $core = $xl.Workbooks.Open((Join-Path $DistDir 'ThinQQ.core.test.xlam'))
    $loaderWb = $null
    if ($targets -contains 'loader') { $loaderWb = $xl.Workbooks.Open((Join-Path $DistDir 'ThinQQ.test.xlam')) }
    # Compilação completa de cada projeto (pega erro de sintaxe em procedimento nunca chamado pelos testes).
    foreach ($pair in @(@('core', $core), @('loader', $loaderWb))) {
        if ($null -eq $pair[1] -or $targets -notcontains $pair[0]) { continue }
        $err = Invoke-VbaCompile $xl $pair[1]
        if ($err) {
            Write-Host "ERRO: erro de compilação VBA no projeto '$($pair[0])' (módulo:linha = $err)" -ForegroundColor Red
            $compileFailed = $true
            throw 'erro de compilação VBA'
        }
    }
    foreach ($t in $targets) {
        $wb = if ($t -eq 'core') { $core } else { $loaderWb }
        $out = Join-Path $resultsDir "$t.tap"
        Remove-Item $out -ErrorAction SilentlyContinue
        $failed = $xl.Run("'$($wb.Name)'!modTestRunner.RunAllTests", $out)
        Write-Host "=== $t ==="
        $tap = Get-Content $out
        $tap | Where-Object { $_ -notmatch '^# time ' }        # as linhas "# time Modulo.Teste segundos" ficam so no arquivo .tap
        $slow = @($tap | Where-Object { $_ -match '^# time \S+ [0-9.]+$' } | ForEach-Object {
            $p = $_ -split ' '
            [pscustomobject]@{ Name = $p[2]; Sec = [double]::Parse($p[3], [Globalization.CultureInfo]::InvariantCulture) } } | Sort-Object Sec -Descending)
        if ($slow.Count -gt 0) {
            Write-Host "--- $t : os 15 testes mais lentos (de $($slow.Count); soma $('{0:N1}' -f ($slow | Measure-Object Sec -Sum).Sum) s) ---"
            $slow | Select-Object -First 15 | ForEach-Object { Write-Host ('{0,8:N2} s  {1}' -f $_.Sec, $_.Name) }
        }
        $plan = $tap | Where-Object { $_ -match '^1\.\.(\d+)$' } | Select-Object -First 1
        if (-not $plan -or [int]($plan -replace '^1\.\.', '') -eq 0) {
            Write-Host "ERRO: nenhum teste encontrado em '$t'" -ForegroundColor Red
            $noTests = $true
        }
        $totalFailed += [int]$failed
    }
} catch {
    if (Test-ExcelWatchdogFired $wd) { $timedOut = $true } elseif (-not $compileFailed) { throw }
} finally {
    Stop-ExcelWatchdog $wd
    if ($compileFailed) {
        # diálogo modal pendente impede Quit: encerra só o processo iniciado por este script
        $xp = Get-ExcelPid $xl
        if ($xp -gt 0) { Stop-Process -Id $xp -Force -ErrorAction SilentlyContinue }
    } else { Close-ExcelApp $xl }
    if ($server) { Stop-Process -Id $server.Id -Force -ErrorAction SilentlyContinue }
}
if ($Filter) {
    Write-Host "Filtro '$Filter' ativo: suites PowerShell puladas"
    if ($compileFailed) { exit 3 }
    if ($timedOut) { exit 2 }
    if ($totalFailed -gt 0) { exit 1 } else { exit 0 }
}
if (-not $timedOut) {
    Write-Host '=== vba-lint ==='
    $global:LASTEXITCODE = 0
    & (Join-Path $RepoRoot 'tests\ps\Test-VbaLint.ps1') -RepoRoot $RepoRoot
    if ($LASTEXITCODE -ne 0) { $totalFailed += 1 }
}
if (-not $timedOut -and $targets -contains 'loader') {
    Write-Host '=== ribbon-pacote ==='
    $global:LASTEXITCODE = 0
    & (Join-Path $RepoRoot 'tests\ps\Test-Ribbon.ps1') -RepoRoot $RepoRoot
    if ($LASTEXITCODE -ne 0) { $totalFailed += 1 }
}
if (-not $timedOut) {
    Write-Host '=== documentacao (tutorial, README, rotulos da faixa) ==='
    $global:LASTEXITCODE = 0
    & (Join-Path $RepoRoot 'tests\ps\Test-Docs.ps1') -RepoRoot $RepoRoot
    if ($LASTEXITCODE -ne 0) { $totalFailed += 1 }
}
if (-not $timedOut) {
    Write-Host '=== privacidade (arquivos versionados: ids, caminhos de perfil, e-mail corporativo) ==='
    $global:LASTEXITCODE = 0
    & (Join-Path $RepoRoot 'tests\ps\Test-Privacy.ps1') -RepoRoot $RepoRoot
    if ($LASTEXITCODE -ne 0) { $totalFailed += 1 }
}
if (-not $timedOut) {
    Write-Host '=== instalador (pwsh e Windows PowerShell 5.1) + guarda de release ==='
    $psExes = @((Get-Process -Id $PID).Path, (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'))
    foreach ($exe in $psExes) {
        foreach ($tst in @('Test-Installer.ps1', 'Test-ReleaseGuard.ps1')) {
            $global:LASTEXITCODE = 0
            & $exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $RepoRoot "tests\ps\$tst")
            if ($LASTEXITCODE -ne 0) { Write-Host "FALHOU: $tst em $exe" -ForegroundColor Red; $totalFailed += 1 }
        }
    }
}
if (-not $timedOut -and $LoaderSwap) {
    Write-Host '=== loader-swap ==='
    $global:LASTEXITCODE = 0
    & (Join-Path $PSScriptRoot 'build.ps1') -Target all -Configuration Release
    if ($LASTEXITCODE -eq 0) { & (Join-Path $RepoRoot 'tests\ps\Test-LoaderSwap.ps1') -RepoRoot $RepoRoot }
    if ($LASTEXITCODE -ne 0) { $totalFailed += 1 }
}
$elapsed = '{0:N1}s' -f $swTotal.Elapsed.TotalSeconds
if ($compileFailed) { Write-Host "Tempo: $elapsed"; exit 3 }
if ($timedOut) { Write-ExcelTimeout $wd; Write-Host "Tempo: $elapsed"; exit 2 }
Write-Host "Tempo: $elapsed"
if ($noTests) { Write-Host 'nenhum teste encontrado' -ForegroundColor Red; exit 1 }
if ($totalFailed -gt 0) { Write-Host "FALHAS: $totalFailed" -ForegroundColor Red; exit 1 }
Write-Host 'Todos os testes passaram' -ForegroundColor Green
