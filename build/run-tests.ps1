param(
    [ValidateSet('loader','core','all')] [string]$Target = 'all',
    [switch]$NoBuild,
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


$totalFailed = 0; $noTests = $false; $timedOut = $false
$xl = New-ExcelApp
$wd = Start-ExcelWatchdog $xl $TimeoutSec
try {
    $xl.EnableEvents = $true
    $core = $xl.Workbooks.Open((Join-Path $DistDir 'ThinQQ.core.test.xlam'))
    foreach ($t in $targets) {
        $wb = if ($t -eq 'core') { $core } else { $xl.Workbooks.Open((Join-Path $DistDir 'ThinQQ.test.xlam')) }
        $out = Join-Path $resultsDir "$t.tap"
        Remove-Item $out -ErrorAction SilentlyContinue
        $failed = $xl.Run("'$($wb.Name)'!modTestRunner.RunAllTests", $out)
        Write-Host "=== $t ==="
        $tap = Get-Content $out
        $tap
        $plan = $tap | Where-Object { $_ -match '^1\.\.(\d+)$' } | Select-Object -First 1
        if (-not $plan -or [int]($plan -replace '^1\.\.', '') -eq 0) {
            Write-Host "ERRO: nenhum teste encontrado em '$t'" -ForegroundColor Red
            $noTests = $true
        }
        $totalFailed += [int]$failed
    }
} catch {
    if (Test-ExcelWatchdogFired $wd) { $timedOut = $true } else { throw }
} finally {
    Stop-ExcelWatchdog $wd
    Close-ExcelApp $xl
    if ($server) { Stop-Process -Id $server.Id -Force -ErrorAction SilentlyContinue }
}
$elapsed = '{0:N1}s' -f $swTotal.Elapsed.TotalSeconds
if ($timedOut) { Write-ExcelTimeout $wd; Write-Host "Tempo: $elapsed"; exit 2 }
Write-Host "Tempo: $elapsed"
if ($noTests) { Write-Host 'nenhum teste encontrado' -ForegroundColor Red; exit 1 }
if ($totalFailed -gt 0) { Write-Host "FALHAS: $totalFailed" -ForegroundColor Red; exit 1 }
Write-Host 'Todos os testes passaram' -ForegroundColor Green
