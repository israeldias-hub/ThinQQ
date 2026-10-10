param([ValidateSet('loader','core','all')] [string]$Target = 'all', [switch]$NoBuild)
. "$PSScriptRoot\lib\Build-Common.ps1"
if (-not $NoBuild) { & "$PSScriptRoot\build.ps1" -Target all -Configuration Test }

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

$totalFailed = 0
$xl = New-ExcelApp
$xl.EnableEvents = $true
try {
    $core = $xl.Workbooks.Open((Join-Path $DistDir 'ThinQQ.core.test.xlam'))
    foreach ($t in $targets) {
        $wb = if ($t -eq 'core') { $core } else { $xl.Workbooks.Open((Join-Path $DistDir 'ThinQQ.test.xlam')) }
        $out = Join-Path $resultsDir "$t.tap"
        Remove-Item $out -ErrorAction SilentlyContinue
        $failed = $xl.Run("'$($wb.Name)'!modTestRunner.RunAllTests", $out)
        Write-Host "=== $t ==="
        Get-Content $out
        $totalFailed += [int]$failed
    }
} finally {
    Close-ExcelApp $xl
    if ($server) { Stop-Process -Id $server.Id -Force -ErrorAction SilentlyContinue }
}
if ($totalFailed -gt 0) { Write-Host "FALHAS: $totalFailed" -ForegroundColor Red; exit 1 }
Write-Host 'Todos os testes passaram' -ForegroundColor Green
