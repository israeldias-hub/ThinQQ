# Instala o ThinQQ para o usuario atual (sem administrador). Compativel com Windows PowerShell 5.1.
param(
    [string]$Root = (Join-Path $env:LOCALAPPDATA 'ThinQQ'),
    [string]$RegistryRoot = 'HKCU:\Software',
    [string]$UpdateSource = '',
    [string]$PackageDir = $PSScriptRoot,
    [switch]$SkipExcelCheck
)
$ErrorActionPreference = 'Stop'
if (-not $SkipExcelCheck) {
    if (Get-Process excel -ErrorAction SilentlyContinue) { throw 'Feche o Excel antes de instalar o ThinQQ.' }
}

# --- Validacao do pacote (nada e copiado antes de tudo conferir) ---
foreach ($f in @('ThinQQ.xlam', 'ThinQQ.core.xlam', 'manifest.json')) {
    if (-not (Test-Path -LiteralPath (Join-Path $PackageDir $f))) { throw "Arquivo '$f' não encontrado em $PackageDir. Descompacte o ZIP inteiro antes de instalar." }
}
$manifestPath = Join-Path $PackageDir 'manifest.json'
try { $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json }
catch { throw 'manifest.json está inválido (não é um JSON legível).' }
$coreVer = [string]$manifest.coreVersion
if ($coreVer -notmatch '^\d+\.\d+\.\d+$') { throw "manifest.json com coreVersion inválida: '$coreVer'." }

function Test-PackageHash([string]$FileName, [string]$Field) {
    $expected = ''
    if ($manifest.PSObject.Properties[$Field]) { $expected = ([string]$manifest.$Field).Trim().ToLower() }
    if ($expected -notmatch '^[0-9a-f]{64}$') { throw "manifest.json sem $Field válido; o pacote não pode ser verificado." }
    $actual = (Get-FileHash -LiteralPath (Join-Path $PackageDir $FileName) -Algorithm SHA256).Hash.ToLower()
    if ($actual -ne $expected) { throw "Pacote corrompido ou adulterado: o SHA-256 de $FileName não confere com o manifest.json. Baixe o ZIP novamente." }
}
Test-PackageHash 'ThinQQ.core.xlam' 'coreSha256'
Test-PackageHash 'ThinQQ.xlam' 'loaderSha256'

# --- Copia ---
$verDir = Join-Path $Root "versions\$coreVer"
foreach ($d in @($verDir, (Join-Path $Root 'tmp'), (Join-Path $Root 'logs'))) { New-Item -ItemType Directory -Force $d | Out-Null }
Copy-Item -LiteralPath (Join-Path $PackageDir 'ThinQQ.xlam') -Destination $Root -Force
Copy-Item -LiteralPath (Join-Path $PackageDir 'ThinQQ.core.xlam') -Destination $verDir -Force
if (Test-Path -LiteralPath (Join-Path $PackageDir 'tutorial.html')) { Copy-Item -LiteralPath (Join-Path $PackageDir 'tutorial.html') -Destination $verDir -Force }
Get-ChildItem -LiteralPath $Root -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

# --- Local confiavel (macros permitidas sem aviso) ---
$tl = Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ'
New-Item -Path $tl -Force | Out-Null
Set-ItemProperty -Path $tl -Name Path -Value ($Root.TrimEnd('\') + '\')
Set-ItemProperty -Path $tl -Name AllowSubfolders -Value 1 -Type DWord
Set-ItemProperty -Path $tl -Name Description -Value 'ThinQQ add-in'

# --- Registro do suplemento: OPEN, OPEN1, OPEN2... (sem lacunas, sem duplicar) ---
$opt = Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Options'
New-Item -Path $opt -Force | Out-Null
$loaderPath = Join-Path $Root 'ThinQQ.xlam'
$existing = @((Get-ItemProperty $opt).PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' })
$already = @($existing | Where-Object { [string]$_.Value -like '*ThinQQ.xlam*' })
if ($already.Count -eq 0) {
    $n = 0
    $names = @($existing | ForEach-Object { $_.Name })
    while ($true) {
        $candidate = 'OPEN'
        if ($n -gt 0) { $candidate = "OPEN$n" }
        if ($names -notcontains $candidate) { break }
        $n++
    }
    Set-ItemProperty -Path $opt -Name $candidate -Value ('/R "' + $loaderPath + '"')
}

# --- Fonte de atualizacao: parametro > valor ja existente > manifestUrl ---
$up = Join-Path $RegistryRoot 'VB and VBA Program Settings\ThinQQ\Updater'
$current = ''
if (Test-Path $up) {
    $p = (Get-ItemProperty $up).PSObject.Properties['UpdateSource']
    if ($p) { $current = [string]$p.Value }
}
$source = $UpdateSource
if (-not $source) { $source = $current }
if (-not $source -and $manifest.PSObject.Properties['manifestUrl']) { $source = [string]$manifest.manifestUrl }
if ($source) {
    New-Item -Path $up -Force | Out-Null
    Set-ItemProperty -Path $up -Name UpdateSource -Value $source
}
Write-Host "ThinQQ $coreVer instalado em $Root. Abra o Excel e procure a aba 'ThinQQ'."
