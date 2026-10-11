# Instala o ThinQQ para o usuario atual (sem administrador). Compativel com Windows PowerShell 5.1.
param(
    [string]$Root = (Join-Path $env:LOCALAPPDATA 'ThinQQ'),
    [string]$RegistryRoot = 'HKCU:\Software',
    [string]$UpdateSource = '',
    [string]$PackageDir = $PSScriptRoot,
    [switch]$SkipExcelCheck,
    [switch]$Force,
    [switch]$SimulateFailureForTest   # uso interno dos testes: falha depois da copia para validar o desfazer
)
$ErrorActionPreference = 'Stop'

# --- Protecoes de caminho (antes de qualquer alteracao) ---
$Root = [IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Root)).TrimEnd('\')
if ([IO.Path]::GetFileName($Root) -ne 'ThinQQ' -or $Root -ieq $env:USERPROFILE.TrimEnd('\')) {
    throw "Pasta de instalação recusada: '$Root'. O nome da pasta precisa ser 'ThinQQ' (padrão: %LOCALAPPDATA%\ThinQQ)."
}
$RegistryRoot = $RegistryRoot.TrimEnd('\')
if ($RegistryRoot -ne 'HKCU:\Software' -and $RegistryRoot -notmatch '^HKCU:\\Software\\ThinQQ_Test_[A-Za-z0-9_]+$') {
    throw "Chave de registro recusada: '$RegistryRoot'. Só HKCU:\Software é permitido."
}
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

# Versao do loader gravada no arquivo (propriedade Comentarios: "ThinQQ loader x.y.z (Release)"); $null se ilegivel
function Get-LoaderFileVersion([string]$Path) {
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $z = [IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $e = $z.GetEntry('docProps/core.xml')
            if (-not $e) { return $null }
            $sr = New-Object IO.StreamReader($e.Open())
            try { $xml = $sr.ReadToEnd() } finally { $sr.Dispose() }
            if ($xml -match 'ThinQQ loader (\d+\.\d+\.\d+)') { return [version]$Matches[1] }
        } finally { $z.Dispose() }
    } catch { }
    return $null
}

# Cria a chave (e so as chaves que faltam) sem -Force: uma chave existente nunca e recriada nem apagada.
function Initialize-RegKey([string]$Path) {
    if (Test-Path -LiteralPath $Path) { return }
    Initialize-RegKey (Split-Path -Path $Path -Parent)
    [void](New-Item -Path $Path)
}
$loaderDest = Join-Path $Root 'ThinQQ.xlam'
$verDir = Join-Path $Root "versions\$coreVer"
$hadLoader = Test-Path -LiteralPath $loaderDest
if ($hadLoader -and -not $Force -and $manifest.PSObject.Properties['loaderVersion'] -and ([string]$manifest.loaderVersion) -match '^\d+\.\d+\.\d+$') {
    $installed = Get-LoaderFileVersion $loaderDest
    if ($installed -and $installed -gt [version]$manifest.loaderVersion) {
        throw "O loader instalado ($installed) é mais novo que o do pacote ($($manifest.loaderVersion)). Use -Force para sobrescrever mesmo assim."
    }
}

# --- Copia em area temporaria, depois move (com desfazer em caso de falha) ---
$rootCreated = -not (Test-Path -LiteralPath $Root)
$verDirCreated = -not (Test-Path -LiteralPath $verDir)
$stage = Join-Path $Root '_staging'
$bak = $loaderDest + '.bak'
$undo = New-Object System.Collections.ArrayList
$loaderMoved = $false; $coreMoved = $false; $tutMoved = $false
$hadCore = Test-Path -LiteralPath (Join-Path $verDir 'ThinQQ.core.xlam')
$hadTut = Test-Path -LiteralPath (Join-Path $verDir 'tutorial.html')
$undoDir = Join-Path $stage '_undo'
try {
    if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force }
    New-Item -ItemType Directory -Path $stage -Force | Out-Null
    foreach ($f in @('ThinQQ.xlam', 'ThinQQ.core.xlam', 'tutorial.html')) {
        if (Test-Path -LiteralPath (Join-Path $PackageDir $f)) {
            Copy-Item -LiteralPath (Join-Path $PackageDir $f) -Destination $stage -Force
            Unblock-File -LiteralPath (Join-Path $stage $f) -ErrorAction SilentlyContinue
        }
    }
    foreach ($d in @($verDir, (Join-Path $Root 'tmp'), (Join-Path $Root 'logs'), $undoDir)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }

    # guarda copias para o desfazer; o loader antigo vira ThinQQ.xlam.bak (so o ultimo e mantido)
    if ($hadLoader) { Copy-Item -LiteralPath $loaderDest -Destination (Join-Path $undoDir 'loader') -Force; Copy-Item -LiteralPath $loaderDest -Destination $bak -Force }
    if ($hadCore) { Copy-Item -LiteralPath (Join-Path $verDir 'ThinQQ.core.xlam') -Destination (Join-Path $undoDir 'core') -Force }
    if ($hadTut) { Copy-Item -LiteralPath (Join-Path $verDir 'tutorial.html') -Destination (Join-Path $undoDir 'tut') -Force }

    Move-Item -LiteralPath (Join-Path $stage 'ThinQQ.xlam') -Destination $loaderDest -Force; $loaderMoved = $true
    Move-Item -LiteralPath (Join-Path $stage 'ThinQQ.core.xlam') -Destination $verDir -Force; $coreMoved = $true
    if (Test-Path -LiteralPath (Join-Path $stage 'tutorial.html')) { Move-Item -LiteralPath (Join-Path $stage 'tutorial.html') -Destination $verDir -Force; $tutMoved = $true }
    if ($SimulateFailureForTest) { throw 'Falha simulada (teste).' }

    # --- Local confiavel (macros permitidas sem aviso) ---
    $tl = Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ'
    if (-not (Test-Path -LiteralPath $tl)) {
        Initialize-RegKey $tl
        [void]$undo.Add({ Remove-Item -LiteralPath $tl -Recurse -Force -ErrorAction SilentlyContinue }.GetNewClosure())
    }
    Set-ItemProperty -LiteralPath $tl -Name Path -Value ($Root + '\')
    Set-ItemProperty -LiteralPath $tl -Name AllowSubfolders -Value 1 -Type DWord
    Set-ItemProperty -LiteralPath $tl -Name Description -Value 'ThinQQ add-in'

    # --- Registro do suplemento: OPEN, OPEN1, OPEN2... (sem lacunas, sem duplicar) ---
    $opt = Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Options'
    Initialize-RegKey $opt
    $newValue = '/R "' + $loaderDest + '"'
    $existing = @((Get-ItemProperty -LiteralPath $opt).PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' })
    # entrada do ThinQQ = nome de arquivo exatamente ThinQQ.xlam (nao confunde com MyThinQQ.xlam)
    $mine = @($existing | Where-Object { ([string]$_.Value) -match '(^|[\\/"\s])ThinQQ\.xlam"?\s*$' })
    if ($mine.Count -gt 0) {
        $nm = $mine[0].Name; $oldVal = [string]$mine[0].Value
        if ($oldVal -ne $newValue) {
            Set-ItemProperty -LiteralPath $opt -Name $nm -Value $newValue
            [void]$undo.Add({ Set-ItemProperty -LiteralPath $opt -Name $nm -Value $oldVal }.GetNewClosure())
        }
    } else {
        $n = 0
        $names = @($existing | ForEach-Object { $_.Name })
        while ($true) {
            $candidate = 'OPEN'
            if ($n -gt 0) { $candidate = "OPEN$n" }
            if ($names -notcontains $candidate) { break }
            $n++
        }
        Set-ItemProperty -LiteralPath $opt -Name $candidate -Value $newValue
        [void]$undo.Add({ Remove-ItemProperty -LiteralPath $opt -Name $candidate -ErrorAction SilentlyContinue }.GetNewClosure())
    }

    # --- Fonte de atualizacao: parametro > valor ja existente > manifestUrl ---
    $up = Join-Path $RegistryRoot 'VB and VBA Program Settings\ThinQQ\Updater'
    $current = $null
    if (Test-Path -LiteralPath $up) {
        $p = (Get-ItemProperty -LiteralPath $up).PSObject.Properties['UpdateSource']
        if ($p) { $current = [string]$p.Value }
    }
    $source = $UpdateSource
    if (-not $source) { $source = $current }
    if (-not $source -and $manifest.PSObject.Properties['manifestUrl']) { $source = [string]$manifest.manifestUrl }
    if ($source -and $source -ne $current) {
        Initialize-RegKey $up
        Set-ItemProperty -LiteralPath $up -Name UpdateSource -Value $source
        [void]$undo.Add({ if ($null -ne $current) { Set-ItemProperty -LiteralPath $up -Name UpdateSource -Value $current } else { Remove-ItemProperty -LiteralPath $up -Name UpdateSource -ErrorAction SilentlyContinue } }.GetNewClosure())
    }
} catch {
    $err = $_
    Write-Host 'Falha na instalação; desfazendo as alterações desta execução...' -ForegroundColor Yellow
    for ($i = $undo.Count - 1; $i -ge 0; $i--) { try { & $undo[$i] } catch { } }
    try {
        if ($loaderMoved) {
            if ($hadLoader) { Copy-Item -LiteralPath (Join-Path $undoDir 'loader') -Destination $loaderDest -Force }
            else { Remove-Item -LiteralPath $loaderDest -Force -ErrorAction SilentlyContinue }
        }
        if ($coreMoved) {
            if ($hadCore) { Copy-Item -LiteralPath (Join-Path $undoDir 'core') -Destination (Join-Path $verDir 'ThinQQ.core.xlam') -Force }
            else { Remove-Item -LiteralPath (Join-Path $verDir 'ThinQQ.core.xlam') -Force -ErrorAction SilentlyContinue }
        }
        if ($tutMoved) {
            if ($hadTut) { Copy-Item -LiteralPath (Join-Path $undoDir 'tut') -Destination (Join-Path $verDir 'tutorial.html') -Force }
            else { Remove-Item -LiteralPath (Join-Path $verDir 'tutorial.html') -Force -ErrorAction SilentlyContinue }
        }
        if (-not $hadLoader) { Remove-Item -LiteralPath $bak -Force -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue }
        if ($verDirCreated -and (Test-Path -LiteralPath $verDir) -and -not (Get-ChildItem -LiteralPath $verDir -Force)) { Remove-Item -LiteralPath $verDir -Force -ErrorAction SilentlyContinue }
        if ($rootCreated -and (Test-Path -LiteralPath $Root)) { Remove-Item -LiteralPath $Root -Recurse -Force -ErrorAction SilentlyContinue }
    } catch { }
    throw $err
}
if (Test-Path -LiteralPath $stage) { Remove-Item -LiteralPath $stage -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host "ThinQQ $coreVer instalado em $Root. Abra o Excel e procure a aba 'ThinQQ'."
