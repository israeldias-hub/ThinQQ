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
    if ((Test-Path -LiteralPath (Join-Path $PackageDir 'ThinQQ.ppam')) -and (Get-Process powerpnt -ErrorAction SilentlyContinue)) { throw 'Feche o PowerPoint antes de instalar o ThinQQ (o suplemento do PowerPoint é atualizado junto).' }
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

# Suplemento do PowerPoint (opcional: ZIPs anteriores a 1.4.0 nao tem). Manifesto que o cita exige o arquivo e o SHA-256.
$ppamSrc = Join-Path $PackageDir 'ThinQQ.ppam'
$hasPpamField = $manifest.PSObject.Properties['pptAddinSha256'] -and ([string]$manifest.pptAddinSha256).Trim()
$withPpam = Test-Path -LiteralPath $ppamSrc
if ($hasPpamField -and -not $withPpam) { throw "Arquivo 'ThinQQ.ppam' não encontrado em $PackageDir. Descompacte o ZIP inteiro antes de instalar." }
if ($withPpam) {
    if (-not $hasPpamField) { throw 'manifest.json sem pptAddinSha256; o suplemento do PowerPoint não pode ser verificado.' }
    $exp = ([string]$manifest.pptAddinSha256).Trim().ToLower()
    $act = (Get-FileHash -LiteralPath $ppamSrc -Algorithm SHA256).Hash.ToLower()
    if ($act -ne $exp) { throw 'Pacote corrompido ou adulterado: o SHA-256 de ThinQQ.ppam não confere com o manifest.json. Baixe o ZIP novamente.' }
}
$ppamDest = Join-Path $Root 'ThinQQ.ppam'
$hadPpam = Test-Path -LiteralPath $ppamDest
$ppamMoved = $false
# Copia/troca do .ppam: se o arquivo instalado estiver em uso (PowerPoint aberto), a falha vira uma orientacao clara (o desfazer restaura o resto)
function Invoke-PpamStep([scriptblock]$Step) {
    try { & $Step } catch { throw 'Feche o PowerPoint (o suplemento ThinQQ está em uso) e rode o instalador de novo.' }
}

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
    foreach ($f in @('ThinQQ.xlam', 'ThinQQ.core.xlam', 'tutorial.html', 'ThinQQ.ppam')) {
        if (Test-Path -LiteralPath (Join-Path $PackageDir $f)) {
            Copy-Item -LiteralPath (Join-Path $PackageDir $f) -Destination $stage -Force
            Unblock-File -LiteralPath (Join-Path $stage $f) -ErrorAction SilentlyContinue
        }
    }
    foreach ($d in @($verDir, (Join-Path $Root 'logs'), $undoDir)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }

    # guarda copias para o desfazer; o loader antigo vira ThinQQ.xlam.bak (so o ultimo e mantido)
    if ($hadLoader) { Copy-Item -LiteralPath $loaderDest -Destination (Join-Path $undoDir 'loader') -Force; Copy-Item -LiteralPath $loaderDest -Destination $bak -Force }
    if ($hadCore) { Copy-Item -LiteralPath (Join-Path $verDir 'ThinQQ.core.xlam') -Destination (Join-Path $undoDir 'core') -Force }
    if ($hadTut) { Copy-Item -LiteralPath (Join-Path $verDir 'tutorial.html') -Destination (Join-Path $undoDir 'tut') -Force }
    if ($hadPpam) { Invoke-PpamStep { Copy-Item -LiteralPath $ppamDest -Destination (Join-Path $undoDir 'ppam') -Force } }

    Move-Item -LiteralPath (Join-Path $stage 'ThinQQ.xlam') -Destination $loaderDest -Force; $loaderMoved = $true
    Move-Item -LiteralPath (Join-Path $stage 'ThinQQ.core.xlam') -Destination $verDir -Force; $coreMoved = $true
    if (Test-Path -LiteralPath (Join-Path $stage 'tutorial.html')) { Move-Item -LiteralPath (Join-Path $stage 'tutorial.html') -Destination $verDir -Force; $tutMoved = $true }
    if ($withPpam) { Invoke-PpamStep { Move-Item -LiteralPath (Join-Path $stage 'ThinQQ.ppam') -Destination $ppamDest -Force }; $ppamMoved = $true }

    # --- Local confiavel (macros permitidas sem aviso) ---
    $tl = Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ'
    if (-not (Test-Path -LiteralPath $tl)) {
        Initialize-RegKey $tl
        [void]$undo.Add({ Remove-Item -LiteralPath $tl -Recurse -Force -ErrorAction SilentlyContinue }.GetNewClosure())
    }
    Set-ItemProperty -LiteralPath $tl -Name Path -Value ($Root + '\')
    Set-ItemProperty -LiteralPath $tl -Name AllowSubfolders -Value 1 -Type DWord
    Set-ItemProperty -LiteralPath $tl -Name Description -Value 'ThinQQ add-in'

    # --- PowerPoint: local confiavel e registro do suplemento (AddIns\ThinQQ: Path, AutoLoad) ---
    if ($withPpam) {
        $ptl = Join-Path $RegistryRoot 'Microsoft\Office\16.0\PowerPoint\Security\Trusted Locations\LocationThinQQ'
        if (-not (Test-Path -LiteralPath $ptl)) {
            Initialize-RegKey $ptl
            [void]$undo.Add({ Remove-Item -LiteralPath $ptl -Recurse -Force -ErrorAction SilentlyContinue }.GetNewClosure())
        }
        Set-ItemProperty -LiteralPath $ptl -Name Path -Value ($Root + '\')
        Set-ItemProperty -LiteralPath $ptl -Name AllowSubfolders -Value 1 -Type DWord
        Set-ItemProperty -LiteralPath $ptl -Name Description -Value 'ThinQQ add-in'
        $pak = Join-Path $RegistryRoot 'Microsoft\Office\16.0\PowerPoint\AddIns\ThinQQ'
        if (-not (Test-Path -LiteralPath $pak)) {
            Initialize-RegKey $pak
            [void]$undo.Add({ Remove-Item -LiteralPath $pak -Recurse -Force -ErrorAction SilentlyContinue }.GetNewClosure())
        }
        Set-ItemProperty -LiteralPath $pak -Name Path -Value $ppamDest
        Set-ItemProperty -LiteralPath $pak -Name AutoLoad -Value 1 -Type DWord
        # versao instalada do suplemento (spec 4.9: o atualizador do Excel compara com pptAddinVersion e so atualiza quem ja tem o .ppam)
        $ppVer = [string]$manifest.pptAddinVersion
        if ($ppVer -notmatch '^\d+\.\d+\.\d+$') { $ppVer = $coreVer }
        $upk = Join-Path $RegistryRoot 'VB and VBA Program Settings\ThinQQ\Updater'
        Initialize-RegKey $upk
        $oldPpVer = (Get-ItemProperty -LiteralPath $upk).PSObject.Properties['PpamVersion']
        $oldPpVal = $null; if ($oldPpVer) { $oldPpVal = [string]$oldPpVer.Value }
        Set-ItemProperty -LiteralPath $upk -Name PpamVersion -Value $ppVer
        [void]$undo.Add({ if ($null -ne $oldPpVal) { Set-ItemProperty -LiteralPath $upk -Name PpamVersion -Value $oldPpVal } else { Remove-ItemProperty -LiteralPath $upk -Name PpamVersion -ErrorAction SilentlyContinue } }.GetNewClosure())
        # pendencia antiga do atualizador (ThinQQ.ppam.new) perde o sentido: o ZIP acabou de instalar a versao dele
        foreach ($n in @('PendingPpam', 'PendingPpamSha256', 'PendingPpamVersion')) { Remove-ItemProperty -LiteralPath $upk -Name $n -ErrorAction SilentlyContinue }
        Remove-Item -LiteralPath ($ppamDest + '.new') -Force -ErrorAction SilentlyContinue
        # sobra de troca interrompida do atualizador (ThinQQ.ppam.old)
        Remove-Item -LiteralPath ($ppamDest + '.old') -Force -ErrorAction SilentlyContinue
    }
    if ($SimulateFailureForTest) { throw 'Falha simulada (teste).' }

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
        if ($ppamMoved) {
            if ($hadPpam) { Copy-Item -LiteralPath (Join-Path $undoDir 'ppam') -Destination $ppamDest -Force }
            else { Remove-Item -LiteralPath $ppamDest -Force -ErrorAction SilentlyContinue }
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
if ($withPpam) { Write-Host "Suplemento do PowerPoint instalado: abra o PowerPoint e procure a aba 'ThinQQ'." }
