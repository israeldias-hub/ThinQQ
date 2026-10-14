# Funcoes do release: URLs, manifesto e verificacao de modulos de teste nos artefatos.
# Compativel com Windows PowerShell 5.1 e pwsh 7.
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Marcadores que nunca podem aparecer num artefato Release (nomes de modulos de teste).
# Os nomes ficam em texto puro no stream PROJECT e, em UTF-16, no diretorio OLE do vbaProject.bin.
$script:TestMarkerRegex = 'modTestDispatch|modTestRunner|modAssert|Module=Test[A-Za-z]+|Class=Test[A-Za-z]+'

function Get-TestMarkersInBytes([byte[]]$Bytes) {
    $text = [Text.Encoding]::GetEncoding(28591).GetString($Bytes).Replace([string][char]0, '')
    $found = New-Object System.Collections.Generic.List[string]
    foreach ($m in [regex]::Matches($text, $script:TestMarkerRegex)) { if (-not $found.Contains($m.Value)) { $found.Add($m.Value) } }
    return $found.ToArray()
}

function Read-EntryBytes($Entry) {
    $ms = New-Object IO.MemoryStream
    $s = $Entry.Open()
    try { $s.CopyTo($ms) } finally { $s.Dispose() }
    return ,$ms.ToArray()
}

# Le o vbaProject.bin dentro de um xlam ou ppam (zip, em memoria; xl/ no Excel, ppt/ no PowerPoint) e devolve os marcadores de teste encontrados.
function Get-TestMarkersInXlamBytes([byte[]]$XlamBytes) {
    $ms = New-Object IO.MemoryStream(,$XlamBytes)
    $z = New-Object IO.Compression.ZipArchive($ms, [IO.Compression.ZipArchiveMode]::Read)
    try {
        $e = $z.GetEntry('xl/vbaProject.bin')
        if (-not $e) { $e = $z.GetEntry('ppt/vbaProject.bin') }
        if (-not $e) { throw 'vbaProject.bin ausente no pacote' }
        return (Get-TestMarkersInBytes (Read-EntryBytes $e))
    } finally { $z.Dispose() }
}

function Get-TestMarkersInXlam([string]$Path) {
    return (Get-TestMarkersInXlamBytes ([IO.File]::ReadAllBytes($Path)))
}

# Varre todos os .xlam e .ppam dentro de um zip de release.
function Get-TestMarkersInZip([string]$ZipPath) {
    $result = New-Object System.Collections.Generic.List[string]
    $z = [IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        foreach ($e in $z.Entries) {
            if ($e.Name -like '*.xlam' -or $e.Name -like '*.ppam') {
                foreach ($m in (Get-TestMarkersInXlamBytes (Read-EntryBytes $e))) { $result.Add("$($e.Name): $m") }
            }
        }
    } finally { $z.Dispose() }
    return $result.ToArray()
}

# Junta base + nome: http(s) usa '/', pasta local/UNC usa '\'.
function Join-ReleaseUrl([string]$Base, [string]$Name) {
    $sep = '\'
    if ($Base -match '^https?://') { $sep = '/' }
    return $Base.TrimEnd('/', '\') + $sep + $Name
}

# Fonte de atualizacao padrao do loader (DEFAULT_UPDATE_SOURCE em src/loader/modules/modSettings.bas: o release/manifest.json da main), lida do
# proprio fonte para as duas listas nunca divergirem.
function Get-DefaultUpdateSource {
    $bas = Join-Path $PSScriptRoot '..\..\src\loader\modules\modSettings.bas'
    $m = [regex]::Match([IO.File]::ReadAllText($bas), 'DEFAULT_UPDATE_SOURCE As String = "([^"]+)"')
    if (-not $m.Success) { throw "DEFAULT_UPDATE_SOURCE nao encontrado em $bas" }
    return $m.Groups[1].Value
}

# manifestUrl = a fonte de atualizacao que o instalador grava para quem instala pelo ZIP. Sem -ManifestUrl: com BaseUrl http(s) (release no
# GitHub, assets versionados) e o endereco ESTAVEL (Get-DefaultUpdateSource), senao o ZIP da v1.1.0 consultaria para sempre o manifesto da
# v1.1.0; com BaseUrl em pasta de rede/local (o manifesto e republicado no mesmo lugar) e BaseUrl\manifest.json.
function New-ReleaseManifest([string]$CoreVersion, [string]$LoaderVersion, [string]$BaseUrl, [string]$CoreSha, [string]$LoaderSha, [string]$Notes, [string]$PublishedAt, [string]$ManifestUrl = '', [string]$PptSha = '') {
    if (-not $ManifestUrl) {
        if ($BaseUrl -match '^https?://') { $ManifestUrl = Get-DefaultUpdateSource } else { $ManifestUrl = Join-ReleaseUrl $BaseUrl 'manifest.json' }
    }
    $m = [ordered]@{
        schemaVersion = 1
        coreVersion   = $CoreVersion
        coreUrl       = Join-ReleaseUrl $BaseUrl 'ThinQQ.core.xlam'
        coreSha256    = $CoreSha
        loaderVersion = $LoaderVersion
        loaderUrl     = Join-ReleaseUrl $BaseUrl 'ThinQQ.xlam'
        loaderSha256  = $LoaderSha
    }
    # suplemento do PowerPoint (spec 4.9): sai junto com o core (mesma versao); sem -PptSha os tres campos nao aparecem
    if ($PptSha) {
        $m['pptAddinVersion'] = $CoreVersion
        $m['pptAddinUrl']     = Join-ReleaseUrl $BaseUrl 'ThinQQ.ppam'
        $m['pptAddinSha256']  = $PptSha
    }
    $m['tutorialUrl'] = Join-ReleaseUrl $BaseUrl 'tutorial.html'
    $m['manifestUrl'] = $ManifestUrl
    $m['publishedAt'] = $PublishedAt
    $m['notes']       = $Notes
    return $m
}

# "Stub" do tutorial: pequeno demais ou com o marcador do esqueleto.
function Test-TutorialIsStub([string]$Path) {
    $len = (Get-Item -LiteralPath $Path).Length
    if ($len -lt 2000) { return $true }
    return ([IO.File]::ReadAllText($Path) -match 'THINQQ-TUTORIAL-STUB')
}

# ---- Conteudo do ZIP do release: o pacote (arquivos soltos de dist), o Excel de demonstracao e um LEIA-ME curto em exemplos/.
# O instalador copia so os arquivos que nomeia (ThinQQ.xlam, ThinQQ.core.xlam, tutorial.html): exemplos/ so viaja no ZIP, nunca e instalado
# nem aberto. O xlsx vem de exemplos/ do repositorio (o gerador o regrava; nao mexa nele aqui) e o LEIA-ME de installer/LEIA-ME-exemplos.txt.
function Get-ReleaseDemoName { return 'ThinQQ-Demonstracao-Materiais-Construcao.xlsx' }

function Get-ReleaseInstallerFiles { return @('install.ps1', 'uninstall.ps1', 'Instalar-ThinQQ.cmd', 'LEIA-ME.txt') }

# Arquivos soltos na raiz do ZIP (e os que -PublishTo copia para a pasta de publicacao).
function Get-ReleaseTopFiles { return @('ThinQQ.xlam', 'ThinQQ.core.xlam', 'ThinQQ.ppam', 'manifest.json', 'tutorial.html') + @(Get-ReleaseInstallerFiles) }

# nome da entrada no ZIP (sempre com '/') -> arquivo de origem
function Get-ReleaseZipLayout([string]$RepoRoot, [string]$DistDir) {
    $layout = [ordered]@{}
    foreach ($n in (Get-ReleaseTopFiles)) { $layout[$n] = Join-Path $DistDir $n }
    $layout["exemplos/$(Get-ReleaseDemoName)"] = Join-Path $RepoRoot "exemplos\$(Get-ReleaseDemoName)"
    $layout['exemplos/LEIA-ME.txt'] = Join-Path $RepoRoot 'installer\LEIA-ME-exemplos.txt'
    return $layout
}

# Cria o ZIP com os nomes de entrada do layout, com '/' em qualquer versao do PowerShell (o Compress-Archive do Windows PowerShell 5.1 grava
# '\', que outras ferramentas leem como parte do nome).
function New-ReleaseZip([string]$ZipPath, $Layout) {
    if (Test-Path -LiteralPath $ZipPath) { Remove-Item -LiteralPath $ZipPath -Force }
    $z = [IO.Compression.ZipFile]::Open($ZipPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($name in $Layout.Keys) {
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($z, [string]$Layout[$name], [string]$name, [IO.Compression.CompressionLevel]::Optimal)
        }
    } finally { $z.Dispose() }
}

function Test-XlsxBytesHaveMacros([byte[]]$Bytes) {
    $ms = New-Object IO.MemoryStream(, $Bytes)
    $z = New-Object IO.Compression.ZipArchive($ms, [IO.Compression.ZipArchiveMode]::Read)
    try { return (@($z.Entries | Where-Object { $_.FullName -match 'vbaProject' }).Count -gt 0) } finally { $z.Dispose() }
}

function Test-XlsxHasMacros([string]$Path) { return (Test-XlsxBytesHaveMacros ([IO.File]::ReadAllBytes($Path))) }

# Problemas do conteudo de um ZIP de release (lista vazia = ok): so os arquivos esperados (nenhuma outra planilha .xls*), nomes com '/', o
# xlsx de demonstracao identico (SHA-256) ao do repositorio (-DemoSource) e sem macros (vbaProject).
function Test-ReleaseZipContent([string]$ZipPath, [string]$DemoSource) {
    $problems = New-Object System.Collections.Generic.List[string]
    $demo = Get-ReleaseDemoName
    $expected = @(Get-ReleaseTopFiles) + @("exemplos/$demo", 'exemplos/LEIA-ME.txt')
    $z = [IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        $names = @($z.Entries | ForEach-Object { $_.FullName } | Where-Object { -not $_.EndsWith('/') })
        foreach ($n in $names) { if ($n.Contains('\')) { $problems.Add("nome com barra invertida no ZIP: $n") } }
        foreach ($n in $expected) { if ($names -cnotcontains $n) { $problems.Add("ausente no ZIP: $n") } }
        foreach ($n in $names) {
            if ($expected -cnotcontains $n) {
                if ($n -match '\.(xlsx|xlsm|xlsb|xls)$') { $problems.Add("planilha inesperada no ZIP: $n (so exemplos/$demo viaja no release)") }
                else { $problems.Add("entrada inesperada no ZIP: $n") }
            }
        }
        $e = $z.GetEntry("exemplos/$demo")
        if ($e) {
            $bytes = Read-EntryBytes $e
            $sha = [Security.Cryptography.SHA256]::Create()
            $inZip = [BitConverter]::ToString($sha.ComputeHash($bytes)).Replace('-', '').ToLower()
            $sha.Dispose()
            if (-not (Test-Path -LiteralPath $DemoSource)) { $problems.Add("arquivo de demonstracao do repositorio ausente: $DemoSource") }
            else {
                $inRepo = (Get-FileHash -LiteralPath $DemoSource -Algorithm SHA256).Hash.ToLower()
                if ($inZip -ne $inRepo) { $problems.Add("exemplos/$demo difere do arquivo do repositorio (SHA-256 $inZip no ZIP, $inRepo no repositorio)") }
            }
            if (Test-XlsxBytesHaveMacros $bytes) { $problems.Add("exemplos/$demo contem vbaProject (macros): a demonstracao deve ser .xlsx sem macros") }
        }
    } finally { $z.Dispose() }
    return $problems.ToArray()
}
