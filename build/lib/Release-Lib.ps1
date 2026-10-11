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

# Le o vbaProject.bin dentro de um xlam (zip, em memoria) e devolve os marcadores de teste encontrados.
function Get-TestMarkersInXlamBytes([byte[]]$XlamBytes) {
    $ms = New-Object IO.MemoryStream(,$XlamBytes)
    $z = New-Object IO.Compression.ZipArchive($ms, [IO.Compression.ZipArchiveMode]::Read)
    try {
        $e = $z.GetEntry('xl/vbaProject.bin')
        if (-not $e) { throw 'xl/vbaProject.bin ausente no xlam' }
        return (Get-TestMarkersInBytes (Read-EntryBytes $e))
    } finally { $z.Dispose() }
}

function Get-TestMarkersInXlam([string]$Path) {
    return (Get-TestMarkersInXlamBytes ([IO.File]::ReadAllBytes($Path)))
}

# Varre todos os .xlam dentro de um zip de release.
function Get-TestMarkersInZip([string]$ZipPath) {
    $result = New-Object System.Collections.Generic.List[string]
    $z = [IO.Compression.ZipFile]::OpenRead($ZipPath)
    try {
        foreach ($e in $z.Entries) {
            if ($e.Name -like '*.xlam') {
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

function New-ReleaseManifest([string]$CoreVersion, [string]$LoaderVersion, [string]$BaseUrl, [string]$CoreSha, [string]$LoaderSha, [string]$Notes, [string]$PublishedAt) {
    return [ordered]@{
        schemaVersion = 1
        coreVersion   = $CoreVersion
        coreUrl       = Join-ReleaseUrl $BaseUrl 'ThinQQ.core.xlam'
        coreSha256    = $CoreSha
        loaderVersion = $LoaderVersion
        loaderUrl     = Join-ReleaseUrl $BaseUrl 'ThinQQ.xlam'
        loaderSha256  = $LoaderSha
        tutorialUrl   = Join-ReleaseUrl $BaseUrl 'tutorial.html'
        manifestUrl   = Join-ReleaseUrl $BaseUrl 'manifest.json'
        publishedAt   = $PublishedAt
        notes         = $Notes
    }
}

# "Stub" do tutorial: pequeno demais ou com o marcador do esqueleto.
function Test-TutorialIsStub([string]$Path) {
    $len = (Get-Item -LiteralPath $Path).Length
    if ($len -lt 2000) { return $true }
    return ([IO.File]::ReadAllText($Path) -match 'THINQQ-TUTORIAL-STUB')
}
