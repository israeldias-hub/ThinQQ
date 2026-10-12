# Testa a verificacao de modulos de teste nos artefatos de release e o manifesto gerado (ida e volta por JSON).
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
. (Join-Path $repo 'build\lib\Release-Lib.ps1')
$dist = Join-Path $repo 'dist'
function Assert($cond, $msg) { if (-not $cond) { throw "FALHA: $msg" } }
$tmp = Join-Path $env:TEMP ("tqq_rg_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $tmp | Out-Null
try {
    # xlam sintetico: vbaProject.bin com nome de modulo em UTF-16 (diretorio OLE) e em texto (stream PROJECT)
    function New-FakeXlam([string]$path, [byte[]]$bin) {
        $z = [IO.Compression.ZipFile]::Open($path, 'Create')
        $e = $z.CreateEntry('xl/vbaProject.bin'); $s = $e.Open(); $s.Write($bin, 0, $bin.Length); $s.Dispose()
        $z.Dispose()
    }
    $clean = [Text.Encoding]::ASCII.GetBytes('Module=modCoreApi' + [char]13 + [char]10 + 'Module=modErrors')
    $dirty16 = [Text.Encoding]::Unicode.GetBytes('Root Entry modTestRunner')
    $dirtyTxt = [Text.Encoding]::ASCII.GetBytes('Module=TestSmoke')
    New-FakeXlam (Join-Path $tmp 'clean.xlam') $clean
    New-FakeXlam (Join-Path $tmp 'dirty16.xlam') $dirty16
    New-FakeXlam (Join-Path $tmp 'dirtyTxt.xlam') $dirtyTxt
    Assert (@(Get-TestMarkersInXlam (Join-Path $tmp 'clean.xlam')).Count -eq 0) 'xlam limpo sem marcadores'
    Assert (@(Get-TestMarkersInXlam (Join-Path $tmp 'dirty16.xlam')) -contains 'modTestRunner') 'detecta nome UTF-16'
    Assert (@(Get-TestMarkersInXlam (Join-Path $tmp 'dirtyTxt.xlam')) -contains 'Module=TestSmoke') 'detecta Module=Test*'
    # zip com xlam sujo
    $zipPath = Join-Path $tmp 'rel.zip'
    $z = [IO.Compression.ZipFile]::Open($zipPath, 'Create')
    [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($z, (Join-Path $tmp 'dirty16.xlam'), 'ThinQQ.core.xlam')
    $z.Dispose()
    Assert (@(Get-TestMarkersInZip $zipPath).Count -gt 0) 'detecta modulo de teste dentro do zip'

    # artefatos reais, se existirem (build Test deve falhar na verificacao; Release deve passar)
    $found = 0
    foreach ($f in (Get-ChildItem $dist -Filter '*.test.xlam' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '~$*' })) {
        Assert (@(Get-TestMarkersInXlam $f.FullName).Count -gt 0) "build de teste $($f.Name) deveria ser reprovado"; $found++
    }
    foreach ($n in @('ThinQQ.xlam', 'ThinQQ.core.xlam')) {
        $p = Join-Path $dist $n
        if (Test-Path $p) { Assert (@(Get-TestMarkersInXlam $p).Count -eq 0) "release $n nao pode ter modulos de teste"; $found++ }
    }
    Write-Host "Artefatos reais verificados: $found"

    # manifesto: ida e volta com aspas, marcador de lista, quebras de linha e base UNC
    $notes = "• Primeira `"versão`" com aspas`r`nsegunda linha`ttab e barra \ final"
    $m = New-ReleaseManifest -CoreVersion '1.2.3' -LoaderVersion '1.0.0' -BaseUrl '\\servidor\ThinQQ\' -CoreSha ('a' * 64) -LoaderSha ('b' * 64) -Notes $notes -PublishedAt '2026-10-10'
    $json = $m | ConvertTo-Json
    $back = $json | ConvertFrom-Json
    Assert ($back.notes -ceq $notes) 'notes volta identico'
    Assert ($back.coreUrl -ceq '\\servidor\ThinQQ\ThinQQ.core.xlam') "coreUrl UNC: $($back.coreUrl)"
    Assert ($back.manifestUrl -ceq '\\servidor\ThinQQ\manifest.json') 'manifestUrl UNC'
    Assert ($back.schemaVersion -eq 1 -and $back.coreSha256 -ceq ('a' * 64) -and $back.loaderSha256 -ceq ('b' * 64)) 'campos obrigatorios'
    Assert ($json -notmatch '[\x00-\x09\x0B\x0C\x0E-\x1F]') 'sem caracteres de controle crus (so CR/LF estrutural)'
    Assert ((($json | ConvertFrom-Json | ConvertTo-Json -Compress)) -notmatch '[\x00-\x1F]') 'forma compacta sem controles'
    $h = New-ReleaseManifest -CoreVersion '1.0.0' -LoaderVersion '1.0.0' -BaseUrl 'https://x.y/z/' -CoreSha ('a' * 64) -LoaderSha ('b' * 64) -Notes '' -PublishedAt '2026-10-10'
    Assert ($h.coreUrl -ceq 'https://x.y/z/ThinQQ.core.xlam') 'URL http usa /'
    # manifestUrl (vira a fonte de atualizacao de quem instala pelo ZIP): com BaseUrl http(s) e o endereco ESTAVEL (DEFAULT_UPDATE_SOURCE do loader,
    # o release/manifest.json da main), nunca o asset versionado; com pasta de rede/local, BaseUrl\manifest.json; -ManifestUrl explicito vence
    $settingsBas = [IO.File]::ReadAllText((Join-Path $repo 'src\loader\modules\modSettings.bas'))
    $stable = [regex]::Match($settingsBas, 'DEFAULT_UPDATE_SOURCE As String = "([^"]+)"').Groups[1].Value
    Assert ($stable -match '^https://') "DEFAULT_UPDATE_SOURCE lido do loader: $stable"
    Assert ($h.manifestUrl -ceq $stable) "manifestUrl com BaseUrl http e o endereco estavel: $($h.manifestUrl)"
    Assert ((Get-DefaultUpdateSource) -ceq $stable) 'Get-DefaultUpdateSource le o mesmo endereco do loader'
    $e = New-ReleaseManifest -CoreVersion '1.0.0' -LoaderVersion '1.0.0' -BaseUrl 'https://x.y/z/' -ManifestUrl 'https://outro/m.json' -CoreSha ('a' * 64) -LoaderSha ('b' * 64) -Notes '' -PublishedAt '2026-10-10'
    Assert ($e.manifestUrl -ceq 'https://outro/m.json') 'ManifestUrl explicito vence'
    $l = New-ReleaseManifest -CoreVersion '1.0.0' -LoaderVersion '1.0.0' -BaseUrl 'C:\temp\ensaio' -CoreSha ('a' * 64) -LoaderSha ('b' * 64) -Notes '' -PublishedAt '2026-10-10'
    Assert ($l.manifestUrl -ceq 'C:\temp\ensaio\manifest.json') "pasta local: $($l.manifestUrl)"

    # stub do tutorial
    $stub = Join-Path $tmp 'stub.html'; Set-Content $stub ('x' * 5000 + '<!-- THINQQ-TUTORIAL-STUB -->')
    Assert (Test-TutorialIsStub $stub) 'marcador => stub'
    Set-Content $stub '<html>pequeno</html>'
    Assert (Test-TutorialIsStub $stub) 'arquivo pequeno => stub'
    Set-Content $stub ('x' * 5000)
    Assert (-not (Test-TutorialIsStub $stub)) 'arquivo grande sem marcador nao e stub'
    Write-Host "Test-ReleaseGuard: OK (PowerShell $($PSVersionTable.PSVersion))" -ForegroundColor Green
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
