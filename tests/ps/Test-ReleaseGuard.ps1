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
    # suplemento do PowerPoint (.ppam): zip com ppt/vbaProject.bin (e nao xl/); o mesmo teste de marcadores vale
    function New-FakePpam([string]$path, [byte[]]$bin, [string]$entry = 'ppt/vbaProject.bin') {
        $z = [IO.Compression.ZipFile]::Open($path, 'Create')
        if ($entry) { $e = $z.CreateEntry($entry); $s = $e.Open(); $s.Write($bin, 0, $bin.Length); $s.Dispose() }
        $z.Dispose()
    }
    New-FakePpam (Join-Path $tmp 'clean.ppam') $clean
    # o .ppam de teste leva os modulos proprios do PowerPoint (modPptTestRunner/modPptAssert), nao os do Excel
    $dirtyPpt16 = [Text.Encoding]::Unicode.GetBytes('Root Entry modPptTestRunner')
    New-FakePpam (Join-Path $tmp 'dirty.ppam') $dirtyPpt16
    New-FakePpam (Join-Path $tmp 'dirtyAssert.ppam') ([Text.Encoding]::Unicode.GetBytes('Root Entry modPptAssert'))
    Assert (@(Get-TestMarkersInXlam (Join-Path $tmp 'dirtyAssert.ppam')) -contains 'modPptAssert') '.ppam com modPptAssert e recusado'
    New-FakePpam (Join-Path $tmp 'sem-bin.ppam') $clean ''
    Assert (@(Get-TestMarkersInXlam (Join-Path $tmp 'clean.ppam')).Count -eq 0) '.ppam limpo sem marcadores'
    Assert (@(Get-TestMarkersInXlam (Join-Path $tmp 'dirty.ppam')) -contains 'modPptTestRunner') '.ppam com modulo de teste (ppt/vbaProject.bin) e recusado'
    $threw = $false; $ppamMsg = ''
    try { [void](Get-TestMarkersInXlam (Join-Path $tmp 'sem-bin.ppam')) } catch { $threw = $true; $ppamMsg = $_.Exception.Message }
    Assert ($threw -and $ppamMsg -match 'vbaProject\.bin ausente') ".ppam sem vbaProject.bin: erro claro ($ppamMsg)"
    $zipP = Join-Path $tmp 'rel-ppam.zip'
    $z = [IO.Compression.ZipFile]::Open($zipP, 'Create')
    [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($z, (Join-Path $tmp 'dirty.ppam'), 'ThinQQ.ppam')
    $z.Dispose()
    Assert ((@(Get-TestMarkersInZip $zipP) -join ' ') -match 'ThinQQ\.ppam: modPptTestRunner') 'detecta modulo de teste no .ppam dentro do zip'

    # artefatos reais, se existirem (build Test deve falhar na verificacao; Release deve passar)
    $found = 0
    foreach ($f in (Get-ChildItem $dist -Filter '*.test.xlam' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '~$*' })) {
        Assert (@(Get-TestMarkersInXlam $f.FullName).Count -gt 0) "build de teste $($f.Name) deveria ser reprovado"; $found++
    }
    foreach ($f in (Get-ChildItem $dist -Filter '*.pptm' -ErrorAction SilentlyContinue | Where-Object { $_.Name -notlike '~$*' })) {
        Assert (@(Get-TestMarkersInXlam $f.FullName).Count -gt 0) "build de teste do PowerPoint $($f.Name) deveria ser reprovado"; $found++
    }
    foreach ($n in @('ThinQQ.xlam', 'ThinQQ.core.xlam', 'ThinQQ.ppam')) {
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

    # suplemento do PowerPoint no manifesto (spec 4.9): pptAddinVersion = coreVersion (saem juntos), Url do asset, SHA-256; sem -PptSha os TRES campos somem
    $pm = New-ReleaseManifest -CoreVersion '1.2.3' -LoaderVersion '1.0.0' -BaseUrl 'https://x.y/z/' -CoreSha ('a' * 64) -LoaderSha ('b' * 64) -PptSha ('c' * 64) -Notes '' -PublishedAt '2026-10-10'
    $pb = ($pm | ConvertTo-Json) | ConvertFrom-Json
    Assert ($pb.pptAddinVersion -ceq '1.2.3' -and $pb.pptAddinVersion -ceq $pb.coreVersion) "pptAddinVersion = coreVersion ($($pb.pptAddinVersion))"
    Assert ($pb.pptAddinUrl -ceq 'https://x.y/z/ThinQQ.ppam' -and $pb.pptAddinUrl.EndsWith('/ThinQQ.ppam')) "pptAddinUrl ($($pb.pptAddinUrl))"
    Assert ($pb.pptAddinSha256 -ceq ('c' * 64)) 'pptAddinSha256'
    Assert ($pb.coreSha256 -ceq ('a' * 64) -and $pb.loaderSha256 -ceq ('b' * 64) -and $pb.schemaVersion -eq 1) 'com o suplemento, os demais campos continuam'
    $pk = @($pm.Keys); $iL = $pk.IndexOf('loaderSha256')
    Assert ($pk.IndexOf('pptAddinVersion') -eq ($iL + 1) -and $pk.IndexOf('pptAddinUrl') -eq ($iL + 2) -and $pk.IndexOf('pptAddinSha256') -eq ($iL + 3)) 'os campos do suplemento vem logo depois de loaderSha256'
    $pu = New-ReleaseManifest -CoreVersion '1.2.3' -LoaderVersion '1.0.0' -BaseUrl '\\servidor\ThinQQ' -CoreSha ('a' * 64) -LoaderSha ('b' * 64) -PptSha ('c' * 64) -Notes '' -PublishedAt '2026-10-10'
    Assert ($pu.pptAddinUrl -ceq '\\servidor\ThinQQ\ThinQQ.ppam') "pptAddinUrl em pasta de rede ($($pu.pptAddinUrl))"
    $mNames = @((($m | ConvertTo-Json) | ConvertFrom-Json).PSObject.Properties.Name)
    foreach ($k in 'pptAddinVersion', 'pptAddinUrl', 'pptAddinSha256') {
        Assert (-not $m.Contains($k) -and -not $h.Contains($k)) "sem -PptSha: $k nao aparece (manifesto de uma release sem o suplemento)"
        Assert ($mNames -cnotcontains $k) "sem -PptSha: $k ausente do JSON"
    }

    # stub do tutorial
    $stub = Join-Path $tmp 'stub.html'; Set-Content $stub ('x' * 5000 + '<!-- THINQQ-TUTORIAL-STUB -->')
    Assert (Test-TutorialIsStub $stub) 'marcador => stub'
    Set-Content $stub '<html>pequeno</html>'
    Assert (Test-TutorialIsStub $stub) 'arquivo pequeno => stub'
    Set-Content $stub ('x' * 5000)
    Assert (-not (Test-TutorialIsStub $stub)) 'arquivo grande sem marcador nao e stub'

    # ---- conteudo do ZIP: o Excel de demonstracao viaja em exemplos/ (so o .xlsx e um LEIA-ME), identico ao do repositorio e sem macros
    $demoName = Get-ReleaseDemoName
    Assert ($demoName -ceq 'ThinQQ-Demonstracao-Materiais-Construcao.xlsx') "nome do arquivo de demonstracao: $demoName"
    $top = @(Get-ReleaseTopFiles)
    foreach ($n in 'ThinQQ.xlam', 'ThinQQ.core.xlam', 'ThinQQ.ppam', 'manifest.json', 'tutorial.html', 'install.ps1', 'uninstall.ps1', 'Instalar-ThinQQ.cmd', 'LEIA-ME.txt') { Assert ($top -contains $n) "o ZIP leva $n" }
    function New-FakeXlsx([string]$path, [string]$marker, [switch]$WithMacros) {
        $z = [IO.Compression.ZipFile]::Open($path, 'Create')
        $names = @('[Content_Types].xml', 'xl/workbook.xml'); if ($WithMacros) { $names += 'xl/vbaProject.bin' }
        foreach ($n in $names) { $e = $z.CreateEntry($n); $w = New-Object IO.StreamWriter($e.Open()); $w.Write("$marker $n"); $w.Dispose() }
        $z.Dispose()
    }
    $pkg = Join-Path $tmp 'pkg'; New-Item -ItemType Directory -Force $pkg | Out-Null
    foreach ($n in $top) { Set-Content -LiteralPath (Join-Path $pkg $n) "conteudo de $n" }
    $demoSrc = Join-Path $tmp 'demo-src.xlsx'; New-FakeXlsx $demoSrc 'repo'
    $demoTxt = Join-Path $tmp 'leia-me-exemplos.txt'; Set-Content -LiteralPath $demoTxt 'exemplos'
    function New-Layout($demoPath) {
        $m = [ordered]@{}
        foreach ($n in $top) { $m[$n] = Join-Path $pkg $n }
        $m["exemplos/$demoName"] = $demoPath
        $m['exemplos/LEIA-ME.txt'] = $demoTxt
        return $m
    }
    $zipOk = Join-Path $tmp 'ok.zip'
    New-ReleaseZip $zipOk (New-Layout $demoSrc)
    $zr = [IO.Compression.ZipFile]::OpenRead($zipOk); $names = @($zr.Entries | ForEach-Object { $_.FullName }); $zr.Dispose()
    Assert ($names -contains "exemplos/$demoName" -and $names -contains 'exemplos/LEIA-ME.txt') "New-ReleaseZip grava exemplos/ com '/': $($names -join ', ')"
    Assert (@($names | Where-Object { $_.Contains([string][char]92) }).Count -eq 0) 'New-ReleaseZip nao grava barra invertida nos nomes'
    $problems = @(Test-ReleaseZipContent $zipOk $demoSrc)
    Assert ($problems.Count -eq 0) "ZIP correto passa na guarda: $($problems -join '; ')"
    # negativos: cada defeito e apontado
    $lay = New-Layout $demoSrc; $lay.Remove("exemplos/$demoName"); $zipNoDemo = Join-Path $tmp 'nodemo.zip'; New-ReleaseZip $zipNoDemo $lay
    Assert ((@(Test-ReleaseZipContent $zipNoDemo $demoSrc) -join ' ') -match 'ausente') 'sem o Excel de demonstracao: reprovado'
    $other = Join-Path $tmp 'outro.xlsx'; New-FakeXlsx $other 'outro'
    $lay = New-Layout $demoSrc; $lay['exemplos/Outro.xlsx'] = $other; $zipExtra = Join-Path $tmp 'extra.zip'; New-ReleaseZip $zipExtra $lay
    Assert ((@(Test-ReleaseZipContent $zipExtra $demoSrc) -join ' ') -match 'Outro\.xlsx') 'outro .xlsx no ZIP: reprovado'
    $lay = New-Layout $demoSrc; $lay['Rascunho.xlsm'] = $other; $zipXlsm = Join-Path $tmp 'xlsm.zip'; New-ReleaseZip $zipXlsm $lay
    Assert ((@(Test-ReleaseZipContent $zipXlsm $demoSrc) -join ' ') -match 'Rascunho\.xlsm') '.xlsm fora do lugar: reprovado'
    $lay = New-Layout $other; $zipDiff = Join-Path $tmp 'diff.zip'; New-ReleaseZip $zipDiff $lay
    Assert ((@(Test-ReleaseZipContent $zipDiff $demoSrc) -join ' ') -match 'SHA-256') 'xlsx diferente do repositorio: reprovado'
    $macro = Join-Path $tmp 'macro.xlsx'; New-FakeXlsx $macro 'repo' -WithMacros
    $lay = New-Layout $macro; $zipMacro = Join-Path $tmp 'macro.zip'; New-ReleaseZip $zipMacro $lay
    Assert ((@(Test-ReleaseZipContent $zipMacro $macro) -join ' ') -match 'vbaProject') 'xlsx com macros: reprovado'
    $lay = New-Layout $demoSrc; $lay.Remove('ThinQQ.core.xlam'); $zipNoCore = Join-Path $tmp 'nocore.zip'; New-ReleaseZip $zipNoCore $lay
    Assert ((@(Test-ReleaseZipContent $zipNoCore $demoSrc) -join ' ') -match 'ThinQQ\.core\.xlam') 'falta de um arquivo do pacote: reprovado'
    $lay = New-Layout $demoSrc; $lay.Remove('ThinQQ.ppam'); $zipNoPpam = Join-Path $tmp 'noppam.zip'; New-ReleaseZip $zipNoPpam $lay
    Assert ((@(Test-ReleaseZipContent $zipNoPpam $demoSrc) -join ' ') -match 'ausente no ZIP: ThinQQ\.ppam') 'ZIP sem o ThinQQ.ppam: reprovado ("ausente no ZIP: ThinQQ.ppam")'
    $loose = @($names | Where-Object { $_ -notlike 'exemplos/*' })
    Assert ($top.Count -eq 9 -and $loose.Count -eq 9 -and $loose -contains 'ThinQQ.ppam') "o ZIP leva 9 arquivos soltos na raiz, com o ThinQQ.ppam ($($loose.Count))"
    $zb = [IO.Compression.ZipFile]::Open((Join-Path $tmp 'bs.zip'), 'Create'); [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zb, $demoSrc, "exemplos\$demoName"); $zb.Dispose()
    Assert ((@(Test-ReleaseZipContent (Join-Path $tmp 'bs.zip') $demoSrc) -join ' ') -match 'barra invertida') 'nome com barra invertida: reprovado'
    # o arquivo de demonstracao do repositorio: existe, sem macros (a guarda do ZIP o compara com ele) e com o LEIA-ME de exemplos ao lado
    $repoDemo = Join-Path $repo "exemplos\$demoName"
    Assert (Test-Path -LiteralPath $repoDemo) "o repositorio tem exemplos\$demoName"
    Assert (-not (Test-XlsxHasMacros $repoDemo)) 'o xlsx de demonstracao do repositorio nao tem vbaProject'
    Assert (Test-Path -LiteralPath (Join-Path $repo 'installer\LEIA-ME-exemplos.txt')) 'installer\LEIA-ME-exemplos.txt existe (vai como exemplos\LEIA-ME.txt)'
    $layout = Get-ReleaseZipLayout $repo $dist
    Assert ($layout.Contains("exemplos/$demoName") -and $layout["exemplos/$demoName"] -ceq $repoDemo) 'o layout do release usa o xlsx do repositorio'
    # ZIP real do release (so o da versao atual de version.json; um ZIP antigo em dist nao conta)
    $curVer = ((Get-Content -LiteralPath (Join-Path $repo 'version.json') -Raw -Encoding UTF8) | ConvertFrom-Json).core
    $realZip = Join-Path $dist "ThinQQ-$curVer.zip"
    if (Test-Path -LiteralPath $realZip) {
        $rp = @(Test-ReleaseZipContent $realZip $repoDemo)
        Assert ($rp.Count -eq 0) "ZIP real $curVer passa na guarda: $($rp -join '; ')"
        Write-Host "ZIP real verificado: $realZip"
    }
    # release.ps1 (NAO e executado aqui: o build do .ppam usa o PowerPoint): confere no texto que cada etapa do suplemento esta no lugar e na ordem
    $relText = [IO.File]::ReadAllText((Join-Path $repo 'build\release.ps1'))
    $iAll = $relText.IndexOf('-Target all -Configuration Release'); $iPpt = $relText.IndexOf('-Target ppt -Configuration Release')
    Assert ($iAll -ge 0 -and $iPpt -gt $iAll) 'release.ps1 constroi o .ppam (-Target ppt -Configuration Release) depois do build all'
    Assert ($relText -match 'Build do suplemento do PowerPoint falhou' -and $relText -match 'VBOM') 'release.ps1: falha do build do .ppam aborta com mensagem clara (VBOM)'
    Assert ($relText.Contains("@('ThinQQ.xlam', 'ThinQQ.core.xlam', 'ThinQQ.ppam')")) 'release.ps1: o .ppam entra na conferencia de marcadores de teste'
    $signLine = @($relText -split "`r?`n" | Where-Object { $_.Contains('Sign-ThinQQ.ps1') -and $_.Contains('-Files') })
    Assert ($signLine.Count -eq 1 -and $signLine[0].Contains("(Join-Path `$DistDir 'ThinQQ.ppam')")) 'release.ps1: assina tambem o .ppam'
    Assert ($relText.IndexOf('Sign-ThinQQ.ps1') -lt $relText.IndexOf('-PptSha') -and $relText.Contains("-PptSha (Get-FileHash (Join-Path `$DistDir 'ThinQQ.ppam') -Algorithm SHA256).Hash.ToLower()")) 'release.ps1: o SHA-256 do .ppam entra no manifesto, depois da assinatura'
    Write-Host "Test-ReleaseGuard: OK (PowerShell $($PSVersionTable.PSVersion))" -ForegroundColor Green
} finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
