# Consistência da documentação: o tutorial real (sem marcador de esqueleto, auto-contido, tamanho razoável,
# XML bem formado), todo rótulo de controle citado existe na faixa de opções ou nos formulários, e
# README/CHANGELOG/LEIA-ME coerentes com o código.
#   Convenção: no tutorial, todo nome de botão/menu do ThinQQ vem em <span class="ui">Rótulo</span>.
#   No README e no LEIA-ME, caminhos de menu são escritos "ThinQQ > Rótulo".
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$ErrorActionPreference = 'Stop'
$script:fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }
function Read-Utf8([string]$p) { [IO.File]::ReadAllText($p, (New-Object Text.UTF8Encoding($false))) }

$tutorialPath = Join-Path $RepoRoot 'docs\tutorial\tutorial.html'
$readmePath = Join-Path $RepoRoot 'README.md'
$changelogPath = Join-Path $RepoRoot 'CHANGELOG.md'
$leiamePath = Join-Path $RepoRoot 'installer\LEIA-ME.txt'
$ribbonPath = Join-Path $RepoRoot 'src\loader\ribbon\customUI14.xml'
$allowedLinks = @('https://github.com/israeldias-hub/ThinQQ', 'https://israeldias-hub.github.io/ThinQQ/')

# ---- rótulos conhecidos: faixa de opções + formulários (loader e core)
$known = New-Object 'System.Collections.Generic.HashSet[string]'
$ribbon = New-Object Xml.XmlDocument
$ribbon.LoadXml((Read-Utf8 $ribbonPath))
foreach ($n in $ribbon.SelectNodes('//*[@label]')) { [void]$known.Add($n.GetAttribute('label')) }
function Add-Captions($node) {
    if ($node -is [System.Management.Automation.PSCustomObject]) {
        foreach ($p in $node.PSObject.Properties) {
            if ($p.Name -ieq 'caption' -and $p.Value -is [string] -and $p.Value.Length -gt 0) { [void]$known.Add($p.Value) }
            else { Add-Captions $p.Value }
        }
    } elseif ($node -is [System.Collections.IEnumerable] -and $node -isnot [string]) {
        foreach ($i in $node) { Add-Captions $i }
    }
}
foreach ($f in @('src\loader\forms\frmUpdate.json', 'src\core\forms\frmAnnotation.json')) {
    Add-Captions ((Read-Utf8 (Join-Path $RepoRoot $f)) | ConvertFrom-Json)
}
Check ($known.Count -ge 40) "rótulos conhecidos carregados ($($known.Count))"
Check ($known.Contains('Colunas empilhadas') -and $known.Contains('Pular esta versão') -and $known.Contains('OK')) 'faixa e formulários presentes no conjunto'

# ---- tutorial
Check (Test-Path $tutorialPath) 'tutorial.html existe'
$html = Read-Utf8 $tutorialPath
$bytes = (Get-Item $tutorialPath).Length
Check ($html -notmatch 'THINQQ-TUTORIAL-STUB') 'tutorial sem marcador de esqueleto'
Check ($bytes -ge 2000) "tutorial com tamanho mínimo ($bytes bytes)"
Check ($bytes -le 61440) "tutorial com até 60 KB ($bytes bytes)"
Check ($html -match '<html lang="pt-BR">' -and $html -match 'name="viewport"') 'idioma pt-BR e viewport'
Check ($html -match 'prefers-color-scheme:\s*dark' -and $html -match '@media print') 'modo escuro e impressão'
Check ($html -match 'class="skip"' -and $html -match '<nav[^>]*aria-label' -and $html -match '<main') 'link de pular, nav e main'

$xml = New-Object Xml.XmlDocument
$xml.XmlResolver = $null
$parsed = $true
try { $xml.LoadXml(($html -replace '^\s*<!DOCTYPE[^>]*>', '')) } catch { $parsed = $false; Write-Host "  $($_.Exception.Message)" }
Check $parsed 'tutorial é XML bem formado (tags balanceadas)'

if ($parsed) {
    $ui = @($xml.SelectNodes("//*[contains(concat(' ', normalize-space(@class), ' '), ' ui ')]") | ForEach-Object { ($_.InnerText -replace '\s+', ' ').Trim() } | Sort-Object -Unique)
    Check ($ui.Count -ge 30) "tutorial cita muitos controles ($($ui.Count) rótulos distintos)"
    foreach ($l in $ui) { Check ($known.Contains($l)) "rótulo do tutorial existe na interface: '$l'" }
    # Todos os rótulos da faixa de opções (exceto itens de lista) aparecem no tutorial
    foreach ($n in $ribbon.SelectNodes('//*[self::button or self::toggleButton or self::dropDown or self::group or self::tab]')) {
        $l = $n.GetAttribute('label')
        Check ($ui -contains $l) "tutorial documenta o controle '$l'"
    }
    # Recursos externos
    $bad = @()
    foreach ($n in $xml.SelectNodes('//*[@src or @href]')) {
        foreach ($a in 'src', 'href') {
            $v = $n.GetAttribute($a)
            if ($v -match '^(https?:)?//') {
                $ok = $false
                foreach ($p in $allowedLinks) { if ($v.StartsWith($p)) { $ok = $true } }
                if (-not $ok -or $n.LocalName -ne 'a') { $bad += "$($n.LocalName) $a=$v" }
            }
        }
    }
    Check ($bad.Count -eq 0) "sem recursos externos (src/href http só em <a> para o repositório): $($bad -join '; ')"
    Check ($xml.SelectNodes('//script|//link|//img|//iframe|//object|//embed|//video|//audio|//source').Count -eq 0) 'sem script/link/img/iframe'
    foreach ($a in $xml.SelectNodes('//a[starts-with(@href,"#")]')) {
        $id = $a.GetAttribute('href').Substring(1)
        if ($xml.SelectSingleNode("//*[@id='$id']") -eq $null) { Check $false "âncora #$id existe" }
    }
    $ids = @($xml.SelectNodes('//*[@id]') | ForEach-Object { $_.GetAttribute('id') })
    Check ($ids.Count -eq @($ids | Sort-Object -Unique).Count) 'ids únicos'
    Check ($xml.SelectNodes('//h1').Count -eq 1) 'um único h1'
}
$cssExternal = [regex]::Matches($html, 'url\(\s*["'']?(https?:)?//|@import') | Measure-Object
Check ($cssExternal.Count -eq 0) 'CSS sem url() externa nem @import'

# ---- README, LEIA-ME: caminhos de menu "ThinQQ > Rótulo" existem
$sorted = @($known | Sort-Object { -$_.Length })
foreach ($doc in @(@{ n = 'README.md'; p = $readmePath }, @{ n = 'LEIA-ME.txt'; p = $leiamePath })) {
    Check (Test-Path $doc.p) "$($doc.n) existe"
    $t = (Read-Utf8 $doc.p) -replace '\s+', ' '
    foreach ($m in [regex]::Matches($t, 'ThinQQ\s>\s(?<rest>[^|`"]{1,80})')) {
        $rest = $m.Groups['rest'].Value
        $hit = $false
        foreach ($l in $sorted) { if ($rest.StartsWith($l)) { $hit = $true; break } }
        Check $hit "$($doc.n): caminho de menu válido: 'ThinQQ > $($rest.Substring(0, [Math]::Min(30, $rest.Length)))'"
    }
}
$readme = Read-Utf8 $readmePath
Check ($readme -match 'https://github\.com/israeldias-hub/ThinQQ') 'README cita o repositório real'
Check ($readme -match 'AccessVBOM|Confiar no acesso ao modelo de objeto') 'README cita a configuração de confiança do VBA'
Check ($readme -match 'Settings.*Pages|Pages' -and $readme -match 'github\.io/ThinQQ/tutorial/tutorial\.html') 'README cita o GitHub Pages do tutorial'
Check ($readme -cnotmatch 'TODO|TBD|FIXME') 'README sem pendências'
# Parâmetros do release citados existem no script
$rel = Read-Utf8 (Join-Path $RepoRoot 'build\release.ps1')
foreach ($p in 'Version', 'BaseUrl', 'PublishTo') { Check ($rel -match "\`$$p" -and $readme -match "-$p") "README e release.ps1 concordam no parâmetro -$p" }

# ---- CHANGELOG
$cl = Read-Utf8 $changelogPath
$ver = ((Read-Utf8 (Join-Path $RepoRoot 'version.json')) | ConvertFrom-Json).core
Check ($cl -match "(?m)^## \[$([regex]::Escape($ver))\] - \d{4}-\d{2}-\d{2}") "CHANGELOG tem a seção da versão $ver"

if ($script:fail -gt 0) { Write-Host "FALHAS Test-Docs: $script:fail"; exit 1 }
Write-Host 'Test-Docs: OK'
