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
Check ($known.Count -ge 60) "rótulos conhecidos carregados ($($known.Count))"
Check ($known.Contains('Colunas empilhadas') -and $known.Contains('Pular esta versão') -and $known.Contains('OK')) 'faixa e formulários presentes no conjunto'

# ---- tutorial
Check (Test-Path $tutorialPath) 'tutorial.html existe'
$html = Read-Utf8 $tutorialPath
$bytes = (Get-Item $tutorialPath).Length
Check ($html -notmatch 'THINQQ-TUTORIAL-STUB') 'tutorial sem marcador de esqueleto'
Check ($bytes -ge 2000) "tutorial com tamanho mínimo ($bytes bytes)"
Check ($bytes -le 122880) "tutorial com até 120 KB ($bytes bytes)"
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
    Check ($ui.Count -ge 45) "tutorial cita muitos controles ($($ui.Count) rótulos distintos)"
    foreach ($l in $ui) { Check ($known.Contains($l)) "rótulo do tutorial existe na interface: '$l'" }
    # Todos os rótulos da faixa de opções (exceto itens de lista) aparecem no tutorial: botões, toggles, listas, menus, grupos e a aba.
    # O customUI tem namespace padrão: o XPath precisa de local-name() (a forma 'self::button' não casa nenhum nó; já houve esse buraco).
    $ctlNodes = @($ribbon.SelectNodes("//*[local-name()='button' or local-name()='toggleButton' or local-name()='dropDown' or local-name()='menu' or local-name()='group' or local-name()='tab']"))
    Check ($ctlNodes.Count -ge 40) "faixa: controles encontrados pelo XPath com local-name() ($($ctlNodes.Count))"
    foreach ($n in $ctlNodes) {
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
    Check ($xml.SelectNodes('//link|//img|//iframe|//object|//embed|//video|//audio|//source').Count -eq 0) 'sem link/img/iframe/object/embed/video/audio/source'
    # Único script permitido: o inline que abre os <details> ao imprimir (sem src, sem rede, sem eval)
    $scripts = @($xml.SelectNodes('//script'))
    Check ($scripts.Count -eq 1) 'exatamente um script inline'
    if ($scripts.Count -eq 1) {
        $js = $scripts[0].InnerText
        Check (-not $scripts[0].HasAttribute('src') -and $js -match 'beforeprint' -and $js -match 'afterprint') 'script sem src, tratando beforeprint/afterprint'
        Check ($js -notmatch 'http|fetch|XMLHttpRequest|WebSocket|import|eval|Function\(|document\.write|location|cookie|localStorage') 'script sem rede, eval nem armazenamento'
    }
    foreach ($a in $xml.SelectNodes('//a[starts-with(@href,"#")]')) {
        $id = $a.GetAttribute('href').Substring(1)
        if ($xml.SelectSingleNode("//*[@id='$id']") -eq $null) { Check $false "âncora #$id existe" }
    }
    $ids = @($xml.SelectNodes('//*[@id]') | ForEach-Object { $_.GetAttribute('id') })
    Check ($ids.Count -eq @($ids | Sort-Object -Unique).Count) 'ids únicos'
    Check ($xml.SelectNodes('//h1').Count -eq 1) 'um único h1'

    # Uma seção por assunto, todas no índice ("Neste tutorial"): 17 tipos navegáveis sem rolar a página inteira
    foreach ($id in 'tipos', 'barras', 'cascata', 'linha-area', 'mekko', 'gantt', 'pizza', 'decoracoes') {
        Check ($null -ne $xml.SelectSingleNode("//*[local-name()='section'][@id='$id']")) "tutorial tem a seção #$id"
        Check ($null -ne $xml.SelectSingleNode("//nav//a[@href='#$id']")) "o índice leva à seção #$id"
    }
    Check ($null -ne $xml.SelectSingleNode("//*[@id='matriz']")) 'tutorial tem a tabela "O que cada tipo aceita" (#matriz)'
    $noToc = @($xml.SelectNodes("//*[local-name()='section'][@id]") | ForEach-Object { $_.GetAttribute('id') } | Where-Object { $null -eq $xml.SelectSingleNode("//nav//a[@href='#$_']") })
    Check ($noToc.Count -eq 0) "toda seção do tutorial está no índice: $($noToc -join ', ')"

    # A matriz de capacidades do tutorial não pode divergir do código: é lida de Kinds_Supports (modChartKinds.bas).
    $kindsSrc = Read-Utf8 (Join-Path $RepoRoot 'src\core\modules\modChartKinds.bas')
    $allKinds = @(([regex]::Match($kindsSrc, 'Private Const KINDS As String = "(?<k>[^"]+)"').Groups['k'].Value).Trim(',').Split(','))
    $matrix = @{}
    foreach ($m in [regex]::Matches($kindsSrc, '(?m)^\s*Case "(?<f>[a-z0-9]+)":\s*list = (?<v>"[^"]*"|allKinds|Replace\(allKinds,\s*"[^"]*",\s*""\))')) {
        $v = $m.Groups['v'].Value
        if ($v.StartsWith('"')) { $list = $v.Trim('"').Split(',') }
        elseif ($v -eq 'allKinds') { $list = $allKinds }
        else { $rep = [regex]::Match($v, '"(?<r>[^"]*)",').Groups['r'].Value; $list = ($allKinds -join ',').Replace($rep, '').Split(',') }
        $matrix[$m.Groups['f'].Value] = $list
    }
    Check ($allKinds.Count -eq 17) "17 tipos de gráfico em modChartKinds.bas ($($allKinds.Count))"
    Check ($matrix.Count -eq 16) "16 recursos na matriz de capacidades de modChartKinds.bas ($($matrix.Count))"
    # A matriz pode ser dividida em tabelas (mesma linha de tipo em cada uma): cada tipo aparece com todos os recursos, cada um uma vez
    $rows = @($xml.SelectNodes('//tr[@data-kind]'))
    $kindsInDoc = @($rows | ForEach-Object { $_.GetAttribute('data-kind') } | Sort-Object -Unique)
    Check ($kindsInDoc.Count -eq $allKinds.Count -and @($allKinds | Where-Object { $kindsInDoc -notcontains $_ }).Count -eq 0) "matriz do tutorial: uma linha por tipo ($($kindsInDoc.Count) de $($allKinds.Count))"
    $badCells = @()
    foreach ($k in $kindsInDoc) {
        if ($allKinds -notcontains $k) { $badCells += "tipo desconhecido '$k'"; continue }
        $kindRows = @($rows | Where-Object { $_.GetAttribute('data-kind') -eq $k })
        foreach ($f in $matrix.Keys) {
            $cells = @($kindRows | ForEach-Object { $_.SelectNodes("td[@data-f='$f']") } | Where-Object { $_ })
            if ($cells.Count -ne 1) { $badCells += "$k/$f com $($cells.Count) célula(s)"; continue }
            $txt = ($cells[0].InnerText -replace '\s+', ' ').Trim()
            $want = ($matrix[$f] -contains $k)
            if ($f -eq 'legend' -and $want -and ($k -eq 'waterfall' -or $k -eq 'barwaterfall')) { $ok = ($txt -like 'Várias*') }   # só com várias séries
            elseif ($want) { $ok = ($txt -eq 'Sim') } else { $ok = ($txt -eq [string][char]0x2014) }
            if (-not $ok) { $badCells += "$k/$f='$txt' (código: $want)" }
        }
    }
    Check ($badCells.Count -eq 0) "matriz do tutorial confere com Kinds_Supports: $($badCells -join '; ')"
}
$cssExternal = [regex]::Matches($html, 'url\(\s*["'']?(https?:)?//|@import') | Measure-Object
Check ($cssExternal.Count -eq 0) 'CSS sem url() externa nem @import'

# Impressão: variáveis de tema restauradas e URLs externas visíveis no papel; details abertos
$printCss = ''; if ($html -match '(?s)@media print \{\s*:root \{(?<v>[^}]*)\}') { $printCss = $Matches['v'] }
foreach ($v in '--navy', '--blue', '--warn-bg', '--warn-ink', '--warn-line', '--ok-bg', '--ok-ink', '--ok-line', '--bg', '--ink', '--link', '--ui-bg') { Check ($printCss -match [regex]::Escape($v + ':')) "impressão redefine $v" }
Check ($html -match 'a\[href\^="http"\]::after') 'impressão mostra a URL dos links externos'

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
Check ($cl -match '(?m)^## \[1\.1\.0\] - \d{4}-\d{2}-\d{2}') 'CHANGELOG tem a seção 1.1.0'
Check ($cl -match '(?m)^## \[1\.0\.0\] - \d{4}-\d{2}-\d{2}') 'CHANGELOG mantém a seção 1.0.0'
Check ($cl -match '(?m)^## \[Não lançado\]\s+## \[') 'CHANGELOG mantém a seção [Não lançado] vazia no topo'

# ---- versões: a faixa em menus (Fase 2) exige o loader 1.1.0 ou mais novo; o tutorial, o README e o LEIA-ME seguem o version.json
$vj = (Read-Utf8 (Join-Path $RepoRoot 'version.json')) | ConvertFrom-Json
Check ($vj.loader -match '^\d+\.\d+\.\d+$' -and [version]$vj.loader -ge [version]'1.1.0') "loader 1.1.0 ou mais novo ($($vj.loader))"
Check ($ver -match '^\d+\.\d+\.\d+$' -and [version]$ver -ge [version]'1.1.0') "core 1.1.0 ou mais novo ($ver)"
Check ($html -match "ThinQQ $([regex]::Escape($ver))\.") "rodapé do tutorial mostra a versão $ver"
Check ($readme -match "v$([regex]::Escape($ver))") "README cita a release v$ver"
Check ($readme -notmatch 'v1\.0\.0') 'README não cita mais a release v1.0.0'
Check ((Read-Utf8 $leiamePath) -match [regex]::Escape("versions\$ver\")) "LEIA-ME cita a pasta versions\$ver\ no exemplo"
$notesLines = @((Read-Utf8 (Join-Path $RepoRoot 'release\notes.txt')) -split "\r?\n" | Where-Object { $_.Trim().Length -gt 0 })
Check ($notesLines.Count -ge 6 -and $notesLines.Count -le 8) "release/notes.txt tem de 6 a 8 linhas ($($notesLines.Count))"
Check (@($notesLines | Where-Object { -not $_.StartsWith([string][char]0x2022) }).Count -eq 0) 'release/notes.txt: toda linha começa com o marcador'

# ---- exemplos: a contagem citada na documentação é a de Samples_Create
$samplesSrc = Read-Utf8 (Join-Path $RepoRoot 'src\core\modules\modSamples.bas')
$exNums = @([regex]::Matches($samplesSrc, 'WriteTitle ws, \d+, "(?<n>\d+)\)') | ForEach-Object { [int]$_.Groups['n'].Value })
$nEx = ($exNums | Measure-Object -Maximum).Maximum
Check ($nEx -ge 18) "Samples_Create tem os exemplos numerados até $nEx"
Check ($html -match "$nEx exemplos" -and $readme -match "$nEx exemplos") "tutorial e README citam $nEx exemplos"
Check ($html -notmatch 'seis tabelas' -and $html -notmatch 'ainda não existem') 'tutorial sem a contagem nem a limitação da Fase 1'
Check ($readme -notmatch 'ficam para a Fase 2') 'README sem a limitação "ficam para a Fase 2"'

# ---- mensagens que o tutorial cita: precisam existir no tutorial E no código do core (texto exato)
$coreSrc = (Get-ChildItem (Join-Path $RepoRoot 'src\core') -Recurse -Include '*.bas', '*.vba' | ForEach-Object { Read-Utf8 $_.FullName }) -join "`n"
$messages = @(
    'A seta de diferença não está disponível para gráficos 100%',
    'A seta de diferença não está disponível para este tipo de gráfico',
    'A seta CAGR não está disponível para este tipo de gráfico',
    'O colchete só está disponível em gráficos Gantt',
    'A cascata aceita no máximo 6 séries',
    'A combinação precisa de pelo menos 2 colunas de valores',
    'Mekko (unidades): a segunda coluna da tabela é a largura; são necessárias pelo menos 3 colunas',
    'Gantt: a tabela precisa de 3 colunas: atividade, início e fim (datas)',
    'Gantt: a coluna de início não tem datas válidas',
    'não se aplica a este tipo de gráfico',
    'Rótulos no fim da linha só existem em gráficos de linha e combinação',
    'Mesma escala: selecione dois ou mais gráficos do ThinQQ',
    'Mesma escala: são necessários dois ou mais gráficos com escala livre',
    'Os dados deste gráfico passaram do máximo da escala fixa',
    'A tabela de origem mudou de tamanho',
    'mostra só a primeira coluna de valores',
    'O gráfico precisa de pelo menos 2 categorias'
)
foreach ($msg in $messages) {
    Check ($html -match [regex]::Escape($msg)) "tutorial cita a mensagem '$msg'"
    Check ($coreSrc -match [regex]::Escape($msg)) "o core emite a mensagem '$msg'"
}

if ($script:fail -gt 0) { Write-Host "FALHAS Test-Docs: $script:fail"; exit 1 }
Write-Host 'Test-Docs: OK'
