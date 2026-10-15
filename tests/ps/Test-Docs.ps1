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
$pptRibbonPath = Join-Path $RepoRoot 'src\ppt\ribbon\customUI14.xml'
$allowedLinks = @('https://github.com/israeldias-hub/ThinQQ', 'https://israeldias-hub.github.io/ThinQQ/')

# ---- rótulos conhecidos: faixa de opções + formulários (loader e core)
$known = New-Object 'System.Collections.Generic.HashSet[string]'
$ribbon = New-Object Xml.XmlDocument
$ribbon.LoadXml((Read-Utf8 $ribbonPath))
foreach ($n in $ribbon.SelectNodes('//*[@label]')) { [void]$known.Add($n.GetAttribute('label')) }
# faixa do suplemento do PowerPoint (ThinQQ.ppam): os rótulos também são "conhecidos" (o tutorial e o README os citam), mas a regra
# "todo controle da faixa aparece no tutorial" abaixo vale só para a faixa do Excel; a do PowerPoint tem a regra própria, na seção #powerpoint
$pptRibbon = New-Object Xml.XmlDocument
$pptRibbon.LoadXml((Read-Utf8 $pptRibbonPath))
foreach ($n in $pptRibbon.SelectNodes('//*[@label]')) { [void]$known.Add($n.GetAttribute('label')) }
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
Check ($known.Contains('Ao abrir') -and $known.Contains('Atualizar sempre') -and $known.Contains('Gráficos do Excel')) 'faixa do PowerPoint presente no conjunto'

# ---- tutorial
Check (Test-Path $tutorialPath) 'tutorial.html existe'
$html = Read-Utf8 $tutorialPath
$bytes = (Get-Item $tutorialPath).Length
Check ($html -notmatch 'THINQQ-TUTORIAL-STUB') 'tutorial sem marcador de esqueleto'
Check ($bytes -ge 2000) "tutorial com tamanho mínimo ($bytes bytes)"
# limite de 640 KB (era 192 KB): o tutorial e um arquivo unico e autocontido e a galeria embute as imagens dos exemplos em base64 (~320 KB)
Check ($bytes -le 655360) "tutorial com até 640 KB ($bytes bytes)"
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
    # Faixa do suplemento do PowerPoint (src\ppt\ribbon\customUI14.xml): a aba, o grupo, os botões, a lista "Ao abrir" e cada item dela são citados,
    # com o rótulo exato, DENTRO da seção #powerpoint (e não só em qualquer parte do tutorial); o botão do Excel (btnPptSync) também.
    $pptSec = $xml.SelectSingleNode("//*[local-name()='section'][@id='powerpoint']")
    $pptUi = @()
    if ($pptSec) { $pptUi = @($pptSec.SelectNodes(".//*[contains(concat(' ', normalize-space(@class), ' '), ' ui ')]") | ForEach-Object { ($_.InnerText -replace '\s+', ' ').Trim() } | Sort-Object -Unique) }
    $pptCtl = @($pptRibbon.SelectNodes("//*[local-name()='tab' or local-name()='group' or local-name()='button' or local-name()='dropDown' or local-name()='item']"))
    Check ($pptCtl.Count -eq 8) "faixa do PowerPoint: 8 controles (aba, grupo, 2 botões, lista e 3 itens: $($pptCtl.Count))"
    foreach ($n in $pptCtl) {
        $l = $n.GetAttribute('label')
        Check ($pptUi -contains $l) "#powerpoint cita o controle da faixa do PowerPoint '$l'"
    }
    $xlSync = $ribbon.SelectSingleNode("//*[@id='btnPptSync']")
    Check ($null -ne $xlSync -and $pptUi -contains $xlSync.GetAttribute('label')) '#powerpoint cita o botão do Excel btnPptSync com o rótulo da faixa'
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
    Check ($xml.SelectNodes('//link|//iframe|//object|//embed|//video|//audio|//source').Count -eq 0) 'sem link/iframe/object/embed/video/audio/source'
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
    # Fase 3 (1.2.0): seções novas, todas no índice
    foreach ($id in 'borboleta', 'dispersao', 'arredondamento', 'corte', 'powerpoint', 'demonstracao', 'instalacao-ti') {
        Check ($null -ne $xml.SelectSingleNode("//*[local-name()='section'][@id='$id']")) "tutorial tem a seção #$id"
        Check ($null -ne $xml.SelectSingleNode("//nav//a[@href='#$id']")) "o índice leva à seção #$id"
    }
    # a tabela de #tipos liga as linhas novas às suas seções
    foreach ($id in 'borboleta', 'dispersao') {
        Check ($null -ne $xml.SelectSingleNode("//*[@id='tipos']//a[@href='#$id']")) "a tabela de #tipos leva à seção #$id"
    }
    # Cascata comparativa (1.3.0): seção própria, no índice, e a linha da tabela de #tipos leva a ela
    Check ($null -ne $xml.SelectSingleNode("//*[local-name()='section'][@id='cascata-comparativa']")) 'tutorial tem a seção #cascata-comparativa'
    Check ($null -ne $xml.SelectSingleNode("//nav//a[@href='#cascata-comparativa']")) 'o índice leva à seção #cascata-comparativa'
    Check ($null -ne $xml.SelectSingleNode("//*[@id='tipos']//a[@href='#cascata-comparativa']")) 'a tabela de #tipos leva à seção #cascata-comparativa'
    # #demonstracao: uma linha por exemplo da lista de cobertura (a mesma que o gerador e o Test-DemoWorkbook usam): número, título e tipo
    . (Join-Path $RepoRoot 'tools\DemoCoverage.ps1')
    $demoCov = @(Get-TqqDemoCoverage)
    $demoSec = $xml.SelectSingleNode("//*[local-name()='section'][@id='demonstracao']")
    if ($demoSec) {
        $demoRows = @($demoSec.SelectNodes(".//tr[@data-ex]"))
        Check ($demoRows.Count -eq $demoCov.Count) "#demonstracao tem uma linha por exemplo da cobertura ($($demoRows.Count) de $($demoCov.Count))"
        foreach ($it in $demoCov) {
            $row = $demoSec.SelectSingleNode(".//tr[@data-ex='$($it.Sheet)']")
            $cells = @(); if ($row) { $cells = @($row.SelectNodes('th|td') | ForEach-Object { ($_.InnerText -replace '\s+', ' ').Trim() }) }
            Check ($cells.Count -eq 3 -and $cells[0] -eq $it.Sheet -and $cells[1] -eq $it.Title -and $cells[2] -eq $it.KindName) "#demonstracao lista o exemplo $($it.Sheet) com o título e o tipo da cobertura"
        }
    }
    # Galeria dos exemplos (tools\Build-TutorialGallery.ps1): uma <figure id="ex-E05"> por exemplo da cobertura, com a(s) imagem(ns) embutida(s)
    # (data URI image/png) e a legenda com o título ATUAL da cobertura: se alguém mudar o título no gerador do demo, o teste falha e pede para
    # refazer a galeria (tools\Export-DemoChartImages.ps1 e tools\Build-TutorialGallery.ps1). Os exemplos com "samescale" têm 2 gráficos (E25).
    Check ($html.Split(@('<!-- galeria:inicio -->'), [StringSplitOptions]::None).Count -eq 2 -and $html.Split(@('<!-- galeria:fim -->'), [StringSplitOptions]::None).Count -eq 2) 'galeria: um par de marcadores <!-- galeria:inicio/fim -->'
    $galFigs = @($xml.SelectNodes("//figure[starts-with(@id,'ex-')]"))
    Check ($galFigs.Count -eq $demoCov.Count) "galeria: uma <figure> por exemplo da cobertura ($($galFigs.Count) de $($demoCov.Count))"
    $expectedImgs = $demoCov.Count + @($demoCov | Where-Object { $_.Features -contains 'samescale' }).Count
    $allImgs = @($xml.SelectNodes('//img'))
    Check ($allImgs.Count -eq $expectedImgs) "galeria: $expectedImgs imagens no tutorial, uma por gráfico ($($allImgs.Count))"
    Check ($xml.SelectNodes('//img[not(ancestor::figure[starts-with(@id,"ex-")])]').Count -eq 0) 'galeria: nenhum <img> fora das figuras ex-E..'
    $galImgBytes = 0
    $pngSig = [byte[]](137, 80, 78, 71, 13, 10, 26, 10)
    foreach ($it in $demoCov) {
        $fig = $xml.SelectSingleNode("//figure[@id='ex-$($it.Sheet)']")
        if ($null -eq $fig) { Check $false "galeria: figura ex-$($it.Sheet) existe (refaça a galeria)"; continue }
        $imgs = @($fig.SelectNodes('.//img'))
        $nImgs = 1; if ($it.Features -contains 'samescale') { $nImgs = 2 }
        Check ($imgs.Count -eq $nImgs) "galeria $($it.Sheet): $nImgs imagem(ns) ($($imgs.Count))"
        $cap = $fig.SelectSingleNode('figcaption')
        $capText = ''; if ($cap) { $capText = ($cap.InnerText -replace '\s+', ' ').Trim() }
        Check ($capText.StartsWith($it.Sheet) -and $capText.Contains($it.Title) -and $capText.Contains($it.Desc)) "galeria $($it.Sheet): legenda cita o título e a descrição atuais da cobertura (senão refaça a galeria com tools\Build-TutorialGallery.ps1)"
        foreach ($img in $imgs) {
            $src = $img.GetAttribute('src')
            $okImg = $false; $info = ''
            if ($src -cmatch '^data:image/png;base64,(?<b>[A-Za-z0-9+/]+={0,2})$') {
                try {
                    $png = [Convert]::FromBase64String($Matches['b'])
                    $sigOk = $png.Length -gt 33
                    if ($sigOk) { for ($i = 0; $i -lt 8; $i++) { if ($png[$i] -ne $pngSig[$i]) { $sigOk = $false } } }
                    if ($sigOk) {
                        $w = ([int]$png[16] -shl 24) -bor ([int]$png[17] -shl 16) -bor ([int]$png[18] -shl 8) -bor [int]$png[19]
                        $h = ([int]$png[20] -shl 24) -bor ([int]$png[21] -shl 16) -bor ([int]$png[22] -shl 8) -bor [int]$png[23]
                        $okImg = ($w -gt 0 -and $h -gt 0 -and $img.GetAttribute('width') -eq [string]$w -and $img.GetAttribute('height') -eq [string]$h -and $png.Length -le 61440)
                        $info = "${w}x$h, $($png.Length) bytes"
                        $galImgBytes += $png.Length
                    }
                } catch { $info = $_.Exception.Message }
            }
            Check $okImg "galeria $($it.Sheet): imagem é data URI PNG decodificável, width/height iguais ao IHDR, até 60 KB ($info)"
            Check ($img.GetAttribute('alt').Trim().Length -gt 0) "galeria $($it.Sheet): imagem com alt não vazio"
        }
    }
    Check ($galImgBytes -le 716800) "galeria: total das imagens até 700 KB ($galImgBytes bytes)"
    # as imagens ficam embutidas: nenhum src/href de arquivo ou de rede (a regra geral de http acima só olha http/https; aqui também file: e relativos)
    $badSrc = @($xml.SelectNodes('//*[@src]') | Where-Object { $_.GetAttribute('src') -notmatch '^data:image/png;base64,' } | ForEach-Object { $_.LocalName + ' src=' + $_.GetAttribute('src') })
    Check ($badSrc.Count -eq 0) "galeria: todo src do tutorial é data URI image/png ($($badSrc -join '; '))"
    Check (@($xml.SelectNodes('//*[@href]') | Where-Object { $_.GetAttribute('href') -match '^(?i)(file:|[a-z]:\|\\)' }).Count -eq 0) 'galeria: nenhum href para file: nem caminho local'
    # âncoras: a tabela da demonstração liga cada número à sua figura e o texto do tutorial remete a toda figura (os links #ex-E.. a ids inexistentes
    # já falham na regra geral das âncoras acima)
    foreach ($it in $demoCov) {
        Check ($null -ne $demoSec.SelectSingleNode(".//tr[@data-ex='$($it.Sheet)']/th/a[@href='#ex-$($it.Sheet)']")) "galeria: o número $($it.Sheet) na tabela é link para a figura"
        $outside = @($xml.SelectNodes("//a[@href='#ex-$($it.Sheet)']") | Where-Object { $null -eq $_.SelectSingleNode("ancestor::*[local-name()='section'][@id='demonstracao']") })
        Check ($outside.Count -ge 1) "galeria: alguma seção do tutorial remete à figura $($it.Sheet) (Veja a figura $($it.Sheet) na galeria)"
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
    Check ($allKinds.Count -eq 21) "21 tipos de gráfico em modChartKinds.bas ($($allKinds.Count))"
    Check ($matrix.Count -eq 23) "23 recursos na matriz de capacidades de modChartKinds.bas ($($matrix.Count))"
    $nk = $allKinds.Count
    Check ($html -match "\b$nk tipos" -and $html -notmatch "\b$($nk - 1) tipos") "tutorial cita os $nk tipos (e nenhum '$($nk - 1) tipos')"
    # Rótulo de cada recurso: Kinds_FeatureLabel (modChartKinds.bas) ou, nos que ela não cobre, o controle da faixa que o aciona
    $featureLabel = @{}
    foreach ($m in [regex]::Matches($kindsSrc, 'Case "(?<f>[a-z0-9]+)":\s*Kinds_FeatureLabel = "(?<l>[^"]*)"')) { $featureLabel[$m.Groups['f'].Value] = $m.Groups['l'].Value }
    foreach ($f in $matrix.Keys) {
        if ($featureLabel.ContainsKey($f)) { continue }
        $node = switch ($f) {
            'lineend'   { $ribbon.SelectSingleNode("//*[local-name()='item'][@id='lineend']") }
            'samescale' { $ribbon.SelectSingleNode("//*[@id='btnSameScale']") }
            default     { $ribbon.SelectSingleNode("//*[@tag='$f']") }
        }
        if ($node) { $featureLabel[$f] = $node.GetAttribute('label') }
    }
    Check (@($matrix.Keys | Where-Object { -not $featureLabel.ContainsKey($_) }).Count -eq 0) 'todo recurso da matriz tem rótulo no código ou na faixa'
    # A matriz pode ser dividida em tabelas (class="mx"). O recurso de cada coluna vem do CABEÇALHO (th data-f), que precisa mostrar o rótulo
    # da faixa; a célula é associada ao recurso pela posição da coluna, e não por um atributo próprio (um atributo por célula podia
    # contradizer o cabeçalho visível sem que o teste visse).
    $tables = @($xml.SelectNodes("//table[contains(concat(' ', normalize-space(@class), ' '), ' mx ')]"))
    Check ($tables.Count -ge 1) "tutorial tem tabelas da matriz de capacidades ($($tables.Count))"
    $badCells = @(); $seenFeature = @{}; $cellText = @{}; $kindsInDoc = @{}
    foreach ($tb in $tables) {
        $heads = @($tb.SelectNodes('thead/tr/th[@data-f]'))
        if (@($tb.SelectNodes('thead/tr/th')).Count -ne $heads.Count + 1) { $badCells += 'cabeçalho com coluna sem data-f (só a 1.ª, "Tipo", fica sem)' }
        foreach ($h in $heads) {
            $f = $h.GetAttribute('data-f')
            $shown = ($h.InnerText -replace '\s+', ' ').Trim()
            if (-not $matrix.ContainsKey($f)) { $badCells += "cabeçalho com recurso desconhecido '$f'" }
            elseif ($shown -ne $featureLabel[$f]) { $badCells += "cabeçalho '$shown' está marcado como '$f', cujo rótulo na faixa é '$($featureLabel[$f])'" }
            if ($seenFeature.ContainsKey($f)) { $badCells += "recurso '$f' aparece em mais de uma coluna" }
            $seenFeature[$f] = $true
        }
        foreach ($row in @($tb.SelectNodes('tbody/tr[@data-kind]'))) {
            $k = $row.GetAttribute('data-kind')
            if ($allKinds -notcontains $k) { $badCells += "tipo desconhecido '$k'"; continue }
            $kindsInDoc[$k] = $true
            $kindNode = $ribbon.SelectSingleNode("//*[local-name()='button'][@tag='$k']")
            $rowShown = ($row.SelectSingleNode("th[@scope='row']").InnerText -replace '\s+', ' ').Trim()
            if ($kindNode -and $rowShown -ne $kindNode.GetAttribute('label')) { $badCells += "linha '$rowShown' está marcada como '$k', cujo rótulo na faixa é '$($kindNode.GetAttribute('label'))'" }
            $cells = @($row.SelectNodes('td'))
            if ($cells.Count -ne $heads.Count) { $badCells += "${k}: $($cells.Count) células para $($heads.Count) colunas"; continue }
            for ($i = 0; $i -lt $cells.Count; $i++) {
                $key = "$k/$($heads[$i].GetAttribute('data-f'))"
                if ($cellText.ContainsKey($key)) { $badCells += "$key duplicada" }
                $cellText[$key] = ($cells[$i].InnerText -replace '\s+', ' ').Trim()
            }
        }
    }
    Check ($kindsInDoc.Count -eq $allKinds.Count -and @($allKinds | Where-Object { -not $kindsInDoc.ContainsKey($_) }).Count -eq 0) "matriz do tutorial: uma linha por tipo ($($kindsInDoc.Count) de $($allKinds.Count))"
    Check (@($matrix.Keys | Where-Object { -not $seenFeature.ContainsKey($_) }).Count -eq 0) 'matriz do tutorial: uma coluna por recurso'
    foreach ($k in $allKinds) {
        foreach ($f in $matrix.Keys) {
            $txt = $cellText["$k/$f"]
            $want = ($matrix[$f] -contains $k)
            if ($f -eq 'legend' -and $want -and ($k -eq 'waterfall' -or $k -eq 'barwaterfall')) { $ok = ($txt -like 'Várias*') }   # só com várias séries
            elseif ($want) { $ok = ($txt -eq 'Sim') } else { $ok = ($txt -eq [string][char]0x2014) }
            if (-not $ok) { $badCells += "$k/$f='$txt' (código: $want)" }
        }
    }
    Check ($badCells.Count -eq 0) "matriz do tutorial confere com Kinds_Supports e com os rótulos da faixa: $($badCells -join '; ')"
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

# ---- README: a seção "Instalação" (para quem não é técnico) e a subseção para a TI (Fase 3, §4.5 e §4.9 da especificação)
Check ($readme -match '(?m)^## Instalação') 'README tem a seção "Instalação"'
Check ($readme -notmatch 'Instalação \(usuários\)') 'README sem o título antigo "Instalação (usuários)"'
foreach ($t in 'Instalar-ThinQQ.cmd', 'Desbloquear', 'Central de Confiabilidade', 'Local Confiável', 'antivírus', 'Desinstala', 'Instalação manual', 'Releases') { Check ($readme -match [regex]::Escape($t)) "README/Instalação cita '$t'" }
foreach ($t in 'Requisitos', 'Passo a passo', 'Como as atualizações chegam', 'Problemas mais comuns') { Check ($readme -match "(?m)^###+ $([regex]::Escape($t))") "README/Instalação tem o subtítulo '$t'" }
Check ($readme -match '(?m)^### Ambientes corporativos \(TI\)') 'README tem "Ambientes corporativos (TI)"'
foreach ($t in 'Editores Confiáveis', 'tools/New-ThinQQCert.ps1', 'tools/Sign-ThinQQ.ps1', 'tools/Trust-ThinQQPublisher.ps1', 'Assinatura Digital', 'Subject Interface Packages', 'Only trust VBA macros that use V3 signatures', '-SignPfx', '-SkipVerify', 'ponta a ponta') { Check ($readme -cmatch [regex]::Escape($t)) "README/TI cita '$t'" }
# repositório privado: o que isso significa para quem instala e as saídas
foreach ($t in 'privado', 'colaborador', 'público', '-PublishTo', 'GitHub Pages') { Check ($readme -match [regex]::Escape($t)) "README/Instalação explica o repositório privado ('$t')" }
# pasta de demonstração: no ZIP, no README (com o link relativo para quem lê no GitHub) e no tutorial
Check ($readme -match 'ThinQQ-Demonstracao-Materiais-Construcao\.xlsx') 'README cita a pasta de demonstração'
Check ($readme -match '\]\(exemplos/ThinQQ-Demonstracao-Materiais-Construcao\.xlsx\)') 'README liga a pasta de demonstração por link relativo'
Check ($html -match 'ThinQQ-Demonstracao-Materiais-Construcao\.xlsx') 'tutorial cita a pasta de demonstração'
. (Join-Path $RepoRoot 'tools\DemoCoverage.ps1')
$nDemo = @(Get-TqqDemoCoverage).Count
$demoRange = 'E01 a E' + ('{0:00}' -f $nDemo)
Check ($nDemo -ge 26) "a cobertura da demonstração tem $nDemo exemplos"
Check ($readme -match "$nDemo exemplos" -and $html -match "$nDemo exemplos") "README e tutorial citam os $nDemo exemplos da demonstração"
Check ($readme -match [regex]::Escape($demoRange) -and $html -match [regex]::Escape($demoRange)) "README e tutorial citam '$demoRange'"
Check ((Read-Utf8 $leiamePath) -match [regex]::Escape("$nDemo exemplos ($demoRange)")) "LEIA-ME cita '$nDemo exemplos ($demoRange)'"
$leiameExPath = Join-Path $RepoRoot 'installer\LEIA-ME-exemplos.txt'
Check ((Read-Utf8 $leiameExPath) -match [regex]::Escape("$nDemo exemplos, de $demoRange")) "LEIA-ME dos exemplos cita '$nDemo exemplos, de $demoRange'"
Check ((Read-Utf8 $leiameExPath) -match [regex]::Escape("planilhas $demoRange")) "LEIA-ME dos exemplos cita 'planilhas $demoRange'"
# Parâmetros do release citados existem no script
$rel = Read-Utf8 (Join-Path $RepoRoot 'build\release.ps1')
foreach ($p in 'Version', 'BaseUrl', 'PublishTo') { Check ($rel -match "\`$$p" -and $readme -match "-$p") "README e release.ps1 concordam no parâmetro -$p" }
# Publicação automatizada (publicar-release.bat + tools/Publish-Release.ps1): README e release/README.md descrevem o fluxo em duas etapas,
# os parâmetros citados existem no script, e o procedimento manual antigo fica como alternativa
$relReadme = Read-Utf8 (Join-Path $RepoRoot 'release\README.md')
$pubScriptPath = Join-Path $RepoRoot 'tools\Publish-Release.ps1'
Check ((Test-Path -LiteralPath $pubScriptPath) -and (Test-Path -LiteralPath (Join-Path $RepoRoot 'publicar-release.bat'))) 'publicar-release.bat e tools/Publish-Release.ps1 existem'
$pubScript = if (Test-Path -LiteralPath $pubScriptPath) { Read-Utf8 $pubScriptPath } else { '' }
foreach ($doc in @(@{ n = 'README.md'; t = $readme }, @{ n = 'release/README.md'; t = $relReadme })) {
    foreach ($t in 'publicar-release.bat', 'tools/Publish-Release.ps1', 'publicar-release.bat x.y.z publicar', '-DryRun', '-Draft', 'alternativa manual', 'git credential fill', 'Create a merge commit', 'squash', 'digest', 'Set-PSDebug', 'Start-Transcript') {
        Check ($doc.t -match [regex]::Escape($t)) "$($doc.n) cita '$t' (publicação automatizada)"
    }
}
foreach ($p in 'Draft', 'DryRun', 'NoBrowser', 'RunTests', 'AllowMainChanges') { Check ($pubScript -match "\[switch\]\`$$p" -and $readme -match "-$p") "README e Publish-Release.ps1 concordam no parâmetro -$p" }

# ---- CHANGELOG
$cl = Read-Utf8 $changelogPath
$vj = $null
try { $vj = (Read-Utf8 (Join-Path $RepoRoot 'version.json')) | ConvertFrom-Json } catch { $vj = $null }
Check ($null -ne $vj) 'version.json é um JSON válido'
$ver = if ($vj) { [string]$vj.core } else { '' }
Check ($cl -match "(?m)^## \[$([regex]::Escape($ver))\] - \d{4}-\d{2}-\d{2}") "CHANGELOG tem a seção da versão $ver"
Check ($cl -match '(?m)^## \[1\.2\.0\] - \d{4}-\d{2}-\d{2}') 'CHANGELOG tem a seção 1.2.0'
Check ($cl -match '(?m)^## \[1\.1\.0\] - \d{4}-\d{2}-\d{2}') 'CHANGELOG tem a seção 1.1.0'
Check ($cl -match '(?m)^## \[1\.0\.0\] - \d{4}-\d{2}-\d{2}') 'CHANGELOG mantém a seção 1.0.0'
Check ($cl.IndexOf('## [Não lançado]') -ge 0 -and $cl.IndexOf('## [Não lançado]') -lt $cl.IndexOf('## [1.2.0]')) 'CHANGELOG tem a seção [Não lançado] no topo, antes das versões (pode ter entradas)'

# ---- versões: loader e core em x.y.z, a partir da 1.2.0 (a primeira publicada); o rodapé do tutorial e o exemplo do LEIA-ME seguem o
# version.json (o publicar-release.bat os atualiza junto com o version.json e o CHANGELOG)
$vjOk = ($null -ne $vj) -and ([string]$vj.loader -match '^\d{1,6}\.\d{1,6}\.\d{1,6}\z') -and ([string]$vj.core -match '^\d{1,6}\.\d{1,6}\.\d{1,6}\z')
Check $vjOk "version.json: loader e core em x.y.z ($(if ($vj) { "$($vj.loader), $($vj.core)" } else { 'ilegível' }))"
$vjMin = $false
if ($vjOk) { try { $vjMin = ([version]$vj.core) -ge [version]'1.2.0' -and ([version]$vj.loader) -ge [version]'1.2.0' } catch { $vjMin = $false } }
Check $vjMin 'version.json: loader e core a partir da 1.2.0'
Check ($html -match "ThinQQ $([regex]::Escape($ver))\.") "rodapé do tutorial mostra a versão $ver"
Check ($readme -match 'v1\.2\.0') 'README cita a release v1.2.0 (a primeira publicada)'
Check ($readme -notmatch 'v1\.0\.0') 'README não cita mais a release v1.0.0'
Check ((Read-Utf8 $leiamePath) -match [regex]::Escape("versions\$ver\")) "LEIA-ME cita a pasta versions\$ver\ no exemplo"
$notesLines = @((Read-Utf8 (Join-Path $RepoRoot 'release\notes.txt')) -split "\r?\n" | Where-Object { $_.Trim().Length -gt 0 })
Check ($notesLines.Count -ge 1 -and $notesLines.Count -le 8) "release/notes.txt tem de 1 a 8 linhas ($($notesLines.Count)): uma versao de correcao pode ter poucas"
Check (@($notesLines | Where-Object { -not $_.StartsWith([string][char]0x2022) }).Count -eq 0) 'release/notes.txt: toda linha começa com o marcador'

# ---- exemplos: a contagem citada na documentação é a de Samples_Create
$samplesSrc = Read-Utf8 (Join-Path $RepoRoot 'src\core\modules\modSamples.bas')
$exNums = @([regex]::Matches($samplesSrc, 'WriteTitle ws, \d+, "(?<n>\d+)\)') | ForEach-Object { [int]$_.Groups['n'].Value })
$nEx = ($exNums | Measure-Object -Maximum).Maximum
Check ($nEx -ge 26) "Samples_Create tem os exemplos numerados até $nEx"
Check ($html -match "$nEx exemplos" -and $readme -match "$nEx exemplos") "tutorial e README citam $nEx exemplos"
Check ($html -notmatch 'seis tabelas' -and $html -notmatch 'ainda não existem') 'tutorial sem a contagem nem a limitação da Fase 1'
Check ($readme -notmatch 'ficam para a Fase 2') 'README sem a limitação "ficam para a Fase 2"'

# ---- mensagens que o tutorial cita: precisam existir no tutorial E no código do core (texto exato)
$coreSrc = (Get-ChildItem (Join-Path $RepoRoot 'src\core') -Recurse -Include '*.bas', '*.vba' | ForEach-Object { Read-Utf8 $_.FullName }) -join "`n"
# Frases completas (as que não têm parte variável) e, nas que têm, os trechos fixos: antes só se conferia o começo da frase.
$messages = @(
    'A seta de diferença não está disponível para gráficos 100%.',
    'A seta de diferença não está disponível para este tipo de gráfico.',
    'A seta CAGR não está disponível para este tipo de gráfico.',
    'O colchete só está disponível em gráficos Gantt.',
    'A cascata aceita no máximo 6 séries.',
    'A combinação precisa de pelo menos 2 colunas de valores: as primeiras viram colunas empilhadas e a última vira linha.',
    'Mekko (unidades): a segunda coluna da tabela é a largura; são necessárias pelo menos 3 colunas.',
    'Gantt: a tabela precisa de 3 colunas: atividade, início e fim (datas). Deixe o fim vazio para um marco.',
    'Gantt: a coluna de início não tem datas válidas (use células de data, não texto).',
    'Rótulos no fim da linha só existem em gráficos de linha e combinação.',
    'Mesma escala: selecione dois ou mais gráficos do ThinQQ (ou nenhum, para usar todos os gráficos da planilha).',
    'Mesma escala: são necessários dois ou mais gráficos com escala livre (100%, Mekko (%), pizza e Gantt não entram).',
    'Os dados deste gráfico passaram do máximo da escala fixa (Mesma escala) e as barras ficam cortadas. Use ''Mesma escala'' de novo ou ''Escala automática''.',
    'A tabela de origem mudou de tamanho. Use ''Alterar dados'' para reconstruir o gráfico (ou reinsira o gráfico).',
    'O gráfico precisa de pelo menos 2 categorias.',
    'O gráfico precisa de pelo menos 1 atividade.',
    'não se aplica a este tipo de gráfico',
    'mostra só a primeira coluna de valores (',
    '). As outras ',
    ' coluna(s) da tabela foram ignoradas.',
    'A borboleta precisa de exatamente 2 colunas de valores (lado esquerdo e lado direito).',
    'Dispersão: a tabela precisa de nome, X e Y (e, nas bolhas, tamanho).',
    'No máximo 6 grupos (coluna de grupo).',
    'Corte de eixo: informe um valor maior que zero e menor que o maior valor do gráfico.',
    'ficaria quase invisível (menos de um quarto dele abaixo do corte).',
    'O PowerPoint não está instalado neste computador (ou não pôde ser iniciado).',
    'Salve a pasta de trabalho antes de copiar o gráfico vinculado: o vínculo aponta para o arquivo salvo.',
    'A pasta tem alterações não salvas: enquanto ela estiver aberta no Excel o vínculo lê esses valores; fechada, o PowerPoint lê o arquivo salvo.',
    'A área de transferência do Windows não está disponível (outro programa a mantém aberta). Tente novamente em instantes.',
    'O PowerPoint está ocupado (uma caixa de diálogo aberta ou um texto em edição). Termine o que está fazendo nele e tente de novo.',
    'Selecione uma única coluna ou linha só com números (pelo menos 2).',
    'Fórmulas de arredondamento com soma consistente gravadas em',
    ' já têm conteúdo. Substituir? Esta ação não pode ser desfeita.',
    'A planilha de destino está protegida sem a permissão de formatar células: desproteja-a ou escolha um destino em outra planilha.',
    'células; o Arredondar intervalo aceita no máximo ',
    'O gráfico saiu do grupo de Mesma escala: corte de eixo e Mesma escala são excludentes.',
    'As setas de variação passam do máximo da escala fixa (Mesma escala) e ficam cortadas. Use ''Escala automática'' ou desligue as setas de variação.',
    'A cascata comparativa precisa de exatamente 2 colunas de valores (uma por série); a tabela tem só 1.',
    'A cascata comparativa compara exatamente 2 séries; a tabela tem ',
    ' colunas de valores. Selecione a coluna das etapas e só as duas colunas a comparar.'
)
# a nota da colagem e o tutorial nao podem voltar a dizer que o vinculo so le o arquivo salvo (a sonda mediu: com a pasta ABERTA no Excel,
# LinkFormat.Update le os valores nao salvos)
Check ($html -notmatch [regex]::Escape('o PowerPoint lê o arquivo salvo nas atualizações do vínculo')) 'tutorial sem a nota antiga da colagem (pasta aberta le os valores nao salvos)'
Check ($html -notmatch [regex]::Escape('Atualizar Dados</strong> relê o arquivo salvo')) 'tutorial: Atualizar Dados qualificado (pasta aberta x fechada)'
foreach ($msg in $messages) {
    Check ($html -match [regex]::Escape($msg)) "tutorial cita a mensagem '$msg'"
    Check ($coreSrc -match [regex]::Escape($msg)) "o core emite a mensagem '$msg'"
}

# loader 1.2.0 com core anterior: o aviso vem do loader (modRibbon.bas), não do core
$loaderSrc = (Get-ChildItem (Join-Path $RepoRoot 'src\loader') -Recurse -Include '*.bas', '*.vba' | ForEach-Object { Read-Utf8 $_.FullName }) -join "`n"
$updMsg = 'Atualize o ThinQQ para usar este recurso.'
Check ($html -match [regex]::Escape($updMsg)) "tutorial cita a mensagem '$updMsg'"
Check ($loaderSrc -match [regex]::Escape($updMsg)) "o loader emite a mensagem '$updMsg'"

# ---- PowerPoint (Parte A, 1.4.0): cada mensagem que o usuário vê (suplemento do PowerPoint, motor do core e atualizador do loader) precisa
# estar no código que a emite (na fonte certa: um texto movido de lugar ou reescrito derruba o teste) E dentro da seção do tutorial onde é
# explicada (#powerpoint para a atualização dos gráficos, #atualizar para o suplemento do PowerPoint): o texto da seção é o decodificado
# (sem tags nem entidades), com os espaços normalizados.
$pptSrc = (Get-ChildItem (Join-Path $RepoRoot 'src\ppt') -Recurse -Include '*.bas', '*.cls' | ForEach-Object { Read-Utf8 $_.FullName }) -join "`n"
$srcByTag = @{ core = $coreSrc; loader = $loaderSrc; ppt = $pptSrc }
function Get-SectionText([string]$id) {
    $n = $xml.SelectSingleNode("//*[local-name()='section'][@id='$id']")
    if ($null -eq $n) { return '' }
    return (($n.InnerText -replace '\s+', ' ').Trim())
}
$pptMessages = @(
    # suplemento do PowerPoint (modPptTqq.bas)
    @('ppt', 'powerpoint', 'Esta apresentação não tem gráficos do ThinQQ vinculados.'),
    @('ppt', 'powerpoint', 'Abra uma apresentação primeiro.'),
    @('ppt', 'powerpoint', ' gráfico(s) do ThinQQ vinculado(s) ao Excel. Atualizar agora com os dados atuais?'),
    @('ppt', 'powerpoint', 'Para não perguntar mais: aba ThinQQ > Ao abrir.'),
    @('ppt', 'powerpoint', 'O Excel não está instalado (ou não pôde ser iniciado): os gráficos do ThinQQ não foram atualizados.'),
    @('ppt', 'powerpoint', 'O ThinQQ não está instalado no Excel deste computador: só os dados foram atualizados; setas e outras anotações do ThinQQ ficaram como estavam.'),
    @('ppt', 'powerpoint', 'Atualize o ThinQQ no Excel para atualizar também as setas.'),
    @('ppt', 'powerpoint', 'O Excel está ocupado (uma célula em edição ou uma caixa de diálogo aberta). Termine o que está fazendo no Excel e use ThinQQ > Atualizar gráficos.'),
    @('ppt', 'powerpoint', 'Dados atualizados em '),
    @('ppt', 'powerpoint', ' sem a pasta de origem (ficaram como estavam)'),
    # motor do core (modPptSync.bas, modCoreApi.bas): o resumo
    @('core', 'powerpoint', "Esta apresentação não tem gráficos do ThinQQ vinculados (copiados com 'Copiar como gráfico vinculado')."),
    @('core', 'powerpoint', "Abra a apresentação no PowerPoint e use de novo 'Atualizar gráficos no PowerPoint'."),
    @('core', 'powerpoint', 'Gráficos do ThinQQ atualizados: '),
    @('core', 'powerpoint', ' Sem mudança: '),
    @('core', 'powerpoint', 'Pasta de origem não encontrada ou sem acesso (o gráfico ficou como estava): '),
    @('core', 'powerpoint', 'Se ela foi movida ou renomeada, use no PowerPoint Arquivo > Informações > Editar Links para Arquivos > Alterar Origem.'),
    @('core', 'powerpoint', 'Gráfico não encontrado na pasta de origem (apagado ou recriado): '),
    @('core', 'powerpoint', 'Sem vínculo com o Excel (Quebrar Vínculo no PowerPoint): '),
    @('core', 'powerpoint', ' (veja o log do ThinQQ).'),
    @('core', 'powerpoint', 'Não foi possível atualizar os gráficos do ThinQQ: '),
    # atualizador do loader (modUpdater.bas: Updater_PpamStatusText; modRibbon.bas: a pergunta de ThinQQ > Configurações)
    @('loader', 'atualizar', 'Suplemento do PowerPoint atualizado.'),
    @('loader', 'atualizar', 'Suplemento do PowerPoint atualizado: feche e abra o PowerPoint para usar a versão nova.'),
    @('loader', 'atualizar', 'O suplemento do PowerPoint será atualizado automaticamente quando o PowerPoint for fechado (o Excel tenta de novo a cada 10 minutos e na próxima abertura).'),
    @('loader', 'atualizar', 'O suplemento do PowerPoint não foi atualizado agora; nova tentativa na próxima verificação.'),
    @('loader', 'atualizar', 'Atualizar também o suplemento do PowerPoint automaticamente?'),
    @('loader', 'atualizar', 'Atualizar também o suplemento do PowerPoint (feche o PowerPoint para aplicar na hora)')
)
foreach ($m in $pptMessages) {
    $tag = $m[0]; $sec = $m[1]; $msg = $m[2]
    $codeText = $srcByTag[$tag]
    if ($tag -eq 'loader') { $codeText = $codeText + "`n" + (Read-Utf8 (Join-Path $RepoRoot 'src\loader\forms\frmUpdate.json')) }
    Check ($codeText.Contains($msg)) "o código ($tag) emite a mensagem '$msg'"
    Check ((Get-SectionText $sec).Contains($msg)) "a seção #$sec do tutorial cita a mensagem '$msg'"
}
$pptMsgNoLoader = (Get-SectionText 'atualizar')
Check ($pptMsgNoLoader -cmatch 'suplemento do PowerPoint é atualizado automaticamente' -and $pptMsgNoLoader.Contains('Atualizar também o suplemento do PowerPoint')) '#atualizar explica a atualização automática do suplemento do PowerPoint e a opção que a desliga'

# ---- PowerPoint: instalação e publicação do ThinQQ.ppam documentadas nos lugares certos (3 arquivos, 6 assets, 9 arquivos soltos no ZIP)
$ppamNames = @((Read-Utf8 (Join-Path $RepoRoot 'tools\lib\Release-Publish.ps1')) -split "`n" | Where-Object { $_ -match 'return @\("ThinQQ-\$Version\.zip"' } | ForEach-Object { [regex]::Matches($_, "'([^']+)'") | ForEach-Object { $_.Groups[1].Value } })
Check ($ppamNames.Count -eq 5 -and $ppamNames -contains 'ThinQQ.ppam') "Get-ReleaseAssetNames lido do código: ZIP + 5 arquivos soltos ($($ppamNames -join ', '))"
$topFiles = @([regex]::Matches((Read-Utf8 (Join-Path $RepoRoot 'build\lib\Release-Lib.ps1')), "function Get-ReleaseTopFiles \{ return @\((?<l>[^)]*)\)"))
Check ($topFiles.Count -eq 1 -and ([regex]::Matches($topFiles[0].Groups['l'].Value, "'[^']+'")).Count -eq 5) 'Get-ReleaseTopFiles: 5 arquivos + os 4 do instalador = 9 soltos no ZIP'
$relReadmeText = ($relReadme -replace '\s+', ' ')
$readmeText = ($readme -replace '\s+', ' ')
foreach ($n in $ppamNames) {
    Check ($relReadmeText.Contains($n)) "release/README.md cita o arquivo da release '$n'"
}
Check ($relReadmeText -cmatch 'guarda os 6 arquivos' -and $relReadmeText -cnotmatch 'guarda os 5 arquivos') 'release/README.md: o fluxo automático guarda os 6 arquivos (ZIP + 5)'
Check ($relReadmeText -cmatch '\.xlam`? e o `?\.ppam`? não vêm assinados') 'release/README.md: .xlam e .ppam não vêm assinados'
Check ($readmeText -cmatch 'envia os 6 assets' -and $readmeText -cnotmatch 'envia os 5 assets') 'README: a etapa 3 envia os 6 assets'
foreach ($n in $ppamNames) { Check ($readmeText.Contains($n)) "README cita o arquivo da release '$n'" }
Check ((Read-Utf8 (Join-Path $RepoRoot 'build\lib\Release-Lib.ps1')) -match "function Get-ReleaseInstallerFiles \{ return @\('install\.ps1', 'uninstall\.ps1', 'Instalar-ThinQQ\.cmd', 'LEIA-ME\.txt'\) \}") 'Get-ReleaseInstallerFiles: 4 arquivos do instalador'
Check ($readmeText -cmatch 'Assinar os dois `\.xlam` e o `\.ppam`') 'README (TI): a assinatura cobre os dois .xlam e o .ppam'
Check ($readmeText -cmatch '9 arquivos soltos') 'README: o ZIP traz 9 arquivos soltos na raiz'
$signSrc = Read-Utf8 (Join-Path $RepoRoot 'tools\Sign-ThinQQ.ps1')
Check ($signSrc -match "(?s)\[string\[\]\]\`$Files = @\([^\r\n]*ThinQQ\.ppam") 'Sign-ThinQQ.ps1: o -Files padrão inclui dist\ThinQQ.ppam'
Check ($signSrc -match '(?s)<#.*\.SYNOPSIS[^\r\n]*\.ppam') 'Sign-ThinQQ.ps1: o .SYNOPSIS cita o .ppam'
$tutInstalar = Get-SectionText 'instalar'
Check ($tutInstalar.Contains('ThinQQ.xlam') -and $tutInstalar.Contains('ThinQQ.core.xlam') -and $tutInstalar.Contains('ThinQQ.ppam') -and $tutInstalar -cmatch 'três arquivos') '#instalar: a instalação tem três arquivos (ThinQQ.xlam, ThinQQ.core.xlam, ThinQQ.ppam)'
Check ($tutInstalar -cmatch 'Feche o Excel e o PowerPoint|feche o Excel e o PowerPoint') '#instalar: pede para fechar o Excel e o PowerPoint'
Check ($tutInstalar.Contains('PowerPoint\AddIns\ThinQQ') -and $tutInstalar.Contains('PowerPoint\Security\Trusted Locations')) '#instalar: "O que o instalador altera" cita as chaves do PowerPoint'
Check ($readmeText.Contains('PowerPoint\AddIns\ThinQQ') -and $readmeText.Contains('PowerPoint\Security\Trusted Locations\LocationThinQQ')) 'README: "O que o instalador altera" cita as chaves do PowerPoint (AddIns\ThinQQ e LocationThinQQ)'
$leiameText = Read-Utf8 $leiamePath
$leiameBlocks = @{}
foreach ($b in ($leiameText -split '(?m)^(?=[A-ZÀ-Ú]{4,}[A-ZÀ-Ú ]*( \(.*\))?\r?$)')) { $t = ($b -split "\r?\n")[0].Trim(); if ($t.Length -gt 0) { $leiameBlocks[$t] = $b } }
$lmManual = ''; foreach ($k in $leiameBlocks.Keys) { if ($k -like 'INSTALAÇÃO MANUAL*') { $lmManual = $leiameBlocks[$k] } }
$lmPpt = ''; foreach ($k in $leiameBlocks.Keys) { if ($k -like 'POWERPOINT*') { $lmPpt = $leiameBlocks[$k] } }
$lmReq = ''; foreach ($k in $leiameBlocks.Keys) { if ($k -like 'REQUISITOS*') { $lmReq = $leiameBlocks[$k] } }
Check ($lmManual -match 'ThinQQ\.ppam' -and $lmManual -match 'Subpastas também são confiáveis' -and $lmManual -match '(?s)faça o mesmo no PowerPoint \(Arquivo > Opções > Central de Confiabilidade >\s+Configurações > Locais Confiáveis > Adicionar novo local: %LOCALAPPDATA%\\ThinQQ, com as subpastas\)') 'LEIA-ME: a INSTALAÇÃO MANUAL copia o ThinQQ.ppam, desbloqueia e cria o Local Confiável do PowerPoint (com subpastas), na linha do PowerPoint'
# polimento da revisao da A8: assinatura cita o .ppam; afirmacoes sem exagero; notas com 1 a 8 linhas
Check ($leiameText -match '(?s)A TI pode assinar os dois \.xlam\s+e o ThinQQ\.ppam' -and $leiameText -match 'os \.xlam e o \.ppam assinados pela TI') 'LEIA-ME: a TI assina os dois .xlam e o .ppam (SEGURANÇA e Central de Confiabilidade)'
Check ($html.Contains('a TI assina os dois <code>.xlam</code> e o <code>.ppam</code>')) 'tutorial: a TI assina os dois .xlam e o .ppam'
Check ($signSrc -match '\.xlam no Excel e o \.ppam no PowerPoint') 'Sign-ThinQQ.ps1 (.NOTES): o .ppam se confere abrindo no PowerPoint'
Check (-not $html.Contains('saem sempre na mesma versão') -and -not $html.Contains('O núcleo foi atualizado normalmente')) 'tutorial: núcleo e suplemento normalmente na mesma versão; a falha do suplemento não afeta o núcleo'
Check ($html -match 'só é gravada ao clicar <span class="ui">Atualizar agora</span>') 'tutorial: a caixa do suplemento só é gravada ao clicar Atualizar agora'
Check ($readme -notmatch 'de 6 a 8' -and (Read-Utf8 (Join-Path $RepoRoot 'tools\lib\Release-Publish.ps1')) -notmatch 'de 6 a 8') 'README e publicar-release: notas de 1 a 8 linhas'
Check ($readme -match 'run-tests\.ps1 -Ppt' -and $readme -match 'tests/ps/Test-PptAddin\.ps1') 'README (comandos): run-tests.ps1 -Ppt e tests/ps/Test-PptAddin.ps1'
Check ((Read-Utf8 (Join-Path $RepoRoot 'release\notes.txt')) -match 'quem instalou o suplemento do PowerPoint pelo ZIP') 'notes.txt: a atualização do suplemento vale para quem o instalou pelo ZIP'
# o intervalo da nova tentativa do suplemento (PPAM_RETRY_MINUTES) e o mesmo no aviso do loader, no tutorial e no README
$updSrc = Read-Utf8 (Join-Path $RepoRoot 'src\loader\modules\modUpdater.bas')
$retryMin = [regex]::Match($updSrc, 'Private Const PPAM_RETRY_MINUTES As Long = (\d+)').Groups[1].Value
Check ($retryMin -ne '' -and $updSrc.Contains("a cada $retryMin minutos") -and (Get-SectionText 'atualizar').Contains("a cada $retryMin minutos") -and $readme.Contains("em até $retryMin minutos")) "PPAM_RETRY_MINUTES ($retryMin) citado no aviso, no tutorial (#atualizar) e no README"
Check (@(($lmManual -split "\r?\n") | Where-Object { $_ -match 'Desbloquear' -and $_ -match 'ThinQQ\.ppam' }).Count -ge 1 -or $lmManual -match '(?s)Em cada arquivo[^\r\n]*\.ppam') 'LEIA-ME: a INSTALAÇÃO MANUAL manda desbloquear também o .ppam'
Check ($lmPpt -match 'macros dessa pasta rodam sem aviso' -and $lmPpt -match '16\.0') 'LEIA-ME: o item POWERPOINT diz que as macros da pasta rodam sem aviso e cita o PowerPoint 16.0'
Check ($lmReq -match 'PowerPoint' -and $lmReq -match '16\.0') 'LEIA-ME: REQUISITOS cita o PowerPoint 16.0'
Check ($leiameText -match '(?s)A aba ThinQQ não aparece no PowerPoint:.{0,200}16\.0') 'LEIA-ME: PROBLEMAS COMUNS: "A aba ThinQQ não aparece no PowerPoint" cita o PowerPoint do Office 16.0'

# ---- README: a lista numerada "Verificações manuais" é contínua (1..N, sem buraco nem repetição) e traz os itens do PowerPoint (Parte A)
$mvBlock = [regex]::Match($readme, '(?s)## Verificações manuais no Excel real(?<b>.*?)## Limitações conhecidas').Groups['b'].Value
$mvNums = @([regex]::Matches($mvBlock, '(?m)^(\d+)\. ') | ForEach-Object { [int]$_.Groups[1].Value })
$mvSeq = ($mvNums.Count -gt 0) -and (@(for ($i = 0; $i -lt $mvNums.Count; $i++) { if ($mvNums[$i] -ne $i + 1) { $i } }).Count -eq 0)
Check $mvSeq "README: a lista de verificações manuais é numerada de 1 a $($mvNums.Count) sem buracos"
# os itens novos (a partir do 43, depois do último de antes da Parte A) cobrem o roteiro manual da especificação (§4.7): sem o corte pelo
# número, os itens antigos do PowerPoint (32 a 36, que também citam "PowerPoint fechado") fariam o teste passar sem os itens novos
$mvNew = (($mvBlock -split "\r?\n") | Where-Object { $_ -match '^(\d+)\. ' -and [int]$Matches[1] -ge 43 }) -join "`n"
Check ($mvNums.Count -ge 54) "README: a lista de verificações manuais tem os itens novos do PowerPoint (a partir do 43; total $($mvNums.Count))"
foreach ($t in 'Perguntar.*Atualizar sempre.*Não atualizar', 'Habilitar Edição', 'pasta movida|Alterar Origem', 'pasta em rede', '30 gráficos', 'Excel aberto', 'sem o ThinQQ no Excel', 'aba ThinQQ no PowerPoint', 'PowerPoint fechado', 'somente leitura', 'Desfazer', 'protegida por senha', 'não verificado em PowerPoint real', 'Assinatura do suplemento do PowerPoint[^\r\n]*ThinQQ\.ppam', 'Instalação manual do suplemento do PowerPoint[^\r\n]*Adicionar novo') {
    Check ($mvNew -match $t) "README (verificações manuais 43 em diante): cobre '$t'"
}

# ---- o limite do Arredondar intervalo vem do código (ROUND_MAX_CELLS) e a documentação cita o mesmo número
$rndSrc = Read-Utf8 (Join-Path $RepoRoot 'src\core\modules\modRounding.bas')
$rndMax = [int][regex]::Match($rndSrc, 'Public Const ROUND_MAX_CELLS As Long = (\d+)').Groups[1].Value
$rndText = ('{0:N0}' -f $rndMax -replace '[\s\u00A0,]', '.') + ' células'
Check ($rndMax -ge 100) "ROUND_MAX_CELLS lido do código ($rndMax)"
Check ($html -match [regex]::Escape($rndText) -and $readme -match [regex]::Escape($rndText)) "tutorial e README citam o limite do Arredondar intervalo ($rndText)"

# ---- limites e avisos que as revisões da Fase 3 deixaram para a documentação: cada item tem uma frase-chave no tutorial
# (tutorialOnly) ou no tutorial e no README (both). Se uma frase for reescrita, atualize-a aqui junto.
$mustDocument = [ordered]@{
    'borboleta: nome longo pode encostar nas barras (Mesma escala)'  = @('both', 'encostar nas barras')
    'borboleta: rótulos de barras finas ficam fora ou ocultos'       = @('tutorial', 'barras muito finas')
    'borboleta: Mesma escala só entre borboletas'                    = @('both', 'só entre borboletas')
    'dispersão: 300 linhas x 4 grupos levam cerca de 28 s'           = @('both', '28 s')
    'dispersão: linha depois da última não é detectada'              = @('both', 'depois da última linha da tabela')
    'dispersão: refresh regrava posição e cor do rótulo de bolha'    = @('tutorial', 'rótulo de bolha que você mudou')
    'dispersão: grupo novo exige Alterar dados'                      = @('tutorial', 'grupo novo')
    'dispersão: colunas extras depois do grupo são ignoradas'        = @('tutorial', 'depois da coluna de grupo')
    'dispersão: eixos seguem o formato da origem e se apertam'       = @('tutorial', 'em volta dos dados')
    'arredondamento: texto na origem dá #VALOR!'                     = @('tutorial', '#VALOR!')
    'arredondamento: a sobrescrita não pode ser desfeita'            = @('tutorial', 'não pode ser desfeita')
    'arredondamento: origem em % mantém %'                           = @('tutorial', 'continua em percentual')
    'arredondamento: planilha protegida precisa de Formatar células' = @('tutorial', 'Formatar células')
    'arredondamento: série só reordenada à mão mapeia por posição'   = @('tutorial', 'reordenou as séries')
    'arredondamento: série apagada à mão volta aos rótulos simples'  = @('tutorial', 'apagar uma série')
    'corte: só colunas e barras empilhadas e agrupadas'              = @('both', 'colunas e barras empilhadas e agrupadas')
    'corte: a linha de valor (média) acima do corte some'            = @('both', 'fica acima do corte')
    'corte: piso do segmento de cima em empilhadas'                  = @('tutorial', 'quarto')
    'corte: Mesma escala e corte são excludentes'                    = @('both', 'excludentes')
    'Mesma escala: rosca, dispersão, bolhas e gráfico com corte não entram' = @('both', 'rosca, dispersão, bolhas')
    'PowerPoint: colar sobrescreve a área de transferência'          = @('both', 'área de transferência')
    'PowerPoint: pasta aberta de URL (SharePoint/OneDrive)'          = @('both', 'SharePoint')
    'PowerPoint: Editar Links para Arquivos e Quebrar Vínculo'          = @('both', 'Editar Links para Arquivos')
    'PowerPoint: Quebrar Vínculo e Atualizar Dados'                  = @('tutorial', 'Quebrar Vínculo')
    'PowerPoint: destino é a apresentação da janela ativa'           = @('tutorial', 'janela ativa')
    'combinação: a linha sobe acima dos totais (e o máximo oculto)'  = @('tutorial', 'faixa acima dos totais')
    'formato inicial dos rótulos vem da primeira célula de valor'    = @('tutorial', 'primeira célula de valor')
    'pizza e rosca: 10 cores sem repetição'                          = @('tutorial', '10 fatias')
    'área: o rótulo da ponta que encosta no vizinho fica oculto'     = @('tutorial', 'rótulo da ponta')
    'texto das setas não se recalcula sem o ThinQQ'                  = @('both', 'não se recalcula')
    'copiar a planilha: avisa uma vez e usa Alterar dados na cópia'  = @('both', 'Alterar dados')
    'assinatura: o ramo signtool + SIPs não rodou ponta a ponta'     = @('both', 'ponta a ponta')
    'assinatura: roteiro manual pelo VBE (Ferramentas > Assinatura Digital)' = @('both', 'Assinatura Digital')
    'comparativa: percentual (B - A) / A'                            = @('both', '(B - A) / A')
    'comparativa: n/d com zero ou sinais opostos'                    = @('both', 'sinais opostos')
    'comparativa: conectores atravessam a barra da outra série'      = @('tutorial', 'atravessam a barra da outra série')
    'comparativa: afastamento das setas só com o ThinQQ'             = @('both', 'afastamento das setas')
    'comparativa: sem cálculo de trás para frente'                   = @('tutorial', 'não calcula para trás')
    'comparativa: no máximo 2 séries'                                = @('both', 'mais de 2 séries')
    'PowerPoint: abre as pastas da apresentação sem macros; Atualizar sempre só com fontes confiáveis' = @('both', 'abre, somente leitura e com as macros desligadas, as pastas de trabalho para as quais a apresentação aponta')
    'PowerPoint: Atualizar sempre pergunta com origem de rede ou da web não aberta' = @('both', 'pergunta assim mesmo')
    'PowerPoint: a pergunta lista as pastas de origem'               = @('tutorial', 'Pastas de origem:')
    'PowerPoint: letras de unidade mapeadas não são detectadas como rede' = @('both', 'letras de unidade mapeadas')
    'comparativa: elipses das setas se sobrepõem em gráfico estreito' = @('both', 'elipses das setas podem se sobrepor')
    'comparativa: sem recalibração com menos de 92 pt'               = @('both', '92 pt')
    'comparativa: Mesma escala pode cortar as setas do topo'         = @('both', 'cortar as setas do topo')
    'comparativa: total negativo mostra o percentual com sinal'      = @('tutorial', '(-50%)')
    'comparativa: o percentual dos totais volta ao alargar'          = @('tutorial', 'volta sozinho')
    'comparativa: o percentual dos totais some sozinho se os rótulos se tocam' = @('both', 'some sozinho quando os dois rótulos de total')
    'comparativa: rótulos de total vizinhos podem se sobrepor'       = @('both', 'ou alargue o gráfico')
    # PowerPoint atualizado ao reabrir (Parte A, 1.4.0)
    'powerpoint: sem o suplemento so os dados atualizam'             = @('both', 'sem o suplemento do PowerPoint')
    'powerpoint: grafico dentro de grupo fica de fora'               = @('tutorial', 'dentro de um grupo')
    'powerpoint: Modo de Exibicao Protegido'                         = @('both', 'Habilitar Edição')
    'powerpoint: pasta movida'                                       = @('both', 'Alterar Origem')
    'powerpoint: apresentacao somente leitura so atualiza na memoria' = @('both', 'somente na memória')
    'powerpoint: a atualizacao pergunta se quer salvar ao fechar'    = @('both', 'se quer salvar ao fechar')
    'powerpoint: o Desfazer do Excel e limpo'                        = @('both', 'Desfazer do Excel')
    'powerpoint: atualizacao nao e transacional'                     = @('both', 'parcialmente atualizado')
    'powerpoint: pasta protegida por senha vai para sem acesso'      = @('both', 'protegida por senha')
    'powerpoint: Salvar como pode perguntar mais uma vez'            = @('tutorial', 'Salvar como')
    'powerpoint: Excel travado nao tem tempo limite'                 = @('both', 'sem tempo limite')
    'powerpoint: formatacao feita no PowerPoint fica; legenda desalinha as setas' = @('tutorial', 'desalinhar as setas')
    'powerpoint: pasta do OneDrive/SharePoint so com a pasta aberta no Excel' = @('tutorial', 'aberta no Excel')
    'powerpoint: o instalador cria o local confiavel do PowerPoint'  = @('both', 'PowerPoint\Security\Trusted Locations\LocationThinQQ')
}
foreach ($k in $mustDocument.Keys) {
    $scope = $mustDocument[$k][0]; $phrase = $mustDocument[$k][1]
    Check ($html -cmatch [regex]::Escape($phrase)) "tutorial documenta: $k ('$phrase')"
    if ($scope -eq 'both') { Check ($readme -cmatch [regex]::Escape($phrase)) "README documenta: $k ('$phrase')" }
}
if ($script:fail -gt 0) { Write-Host "FALHAS Test-Docs: $script:fail"; exit 1 }
Write-Host 'Test-Docs: OK'
