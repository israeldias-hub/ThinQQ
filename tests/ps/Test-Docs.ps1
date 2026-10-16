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
    $ctlNodes = @($ribbon.SelectNodes("//*[local-name()='button' or local-name()='toggleButton' or local-name()='checkBox' or local-name()='dropDown' or local-name()='menu' or local-name()='group' or local-name()='tab']"))
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
    $pptCtl = @($pptRibbon.SelectNodes("//*[local-name()='tab' or local-name()='group' or local-name()='button' or local-name()='toggleButton' or local-name()='dropDown' or local-name()='item']"))
    Check ($pptCtl.Count -eq 34) "faixa do PowerPoint: 34 controles (aba, 3 grupos, 25 botões (21 tipos de Inserir gráfico), 1 alternância, lista e 3 itens: $($pptCtl.Count))"
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
    Check ($matrix.Count -eq 24) "24 recursos na matriz de capacidades de modChartKinds.bas ($($matrix.Count))"
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
# Parte B (delta, 1.5.0) depois da 1.4.0 publicada: o publicar-release transforma [Não lançado] em [1.4.0] e um merge de 3 vias deste ramo
# poria as linhas 'Delta (1.5.0)' SEM conflito dentro da [1.4.0]. Guarda (inerte enquanto nao houver [1.4.0]): com a [1.4.0] no arquivo, as
# linhas do delta ficam em [Não lançado] e nenhuma na [1.4.0]. Devolve 'inerte', 'ok' ou o motivo da falha.
function Test-DeltaChangelogSection([string]$text) {
    $t = $text -replace "`r`n", "`n"
    if ($t -notmatch '(?m)^## \[1\.4\.0\]') { return 'inerte' }
    if ($t -match '(?m)^## \[1\.5\.0\]') { return 'inerte' }   # 1.5.0 ja publicada: as linhas Delta (1.5.0) pertencem a [1.5.0], nao ha mais o que guardar
    $un = [regex]::Match($t, '(?ms)^## \[Não lançado\][^\n]*\n(?<b>.*?)(?=^## \[|\z)').Groups['b'].Value
    $v14 = [regex]::Match($t, '(?ms)^## \[1\.4\.0\][^\n]*\n(?<b>.*?)(?=^## \[|\z)').Groups['b'].Value
    if (-not $un.Contains('Delta (1.5.0)')) { return 'as linhas Delta (1.5.0) não estão em [Não lançado]' }
    if ($v14.Contains('Delta (1.5.0)')) { return 'há linhas Delta (1.5.0) dentro da [1.4.0]' }
    return 'ok'
}
$clBad = "# Changelog`n`n## [Não lançado]`n`n## [1.4.0] - 2026-10-20`n`n- PowerPoint atualizado ao reabrir`n- Delta (1.5.0): Delta entre barras`n`n## [1.3.0] - 2026-10-14`n`n- x`n"
$clGood = "# Changelog`n`n## [Não lançado]`n`n- Delta (1.5.0): Delta entre barras`n`n## [1.4.0] - 2026-10-20`n`n- PowerPoint atualizado ao reabrir`n`n## [1.3.0] - 2026-10-14`n`n- x`n"
$clEmpty = "# Changelog`r`n`r`n## [Não lançado]`r`n`r`n## [1.4.0] - 2026-10-20`r`n`r`n- PowerPoint`r`n"
Check ((Test-DeltaChangelogSection $clBad) -ne 'ok' -and (Test-DeltaChangelogSection $clBad) -ne 'inerte') 'guarda do CHANGELOG (sintético): Delta (1.5.0) caído dentro da [1.4.0] depois do merge = falha'
Check ((Test-DeltaChangelogSection $clEmpty) -ne 'ok' -and (Test-DeltaChangelogSection $clEmpty) -ne 'inerte') 'guarda do CHANGELOG (sintético, CRLF): [1.4.0] publicada e [Não lançado] sem o delta = falha'
Check ((Test-DeltaChangelogSection $clGood) -eq 'ok') 'guarda do CHANGELOG (sintético): Delta (1.5.0) em [Não lançado] acima da [1.4.0] = ok'
Check ((Test-DeltaChangelogSection "## [Não lançado]`n`n- Delta (1.5.0): x`n`n## [1.3.0] - 2026-10-14`n") -eq 'inerte') 'guarda do CHANGELOG (sintético): sem [1.4.0] = inerte'
Check ($readme.Contains('**Publicando a Parte B depois da 1.4.0.**') -and $readme -match '(?s)Publicando a Parte B depois da 1\.4\.0.{0,900}mova essas linhas para a nova `\[Não lançado\]`') 'README (Publicar uma versão): como publicar a Parte B depois da 1.4.0 (mover as linhas Delta (1.5.0) para [Não lançado])'
$clGuard = Test-DeltaChangelogSection $cl
Check ($clGuard -eq 'ok' -or $clGuard -eq 'inerte') "CHANGELOG: linhas Delta (1.5.0) na seção certa depois da 1.4.0 ($clGuard)"

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
    ' colunas de valores. Selecione a coluna das etapas e só as duas colunas a comparar.',
    # delta entre barras (Parte B, 1.5.0): avisos (ANN_DELTA_REFUSAL, DELTA_BAD_PICK, DELTA_NONE, DELTAPICK_MSG_PROTECTED/ARMFAIL e o aviso sem gráfico)
    # e as mensagens da barra de status do modo de dois cliques (DELTAPICK_MSG_*)
    'O delta entre barras não está disponível para este tipo de gráfico.',
    'Escolha duas barras diferentes do gráfico.',
    'Este gráfico não tem delta entre barras. Use Setas > Delta clicando nas barras.',
    'A planilha está protegida. Desproteja a planilha para inserir o delta.',
    'Selecione um gráfico do ThinQQ antes de inserir o delta.',
    'Selecione um gráfico do ThinQQ antes de alternar o texto do delta.',
    'Não foi possível ativar o delta clicando nas barras neste gráfico. Selecione o gráfico e tente de novo.',
    'ThinQQ - Delta: clique na barra de ORIGEM (Esc ou um clique numa célula cancela).',
    'ThinQQ - Delta: agora clique na barra de DESTINO.',
    'ThinQQ - Delta: escolha uma barra diferente da origem.',
    'ThinQQ - Delta: clique numa barra (não em rótulos, setas ou linhas auxiliares).',
    'ThinQQ - Delta: clique de novo na mesma barra para escolher só ela.',
    'ThinQQ - Delta cancelado.',
    'ThinQQ - Delta cancelado (outro gráfico).',
    'ThinQQ - Delta cancelado (60 s sem clique).'
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
    @('ppt', 'powerpoint', 'O Excel está ocupado (uma célula em edição ou uma caixa de diálogo aberta). Termine o que está fazendo no Excel e tente de novo.'),
    @('ppt', 'powerpoint', 'Selecione um ou mais gráficos do ThinQQ vinculados ao Excel.'),
    @('ppt', 'powerpoint', 'Atualize o ThinQQ no Excel para a atualização automática funcionar (a marca foi gravada).'),
    @('ppt', 'powerpoint', 'Marcado para atualizar automaticamente. Para os gráficos acompanharem a pasta de origem, abra o Excel com o ThinQQ (1.6.0 ou mais novo).'),
    @('ppt', 'powerpoint', 'A atualização automática está desligada no Excel: aba ThinQQ > PowerPoint > Atualização automática ligada.'),
    @('ppt', 'powerpoint', 'Atualize o ThinQQ no Excel para atualizar só os gráficos selecionados (desta vez a apresentação inteira foi atualizada).'),
    @('ppt', 'powerpoint', 'Não foi possível marcar o gráfico como automático: '),
    @('ppt', 'powerpoint', 'Não foi possível voltar o gráfico para a atualização manual: '),
    @('ppt', 'powerpoint', 'Dados atualizados em '),
    @('ppt', 'powerpoint', ' sem a pasta de origem (ficaram como estavam)'),
    # motor do core (modPptSync.bas, modCoreApi.bas): o resumo
    @('core', 'powerpoint', "Esta apresentação não tem gráficos do ThinQQ vinculados (copiados com 'Copiar como gráfico vinculado')."),
    @('core', 'powerpoint', "Abra a apresentação no PowerPoint e use de novo 'Atualizar gráficos no PowerPoint'."),
    @('core', 'powerpoint', 'Os gráficos selecionados não foram encontrados na apresentação (renomeados?): selecione-os de novo.'),
    @('core', 'powerpoint', 'Gráficos do ThinQQ atualizados: '),
    # Inserir também no slide (1.7.0): modCoreApi.bas e o aviso de núcleo antigo da faixa
    @('core', 'powerpoint', 'Salve a pasta de trabalho antes de inserir no PowerPoint: o gráfico vinculado aponta para o arquivo salvo. Nada foi criado.'),
    @('core', 'powerpoint', "Esta pasta não tem molduras de intervalos vinculados ao PowerPoint (elas aparecem ao usar 'Inserir também no slide')."),
    @('loader', 'powerpoint', 'Atualize o ThinQQ para usar este recurso.'),
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
# notas da 1.6.0 (Parte C, atualização automática): o texto fica em $notesText; as frases-chave e a contagem de linhas são conferidas no bloco
# "Atualização automática no PowerPoint" mais abaixo; sem o aviso de reinstalar pelo ZIP (quem tem a 1.5.0 atualiza pelo Excel)
$notesText = Read-Utf8 (Join-Path $RepoRoot 'release\notes.txt')
Check ($notesLines.Count -ge 1 -and $notesLines.Count -le 8 -and $notesText -notmatch 'reinstal') 'notes.txt (1.6.0): de 1 a 8 linhas, sem o aviso de reinstalar pelo ZIP'
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
Check ($mvNums.Count -ge 68) "README: a lista de verificações manuais tem os itens novos do PowerPoint, do delta e da atualização automática (a partir do 43; total $($mvNums.Count))"
foreach ($t in 'Perguntar.*Atualizar sempre.*Não atualizar', 'Habilitar Edição', 'pasta movida|Alterar Origem', 'pasta em rede', '30 gráficos', 'Excel aberto', 'sem o ThinQQ no Excel', 'aba ThinQQ no PowerPoint', 'PowerPoint fechado', 'somente leitura', 'Desfazer', 'protegida por senha', 'não verificado em PowerPoint real', 'Assinatura do suplemento do PowerPoint[^\r\n]*ThinQQ\.ppam', 'Instalação manual do suplemento do PowerPoint[^\r\n]*Adicionar novo') {
    Check ($mvNew -match $t) "README (verificações manuais 43 em diante): cobre '$t'"
}

# delta entre barras (1.5.0): os itens 55 em diante cobrem a lista manual da B3 (cliques reais nos 9 tipos, 1.o clique num gráfico não ativado,
# Esc com o gráfico ativo, zoom/DPI, cancelamentos reais, casos de borda, PowerPoint) e cada um se declara não verificado com mouse real
$mvDelta = (($mvBlock -split "\r?\n") | Where-Object { $_ -match '^(\d+)\. ' -and [int]$Matches[1] -ge 55 -and [int]$Matches[1] -le 60 }) -join "`n"
Check (@($mvNums | Where-Object { $_ -ge 55 -and $_ -le 60 }).Count -eq 6) "README: itens 55 a 60 (delta): $(@($mvNums | Where-Object { $_ -ge 55 -and $_ -le 60 }).Count)"
foreach ($t in 'cliques reais com o mouse nos 9 tipos', 'primeiro clique num gráfico que não estava ativado', 'clique único em barra de série já selecionada', 'zoom 75% e 150% e escala de tela 125% e 150%', 'Esc com o gráfico ativo', '60 s', 'outro gráfico', 'grupo de formas', 'planilha protegida', 'Delta no PowerPoint', 'Ctrl\+Z') {
    Check ($mvDelta -match $t) "README (verificações manuais 55 em diante): cobre '$t'"
}
Check (@(($mvDelta -split "`n") | Where-Object { $_ -notmatch 'não verificado com mouse real' }).Count -eq 0) 'README (verificações manuais 55 em diante): todo item do delta marca "não verificado com mouse real"'

# delta: a especificação (as duas copias que o repositorio guarda: so a de docs/superpowers e rastreada) cita o manipulador certo do OnTime e a regra do lado de baixo
$specDeltaPath = Join-Path $RepoRoot 'docs\superpowers\specs\2026-10-14-thinqq-ppt-delta-design.md'
$specDelta = Read-Utf8 $specDeltaPath
$bt = [string][char]96
Check ($specDelta.Contains('DeltaPick_TimeoutFired') -and -not $specDelta.Contains($bt + 'DeltaPick_Timeout' + $bt + ', ' + $bt + 'DeltaPick_IsArmedFor')) 'especificação do delta: o manipulador do OnTime é DeltaPick_TimeoutFired'
Check ($specDelta.Contains('Emenda (revisão da B4), lado de baixo') -and $specDelta.Contains('Delta_IsStackBottomNeg') -and $specDelta.Contains('NÃO deslocam')) 'especificação do delta (5.3): regra do lado de baixo'
Check ($specDelta.Contains('Emenda (revisão da B3)') -and $specDelta.Contains('DELTAPICK_MSG_ARMFAIL')) 'especificação do delta (5.1): cancelamentos e mensagens da B3'
# B6: o prazo do modo de dois cliques e renovado no clique de origem aceito: tutorial, README (limitações e item manual) e especificação dizem
# "60 s sem clique, contados do último clique aceito" e nenhum volta a dizer que o prazo conta só de quando o modo é armado
$prazo = '60 s sem clique, contados do último clique aceito'
Check ($html.Contains($prazo)) "tutorial: '$prazo'"
Check (([regex]::Matches($readme, [regex]::Escape($prazo))).Count -ge 2) "README (limitações e item manual 58): '$prazo'"
Check ($specDelta.Contains($prazo)) "especificação do delta: '$prazo'"
$undoDelta = 'Inserir ou alternar o texto do delta limpa o Desfazer do Excel (como as demais setas)'
Check ($html.Contains($undoDelta) -and $readme.Contains($undoDelta)) "tutorial (#delta, Editar, atualizar e apagar) e README (Limitações): '$undoDelta'"
Check ($mvDelta.Contains('ao fechar o Excel pode perguntar se quer salvar (o marcador de origem é removido ao fechar)')) 'README item manual 58: Ctrl+S com o modo armado - o Excel pode perguntar se quer salvar ao fechar'
foreach ($velho in 'os cliques não o renovam', 'um clique não o renova', 'conta a partir de quando o modo é armado', 'conta de quando o modo é armado', '60 s depois de armado') {
    Check (-not $html.Contains($velho) -and -not $readme.Contains($velho) -and -not $specDelta.Contains($velho)) "prazo do delta: nenhum documento diz '$velho'"
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
# ---- delta entre barras (1.5.0): cada limite e conferido POR LUGAR (secao #setas e #limites do tutorial, "Limitações conhecidas" do README):
# tirar UMA das ocorrencias derruba o teste (uma frase presente so em outro lugar nao a cobre)
$setasTxt = Get-SectionText 'setas'
$limitesTxt = Get-SectionText 'limites'
$readmeLimTxt = ([regex]::Match($readme, '(?s)## Limitações conhecidas(?<b>.*?)## Licença').Groups['b'].Value -replace '\s+', ' ')
Check ($setasTxt.Contains('Delta entre barras') -and $limitesTxt.Length -gt 500 -and $readmeLimTxt.Length -gt 500) 'delta: secoes #setas e #limites e a lista de limitações do README lidas'
$deltaPins = @(
    @('seta nao arrastavel',                          'não é arrastável',                           @('setas', 'limites', 'readme')),
    @('grafico em grupo de formas',                   'grupo de formas',                            @('setas', 'limites', 'readme')),
    @('percentual diferente da seta de variacao',     '-75%',                                       @('setas', 'limites', 'readme')),
    @('Alterar dados nao preserva os deltas',         'não preserva os deltas',                     @('setas', 'limites', 'readme')),
    @('a sala so encolhe a area de plotagem',         'só encolhe a área de plotagem',              @('setas', 'limites', 'readme')),
    @('redimensionar a mao nao reabre a sala',        'não reabre o espaço',                        @('setas', 'limites', 'readme')),
    @('formato percentual sem p.p.',                  'p.p.',                                       @('setas', 'readme')),
    @('a sala so e refeita ate o proximo comando',    'até o próximo desses comandos',              @('setas', 'readme')),
    @('larguras de rotulo estimam fonte 9 pt negrito', 'fonte de 9 pt em negrito',                  @('setas', 'readme')),
    @('50 ou mais categorias em barras horizontais',  '50 ou mais categorias',                      @('setas', 'readme')),
    @('muitos deltas e o piso de 30%',                'abaixo de 30% do gráfico',                   @('setas', 'readme')),
    @('comparativa: perna compartilha o x da seta de variacao', 'mesmo x da perna da seta de variação', @('setas', 'readme')),
    @('comparativa: so com as setas de variacao ligadas', 'com as setas de variação ligadas',       @('setas', 'readme')),
    @('empilhadas: perna de baixo cruza o rotulo de cima', 'cruza o rótulo do segmento de cima',    @('setas', 'readme')),
    @('Esc: atribuicao de outro suplemento se perde', 'outro suplemento ou macro se perde',         @('setas', 'readme')),
    @('cliques reais nao verificados com mouse real', 'não verificado com mouse real',              @('setas', 'readme')),
    @('cancela ao trocar de planilha ou de pasta',    'planilha ou de pasta de trabalho',           @('setas')),
    @('planilha protegida so com o grafico desbloqueado', 'nem pode ser selecionado',               @('setas')),
    @('clicar de novo na origem so avisa na barra de status', 'Clicar de novo na própria origem',   @('setas'))
)
$pinText = @{ setas = $setasTxt; limites = $limitesTxt; readme = $readmeLimTxt }
$pinName = @{ setas = 'tutorial #setas (Delta entre barras)'; limites = 'tutorial #limites'; readme = 'README Limitações conhecidas' }
foreach ($d in $deltaPins) {
    foreach ($w in $d[2]) { Check ($pinText[$w].Contains($d[1])) "delta: $($d[0]) ($($pinName[$w]): '$($d[1])')" }
}
# a sala do delta NAO e refeita ao redimensionar a mao (Cht_Resize so reposiciona): nenhum texto pode listar o redimensionamento/tamanho como
# gatilho da sala; a frase positiva com os gatilhos reais (inserir, alternar o texto, legenda, formato, decoracoes, corte de eixo) fica no tutorial, no README e na especificacao
$roomWrong = '(?i)formato e (tamanho|redimensionamento)|legenda, formato, redimensionar|<span class="ui">Formato</span> e redimensionamento|próximo redimensionamento'
Check ($html -notmatch $roomWrong -and $readme -notmatch $roomWrong -and $specDelta -notmatch $roomWrong) 'delta: nenhum texto lista o redimensionamento como gatilho da sala do delta'
Check ($specDelta.Contains('não reabre o espaço') -and $specDelta.Contains('Chart_EnsureRoomAfterChange') -and $specDelta.Contains('corte de eixo')) 'delta: a especificação lista os gatilhos reais da sala e diz que o redimensionamento a mão não reabre o espaço'
Check ($limitesTxt.Contains('legenda, formato, decorações e corte de eixo') -and $readmeLimTxt.Contains('legenda, formato, decorações') -and $setasTxt.Contains('Corte de eixo...')) 'delta: tutorial e README citam os gatilhos reais da sala (legenda, formato, decorações, corte de eixo)'

# ---- Atualização automática no PowerPoint (1.6.0, Parte C): cada afirmação é conferida contra o CÓDIGO e POR LUGAR (subseção #ppt-auto do
# tutorial, "Recursos" e "Limitações conhecidas" do README, itens manuais 61 a 68, [Não lançado] do CHANGELOG e notes.txt): tirar a frase de UM
# lugar derruba o teste, e um número que muda no código (espera, tentativas, tempo máximo) obriga a reescrever os textos.
$autoSrc = Read-Utf8 (Join-Path $RepoRoot 'src\core\modules\modPptAuto.bas')
$syncSrc = Read-Utf8 (Join-Path $RepoRoot 'src\core\modules\modPptSync.bas')
$evSrc = Read-Utf8 (Join-Path $RepoRoot 'src\core\classes\clsAppEvents.cls')
$pptTqqSrc = Read-Utf8 (Join-Path $RepoRoot 'src\ppt\modules\modPptTqq.bas')
$debounce = [int][regex]::Match($autoSrc, 'Public Const PPTAUTO_DEBOUNCE_S As Long = (\d+)').Groups[1].Value
$maxRetries = [int][regex]::Match($autoSrc, 'Public Const PPTAUTO_MAX_RETRIES As Long = (\d+)').Groups[1].Value
$maxWait = [int][regex]::Match($autoSrc, 'Public Const PPTAUTO_MAX_WAIT_S As Long = (\d+)').Groups[1].Value
$tagAuto = [regex]::Match($syncSrc, 'Public Const PPTSYNC_TAG_AUTO As String = "([^"]+)"').Groups[1].Value
$tagIdx = [regex]::Match($syncSrc, 'Public Const PPTSYNC_TAG_INDEX As String = "([^"]+)"').Groups[1].Value
Check ($debounce -ge 1 -and $maxRetries -ge 2 -and $maxWait -gt $debounce) "atualização automática: constantes lidas do código (espera $debounce s, $maxRetries tentativas, teto $maxWait s)"
Check ($tagAuto -ceq 'TQQ_AUTO' -and $tagIdx -ceq 'TQQ_AUTOIDX' -and $pptTqqSrc.Contains('PPT_TAG_AUTO As String = "TQQ_AUTO"') -and $pptTqqSrc.Contains('PPT_TAG_INDEX As String = "TQQ_AUTOIDX"')) 'atualização automática: as etiquetas TQQ_AUTO e TQQ_AUTOIDX são as mesmas no core e no suplemento'
# as novas tentativas seguem DEBOUNCE * (n + 1): 4, 6, 8, 10 e 12 s com os valores atuais
Check ($autoSrc.Contains('PptAuto_Schedule PPTAUTO_DEBOUNCE_S * (mRetries + 1)') -and $autoSrc.Contains('mRetries < PPTAUTO_MAX_RETRIES')) 'atualização automática: a nova tentativa espera DEBOUNCE * (n + 1) e há um teto de tentativas'
$retrySecs = @(); for ($i = 1; $i -le $maxRetries; $i++) { $retrySecs += ($debounce * ($i + 1)) }
$retryText = (($retrySecs[0..($retrySecs.Count - 2)]) -join ', ') + ' e ' + $retrySecs[$retrySecs.Count - 1] + ' s'
$retryTotal = ($retrySecs | Measure-Object -Sum).Sum
$waitText = "$maxWait a $($maxWait + 1) s"
$waitTextLong = "$maxWait a $($maxWait + 1) segundos"
$debText = "cerca de $debounce s"
$debTextLong = "cerca de $debounce segundos"
# padrões das duas opções, eventos, ciclo que nunca abre pasta, somente leitura e PowerPoint já aberto
Check ($autoSrc.Contains('GetSetting(APP_NAME, PPTAUTO_SECTION, "AutoUpdate", "1")') -and $autoSrc.Contains('GetSetting(APP_NAME, PPTAUTO_SECTION, "AutoNew", "0")')) 'atualização automática (código): Atualização automática ligada começa ligada e Novos vínculos começa desligada'
Check ($evSrc -match '(?s)App_SheetChange.{0,200}PptAuto_OnChange.{0,200}App_SheetCalculate.{0,200}PptAuto_OnChange.{0,260}App_WorkbookAfterSave.{0,200}PptAuto_OnChange') 'atualização automática (código): digitar, recalcular e salvar chamam PptAuto_OnChange'
Check ($autoSrc.Contains('PptSync_Items(items, False, True)') -and $autoSrc.Contains('PptAuto_KeepOpenBooks mDirty')) 'atualização automática (código): o ciclo nunca abre pasta (noOpen e a poda das pastas fechadas)'
Check ($autoSrc.Contains('If CLng(pres.ReadOnly) <> 0 Then Exit Sub') -and $autoSrc.Contains('Set app = Ppt_GetApp(False)')) 'atualização automática (código): o ciclo pula somente leitura e só usa um PowerPoint já aberto'
Check ($autoSrc.Contains('If Len(wb.Path) = 0 Then Exit Function') -and $autoSrc.Contains('Set ws = wb.Worksheets(CALC_SHEET)')) 'atualização automática (código): só pastas salvas e com gráficos do ThinQQ são candidatas'
Check ($autoSrc.Contains('PptAuto_Log "PptAuto: desistiu"') -and $autoSrc.Contains('If mInTick Or Ppt_InOperation() Then PptAuto_Schedule PPTAUTO_DEBOUNCE_S')) 'atualização automática (código): desiste com o registro "PptAuto: desistiu" e espera uma operação do core em andamento'
Check ($pptTqqSrc -match '(?s)If turnOn Then PptTqq_SyncEntries pres, PptTqq_EntriesOf\(shapes\), True, True') 'atualização automática (código): ligar o automático já atualiza uma vez só as formas do botão'
Check ($pptTqqSrc.Contains('PPT_API_SHAPES As Long = 2') -and $syncSrc -match 'CORE_PPT_API As Long = ([2-9]|\d{2,})\b') 'atualização automática (código): o contrato do suplemento com o core passa à API 2 (ou maior: 3 = 1.8.0)'

# rótulos exatos das duas faixas (a do Excel tem as caixas, a do PowerPoint o botão de alternância e o botão)
$lblOn = 'Atualização automática ligada'; $lblNew = 'Novos vínculos atualizam automaticamente'; $lblAuto = 'Atualizar automaticamente'; $lblOnce = 'Atualizar uma vez'
$xlOnNode = $ribbon.SelectSingleNode("//*[@id='chkPptAutoOn']"); $xlNewNode = $ribbon.SelectSingleNode("//*[@id='chkPptAutoNew']")
$pAutoNode = $pptRibbon.SelectSingleNode("//*[@id='tglPptAuto']"); $pOnceNode = $pptRibbon.SelectSingleNode("//*[@id='btnPptUpdateSel']")
Check ($null -ne $xlOnNode -and $xlOnNode.LocalName -eq 'checkBox' -and $xlOnNode.GetAttribute('label') -ceq $lblOn -and $xlOnNode.GetAttribute('tag') -ceq 'autoon') "faixa do Excel: a caixa chkPptAutoOn tem o rótulo '$lblOn'"
Check ($null -ne $xlNewNode -and $xlNewNode.LocalName -eq 'checkBox' -and $xlNewNode.GetAttribute('label') -ceq $lblNew -and $xlNewNode.GetAttribute('tag') -ceq 'autonew') "faixa do Excel: a caixa chkPptAutoNew tem o rótulo '$lblNew'"
Check ($null -ne $pAutoNode -and $pAutoNode.GetAttribute('label') -ceq $lblAuto -and $null -ne $pOnceNode -and $pOnceNode.GetAttribute('label') -ceq $lblOnce) "faixa do PowerPoint: tglPptAuto '$lblAuto' e btnPptUpdateSel '$lblOnce'"
$xlChk = @($ribbon.SelectNodes("//*[local-name()='checkBox']"))
Check ($xlChk.Count -ge 2) "faixa do Excel: as caixas de seleção entram na regra 'todo controle aparece no tutorial' ($($xlChk.Count))"
if ($parsed) {
    Check ($null -ne $xml.SelectSingleNode("//*[local-name()='section'][@id='powerpoint']//*[local-name()='h3'][@id='ppt-auto']")) 'tutorial: a subseção id="ppt-auto" está dentro de #powerpoint'
    Check ($xml.SelectNodes("//a[@href='#ppt-auto']").Count -ge 3) 'tutorial: a subseção #ppt-auto tem links de #powerpoint, da referência da faixa e dos limites'
    foreach ($l in $lblOn, $lblNew, $lblAuto, $lblOnce) { Check ($pptUi -contains $l) "tutorial: '$l' citado como span.ui dentro de #powerpoint" }
}
# o texto da subseção #ppt-auto (do h3 até o próximo h3) sem tags nem entidades
function Get-SubsectionText([string]$id) {
    $h = $xml.SelectSingleNode("//*[@id='$id']")
    if ($null -eq $h) { return '' }
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append($h.InnerText).Append(' ')
    $n = $h.NextSibling
    while ($null -ne $n -and -not ($n.NodeType -eq 'Element' -and $n.LocalName -eq 'h3')) { [void]$sb.Append($n.InnerText).Append(' '); $n = $n.NextSibling }
    return (($sb.ToString() -replace '\s+', ' ').Trim())
}
$autoTut = Get-SubsectionText 'ppt-auto'
$recRaw = [regex]::Match($readme, '(?s)\*\*Atualização automática no PowerPoint \(1\.6\.0\)\*\*(?<b>.*?)### Galeria').Groups['b'].Value
$autoRec = ($recRaw -replace '\s+', ' ')
$clUnrel = [regex]::Match(($cl -replace "`r`n", "`n"), '(?ms)^## \[Não lançado\][^\n]*\n(?<b>.*?)(?=^## \[|\z)').Groups['b'].Value
# As linhas da 1.6.0 ficam em [Não lançado] até o publicar-release, que as move para "## [1.6.0] - data" (1.6.0 publicada antes) ou, se a 1.6.0
# e a 1.7.0 saírem juntas, para "## [1.7.0] - data" (mesma classe do incidente do delta, dc446b0): os pinos leem a primeira seção que tem as
# linhas (TQQ_AUTOIDX), nesta ordem: [Não lançado], [1.7.0], [1.6.0].
function Get-ChangelogSection([string]$text, [string]$head) {
    $t = $text -replace "`r`n", "`n"
    return [regex]::Match($t, "(?ms)^## $head[^\n]*\n(?<b>.*?)(?=^## \[|\z)").Groups['b'].Value
}
function Get-AutoChangelogBlock([string]$text) {
    foreach ($h in '\[Não lançado\]', '\[1\.7\.0\]') {
        $b = Get-ChangelogSection $text $h
        if ($b.Contains('TQQ_AUTOIDX')) { return $b }
    }
    return (Get-ChangelogSection $text '\[1\.6\.0\]')
}
$clAuto = Get-AutoChangelogBlock $cl
$notesNorm = ($notesText -replace '\s+', ' ')
Check ($autoTut.Length -gt 3000 -and $autoRec.Length -gt 1500 -and $clAuto.Length -gt 1000) "atualização automática: subseção do tutorial ($($autoTut.Length)), Recursos do README ($($autoRec.Length)) e CHANGELOG da 1.6.0 ($($clAuto.Length)) lidos"

$labelPins = @($lblOn, $lblNew, $lblAuto, $lblOnce)
$clAutoPhrases = @(
    $lblOn, $lblNew, $lblAuto, $lblOnce, $debTextLong, $waitText, $retryText, $tagAuto, $tagIdx, 'nunca abre pasta', 'Core_PptSyncShapes', 'verificado em PowerPoint real',
    '### Adicionado', '### Alterado', 'sem diferenciar maiúsculas', 'Desfazer do Excel')
# Estrutura que o publicar-release espera no bloco da 1.6.0: "### Adicionado" e pelo menos 4 entradas "- ". Devolve as frases que faltam.
function Get-AutoChangelogMissing([string]$block) {
    $miss = @($clAutoPhrases | Where-Object { -not $block.Contains($_) })
    if (-not ($block -match '(?m)^### Adicionado\s*$' -and @([regex]::Matches($block, '(?m)^\s*[-*]\s+\S')).Count -ge 4)) { $miss += 'estrutura (### Adicionado e 4 entradas)' }
    return $miss
}
# sintético: o CHANGELOG depois do publicar-release 1.6.0 ([Não lançado] vazia, as linhas na [1.6.0]) passa; sem as linhas, falha
$clRel = "# Changelog`r`n`r`n## [Não lançado]`r`n`r`n## [1.6.0] - 2026-10-20`r`n" + ($clAuto -replace "`n", "`r`n") + "## [1.5.0] - 2026-10-15`r`n`r`n- x`r`n"
$clRelEmpty = "# Changelog`n`n## [Não lançado]`n`n## [1.6.0] - 2026-10-20`n`n- outra coisa`n`n## [1.5.0] - 2026-10-15`n`n- x`n"
Check (@(Get-AutoChangelogMissing (Get-AutoChangelogBlock $clRel)).Count -eq 0) 'CHANGELOG (sintético): depois do publicar-release 1.6.0 as linhas na [1.6.0] satisfazem os pinos'
Check (@(Get-AutoChangelogMissing (Get-AutoChangelogBlock $clRelEmpty)).Count -gt 0) 'CHANGELOG (sintético): [1.6.0] publicada sem as linhas da atualização automática = falha'
Check (@(Get-AutoChangelogMissing (Get-AutoChangelogBlock "## [Não lançado]`n`n## [1.5.0] - 2026-10-15`n`n- x`n")).Count -gt 0) 'CHANGELOG (sintético): sem as linhas em [Não lançado] nem [1.6.0] = falha'
$autoPlaces = [ordered]@{
    'tutorial #ppt-auto' = @{ text = $autoTut; phrases = @(
        'por gráfico', $tagAuto, $tagIdx, 'Nada vira automático sem uma ação sua', '(padrão: ligada)', '(padrão: desligada)',
        $debTextLong, 'de 1 a 2 s', $waitTextLong, 'nunca abre uma pasta', 'já abertas', 'somente leitura', $retryText, 'PptAuto: desistiu',
        'Pasta não salva', 'Excel fechado', 'Desfazer do Excel', 'sem diferenciar maiúsculas', 'ignora a caixa', 'Slides copiados de outra apresentação',
        'Atualize o ThinQQ no Excel para a atualização automática funcionar (a marca foi gravada).', 'verificado em PowerPoint real', 'aguardando ação OLE',
        'itens 61 a 68', 'só lê, de cada apresentação', 'volta ao texto que estava antes', 'espera e roda depois', 'não muda o modo', 'valem para todas as suas pastas',
        'Valores ainda não salvos entram', 'não usa a janela do PowerPoint') }
    'README Recursos' = @{ text = $autoRec; phrases = @(
        $tagAuto, $tagIdx, 'Nada vira automático sem uma ação do usuário', '(padrão: ligada;', '(padrão: desligada;', $debTextLong, $waitText, 'pasta fechada nunca é aberta',
        'já abertas', $retryText, 'Desfazer do Excel', 'maiúsculas de minúsculas', 'verificado em PowerPoint real', 'itens 61 a 68', 'aguardando ação OLE',
        'Atualize o ThinQQ no Excel para a atualização automática funcionar', 'somente leitura', 'sem usar a janela do PowerPoint', 'Slides copiados de outra apresentação'.ToLowerInvariant()) }
    'README Limitações' = @{ text = $readmeLimTxt; phrases = @(
        'Atualização automática no PowerPoint (1.6.0)', 'nada vira automático sem uma ação do usuário', 'já abertas', 'nunca abre o PowerPoint, uma apresentação nem uma pasta',
        $debText, $waitText, $retryText, 'Desfazer do Excel pode ser limpo', 'marca a apresentação como modificada', 'Atualize o ThinQQ no Excel para a atualização automática funcionar',
        'verificado em PowerPoint real', 'itens 61 a 68', 'Slides copiados de outra apresentação', 'somente leitura') }
    'CHANGELOG ([Não lançado], [1.7.0] ou [1.6.0])' = @{ text = $clAuto; phrases = $clAutoPhrases }
    'release/notes.txt' = @{ text = $notesNorm; phrases = @($lblAuto, $lblOnce, $lblOn, $lblNew, "$debounce segundos", 'Nada vira automático sem uma ação sua', 'por gráfico') }
}
foreach ($place in $autoPlaces.Keys) {
    $pt = [string]$autoPlaces[$place].text
    foreach ($ph in $autoPlaces[$place].phrases) { Check ($pt.Contains($ph)) "atualização automática, $place cita '$ph'" }
}
# os quatro rótulos também aparecem na seção #powerpoint, no README (Recursos) e nos itens manuais, com o caminho de menu certo
foreach ($l in $labelPins) { Check ($autoRec.Contains($l)) "README Recursos cita o rótulo '$l'" }
Check ($autoRec.Contains('aba ThinQQ > Gráfico selecionado > Atualizar automaticamente') -and $autoRec.Contains('No Excel, ThinQQ > PowerPoint:')) 'README Recursos: caminhos de menu do PowerPoint e do Excel'
Check ($readmeText.Contains('Arquitetura') -and $readme.Contains('**Atualização automática (1.6.0):** `clsAppEvents`') -and $readme.Contains('`PptAuto_Tick`') -and $readme.Contains('API 2')) 'README (Arquitetura): o motor automático (clsAppEvents, PptAuto_OnChange, PptAuto_Tick, API 2)'
# a CHANGELOG: estrutura que o publicar-release espera (cabeçalho, ### e entradas "- ") e as linhas do delta ficam em [1.5.0]
Check ($clAuto -match '(?m)^### Adicionado\s*$' -and @([regex]::Matches($clAuto, '(?m)^\s*[-*]\s+\S')).Count -ge 4) 'CHANGELOG da 1.6.0 ([Não lançado] ou, depois do publicar-release, [1.6.0]): seção ### Adicionado e pelo menos 4 entradas "- "'
Check ($clUnrel -notmatch 'Delta \(1\.5\.0\)') 'CHANGELOG [Não lançado]: sem linhas da 1.5.0 (o delta já está na [1.5.0])'
# notes.txt (1.6.0): de 1 a 8 linhas com o marcador, sem texto da 1.5.0
Check ($notesLines.Count -ge 3 -and $notesLines.Count -le 8) "notes.txt (1.6.0): de 3 a 8 linhas ($($notesLines.Count))"
Check ($notesText -notmatch 'delta entre quaisquer duas barras') 'notes.txt (1.6.0): sem as notas da 1.5.0 (o delta)'

# README: os itens manuais 61 a 68 (um por frase-chave do roteiro, por número) e todos se declaram não verificados em PowerPoint real
$mvAutoLines = @(($mvBlock -split "\r?\n") | Where-Object { $_ -match '^(\d+)\. ' -and [int]$Matches[1] -ge 61 -and [int]$Matches[1] -le 68 })
Check ($mvAutoLines.Count -eq 8 -and @($mvNums | Where-Object { $_ -ge 61 -and $_ -le 68 }).Count -eq 8) "README: os itens manuais 61 a 68 existem, numerados sem buracos ($($mvAutoLines.Count))"
Check ($mvNums.Count -ge 68) "README: a lista de verificações manuais vai pelo menos até o 68 ($($mvNums.Count))"
Check (@($mvAutoLines | Where-Object { $_ -notmatch 'não verificado em PowerPoint real' }).Count -eq 0) 'README (itens 61 a 68): todo item marca "não verificado em PowerPoint real"'
$mvAutoPins = @{
    61 = @('digitar valores', 'em até 3 s', 'Ctrl+Z', $lblAuto, 'volta ao manual')
    62 = @('duas apresentações abertas', 'mesma pasta', 'somente leitura')
    63 = @('texto em edição', 'aguardando ação OLE', 'sonda P3', $retryText, 'PptAuto: desistiu')
    64 = @('somente leitura', $lblOnce, 'só na memória')
    65 = @($lblOn, $lblNew, 'Copiar como gráfico vinculado', 'nasce automático', 'nasce manual')
    66 = @('20 gráficos', 'tempo', 'Excel fica ocupado', $waitText)
    67 = @('salvar a apresentação', 'continua automático', 'descarregar o ThinQQ', 'ciclo pendente')
    68 = @('.ppam', '1.5.x', 'só grava a marca', 'Atualize o ThinQQ no Excel para a atualização automática funcionar', 'Atualize o ThinQQ no Excel para atualizar só os gráficos selecionados')
}
foreach ($n in ($mvAutoPins.Keys | Sort-Object)) {
    $line = @($mvAutoLines | Where-Object { $_ -match "^$n\. " })
    foreach ($ph in $mvAutoPins[$n]) { Check ($line.Count -eq 1 -and $line[0].Contains($ph)) "README item manual ${n}: cita '$ph'" }
}

# ---- Inserir também no slide (1.7.0, Tarefa D3): subseção #ppt-inserir, Recursos e Limitações do README, itens manuais 69 a 74 e a [Não lançado] do
# CHANGELOG. As linhas da 1.7.0 ficam em [Não lançado] (junto com as da 1.6.0) até o publicar-release 1.7.0, que as move para "## [1.7.0]": os pinos
# leem só essas duas seções. Dentro de "## [1.6.0]" elas anunciariam a 1.7.0 como entregue na 1.6.0 (merge depois do publicar-release 1.6.0, como
# no delta): um pino recusa (README, "Publicar uma versão").
$lblIns = 'Inserir também no slide'; $lblFrames = 'Mostrar/ocultar molduras'
$insNode = $ribbon.SelectSingleNode("//*[@id='chkPptInsert']"); $frNode = $ribbon.SelectSingleNode("//*[@id='btnPptFrames']")
Check ($null -ne $insNode -and $insNode.LocalName -eq 'checkBox' -and $insNode.GetAttribute('label') -ceq $lblIns) "faixa do Excel: a caixa chkPptInsert tem o rótulo '$lblIns'"
Check ($null -ne $frNode -and $frNode.LocalName -eq 'button' -and $frNode.GetAttribute('label') -ceq $lblFrames) "faixa do Excel: o botão btnPptFrames tem o rótulo '$lblFrames'"
if ($parsed) {
    Check ($null -ne $xml.SelectSingleNode("//*[local-name()='section'][@id='powerpoint']//*[local-name()='h3'][@id='ppt-inserir']")) 'tutorial: a subseção id="ppt-inserir" está dentro de #powerpoint'
    Check ($xml.SelectNodes("//a[@href='#ppt-inserir']").Count -ge 3) 'tutorial: a subseção #ppt-inserir tem links de #powerpoint, da referência da faixa e dos limites'
    foreach ($l in $lblIns, $lblFrames) { Check ($pptUi -contains $l) "tutorial: '$l' citado como span.ui dentro de #powerpoint" }
}
$insTut = Get-SubsectionText 'ppt-inserir'
$insRec = ([regex]::Match($readme, '(?s)\*\*Inserir também no slide \(1\.7\.0\)\*\*(?<b>.*?)### Galeria').Groups['b'].Value -replace '\s+', ' ')
function Get-InsertChangelogBlock([string]$text) {
    $acc = ''
    foreach ($h in '\[Não lançado\]', '\[1\.7\.0\]') { $acc += "`n" + (Get-ChangelogSection $text $h) }
    return $acc
}
# True se existe "## [1.6.0]" e ela contém linhas da 1.7.0 (a 1.7.0 anunciada como entregue na 1.6.0)
function Test-InsertInside160([string]$text) {
    $t = $text -replace "`r`n", "`n"
    if ($t -notmatch '(?m)^## \[1\.6\.0\]') { return $false }
    return (Get-ChangelogSection $text '\[1\.6\.0\]').Contains('Inserir no slide (1.7.0)')
}
$clInsPhrases = @($lblIns, $lblFrames, 'Inserir no slide (1.7.0)', 'moldura verde', 'não muda o vínculo', 'Alterar dados', 'Atualize o ThinQQ para usar este recurso.', 'não cria nada', 'sem moldura')
function Get-InsertChangelogMissing([string]$block) { return @($clInsPhrases | Where-Object { -not $block.Contains($_) }) }
$clIns = Get-InsertChangelogBlock $cl
Check ($insTut.Length -gt 2500 -and $insRec.Length -gt 1500 -and $clIns.Length -gt 1000) "inserir no slide: subseção do tutorial ($($insTut.Length)), Recursos do README ($($insRec.Length)) e CHANGELOG ($($clIns.Length)) lidos"
$clInsLine = "- Inserir no slide (1.7.0): " + (($clInsPhrases | Where-Object { $_ -ne 'Inserir no slide (1.7.0)' }) -join ' ') + "`n"
Check (@(Get-InsertChangelogMissing (Get-InsertChangelogBlock "## [Não lançado]`n`n$clInsLine`n## [1.5.0] - 2026-10-15`n`n- x`n")).Count -eq 0) 'CHANGELOG (sintético): as linhas da 1.7.0 em [Não lançado] satisfazem os pinos'
$clIn160 = "## [Não lançado]`n`n## [1.6.0] - 2026-10-20`n`n### Adicionado`n$clInsLine`n## [1.5.0] - 2026-10-15`n`n- x`n"
Check ((Test-InsertInside160 $clIn160) -and @(Get-InsertChangelogMissing (Get-InsertChangelogBlock $clIn160)).Count -gt 0) 'CHANGELOG (sintético): as linhas da 1.7.0 dentro da [1.6.0] (merge depois do publicar-release 1.6.0) = falha'
# tudo junto como 1.7.0: o publicar-release 1.7.0 move a [Não lançado] inteira (linhas da 1.6.0 e da 1.7.0) para a [1.7.0]: ok para os dois blocos
$clRoll = "# Changelog`n`n## [Não lançado]`n`n## [1.7.0] - 2026-10-30`n" + ($clIns -replace "`r`n", "`n") + "`n## [1.5.0] - 2026-10-15`n`n- x`n"
Check (@(Get-AutoChangelogMissing (Get-AutoChangelogBlock $clRoll)).Count -eq 0 -and @(Get-InsertChangelogMissing (Get-InsertChangelogBlock $clRoll)).Count -eq 0 -and -not (Test-InsertInside160 $clRoll)) 'CHANGELOG (sintético): 1.6.0 e 1.7.0 publicadas juntas na [1.7.0] satisfazem os pinos das duas'
Check (-not (Test-InsertInside160 $cl)) 'CHANGELOG: a [1.6.0] não contém linhas "Inserir no slide (1.7.0)" (mova-as para a [Não lançado]; README, Publicar uma versão)'
Check ($cl -notmatch '\.- [A-Z]') 'CHANGELOG: nenhuma entrada colada no fim de outra (".- X" sem quebra de linha)'
Check (@(Get-InsertChangelogMissing (Get-InsertChangelogBlock "## [Não lançado]`n`n## [1.7.0] - 2026-10-30`n`n$clInsLine`n## [1.6.0] - 2026-10-20`n`n- outra coisa`n")).Count -eq 0) 'CHANGELOG (sintético): depois do publicar-release 1.7.0 as linhas na [1.7.0] satisfazem os pinos'
Check (@(Get-InsertChangelogMissing (Get-InsertChangelogBlock "## [Não lançado]`n`n## [1.6.0] - 2026-10-20`n`n- outra coisa`n`n## [1.5.0] - 2026-10-15`n`n- x`n")).Count -gt 0) 'CHANGELOG (sintético): sem as linhas da 1.7.0 em lugar nenhum = falha'
$insPlaces = [ordered]@{
    'tutorial #ppt-inserir' = @{ text = $insTut; phrases = @(
        $lblIns, $lblFrames, 'Salve a pasta de trabalho antes de inserir no PowerPoint', 'Nada foi criado', 'Esta pasta não tem molduras de intervalos vinculados ao PowerPoint',
        'moldura verde', 'não muda o vínculo', 'Alterar dados', 'Só os gráficos inseridos por este caminho ganham moldura', 'Não é impressa', 'vermelha e azul',
        'sem moldura', 'sem vínculo com a pasta', 'desmarcada', 'cinco menus', 'sonda P4', 'itens 69 a 74', 'Atualize o ThinQQ para usar este recurso.', 'não insere nada',
        'Não verificado em PowerPoint e Excel reais', 'é só um indicador', 'também existe no Excel') }
    'README Recursos' = @{ text = $insRec; phrases = @(
        $lblIns, $lblFrames, 'Moldura verde', 'não muda o vínculo', 'Alterar dados', 'só os gráficos inseridos por este caminho ganham moldura', 'vermelha e azul', 'não é impressa',
        'Nada foi criado', 'sem moldura', 'Atualize o ThinQQ para usar este recurso.', 'itens 69 a 74', 'sonda P4', 'Não verificado em Excel e PowerPoint reais', 'volta desmarcada') }
    'README Limitações' = @{ text = $readmeLimTxt; phrases = @(
        'Inserir também no slide (1.7.0)', 'não muda o vínculo', 'só os gráficos inseridos por este caminho ganham moldura', 'vermelha e azul', 'não é impressa',
        'Atualize o ThinQQ para usar este recurso.', 'itens 69 a 74', 'sonda P4', 'Alterar dados') }
    'CHANGELOG ([Não lançado] ou [1.7.0])' = @{ text = $clIns; phrases = $clInsPhrases }
}
foreach ($place in $insPlaces.Keys) {
    $pt = [string]$insPlaces[$place].text
    foreach ($ph in $insPlaces[$place].phrases) { Check ($pt.Contains($ph)) "inserir no slide, $place cita '$ph'" }
}
Check ($readme.Contains('`Core_InsertToPowerPoint(kind)`') -and $readme.Contains('`modLinkFrame`') -and $readme.Contains('`Core_ToggleLinkFrames`')) 'README (Arquitetura): o fluxo de inserir no slide (Core_InsertToPowerPoint, modLinkFrame, Core_ToggleLinkFrames)'
$mvInsLines = @(($mvBlock -split "\r?\n") | Where-Object { $_ -match '^(\d+)\. ' -and [int]$Matches[1] -ge 69 })
Check ($mvInsLines.Count -eq 6 -and @($mvNums | Where-Object { $_ -ge 69 -and $_ -le 74 }).Count -eq 6 -and $mvNums.Count -ge 74) "README: os itens manuais 69 a 74 existem, numerados sem buracos ($($mvInsLines.Count); total $($mvNums.Count))"
Check (@($mvInsLines | Where-Object { $_ -notmatch 'não verificado em (Excel|PowerPoint) real' }).Count -eq 0) 'README (itens 69 a 74): todo item marca "não verificado em Excel real" ou "em PowerPoint real"'
$mvInsPins = @{
    69 = @('sonda P4', 'clicar', $lblIns, 'quatro linhas')
    70 = @('imprimir', 'não aparece no papel')
    71 = @('inserir uma linha', 'apagar uma linha', 'Alterar dados')
    72 = @('Mover ou copiar', 'moldura órfã')
    73 = @('PowerPoint fechado', 'cinco menus', 'Novos vínculos atualizam automaticamente')
    74 = @('Nada foi criado', $lblFrames, 'Atualize o ThinQQ para usar este recurso.', 'sem moldura')
}
foreach ($n in ($mvInsPins.Keys | Sort-Object)) {
    $line = @($mvInsLines | Where-Object { $_ -match "^$n\. " })
    foreach ($ph in $mvInsPins[$n]) { Check ($line.Count -eq 1 -and $line[0].Contains($ph)) "README item manual ${n}: cita '$ph'" }
}
if ($script:fail -gt 0) { Write-Host "FALHAS Test-Docs: $script:fail"; exit 1 }
Write-Host 'Test-Docs: OK'
