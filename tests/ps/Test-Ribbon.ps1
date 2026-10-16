# Verifica o pacote ThinQQ.test.xlam: customUI14.xml (identico ao fonte no modo Png), a relacao ui/extensibility, e os icones proprios:
# para cada image="x" do XML do pacote, customUI/images/x.png (PNG valido, 32 px ou 16 px conforme o sufixo "16"), a Relationship Id="x"
# em customUI/_rels/customUI14.xml.rels e o content type png. No modo Mso (sem image=) nao pode haver imagens.
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$ErrorActionPreference = 'Stop'   # um erro inesperado nunca pode virar verde em silencio (PowerShell 5.1 e 7)
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Drawing
. (Join-Path $RepoRoot 'build\lib\Ribbon-Icons.ps1')
$pkg = Join-Path $RepoRoot 'dist\ThinQQ.test.xlam'
$src = Join-Path $RepoRoot 'src\loader\ribbon\customUI14.xml'
$script:fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }
# Nomes image= distintos dos ELEMENTOS do XML (DOM: um image= citado num comentario nao conta)
function Get-DomImages([string]$x) { $d = New-Object Xml.XmlDocument; $d.LoadXml($x); return ,[string[]]@($d.SelectNodes('//*[@image]') | ForEach-Object { $_.GetAttribute('image') } | Sort-Object -Unique) }
function Read-Entry($zip, [string]$name) { $e = $zip.GetEntry($name); if (-not $e) { return $null }; $ms = New-Object System.IO.MemoryStream; $s = $e.Open(); $s.CopyTo($ms); $s.Dispose(); return ,$ms.ToArray() }
if (-not (Test-Path -LiteralPath $pkg)) { Write-Host "not ok - pacote ausente: $pkg"; exit 1 }
$fs = New-Object System.IO.FileStream($pkg, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
$zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
try {
    $xmlBytes = Read-Entry $zip 'customUI/customUI14.xml'
    Check ($null -ne $xmlBytes) 'pacote contem customUI/customUI14.xml'
    $xml = [Text.Encoding]::UTF8.GetString($xmlBytes)
    $images = Get-DomImages $xml
    $srcXml = [IO.File]::ReadAllText($src)
    $srcImages = Get-DomImages $srcXml
    Check ($srcImages.Count -ge 13) "fonte: icones proprios em image= ($($srcImages.Count) nomes distintos)"
    if ($images.Count -gt 0 -or $srcImages.Count -eq 0) {
        # modo Png (ou fonte ainda sem image=, Tarefas 6 e 7): o XML do pacote e o fonte, byte a byte
        Check (([Convert]::ToBase64String($xmlBytes)) -eq ([Convert]::ToBase64String([IO.File]::ReadAllBytes($src)))) 'customUI14.xml do pacote identico ao fonte'
        Check (($images -join ',') -eq ($srcImages -join ',')) "modo Png: as imagens do pacote sao as do fonte ($($images.Count))"
    } else {
        # modo Mso: o XML do pacote e exatamente o fonte convertido (cada image= trocado pela reserva da tabela)
        Check ($xml -eq (Convert-TqqRibbonToMso -Xml $srcXml)) 'modo Mso: customUI14.xml do pacote = fonte convertido para imageMso'
    }
    if ($images.Count -eq 0) { Check ($null -eq ($zip.Entries | Where-Object { $_.FullName -like 'customUI/images/*' })) 'pacote sem image=: sem imagens' }
    $r = Read-Entry $zip '_rels/.rels'
    Check ($null -ne $r) 'pacote contem _rels/.rels'
    if ($r) {
        $txt = [Text.Encoding]::UTF8.GetString($r)
        Check ($txt -match 'http://schemas\.microsoft\.com/office/2007/relationships/ui/extensibility') 'relationship ui/extensibility presente'
        Check ($txt -match 'customUI/customUI14\.xml') 'relationship aponta para customUI/customUI14.xml'
    }
    $specs = Get-TqqIconSpecs
    $relsBytes = Read-Entry $zip 'customUI/_rels/customUI14.xml.rels'
    $ct = [Text.Encoding]::UTF8.GetString((Read-Entry $zip '[Content_Types].xml'))
    if ($images.Count -gt 0) {
        Check ($null -ne $relsBytes) 'pacote contem customUI/_rels/customUI14.xml.rels'
        Check ($ct -match '<Default Extension="png" ContentType="image/png"\s*/>') 'content type png declarado'
        $rels = if ($relsBytes) { [Text.Encoding]::UTF8.GetString($relsBytes) } else { '' }
        foreach ($img in $images) {
            $base = $img -replace '16$', ''
            Check ($specs.Contains($base)) "icone '$img' existe na tabela de icones ($base)"
            $png = Read-Entry $zip "customUI/images/$img.png"
            Check ($null -ne $png) "pacote contem customUI/images/$img.png"
            if ($png) {
                $sig = [Convert]::ToBase64String($png[0..7])
                Check ($sig -eq [Convert]::ToBase64String([byte[]]@(0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A))) "$img.png: assinatura PNG"
                $w = ($png[16] -shl 24) -bor ($png[17] -shl 16) -bor ($png[18] -shl 8) -bor $png[19]
                $expected = if ($img -like '*16') { 16 } else { 32 }
                Check ($w -eq $expected) "$img.png: largura $expected ($w)"
                try { $ms = New-Object IO.MemoryStream(,$png); $bmp = [System.Drawing.Image]::FromStream($ms); Check ($bmp.Width -eq $expected -and $bmp.Height -eq $expected) "$img.png: System.Drawing abre $($bmp.Width)x$($bmp.Height)"; $bmp.Dispose() } catch { Check $false "$img.png: System.Drawing nao abre: $($_.Exception.Message)" }
            }
            Check ($rels -match ('Id="' + [regex]::Escape($img) + '"[^>]*Target="images/' + [regex]::Escape($img) + '\.png"')) "relacao Id=$img -> images/$img.png"
        }
        Check ($rels -match 'http://schemas\.openxmlformats\.org/officeDocument/2006/relationships/image') 'relacoes do tipo image'
    }
    # vetor de teste padrao do CRC-32 (IEEE 802.3, o do PNG): "123456789" -> CBF43926
    $crc = '{0:X8}' -f (Get-TqqCrc32 ([Text.Encoding]::ASCII.GetBytes('123456789')))
    Check ($crc -eq 'CBF43926') "CRC-32('123456789') = CBF43926 ($crc)"
    # determinismo do gravador: duas gravacoes, mesmos bytes
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ("tqq_icons_" + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force $tmp | Out-Null
    try {
        $a = Export-TqqIcons -OutDir $tmp
        $h1 = ($a | ForEach-Object { (Get-FileHash $_).Hash }) -join ','
        Remove-Item (Join-Path $tmp '*') -Force
        $b = Export-TqqIcons -OutDir $tmp
        $h2 = ($b | ForEach-Object { (Get-FileHash $_).Hash }) -join ','
        Check ($h1 -eq $h2 -and $a.Count -eq 2 * $specs.Count) "gravador de PNG deterministico ($($a.Count) arquivos para $($specs.Count) icones)"
        Check ((Get-DomImages (Convert-TqqRibbonToMso -Xml $srcXml)).Count -eq 0) 'conversao para imageMso nao deixa image= em nenhum controle'
    } finally { Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue }
} finally { $zip.Dispose(); $fs.Dispose() }
# Empacotamento com image= (o XML fonte so ganha image= na Tarefa 8): monta pacotes sinteticos numa pasta temporaria, nos dois modos,
# e confere imagens, relacoes, content type, idempotencia e a recusa de nome fora da tabela.
. (Join-Path $RepoRoot 'build\lib\Build-Common.ps1')
$tmpP = Join-Path ([IO.Path]::GetTempPath()) ("tqq_pkg_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $tmpP | Out-Null
# Onda B: as pastas temporarias de icones das chamadas deste teste nascem numa pasta propria (um build concorrente cria as suas em %TEMP%:
# contar thinqq_icons_* la ficava instavel); o teste confere que nenhuma sobra nela.
$iconRoot = Join-Path $tmpP 'icons'; New-Item -ItemType Directory -Force $iconRoot | Out-Null
$PSDefaultParameterValues['Add-RibbonToPackage:IconTempRoot'] = $iconRoot
try {
    $synXml = '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="t" label="T"><group id="g" label="G"><button id="b1" label="A" image="tqq_columns" size="large"/><button id="b2" label="B" image="tqq_columns16"/><button id="b3" label="C" image="tqq_ppt"/></group></tab></tabs></ribbon></customUI>'
    $synPath = Join-Path $tmpP 'syn.xml'; [IO.File]::WriteAllText($synPath, $synXml, (New-Object Text.UTF8Encoding($false)))
    $badPath = Join-Path $tmpP 'bad.xml'; [IO.File]::WriteAllText($badPath, $synXml.Replace('tqq_ppt', 'tqq_inexistente'), (New-Object Text.UTF8Encoding($false)))
    $pp = Join-Path $tmpP 'p.xlam'; Copy-Item -LiteralPath $pkg -Destination $pp
    function Get-PkgInfo([string]$p) {
        $z = [System.IO.Compression.ZipFile]::OpenRead($p)
        try {
            $names = @($z.Entries | ForEach-Object { $_.FullName })
            $rd = { param($n) $e = $z.GetEntry($n); if (-not $e) { return $null }; $ms = New-Object IO.MemoryStream; $s = $e.Open(); $s.CopyTo($ms); $s.Dispose(); , $ms.ToArray() }
            return @{ Names = $names; Xml = (& $rd 'customUI/customUI14.xml'); Rels = (& $rd 'customUI/_rels/customUI14.xml.rels'); Ct = [Text.Encoding]::UTF8.GetString((& $rd '[Content_Types].xml')); Root = [Text.Encoding]::UTF8.GetString((& $rd '_rels/.rels')); Png32 = (& $rd 'customUI/images/tqq_columns.png'); Png16 = (& $rd 'customUI/images/tqq_columns16.png') }
        } finally { $z.Dispose() }
    }
    Add-RibbonToPackage -PackagePath $pp -RibbonXmlPath $synPath -Icons Png
    Add-RibbonToPackage -PackagePath $pp -RibbonXmlPath $synPath -Icons Png   # 2a vez: idempotente
    $i = Get-PkgInfo $pp
    Check (([Convert]::ToBase64String($i.Xml)) -eq ([Convert]::ToBase64String([IO.File]::ReadAllBytes($synPath)))) 'sintetico Png: XML gravado byte a byte'
    foreach ($nm in 'tqq_columns', 'tqq_columns16', 'tqq_ppt') {
        Check ($i.Names -contains "customUI/images/$nm.png") "sintetico Png: imagem $nm.png no pacote"
        Check ([Text.Encoding]::UTF8.GetString($i.Rels) -match ('Id="' + $nm + '"[^>]*Target="images/' + $nm + '\.png"')) "sintetico Png: relacao Id=$nm"
    }
    Check (@($i.Names | Where-Object { $_ -like 'customUI/images/*' }).Count -eq 3) 'sintetico Png: so as 3 imagens referenciadas'
    Check ((@([regex]::Matches($i.Ct, 'Extension="png"'))).Count -eq 1) 'sintetico Png: content type png declarado uma vez (idempotente)'
    Check ((@([regex]::Matches($i.Root, 'customUI/customUI14\.xml'))).Count -eq 1) 'sintetico Png: relationship ui/extensibility uma vez'
    foreach ($pair in @(@($i.Png32, 32), @($i.Png16, 16))) {
        $ms = New-Object IO.MemoryStream(, [byte[]]$pair[0]); $bm = [System.Drawing.Image]::FromStream($ms)
        Check ($bm.Width -eq $pair[1] -and $bm.Height -eq $pair[1]) "sintetico Png: PNG $($pair[1]) px abre no System.Drawing"
        if ($pair[1] -eq 32) { $c = ([System.Drawing.Bitmap]$bm).GetPixel(8, 20); Check ($c.A -eq 255 -and $c.B -eq 100) "sintetico Png: pixel do 1o retangulo (navy) = $($c.R),$($c.G),$($c.B)" }
        $bm.Dispose()
    }
    # Mso sobre um pacote que ja tem imagens: remove imagens e relacoes, XML convertido
    Add-RibbonToPackage -PackagePath $pp -RibbonXmlPath $synPath -Icons Mso
    $i = Get-PkgInfo $pp
    Check ($null -eq ($i.Names | Where-Object { $_ -like 'customUI/images/*' -or $_ -eq 'customUI/_rels/customUI14.xml.rels' })) 'sintetico Mso: sem imagens nem relacoes de imagem'
    Check ([Text.Encoding]::UTF8.GetString($i.Xml) -eq (Convert-TqqRibbonToMso -Xml $synXml)) 'sintetico Mso: XML convertido'
    Check ([Text.Encoding]::UTF8.GetString($i.Xml) -notmatch 'image="' -and [Text.Encoding]::UTF8.GetString($i.Xml) -match 'imageMso="ChartTypeColumnInsertGallery"') 'sintetico Mso: image= virou imageMso'
    # nome fora da tabela: recusa clara e nenhuma pasta temporaria de icones sobra
    $thrown = $false; try { Add-RibbonToPackage -PackagePath $pp -RibbonXmlPath $badPath -Icons Png } catch { $thrown = $_.Exception.Message -match 'tqq_inexistente' }
    Check $thrown 'sintetico: image= fora da tabela e recusado pelo nome'
    Check ((Get-Command Add-RibbonToPackage).Parameters.ContainsKey('IconTempRoot')) 'Add-RibbonToPackage aceita a pasta das pastas temporarias de icones (o padrao deste teste vale)'
    Check (@(Get-ChildItem $iconRoot -Directory -Filter 'thinqq_icons_*' -ErrorAction SilentlyContinue).Count -eq 0) 'sintetico: nenhuma pasta temporaria de icones sobra'
    # Revalidacao (Tarefa 8): comentarios nao contam; image + imageMso e nome fora da tabela sao recusados nos dois modos; o XML convertido
    # para Mso e relido como XML (um controle com image e imageMso viraria um atributo imageMso repetido) e nao pode sobrar image=.
    function Write-Syn([string]$name, [string]$body) { $p = Join-Path $tmpP $name; [IO.File]::WriteAllText($p, '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui"><ribbon><tabs><tab id="t" label="T"><group id="g" label="G">' + $body + '</group></tab></tabs></ribbon></customUI>', (New-Object Text.UTF8Encoding($false))); return $p }
    function Test-Throws([scriptblock]$sb, [string]$pattern) { try { & $sb; return $false } catch { return ($_.Exception.Message -match $pattern) } }
    $cmt = Write-Syn 'cmt.xml' '<!-- exemplo: image="tqq_inexistente" --><button id="b1" label="A" image="tqq_columns"/>'
    Add-RibbonToPackage -PackagePath $pp -RibbonXmlPath $cmt -Icons Png
    $i = Get-PkgInfo $pp
    Check ((@($i.Names | Where-Object { $_ -like 'customUI/images/*' }) -join ',') -eq 'customUI/images/tqq_columns.png') 'revalidacao Png: image= citado em comentario nao vira imagem nem recusa'
    Add-RibbonToPackage -PackagePath $pp -RibbonXmlPath $cmt -Icons Mso
    $i = Get-PkgInfo $pp; $mx = [Text.Encoding]::UTF8.GetString($i.Xml)
    Check ($mx.Contains('<!-- exemplo: image="tqq_inexistente" -->') -and $mx.Contains('id="b1" label="A" imageMso="ChartTypeColumnInsertGallery"')) 'revalidacao Mso: comentario intacto, so o atributo do controle convertido'
    $both = Write-Syn 'both.xml' '<button id="bAmbos" label="A" image="tqq_columns" imageMso="Help"/>'
    Check (Test-Throws { Add-RibbonToPackage -PackagePath $pp -RibbonXmlPath $both -Icons Png } 'bAmbos.*image e imageMso') 'revalidacao Png: controle com image e imageMso recusado pelo id'
    Check (Test-Throws { Add-RibbonToPackage -PackagePath $pp -RibbonXmlPath $both -Icons Mso } 'bAmbos.*image e imageMso') 'revalidacao Mso: controle com image e imageMso recusado pelo id'
    $foo = Write-Syn 'foo.xml' '<button id="bFoo" label="A" image="foo"/>'
    Check (Test-Throws { Add-RibbonToPackage -PackagePath $pp -RibbonXmlPath $foo -Icons Mso } '"foo"') 'revalidacao Mso: image fora da tabela (sem prefixo tqq_) recusado'
    Check (Test-Throws { Add-RibbonToPackage -PackagePath $pp -RibbonXmlPath $foo -Icons Png } '"foo"') 'revalidacao Png: image fora da tabela recusado'
    # a validacao do XML convertido, direto: image= que sobrou, atributo repetido (o que a conversao faria com image + imageMso) e controle sem id
    $ns = '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">'
    Check (Test-Throws { Assert-TqqRibbonXml -Xml ($ns + '<button id="bX" image="tqq_columns"/></customUI>') -Mode Mso } 'sobrou.*bX') 'Assert-TqqRibbonXml Mso: image= que sobrou e recusado'
    Check (Test-Throws { Assert-TqqRibbonXml -Xml ($ns + '<button id="bY" imageMso="Help" imageMso="ChartAxes"/></customUI>') -Mode Mso } 'XML invalido') 'Assert-TqqRibbonXml Mso: atributo imageMso repetido (XML invalido) e recusado'
    Check (Test-Throws { Assert-TqqRibbonXml -Xml ($ns + '<item image="tqq_columns" imageMso="Help"/></customUI>') -Mode Png } 'sem id') 'Assert-TqqRibbonXml: controle sem id na mensagem (sem erro de nulo)'
    $okNames = Assert-TqqRibbonXml -Xml ($ns + '<!-- image="tqq_nada" --><button id="a" image="tqq_ppt16"/><button id="b" image="tqq_ppt16"/><button id="c" image="tqq_bars"/></customUI>') -Mode Png
    Check (($okNames -join ',') -eq 'tqq_bars,tqq_ppt16') "Assert-TqqRibbonXml Png: nomes distintos, sem os de comentario ($($okNames -join ','))"
    $none = Assert-TqqRibbonXml -Xml ($ns + '<button id="a" imageMso="Help"/></customUI>') -Mode Mso
    Check ($null -ne $none -and $none.Count -eq 0) 'Assert-TqqRibbonXml Mso valido: lista vazia'
    Check (@(Get-ChildItem $iconRoot -Directory -Filter 'thinqq_icons_*' -ErrorAction SilentlyContinue).Count -eq 0) 'revalidacao: nenhuma pasta temporaria de icones sobra'
} finally {
    $PSDefaultParameterValues.Remove('Add-RibbonToPackage:IconTempRoot')   # rodada 1: o padrao nao vaza para o build Release seguinte (mesma sessao)
    Remove-Item $tmpP -Recurse -Force -ErrorAction SilentlyContinue
}
Check (-not $PSDefaultParameterValues.ContainsKey('Add-RibbonToPackage:IconTempRoot')) 'padrao de IconTempRoot removido da sessao'
Check (-not (Test-Path -LiteralPath $tmpP)) 'pasta tqq_pkg_* do teste removida'
# ---- Faixa do suplemento do PowerPoint (src/ppt/ribbon/customUI14.xml): bem formada, so imageMso, ids e callbacks esperados
$pptXmlPath = Join-Path $RepoRoot 'src\ppt\ribbon\customUI14.xml'
[xml]$px = [IO.File]::ReadAllText($pptXmlPath, (New-Object Text.UTF8Encoding($false)))
$pns = New-Object Xml.XmlNamespaceManager($px.NameTable); $pns.AddNamespace('u', 'http://schemas.microsoft.com/office/2009/07/customui')
Check ($px.customUI.onLoad -eq 'PptRx_OnLoad') 'ppt: onLoad PptRx_OnLoad'
# 1.8.0: o grupo "Inserir gráfico" usa os icones proprios (o build do .ppam embute PNG, como o do Excel): cada image= existe na tabela,
# nenhum controle tem image= e imageMso juntos, e filhos de menu nao tem size
$pptImgNames = Assert-TqqRibbonXml -Xml ([IO.File]::ReadAllText($pptXmlPath, (New-Object Text.UTF8Encoding($false)))) -Mode Png -What 'ppt customUI14.xml'
Check ($pptImgNames.Count -ge 5) "ppt: icones proprios no grupo Inserir gráfico ($($pptImgNames -join ','))"
Check ($null -eq $px.SelectSingleNode('//u:*[@image and @imageMso]', $pns)) 'ppt: nenhum controle com image= e imageMso juntos'
Check ($null -eq $px.SelectSingleNode('//u:menu/u:button[@size]', $pns)) 'ppt: botoes de menu sem size'
Check ($null -eq $px.SelectSingleNode("//u:*[@image and not(ancestor-or-self::u:group[@id='grpPptInsert'])]", $pns)) 'ppt: image= so no grupo Inserir gráfico (os demais seguem com imageMso)'
Check ($null -ne $px.SelectSingleNode("//u:button[@id='btnPptSyncNow'][@onAction='PptRx_OnSync']", $pns)) 'ppt: Atualizar gráficos'
Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptElem']/u:toggleButton[@id='tglPptAuto'][@getPressed='PptRx_GetAuto'][@onAction='PptRx_OnAuto'][@getEnabled='PptRx_GetElemEnabled']", $pns)) 'ppt: Atualizar automaticamente'
Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptElem']/u:button[@id='btnPptUpdateSel'][@onAction='PptRx_OnUpdateSel'][@getEnabled='PptRx_GetElemEnabled']", $pns)) 'ppt: Atualizar uma vez'
# Cada callback da faixa do PowerPoint existe como Public Sub em src/ppt/modules (nome errado = erro silencioso so no clique), com a
# assinatura certa nos de toggle/getPressed/getEnabled; ids unicos.
$pptBas = (Get-ChildItem (Join-Path $RepoRoot 'src\ppt\modules') -Filter *.bas | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n"
$pptCbs = @($px.SelectNodes('//@*') | Where-Object { $_.LocalName -match '^(onLoad|onAction|get[A-Z]\w*)$' } | ForEach-Object { $_.Value } | Sort-Object -Unique)
Check ($pptCbs.Count -ge 7) "ppt: callbacks lidos do XML ($($pptCbs -join ','))"
foreach ($cb in $pptCbs) { Check ($pptBas -match "(?m)^Public Sub $([regex]::Escape($cb))\(") "ppt: callback $cb existe (Public Sub)" }
Check ($pptBas -match '(?m)^Public Sub PptRx_OnAuto\(control As IRibbonControl, pressed As Boolean\)') 'ppt: assinatura do onAction do toggleButton'
Check ($pptBas -match '(?m)^Public Sub PptRx_GetAuto\(control As IRibbonControl, ByRef \w+\)') 'ppt: assinatura do getPressed'
Check ($pptBas -match '(?m)^Public Sub PptRx_GetElemEnabled\(control As IRibbonControl, ByRef \w+\)') 'ppt: assinatura do getEnabled'
$pptIds = @($px.SelectNodes('//u:*[@id]', $pns) | ForEach-Object { $_.GetAttribute('id') })
Check ($pptIds.Count -eq @($pptIds | Sort-Object -Unique).Count) "ppt: ids unicos ($($pptIds.Count))"
Check ((@($px.SelectNodes("//u:dropDown[@id='ddOpenMode']/u:item", $pns) | ForEach-Object { $_.id }) -join ',') -eq 'modeask,modealways,modenever') 'ppt: itens de Ao abrir'
# 1.8.0 (G4): "Inserir gráfico" com os 21 tipos do Excel, na mesma ordem, e "Editar dados"
$xx = New-Object Xml.XmlDocument; $xx.Load((Join-Path $RepoRoot 'src\loader\ribbon\customUI14.xml'))
$xns = New-Object Xml.XmlNamespaceManager($xx.NameTable); $xns.AddNamespace('u', 'http://schemas.microsoft.com/office/2009/07/customui')
$excelTags = @($xx.SelectNodes("//u:group[@id='grpInsert']//u:button[@onAction='Rx_OnInsert']", $xns) | ForEach-Object { $_.GetAttribute('tag') })
$pptTags = @($px.SelectNodes("//u:group[@id='grpPptInsert']//u:button[@onAction='PptRx_OnInsertChart']", $pns) | ForEach-Object { $_.GetAttribute('tag') })
Check ($pptTags.Count -eq 21) "ppt: 21 tipos em Inserir gráfico ($($pptTags.Count))"
Check (($excelTags -join ',') -eq ($pptTags -join ',')) 'ppt: mesmos tipos e mesma ordem dos menus do Excel'
Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptElem']/u:button[@id='btnPptEditData'][@onAction='PptRx_OnEditData'][@getEnabled='PptRx_GetElemEnabled']", $pns)) 'ppt: Editar dados'
Check ('grpPptInsert' -eq $px.SelectSingleNode("//u:tab[@id='tabThinQQPpt']/u:group[1]", $pns).GetAttribute('id')) 'ppt: Inserir gráfico é o primeiro grupo'
# mesmos rotulos e ids (prefixo "p") dos menus e botoes do Excel; o estado de "Editar dados" segue a selecao (invalidado com os outros)
$xlInsNodes = @($xx.SelectNodes("//u:group[@id='grpInsert']//u:*[self::u:menu or self::u:button]", $xns))
$ppInsNodes = @($px.SelectNodes("//u:group[@id='grpPptInsert']//u:*[self::u:menu or self::u:button]", $pns))
$sameIns = ($xlInsNodes.Count -eq $ppInsNodes.Count)
for ($i = 0; $sameIns -and $i -lt $xlInsNodes.Count; $i++) {
    if (('p' + $xlInsNodes[$i].GetAttribute('id')) -cne $ppInsNodes[$i].GetAttribute('id') -or $xlInsNodes[$i].GetAttribute('label') -cne $ppInsNodes[$i].GetAttribute('label') -or $xlInsNodes[$i].GetAttribute('image') -cne $ppInsNodes[$i].GetAttribute('image')) { $sameIns = $false }
}
Check $sameIns "ppt: menus e botoes de Inserir gráfico com os ids (prefixo p), rotulos e icones do Excel ($($ppInsNodes.Count))"
Check ($pptBas -match 'InvalidateControl "btnPptEditData"') 'ppt: Editar dados invalidado com a selecao'
Check ($pptBas -match '(?m)^Public Sub PptRx_OnInsertChart\(control As IRibbonControl\)' -and $pptBas -match '(?m)^Public Sub PptRx_OnEditData\(control As IRibbonControl\)') 'ppt: assinaturas de Inserir gráfico e Editar dados'
# 1.9.0 (G8): grupo "Gráfico do ThinQQ" depois do "Gráfico selecionado": 3 alternâncias, 2 listas e 3 botões, com os callbacks e os tags das opções
$gc = $px.SelectSingleNode("//u:group[@id='grpPptChart']", $pns)
Check ($null -ne $gc -and $gc.GetAttribute('label') -ceq 'Gráfico do ThinQQ') 'ppt: grupo Gráfico do ThinQQ'
$gcPrev = $px.SelectSingleNode("//u:group[@id='grpPptChart']/preceding-sibling::u:group[1]", $pns)
Check ($null -ne $gcPrev -and $gcPrev.GetAttribute('id') -eq 'grpPptElem') 'ppt: Gráfico do ThinQQ logo depois de Gráfico selecionado'
foreach ($t in @(@('tglPptLabels', 'labels'), @('tglPptTotals', 'totals'), @('tglPptConnectors', 'connectors'))) {
    Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptChart']/u:toggleButton[@id='$($t[0])'][@tag='$($t[1])'][@getPressed='PptRx_GetOpt'][@onAction='PptRx_OnToggleOpt'][@getEnabled='PptRx_GetOptEnabled']", $pns)) "ppt: alternância $($t[0])"
}
foreach ($t in @(@('ddPptNumFmt', 'numfmt', 'int,dec1,dec2,pct0,pct1,mil,mi'), @('ddPptLegend', 'legend', 'right,none'))) {
    $dd = $px.SelectSingleNode("//u:group[@id='grpPptChart']/u:dropDown[@id='$($t[0])'][@tag='$($t[1])'][@getSelectedItemID='PptRx_GetOptItem'][@onAction='PptRx_OnOptItem'][@getEnabled='PptRx_GetOptEnabled']", $pns)
    Check ($null -ne $dd) "ppt: lista $($t[0])"
    $want = @($t[2] -split ',' | ForEach-Object { $t[1] + $_ }) -join ','
    Check ($null -ne $dd -and (@($dd.SelectNodes('u:item', $pns) | ForEach-Object { $_.GetAttribute('id') }) -join ',') -ceq $want) "ppt: itens de $($t[0]) = tag + valor do core ($want)"
}
# mesmos rotulos dos itens do Excel (ids do Excel = valor do core)
foreach ($t in @(@('ddPptNumFmt', 'ddFormat', 'numfmt'), @('ddPptLegend', 'ddLegend', 'legend'))) {
    $ok = $true
    foreach ($it in @($px.SelectNodes("//u:dropDown[@id='$($t[0])']/u:item", $pns))) {
        $xi = $xx.SelectSingleNode("//u:dropDown[@id='$($t[1])']/u:item[@id='$($it.GetAttribute('id').Substring($t[2].Length))']", $xns)
        if ($null -eq $xi -or $xi.GetAttribute('label') -cne $it.GetAttribute('label')) { $ok = $false }
    }
    Check $ok "ppt: itens de $($t[0]) com os valores e rotulos do $($t[1]) do Excel"
}
Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptChart']/u:button[@id='btnPptCagr'][@tag='cagr'][@onAction='PptRx_OnArrow'][@getEnabled='PptRx_GetOptEnabled']", $pns)) 'ppt: Seta CAGR'
Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptChart']/u:button[@id='btnPptDiff'][@tag='diff'][@onAction='PptRx_OnArrow'][@getEnabled='PptRx_GetOptEnabled']", $pns)) 'ppt: Seta diferença'
Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptChart']/u:button[@id='btnPptClearAnn'][@onAction='PptRx_OnClearAnn'][@getEnabled='PptRx_GetElemEnabled']", $pns)) 'ppt: Remover setas'
Check (@($px.SelectNodes("//u:group[@id='grpPptChart']/u:*", $pns)).Count -eq 8) 'ppt: Gráfico do ThinQQ com 8 controles'
Check ($pptBas -match '(?m)^Public Sub PptRx_OnToggleOpt\(control As IRibbonControl, pressed As Boolean\)' -and $pptBas -match '(?m)^Public Sub PptRx_GetOpt\(control As IRibbonControl, ByRef \w+\)' -and $pptBas -match '(?m)^Public Sub PptRx_GetOptEnabled\(control As IRibbonControl, ByRef \w+\)') 'ppt: assinaturas das alternâncias do Gráfico do ThinQQ'
Check ($pptBas -match '(?m)^Public Sub PptRx_GetOptItem\(control As IRibbonControl, ByRef \w+\)' -and $pptBas -match '(?m)^Public Sub PptRx_OnOptItem\(control As IRibbonControl, selectedId As String, selectedIndex As Integer\)') 'ppt: assinaturas das listas do Gráfico do ThinQQ'
Check ($pptBas -match '(?m)^Public Sub PptRx_OnArrow\(control As IRibbonControl\)' -and $pptBas -match '(?m)^Public Sub PptRx_OnClearAnn\(control As IRibbonControl\)') 'ppt: assinaturas das setas do Gráfico do ThinQQ'
# 1.9.0 (G9): "Empacotar para enviar" no grupo "Gráficos do Excel" (imageMso: image= so no Inserir gráfico)
Check ($null -ne $px.SelectSingleNode("//u:group[@id='grpPptSync']/u:button[@id='btnPptPackage'][@label='Empacotar para enviar'][@onAction='PptRx_OnPackage'][@imageMso]", $pns)) 'ppt: Empacotar para enviar'
Check ($pptBas -match '(?m)^Public Sub PptRx_OnPackage\(control As IRibbonControl\)') 'ppt: assinatura de Empacotar para enviar'
foreach ($id in 'tglPptLabels', 'tglPptTotals', 'tglPptConnectors', 'ddPptNumFmt', 'ddPptLegend', 'btnPptCagr', 'btnPptDiff', 'btnPptClearAnn') {
    Check ($pptBas.Contains('"' + $id + '"')) "ppt: $id invalidado com a selecao"
}
# 1.9.0 (E3): "Texto vinculado ao Excel" no grupo "Elemento selecionado" (ultimo botao; imageMso ja usado; sem getEnabled: vale para caixa de texto)
Check ($px.SelectSingleNode("//u:group[@id='grpPptElem']", $pns).GetAttribute('label') -ceq 'Elemento selecionado') 'ppt: grupo Elemento selecionado (1.9.0)'
$lt = $px.SelectSingleNode("//u:group[@id='grpPptElem']/u:button[@id='btnPptLinkText'][@label='Texto vinculado ao Excel'][@onAction='PptRx_OnLinkText'][@imageMso='TextBoxInsert']", $pns)
Check ($null -ne $lt) 'ppt: Texto vinculado ao Excel'
Check ($null -ne $lt -and -not $lt.HasAttribute('getEnabled')) 'ppt: Texto vinculado ao Excel sempre ativo'
Check ($null -ne $lt -and $null -eq $lt.NextSibling) 'ppt: Texto vinculado ao Excel e o ultimo do grupo'
Check (@($px.SelectNodes("//u:group[@id='grpPptElem']/u:*", $pns)).Count -eq 4) 'ppt: Elemento selecionado com 4 controles'
Check ($pptBas -match '(?m)^Public Sub PptRx_OnLinkText\(control As IRibbonControl\)') 'ppt: assinatura de Texto vinculado ao Excel'
# e no Excel: Vincular celula ao texto selecionado (menu PowerPoint) e Titulo ligado a uma celula (Mais opcoes)
Check ($null -ne $xx.SelectSingleNode("//u:menu[@id='mnuPpt']/u:button[@id='btnPptLinkText'][@onAction='Rx_OnPptLinkText'][@image='tqq_ppt16']", $xns)) 'excel: Vincular célula ao texto selecionado'
Check ($null -ne $xx.SelectSingleNode("//u:menu[@id='mnuMore']/u:button[@id='btnLinkTitle'][@onAction='Rx_OnLinkTitle'][@getEnabled='Rx_GetEnabled'][@image='tqq_options16']", $xns)) 'excel: Título ligado a uma célula'
# ---- Carimbo do .ppam (Set-PackageDescription): o PowerPoint grava o suplemento sem docProps; pacotes sinteticos, sem Office
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
$tmpD = Join-Path ([IO.Path]::GetTempPath()) ("tqq_pkgd_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $tmpD | Out-Null
try {
    function New-SynPkg([string]$path, [hashtable]$parts) {
        $z = [System.IO.Compression.ZipFile]::Open($path, [System.IO.Compression.ZipArchiveMode]::Create)
        try { foreach ($k in $parts.Keys) { $e = $z.CreateEntry($k); $w = New-Object IO.StreamWriter($e.Open(), (New-Object Text.UTF8Encoding($false))); $w.Write($parts[$k]); $w.Dispose() } } finally { $z.Dispose() }
    }
    function Read-PkgText([string]$path, [string]$name) { $z = [System.IO.Compression.ZipFile]::OpenRead($path); try { $b = Read-Entry $z $name; if ($null -eq $b) { return $null }; return [Text.Encoding]::UTF8.GetString($b) } finally { $z.Dispose() } }
    $ctBare = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>'
    $relBare = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>'
    $a = Join-Path $tmpD 'a.ppam'
    New-SynPkg $a @{ '[Content_Types].xml' = $ctBare; '_rels/.rels' = $relBare; 'ppt/presentation.xml' = '' }
    Set-PackageDescription -PackagePath $a -Description 'ThinQQ ppt 9.8.7 (Release)'
    Set-PackageDescription -PackagePath $a -Description 'ThinQQ ppt 9.8.7 (Release)'   # repetir nao duplica
    $core = Read-PkgText $a 'docProps/core.xml'
    Check ($null -ne $core -and ([xml]$core).coreProperties.description -eq 'ThinQQ ppt 9.8.7 (Release)') 'carimbo: pacote sem docProps ganha core.xml com dc:description'
    Check ($core -match 'ThinQQ ppt (\d+\.\d+\.\d+)' -and $Matches[1] -eq '9.8.7') 'carimbo: versao legivel pela mesma regex do instalador'
    $ctA = Read-PkgText $a '[Content_Types].xml'; $relA = Read-PkgText $a '_rels/.rels'
    Check (([regex]::Matches($ctA, 'PartName="/docProps/core\.xml" ContentType="application/vnd\.openxmlformats-package\.core-properties\+xml"')).Count -eq 1) 'carimbo: content type de core.xml (uma vez)'
    Check (([regex]::Matches($relA, 'relationships/metadata/core-properties" Target="docProps/core\.xml"')).Count -eq 1) 'carimbo: relacao core-properties (uma vez)'
    Check ($null -ne ([xml]$ctA) -and $null -ne ([xml]$relA)) 'carimbo: [Content_Types].xml e _rels/.rels continuam XML valido'
    $b = Join-Path $tmpD 'b.pptm'
    New-SynPkg $b @{ '[Content_Types].xml' = $ctBare; '_rels/.rels' = $relBare; 'docProps/core.xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>X</dc:title><dc:description>velho</dc:description></cp:coreProperties>' }
    Set-PackageDescription -PackagePath $b -Description 'ThinQQ ppt 1.2.3 (Test)'
    $coreB = Read-PkgText $b 'docProps/core.xml'
    Check (([xml]$coreB).coreProperties.description -eq 'ThinQQ ppt 1.2.3 (Test)' -and ([xml]$coreB).coreProperties.title -eq 'X') 'carimbo: core.xml existente troca so o dc:description'
    Check ((Read-PkgText $b '[Content_Types].xml') -eq $ctBare) 'carimbo: core.xml existente nao mexe no content type'
} finally { Remove-Item -LiteralPath $tmpD -Recurse -Force -ErrorAction SilentlyContinue }
Check (-not (Test-Path -LiteralPath $tmpD)) 'pasta tqq_pkgd_* do teste removida'
# ---- Close-PptApp com duble (sem PowerPoint): fecha as do script e as sem janela; com apresentacao de janela ou Modo Protegido nao da Quit
function New-FakePres([int]$w) { $o = [pscustomobject]@{ Closed = $false; Saved = 0; Windows = [pscustomobject]@{ Count = $w } }; $o | Add-Member -MemberType ScriptMethod -Name Close -Value { $this.Closed = $true }; return $o }
function New-FakePpt([object[]]$pres, [int]$pv = 0) {
    $col = [pscustomobject]@{ L = [System.Collections.ArrayList]::new(@($pres)) }
    $col | Add-Member -MemberType ScriptProperty -Name Count -Value { $this.L.Count }
    $col | Add-Member -MemberType ScriptMethod -Name Item -Value { param($i) $this.L[$i - 1] }
    $a = [pscustomobject]@{ Presentations = $col; ProtectedViewWindows = [pscustomobject]@{ Count = $pv }; QuitCalled = $false }
    $a | Add-Member -MemberType ScriptMethod -Name Quit -Value { $this.QuitCalled = $true }
    return $a
}
$script:PptPid = @()
$p1 = New-FakePres 0; $p2 = New-FakePres 1; $fa = New-FakePpt @($p1, $p2)
Close-PptApp $fa -Owned @($p1) -WaitSec 1 6>$null
Check ($p1.Closed -and -not $p2.Closed -and -not $fa.QuitCalled) 'Close-PptApp: apresentacao com janela (do usuario) fica aberta e o PowerPoint nao recebe Quit'
$p3 = New-FakePres 0; $fb = New-FakePpt @($p3)
Close-PptApp $fb -WaitSec 1
Check ($p3.Closed -and $fb.QuitCalled) 'Close-PptApp: sobra sem janela (do codigo) e fechada e o PowerPoint recebe Quit'
$fc = New-FakePpt @() 1
Close-PptApp $fc -WaitSec 1 6>$null
Check (-not $fc.QuitCalled) 'Close-PptApp: janela do Modo de Exibicao Protegido aberta: sem Quit'
# ---- run-tests: POWERPNT antes/depois (so avisa: nunca encerra um PowerPoint que o script nao criou) e testes de PowerPoint pulados
$snapA = @([pscustomobject]@{ Id = 10; ParentId = 1; CommandLine = 'POWERPNT.EXE' })
$snapB = @($snapA[0], [pscustomobject]@{ Id = 20; ParentId = 900; CommandLine = '"POWERPNT.EXE" /AUTOMATION -Embedding' }, [pscustomobject]@{ Id = 30; ParentId = 5; CommandLine = '"POWERPNT.EXE" "C:\x\a.pptx"' })
$left = @(Get-PptLeftovers -Before $snapA -After $snapB)
Check ($left.Count -eq 2 -and ($left | ForEach-Object Id) -join ',' -eq '20,30') 'run-tests: sobra = POWERPNT que nao existia antes (o do usuario aberto antes nao conta)'
Check (@(Get-PptLeftovers -Before @() -After @()).Count -eq 0) 'run-tests: sem PowerPoint antes nem depois = nenhuma sobra'
Check ((Test-PptAutomationProcess $snapB[1]) -and -not (Test-PptAutomationProcess $snapB[2]) -and -not (Test-PptAutomationProcess $snapA[0])) 'run-tests: so -Embedding e PowerPoint de automacao (orfao de teste); o aberto pelo usuario nao'
$msgL = Format-PptLeftover $snapB[1]
Check ($msgL -match 'PID 20' -and $msgL -match 'pai 900' -and $msgL -match 'Embedding') "run-tests: aviso da sobra traz PID, pai e argumentos ($msgL)"
$logLines = @("2026-10-14 10:00:00`tcore`tTestPowerPoint: pulado (ha um PowerPoint aberto)", "2026-10-14 10:00:01`tcore`tTestPptSync: pulado (processo POWERPNT existente ou contagem indisponivel: 1)",
    "2026-10-14 10:00:02`tcore`tTestPptSync: colagem pulada (area de transferencia indisponivel)", "2026-10-14 10:00:03`tcore`tTestPowerPoint: resto do teste pulado (uma janela apareceu)", "2026-10-14 10:00:04`tcore`tCore_Insert`t5`toutro erro")
Check ((Get-PptSkipCount $logLines) -eq 4) 'run-tests: conta os testes de PowerPoint pulados pelas linhas do log de teste'
Check ((Get-PptSkipCount @()) -eq 0) 'run-tests: log vazio = 0 pulados'
$tmpLog = Join-Path ([IO.Path]::GetTempPath()) ('tqq_log_' + [guid]::NewGuid().ToString('N') + '.log')
try {
    Check ((Get-FileLengthOrZero $tmpLog) -eq 0) 'run-tests: log ausente = tamanho 0'
    [IO.File]::WriteAllText($tmpLog, "antigo`r`n", [Text.Encoding]::ASCII); $off = Get-FileLengthOrZero $tmpLog
    [IO.File]::AppendAllText($tmpLog, "novo1`r`nnovo2`r`n", [Text.Encoding]::ASCII)
    $tail = @(Read-LogTail $tmpLog $off)
    Check ($tail.Count -eq 2 -and $tail[0] -eq 'novo1' -and $tail[1] -eq 'novo2') "run-tests: le so as linhas do log gravadas durante a execucao ($($tail -join '|'))"
    Check (@(Read-LogTail (Join-Path ([IO.Path]::GetTempPath()) 'tqq_nao_existe.log') 0).Count -eq 0) 'run-tests: log ausente = nenhuma linha'
} finally { Remove-Item -LiteralPath $tmpLog -Force -ErrorAction SilentlyContinue }
$rt = [IO.File]::ReadAllText((Join-Path $RepoRoot 'build\run-tests.ps1'))
Check ($rt -match '\[int\]\$TimeoutSec = 1200') 'run-tests: -TimeoutSec padrao 1200 (a suite completa passa de 800 s)'
Check (([IO.File]::ReadAllText((Join-Path $RepoRoot 'tools\lib\Release-Publish.ps1'))) -match 'run-tests\.ps1''\) -Target all -LoaderSwap -TimeoutSec 1200') 'publicar-release -RunTests: suite com -TimeoutSec 1200'
Check ($rt -match 'Get-PptProcessSnapshot' -and $rt -match 'Get-PptLeftovers' -and $rt -match 'Get-PptSkipCount') 'run-tests: registra POWERPNT antes/depois e conta os testes de PowerPoint pulados'
Check ($rt -notmatch 'Stop-Process[^\r\n]*(left|Leftover|POWERPNT)') 'run-tests: nunca encerra POWERPNT (so avisa)'
$fin = [regex]::Match($rt, '(?s)\} finally \{\s*\r?\n\s*Stop-ExcelWatchdog \$wd.*?\r?\n\}').Value
Check ($fin -match 'Test-PptAfterSuite') 'run-tests: o aviso de PowerPoint orfao sai tambem quando o bloco VBA lanca (0x800706BE): Test-PptAfterSuite no finally'
if ($script:fail -gt 0) { Write-Host "FALHAS ribbon-pacote: $script:fail"; exit 1 }
