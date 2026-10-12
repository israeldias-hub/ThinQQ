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
if ($script:fail -gt 0) { Write-Host "FALHAS ribbon-pacote: $script:fail"; exit 1 }
