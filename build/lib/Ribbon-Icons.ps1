# Icones proprios da faixa: gravador de PNG puro (RGBA 8 bits; zlib com blocos "stored"; CRC-32 e Adler-32 proprios: bytes identicos em
# qualquer runtime, sem System.Drawing) e a tabela de icones (retangulos numa grade de 16 x 16, com o imageMso de reserva de cada um).
Set-StrictMode -Version Latest

$script:TqqNavy = @(31, 56, 100); $script:TqqBlue = @(68, 114, 196); $script:TqqLight = @(157, 195, 230); $script:TqqGray = @(127, 127, 127)
# nunca "R": o alias embutido r (Invoke-History) tem precedencia sobre funcoes e a chamada falharia
function New-TqqRect([double]$x, [double]$y, [double]$w, [double]$h, $c) { return @{ x = $x; y = $y; w = $w; h = $h; c = $c } }

function Get-TqqIconSpecs {
    $n = $script:TqqNavy; $b = $script:TqqBlue; $l = $script:TqqLight; $g = $script:TqqGray
    return [ordered]@{
        tqq_columns   = @{ Mso = 'ChartTypeColumnInsertGallery'; Rects = @((New-TqqRect 1 8 4 7 $n), (New-TqqRect 6 3 4 12 $b), (New-TqqRect 11 6 4 9 $l)) }
        tqq_bars      = @{ Mso = 'ChartTypeBarInsertGallery';    Rects = @((New-TqqRect 1 1 13 4 $n), (New-TqqRect 1 6 8 4 $b), (New-TqqRect 1 11 5 4 $l)) }
        tqq_waterfall = @{ Mso = 'ChartTypeOtherInsertGallery';  Rects = @((New-TqqRect 1 7 3 8 $n), (New-TqqRect 5 3 3 4 $b), (New-TqqRect 9 3 3 3 $l), (New-TqqRect 13 3 2 12 $n)) }
        tqq_lines     = @{ Mso = 'ChartTypeLineInsertGallery';   Rects = @((New-TqqRect 1 12 3 2 $b), (New-TqqRect 4 9 3 2 $b), (New-TqqRect 7 10 3 2 $b), (New-TqqRect 10 5 3 2 $b), (New-TqqRect 13 2 2 2 $b), (New-TqqRect 1 14 15 1 $g)) }
        tqq_more      = @{ Mso = 'ChartTypePieInsertGallery';    Rects = @((New-TqqRect 1 1 6 6 $n), (New-TqqRect 9 1 6 6 $b), (New-TqqRect 1 9 6 6 $l), (New-TqqRect 9 9 6 6 $g)) }
        tqq_butterfly = @{ Mso = 'ChartTypeBarInsertGallery';    Rects = @((New-TqqRect 1 2 5 3 $n), (New-TqqRect 3 6 3 3 $n), (New-TqqRect 0 10 6 3 $n), (New-TqqRect 10 2 5 3 $b), (New-TqqRect 10 6 4 3 $b), (New-TqqRect 10 10 3 3 $b)) }
        tqq_scatter   = @{ Mso = 'ChartTypeXYScatterInsertGallery'; Rects = @((New-TqqRect 2 10 3 3 $b), (New-TqqRect 6 7 3 3 $b), (New-TqqRect 9 11 3 3 $n), (New-TqqRect 11 3 3 3 $b), (New-TqqRect 5 2 3 3 $n), (New-TqqRect 0 0 1 16 $g), (New-TqqRect 0 15 16 1 $g)) }
        tqq_bubble    = @{ Mso = 'ChartTypeXYScatterInsertGallery'; Rects = @((New-TqqRect 1 8 7 7 $b), (New-TqqRect 9 2 6 6 $n), (New-TqqRect 7 1 3 3 $l), (New-TqqRect 0 0 1 16 $g), (New-TqqRect 0 15 16 1 $g)) }
        tqq_options   = @{ Mso = 'PropertySheet';                Rects = @((New-TqqRect 1 3 14 1 $g), (New-TqqRect 1 8 14 1 $g), (New-TqqRect 1 13 14 1 $g), (New-TqqRect 9 2 3 3 $n), (New-TqqRect 4 7 3 3 $b), (New-TqqRect 11 12 3 3 $n)) }
        tqq_arrows    = @{ Mso = 'ArrowsMore';                   Rects = @((New-TqqRect 1 7 10 2 $n), (New-TqqRect 9 4 2 8 $n), (New-TqqRect 11 5 2 6 $n), (New-TqqRect 13 6 2 4 $n)) }
        tqq_ppt       = @{ Mso = 'SlideNew';                     Rects = @((New-TqqRect 1 2 14 11 $n), (New-TqqRect 3 4 10 7 $l), (New-TqqRect 4 7 2 3 $b), (New-TqqRect 7 5 2 5 $b), (New-TqqRect 10 6 2 4 $b), (New-TqqRect 6 13 4 2 $g)) }
        tqq_round     = @{ Mso = 'DecimalsDecrease';             Rects = @((New-TqqRect 1 2 14 4 $l), (New-TqqRect 1 7 14 4 $b), (New-TqqRect 1 12 14 3 $n), (New-TqqRect 14 1 2 14 $g)) }
        tqq_break     = @{ Mso = 'ChartAxes';                    Rects = @((New-TqqRect 5 9 6 6 $b), (New-TqqRect 5 1 6 5 $b), (New-TqqRect 3 6 10 1 $g), (New-TqqRect 3 8 10 1 $g)) }
    }
}

# tabela do CRC-32 montada num bloco de script: $poly, $i, $c e $k nao vazam para os scripts que fazem dot-source do Build-Common
$script:TqqCrcTable = & {
    $t = New-Object uint64[] 256
    $poly = [uint64]3988292384
    for ($i = 0; $i -lt 256; $i++) { $c = [uint64]$i; for ($k = 0; $k -lt 8; $k++) { if (($c -band 1) -eq 1) { $c = $poly -bxor ($c -shr 1) } else { $c = $c -shr 1 } }; $t[$i] = $c }
    , $t
}
function Get-TqqCrc32([byte[]]$Bytes) {
    $c = [uint64]4294967295
    foreach ($b in $Bytes) { $c = $script:TqqCrcTable[[int](($c -bxor [uint64]$b) -band 255)] -bxor ($c -shr 8) }
    return [uint32]($c -bxor [uint64]4294967295)
}
function Get-TqqBE32([uint32]$v) { return [byte[]]@([byte](($v -shr 24) -band 255), [byte](($v -shr 16) -band 255), [byte](($v -shr 8) -band 255), [byte]($v -band 255)) }
function New-TqqPngChunk([string]$Type, [byte[]]$Data) {
    $t = [Text.Encoding]::ASCII.GetBytes($Type)
    return [byte[]]((Get-TqqBE32 ([uint32]$Data.Length)) + $t + $Data + (Get-TqqBE32 (Get-TqqCrc32 ([byte[]]($t + $Data)))))
}
function Get-TqqZlibStored([byte[]]$Raw) {
    $out = New-Object System.Collections.Generic.List[byte]
    $out.Add([byte]0x78); $out.Add([byte]0x01)
    $pos = 0
    do {
        $len = [Math]::Min(65535, $Raw.Length - $pos); $final = if ($pos + $len -ge $Raw.Length) { 1 } else { 0 }; $nlen = 65535 - $len
        $out.Add([byte]$final); $out.Add([byte]($len -band 255)); $out.Add([byte](($len -shr 8) -band 255)); $out.Add([byte]($nlen -band 255)); $out.Add([byte](($nlen -shr 8) -band 255))
        for ($i = 0; $i -lt $len; $i++) { $out.Add($Raw[$pos + $i]) }
        $pos += $len
    } while ($pos -lt $Raw.Length)
    [uint64]$a = 1; [uint64]$b = 0
    foreach ($x in $Raw) { $a = ($a + [uint64]$x) % 65521; $b = ($b + $a) % 65521 }
    foreach ($y in (Get-TqqBE32 ([uint32](($b -shl 16) -bor $a)))) { $out.Add([byte]$y) }
    return $out.ToArray()
}
function Write-TqqPng([string]$Path, [int]$Width, [int]$Height, [byte[]]$Rgba) {
    $raw = New-Object byte[] (($Width * 4 + 1) * $Height)
    for ($y = 0; $y -lt $Height; $y++) { $raw[$y * ($Width * 4 + 1)] = 0; [Array]::Copy($Rgba, $y * $Width * 4, $raw, $y * ($Width * 4 + 1) + 1, $Width * 4) }
    $ihdr = [byte[]]((Get-TqqBE32 ([uint32]$Width)) + (Get-TqqBE32 ([uint32]$Height)) + [byte[]]@(8, 6, 0, 0, 0))
    $png = [byte[]]([byte[]]@(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A) + (New-TqqPngChunk 'IHDR' $ihdr) + (New-TqqPngChunk 'IDAT' (Get-TqqZlibStored $raw)) + (New-TqqPngChunk 'IEND' ([byte[]]@())))
    [IO.File]::WriteAllBytes($Path, $png)
}
function New-TqqIconRgba([int]$Size, $Rects) {
    $px = New-Object byte[] ($Size * $Size * 4)
    foreach ($r in $Rects) {
        $x0 = [int][Math]::Floor($r.x * $Size / 16); $y0 = [int][Math]::Floor($r.y * $Size / 16)
        $x1 = [Math]::Min($Size, [int][Math]::Ceiling(($r.x + $r.w) * $Size / 16)); $y1 = [Math]::Min($Size, [int][Math]::Ceiling(($r.y + $r.h) * $Size / 16))
        for ($y = $y0; $y -lt $y1; $y++) { for ($x = $x0; $x -lt $x1; $x++) { $o = ($y * $Size + $x) * 4; $px[$o] = [byte]$r.c[0]; $px[$o + 1] = [byte]$r.c[1]; $px[$o + 2] = [byte]$r.c[2]; $px[$o + 3] = 255 } }
    }
    return $px
}
# Grava <nome>.png (32) e <nome>16.png (16) de cada icone em OutDir; devolve os caminhos (ordem fixa).
function Export-TqqIcons([string]$OutDir) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($e in (Get-TqqIconSpecs).GetEnumerator()) {
        $p32 = Join-Path $OutDir "$($e.Key).png"; Write-TqqPng $p32 32 32 (New-TqqIconRgba 32 $e.Value.Rects); $out.Add($p32)
        $p16 = Join-Path $OutDir "$($e.Key)16.png"; Write-TqqPng $p16 16 16 (New-TqqIconRgba 16 $e.Value.Rects); $out.Add($p16)
    }
    return $out.ToArray()
}
# Fallback: troca image="tqq_x" / image="tqq_x16" por imageMso="<reserva>" (so no texto que vai para o pacote; o fonte nao muda). Comentarios
# XML passam intactos (um image= citado num comentario nao e atributo de controle).
function Convert-TqqRibbonToMso([string]$Xml) {
    $specs = Get-TqqIconSpecs
    return [regex]::Replace($Xml, '(?s)(<!--.*?-->)|(?<=\s)image="(tqq_[a-z0-9_]+?)(16)?"', {
        param($m)
        if ($m.Groups[1].Success) { return $m.Value }
        $k = $m.Groups[2].Value
        if ($specs.Contains($k)) { return 'imageMso="' + $specs[$k].Mso + '"' } else { throw "icone sem reserva: $k" }
    })
}

# Valida o customUI14 que vai para o pacote e devolve os nomes image= distintos (ordenados; [string[]] vazio se nao houver). Le pelo DOM, entao
# comentarios nao contam. Recusa (throw, com o id do controle): XML malformado (inclusive atributo repetido, o que a conversao Mso produziria
# num controle com image e imageMso), controle com image e imageMso, nome fora da tabela de icones (Png) e qualquer image= que sobrou (Mso).
function Assert-TqqRibbonXml([string]$Xml, [ValidateSet('Png','Mso')][string]$Mode, [string]$What = 'customUI14.xml') {
    $doc = New-Object Xml.XmlDocument
    try { $doc.LoadXml($Xml) } catch { throw "$What`: XML invalido (modo $Mode): $($_.Exception.Message)" }
    $specs = Get-TqqIconSpecs
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($n in @($doc.SelectNodes('//*[@image]'))) {
        $img = $n.GetAttribute('image'); $id = $n.GetAttribute('id')
        if (-not $id) { $id = "<$($n.LocalName) sem id>" }
        if ($n.HasAttribute('imageMso')) { throw "$What`: o controle '$id' tem image e imageMso (use so um dos dois)" }
        if ($Mode -eq 'Mso') { throw "$What`: image=`"$img`" sobrou no controle '$id' (nome fora da tabela de icones?)" }
        if (-not $specs.Contains(($img -replace '16$', ''))) { throw "$What usa image=`"$img`" sem icone na tabela (Ribbon-Icons.ps1)" }
        if (-not $names.Contains($img)) { $names.Add($img) }
    }
    $out = [string[]]@($names | Sort-Object)
    return , $out
}
