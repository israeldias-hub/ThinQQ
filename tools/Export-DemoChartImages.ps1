<#
.SYNOPSIS  Exporta os gráficos da pasta de demonstração (exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx) para PNG, na ordem E01..E26
           (E25 tem dois gráficos: E25a e E25b), e reduz cada PNG. As imagens alimentam a galeria do tutorial (tools\Build-TutorialGallery.ps1).
.EXAMPLE   pwsh tools\Export-DemoChartImages.ps1 -OutDir C:\temp\tqq-galeria-img
           pwsh tools\Build-TutorialGallery.ps1 -ImageDir C:\temp\tqq-galeria-img
.NOTES     Abre uma CÓPIA do .xlsx (somente leitura, numa subpasta de -OutDir apagada no fim) num Excel próprio, invisível e SEM o ThinQQ: a pasta
           abre sem o suplemento (é o que o usuário sem o ThinQQ vê). O Excel de outra pessoa nunca é tocado (New-ExcelApp recusa um processo já
           existente; o watchdog mata só o Excel iniciado aqui). Cada gráfico é exportado por Chart.Export (PNG, tamanho natural: 96 dpi, ou seja
           4/3 de pixel por ponto), depois de ativar a planilha e o gráfico e esperar o layout assentar, como na verificação visual do gerador do demo.
           Redução de cada PNG: (1) descarta os blocos auxiliares (texto, tempo, gama, dimensões físicas), ficando só IHDR, PLTE, tRNS, IDAT e IEND;
           (2) se a imagem não é indexada mas tem no máximo -MaxColors cores distintas, converte para paleta de 8 bits SEM perda
           (System.Windows.Media.Imaging: FormatConvertedBitmap + BitmapPalette) e fica com o MENOR dos dois; com mais cores a imagem fica como está
           (a quantização aproximada deixou manchas numa imagem de bolhas). O Excel já exporta quase tudo indexado, então na prática basta o passo (1).
           Funciona no Windows PowerShell 5.1 e no pwsh 7. Sai com 0 (imagens gravadas em -OutDir) ou 1 (falha).
#>
param(
    [string]$Workbook = (Join-Path $PSScriptRoot '..\exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx'),
    [string]$OutDir = (Join-Path ([IO.Path]::GetTempPath()) 'tqq-demo-imagens'),
    [ValidateRange(2, 256)][int]$MaxColors = 256,
    [int]$TimeoutSec = 300
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '..\build\lib\Build-Common.ps1')
. (Join-Path $PSScriptRoot 'DemoCoverage.ps1')
# (o WPF é carregado só na redução, depois que o Excel fecha)

# Lê o PNG: dimensões, profundidade, tipo de cor e a lista de blocos (nome, deslocamento, tamanho total com comprimento e CRC).
function Get-PngChunks([byte[]]$Bytes) {
    $sig = [byte[]](137, 80, 78, 71, 13, 10, 26, 10)
    for ($i = 0; $i -lt 8; $i++) { if ($Bytes[$i] -ne $sig[$i]) { throw 'não é um PNG (assinatura inválida)' } }
    $chunks = New-Object System.Collections.Generic.List[object]
    $pos = 8
    while ($pos + 12 -le $Bytes.Length) {
        $len = ([int]$Bytes[$pos] -shl 24) -bor ([int]$Bytes[$pos + 1] -shl 16) -bor ([int]$Bytes[$pos + 2] -shl 8) -bor [int]$Bytes[$pos + 3]
        $name = [Text.Encoding]::ASCII.GetString($Bytes, $pos + 4, 4)
        $chunks.Add([pscustomobject]@{ Name = $name; Offset = $pos; Size = $len + 12 })
        $pos += $len + 12
        if ($name -eq 'IEND') { break }
    }
    $ihdr = $chunks[0]
    if ($ihdr.Name -ne 'IHDR') { throw 'PNG sem IHDR no começo' }
    $o = $ihdr.Offset + 8
    $w = ([int]$Bytes[$o] -shl 24) -bor ([int]$Bytes[$o + 1] -shl 16) -bor ([int]$Bytes[$o + 2] -shl 8) -bor [int]$Bytes[$o + 3]
    $h = ([int]$Bytes[$o + 4] -shl 24) -bor ([int]$Bytes[$o + 5] -shl 16) -bor ([int]$Bytes[$o + 6] -shl 8) -bor [int]$Bytes[$o + 7]
    return [pscustomobject]@{ Width = $w; Height = $h; BitDepth = [int]$Bytes[$o + 8]; ColorType = [int]$Bytes[$o + 9]; Chunks = $chunks.ToArray() }
}

# Cópia do PNG só com os blocos que desenham a imagem (IHDR, PLTE, tRNS, IDAT, IEND).
function Remove-PngMetadata([byte[]]$Bytes) {
    $info = Get-PngChunks $Bytes
    $ms = New-Object IO.MemoryStream
    $ms.Write($Bytes, 0, 8)
    foreach ($c in $info.Chunks) { if ($c.Name -in 'IHDR', 'PLTE', 'tRNS', 'IDAT', 'IEND') { $ms.Write($Bytes, $c.Offset, $c.Size) } }
    return , $ms.ToArray()
}

# Quantização para paleta de 8 bits, SÓ quando não perde nada (a imagem tem no máximo $Colors cores distintas); senão devolve $null: a paleta
# aproximada do WPF deixou manchas e linhas fantasma numa imagem de bolhas, e a economia não compensa. Devolve os bytes do PNG indexado.
function ConvertTo-IndexedPng([byte[]]$Bytes, [int]$Colors) {
    Add-Type -AssemblyName PresentationCore, WindowsBase
    $in = New-Object IO.MemoryStream(, $Bytes)
    $dec = New-Object Windows.Media.Imaging.PngBitmapDecoder($in, [Windows.Media.Imaging.BitmapCreateOptions]::PreservePixelFormat, [Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
    $bgra = New-Object Windows.Media.Imaging.FormatConvertedBitmap($dec.Frames[0], [Windows.Media.PixelFormats]::Bgra32, $null, 0)
    $px = New-Object byte[] ($bgra.PixelWidth * $bgra.PixelHeight * 4)
    $bgra.CopyPixels($px, $bgra.PixelWidth * 4, 0)
    $seen = New-Object 'System.Collections.Generic.HashSet[uint32]'
    for ($i = 0; $i -lt $px.Length; $i += 4) { [void]$seen.Add([BitConverter]::ToUInt32($px, $i)); if ($seen.Count -gt $Colors) { return $null } }
    $pal = New-Object Windows.Media.Imaging.BitmapPalette($bgra, $Colors)
    $idx = New-Object Windows.Media.Imaging.FormatConvertedBitmap($bgra, [Windows.Media.PixelFormats]::Indexed8, $pal, 0)
    $enc = New-Object Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($idx))
    $out = New-Object IO.MemoryStream
    $enc.Save($out)
    return , (Remove-PngMetadata $out.ToArray())
}

function Reduce-DemoPng([string]$Path, [int]$Colors) {
    $orig = [IO.File]::ReadAllBytes($Path)
    $best = Remove-PngMetadata $orig
    $how = 'sem metadados'
    $info = Get-PngChunks $best
    if ($info.ColorType -ne 3) {                      # truecolor: tenta a paleta exata de 8 bits e fica com o menor
        $q = ConvertTo-IndexedPng $best $Colors
        if ($null -ne $q -and $q.Length -lt $best.Length) { $best = $q; $how = "paleta exata" }
    }
    [IO.File]::WriteAllBytes($Path, $best)
    return [pscustomobject]@{ Before = $orig.Length; After = $best.Length; How = $how }
}

if (-not (Test-Path -LiteralPath $Workbook)) { Write-Host "pasta de demonstração não encontrada: $Workbook" -ForegroundColor Red; exit 1 }
$Workbook = (Resolve-Path -LiteralPath $Workbook).Path
$OutDir = [IO.Path]::GetFullPath($OutDir)
$work = Join-Path $OutDir '_trabalho'
New-Item -ItemType Directory -Force $work | Out-Null
Get-ChildItem -LiteralPath $OutDir -Filter 'E*.png' -File -ErrorAction SilentlyContinue | Remove-Item -Force      # só os PNG que este script gera
$copy = Join-Path $work 'demo.xlsx'
Copy-Item -LiteralPath $Workbook -Destination $copy -Force

$cov = @(Get-TqqDemoCoverage)
$xl = $null; $wd = $null; $code = 0; $names = @()
try {
    $xl = New-ExcelApp
    $wd = Start-ExcelWatchdog $xl $TimeoutSec
    $wb = $xl.Workbooks.Open($copy, 0, $true)
    foreach ($it in $cov) {
        $ws = $wb.Worksheets.Item($it.Sheet)
        $ws.Activate()
        $objs = @(); foreach ($co in $ws.ChartObjects()) { $objs += $co }
        if ($objs.Count -lt 1) { throw "$($it.Sheet) não tem gráfico" }
        $objs = @($objs | Sort-Object { [double]$_.Left }, { [double]$_.Top })        # E25: o da esquerda é o "a"
        for ($k = 0; $k -lt $objs.Count; $k++) {
            $name = $it.Sheet; if ($objs.Count -gt 1) { $name += [string][char](97 + $k) }
            $file = Join-Path $OutDir "$name.png"
            $objs[$k].Activate()
            Start-Sleep -Milliseconds 400                                              # o layout do gráfico assenta depois de ativá-lo
            [void]$xl.ActiveChart.Export([string]$file, 'PNG')      # [string]: o texto do Join-Path (PSObject) faz o Chart.Export travar no Windows PowerShell 5.1
            if (-not (Test-Path -LiteralPath $file)) { throw "Chart.Export não gravou $file" }
            $names += $name
        }
        [void]$ws.Range('A1').Select()
    }
    $wb.Close($false)
    Write-Host "$($names.Count) gráficos exportados"
} catch {
    Write-Host "falha: $($_.Exception.Message)" -ForegroundColor Red
    $code = 1
} finally {
    Stop-ExcelWatchdog $wd
    Close-ExcelApp $xl
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue      # só a cópia de trabalho criada acima
}
if ($code -ne 0) { exit $code }

# Redução (com o Excel já fechado)
$total = 0
foreach ($name in $names) {
    $file = Join-Path $OutDir "$name.png"
    $r = Reduce-DemoPng $file $MaxColors
    $inf = Get-PngChunks ([IO.File]::ReadAllBytes($file))
    Write-Host ('{0,-5} {1,4}x{2,-4} {3,7:N0} -> {4,7:N0} bytes ({5})' -f $name, $inf.Width, $inf.Height, $r.Before, $r.After, $r.How)
    $total += $r.After
}
Write-Host ('{0} imagens, {1:N0} bytes ({2:N1} KB, média {3:N1} KB) em {4}' -f $names.Count, $total, ($total / 1KB), ($total / 1KB / $names.Count), $OutDir)
exit 0
