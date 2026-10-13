<#
.SYNOPSIS  Monta a galeria dos exemplos da pasta de demonstração no tutorial (docs\tutorial\tutorial.html), com as imagens EMBUTIDAS (data URI base64):
           o tutorial é um arquivo único e autocontido (vai solto no ZIP e é baixado pelo atualizador), então nada de arquivo ou link externo.
.EXAMPLE   pwsh tools\Export-DemoChartImages.ps1 -OutDir C:\temp\tqq-galeria-img
           pwsh tools\Build-TutorialGallery.ps1 -ImageDir C:\temp\tqq-galeria-img
.NOTES     Entrada: os PNG de Export-DemoChartImages.ps1 (E01.png ... E26.png; E25 com E25a.png e E25b.png). Os títulos, tipos, menus e descrições
           vêm de tools\DemoCoverage.ps1 (Get-TqqDemoCoverage), nunca copiados à mão: se a cobertura mudar, rode de novo (o Test-Docs falha se a
           legenda de alguma figura não citar o título atual).
           Saída (idempotente: duas execuções dão o mesmo arquivo): o bloco entre <!-- galeria:inicio --> e <!-- galeria:fim --> é substituído;
           sem os marcadores, o bloco entra logo depois da tabela de exemplos da seção #demonstracao. Na tabela, o número do exemplo (E05) vira
           link para a figura (#ex-E05). Cada <figure id="ex-E05"> tem um <img> de largura/altura reais (lidas do IHDR do PNG) e alt em português.
           Sai com 0 (gravado) ou 1 (falha; o tutorial não é alterado).
#>
param(
    [string]$ImageDir = (Join-Path ([IO.Path]::GetTempPath()) 'tqq-demo-imagens'),      # o mesmo padrão de Export-DemoChartImages.ps1 (sem Mandatory: no 5.1 ele esvazia $PSScriptRoot nos padrões seguintes)
    [string]$Tutorial = (Join-Path $PSScriptRoot '..\docs\tutorial\tutorial.html')
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'DemoCoverage.ps1')

$begin = '<!-- galeria:inicio -->'
$end = '<!-- galeria:fim -->'

function ConvertTo-HtmlText([string]$s) { return [Security.SecurityElement]::Escape($s) }

# Largura e altura (IHDR) de um PNG; recusa o que não for PNG.
function Get-PngSize([byte[]]$Bytes) {
    $sig = [byte[]](137, 80, 78, 71, 13, 10, 26, 10)
    if ($Bytes.Length -lt 33) { throw 'PNG curto demais' }
    for ($i = 0; $i -lt 8; $i++) { if ($Bytes[$i] -ne $sig[$i]) { throw 'não é um PNG (assinatura inválida)' } }
    $w = ([int]$Bytes[16] -shl 24) -bor ([int]$Bytes[17] -shl 16) -bor ([int]$Bytes[18] -shl 8) -bor [int]$Bytes[19]
    $h = ([int]$Bytes[20] -shl 24) -bor ([int]$Bytes[21] -shl 16) -bor ([int]$Bytes[22] -shl 8) -bor [int]$Bytes[23]
    if ($w -le 0 -or $h -le 0) { throw 'PNG com dimensões inválidas' }
    return [pscustomobject]@{ Width = $w; Height = $h }
}

try {
    if (-not (Test-Path -LiteralPath $Tutorial)) { throw "tutorial não encontrado: $Tutorial" }
    if (-not (Test-Path -LiteralPath $ImageDir)) { throw "pasta de imagens não encontrada: $ImageDir (rode tools\Export-DemoChartImages.ps1)" }
    $Tutorial = (Resolve-Path -LiteralPath $Tutorial).Path
    $ImageDir = (Resolve-Path -LiteralPath $ImageDir).Path
    $utf8 = New-Object Text.UTF8Encoding($false)
    $html = [IO.File]::ReadAllText($Tutorial, $utf8)
    $nl = "`n"; if ($html.Contains("`r`n")) { $nl = "`r`n" }

    $cov = @(Get-TqqDemoCoverage)
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append($begin + $nl)
    [void]$sb.Append('  <h3 id="galeria">Galeria: como cada exemplo fica</h3>' + $nl)
    [void]$sb.Append("  <p>Os $($cov.Count) exemplos da pasta, tais como aparecem no Excel (dados fictícios), na ordem das planilhas <strong>E01</strong> a <strong>E$('{0:00}' -f $cov.Count)</strong>. Cada legenda diz o que observar e qual caminho da aba <span class=""ui"">ThinQQ</span> cria o gráfico. São imagens fixas: no Excel os gráficos acompanham os números.</p>" + $nl)
    [void]$sb.Append('  <div class="gal">' + $nl)
    $imgCount = 0
    foreach ($it in $cov) {
        $files = @(Get-ChildItem -LiteralPath $ImageDir -File | Where-Object { $_.Name -match "^$($it.Sheet)[a-z]?\.png$" } | Sort-Object Name)
        if ($files.Count -lt 1) { throw "sem imagem para $($it.Sheet) em $ImageDir" }
        $sheet = $it.Sheet
        [void]$sb.Append("    <figure class=""gx"" id=""ex-$sheet"">" + $nl)
        [void]$sb.Append('      <div class="gx-img">' + $nl)
        $k = 0
        foreach ($f in $files) {
            $bytes = [IO.File]::ReadAllBytes($f.FullName)
            $size = Get-PngSize $bytes
            $alt = "Gráfico do exemplo $sheet"
            if ($files.Count -gt 1) { $alt += " ($($k + 1) de $($files.Count))" }
            $alt += ": $($it.Title). Tipo: $($it.KindName)."
            [void]$sb.Append("        <img src=""data:image/png;base64,$([Convert]::ToBase64String($bytes))"" width=""$($size.Width)"" height=""$($size.Height)"" alt=""$(ConvertTo-HtmlText $alt)"" />" + $nl)
            $k++; $imgCount++
        }
        [void]$sb.Append('      </div>' + $nl)
        [void]$sb.Append('      <figcaption>' + $nl)
        [void]$sb.Append("        <strong>$sheet, $(ConvertTo-HtmlText $it.Title)</strong>" + $nl)
        [void]$sb.Append("        <span class=""gx-menu"">Como criar: $(ConvertTo-HtmlText $it.Menu)</span>" + $nl)
        [void]$sb.Append("        <span class=""gx-desc"">$(ConvertTo-HtmlText $it.Desc)</span>" + $nl)
        [void]$sb.Append('      </figcaption>' + $nl)
        [void]$sb.Append('    </figure>' + $nl)
    }
    [void]$sb.Append('  </div>' + $nl)
    [void]$sb.Append($end)
    $block = $sb.ToString()

    $bi = $html.IndexOf($begin)
    $ei = $html.IndexOf($end)
    if ($bi -ge 0 -and $ei -gt $bi) {
        $html = $html.Substring(0, $bi) + $block + $html.Substring($ei + $end.Length)
    } elseif ($bi -ge 0 -or $ei -ge 0) {
        throw 'marcadores da galeria incompletos (só um dos dois existe)'
    } else {
        $sec = $html.IndexOf('<section id="demonstracao"')
        if ($sec -lt 0) { throw 'seção #demonstracao não encontrada' }
        $tbl = $html.IndexOf('</table>', $sec)
        if ($tbl -lt 0) { throw 'tabela de exemplos da seção #demonstracao não encontrada' }
        $m = [regex]::Match($html.Substring($tbl), '^</table>\s*</div>')
        if (-not $m.Success) { throw 'a tabela de exemplos não fecha com </table></div>' }
        $at = $tbl + $m.Length
        $html = $html.Substring(0, $at) + $nl + $block + $html.Substring($at)
    }

    # o número do exemplo na tabela vira link para a figura (idempotente: só o que ainda não é link)
    foreach ($it in $cov) {
        $row = "<tr data-ex=""$($it.Sheet)""><th scope=""row"">$($it.Sheet)</th>"
        $linked = "<tr data-ex=""$($it.Sheet)""><th scope=""row""><a href=""#ex-$($it.Sheet)"">$($it.Sheet)</a></th>"
        if ($html.Contains($row)) { $html = $html.Replace($row, $linked) }
        elseif (-not $html.Contains($linked)) { throw "linha da tabela do $($it.Sheet) não encontrada na seção #demonstracao" }
    }

    [IO.File]::WriteAllText($Tutorial, $html, $utf8)
    $b64 = [regex]::Matches($block, 'base64,([A-Za-z0-9+/=]+)') | ForEach-Object { $_.Groups[1].Value.Length }
    $imgBytes = ($b64 | ForEach-Object { [int]($_ * 3 / 4) } | Measure-Object -Sum).Sum
    Write-Host ('galeria gravada: {0} figuras, {1} imagens, {2:N0} bytes de imagem; tutorial com {3:N0} bytes ({4})' -f $cov.Count, $imgCount, $imgBytes, (Get-Item -LiteralPath $Tutorial).Length, $Tutorial)
    exit 0
} catch {
    Write-Host "falha: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
