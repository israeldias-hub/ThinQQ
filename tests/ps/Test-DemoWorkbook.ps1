# Pasta de demonstracao: indice com hiperlinks internos, um grafico por exemplo, cobertura = tipos e recursos da MATRIZ (lidos de
# modChartKinds.bas) + a lista fixa de recursos, sem erros, _ThinQQ muito oculta, < 3 MB, docProps sem autor/empresa, sem vinculo externo
# (abre sem o ThinQQ). Excel proprio (New-ExcelApp), somente leitura, nada e gravado. O run-tests o chama num processo proprio; termina
# sempre com exit 0 ou exit 1.
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$script:fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }
try {
    . (Join-Path $RepoRoot 'build\lib\Build-Common.ps1')
    . (Join-Path $RepoRoot 'tools\DemoCoverage.ps1')
    $path = Join-Path $RepoRoot 'exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx'
    Check (Test-Path $path) "pasta de demonstracao existe: $path"
    if (-not (Test-Path $path)) { Write-Host 'FALHAS demonstracao: arquivo ausente'; exit 1 }
    Check ((Get-Item $path).Length -lt 3MB) "tamanho < 3 MB ($((Get-Item $path).Length) bytes)"
    $cov = @(Get-TqqDemoCoverage)
    $feats = @(Get-TqqDemoFeatures)
    $kindsSrc = [IO.File]::ReadAllText((Join-Path $RepoRoot 'src\core\modules\modChartKinds.bas'))
    $allKinds = @(([regex]::Match($kindsSrc, 'Private Const KINDS As String = "(?<k>[^"]+)"').Groups['k'].Value).Trim(',').Split(','))
    Check ($allKinds.Count -ge 20) "tipos lidos de modChartKinds.bas ($($allKinds.Count))"
    foreach ($k in $allKinds) { Check (@($cov | Where-Object { $_.Kind -eq $k }).Count -ge 1) "cobertura tem o tipo '$k'" }
    # recursos da MATRIZ (a mesma leitura do Test-Docs), nao so a lista fixa de DemoCoverage.ps1 (que viria do mesmo arquivo da cobertura):
    # um recurso conta coberto por um item com o mesmo nome ou com o prefixo "<recurso>-" (legend -> legend-none)
    $matrixFeatures = @([regex]::Matches($kindsSrc, 'Case "(?<f>[a-z0-9]+)":\s*list =') | ForEach-Object { $_.Groups['f'].Value })
    Check ($matrixFeatures.Count -ge 20) "recursos da matriz lidos de modChartKinds.bas ($($matrixFeatures.Count))"
    foreach ($f in $matrixFeatures) { Check (@($cov | Where-Object { $_.Features -contains $f -or @($_.Features | Where-Object { $_ -like "$f-*" }).Count -gt 0 }).Count -ge 1) "cobertura tem o recurso da matriz '$f'" }
    foreach ($f in $feats) { Check (@($cov | Where-Object { $_.Features -contains $f }).Count -ge 1) "cobertura tem o recurso '$f'" }
    Check (@($cov | ForEach-Object { $_.Sheet } | Sort-Object -Unique).Count -eq $cov.Count) 'planilhas da cobertura sem repeticao'
    # privacidade do binario (o Test-Privacy pula arquivo com byte NUL): docProps sem autor nem empresa e nenhuma parte XML com id, caminho de
    # perfil ou e-mail (as expressoes do Test-Privacy, montadas em pedacos para este arquivo nao casar com elas); sem vinculo externo nem
    # referencia ao suplemento (a pasta abre sem o ThinQQ e sem pedir para atualizar vinculos)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $privPatterns = @(('(?<!\d)14' + '3272(?!\d)'), ('(?i)[a-z]:[\\/]+' + 'Us' + 'ers[\\/]+(?![<%$\{]|\.\.\.)[^\\/\s''"`<>]+'),
                      ('(?i)quero-?' + 'quero\.[a-z]+'), ('(?i)israel\.' + 'dias@'), '(?i)absPath', '(?i)\.xlam',
                      '(?i)Casa\s*(&amp;|&|e)\s*Obra')         # nome antigo da rede (generico, pode coincidir com lojas reais): a pasta usa um nome claramente ficticio
    $privHits = @()
    $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
    try {
        $names = @($zip.Entries | ForEach-Object { $_.FullName })
        Check (@($names | Where-Object { $_ -like 'xl/externalLinks/*' }).Count -eq 0) 'sem vinculos externos (xl/externalLinks)'
        Check (@($names | Where-Object { $_ -like '*vbaProject*' }).Count -eq 0) 'sem projeto VBA (pasta .xlsx sem macros)'
        Check ($names -contains 'docProps/core.xml') 'docProps/core.xml presente'
        foreach ($e in @($zip.Entries | Where-Object { $_.FullName -like '*.xml' -or $_.FullName -like '*.rels' })) {
            $sr = New-Object IO.StreamReader($e.Open()); $txt = $sr.ReadToEnd(); $sr.Dispose()
            # valores NUMERICOS de celula (<v>123</v>) sao dados ficticios, nao identificadores: uma venda de R$ 14..72 casaria com o id; textos
            # (inclusive resultados de formula em <v>) continuam conferidos
            $scan = $txt -replace '<v>-?[0-9.]+(E[-+]?[0-9]+)?</v>', '<v/>'
            foreach ($p in $privPatterns) { if ($scan -match $p) { $privHits += "$($e.FullName) ($p)" } }
            if ($e.FullName -eq 'docProps/core.xml') {
                Check ($txt -notmatch '<dc:creator>[^<]+</dc:creator>') 'docProps/core.xml sem autor (dc:creator vazio ou ausente)'
                Check ($txt -notmatch '<cp:lastModifiedBy>[^<]+</cp:lastModifiedBy>') 'docProps/core.xml sem cp:lastModifiedBy'
            }
            if ($e.FullName -eq 'docProps/app.xml') {
                Check ($txt -notmatch '<Company>[^<]+</Company>') 'docProps/app.xml sem empresa'
                Check ($txt -notmatch '<Manager>[^<]+</Manager>') 'docProps/app.xml sem gerente'
            }
        }
    } finally { $zip.Dispose() }
    Check ($privHits.Count -eq 0) "partes XML sem id de maquina, caminho de perfil, e-mail, absPath, .xlam nem o nome antigo da rede ($($privHits -join ', '))"
    $xl = New-ExcelApp
    $wd = Start-ExcelWatchdog $xl 240
    try {
        $wb = $xl.Workbooks.Open($path, 0, $true)
        Check ($wb.Worksheets.Item(1).Name -eq 'Indice') 'primeira planilha e Indice'
        $idx = $wb.Worksheets.Item('Indice')
        $links = @(); foreach ($h in $idx.Hyperlinks) { $links += $h }
        Check ($links.Count -eq $cov.Count) "indice com um hiperlink por exemplo ($($links.Count) de $($cov.Count))"
        $sheetNames = @(); foreach ($s in $wb.Sheets) { $sheetNames += $s.Name }
        $linked = @()
        foreach ($h in $links) {
            Check ([string]::IsNullOrEmpty($h.Address)) "hiperlink interno (sem endereco externo): $($h.SubAddress)"
            $target = ($h.SubAddress -replace "^'?(.*?)'?!.*$", '$1')
            $linked += $target
            Check ($sheetNames -contains $target) "destino existe: $($h.SubAddress)"
            if ($sheetNames -contains $target) {
                $ws = $wb.Worksheets.Item($target)
                Check ($ws.ChartObjects().Count -ge 1) "'$target' tem grafico"
                $back = $ws.Range('A1').Hyperlinks
                Check ($back.Count -eq 1 -and $back.Item(1).SubAddress -match 'Indice') "'$target'!A1 volta ao indice"
            }
        }
        foreach ($it in $cov) {
            Check ($sheetNames -contains $it.Sheet) "planilha $($it.Sheet) do item $($it.Num)"
            Check ($linked -contains $it.Sheet) "o indice leva a $($it.Sheet)"
        }
        # formatos numericos das tabelas: o texto exibido segue o codigo em ingles com os separadores do Excel (um codigo gravado como local
        # pelo PowerShell num Windows pt-BR mostra "jan/yy" e "1234567,8910"); expectativas derivadas de Application.International
        $nfi = New-Object Globalization.NumberFormatInfo
        $nfi.NumberDecimalSeparator = [string]$xl.International(3); $nfi.NumberGroupSeparator = [string]$xl.International(4)
        $dsep = [string]$xl.International(17)
        $fmtChecks = @(
            @('Dados', 'F2', { param($v) ([double]$v).ToString('#,##0', $nfi) }),
            @('Dados', 'J2', { param($v) ([double]$v).ToString('#,##0.00', $nfi) }),
            @('Lojas', 'F2', { param($v) [datetime]::FromOADate([double]$v).ToString('dd') + $dsep + [datetime]::FromOADate([double]$v).ToString('MM') + $dsep + [datetime]::FromOADate([double]$v).ToString('yyyy') }),
            @('E02', 'B7', { param($v) ([double]$v).ToString('#,##0.0', $nfi) }),
            @('E12', 'D7', { param($v) (100 * [double]$v).ToString('0.0', $nfi) + '%' }))
        foreach ($fc in $fmtChecks) {
            $c = $wb.Worksheets.Item($fc[0]).Range($fc[1]); $exp = & $fc[2] $c.Value2
            Check ([string]$c.Text -eq $exp) "formato de $($fc[0])!$($fc[1]): texto [$($c.Text)] esperado [$exp]"
        }
        foreach ($sh in 'E04', 'E12', 'E26') {
            $t = [string]$wb.Worksheets.Item($sh).Range('A7').Text
            Check ($t -match '\d{2}$' -and $t -notmatch '(?i)yy|aa') "mes/ano em $sh!A7 com o ano em 2 digitos ([$t])"
        }
        # o mesmo total em R$ mil e o mesmo numero em todas as planilhas (as partes fecham com o total arredondado): cimento 2025 em E01 (canais),
        # E24 e E22 (total); Sudeste 2025 em E03/E06 (canais), E13/E23 (grupos de produto) e E17/E18 (total)
        function Get-RowSum($sh, [string]$name) {
            $ws = $wb.Worksheets.Item($sh); $r = 7
            while ([string]$ws.Cells.Item($r, 1).Value2 -ne $name -and $r -lt 40) { $r++ }
            $t = 0.0; for ($c = 2; $c -le 5; $c++) { $v = $ws.Cells.Item($r, $c).Value2; if ($v -is [double]) { $t += $v } }
            return $t
        }
        $cim = @('E01', 'E24' | ForEach-Object { Get-RowSum $_ 'Cimento e argamassa' }) + @([double]$wb.Worksheets.Item('E22').Range('B8').Value2)
        Check (@($cim | Sort-Object -Unique).Count -eq 1) "cimento 2025 igual em E01, E24 e E22 ($($cim -join ', '))"
        $sud = @('E03', 'E06', 'E13', 'E23', 'E17', 'E18' | ForEach-Object { Get-RowSum $_ 'Sudeste' })
        Check (@($sud | Sort-Object -Unique).Count -eq 1) "Sudeste 2025 igual em E03, E06, E13, E23, E17 e E18 ($($sud -join ', '))"
        # graficos vivos sem o ThinQQ (nenhum suplemento aberto neste Excel): +1.000.000 em Dados (Vendas e Vendas B2C de uma linha de cimento em 2025) muda a tabela
        # do E01 (SOMASES), a serie B2C do grafico e o rotulo do total (ligado a _ThinQQ); nada e gravado (pasta somente leitura, fechada sem salvar)
        $wsDados = $wb.Worksheets.Item('Dados')
        $keys = $wsDados.Range('D2').Resize(($wsDados.UsedRange.Rows.Count - 1), 2).Value2
        $hit = 0
        for ($i = 1; $i -le $keys.GetLength(0); $i++) {
            $serial = [double]$keys[$i, 2]                   # fora da chamada: no 5.1 a virgula do indice [i, 2] dentro de (...) vira separador de argumento
            if ($keys[$i, 1] -eq 'Cimento e argamassa' -and [datetime]::FromOADate($serial).Year -eq 2025) { $hit = $i + 1; break }
        }
        Check ($hit -gt 0) "linha de cimento em 2025 encontrada em Dados ($hit)"
        $e01 = $wb.Worksheets.Item('E01'); $cht = $e01.ChartObjects(1).Chart
        $ser = $null; $tot = $null
        foreach ($s in $cht.SeriesCollection()) { if ($s.Name -eq 'B2C') { $ser = $s }; if ($s.Name -eq 'tqq_total') { $tot = $s } }
        Check ($null -ne $ser -and $null -ne $tot) 'E01 com a serie B2C e o total (tqq_total)'
        if ($hit -gt 0 -and $null -ne $ser -and $null -ne $tot) {
            $cellBefore = [double]$e01.Range('B7').Value2; $serBefore = [double]@($ser.Values)[0]; $lblBefore = [string]$tot.Points(1).DataLabel.Text
            foreach ($col in 6, 11) { $kCell = $wsDados.Cells.Item($hit, $col); $kCell.Value2 = [double]$kCell.Value2 + 1000000 }   # Vendas e Vendas B2C
            $xl.Calculate()
            Check ([double]$e01.Range('B7').Value2 -eq $cellBefore + 1000) "E01!B7 acompanha Dados ($cellBefore -> $($e01.Range('B7').Value2))"
            Check ([double]@($ser.Values)[0] -eq $serBefore + 1000) "serie B2C do E01 acompanha a tabela ($serBefore -> $(@($ser.Values)[0]))"
            Check ([string]$tot.Points(1).DataLabel.Text -ne $lblBefore) "rotulo do total do E01 acompanha ($lblBefore -> $($tot.Points(1).DataLabel.Text))"
        }
        # opcoes que o core gravou no bloco do grafico (_ThinQQ, linha de meta, coluna 4: "chave=valor;...") e o tipo (coluna 2)
        function Get-ChartBlockInfo($co) {
            $blk = $wb.Names.Item('tqq_' + $co.Name.Substring(4)).RefersToRange
            $o = @{}
            foreach ($p in ([string]$blk.Cells.Item(1, 4).Value2).Split(';')) { $q = $p.IndexOf('='); if ($q -gt 0) { $o[$p.Substring(0, $q)] = $p.Substring($q + 1) } }
            return @{ Kind = [string]$blk.Cells.Item(1, 2).Value2; Opt = $o; Block = $blk }
        }
        # cada recurso da cobertura deixou no grafico o que o comando grava (opcao no bloco, anotacao tqq_ann_*, escala do Gantt, formulas do
        # Arredondar intervalo): um comando que nao fez nada nao passa. O mapa e deste teste (independente da lista de DemoCoverage.ps1) e cobre
        # toda a lista fixa de recursos.
        function Test-HasShape($co, [string]$prefix) { foreach ($sh in $co.Chart.Shapes) { if (([string]$sh.Name).StartsWith($prefix)) { return $true } }; return $false }
        $featCheck = @{
            'totals'       = { param($i, $co, $ws) $i.Opt['totals'] -eq '1' }
            'labels'       = { param($i, $co, $ws) $i.Opt['labels'] -eq '1' }
            'legend-none'  = { param($i, $co, $ws) $i.Opt['legend'] -eq 'none' -and -not $co.Chart.HasLegend }
            'lineend'      = { param($i, $co, $ws) $i.Opt['legend'] -eq 'lineend' }
            'numfmt'       = { param($i, $co, $ws) $i.Opt['numfmt'] -eq 'mil' }
            'gap'          = { param($i, $co, $ws) $i.Opt['gap'] -eq '120' }
            'palette'      = { param($i, $co, $ws) $i.Opt['palette'] -eq 'cinza' }
            'cagr'         = { param($i, $co, $ws) Test-HasShape $co 'tqq_ann_cagr_' }
            'diff'         = { param($i, $co, $ws) Test-HasShape $co 'tqq_ann_diff_' }
            'bracket'      = { param($i, $co, $ws) Test-HasShape $co 'tqq_ann_bracket_' }
            'vline'        = { param($i, $co, $ws) $i.Opt['vline'] -eq '1' }
            'ind100'       = { param($i, $co, $ws) $i.Opt['ind100'] -eq '1' -and (Test-HasShape $co 'tqq_ann_ind') }
            'samescale'    = { param($i, $co, $ws) $i.Opt['vmax'] -match '^\d' -and [double]$co.Chart.Axes(2).MaximumScale -eq [double]$i.Opt['vmax'] }
            'secondary'    = { param($i, $co, $ws) $i.Opt['secondary'] -eq '1' }
            'pieout'       = { param($i, $co, $ws) $i.Opt['pieout'] -eq '1' }
            'gantt-weeks'  = { param($i, $co, $ws) [string]$i.Block.Cells.Item(3, 12).Value2 -eq 'w' }
            'gantt-months' = { param($i, $co, $ws) [string]$i.Block.Cells.Item(3, 12).Value2 -eq 'm' }
            'roundsum'     = { param($i, $co, $ws) $i.Opt['roundsum'] -eq '1' }
            'roundrange'   = { param($i, $co, $ws) $n = 0; foreach ($c in $ws.UsedRange.Cells) { if ([string]$c.Formula -match 'SUMPRODUCT') { $n++ } }; $n -ge 2 }
            'break'        = { param($i, $co, $ws) $i.Opt['vbreak'] -match '^\d' -and [double]$co.Chart.Axes(2).MaximumScale -eq [double]$i.Opt['vbreak'] }
            'trend'        = { param($i, $co, $ws) $n = 0; foreach ($sr in $co.Chart.SeriesCollection()) { $n += $sr.Trendlines().Count }; $i.Opt['trend'] -eq '1' -and $n -ge 1 }
            'quad'         = { param($i, $co, $ws) $i.Opt['quad'] -eq '1' }
            'connectors'   = { param($i, $co, $ws) $i.Opt['connectors'] -eq '1' }
        }
        foreach ($f in $feats) { Check ($featCheck.ContainsKey($f)) "o teste sabe conferir o recurso '$f'" }
        foreach ($it in $cov) {
            $ws = $wb.Worksheets.Item($it.Sheet)
            foreach ($co in @($ws.ChartObjects())) {
                $info = Get-ChartBlockInfo $co
                Check ($info.Kind -eq $it.Kind) "$($it.Sheet): grafico do tipo '$($it.Kind)' ($($info.Kind))"
                foreach ($f in $it.Features) {
                    $ok = $false
                    if ($featCheck.ContainsKey($f)) { try { $ok = [bool](& $featCheck[$f] $info $co $ws) } catch { $ok = $false } }
                    Check $ok "$($it.Sheet): recurso '$f' presente no grafico (opcoes: $(($info.Opt.Keys | Sort-Object | ForEach-Object { "$_=$($info.Opt[$_])" }) -join ';'))"
                }
            }
        }
        # E02: a tabela tem 1 casa decimal; com o formato padrao (inteiro) dois valores diferentes viram o mesmo rotulo
        $e02 = Get-ChartBlockInfo ($wb.Worksheets.Item('E02').ChartObjects(1))
        Check ($e02.Opt['numfmt'] -eq 'dec1') "E02: rotulos com 1 casa como a tabela (numfmt=$($e02.Opt['numfmt']))"
        # E12: com o eixo secundario a linha da margem fica acima dos totais (core 1.2.0): os totais das colunas ficam ligados
        $e12 = Get-ChartBlockInfo ($wb.Worksheets.Item('E12').ChartObjects(1))
        Check ($e12.Opt['totals'] -eq '1' -and $e12.Opt['secondary'] -eq '1') "E12: totais e eixo secundario ligados (totals=$($e12.Opt['totals']), secondary=$($e12.Opt['secondary']))"
        # E20/E21: a margem e uma fracao com formato % na tabela e os eixos de valores mostram % (o eixo segue o formato da origem); o texto do
        # 1o rotulo do eixo vertical e conferido pelo codigo do eixo, com o separador decimal do Excel
        foreach ($sh in 'E20', 'E21') {
            $wsS = $wb.Worksheets.Item($sh)
            $m = $wsS.Range('C7')
            Check ([double]$m.Value2 -gt 0 -and [double]$m.Value2 -lt 1 -and [string]$m.Text -match '%$') "${sh}: margem como fracao com % na tabela ($($m.Value2) -> [$($m.Text)])"
            $ay = $wsS.ChartObjects(1).Chart.Axes(2)
            Check ([string]$ay.TickLabels.NumberFormat -match '%') "${sh}: eixo da margem em % ($($ay.TickLabels.NumberFormat))"
            Check ([double]$ay.MinimumScale -gt 0) "${sh}: eixo da margem nao comeca em zero (minimo $($ay.MinimumScale))"
        }
        # E24 (corte de eixo): na coluna cortada (total acima do maximo do eixo) nenhum segmento fica inteiro escondido: a base de cada serie de
        # dados fica abaixo do corte, o ponto tem altura desenhada e o rotulo continua ligado
        $cb = $wb.Worksheets.Item('E24').ChartObjects(1).Chart
        $vmax = [double]$cb.Axes(2).MaximumScale
        $dataSer = @(); foreach ($s in $cb.SeriesCollection()) { if (-not ([string]$s.Name).StartsWith('tqq_')) { $dataSer += $s } }
        $nCat = @($dataSer[0].Values).Count; $cutCats = 0
        for ($ci = 1; $ci -le $nCat; $ci++) {
            $tot = 0.0; foreach ($s in $dataSer) { $tot += [double]@($s.Values)[$ci - 1] }
            if ($tot -le $vmax) { continue }
            $cutCats++; $base = 0.0
            foreach ($s in $dataSer) {
                $v = [double]@($s.Values)[$ci - 1]; $pt = $s.Points($ci)
                Check ($base -lt $vmax -and [double]$pt.Height -gt 1 -and $pt.HasDataLabel) "E24: segmento '$($s.Name)' da coluna cortada $ci visivel (base $base < corte $vmax, altura $([Math]::Round([double]$pt.Height, 1)) pt, rotulo $($pt.HasDataLabel))"
                $base += $v
            }
        }
        Check ($cutCats -eq 1) "E24: exatamente uma coluna cortada ($cutCats)"
        $calc = $null; try { $calc = $wb.Worksheets.Item('_ThinQQ') } catch { }
        Check ($null -ne $calc -and $calc.Visible -eq 2) '_ThinQQ muito oculta'
        $errs = 0
        foreach ($ws in $wb.Worksheets) {
            if ($ws.Name -eq '_ThinQQ') { continue }
            $vals = $ws.UsedRange.Value2
            if ($null -eq $vals) { continue }
            if ($vals -is [array]) { foreach ($v in $vals) { if ($v -is [int] -and $v -lt -2146826000 -and $v -gt -2146827000) { $errs++ } } }
            elseif ($vals -is [int] -and $vals -lt -2146826000 -and $vals -gt -2146827000) { $errs++ }
        }
        Check ($errs -eq 0) "nenhuma celula com valor de erro ($errs)"
        $wb.Close($false)
    } catch { Check $false "excecao: $($_.Exception.Message)" }
    finally { Stop-ExcelWatchdog $wd; Close-ExcelApp $xl }
} catch {
    Check $false "excecao: $($_.Exception.Message)"
}
if ($script:fail -gt 0) { Write-Host "FALHAS demonstracao: $script:fail"; exit 1 }
Write-Host 'Test-DemoWorkbook: OK'
exit 0
