<#
.SYNOPSIS  Gera exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx: dados FICTÍCIOS determinísticos de uma rede de materiais de construção e um
           exemplo por funcionalidade do ThinQQ, pela API real do core (Application.Run num Excel novo e invisível).
.EXAMPLE   robocopy <repo> C:\temp\tqq-f3 /MIR /XD .git .superpowers dist ; pwsh C:\temp\tqq-f3\build\build.ps1 -Target core -Configuration Test
           pwsh tools\New-DemoWorkbook.ps1 -CoreXlam C:\temp\tqq-f3\dist\ThinQQ.core.test.xlam
.NOTES     Core de TESTE de propósito: em Release, Core_Notify abre MsgBox e travaria a automação. O .xlsx não leva código.
           Reprodutível: semente fixa, datas fixas (nada depende da data nem do idioma da máquina; fórmulas em inglês por .Formula, formatos
           numéricos por Set-NumFmt, com os códigos em inglês e o LCID en-US na chamada COM). A lista de exemplos vem de tools\DemoCoverage.ps1 (a mesma do teste
           tests\ps\Test-DemoWorkbook.ps1). Nunca altera Application.UserName (grava no registro do Office do usuário): o autor sai do arquivo
           por RemovePersonalInformation. Sai com 0 (gerado) ou 1 (falha; nada é gravado em -OutFile).
#>
param(
    [Parameter(Mandatory)][string]$CoreXlam,
    [string]$OutFile = (Join-Path $PSScriptRoot '..\exemplos\ThinQQ-Demonstracao-Materiais-Construcao.xlsx'),
    [int]$TimeoutSec = 600
)
$ErrorActionPreference = 'Stop'
$sw = [Diagnostics.Stopwatch]::StartNew()
. (Join-Path $PSScriptRoot '..\build\lib\Build-Common.ps1')
. (Join-Path $PSScriptRoot 'DemoCoverage.ps1')
if (-not (Test-Path $CoreXlam)) { Write-Host "core de teste não encontrado: $CoreXlam" -ForegroundColor Red; exit 1 }
$CoreXlam = (Resolve-Path $CoreXlam).Path
$OutFile = [IO.Path]::GetFullPath($OutFile)
New-Item -ItemType Directory -Force (Split-Path $OutFile -Parent) | Out-Null

# =====================================================================================================================================
# Dados fictícios determinísticos (gerador congruencial com semente fixa; nenhuma pessoa, empresa real ou dado pessoal)
# =====================================================================================================================================
$script:seed = 20261011
function Get-DemoRandom { $script:seed = ([int64]$script:seed * 1103515245 + 12345) % 2147483648; return [double]$script:seed / 2147483648 }

$company = 'Rede Demo ThinQQ - Materiais de Construção (fictícia)'      # nome claramente fictício (nenhuma empresa real)
# Lojas (nomes e cidades inventados), na ordem das regiões. Size = porte relativo da loja.
$stores = @(
    @{ Name = 'Aurora Centro';  City = 'Vila Aurora';    Region = 'Sudeste';      Size = 1.45; Open = [datetime]::new(2009, 3, 14) }
    @{ Name = 'Aurora Norte';   City = 'Vila Aurora';    Region = 'Sudeste';      Size = 0.95; Open = [datetime]::new(2016, 8, 6) }
    @{ Name = 'Porto Sereno';   City = 'Porto Sereno';   Region = 'Sul';          Size = 1.15; Open = [datetime]::new(2011, 5, 21) }
    @{ Name = 'Serra Clara';    City = 'Serra Clara';    Region = 'Sul';          Size = 0.75; Open = [datetime]::new(2018, 10, 3) }
    @{ Name = 'Lago Prateado';  City = 'Lago Prateado';  Region = 'Nordeste';     Size = 0.85; Open = [datetime]::new(2014, 2, 8) }
    @{ Name = 'Monte Alvorada'; City = 'Monte Alvorada'; Region = 'Nordeste';     Size = 0.65; Open = [datetime]::new(2020, 11, 14) }
    @{ Name = 'Campo Dourado';  City = 'Campo Dourado';  Region = 'Centro-Oeste'; Size = 1.05; Open = [datetime]::new(2013, 7, 27) }
    @{ Name = 'Rio Turquesa';   City = 'Rio Turquesa';   Region = 'Norte';        Size = 0.55; Open = [datetime]::new(2022, 4, 9) }
)
$regions = @('Sudeste', 'Sul', 'Nordeste', 'Centro-Oeste', 'Norte')
# Departamentos: Weight = peso nas vendas; Price = preço médio do item (R$); Ticket = ticket médio de referência (R$); Margin = margem bruta;
# Growth = crescimento anual próprio, somado ao da loja (o mix muda ao longo dos 24 meses).
$depts = @(
    @{ Name = 'Cimento e argamassa';   Group = 'Construção';  Weight = 4.0; Price = 34;  Ticket = 420; Margin = 0.22; Growth =  0.01 }
    @{ Name = 'Pisos e revestimentos'; Group = 'Acabamento';  Weight = 1.8; Price = 58;  Ticket = 690; Margin = 0.34; Growth =  0.12 }
    @{ Name = 'Tintas';                Group = 'Acabamento';  Weight = 1.3; Price = 165; Ticket = 360; Margin = 0.38; Growth =  0.05 }
    @{ Name = 'Hidráulica';            Group = 'Instalações'; Weight = 1.1; Price = 28;  Ticket = 210; Margin = 0.36; Growth =  0.03 }
    @{ Name = 'Elétrica';              Group = 'Instalações'; Weight = 1.0; Price = 24;  Ticket = 120; Margin = 0.40; Growth =  0.06 }
    @{ Name = 'Madeiras';              Group = 'Construção';  Weight = 0.9; Price = 42;  Ticket = 380; Margin = 0.29; Growth = -0.03 }
    @{ Name = 'Ferramentas';           Group = 'Instalações'; Weight = 0.7; Price = 85;  Ticket = 160; Margin = 0.43; Growth =  0.08 }
    @{ Name = 'Ferragens';             Group = 'Construção';  Weight = 0.5; Price = 9;   Ticket = 75;  Margin = 0.46; Growth =  0.02 }
)
$groups = @('Construção', 'Acabamento', 'Instalações')
# Perfil regional: multiplicador das vendas de cada grupo de produto (mais materiais básicos no Norte e Nordeste, mais acabamento no Sudeste e Sul)
$regionMix = @{
    'Sudeste' = @{ 'Construção' = 0.85; 'Acabamento' = 1.25; 'Instalações' = 1.05 }; 'Sul' = @{ 'Construção' = 0.95; 'Acabamento' = 1.10; 'Instalações' = 1.05 }
    'Nordeste' = @{ 'Construção' = 1.15; 'Acabamento' = 0.85; 'Instalações' = 0.95 }; 'Centro-Oeste' = @{ 'Construção' = 1.10; 'Acabamento' = 0.95; 'Instalações' = 0.90 }
    'Norte' = @{ 'Construção' = 1.25; 'Acabamento' = 0.75; 'Instalações' = 0.90 }
}
$nS = $stores.Count; $nD = $depts.Count; $nM = 24
$months = @(for ($m = 0; $m -lt $nM; $m++) { [datetime]::new(2024 + [Math]::Floor($m / 12), ($m % 12) + 1, 1) })
# por loja: participação B2B (vendas a construtoras e empreiteiros), crescimento anual e diferenças próprias de ticket e de margem. As faixas
# de ticket e de margem das lojas ficam estreitas (máximo - mínimo < 1/6 do máximo) de propósito: assim a escala automática da dispersão (E20)
# não começa em zero e os pontos se espalham pelo gráfico.
foreach ($s in $stores) { $s.B2B = 0.18 + 0.30 * (Get-DemoRandom); $s.Growth = 0.03 + 0.07 * (Get-DemoRandom); $s.TicketF = 0.98 + 0.04 * (Get-DemoRandom); $s.MarginAdj = 0.01 * ((Get-DemoRandom) - 0.5) }

# Cubo loja x departamento x mês (índice plano) e as linhas da planilha Dados
$cube = @{}
foreach ($k in 'sales', 'cost', 'b2c', 'b2b', 'receipts') { $cube[$k] = New-Object 'double[]' ($nS * $nD * $nM) }
$rows = New-Object System.Collections.Generic.List[object]
for ($si = 0; $si -lt $nS; $si++) {
    $s = $stores[$si]
    for ($di = 0; $di -lt $nD; $di++) {
        $d = $depts[$di]
        for ($m = 0; $m -lt $nM; $m++) {
            $season = 1 + 0.10 * [Math]::Cos(2 * [Math]::PI * (($m % 12) - 9) / 12)      # pico em outubro (estação seca), vale em abril
            $trend = [Math]::Pow(1 + $s.Growth + $d.Growth, $m / 12)
            $sales = [Math]::Round(110000 * $s.Size * $d.Weight * $regionMix[$s.Region][$d.Group] * $season * $trend * (0.94 + 0.12 * (Get-DemoRandom)), 0)
            $margin = $d.Margin - 0.14 * ($s.B2B - 0.33) + $s.MarginAdj + 0.03 * ((Get-DemoRandom) - 0.5)   # mais B2B: margem menor
            $cost = [Math]::Round($sales * (1 - $margin), 0)
            $ticket = [Math]::Round($d.Ticket * (0.85 + 0.45 * $s.B2B) * $s.TicketF * (0.95 + 0.1 * (Get-DemoRandom)), 2)   # mais B2B: ticket maior
            $units = [Math]::Max(1, [Math]::Round($sales / ($d.Price * (0.95 + 0.1 * (Get-DemoRandom))), 0))
            $b2b = [Math]::Round($sales * $s.B2B * (0.95 + 0.1 * (Get-DemoRandom)), 0)
            $b2c = $sales - $b2b
            $ix = ($si * $nD + $di) * $nM + $m
            $cube.sales[$ix] = $sales; $cube.cost[$ix] = $cost; $cube.b2c[$ix] = $b2c; $cube.b2b[$ix] = $b2b; $cube.receipts[$ix] = $sales / $ticket
            $rows.Add(@($s.Name, $s.City, $s.Region, $d.Name, $months[$m], $sales, $cost, ($sales - $cost), $units, $ticket, $b2c, $b2b))
        }
    }
}
$allS = @(0..($nS - 1)); $allD = @(0..($nD - 1)); $y24 = @(0..11); $y25 = @(12..23)
# Soma de uma medida do cubo sobre conjuntos de lojas, departamentos e meses (R$, sem arredondar)
function Get-DemoSum([string]$metric, [int[]]$sIdx, [int[]]$dIdx, [int[]]$mIdx) {
    $a = $cube[$metric]; $t = 0.0
    foreach ($si in $sIdx) { foreach ($di in $dIdx) { $b = ($si * $nD + $di) * $nM; foreach ($m in $mIdx) { $t += $a[$b + $m] } } }
    return $t
}
function Get-DemoK([double]$v) { return [Math]::Round($v / 1000, 0) }                       # R$ -> R$ mil
# Partes em R$ mil que fecham com o total arredondado: a última parte é K(total) - soma das outras (já arredondadas). Assim o mesmo total
# (ex.: vendas 2025 do Sudeste ou do cimento) mostra o mesmo número em todas as planilhas, venha de canais, de grupos ou do total.
function Get-DemoParts([double]$total, [double[]]$parts) {
    $out = @(); $acc = 0.0
    for ($i = 0; $i -lt $parts.Count - 1; $i++) { $k = Get-DemoK $parts[$i]; $out += $k; $acc += $k }
    return ,@($out + @((Get-DemoK $total) - $acc))
}
function Get-StoresOf([string]$region) { return @(for ($i = 0; $i -lt $nS; $i++) { if ($stores[$i].Region -eq $region) { $i } }) }
function Get-DeptsOf([string]$group) { return @(for ($i = 0; $i -lt $nD; $i++) { if ($depts[$i].Group -eq $group) { $i } }) }
# ordens: departamentos e lojas pelas vendas de 2025 (maior primeiro)
$deptOrder = @($allD | Sort-Object { - (Get-DemoSum 'sales' $allS @($_) $y25) })
$storeOrder = @($allS | Sort-Object { - (Get-DemoSum 'sales' @($_) $allD $y25) })
$s1 = $storeOrder[0]; $sLast = $storeOrder[$nS - 1]                                             # maior e menor loja
if ($stores[$s1].Name -ne 'Aurora Centro') { throw 'tools\DemoCoverage.ps1 cita a loja Aurora Centro como a maior (E04, E12, E26): ajuste os títulos' }
$top3 = @($deptOrder[0..2])

# =====================================================================================================================================
# Excel
# =====================================================================================================================================
$xl = $null; $wd = $null; $exitCode = 1
$tmpOut = Join-Path ([IO.Path]::GetTempPath()) ('tqq_demo_' + [Guid]::NewGuid().ToString('N') + '.xlsx')
try {
    $xl = New-ExcelApp
    $wd = Start-ExcelWatchdog $xl $TimeoutSec
    $xl.EnableEvents = $false
    $core = $xl.Workbooks.Open($CoreXlam)
    $run = "'$($core.Name)'!"
    function Invoke-Core([string]$proc) {
        $a = @($args)
        switch ($a.Count) {
            0 { return $xl.Run($run + $proc) }
            1 { return $xl.Run($run + $proc, $a[0]) }
            2 { return $xl.Run($run + $proc, $a[0], $a[1]) }
            default { return $xl.Run($run + $proc, $a[0], $a[1], $a[2]) }
        }
    }
    # Comando do core que não pode deixar aviso (Core_Notify = recusa, validação ou falha registrada): limpa antes e confere depois.
    function Invoke-CoreQuiet([string]$ctx, [string]$proc) {
        Invoke-Core 'Core_ClearLastNotice' | Out-Null
        $r = Invoke-Core $proc @args
        $n = Invoke-Core 'Core_LastNotice'
        if ($n) { throw "aviso do core em ${ctx} ($proc): $n" }
        return $r
    }

    $wb = $xl.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
    $idx = $wb.Worksheets.Item(1); $idx.Name = 'Indice'

    # Formato numérico com os códigos em INGLÊS (como Range.NumberFormat no VBA). Pelo PowerShell a chamada COM leva o LCID da cultura do
    # processo e, num Windows pt-BR, Range.NumberFormat passa a ler o código como local: 'mmm/yy' mostra "jan/yy" e '#,##0.0' vira
    # '#,##00' (medido). InvokeMember com a cultura en-US fixa o LCID da chamada: o resultado é o mesmo em qualquer idioma do Windows.
    $script:enUS = [Globalization.CultureInfo]::GetCultureInfo('en-US')
    function Set-NumFmt($rng, [string]$code) { [void][System.__ComObject].InvokeMember('NumberFormat', [Reflection.BindingFlags]::SetProperty, $null, $rng, @($code), $script:enUS) }

    # ---- Dados (uma linha por loja x departamento x mês)
    $wsD = $wb.Worksheets.Add([Type]::Missing, $idx); $wsD.Name = 'Dados'
    $hdr = [object[]]('Loja', 'Cidade', 'Região', 'Departamento', 'Mês', 'Vendas (R$)', 'Custo (R$)', 'Margem (R$)', 'Unidades', 'Ticket médio (R$)', 'Vendas B2C (R$)', 'Vendas B2B (R$)')
    $wsD.Range('A1:L1').Value2 = $hdr
    $arr = New-Object 'object[,]' $rows.Count, 12
    for ($i = 0; $i -lt $rows.Count; $i++) { for ($j = 0; $j -lt 12; $j++) { $arr[$i, $j] = $rows[$i][$j] } }
    $nRows = $rows.Count
    $wsD.Range('A2').Resize($nRows, 12).Value2 = $arr
    Set-NumFmt ($wsD.Range('E2').Resize($nRows, 1)) 'dd/mm/yyyy'
    Set-NumFmt ($wsD.Range('E2').Resize($nRows, 1)) 'mmm/yyyy'
    Set-NumFmt ($wsD.Range('F2').Resize($nRows, 4)) '#,##0'
    Set-NumFmt ($wsD.Range('J2').Resize($nRows, 1)) '#,##0.00'
    Set-NumFmt ($wsD.Range('K2').Resize($nRows, 2)) '#,##0'
    $wsD.Range('A1:L1').Font.Bold = $true
    $wsD.Range('A1:L1').AutoFilter() | Out-Null
    $wsD.Columns('A:L').AutoFit() | Out-Null

    # ---- Lojas e Obras
    $wsL = $wb.Worksheets.Add([Type]::Missing, $wsD); $wsL.Name = 'Lojas'
    $wsL.Range('A1:F1').Value2 = [object[]]('Loja', 'Cidade', 'Região', 'Vendas 2025 (R$ mil)', 'Meta 2026 (R$ mil)', 'Inauguração')
    for ($i = 0; $i -lt $nS; $i++) {
        $s = $stores[$i]; $v25 = Get-DemoK (Get-DemoSum 'sales' @($i) $allD $y25)
        $wsL.Range("A$($i + 2):F$($i + 2)").Value2 = [object[]]($s.Name, $s.City, $s.Region, $v25, [Math]::Round($v25 * 1.08, 0), $s.Open)
    }
    Set-NumFmt ($wsL.Range("D2:E$($nS + 1)")) '#,##0'
    Set-NumFmt ($wsL.Range("F2:F$($nS + 1)")) 'dd/mm/yyyy'
    $wsL.Range('A1:F1').Font.Bold = $true; $wsL.Columns('A:F').AutoFit() | Out-Null

    $obrasQ = @(@('Reforma da fachada (Aurora Centro)', [datetime]::new(2026, 3, 2), [datetime]::new(2026, 3, 27)),
                @('Troca do telhado (Porto Sereno)', [datetime]::new(2026, 3, 16), [datetime]::new(2026, 4, 24)),
                @('Novo depósito (Serra Clara)', [datetime]::new(2026, 4, 6), [datetime]::new(2026, 5, 15)),
                @('Inauguração do depósito', [datetime]::new(2026, 5, 18), $null))
    $obrasY = @(@('Projeto e licenças', [datetime]::new(2026, 1, 5), [datetime]::new(2026, 3, 27)),
                @('Obra civil', [datetime]::new(2026, 4, 6), [datetime]::new(2026, 12, 18)),
                @('Instalações e acabamento', [datetime]::new(2027, 1, 4), [datetime]::new(2027, 4, 30)),
                @('Montagem e estoque', [datetime]::new(2027, 5, 3), [datetime]::new(2027, 6, 11)),
                @('Inauguração', [datetime]::new(2027, 6, 15), $null))
    $wsO = $wb.Worksheets.Add([Type]::Missing, $wsL); $wsO.Name = 'Obras'
    $wsO.Range('A1:D1').Value2 = [object[]]('Obra', 'Início', 'Fim', 'Plano')
    $r = 2
    foreach ($o in $obrasQ) { $wsO.Cells.Item($r, 1).Value2 = $o[0]; $wsO.Cells.Item($r, 2).Value = $o[1]; if ($o[2]) { $wsO.Cells.Item($r, 3).Value = $o[2] }; $wsO.Cells.Item($r, 4).Value2 = 'Reformas 2026'; $r++ }
    foreach ($o in $obrasY) { $wsO.Cells.Item($r, 1).Value2 = $o[0]; $wsO.Cells.Item($r, 2).Value = $o[1]; if ($o[2]) { $wsO.Cells.Item($r, 3).Value = $o[2] }; $wsO.Cells.Item($r, 4).Value2 = 'Nova loja Vale do Ipê'; $r++ }
    Set-NumFmt ($wsO.Range("B2:C$($r - 1)")) 'dd/mm/yyyy'; $wsO.Range('A1:D1').Font.Bold = $true; $wsO.Columns('A:D').AutoFit() | Out-Null

    # ---- auxiliares dos exemplos
    function Write-Table($ws, [string]$topLeft, [object[]]$header, [object[][]]$body, [string]$numFmt = '#,##0') {
        $ws.Range($topLeft).Resize(1, $header.Count).Value2 = $header
        $ws.Range($topLeft).Resize(1, $header.Count).Font.Bold = $true
        if ($header.Count -gt 1) { $ws.Range($topLeft).Offset(0, 1).Resize(1, $header.Count - 1).HorizontalAlignment = -4152 }   # xlRight: sobre os números
        for ($i = 0; $i -lt $body.Count; $i++) {
            # valores sem o invólucro PSObject: no Windows PowerShell 5.1 um $_ de ForEach-Object (ex.: o nome da região) chega embrulhado e o
            # Range.Value2 falha com E_OUTOFMEMORY (0x8007000E, medido); o pwsh 7 desembrulha sozinho
            $row = New-Object 'object[]' $body[$i].Count
            for ($j = 0; $j -lt $row.Count; $j++) { $v = $body[$i][$j]; if ($null -ne $v) { $v = $v.PSObject.BaseObject }; $row[$j] = $v }
            $ws.Range($topLeft).Offset($i + 1, 0).Resize(1, $row.Count).Value2 = $row
        }
        $rng = $ws.Range($topLeft).Resize($body.Count + 1, $header.Count)
        if ($numFmt) { Set-NumFmt ($rng.Offset(1, 1).Resize($body.Count, $header.Count - 1)) $numFmt }
        # largura: cabe o cabeçalho e os números (nunca encolhe uma coluna já usada por outra tabela da planilha)
        for ($c = 1; $c -le $header.Count; $c++) {
            $col = $rng.Columns.Item($c); $old = [double]$col.ColumnWidth
            $col.AutoFit() | Out-Null
            $min = [Math]::Max($old, $(if ($c -eq 1) { 24 } else { 13 }))
            if ([double]$col.ColumnWidth -lt $min) { $col.ColumnWidth = $min }
        }
        return ,$rng                                   # a vírgula: sem ela o PowerShell desenrola o Range célula por célula
    }
    # Datas de mês na 1a coluna: passam por dd/mm/yyyy antes de mmm/yy (de General direto, o rótulo ligado mostra o mês em inglês).
    function Set-MonthFormat($rng) { Set-NumFmt ($rng) 'dd/mm/yyyy'; Set-NumFmt ($rng) 'mmm/yy' }
    function New-ExampleSheet($after, $item, [string]$unit) {
        $ws = $wb.Worksheets.Add([Type]::Missing, $after)
        $ws.Name = $item.Sheet
        $ws.Hyperlinks.Add($ws.Range('A1'), '', "'Indice'!A1", 'Voltar ao índice', '« Voltar ao índice') | Out-Null
        $ws.Range('A2').Value2 = "$($item.Num)) $($item.Title)"
        $ws.Range('A2').Font.Bold = $true; $ws.Range('A2').Font.Size = 13
        $ws.Range('A3').Value2 = $item.Desc
        $ws.Range('A3').Font.Italic = $true
        $note = "Como criar no ThinQQ: $($item.Menu)"
        if (-not $note.EndsWith('.')) { $note = "$note." }
        if ($unit) { $note = "$note  Valores: $unit." }
        $ws.Range('A4').Value2 = $note
        $ws.Range('A4').Font.Color = 0x595959
        $ws.Columns('A').ColumnWidth = 24; $ws.Columns('B:D').ColumnWidth = 13
        $ws.Tab.Color = 0x9F6B1F
        return $ws
    }
    function Add-Chart($ws, $rng, [string]$kind) {
        $before = $ws.ChartObjects().Count
        $ws.Activate(); $rng.Select() | Out-Null
        Invoke-CoreQuiet "Core_Insert $kind em $($ws.Name)" 'Core_Insert' $kind | Out-Null
        if ($ws.ChartObjects().Count -ne $before + 1) { throw "Core_Insert $kind não criou o gráfico em $($ws.Name)" }
        return $ws.ChartObjects().Item($before + 1)
    }
    # Muda o tamanho/posição de um gráfico e refaz o layout pelo core (rótulos, anotações, escala), como o evento faria no Excel do usuário.
    function Set-ChartBox($ws, $co, [double]$width, [double]$height, [double]$left = -1, [double]$top = -1) {
        $co.Width = $width; $co.Height = $height
        if ($left -ge 0) { $co.Left = $left }
        if ($top -ge 0) { $co.Top = $top }
        $ws.Activate(); $co.Select() | Out-Null
        Invoke-CoreQuiet "Core_Refresh em $($ws.Name)" 'Core_Refresh' | Out-Null
    }
    function Get-SumIfsFormula([string]$col, [string]$critCol, [string]$crit, [int]$year) {
        return "=ROUND(SUMIFS(Dados!`$${col}:`$${col},Dados!`$${critCol}:`$${critCol},$crit,Dados!`$E:`$E,"">=""&DATE($year,1,1),Dados!`$E:`$E,""<""&DATE($($year + 1),1,1))/1000,0)"
    }

    # ---- tabelas (R$ mil, salvo indicação)
    $deptChannelBody = @($deptOrder | ForEach-Object { ,@(@($depts[$_].Name) + (Get-DemoParts (Get-DemoSum 'sales' $allS @($_) $y25) @((Get-DemoSum 'b2c' $allS @($_) $y25), 0))) })
    $storeYearMiBody = @($storeOrder | ForEach-Object { ,@($stores[$_].Name, [Math]::Round((Get-DemoSum 'sales' @($_) $allD $y24) / 1e6, 1), [Math]::Round((Get-DemoSum 'sales' @($_) $allD $y25) / 1e6, 1)) })
    $storeYearRsBody = @($storeOrder | ForEach-Object { ,@($stores[$_].Name, (Get-DemoSum 'sales' @($_) $allD $y24), (Get-DemoSum 'sales' @($_) $allD $y25)) })
    $regionChannelBody = @($regions | ForEach-Object { $ss = @(Get-StoresOf $_); ,@(@([string]$_) + (Get-DemoParts (Get-DemoSum 'sales' $ss $allD $y25) @((Get-DemoSum 'b2c' $ss $allD $y25), 0))) })
    $storeChannelBody = @($storeOrder | ForEach-Object { ,@(@($stores[$_].Name) + (Get-DemoParts (Get-DemoSum 'sales' @($_) $allD $y25) @((Get-DemoSum 'b2c' @($_) $allD $y25), 0))) })
    $store1MonthBody = @($y25 | ForEach-Object { ,@($months[$_], (Get-DemoK (Get-DemoSum 'sales' @($s1) $allD @($_))), (Get-DemoK (Get-DemoSum 'cost' @($s1) $allD @($_)))) })
    $comboBody = @($y25 | ForEach-Object { $v = Get-DemoSum 'sales' @($s1) $allD @($_); $c = Get-DemoSum 'cost' @($s1) $allD @($_)
                                           ,@(@($months[$_]) + (Get-DemoParts $v @((Get-DemoSum 'b2c' @($s1) $allD @($_)), 0)) + @([Math]::Round(($v - $c) / $v, 3))) })
    $receita = Get-DemoSum 'sales' $allS $allD $y25; $cmv = Get-DemoSum 'cost' $allS $allD $y25
    $despesas = 0.21 * $receita; $ebitda = $receita - $cmv - $despesas; $deprec = 0.025 * $receita; $impostos = 0.34 * ($ebitda - $deprec)
    $wfBody = @(@('Receita bruta', (Get-DemoK $receita)), @('Custo das mercadorias', - (Get-DemoK $cmv)), @('Despesas operacionais', - (Get-DemoK $despesas)),
                @('EBITDA', 'e'), @('Depreciação e juros', - (Get-DemoK $deprec)), @('Impostos', - (Get-DemoK $impostos)), @('Lucro líquido', 'e'))
    # área (E10/E11): trimestres de 2024 e 2025 (12 meses de rótulos de 5 dígitos colidem nas bordas); o maior departamento por cima, longe da média
    $areaDepts = @($top3[2], $top3[1], $top3[0])
    $quarterDeptBody = @(for ($q = 0; $q -lt 8; $q++) { $qm = @((3 * $q)..(3 * $q + 2)); ,@(@("T$($q % 4 + 1)/$(24 + [Math]::Floor($q / 4))") + @($areaDepts | ForEach-Object { Get-DemoK (Get-DemoSum 'sales' $allS @($_) $qm) })) })
    $regionGroupBody = @($regions | ForEach-Object { $ss = @(Get-StoresOf $_)
                                                     ,@(@([string]$_) + (Get-DemoParts (Get-DemoSum 'sales' $ss $allD $y25) @($groups | ForEach-Object { Get-DemoSum 'sales' $ss (Get-DeptsOf $_) $y25 }))) })
    $mekkoUnitsBody = @($regions | ForEach-Object { $ss = @(Get-StoresOf $_); $n = $ss.Count
                                                    ,@(@($_, $n) + @($groups | ForEach-Object { [Math]::Round((Get-DemoSum 'sales' $ss (Get-DeptsOf $_) $y25) / 1000 / $n, 0) })) })
    $deptSalesBody = @($deptOrder | ForEach-Object { ,@($depts[$_].Name, (Get-DemoK (Get-DemoSum 'sales' $allS @($_) $y25))) })
    # pizza e rosca: vendas por região (5 fatias; a paleta tem 6 cores distintas e, com 8 departamentos, as cores se repetiriam)
    $pieBody = @($regions | ForEach-Object { ,@($_, (Get-DemoK (Get-DemoSum 'sales' @(Get-StoresOf $_) $allD $y25))) })
    $scatterBody = @($storeOrder | ForEach-Object { $v = Get-DemoSum 'sales' @($_) $allD $y25; $c = Get-DemoSum 'cost' @($_) $allD $y25; $rc = Get-DemoSum 'receipts' @($_) $allD $y25
                                                    ,@($stores[$_].Name, [Math]::Round($v / $rc, 0), [Math]::Round(100 * ($v - $c) / $v, 1)) })
    $bubbleBody = @($deptOrder | ForEach-Object { $v = Get-DemoSum 'sales' $allS @($_) $y25; $c = Get-DemoSum 'cost' $allS @($_) $y25; $rc = Get-DemoSum 'receipts' $allS @($_) $y25
                                                  ,@($depts[$_].Name, [Math]::Round($v / $rc, 0), [Math]::Round(100 * ($v - $c) / $v, 1), (Get-DemoK $v), $depts[$_].Group) })
    # Arredondar intervalo (E23): participação dos departamentos nas vendas de 2025 de uma loja; a primeira loja (pela ordem de vendas) cujo
    # arredondamento simples a 0% NÃO soma 100% mostra a diferença (determinístico; sem nenhuma, a maior loja)
    $roundStore = $s1
    foreach ($si in $storeOrder) {
        $tot = Get-DemoSum 'sales' @($si) $allD $y25
        $naive = 0.0; foreach ($di in $deptOrder) { $naive += [Math]::Round((Get-DemoSum 'sales' @($si) @($di) $y25) / $tot, 2, [MidpointRounding]::AwayFromZero) }
        if ([Math]::Abs($naive - 1) -gt 0.005) { $roundStore = $si; break }
    }
    $roundTot = Get-DemoSum 'sales' @($roundStore) $allD $y25
    $roundBody = @($deptOrder | ForEach-Object { ,@($depts[$_].Name, ((Get-DemoSum 'sales' @($roundStore) @($_) $y25) / $roundTot)) })

    # ---- um exemplo por item da cobertura
    $prev = $wsO
    $cov = @(Get-TqqDemoCoverage)
    foreach ($item in $cov) {
        $unit = 'R$ mil'
        switch ($item.Num) { 2 { $unit = 'R$ milhões' } 7 { $unit = 'R$ na tabela, rótulos em R$ mil' } { $_ -in 15, 16 } { $unit = '' }
                             20 { $unit = 'ticket em R$, margem bruta em %' } 21 { $unit = 'ticket em R$, margem bruta em %, vendas em R$ mil' } 23 { $unit = 'R$ mil e participações' } }
        $ws = New-ExampleSheet $prev $item $unit
        $prev = $ws
        switch ($item.Num) {
            1  { $rng = Write-Table $ws 'A6' @('Departamento', 'B2C', 'B2B') $deptChannelBody
                 for ($i = 0; $i -lt $deptChannelBody.Count; $i++) {
                     $ws.Cells.Item(7 + $i, 2).Formula = Get-SumIfsFormula 'K' 'D' "`$A$(7 + $i)" 2025
                     $ws.Cells.Item(7 + $i, 3).Formula = (Get-SumIfsFormula 'F' 'D' "`$A$(7 + $i)" 2025) + "-`$B$(7 + $i)"     # B2B = total arredondado - B2C
                 }
                 $co = Add-Chart $ws $rng 'stacked'; Set-ChartBox $ws $co 620 320 }
            2  { $rng = Write-Table $ws 'A6' @('Loja', '2024', '2025') $storeYearMiBody '#,##0.0'; $co = Add-Chart $ws $rng 'clustered'
                 Invoke-CoreQuiet 'E02' 'Core_SetOption' 'gap' '120' | Out-Null
                 Invoke-CoreQuiet 'E02' 'Core_SetOption' 'numfmt' 'dec1' | Out-Null      # a tabela tem 1 casa: no padrão (inteiro) 8,7 e 9,4 viravam 9 e 9
                 Set-ChartBox $ws $co 620 300 }
            3  { $rng = Write-Table $ws 'A6' @('Região', 'B2C', 'B2B') $regionChannelBody; $co = Add-Chart $ws $rng 'stacked100' }   # o indicador 100% nasce ligado nos 100%
            4  { $rng = Write-Table $ws 'A6' @('Mês', 'Vendas', 'Custo') $store1MonthBody; Set-MonthFormat $ws.Range('A7:A18')
                 $co = Add-Chart $ws $rng 'line'; Invoke-CoreQuiet 'E04' 'Core_SetOption' 'legend' 'lineend' | Out-Null }
            5  { $rng = Write-Table $ws 'A6' @('Etapa', 'R$ mil') $wfBody; $co = Add-Chart $ws $rng 'waterfall'; Set-ChartBox $ws $co 560 300 }
            6  { $rng = Write-Table $ws 'A6' @('Região', 'B2C', 'B2B') $regionChannelBody; $co = Add-Chart $ws $rng 'barstacked'
                 Invoke-CoreQuiet 'E06' 'Core_AddAnnotation' 'diff' 1 $regions.Count | Out-Null }
            7  { $rng = Write-Table $ws 'A6' @('Loja', '2024', '2025') $storeYearRsBody; $co = Add-Chart $ws $rng 'barclustered'
                 Invoke-CoreQuiet 'E07' 'Core_SetOption' 'numfmt' 'mil' | Out-Null; Set-ChartBox $ws $co 520 360 }
            8  { $rng = Write-Table $ws 'A6' @('Loja', 'B2C', 'B2B') $storeChannelBody; $co = Add-Chart $ws $rng 'barstacked100'; Set-ChartBox $ws $co 520 320 }
            9  { $rng = Write-Table $ws 'A6' @('Etapa', 'R$ mil') $wfBody; $co = Add-Chart $ws $rng 'barwaterfall'; Set-ChartBox $ws $co 520 300 }
            10 { $rng = Write-Table $ws 'A6' (@('Trimestre') + @($areaDepts | ForEach-Object { $depts[$_].Name })) $quarterDeptBody
                 $co = Add-Chart $ws $rng 'area'; Invoke-CoreQuiet 'E10' 'Core_SetToggle' 'vline' $true | Out-Null; Set-ChartBox $ws $co 560 300 }
            11 { $rng = Write-Table $ws 'A6' (@('Trimestre') + @($areaDepts | ForEach-Object { $depts[$_].Name })) $quarterDeptBody
                 $co = Add-Chart $ws $rng 'area100'; Set-ChartBox $ws $co 560 300 }
            12 { $rng = Write-Table $ws 'A6' @('Mês', 'B2C', 'B2B', 'Margem %') $comboBody; Set-MonthFormat $ws.Range('A7:A18'); Set-NumFmt ($ws.Range('D7:D18')) '0.0%'
                 $co = Add-Chart $ws $rng 'combo'
                 if (-not (Invoke-Core 'Core_GetToggle' 'secondary')) { throw 'E12: o eixo secundário não ligou sozinho' }
                 # sem os totais das colunas: a linha da margem (eixo secundário, perto do topo) cairia sobre eles
                 Invoke-CoreQuiet 'E12' 'Core_SetToggle' 'totals' $false | Out-Null
                 Set-ChartBox $ws $co 600 320 }
            13 { $rng = Write-Table $ws 'A6' (@('Região') + $groups) $regionGroupBody; $co = Add-Chart $ws $rng 'mekko'; Set-ChartBox $ws $co 560 320 }
            14 { $rng = Write-Table $ws 'A6' (@('Região', 'Lojas') + $groups) $mekkoUnitsBody; Set-NumFmt ($ws.Range('B7:B11')) '0'; $co = Add-Chart $ws $rng 'mekkounits'; Set-ChartBox $ws $co 560 320 }
            15 { $rng = Write-Table $ws 'A6' @('Obra', 'Início', 'Fim') @($obrasQ | ForEach-Object { ,@($_[0], $_[1].ToOADate(), $(if ($_[2]) { $_[2].ToOADate() } else { $null })) }) 'dd/mm/yyyy'
                 $ws.Columns('A').ColumnWidth = 32
                 $co = Add-Chart $ws $rng 'gantt'; Invoke-CoreQuiet 'E15' 'Core_AddAnnotation' 'bracket' 1 2 | Out-Null; Set-ChartBox $ws $co 600 260 }
            16 { $rng = Write-Table $ws 'A6' @('Etapa', 'Início', 'Fim') @($obrasY | ForEach-Object { ,@($_[0], $_[1].ToOADate(), $(if ($_[2]) { $_[2].ToOADate() } else { $null })) }) 'dd/mm/yyyy'
                 $co = Add-Chart $ws $rng 'gantt'; Set-ChartBox $ws $co 600 260 }
            17 { $rng = Write-Table $ws 'A6' @('Região', 'Vendas 2025') $pieBody; $co = Add-Chart $ws $rng 'pie'
                 Invoke-CoreQuiet 'E17' 'Core_SetToggle' 'pieout' $true | Out-Null; Set-ChartBox $ws $co 560 340 }
            18 { $rng = Write-Table $ws 'A6' @('Região', 'Vendas 2025') $pieBody; $co = Add-Chart $ws $rng 'doughnut'
                 Invoke-CoreQuiet 'E18' 'Core_SetOption' 'palette' 'cinza' | Out-Null; Set-ChartBox $ws $co 520 320 }
            19 { $rng = Write-Table $ws 'A6' @('Loja', 'B2C', 'B2B') $storeChannelBody; $co = Add-Chart $ws $rng 'butterfly'; Set-ChartBox $ws $co 640 320 }
            20 { $rng = Write-Table $ws 'A6' @('Loja', 'Ticket médio (R$)', 'Margem bruta (%)') $scatterBody '#,##0'; Set-NumFmt ($ws.Range('C7:C14')) '0.0'
                 $co = Add-Chart $ws $rng 'scatter'
                 Invoke-CoreQuiet 'E20' 'Core_SetToggle' 'trend' $true | Out-Null
                 Invoke-CoreQuiet 'E20' 'Core_SetToggle' 'quad' $true | Out-Null
                 Invoke-CoreQuiet 'E20' 'Core_SetOption' 'legend' 'none' | Out-Null
                 Set-ChartBox $ws $co 560 340 }
            21 { $rng = Write-Table $ws 'A6' @('Departamento', 'Ticket médio (R$)', 'Margem bruta (%)', 'Vendas 2025', 'Grupo') $bubbleBody '#,##0'; Set-NumFmt ($ws.Range('C7:C14')) '0.0'
                 $ws.Columns('E').ColumnWidth = 13
                 $co = Add-Chart $ws $rng 'bubble'; Set-ChartBox $ws $co 600 360 }
            22 { # valores (não SOMASES): a seta CAGR é uma forma posicionada pelo ThinQQ e não acompanharia uma mudança em Dados sem ele
                 $cagrBody = @((@(2024) + @($top3 | ForEach-Object { Get-DemoK (Get-DemoSum 'sales' $allS @($_) $y24) })), (@(2025) + @($top3 | ForEach-Object { Get-DemoK (Get-DemoSum 'sales' $allS @($_) $y25) })))
                 $rng = Write-Table $ws 'A6' (@('Ano') + @($top3 | ForEach-Object { $depts[$_].Name })) $cagrBody
                 Set-NumFmt ($ws.Range('A7:A8')) '0'
                 $co = Add-Chart $ws $rng 'stacked'; Invoke-CoreQuiet 'E22' 'Core_AddAnnotation' 'cagr' 1 2 | Out-Null }
            23 { $rng = Write-Table $ws 'A6' (@('Região') + $groups) $regionGroupBody; $co = Add-Chart $ws $rng 'stacked100'
                 Invoke-CoreQuiet 'E23' 'Core_SetToggle' 'roundsum' $true | Out-Null
                 $ws.Range('A14').Value2 = "Arredondar intervalo: loja $($stores[$roundStore].Name), vendas 2025"
                 $ws.Range('A14').Font.Bold = $true
                 $rt = Write-Table $ws 'A15' @('Departamento', 'Exata', 'Arredondamento simples', 'Soma consistente') @($roundBody | ForEach-Object { ,@($_[0], $_[1]) }) ''
                 $n = $roundBody.Count
                 Set-NumFmt ($ws.Range("B16:B$(15 + $n)")) '0.000%'
                 for ($i = 0; $i -lt $n; $i++) { $ws.Cells.Item(16 + $i, 3).Formula = "=ROUND(B$(16 + $i),2)" }
                 Set-NumFmt ($ws.Range("C16:C$(15 + $n)")) '0%'
                 Invoke-CoreQuiet 'E23 Round_WriteFormulas' 'Round_WriteFormulas' $ws.Range("B16:B$(15 + $n)") $ws.Range('D16') 2 | Out-Null
                 $tr = 16 + $n
                 $ws.Cells.Item($tr, 1).Value2 = 'Total'
                 foreach ($c in 2, 3, 4) { $col = [char](64 + $c); $ws.Cells.Item($tr, $c).Formula = "=SUM(${col}16:${col}$($tr - 1))" }
                 Set-NumFmt ($ws.Range("B${tr}")) '0.000%'; Set-NumFmt ($ws.Range("C${tr}:D${tr}")) '0%'
                 $ws.Range("A${tr}:D${tr}").Font.Bold = $true
                 $ws.Range("A${tr}:D${tr}").Borders.Item(8).LineStyle = 1        # xlEdgeTop
                 $ws.Range('C15:D15').WrapText = $true; $ws.Columns('C:D').ColumnWidth = 16
                 $ws.Rows(15).AutoFit() | Out-Null }                                 # cabeçalho em 2 linhas inteiro
            24 { $rng = Write-Table $ws 'A6' @('Departamento', 'B2C', 'B2B') $deptChannelBody; $co = Add-Chart $ws $rng 'stacked'
                 # corte DENTRO do segmento de cima da coluna cortada (B2C + metade do B2B do cimento): os dois segmentos aparecem e o B2B mantém o
                 # rótulo; um corte abaixo do topo do B2C (ex.: 1,2 x o 2º maior total) esconderia o B2B inteiro e a coluna pareceria só B2C
                 $cut = $deptChannelBody[0]; $second = [double]$deptChannelBody[1][1] + [double]$deptChannelBody[1][2]
                 $limit = [Math]::Ceiling(([double]$cut[1] + 0.5 * [double]$cut[2]) / 1000) * 1000
                 if ($limit -le $second -or $limit -ge [double]$cut[1] + [double]$cut[2]) { throw "E24: corte $limit fora de (2º maior total $second, total do cimento)" }
                 Invoke-CoreQuiet 'E24 corte de eixo' 'Core_SetAxisBreakAt' ([double]$limit) | Out-Null
                 Set-ChartBox $ws $co 620 320 }
            25 { $big = $stores[$s1].Name; $small = $stores[$sLast].Name
                 $b1 = @(@(2024, (Get-DemoK (Get-DemoSum 'sales' @($s1) $allD $y24)), (Get-DemoK (Get-DemoSum 'cost' @($s1) $allD $y24))), @(2025, (Get-DemoK (Get-DemoSum 'sales' @($s1) $allD $y25)), (Get-DemoK (Get-DemoSum 'cost' @($s1) $allD $y25))))
                 $b2 = @(@(2024, (Get-DemoK (Get-DemoSum 'sales' @($sLast) $allD $y24)), (Get-DemoK (Get-DemoSum 'cost' @($sLast) $allD $y24))), @(2025, (Get-DemoK (Get-DemoSum 'sales' @($sLast) $allD $y25)), (Get-DemoK (Get-DemoSum 'cost' @($sLast) $allD $y25))))
                 $rng = Write-Table $ws 'A6' @($big, 'Vendas', 'Custo') $b1; Set-NumFmt ($ws.Range('A7:A8')) '0'
                 $co1 = Add-Chart $ws $rng 'clustered'
                 $rng2 = Write-Table $ws 'A11' @($small, 'Vendas', 'Custo') $b2; Set-NumFmt ($ws.Range('A12:A13')) '0'
                 $co2 = Add-Chart $ws $rng2 'clustered'
                 Set-ChartBox $ws $co1 330 250
                 Set-ChartBox $ws $co2 330 250 ($co1.Left + $co1.Width + 24) $co1.Top
                 $ws.Activate(); $ws.Range('A1').Select() | Out-Null
                 Invoke-Core 'Core_ClearLastNotice' | Out-Null
                 Invoke-Core 'Core_SameScale' | Out-Null
                 $n = Invoke-Core 'Core_LastNotice'
                 if ($n -notmatch 'aplicada a 2 ') { throw "E25 mesma escala: $n" }
                 # nome da loja acima de cada gráfico (linha 5 fica livre: a tabela começa na 6 e não encosta nela)
                 foreach ($pair in @(@($co1, $big), @($co2, $small))) { $cell = $pair[0].TopLeftCell; $lbl = $ws.Cells.Item($cell.Row - 1, $cell.Column); $lbl.Value2 = "$($pair[1]) (R$ mil)"; $lbl.Font.Bold = $true } }
            26 { $rng = Write-Table $ws 'A6' @('Mês', 'Vendas') @($store1MonthBody | ForEach-Object { ,@($_[0], $_[1]) }); Set-MonthFormat $ws.Range('A7:A18')
                 $co = Add-Chart $ws $rng 'line'; Invoke-CoreQuiet 'E26' 'Core_SetOption' 'legend' 'none' | Out-Null }
        }
        Invoke-Core 'Core_ClearLastNotice' | Out-Null
    }

    # ---- índice
    $idx.Activate()
    $idx.Tab.Color = 0x79531F
    $idx.Range('A1').Value2 = 'ThinQQ — pasta de demonstração'
    $idx.Range('A1').Font.Bold = $true; $idx.Range('A1').Font.Size = 16
    $idx.Range('A2').Value2 = "${company}: $nS lojas em $($regions.Count) regiões. Nenhuma empresa ou pessoa real; lojas, cidades e números são inventados (gerados por script)."
    $idx.Range('A3').Value2 = 'Clique no nome de um exemplo para ir ao gráfico; em cada exemplo, a célula A1 ("Voltar ao índice") traz você de volta. Os dados estão nas planilhas Dados, Lojas e Obras.'
    $idx.Range('A4').Value2 = 'Os gráficos são gráficos nativos do Excel e continuam ligados às tabelas: funcionam com ou sem o ThinQQ instalado. Setas, colchetes, marcas de corte e a escala do Gantt são recalculados só com o ThinQQ instalado; sem ele ficam como estão se os números mudarem.'
    $idx.Range('A2:A4').Font.Color = 0x404040
    foreach ($rr in 2, 3, 4) { $m = $idx.Range("A${rr}:E${rr}"); $m.Merge() | Out-Null; $m.WrapText = $true; $m.VerticalAlignment = -4160 }   # textos dentro da largura da tabela
    $idx.Rows(4).RowHeight = 30                                                  # células mescladas não se ajustam sozinhas: A4 tem 2 linhas
    $hdrRow = 6
    $idx.Range("A${hdrRow}:E${hdrRow}").Value2 = [object[]]('Nº', 'Exemplo (clique para abrir)', 'Tipo de gráfico', 'Recursos em destaque', 'Como criar no ThinQQ')
    $h = $idx.Range("A${hdrRow}:E${hdrRow}"); $h.Font.Bold = $true; $h.Font.Color = 0xFFFFFF; $h.Interior.Color = 0x79531F
    $featNames = @{ totals = 'Totais'; labels = 'Rótulos'; 'legend-none' = 'Legenda: nenhuma'; lineend = 'Legenda: nome no fim da linha'; numfmt = 'Formato (milhares)'
                    gap = 'Espaço (muito largo)'; palette = 'Paleta (tons de cinza)'; cagr = 'Seta CAGR'; diff = 'Seta diferença'; bracket = 'Colchete'
                    vline = 'Linha de valor (média)'; ind100 = 'Indicador 100%'; samescale = 'Mesma escala'; secondary = 'Eixo secundário'; pieout = 'Rótulos fora da pizza'
                    'gantt-weeks' = 'Escala em semanas'; 'gantt-months' = 'Escala em meses'; roundsum = 'Rótulos com soma consistente'; roundrange = 'Arredondar intervalo'
                    break = 'Corte de eixo'; trend = 'Linha de tendência'; quad = 'Quadrantes (médias)'; connectors = 'Conectores' }
    $r = $hdrRow + 1
    foreach ($item in $cov) {
        $idx.Cells.Item($r, 1).Value2 = $item.Num
        $idx.Hyperlinks.Add($idx.Cells.Item($r, 2), '', "'$($item.Sheet)'!A1", "Ir para o exemplo $($item.Num) (planilha $($item.Sheet))", $item.Title) | Out-Null
        $idx.Cells.Item($r, 3).Value2 = $item.KindName
        $fl = @($item.Features | ForEach-Object { $featNames[$_] })
        $idx.Cells.Item($r, 4).Value2 = $(if ($fl.Count -gt 0) { $fl -join ', ' } else { '—' })
        $idx.Cells.Item($r, 5).Value2 = $item.Menu
        if (($r - $hdrRow) % 2 -eq 0) { $idx.Range("A${r}:E${r}").Interior.Color = 0xF7F2EC }
        $r++
    }
    $last = $r - 1
    # ~175 caracteres de largura: cabe numa tela de 1366 px; títulos, recursos e menus longos quebram a linha
    $idx.Columns('A').ColumnWidth = 5; $idx.Columns('B').ColumnWidth = 60; $idx.Columns('C').ColumnWidth = 22; $idx.Columns('D').ColumnWidth = 30; $idx.Columns('E').ColumnWidth = 58
    $idx.Range("A${hdrRow}:A${last}").HorizontalAlignment = -4108                   # xlCenter
    $idx.Range("A${hdrRow}:E${last}").VerticalAlignment = -4160                     # xlTop
    $idx.Range("B$($hdrRow + 1):E${last}").WrapText = $true
    $idx.Range("A$($hdrRow + 1):E${last}").Rows.AutoFit() | Out-Null
    $idx.Range("A${hdrRow}:E${last}").Borders.Item(12).LineStyle = 1               # xlInsideHorizontal
    $idx.Range("A${hdrRow}:E${last}").Borders.Item(12).Color = 0xD9D9D9
    $idx.Range("A${hdrRow}:E${last}").Borders.Item(9).LineStyle = 1                # xlEdgeBottom
    $idx.Range("A${hdrRow}:E${last}").Borders.Item(9).Color = 0xD9D9D9

    # ---- acabamento: cada planilha abre em A1, índice sem linhas de grade e congelado no cabeçalho, pasta abre no índice
    foreach ($ws in @($wb.Worksheets)) {
        if ($ws.Visible -ne -1) { continue }
        $ws.Activate(); $ws.Range('A1').Select() | Out-Null
        $xl.ActiveWindow.ScrollRow = 1; $xl.ActiveWindow.ScrollColumn = 1
        if ($ws.Name -eq 'Dados') { $xl.ActiveWindow.SplitRow = 1; $xl.ActiveWindow.SplitColumn = 0; $xl.ActiveWindow.FreezePanes = $true }
        if ($ws.Name -in 'Dados', 'Lojas', 'Obras') { $ws.Tab.Color = 0x808080 }
    }
    $idx.Activate(); $xl.ActiveWindow.DisplayGridlines = $false
    $xl.ActiveWindow.SplitRow = $hdrRow; $xl.ActiveWindow.SplitColumn = 0; $xl.ActiveWindow.FreezePanes = $true
    $idx.Range('B7').Select() | Out-Null
    Set-DocProperty $wb 'Title' 'ThinQQ - pasta de demonstração (dados fictícios)'
    Set-DocProperty $wb 'Subject' 'Um exemplo de gráfico por funcionalidade do ThinQQ'
    $wb.RemovePersonalInformation = $true             # docProps sem autor/empresa (nunca altere $xl.UserName: grava no registro do Office do usuário)
    $wb.SaveAs($tmpOut, 51)
    $wb.Close($false)
    $core.Close($false)
    if (Test-Path $OutFile) { Remove-Item $OutFile -Force }
    Move-Item $tmpOut $OutFile
    Write-Host ("Gerado: $OutFile ($((Get-Item $OutFile).Length) bytes) em {0:N1} s" -f $sw.Elapsed.TotalSeconds)
    $exitCode = 0
} catch {
    Write-Host "FALHOU: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace
    if ($wd -and (Test-ExcelWatchdogFired $wd)) { Write-ExcelTimeout $wd }
} finally {
    Stop-ExcelWatchdog $wd
    Close-ExcelApp $xl
    if (Test-Path $tmpOut) { Remove-Item $tmpOut -Force -ErrorAction SilentlyContinue }
}
exit $exitCode
