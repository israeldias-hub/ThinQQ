# Funções compartilhadas de build (carregar com dot-source).
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$script:DistDir  = Join-Path $script:RepoRoot 'dist'
$script:PptPid   = @()   # PID do PowerPoint criado por New-PptApp (encerrado por Close-PptApp)
. "$PSScriptRoot\Ribbon-Icons.ps1"

function Assert-VbomTrusted {
    $key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    $v = (Get-ItemProperty $key -ErrorAction SilentlyContinue).AccessVBOM
    if ($v -ne 1) {
        throw "Habilite 'Confiar no acesso ao modelo de objeto do projeto VBA' (Arquivo > Opções > Central de Confiabilidade > Configurações de Macro) ou execute: Set-ItemProperty '$key' AccessVBOM 1"
    }
}

function Test-PptVbomTrusted {
    $v = (Get-ItemProperty 'HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Security' -ErrorAction SilentlyContinue).AccessVBOM
    return ($v -eq 1)
}
function Assert-PptVbomTrusted {
    if (-not (Test-PptVbomTrusted)) {
        throw "Habilite no POWERPOINT 'Confiar no acesso ao modelo de objeto do projeto VBA' (Arquivo > Opções > Central de Confiabilidade > Configurações de Macro). Sem isso o suplemento ThinQQ.ppam não pode ser montado."
    }
}
# PowerPoint e de instancia unica: so cria se NENHUM POWERPNT estiver em execucao (senao CreateObject anexaria ao do usuario).
# O PID criado fica em $script:PptPid; o processo tem que ser de automacao (linha de comando com -Embedding), senao nao e nosso.
function New-PptApp {
    if (Get-Process POWERPNT -ErrorAction SilentlyContinue) { throw 'Há um PowerPoint aberto. Feche-o para montar/testar o suplemento (o PowerPoint é de instância única).' }
    $script:PptPid = @()
    $app = New-Object -ComObject PowerPoint.Application
    $ids = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
    $cmd = ''
    if ($ids.Count -eq 1) { try { $cmd = [string](Get-CimInstance Win32_Process -Filter "ProcessId = $($ids[0])" -ErrorAction Stop).CommandLine } catch { $cmd = '' } }
    if ($ids.Count -ne 1 -or $cmd -notmatch 'Embedding') {
        # Nao chama Quit nem encerra: pode ser um PowerPoint do usuario aberto no mesmo instante.
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { }
        throw "PowerPoint do build não identificado ($($ids.Count) processos; linha de comando '$cmd'); nada foi encerrado."
    }
    $script:PptPid = $ids
    $app.DisplayAlerts = 1          # ppAlertsNone
    return $app
}
# Fecha o PowerPoint criado por New-PptApp. Fecha as apresentacoes que o script abriu ($Owned) e as que ficaram SEM janela (o PowerPoint do
# script e invisivel: so codigo cria apresentacao sem janela). Se sobrar apresentacao COM janela ou janela do Modo de Exibicao Protegido (o
# usuario abriu um arquivo nesta instancia: o PowerPoint e de instancia unica), nao da Quit nem encerra: avisa e so solta o objeto. Senao
# Quit e, se o processo (PID registrado) nao terminar em $WaitSec s, encerra esse PID (tambem quando o PowerPoint ja nao responde).
function Close-PptApp($app, [object[]]$Owned = @(), [int]$WaitSec = 15) {
    $rcw = [System.Runtime.InteropServices.Marshal]
    $foreign = 0
    if ($null -ne $app) {
        foreach ($o in @($Owned)) {
            if ($null -eq $o) { continue }
            try { $o.Saved = -1; $o.Close() } catch { }
            try { [void]$rcw::ReleaseComObject($o) } catch { }
        }
        $o = $null
        # Pelo indice, soltando cada objeto (enumerar a colecao com foreach deixa referencias que impedem o Quit: medido).
        $col = $null; $p = $null; $responds = $true
        try {
            $col = $app.Presentations
            for ($i = [int]$col.Count; $i -ge 1; $i--) {
                $p = $col.Item($i)
                try { if ([int]$p.Windows.Count -eq 0) { $p.Saved = -1; $p.Close() } else { $foreign++ } } finally { try { [void]$rcw::ReleaseComObject($p) } catch { }; $p = $null }
            }
            $pv = $app.ProtectedViewWindows; $foreign += [int]$pv.Count; try { [void]$rcw::ReleaseComObject($pv) } catch { }; $pv = $null
        } catch { $responds = $false } finally { if ($null -ne $col) { try { [void]$rcw::ReleaseComObject($col) } catch { } }; $col = $null }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        if ($foreign -gt 0) {
            Write-Host "AVISO: o PowerPoint (PID $(@($script:PptPid) -join ',')) tem $foreign apresentação(ões)/janela(s) que não são deste script: não foi fechado nem encerrado." -ForegroundColor Yellow
            try { [void]$rcw::ReleaseComObject($app) } catch { }
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            $script:PptPid = @()
            return
        }
        if ($responds) { try { $app.Quit() } catch { } }
        try { [void]$rcw::ReleaseComObject($app) } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
    foreach ($id in @($script:PptPid)) {
        if (-not $id) { continue }
        $proc = Get-Process -Id $id -ErrorAction SilentlyContinue
        # Dispose: um handle aberto deixa o processo que ja saiu ainda listado por Get-Process (contagem falsa logo depois)
        if ($proc) { try { [void]$proc.WaitForExit($WaitSec * 1000) } catch { } finally { $proc.Dispose(); $proc = $null } }
        if (Get-Process -Id $id -ErrorAction SilentlyContinue) {
            Write-Host "PowerPoint (PID $id) não fechou em $WaitSec s: encerrado pelo PID." -ForegroundColor Yellow
            Stop-Process -Id $id -Force -ErrorAction SilentlyContinue
            for ($i = 0; $i -lt 50 -and (Get-Process -Id $id -ErrorAction SilentlyContinue); $i++) { Start-Sleep -Milliseconds 200 }
        }
    }
    $script:PptPid = @()
}

function New-ExcelApp {
    # Snapshot dos Excel já abertos: o app novo tem que ser um processo NOVO (nunca matar/fechar o do usuário).
    $before = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
    $xl = New-Object -ComObject Excel.Application
    $xlPid = Get-ExcelPid $xl
    if ($xlPid -le 0 -or $before -contains $xlPid) {
        # Não chama Quit: pode ser uma instância do usuário.
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
        throw "Não foi possível garantir que o Excel é uma instância nova (PID $xlPid já existia ou é inválido); abortando para não afetar o seu Excel."
    }
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    return $xl
}

function Close-ExcelApp($xl) {
    if ($null -ne $xl) {
        try { $xl.Quit() } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } catch { }
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

function Write-BuildInfoModule([string]$Version, [string]$Configuration, [string]$OutPath) {
    $suffix = if ($Configuration -eq 'Test') { '.test' } else { '' }
    $content = @"
Attribute VB_Name = "modBuildInfo"
' Gerado pelo build - nao editar
Option Explicit
Public Const APP_VERSION As String = "$Version"
Public Const BUILD_CONFIG As String = "$Configuration"
Public Const BUILD_SUFFIX As String = "$suffix"
"@
    Set-Content -Path $OutPath -Value $content -Encoding ascii
}

# O VBA importa .bas/.cls em ANSI (Windows-1252) com CRLF; os fontes no Git ficam em UTF-8.
# Converte para um arquivo temporário antes do Import.
function Import-VbaFile($vbproj, [string]$Path) {
    $text = Get-Content -Path $Path -Raw -Encoding utf8
    $text = ($text -replace "`r`n", "`n") -replace "`n", "`r`n"
    $tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) 'thinqq_import'
    New-Item -ItemType Directory -Force $tmpDir | Out-Null
    $tmp = Join-Path $tmpDir ([System.IO.Path]::GetFileName($Path))
    [System.IO.File]::WriteAllText($tmp, $text, [System.Text.Encoding]::GetEncoding(1252))
    try { [void]$vbproj.VBComponents.Import($tmp) } finally { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
}

function Read-VbaCode([string]$CodePath) {
    $t = Get-Content -Path $CodePath -Raw -Encoding utf8
    return ($t -replace "`r`n", "`n") -replace "`n", "`r`n"
}

function Add-CodeToComponent($vbproj, [string]$ComponentName, [string]$CodePath) {
    $cm = $vbproj.VBComponents.Item($ComponentName).CodeModule
    $cm.AddFromString((Read-VbaCode $CodePath))
}

# O PowerShell devolve $null em $ctl.Font de controles MSForms; InvokeMember contorna.
function Set-FontProp($ctl, [string]$Prop, $Value) {
    $bf = [System.Reflection.BindingFlags]
    $f = [System.__ComObject].InvokeMember('Font', $bf::GetProperty, $null, $ctl, $null)
    [void][System.__ComObject].InvokeMember($Prop, $bf::SetProperty, $null, $f, @($Value))
}

# $wb.BuiltinDocumentProperties vem $null no PowerShell; InvokeMember contorna.
function Set-DocProperty($wb, [string]$Name, [string]$Value) {
    $bf = [System.Reflection.BindingFlags]
    $props = [System.__ComObject].InvokeMember('BuiltinDocumentProperties', $bf::GetProperty, $null, $wb, $null)
    $prop = [System.__ComObject].InvokeMember('Item', $bf::GetProperty, $null, $props, @($Name))
    [void][System.__ComObject].InvokeMember('Value', $bf::SetProperty, $null, $prop, @($Value))
}

# Cria um UserForm a partir de um JSON declarativo e injeta o code-behind (.vba).
function Add-FormFromJson($vbproj, [string]$JsonPath, [string]$CodePath) {
    $spec = Get-Content -Path $JsonPath -Raw -Encoding utf8 | ConvertFrom-Json
    $comp = $vbproj.VBComponents.Add(3)   # 3 = vbext_ct_MSForm
    $comp.Name = $spec.name
    $d = $comp.Designer
    $d.Caption = $spec.caption
    $comp.Properties.Item("Width").Value  = [double]$spec.width
    $comp.Properties.Item("Height").Value = [double]$spec.height
    foreach ($c in $spec.controls) {
        $ctl = $d.Controls.Add("Forms.$($c.type).1", $c.name, $true)
        foreach ($p in $c.PSObject.Properties) {
            switch ($p.Name) {
                'type'     { }
                'name'     { }
                'fontName' { Set-FontProp $ctl 'Name' ([string]$p.Value) }
                'fontSize' { Set-FontProp $ctl 'Size' ([double]$p.Value) }
                'fontBold' { Set-FontProp $ctl 'Bold' ([bool]$p.Value) }
                default    { $ctl.($p.Name) = $p.Value }
            }
        }
    }
    if ($CodePath -and (Test-Path $CodePath)) {
        $comp.CodeModule.AddFromString((Read-VbaCode $CodePath))
    }
}

# Insere customUI/customUI14.xml no pacote (zip), a relacao em _rels/.rels e, no modo Png, os icones proprios: customUI/images/<nome>.png
# para cada image= do XML, customUI/_rels/customUI14.xml.rels (Id = nome) e o content type png. Modo Mso: XML com imageMso de reserva, sem imagens.
# IconTempRoot: onde nasce (e e apagada) a pasta temporaria thinqq_icons_<guid> dos PNG; os testes passam uma pasta propria.
function Add-RibbonToPackage([string]$PackagePath, [string]$RibbonXmlPath, [ValidateSet('Png','Mso')][string]$Icons = 'Png', [string]$IconTempRoot = ([IO.Path]::GetTempPath())) {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $xmlText = [IO.File]::ReadAllText($RibbonXmlPath, (New-Object Text.UTF8Encoding($false)))
    # fonte: XML bem formado, nomes image= (so dos elementos; comentarios nao contam) na tabela e nenhum controle com image e imageMso
    $images = Assert-TqqRibbonXml -Xml $xmlText -Mode Png -What 'customUI14.xml'
    if ($Icons -eq 'Mso') {
        $xmlText = Convert-TqqRibbonToMso -Xml $xmlText
        [void](Assert-TqqRibbonXml -Xml $xmlText -Mode Mso -What 'customUI14.xml convertido para imageMso')   # relido como XML: nada de image= sobrando
        $images = @()
    }
    $iconDir = $null
    if ($images.Count -gt 0) { $iconDir = Join-Path $IconTempRoot ("thinqq_icons_" + [guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Force $iconDir | Out-Null; [void](Export-TqqIcons -OutDir $iconDir) }
    $zip = [System.IO.Compression.ZipFile]::Open($PackagePath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        function Set-ZipText($zip, [string]$name, [string]$text) {
            $old = $zip.GetEntry($name); if ($old) { $old.Delete() }
            $e = $zip.CreateEntry($name); $w = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false))); $w.Write($text); $w.Dispose()
        }
        if ($Icons -eq 'Mso') {
            Set-ZipText $zip 'customUI/customUI14.xml' $xmlText
        } else {
            $existing = $zip.GetEntry('customUI/customUI14.xml'); if ($existing) { $existing.Delete() }
            $entry = $zip.CreateEntry('customUI/customUI14.xml'); $bytes = [IO.File]::ReadAllBytes($RibbonXmlPath); $s = $entry.Open(); $s.Write($bytes, 0, $bytes.Length); $s.Dispose()
        }
        $relsEntry = $zip.GetEntry('_rels/.rels')
        $reader = New-Object System.IO.StreamReader($relsEntry.Open()); $rels = $reader.ReadToEnd(); $reader.Dispose()
        if ($rels -notmatch 'customUI/customUI14\.xml') {
            $rel = '<Relationship Id="rIdThinQQ" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
            Set-ZipText $zip '_rels/.rels' ($rels -replace '</Relationships>', "$rel</Relationships>")
        }
        foreach ($old in @($zip.Entries | Where-Object { $_.FullName -like 'customUI/images/*' -or $_.FullName -eq 'customUI/_rels/customUI14.xml.rels' })) { $old.Delete() }
        if ($images.Count -gt 0) {
            foreach ($img in $images) {
                $e = $zip.CreateEntry("customUI/images/$img.png"); $b = [IO.File]::ReadAllBytes((Join-Path $iconDir "$img.png")); $s = $e.Open(); $s.Write($b, 0, $b.Length); $s.Dispose()
            }
            $imgRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                (($images | ForEach-Object { '<Relationship Id="' + $_ + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/' + $_ + '.png"/>' }) -join '') + '</Relationships>'
            Set-ZipText $zip 'customUI/_rels/customUI14.xml.rels' $imgRels
            $ctEntry = $zip.GetEntry('[Content_Types].xml')
            $reader = New-Object System.IO.StreamReader($ctEntry.Open()); $ct = $reader.ReadToEnd(); $reader.Dispose()
            if ($ct -notmatch 'Extension="png"') { Set-ZipText $zip '[Content_Types].xml' ($ct -replace '<Default Extension="rels"', '<Default Extension="png" ContentType="image/png"/><Default Extension="rels"') }
        }
    } finally { $zip.Dispose(); if ($iconDir) { Remove-Item $iconDir -Recurse -Force -ErrorAction SilentlyContinue } }
}

# Garante docProps/core.xml com dc:description = $Description (o "Comments" do Office) no pacote. O PowerPoint grava o .ppam so com
# ppt/presentation.xml e ppt/vbaProject.bin, sem propriedades do documento (medido): o carimbo de versao lido sem abrir o arquivo
# (como o do loader, docProps/core.xml) e inserido aqui, com o content type e a relacao core-properties. Pacote que ja tem core.xml:
# troca (ou acrescenta) so o dc:description.
function Set-PackageDescription([string]$PackagePath, [string]$Description, [string]$Title = 'ThinQQ') {
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $esc = [Security.SecurityElement]::Escape($Description)
    $zip = [System.IO.Compression.ZipFile]::Open($PackagePath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        function Get-ZipText($zip, [string]$name) { $e = $zip.GetEntry($name); if (-not $e) { return $null }; $r = New-Object System.IO.StreamReader($e.Open()); try { return $r.ReadToEnd() } finally { $r.Dispose() } }
        function Set-ZipText($zip, [string]$name, [string]$text) {
            $old = $zip.GetEntry($name); if ($old) { $old.Delete() }
            $e = $zip.CreateEntry($name); $w = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false))); $w.Write($text); $w.Dispose()
        }
        $core = Get-ZipText $zip 'docProps/core.xml'
        if ($null -ne $core) {
            if ($core -match '<dc:description>[^<]*</dc:description>') { $core = $core -replace '<dc:description>[^<]*</dc:description>', "<dc:description>$esc</dc:description>" }
            elseif ($core -match '<dc:description\s*/>') { $core = $core -replace '<dc:description\s*/>', "<dc:description>$esc</dc:description>" }
            else { $core = $core -replace '</cp:coreProperties>', "<dc:description>$esc</dc:description></cp:coreProperties>" }
            Set-ZipText $zip 'docProps/core.xml' $core
        } else {
            $t = [Security.SecurityElement]::Escape($Title)
            $core = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
                "<dc:title>$t</dc:title><dc:description>$esc</dc:description></cp:coreProperties>"
            Set-ZipText $zip 'docProps/core.xml' $core
            $ct = Get-ZipText $zip '[Content_Types].xml'
            if ($ct -notmatch 'PartName="/docProps/core\.xml"') {
                Set-ZipText $zip '[Content_Types].xml' ($ct -replace '</Types>', '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>')
            }
            $rels = Get-ZipText $zip '_rels/.rels'
            if ($rels -notmatch 'metadata/core-properties') {
                Set-ZipText $zip '_rels/.rels' ($rels -replace '</Relationships>', '<Relationship Id="rIdThinQQCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>')
            }
        }
    } finally { $zip.Dispose() }
}

# ---- Watchdog do Excel -------------------------------------------------------
# Erro de compilação VBA ou diálogo modal trava a chamada COM para sempre. Um processo pwsh
# separado mata SOMENTE o Excel que este script iniciou (PID obtido via $xl.Hwnd) após o timeout;
# a chamada COM pendente então lança exceção e o chamador sai com código 2.
Add-Type -Namespace ThinQQ -Name Win32 -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint pid);
'@ -ErrorAction SilentlyContinue

function Get-ExcelPid($xl) {
    [uint32]$procId = 0
    [void][ThinQQ.Win32]::GetWindowThreadProcessId([IntPtr]::new([int64]$xl.Hwnd), [ref]$procId)
    return [int]$procId
}

# Watchdog genérico por PID (valores escapados para string entre aspas simples; usado também em testes com processo fictício).
function Start-PidWatchdog([int]$TargetPid, [int]$TimeoutSec, [string]$Marker) {
    if ($TargetPid -le 0) { throw "PID inválido para o watchdog ($TargetPid)." }
    $m = $Marker.Replace("'", "''")
    $cmd = "Wait-Process -Id $PID -Timeout $TimeoutSec -ErrorAction SilentlyContinue; " +
           "if (Get-Process -Id $PID -ErrorAction SilentlyContinue) { Set-Content -LiteralPath '$m' 'timeout'; Stop-Process -Id $TargetPid -Force -ErrorAction SilentlyContinue }"
    $enc = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($cmd))
    $pwsh = (Get-Process -Id $PID).Path
    $proc = Start-Process -FilePath $pwsh -ArgumentList @('-NoProfile', '-NonInteractive', '-EncodedCommand', $enc) -PassThru -WindowStyle Hidden
    return [pscustomobject]@{ Proc = $proc; Marker = $Marker; ExcelPid = $TargetPid }
}

function Start-ExcelWatchdog($xl, [int]$TimeoutSec) {
    $xlPid = Get-ExcelPid $xl
    if ($xlPid -le 0) { throw 'Não foi possível obter o PID do Excel (Hwnd inválido); watchdog não iniciado.' }
    $marker = Join-Path ([System.IO.Path]::GetTempPath()) "thinqq_watchdog_$PID`_$xlPid.flag"
    Remove-Item -LiteralPath $marker -Force -ErrorAction SilentlyContinue
    return Start-PidWatchdog $xlPid $TimeoutSec $marker
}
function Stop-ExcelWatchdog($wd) {
    if ($null -eq $wd) { return }
    if ($wd.Proc -and -not $wd.Proc.HasExited) { Stop-Process -Id $wd.Proc.Id -Force -ErrorAction SilentlyContinue }
}

function Test-ExcelWatchdogFired($wd) {
    return ($null -ne $wd -and (Test-Path -LiteralPath $wd.Marker))
}

function Write-ExcelTimeout($wd) {
    Remove-Item -LiteralPath $wd.Marker -Force -ErrorAction SilentlyContinue
    Write-Host 'TIMEOUT: o Excel não terminou (provável erro de compilação VBA ou diálogo modal)' -ForegroundColor Red
}

# ---- PowerPoint antes/depois da suite (run-tests). O PowerPoint que os testes do core abrem e criado pelo COM a pedido do Excel de teste
# (pai = o servico DcomLaunch, nao o Excel: nao ha como provar pelo PID/pai que e do teste); por isso o script NUNCA encerra POWERPNT aqui:
# so registra antes/depois e avisa com PID, pai e argumentos. Sobra de automacao (-Embedding) conta como falha (teste vazou o PowerPoint).
function Get-PptProcessSnapshot {
    $list = @()
    try {
        $list = @(Get-CimInstance Win32_Process -Filter "Name = 'POWERPNT.EXE'" -ErrorAction Stop | ForEach-Object {
            [pscustomobject]@{ Id = [int]$_.ProcessId; ParentId = [int]$_.ParentProcessId; CommandLine = [string]$_.CommandLine } })
    } catch {
        $list = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | ForEach-Object { [pscustomobject]@{ Id = [int]$_.Id; ParentId = 0; CommandLine = '' } })
    }
    return $list
}
function Get-PptLeftovers([object[]]$Before = @(), [object[]]$After = @()) {
    $ids = @(@($Before) | Where-Object { $null -ne $_ } | ForEach-Object { [int]$_.Id })
    return @(@($After) | Where-Object { $null -ne $_ -and $ids -notcontains [int]$_.Id })
}
function Test-PptAutomationProcess($Proc) { return ($null -ne $Proc -and [string]$Proc.CommandLine -match '(?i)[-/]Embedding') }
function Format-PptLeftover($Proc) {
    $kind = if (Test-PptAutomationProcess $Proc) { 'automação (provável órfão de teste)' } else { 'aberto com janela (provável do usuário)' }
    return "POWERPNT PID $($Proc.Id) (pai $($Proc.ParentId)), $kind; argumentos: $($Proc.CommandLine)"
}
# Testes do core com PowerPoint real (TestPowerPoint/TestPptSync) que pularam: cada pulo grava "TestX: ... pulado/pulada" no log de teste.
function Get-PptSkipCount([string[]]$LogLines = @()) {
    return @(@($LogLines) | Where-Object { $_ -match 'Test(PowerPoint|PptSync): .*\bpulad[oa]\b' }).Count
}
function Get-FileLengthOrZero([string]$Path) {
    if (Test-Path -LiteralPath $Path -PathType Leaf) { return [long](Get-Item -LiteralPath $Path).Length }
    return [long]0
}
# Linhas gravadas no log depois do deslocamento $Offset (bytes; o log do VBA e ANSI). Log ausente ou menor que o deslocamento = do inicio.
function Read-LogTail([string]$Path, [long]$Offset) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return }
    $fs = New-Object IO.FileStream($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
    try {
        if ($Offset -gt $fs.Length) { $Offset = 0 }
        [void]$fs.Seek($Offset, [IO.SeekOrigin]::Begin)
        $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::GetEncoding(1252))
        $text = $sr.ReadToEnd()
    } finally { $fs.Dispose() }
    return @($text -split "`r?`n" | Where-Object { $_ -ne '' })
}

# Gera modTestDispatch com chamadas DIRETAS a cada Test_* (Select Case). Application.Run abre diálogo modal
# em erro não tratado no Excel invisível; chamada direta respeita o On Error do modTestRunner.
function Write-TestDispatchModule([string[]]$TestFiles, [string]$OutPath) {
    $cases = New-Object System.Collections.Generic.List[string]
    foreach ($f in $TestFiles) {
        $text = Get-Content -Path $f -Raw -Encoding utf8
        if ($text -notmatch '(?m)^Attribute VB_Name = "(Test\w*)"') { continue }
        $mod = $Matches[1]
        foreach ($w in [regex]::Matches($text, '(?m)^\s*(?:(?:Public|Private|Friend)\s+)?Sub\s+(Test_\w+)[^\r\n]*')) {
            if ($w.Value -notmatch '^\s*Public Sub Test_\w+\s*\(\s*\)') { Write-Warning "$mod.$($w.Groups[1].Value) não é 'Public Sub Test_x()'; não será executado." }
        }
        foreach ($m in [regex]::Matches($text, '(?m)^\s*Public Sub (Test_\w+)\s*\(\s*\)')) {
            $cases.Add("        Case ""$mod.$($m.Groups[1].Value)"": $mod.$($m.Groups[1].Value)")
        }
    }
    $body = ($cases -join "`r`n")
    $names = @($cases | ForEach-Object { if ($_ -match 'Case "([^"]+)"') { '    s = s & "' + $Matches[1] + '|"' } })
    $list = ($names -join "`r`n")
    $content = @"
Attribute VB_Name = "modTestDispatch"
' Gerado pelo build - nao editar
Option Explicit

Public Sub Dispatch(ByVal qualifiedName As String)
    Select Case qualifiedName
$body
        Case Else: Err.Raise vbObjectError + 514, "Dispatch", "teste nao despachado: " & qualifiedName
    End Select
End Sub

Public Function AllTests() As String
    Dim s As String
$list
    AllTests = s
End Function
"@
    Set-Content -Path $OutPath -Value $content -Encoding ascii
}

# Compila o projeto VBA inteiro (Depurar > Compilar), pegando erros de sintaxe em procedimentos nunca chamados.
# Retorna $null se compilou; senão "Modulo:linha" do erro. Em caso de erro o Excel fica com diálogo modal pendente:
# o chamador deve encerrar o processo (Stop-Process no PID do Excel iniciado pelo build).
function Invoke-VbaCompile($xl, $wb) {
    $xl.VBE.MainWindow.Visible = $true
    $xl.VBE.ActiveVBProject = $wb.VBProject
    $ctl = $xl.VBE.CommandBars.FindControl(1, 578)   # 578 = Compilar VBAProject
    if ($null -eq $ctl) { throw 'Controle "Compilar VBAProject" (Id 578) não encontrado no VBE.' }
    $ctl.Execute()
    if (-not $xl.VBE.CommandBars.FindControl(1, 578).Enabled) { return $null }   # desabilitado = projeto já compilado
    $where = '?'
    try {
        $p = $xl.VBE.ActiveCodePane
        $a = [ref]0; $b = [ref]0; $c = [ref]0; $d = [ref]0
        $p.GetSelection($a, $b, $c, $d)
        $where = "$($p.CodeModule.Parent.Name):$($a.Value)"
    } catch { }
    return $where
}
