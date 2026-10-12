# Funções compartilhadas de build (carregar com dot-source).
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$script:DistDir  = Join-Path $script:RepoRoot 'dist'
. "$PSScriptRoot\Ribbon-Icons.ps1"

function Assert-VbomTrusted {
    $key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    $v = (Get-ItemProperty $key -ErrorAction SilentlyContinue).AccessVBOM
    if ($v -ne 1) {
        throw "Habilite 'Confiar no acesso ao modelo de objeto do projeto VBA' (Arquivo > Opções > Central de Confiabilidade > Configurações de Macro) ou execute: Set-ItemProperty '$key' AccessVBOM 1"
    }
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
function Add-RibbonToPackage([string]$PackagePath, [string]$RibbonXmlPath, [ValidateSet('Png','Mso')][string]$Icons = 'Png') {
    [void]([xml](Get-Content $RibbonXmlPath -Raw))   # falha cedo se o XML estiver malformado
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $xmlText = [IO.File]::ReadAllText($RibbonXmlPath, (New-Object Text.UTF8Encoding($false)))
    $images = @([regex]::Matches($xmlText, 'image="([^"]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
    if ($Icons -eq 'Mso') { $xmlText = Convert-TqqRibbonToMso -Xml $xmlText; $images = @() }
    $specs = Get-TqqIconSpecs
    foreach ($img in $images) { if (-not $specs.Contains(($img -replace '16$', ''))) { throw "customUI14.xml usa image=`"$img`" sem icone na tabela (Ribbon-Icons.ps1)" } }
    $iconDir = $null
    if ($images.Count -gt 0) { $iconDir = Join-Path ([IO.Path]::GetTempPath()) ("thinqq_icons_" + [guid]::NewGuid().ToString('N')); New-Item -ItemType Directory -Force $iconDir | Out-Null; [void](Export-TqqIcons -OutDir $iconDir) }
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
