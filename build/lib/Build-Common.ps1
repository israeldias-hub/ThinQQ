# Funções compartilhadas de build (carregar com dot-source).
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$script:DistDir  = Join-Path $script:RepoRoot 'dist'

function Assert-VbomTrusted {
    $key = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Security'
    $v = (Get-ItemProperty $key -ErrorAction SilentlyContinue).AccessVBOM
    if ($v -ne 1) {
        throw "Habilite 'Confiar no acesso ao modelo de objeto do projeto VBA' (Arquivo > Opções > Central de Confiabilidade > Configurações de Macro) ou execute: Set-ItemProperty '$key' AccessVBOM 1"
    }
}

function New-ExcelApp {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false
    return $xl
}

function Close-ExcelApp($xl) {
    if ($null -ne $xl) {
        try { $xl.Quit() } catch { }
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
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

# Insere customUI/customUI14.xml no pacote (zip) e a relação em _rels/.rels.
function Add-RibbonToPackage([string]$PackagePath, [string]$RibbonXmlPath) {
    [void]([xml](Get-Content $RibbonXmlPath -Raw))   # falha cedo se o XML estiver malformado
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::Open($PackagePath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $existing = $zip.GetEntry('customUI/customUI14.xml')
        if ($existing) { $existing.Delete() }
        $entry = $zip.CreateEntry('customUI/customUI14.xml')
        $bytes = [System.IO.File]::ReadAllBytes($RibbonXmlPath)
        $s = $entry.Open(); $s.Write($bytes, 0, $bytes.Length); $s.Dispose()

        $relsEntry = $zip.GetEntry('_rels/.rels')
        $reader = New-Object System.IO.StreamReader($relsEntry.Open())
        $rels = $reader.ReadToEnd(); $reader.Dispose()
        if ($rels -notmatch 'customUI/customUI14\.xml') {
            $rel = '<Relationship Id="rIdThinQQ" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
            $rels = $rels -replace '</Relationships>', "$rel</Relationships>"
            $relsEntry.Delete()
            $newRels = $zip.CreateEntry('_rels/.rels')
            $w = New-Object System.IO.StreamWriter($newRels.Open(), (New-Object System.Text.UTF8Encoding($false)))
            $w.Write($rels); $w.Dispose()
        }
    } finally { $zip.Dispose() }
}
