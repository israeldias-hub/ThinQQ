# ThinQQ — Plano de Implementação (Fase 1)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Entregar o ThinQQ 1.0.0: suplemento Excel (VBA, dois arquivos .xlam) que insere colunas empilhadas/agrupadas/100%/linha e cascata com totais, rótulos vivos, conectores e setas de CAGR/diferença; com aba própria na faixa, instalador, tutorial e atualização por download dentro do Excel.

**Architecture:** `ThinQQ.xlam` (loader: faixa de opções, boot, atualizador, configurações) abre `ThinQQ.core.xlam` (motor de gráficos) da pasta `%LOCALAPPDATA%\ThinQQ\versions\<x.y.z>\` e encaminha os comandos via `Application.Run`. Os gráficos são gráficos Excel nativos ligados a um bloco de fórmulas numa planilha muito oculta `_ThinQQ`; anotações são formas dentro do gráfico reposicionadas nos eventos `Calculate`/`Resize`. Fonte em texto no Git; `build.ps1` monta os .xlam via COM e `run-tests.ps1` roda testes VBA reais no Excel invisível.

**Tech Stack:** VBA 7 (Excel 2016+ Windows, 32/64 bits), customUI14 (ribbon XML), PowerShell 7 (build/testes) e PowerShell 5.1 (instalador), Node.js (apenas servidor de fixtures HTTP nos testes), Git.

**Spec:** `docs/superpowers/specs/2026-10-10-thinqq-design.md`

## Global Constraints

- Excel 2016 ou superior no Windows, 32 e 64 bits: **nenhuma** instrução `Declare` (só objetos COM por `CreateObject`); sem referências extras no projeto VBA (ligação tardia).
- Fórmulas sempre em inglês via `Range.Formula`/`FormulaR1C1`; códigos de formato numérico em inglês (`#,##0`), nunca `FormulaLocal`/`NumberFormatLocal`.
- Texto de interface em pt-BR; identificadores de código em inglês; prefixos de módulo: `Loader_`, `Updater_`, `Transport_`, `Settings_`, `Rx_` (loader) e `Core_`, `ChartData_`, `CalcSheet_`, `Registry_`, `Style_`, `Stacked_`, `Waterfall_`, `Geom_`, `Ann_`, `Labels_`, `Watchers_`, `Samples_`, `Util_` (core).
- Nomes reservados: planilha oculta `_ThinQQ`; `ChartObject.Name = "tqq_<id>"`; nome definido `tqq_<id>`; séries auxiliares `tqq_*`; formas `tqq_ann_*`; id = 6 caracteres `[a-z0-9]`.
- Versões no formato `x.y.z` (três inteiros). Versão do core e do loader em `version.json`; o build gera `modBuildInfo` com `APP_VERSION`, `BUILD_CONFIG` (`Release`|`Test`) e `BUILD_SUFFIX` (`""`|`".test"`).
- Sem VBScript. Instalador compatível com Windows PowerShell 5.1 (sem operadores `??`/ternário). Build e testes usam `pwsh` 7.
- Toda função pública do core chamada pelo loader (`Core_*`) trata seus erros (`Errors_Handle`) e nunca propaga exceção para o `Application.Run`.
- O updater só instala arquivo cujo SHA-256 confere com o manifesto (quando o manifesto informa o hash).
- Commits pequenos com mensagens `feat:`, `fix:`, `test:`, `build:`, `docs:` em português.

## Review Focus

1. **Tabela com células vazias ou texto ("n/d") no meio dos valores**: o gráfico insere sem erro e trata como zero/não plotado. Teste em Tarefa 8 (`Test_Stacked_TextCellTreatedAsZero`).
2. **Totais negativos e cascata que cruza o zero**: rótulos ficam abaixo da barra e os segmentos não se sobrepõem. Testes em Tarefa 8 (`Test_Stacked_NegativeTotal_LabelBelow`) e Tarefa 9 (`Test_Waterfall_Block_CrossingZero`).
3. **Nome de planilha com apóstrofo e espaço (`Plano 'A' 2026`)**: fórmulas e séries continuam válidas. Testes em Tarefa 5 (`Test_ChartData_ExternalAddress_EscapesQuote`) e Tarefa 8 (`Test_Stacked_Insert_SheetNameWithQuote`).
4. **Manifesto inacessível, malformado ou com SHA errado**: mensagem amigável e nada é instalado. Testes em Tarefa 13 (`Test_Updater_Evaluate_MalformedJson_IsNone`, `Test_Updater_Apply_BadSha_DoesNotInstall`).
5. **Usuário exclui o gráfico e insere outro**: o bloco órfão e o nome definido são removidos. Teste em Tarefa 6 (`Test_CalcSheet_Cleanup_RemovesOrphan`).

---

## Estrutura de arquivos

```
<repo>/
├── .gitignore
├── README.md
├── CHANGELOG.md
├── version.json                      {"loader":"1.0.0","core":"1.0.0"}
├── build/
│   ├── lib/Build-Common.ps1          funções: Excel COM, importar fontes, forms de JSON, injetar ribbon
│   ├── build.ps1                     monta dist/ThinQQ[.test].xlam e dist/ThinQQ.core[.test].xlam
│   ├── run-tests.ps1                 build Test + Excel invisível + RunAllTests (TAP) + servidor Node
│   └── release.ps1                   bump versão, build Release, manifest.json, zip, publicação
├── src/
│   ├── loader/
│   │   ├── ThisWorkbook.vba          Workbook_Open/BeforeClose
│   │   ├── modules/modLoaderMain.bas boot, pastas, abrir core, Application.Run
│   │   ├── modules/modVersion.bas    IsValidVersion, CompareVersions
│   │   ├── modules/modJsonLite.bas   JsonGetString, JsonGetNumber
│   │   ├── modules/modTransport.bas  HTTP/arquivo, download, SHA-256, UTF-8
│   │   ├── modules/modSettings.bas   registro HKCU (SaveSetting)
│   │   ├── modules/modUpdater.bas    verificação, formulário, instalação
│   │   ├── modules/modRibbon.bas     callbacks Rx_* → Core_*
│   │   ├── forms/frmUpdate.json + frmUpdate.vba
│   │   └── ribbon/customUI14.xml
│   └── core/
│       ├── ThisWorkbook.vba
│       ├── modules/modConstants.bas, modCoreApi.bas, modErrors.bas, modChartUtil.bas,
│       │   modChartData.bas, modCalcSheet.bas, modChartRegistry.bas, modChartStyle.bas,
│       │   modLabels.bas, modChartStacked.bas, modChartWaterfall.bas, modGeometry.bas,
│       │   modAnnotations.bas, modWatchers.bas, modSamples.bas
│       ├── classes/clsChartData.cls, clsChartWatcher.cls, clsAppEvents.cls
│       └── forms/frmAnnotation.json + frmAnnotation.vba
├── tests/
│   ├── vba/common/modAssert.bas, modTestRunner.bas
│   ├── vba/loader/Test*.bas
│   ├── vba/core/Test*.bas
│   ├── http/server.js
│   ├── fixtures/http/hello.txt
│   └── ps/Test-Installer.ps1
├── installer/Instalar-ThinQQ.cmd, install.ps1, uninstall.ps1, LEIA-ME.txt
├── release/manifest.json, notes.txt
└── docs/tutorial/tutorial.html
```

Comandos de referência (sempre a partir da raiz do repositório, em `pwsh`):

```powershell
pwsh build/build.ps1 -Target all -Configuration Release        # monta dist/
pwsh build/build.ps1 -Target all -Configuration Release -Install # + copia para %LOCALAPPDATA%\ThinQQ (Excel fechado)
pwsh build/run-tests.ps1                                        # build Test + testes core e loader
pwsh build/run-tests.ps1 -Target core -NoBuild                  # reaproveita o build Test
```

---

### Tarefa 1: Esqueleto do repositório, build via COM e runner de testes

**Files:**
- Create: `.gitignore`, `version.json`, `README.md`
- Create: `build/lib/Build-Common.ps1`, `build/build.ps1`, `build/run-tests.ps1`
- Create: `src/loader/ThisWorkbook.vba`, `src/loader/modules/modLoaderMain.bas`, `src/loader/ribbon/customUI14.xml`, `src/loader/modules/modRibbon.bas`
- Create: `src/loader/forms/frmUpdate.json`, `src/loader/forms/frmUpdate.vba`
- Create: `src/core/ThisWorkbook.vba`, `src/core/modules/modConstants.bas`, `src/core/modules/modCoreApi.bas`
- Create: `tests/vba/common/modAssert.bas`, `tests/vba/common/modTestRunner.bas`, `tests/vba/core/TestSmoke.bas`, `tests/vba/loader/TestSmokeLoader.bas`, `tests/http/server.js`, `tests/fixtures/http/hello.txt`

**Interfaces:**
- Produces (loader): `Loader_Ping() As String`; `frmUpdate` com `Prepare(title, versionLine, notes, Optional showSkip)`, propriedade `Result` (`"update"|"later"|"skip"`); `Rx_OnLoad(ribbon)`, `Rx_OnAbout(control)`, `Loader_InvalidateRibbon()`.
- Produces (core): `Core_Ping() As String`, `Core_Version() As String`; constantes `APP_NAME`, `CALC_SHEET`, `ID_PREFIX`, `ANN_PREFIX`, `HELPER_PREFIX`.
- Produces (testes): `AssertTrue`, `AssertFalse`, `AssertEqual`, `AssertAlmost`, `NewTempWorkbook()`, `CloseTemp(wb)`, `TempDir()`, `RepoPath()`, `RunAllTests(outputPath) As Long`.

- [ ] **Passo 1: Iniciar o repositório e arquivos base**

```powershell
git init -b main
```

`.gitignore`:

```
dist/
tests/results/
tests/fixtures/release/
*.tmp
~$*
```

`version.json`:

```json
{ "loader": "1.0.0", "core": "1.0.0" }
```

`README.md` (conteúdo inicial; completado na Tarefa 15):

```markdown
# ThinQQ

Suplemento Excel (VBA) para gráficos de apresentação no estilo think-cell.
Veja `docs/superpowers/specs/2026-10-10-thinqq-design.md` (design) e
`docs/superpowers/plans/2026-10-10-thinqq-fase1.md` (plano).

## Build e testes
```powershell
pwsh build/build.ps1 -Target all -Configuration Release
pwsh build/run-tests.ps1
```
```

- [ ] **Passo 2: Escrever o teste de fumaça (vai falhar: não existe build nem `Core_Ping`)**

`tests/vba/core/TestSmoke.bas`:

```vb
Attribute VB_Name = "TestSmoke"
Option Explicit

Public Sub Test_Core_Ping()
    AssertEqual "pong", Core_Ping()
End Sub

Public Sub Test_BuildInfo_IsTestConfig()
    AssertEqual "Test", BUILD_CONFIG
    AssertEqual ".test", BUILD_SUFFIX
End Sub
```

`tests/vba/loader/TestSmokeLoader.bas`:

```vb
Attribute VB_Name = "TestSmokeLoader"
Option Explicit

Public Sub Test_Loader_Ping()
    AssertEqual "pong", Loader_Ping()
End Sub

' Prova que o pipeline de UserForm por JSON funciona (o formulário é consumido na Tarefa 13).
Public Sub Test_Loader_FormUpdate_Exists()
    Dim frm As Object
    Set frm = VBA.UserForms.Add("frmUpdate")
    AssertTrue Not frm Is Nothing, "frmUpdate não foi criado pelo build"
    frm.Prepare "Título", "1.0.0 -> 1.1.0", "notas"
    AssertEqual "Título", frm.Controls("lblTitle").Caption
    Unload frm
End Sub
```

- [ ] **Passo 3: Escrever o mini-framework de testes**

`tests/vba/common/modAssert.bas`:

```vb
Attribute VB_Name = "modAssert"
Option Explicit
Public Const ASSERT_ERR As Long = vbObjectError + 513

Public Sub AssertTrue(cond As Boolean, Optional msg As String = "esperado True")
    If Not cond Then Err.Raise ASSERT_ERR, "Assert", msg
End Sub

Public Sub AssertFalse(cond As Boolean, Optional msg As String = "esperado False")
    If cond Then Err.Raise ASSERT_ERR, "Assert", msg
End Sub

Public Sub AssertEqual(expected As Variant, actual As Variant, Optional msg As String = "")
    Dim same As Boolean
    If IsNumeric(expected) And IsNumeric(actual) Then
        same = (CDbl(expected) = CDbl(actual))
    Else
        same = (CStr(expected) = CStr(actual))
    End If
    If Not same Then Err.Raise ASSERT_ERR, "Assert", msg & " esperado=<" & CStr(expected) & "> obtido=<" & CStr(actual) & ">"
End Sub

Public Sub AssertAlmost(expected As Double, actual As Double, Optional tol As Double = 0.000001, Optional msg As String = "")
    If Abs(expected - actual) > tol Then Err.Raise ASSERT_ERR, "Assert", msg & " esperado~" & expected & " obtido=" & actual
End Sub

Public Function NewTempWorkbook() As Workbook
    Set NewTempWorkbook = Application.Workbooks.Add
End Function

Public Sub CloseTemp(wb As Workbook)
    On Error Resume Next
    wb.Close SaveChanges:=False
End Sub

' Pasta temporária nova e vazia (o chamador apaga se quiser).
Public Function TempDir() As String
    Dim p As String
    Randomize
    p = Environ$("TEMP") & "\thinqq_test_" & Format$(Now, "hhnnss") & "_" & CLng(Rnd * 1000000)
    MkDir p
    TempDir = p
End Function

' Raiz do repositório (o .xlam de teste fica em <repo>\dist).
Public Function RepoPath() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    RepoPath = fso.GetParentFolderName(ThisWorkbook.Path)
End Function
```

`tests/vba/common/modTestRunner.bas`:

```vb
Attribute VB_Name = "modTestRunner"
Option Explicit
' Executa todo Sub público "Test_*" de todo módulo "Test*" e grava saída no formato TAP.
Private mPassed As Long, mFailed As Long, mLog As String

Public Function RunAllTests(outputPath As String) As Long
    Dim comp As Object, cm As Object, line As Long, procName As String, kind As Long, lastProc As String
    mPassed = 0: mFailed = 0: mLog = ""
    For Each comp In ThisWorkbook.VBProject.VBComponents
        If comp.Type = 1 And Left$(comp.Name, 4) = "Test" Then    ' 1 = vbext_ct_StdModule
            Set cm = comp.CodeModule
            lastProc = ""
            For line = cm.CountOfDeclarationLines + 1 To cm.CountOfLines
                procName = cm.ProcOfLine(line, kind)
                If procName <> lastProc Then
                    lastProc = procName
                    If Left$(procName, 5) = "Test_" Then RunOne comp.Name, procName
                End If
            Next
        End If
    Next
    mLog = mLog & "1.." & (mPassed + mFailed) & vbLf & "# passed " & mPassed & " failed " & mFailed & vbLf
    WriteText outputPath, mLog
    RunAllTests = mFailed
End Function

Private Sub RunOne(moduleName As String, procName As String)
    On Error GoTo Fail
    Application.Run "'" & ThisWorkbook.Name & "'!" & moduleName & "." & procName
    mPassed = mPassed + 1
    mLog = mLog & "ok " & (mPassed + mFailed) & " - " & moduleName & "." & procName & vbLf
    Exit Sub
Fail:
    mFailed = mFailed + 1
    mLog = mLog & "not ok " & (mPassed + mFailed) & " - " & moduleName & "." & procName & " # " & Replace(Replace(Err.Description, vbCr, " "), vbLf, " ") & vbLf
End Sub

Private Sub WriteText(path As String, text As String)
    Dim f As Integer
    f = FreeFile
    Open path For Output As #f
    Print #f, text
    Close #f
End Sub
```

- [ ] **Passo 4: Escrever os fontes mínimos do loader e do core**

`src/loader/modules/modLoaderMain.bas` (versão mínima; cresce na Tarefa 4):

```vb
Attribute VB_Name = "modLoaderMain"
Option Explicit
Public Const APP_NAME As String = "ThinQQ"

Public Function Loader_Ping() As String
    Loader_Ping = "pong"
End Function
```

`src/loader/ThisWorkbook.vba`:

```vb
Option Explicit

Private Sub Workbook_Open()
    If BUILD_CONFIG = "Test" Then Exit Sub
    Loader_Startup
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If BUILD_CONFIG = "Test" Then Exit Sub
    Loader_Shutdown
End Sub
```

Como `Loader_Startup`/`Loader_Shutdown` só existem na Tarefa 4, acrescente ao `modLoaderMain.bas` desta tarefa estes dois stubs (a Tarefa 4 os substitui):

```vb
Public Sub Loader_Startup()
    ' substituído na Tarefa 4
End Sub

Public Sub Loader_Shutdown()
    ' substituído na Tarefa 4
End Sub
```

`src/loader/modules/modRibbon.bas` (mínimo; completo na Tarefa 12):

```vb
Attribute VB_Name = "modRibbon"
Option Explicit
Private mRibbon As IRibbonUI

Public Sub Rx_OnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

Public Sub Loader_InvalidateRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

Public Sub Rx_OnAbout(control As IRibbonControl)
    MsgBox "ThinQQ loader " & APP_VERSION, vbInformation, APP_NAME
End Sub
```

`src/loader/ribbon/customUI14.xml` (mínimo; completo na Tarefa 12):

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Rx_OnLoad">
  <ribbon>
    <tabs>
      <tab id="tqqTab" label="ThinQQ" insertAfterMso="TabInsert">
        <group id="grpAbout" label="ThinQQ">
          <button id="btnAbout" label="Sobre" size="large" imageMso="Info" onAction="Rx_OnAbout"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

`src/loader/forms/frmUpdate.json`:

```json
{
  "name": "frmUpdate",
  "caption": "Atualização do ThinQQ",
  "width": 430,
  "height": 320,
  "controls": [
    { "type": "Label", "name": "lblTitle", "Left": 16, "Top": 12, "Width": 396, "Height": 22,
      "Caption": "Nova versão disponível", "fontName": "Segoe UI", "fontSize": 12, "fontBold": true },
    { "type": "Label", "name": "lblVersion", "Left": 16, "Top": 38, "Width": 396, "Height": 16,
      "Caption": "", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "Label", "name": "lblNotesTitle", "Left": 16, "Top": 62, "Width": 396, "Height": 14,
      "Caption": "Novidades desta versão:", "fontName": "Segoe UI", "fontSize": 9, "fontBold": true },
    { "type": "TextBox", "name": "txtNotes", "Left": 16, "Top": 80, "Width": 396, "Height": 150,
      "MultiLine": true, "Locked": true, "ScrollBars": 2, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnUpdate", "Left": 196, "Top": 246, "Width": 110, "Height": 28,
      "Caption": "Atualizar agora", "Default": true, "fontName": "Segoe UI", "fontSize": 9, "fontBold": true },
    { "type": "CommandButton", "name": "btnLater", "Left": 312, "Top": 246, "Width": 100, "Height": 28,
      "Caption": "Depois", "Cancel": true, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnSkip", "Left": 16, "Top": 246, "Width": 120, "Height": 28,
      "Caption": "Pular esta versão", "fontName": "Segoe UI", "fontSize": 9 }
  ]
}
```

`src/loader/forms/frmUpdate.vba`:

```vb
Option Explicit
Public Result As String

Public Sub Prepare(title As String, versionLine As String, notes As String, Optional showSkip As Boolean = True)
    lblTitle.Caption = title
    lblVersion.Caption = versionLine
    txtNotes.Text = notes
    btnSkip.Visible = showSkip
    Result = "later"
End Sub

Private Sub btnUpdate_Click()
    Result = "update"
    Me.Hide
End Sub

Private Sub btnLater_Click()
    Result = "later"
    Me.Hide
End Sub

Private Sub btnSkip_Click()
    Result = "skip"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Cancel = True
        Result = "later"
        Me.Hide
    End If
End Sub
```

`src/core/modules/modConstants.bas`:

```vb
Attribute VB_Name = "modConstants"
Option Explicit
Public Const APP_NAME As String = "ThinQQ"
Public Const CALC_SHEET As String = "_ThinQQ"
Public Const ID_PREFIX As String = "tqq_"
Public Const ANN_PREFIX As String = "tqq_ann_"
Public Const HELPER_PREFIX As String = "tqq_"
```

`src/core/modules/modCoreApi.bas` (mínimo; completo na Tarefa 11):

```vb
Attribute VB_Name = "modCoreApi"
Option Explicit

Public Function Core_Ping() As String
    Core_Ping = "pong"
End Function

Public Function Core_Version() As String
    Core_Version = APP_VERSION
End Function
```

`src/core/ThisWorkbook.vba`:

```vb
Option Explicit

Private Sub Workbook_Open()
    ' Core_OnLoaded passa a existir na Tarefa 11; até lá nada a fazer.
End Sub
```

- [ ] **Passo 5: Escrever as funções de build compartilhadas**

`build/lib/Build-Common.ps1`:

```powershell
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

function Import-VbaFile($vbproj, [string]$Path) {
    [void]$vbproj.VBComponents.Import($Path)
}

function Add-CodeToComponent($vbproj, [string]$ComponentName, [string]$CodePath) {
    $cm = $vbproj.VBComponents.Item($ComponentName).CodeModule
    $cm.AddFromString((Get-Content -Path $CodePath -Raw -Encoding utf8))
}

# Cria um UserForm a partir de um JSON declarativo e injeta o code-behind (.vba).
function Add-FormFromJson($vbproj, [string]$JsonPath, [string]$CodePath) {
    $spec = Get-Content -Path $JsonPath -Raw -Encoding utf8 | ConvertFrom-Json
    $comp = $vbproj.VBComponents.Add(3)   # 3 = vbext_ct_MSForm
    $comp.Name = $spec.name
    $d = $comp.Designer
    $d.Caption = $spec.caption
    $d.Width   = [double]$spec.width
    $d.Height  = [double]$spec.height
    foreach ($c in $spec.controls) {
        $ctl = $d.Controls.Add("Forms.$($c.type).1", $c.name, $true)
        foreach ($p in $c.PSObject.Properties) {
            switch ($p.Name) {
                'type'     { }
                'name'     { }
                'fontName' { $ctl.Font.Name = [string]$p.Value }
                'fontSize' { $ctl.Font.Size = [double]$p.Value }
                'fontBold' { $ctl.Font.Bold = [bool]$p.Value }
                default    { $ctl.($p.Name) = $p.Value }
            }
        }
    }
    if ($CodePath -and (Test-Path $CodePath)) {
        $comp.CodeModule.AddFromString((Get-Content -Path $CodePath -Raw -Encoding utf8))
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
```

- [ ] **Passo 6: Escrever `build/build.ps1`**

```powershell
<#
.SYNOPSIS  Monta ThinQQ.xlam (loader) e ThinQQ.core.xlam (core) a partir de src/.
.EXAMPLE   pwsh build/build.ps1 -Target all -Configuration Release -Install
#>
param(
    [ValidateSet('loader','core','all')] [string]$Target = 'all',
    [ValidateSet('Release','Test')]      [string]$Configuration = 'Release',
    [switch]$Install
)
. "$PSScriptRoot\lib\Build-Common.ps1"
Assert-VbomTrusted
$versions = Get-Content (Join-Path $RepoRoot 'version.json') -Raw | ConvertFrom-Json
New-Item -ItemType Directory -Force $DistDir | Out-Null
$targets = if ($Target -eq 'all') { @('loader','core') } else { @($Target) }

foreach ($t in $targets) {
    $srcDir  = Join-Path $RepoRoot "src\$t"
    $version = $versions.$t
    $suffix  = if ($Configuration -eq 'Test') { '.test' } else { '' }
    $outName = if ($t -eq 'loader') { "ThinQQ$suffix.xlam" } else { "ThinQQ.core$suffix.xlam" }
    $outPath = Join-Path $DistDir $outName
    if (Test-Path $outPath) { Remove-Item $outPath -Force }

    $buildInfo = Join-Path $env:TEMP "modBuildInfo_$t.bas"
    Write-BuildInfoModule -Version $version -Configuration $Configuration -OutPath $buildInfo

    $xl = New-ExcelApp
    try {
        $wb = $xl.Workbooks.Add()
        while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
        $wb.Worksheets.Item(1).Name = 'ThinQQ'
        $vbproj = $wb.VBProject
        $vbproj.Name = "ThinQQ_$t"

        Import-VbaFile $vbproj $buildInfo
        Get-ChildItem (Join-Path $srcDir 'modules') -Filter *.bas | ForEach-Object { Import-VbaFile $vbproj $_.FullName }
        $clsDir = Join-Path $srcDir 'classes'
        if (Test-Path $clsDir) { Get-ChildItem $clsDir -Filter *.cls | ForEach-Object { Import-VbaFile $vbproj $_.FullName } }
        $formsDir = Join-Path $srcDir 'forms'
        if (Test-Path $formsDir) {
            Get-ChildItem $formsDir -Filter *.json | ForEach-Object {
                Add-FormFromJson $vbproj $_.FullName (Join-Path $formsDir ($_.BaseName + '.vba'))
            }
        }
        $twb = Join-Path $srcDir 'ThisWorkbook.vba'
        if (Test-Path $twb) { Add-CodeToComponent $vbproj 'ThisWorkbook' $twb }

        if ($Configuration -eq 'Test') {
            Get-ChildItem (Join-Path $RepoRoot 'tests\vba\common') -Filter *.bas | ForEach-Object { Import-VbaFile $vbproj $_.FullName }
            $tDir = Join-Path $RepoRoot "tests\vba\$t"
            if (Test-Path $tDir) { Get-ChildItem $tDir -Filter *.bas | ForEach-Object { Import-VbaFile $vbproj $_.FullName } }
        }

        $wb.BuiltinDocumentProperties.Item('Title').Value    = 'ThinQQ'
        $wb.BuiltinDocumentProperties.Item('Comments').Value = "ThinQQ $t $version ($Configuration)"
        $wb.IsAddin = $true
        $wb.SaveAs($outPath, 55)   # 55 = xlOpenXMLAddIn
        $wb.Close($false)
    } finally { Close-ExcelApp $xl }

    if ($t -eq 'loader') { Add-RibbonToPackage -PackagePath $outPath -RibbonXmlPath (Join-Path $srcDir 'ribbon\customUI14.xml') }
    Write-Host "OK  $outPath"
}

if ($Install -and $Configuration -eq 'Release') {
    $root = Join-Path $env:LOCALAPPDATA 'ThinQQ'
    $verDir = Join-Path $root "versions\$($versions.core)"
    New-Item -ItemType Directory -Force $verDir, (Join-Path $root 'tmp'), (Join-Path $root 'logs') | Out-Null
    Copy-Item (Join-Path $DistDir 'ThinQQ.xlam') $root -Force
    Copy-Item (Join-Path $DistDir 'ThinQQ.core.xlam') $verDir -Force
    Write-Host "Instalado em $root (abra o Excel; na primeira vez registre o suplemento em Arquivo > Opções > Suplementos > Ir... > Procurar)"
}
```

- [ ] **Passo 7: Escrever `build/run-tests.ps1` e o servidor de fixtures**

`tests/http/server.js`:

```js
// Servidor estático mínimo para fixtures de teste: node server.js <porta> <pasta>
const http = require('http'), fs = require('fs'), path = require('path');
const port = Number(process.argv[2] || 18080);
const root = process.argv[3] || '.';
http.createServer((req, res) => {
  const file = path.join(root, decodeURIComponent(req.url.split('?')[0]));
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); res.end('not found'); return; }
    res.writeHead(200, { 'Content-Type': 'application/octet-stream', 'Cache-Control': 'no-store' });
    res.end(data);
  });
}).listen(port, '127.0.0.1');
```

`tests/fixtures/http/hello.txt` (salvar em UTF-8 sem BOM, sem quebra de linha final):

```
olá ThinQQ
```

`build/run-tests.ps1`:

```powershell
param([ValidateSet('loader','core','all')] [string]$Target = 'all', [switch]$NoBuild)
. "$PSScriptRoot\lib\Build-Common.ps1"
if (-not $NoBuild) { & "$PSScriptRoot\build.ps1" -Target all -Configuration Test }

$resultsDir = Join-Path $RepoRoot 'tests\results'
New-Item -ItemType Directory -Force $resultsDir | Out-Null
Remove-Item (Join-Path $resultsDir 'http-base.txt') -ErrorAction SilentlyContinue
$targets = if ($Target -eq 'all') { @('core','loader') } else { @($Target) }

# Fixtures de release para o teste do updater (core de teste + manifesto apontando para o servidor local)
$relDir = Join-Path $RepoRoot 'tests\fixtures\release'
New-Item -ItemType Directory -Force $relDir | Out-Null
Copy-Item (Join-Path $DistDir 'ThinQQ.core.test.xlam') $relDir -Force
$sha = (Get-FileHash (Join-Path $relDir 'ThinQQ.core.test.xlam') -Algorithm SHA256).Hash.ToLower()
@{ schemaVersion = 1; coreVersion = '9.9.9'; coreUrl = 'http://127.0.0.1:18080/release/ThinQQ.core.test.xlam'
   coreSha256 = $sha; loaderVersion = '0.0.1'; loaderUrl = ''; loaderSha256 = ''; tutorialUrl = ''
   publishedAt = '2099-01-01'; notes = "• fixture de teste" } | ConvertTo-Json | Set-Content (Join-Path $relDir 'manifest.json') -Encoding utf8

$server = $null
$node = Get-Command node -ErrorAction SilentlyContinue
if ($node) {
    $server = Start-Process -FilePath $node.Source -ArgumentList @((Join-Path $RepoRoot 'tests\http\server.js'), '18080', (Join-Path $RepoRoot 'tests\fixtures')) -PassThru -WindowStyle Hidden
    Start-Sleep -Milliseconds 800
    Set-Content (Join-Path $resultsDir 'http-base.txt') 'http://127.0.0.1:18080/' -NoNewline
}

$totalFailed = 0
$xl = New-ExcelApp
$xl.EnableEvents = $true
try {
    $core = $xl.Workbooks.Open((Join-Path $DistDir 'ThinQQ.core.test.xlam'))
    foreach ($t in $targets) {
        $wb = if ($t -eq 'core') { $core } else { $xl.Workbooks.Open((Join-Path $DistDir 'ThinQQ.test.xlam')) }
        $out = Join-Path $resultsDir "$t.tap"
        Remove-Item $out -ErrorAction SilentlyContinue
        $failed = $xl.Run("'$($wb.Name)'!modTestRunner.RunAllTests", $out)
        Write-Host "=== $t ==="
        Get-Content $out
        $totalFailed += [int]$failed
    }
} finally {
    Close-ExcelApp $xl
    if ($server) { Stop-Process -Id $server.Id -Force -ErrorAction SilentlyContinue }
}
if ($totalFailed -gt 0) { Write-Host "FALHAS: $totalFailed" -ForegroundColor Red; exit 1 }
Write-Host 'Todos os testes passaram' -ForegroundColor Green
```

- [ ] **Passo 8: Rodar o build e os testes**

Run: `pwsh build/run-tests.ps1`
Expected: duas seções (`=== core ===`, `=== loader ===`), todas as linhas `ok`, final `Todos os testes passaram`.

Se o Excel travar (diálogo de erro de compilação invisível): finalize `EXCEL.EXE` no Gerenciador de Tarefas, abra `dist\ThinQQ.core.test.xlam` no Excel, `Alt+F11` > Depurar > Compilar e corrija.
Se `Add-FormFromJson` falhar em `.Designer`: crie o formulário manualmente no VBE uma vez, exporte `frmUpdate.frm/.frx` para `src/loader/forms/` e troque, no `build.ps1`, a chamada `Add-FormFromJson` por `Import-VbaFile $vbproj (caminho .frm)`.

- [ ] **Passo 9: Commit**

```bash
git add -A
git commit -m "build: esqueleto do ThinQQ com build via COM, ribbon injetado e runner de testes TAP"
```

---

### Tarefa 2: Versões semânticas e leitor de JSON plano (loader)

**Files:**
- Create: `src/loader/modules/modVersion.bas`, `src/loader/modules/modJsonLite.bas`
- Test: `tests/vba/loader/TestVersion.bas`, `tests/vba/loader/TestJsonLite.bas`

**Interfaces:**
- Produces: `IsValidVersion(v As String) As Boolean`; `CompareVersions(a As String, b As String) As Long` (1, 0, -1; inválida conta como `0.0.0`); `JsonGetString(json As String, key As String) As String` (`""` se ausente; desfaz escapes `\n \" \\ \/ \uXXXX`); `JsonGetNumber(json, key) As Double`.

- [ ] **Passo 1: Escrever os testes**

`tests/vba/loader/TestVersion.bas`:

```vb
Attribute VB_Name = "TestVersion"
Option Explicit

Public Sub Test_Version_Valid()
    AssertTrue IsValidVersion("1.0.0")
    AssertTrue IsValidVersion("10.20.30")
End Sub

Public Sub Test_Version_Invalid()
    AssertFalse IsValidVersion("")
    AssertFalse IsValidVersion("1.0")
    AssertFalse IsValidVersion("v1.0.0")
    AssertFalse IsValidVersion("1.0.0-beta")
    AssertFalse IsValidVersion("1..0")
End Sub

Public Sub Test_Version_Compare()
    AssertEqual 1, CompareVersions("1.2.10", "1.2.9")
    AssertEqual -1, CompareVersions("1.9.0", "1.10.0")
    AssertEqual 0, CompareVersions("2.0.0", "2.0.0")
    AssertEqual 1, CompareVersions("0.0.1", "")
    AssertEqual -1, CompareVersions("abc", "0.0.1")
End Sub
```

`tests/vba/loader/TestJsonLite.bas`:

```vb
Attribute VB_Name = "TestJsonLite"
Option Explicit

Public Sub Test_Json_String()
    AssertEqual "1.2.3", JsonGetString("{""coreVersion"": ""1.2.3"", ""x"": 1}", "coreVersion")
End Sub

Public Sub Test_Json_Escapes()
    AssertEqual "a" & vbLf & "b ""c"" \ é /", JsonGetString("{""notes"":""a\nb \""c\"" \\ \u00e9 \/""}", "notes")
End Sub

Public Sub Test_Json_MissingKey_ReturnsEmpty()
    AssertEqual "", JsonGetString("{""a"":""b""}", "coreUrl")
End Sub

Public Sub Test_Json_Number_And_Bool()
    AssertEqual 2, JsonGetNumber("{""schemaVersion"":2,""a"":""b""}", "schemaVersion")
    AssertEqual "true", JsonGetString("{""flag"": true}", "flag")
End Sub

Public Sub Test_Json_KeyPrefixNotConfused()
    AssertEqual "B", JsonGetString("{""coreUrlX"":""A"",""coreUrl"":""B""}", "coreUrl")
End Sub

Public Sub Test_Json_MultilinePretty()
    Dim j As String
    j = "{" & vbCrLf & "  ""coreVersion"": ""1.0.0""," & vbCrLf & "  ""notes"": ""x""" & vbCrLf & "}"
    AssertEqual "1.0.0", JsonGetString(j, "coreVersion")
    AssertEqual "x", JsonGetString(j, "notes")
End Sub
```

- [ ] **Passo 2: Rodar os testes para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target loader`
Expected: linhas `not ok ... TestVersion.*` e `TestJsonLite.*` (função não definida).

- [ ] **Passo 3: Implementar `modVersion.bas`**

```vb
Attribute VB_Name = "modVersion"
Option Explicit

Public Function IsValidVersion(v As String) As Boolean
    Dim parts() As String, i As Long, k As Long
    If Len(v) = 0 Then Exit Function
    parts = Split(v, ".")
    If UBound(parts) <> 2 Then Exit Function
    For i = 0 To 2
        If Len(parts(i)) = 0 Or Len(parts(i)) > 6 Then Exit Function
        For k = 1 To Len(parts(i))
            If InStr("0123456789", Mid$(parts(i), k, 1)) = 0 Then Exit Function
        Next
    Next
    IsValidVersion = True
End Function

' 1 se a > b, -1 se a < b, 0 se iguais. Versão inválida vale 0.0.0.
Public Function CompareVersions(a As String, b As String) As Long
    Dim pa() As String, pb() As String, i As Long, na As Long, nb As Long
    pa = Split(IIf(IsValidVersion(a), a, "0.0.0"), ".")
    pb = Split(IIf(IsValidVersion(b), b, "0.0.0"), ".")
    For i = 0 To 2
        na = CLng(pa(i)): nb = CLng(pb(i))
        If na > nb Then CompareVersions = 1: Exit Function
        If na < nb Then CompareVersions = -1: Exit Function
    Next
    CompareVersions = 0
End Function
```

- [ ] **Passo 4: Implementar `modJsonLite.bas`**

```vb
Attribute VB_Name = "modJsonLite"
Option Explicit
' Leitor mínimo de JSON "plano": chaves de primeiro nível com valores string/número/booleano.

Public Function JsonGetString(json As String, key As String) As String
    Dim p As Long, q As Long, needle As String
    needle = """" & key & """"
    p = InStr(1, json, needle, vbBinaryCompare)
    If p = 0 Then Exit Function
    p = InStr(p + Len(needle), json, ":")
    If p = 0 Then Exit Function
    p = p + 1
    Do While p <= Len(json)
        If InStr(1, " " & vbTab & vbCr & vbLf, Mid$(json, p, 1)) = 0 Then Exit Do
        p = p + 1
    Loop
    If p > Len(json) Then Exit Function
    If Mid$(json, p, 1) = """" Then
        JsonGetString = JsonReadString(json, p + 1)
    Else
        q = p
        Do While q <= Len(json)
            If InStr(1, ",}]" & vbCr & vbLf & " " & vbTab, Mid$(json, q, 1)) > 0 Then Exit Do
            q = q + 1
        Loop
        JsonGetString = Mid$(json, p, q - p)
    End If
End Function

Public Function JsonGetNumber(json As String, key As String) As Double
    Dim s As String
    s = JsonGetString(json, key)
    If Len(s) = 0 Then Exit Function
    JsonGetNumber = Val(s)   ' Val usa ponto decimal, independente do locale
End Function

' Lê uma string JSON a partir da posição após a aspa de abertura, desfazendo escapes.
Private Function JsonReadString(json As String, startPos As Long) As String
    Dim i As Long, ch As String, out As String, esc As String
    i = startPos
    Do While i <= Len(json)
        ch = Mid$(json, i, 1)
        If ch = """" Then Exit Do
        If ch = "\" Then
            esc = Mid$(json, i + 1, 1)
            Select Case esc
                Case "n": out = out & vbLf
                Case "r": out = out & vbCr
                Case "t": out = out & vbTab
                Case "b": out = out & Chr$(8)
                Case "f": out = out & Chr$(12)
                Case "u": out = out & ChrW$(CLng("&H" & Mid$(json, i + 2, 4))): i = i + 4
                Case Else: out = out & esc    ' \" \\ \/
            End Select
            i = i + 2
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    JsonReadString = out
End Function
```

- [ ] **Passo 5: Rodar os testes para ver passar**

Run: `pwsh build/run-tests.ps1 -Target loader`
Expected: todas as linhas `TestVersion.*` e `TestJsonLite.*` com `ok`.

- [ ] **Passo 6: Commit**

```bash
git add src/loader/modules/modVersion.bas src/loader/modules/modJsonLite.bas tests/vba/loader/TestVersion.bas tests/vba/loader/TestJsonLite.bas
git commit -m "feat(loader): comparação de versões semânticas e leitor de JSON plano"
```

---

### Tarefa 3: Transporte (HTTP e arquivo), SHA-256 e configurações (loader)

**Files:**
- Create: `src/loader/modules/modTransport.bas`, `src/loader/modules/modSettings.bas`
- Test: `tests/vba/loader/TestTransport.bas`, `tests/vba/loader/TestSettings.bas`

**Interfaces:**
- Produces: `Transport_IsHttp(source) As Boolean`; `Transport_ExpandPath(p) As String` (expande `%VAR%`); `Transport_FetchText(source) As String` (UTF-8; erro se HTTP ≠ 200 ou arquivo ausente); `Transport_DownloadFile(source, destPath)`; `Transport_Sha256File(path) As String` (hex minúsculo); `Transport_ReadTextFile(path)`, `Transport_WriteTextFile(path, text)`.
- Produces: `Settings_Get(key, Optional default) As String`, `Settings_Set(key, value)`, `Settings_Delete(key)`, `Settings_UpdateSource() As String`; constantes `KEY_UPDATE_SOURCE`, `KEY_LAST_CHECK`, `KEY_SKIPPED`, `KEY_PENDING_LOADER`, `KEY_INTERVAL`, `KEY_LAST_NOTIFIED_CORE`, `DEFAULT_UPDATE_SOURCE`.
- Consumes: `TempDir()`, `RepoPath()` (Tarefa 1).

- [ ] **Passo 1: Escrever os testes**

`tests/vba/loader/TestTransport.bas`:

```vb
Attribute VB_Name = "TestTransport"
Option Explicit

Private Function HttpBase() As String
    Dim p As String
    p = RepoPath() & "\tests\results\http-base.txt"
    If Len(Dir(p)) = 0 Then Err.Raise vbObjectError + 900, "TestTransport", "servidor HTTP de fixtures ausente (Node não encontrado?)"
    HttpBase = Transport_ReadTextFile(p)
End Function

Public Sub Test_Transport_IsHttp()
    AssertTrue Transport_IsHttp("https://x/y.json")
    AssertTrue Transport_IsHttp("HTTP://x")
    AssertFalse Transport_IsHttp("\\servidor\pasta\manifest.json")
    AssertFalse Transport_IsHttp("C:\a.json")
End Sub

Public Sub Test_Transport_ExpandPath()
    AssertEqual Environ$("TEMP") & "\x\y", Transport_ExpandPath("%TEMP%\x\y")
    AssertEqual "C:\sem\variavel", Transport_ExpandPath("C:\sem\variavel")
End Sub

Public Sub Test_Transport_File_RoundTrip_Utf8()
    Dim d As String, p As String
    d = TempDir(): p = d & "\t.txt"
    Transport_WriteTextFile p, "ação & café"
    AssertEqual "ação & café", Transport_ReadTextFile(p)
    AssertEqual "ação & café", Transport_FetchText(p)
End Sub

Public Sub Test_Transport_DownloadFile_FromPath()
    Dim d As String
    d = TempDir()
    Transport_WriteTextFile d & "\src.txt", "abc"
    Transport_DownloadFile d & "\src.txt", d & "\dst.txt"
    AssertEqual "abc", Transport_ReadTextFile(d & "\dst.txt")
End Sub

Public Sub Test_Transport_Sha256_KnownVector()
    Dim d As String, f As Integer, b() As Byte
    d = TempDir()
    b = StrConv("abc", vbFromUnicode)
    f = FreeFile: Open d & "\abc.bin" For Binary As #f: Put #f, , b: Close #f
    AssertEqual "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad", Transport_Sha256File(d & "\abc.bin")
    f = FreeFile: Open d & "\empty.bin" For Binary As #f: Close #f
    AssertEqual "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855", Transport_Sha256File(d & "\empty.bin")
End Sub

Public Sub Test_Transport_FetchText_MissingFile_Raises()
    Dim raised As Boolean
    On Error Resume Next
    Transport_FetchText "C:\nao\existe\manifest.json"
    raised = (Err.Number <> 0)
    On Error GoTo 0
    AssertTrue raised, "deveria lançar erro para arquivo inexistente"
End Sub

Public Sub Test_Transport_Http_FetchText()
    AssertEqual "olá ThinQQ", Transport_FetchText(HttpBase() & "http/hello.txt")
End Sub

Public Sub Test_Transport_Http_404_Raises()
    Dim raised As Boolean
    On Error Resume Next
    Transport_FetchText HttpBase() & "http/nao-existe.txt"
    raised = (Err.Number <> 0)
    On Error GoTo 0
    AssertTrue raised, "deveria lançar erro em HTTP 404"
End Sub

Public Sub Test_Transport_Http_DownloadFile()
    Dim d As String
    d = TempDir()
    Transport_DownloadFile HttpBase() & "http/hello.txt", d & "\hello.txt"
    AssertEqual "olá ThinQQ", Transport_ReadTextFile(d & "\hello.txt")
End Sub
```

`tests/vba/loader/TestSettings.bas`:

```vb
Attribute VB_Name = "TestSettings"
Option Explicit

Public Sub Test_Settings_RoundTrip()
    Settings_Set "TestKey", "valor"
    AssertEqual "valor", Settings_Get("TestKey")
    Settings_Delete "TestKey"
    AssertEqual "padrao", Settings_Get("TestKey", "padrao")
End Sub

Public Sub Test_Settings_UpdateSource_Default()
    Dim saved As String
    saved = Settings_Get(KEY_UPDATE_SOURCE)
    Settings_Delete KEY_UPDATE_SOURCE
    AssertEqual DEFAULT_UPDATE_SOURCE, Settings_UpdateSource()
    If Len(saved) > 0 Then Settings_Set KEY_UPDATE_SOURCE, saved
End Sub
```

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target loader`
Expected: `not ok` em todos os `TestTransport.*` e `TestSettings.*`.

- [ ] **Passo 3: Implementar `modTransport.bas`**

```vb
Attribute VB_Name = "modTransport"
Option Explicit
' Acesso a manifesto/arquivos por HTTP(S) (MSXML2.XMLHTTP, respeita proxy do usuário) ou por caminho (UNC/local).

Public Function Transport_IsHttp(source As String) As Boolean
    Transport_IsHttp = (LCase$(Left$(source, 7)) = "http://") Or (LCase$(Left$(source, 8)) = "https://")
End Function

' Expande %VARIAVEIS% de ambiente (ex.: %USERPROFILE%\OneDrive\ThinQQ\manifest.json).
Public Function Transport_ExpandPath(p As String) As String
    Dim out As String, i As Long, j As Long, name As String
    out = p
    i = InStr(1, out, "%")
    Do While i > 0
        j = InStr(i + 1, out, "%")
        If j = 0 Then Exit Do
        name = Mid$(out, i + 1, j - i - 1)
        out = Left$(out, i - 1) & Environ$(name) & Mid$(out, j + 1)
        i = InStr(i + Len(Environ$(name)), out, "%")
    Loop
    Transport_ExpandPath = out
End Function

Public Function Transport_FetchText(source As String) As String
    If Transport_IsHttp(source) Then
        Transport_FetchText = BytesToUtf8(HttpGetBytes(source))
    Else
        Transport_FetchText = Transport_ReadTextFile(Transport_ExpandPath(source))
    End If
End Function

Public Sub Transport_DownloadFile(source As String, destPath As String)
    If Transport_IsHttp(source) Then
        SaveBytes HttpGetBytes(source), destPath
    Else
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile Transport_ExpandPath(source), destPath, True
    End If
End Sub

Public Function Transport_ReadTextFile(path As String) As String
    Dim stm As Object
    If Len(Dir(path)) = 0 Then Err.Raise vbObjectError + 1002, "Transport", "Arquivo não encontrado: " & path
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.LoadFromFile path
    Transport_ReadTextFile = stm.ReadText(-1)
    stm.Close
End Function

Public Sub Transport_WriteTextFile(path As String, text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "utf-8": stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2    ' adSaveCreateOverWrite
    stm.Close
End Sub

Public Function Transport_Sha256File(path As String) As String
    Dim stm As Object, data() As Byte, sha As Object, hash() As Byte, i As Long, out As String
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open: stm.LoadFromFile path
    If stm.Size = 0 Then
        stm.Close
        Transport_Sha256File = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
        Exit Function
    End If
    data = stm.Read(-1)
    stm.Close
    Set sha = CreateObject("System.Security.Cryptography.SHA256Managed")
    hash = sha.ComputeHash_2((data))
    For i = LBound(hash) To UBound(hash)
        out = out & Right$("0" & LCase$(Hex$(hash(i))), 2)
    Next
    Transport_Sha256File = out
End Function

Private Function HttpGetBytes(url As String) As Byte()
    Dim http As Object, sep As String
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    sep = IIf(InStr(url, "?") > 0, "&", "?")
    http.Open "GET", url & sep & "_ts=" & CStr(CLng(CDbl(Now) * 86400)), False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.Send
    If http.Status <> 200 Then Err.Raise vbObjectError + 1001, "Transport", "HTTP " & http.Status & " ao acessar " & url
    HttpGetBytes = http.responseBody
End Function

Private Sub SaveBytes(data() As Byte, destPath As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open
    stm.Write data
    stm.SaveToFile destPath, 2
    stm.Close
End Sub

Private Function BytesToUtf8(data() As Byte) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open
    stm.Write data
    stm.Position = 0: stm.Type = 2: stm.Charset = "utf-8"
    BytesToUtf8 = stm.ReadText(-1)
    stm.Close
End Function
```

- [ ] **Passo 4: Implementar `modSettings.bas`**

```vb
Attribute VB_Name = "modSettings"
Option Explicit
' Configurações por usuário em HKCU\Software\VB and VBA Program Settings\ThinQQ\Updater
Private Const REG_APP As String = "ThinQQ"
Private Const REG_SECTION As String = "Updater"
Public Const KEY_UPDATE_SOURCE As String = "UpdateSource"
Public Const KEY_LAST_CHECK As String = "LastCheck"
Public Const KEY_SKIPPED As String = "SkippedVersion"
Public Const KEY_PENDING_LOADER As String = "PendingLoaderPath"
Public Const KEY_INTERVAL As String = "CheckIntervalHours"
Public Const KEY_LAST_NOTIFIED_CORE As String = "LastNotifiedCore"
Public Const DEFAULT_UPDATE_SOURCE As String = "https://raw.githubusercontent.com/israeldias-hub/ThinQQ/main/release/manifest.json"

Public Function Settings_Get(key As String, Optional defaultValue As String = "") As String
    Settings_Get = GetSetting(REG_APP, REG_SECTION, key, defaultValue)
End Function

Public Sub Settings_Set(key As String, value As String)
    SaveSetting REG_APP, REG_SECTION, key, value
End Sub

Public Sub Settings_Delete(key As String)
    On Error Resume Next
    DeleteSetting REG_APP, REG_SECTION, key
End Sub

Public Function Settings_UpdateSource() As String
    Settings_UpdateSource = Settings_Get(KEY_UPDATE_SOURCE, DEFAULT_UPDATE_SOURCE)
End Function
```

- [ ] **Passo 5: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1 -Target loader`
Expected: todos `ok`, incluindo os três testes `Test_Transport_Http_*` (o servidor Node é iniciado pelo `run-tests.ps1`).

- [ ] **Passo 6: Commit**

```bash
git add src/loader/modules/modTransport.bas src/loader/modules/modSettings.bas tests/vba/loader/TestTransport.bas tests/vba/loader/TestSettings.bas
git commit -m "feat(loader): transporte HTTP/arquivo com SHA-256 e configurações no registro"
```

---
### Tarefa 4: Boot do loader e carregamento do core

**Files:**
- Modify: `src/loader/modules/modLoaderMain.bas` (substituir inteiro)
- Test: `tests/vba/loader/TestLoaderMain.bas`

**Interfaces:**
- Produces: `Loader_Root()`, `Loader_VersionsDir()`, `Loader_TmpDir()`, `Loader_EnsureFolders()`, `Loader_SetRootOverride(path)` (testes), `Loader_CoreFileName() As String` (`"ThinQQ.core" & BUILD_SUFFIX & ".xlam"`), `Loader_FindNewestCoreVersion(versionsDir) As String`, `Loader_CoreWorkbook() As Workbook`, `Loader_OpenCore(path) As Boolean`, `Loader_CloseCore()`, `Loader_RunCore(procName, ParamArray args()) As Variant` (Empty se core ausente), `Loader_CurrentCoreVersion() As String` (`"0.0.0"` se ausente), `Loader_Startup()`, `Loader_Shutdown()`, `Loader_Log(msg)`.
- Consumes: `IsValidVersion`, `CompareVersions` (Tarefa 2); `Core_Ping`, `Core_Version` (Tarefa 1). `Core_RegisterLoader`/`Core_OnUnload` (Tarefa 11) são chamados com `On Error Resume Next` até existirem.

- [ ] **Passo 1: Escrever os testes**

`tests/vba/loader/TestLoaderMain.bas`:

```vb
Attribute VB_Name = "TestLoaderMain"
Option Explicit

Private Sub MakeVersion(root As String, ver As String)
    MkDir root & "\" & ver
    Dim f As Integer
    f = FreeFile: Open root & "\" & ver & "\" & Loader_CoreFileName() For Output As #f: Print #f, "x": Close #f
End Sub

Public Sub Test_CoreFileName_UsesSuffix()
    AssertEqual "ThinQQ.core.test.xlam", Loader_CoreFileName()
End Sub

Public Sub Test_FindNewestCoreVersion()
    Dim d As String
    d = TempDir()
    MakeVersion d, "1.0.0"
    MakeVersion d, "1.10.0"
    MakeVersion d, "1.2.0"
    MkDir d & "\abc"                      ' pasta inválida: ignorada
    MkDir d & "\2.0.0"                    ' sem arquivo do core: ignorada
    AssertEqual "1.10.0", Loader_FindNewestCoreVersion(d)
End Sub

Public Sub Test_FindNewestCoreVersion_EmptyOrMissing()
    AssertEqual "", Loader_FindNewestCoreVersion(TempDir())
    AssertEqual "", Loader_FindNewestCoreVersion("C:\nao\existe")
End Sub

Public Sub Test_RootOverride_And_EnsureFolders()
    Dim d As String
    d = TempDir()
    Loader_SetRootOverride d & "\root"
    Loader_EnsureFolders
    AssertTrue Len(Dir(d & "\root\versions", vbDirectory)) > 0
    AssertTrue Len(Dir(d & "\root\tmp", vbDirectory)) > 0
    AssertEqual d & "\root\versions", Loader_VersionsDir()
    Loader_SetRootOverride ""
    AssertEqual Environ$("LOCALAPPDATA") & "\ThinQQ", Loader_Root()
End Sub

Public Sub Test_RunCore_Ping()
    ' O run-tests.ps1 abre dist\ThinQQ.core.test.xlam antes dos testes do loader.
    AssertTrue Not Loader_CoreWorkbook() Is Nothing, "core de teste não está aberto"
    AssertEqual "pong", Loader_RunCore("Core_Ping")
    AssertTrue IsValidVersion(Loader_CurrentCoreVersion())
End Sub

Public Sub Test_OpenCore_SamePath_IsNoop()
    Dim p As String
    p = Loader_CoreWorkbook().FullName
    AssertTrue Loader_OpenCore(p)
    AssertEqual p, Loader_CoreWorkbook().FullName
End Sub
```

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target loader`
Expected: `not ok` em `TestLoaderMain.*`.

- [ ] **Passo 3: Substituir `modLoaderMain.bas` pela implementação completa**

```vb
Attribute VB_Name = "modLoaderMain"
Option Explicit
Public Const APP_NAME As String = "ThinQQ"
Private mRootOverride As String

Public Function Loader_Ping() As String
    Loader_Ping = "pong"
End Function

Public Sub Loader_SetRootOverride(path As String)
    mRootOverride = path
End Sub

Public Function Loader_Root() As String
    If Len(mRootOverride) > 0 Then Loader_Root = mRootOverride Else Loader_Root = Environ$("LOCALAPPDATA") & "\ThinQQ"
End Function

Public Function Loader_VersionsDir() As String
    Loader_VersionsDir = Loader_Root() & "\versions"
End Function

Public Function Loader_TmpDir() As String
    Loader_TmpDir = Loader_Root() & "\tmp"
End Function

Public Function Loader_CoreFileName() As String
    Loader_CoreFileName = "ThinQQ.core" & BUILD_SUFFIX & ".xlam"
End Function

Public Sub Loader_EnsureFolders()
    EnsureDir Loader_Root()
    EnsureDir Loader_VersionsDir()
    EnsureDir Loader_TmpDir()
    EnsureDir Loader_Root() & "\logs"
End Sub

Private Sub EnsureDir(path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    If Not fso.FolderExists(fso.GetParentFolderName(path)) Then EnsureDir fso.GetParentFolderName(path)
    fso.CreateFolder path
End Sub

Public Sub Loader_Log(msg As String)
    On Error Resume Next
    Dim f As Integer
    Loader_EnsureFolders
    f = FreeFile
    Open Loader_Root() & "\logs\thinqq.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & "loader" & vbTab & msg
    Close #f
End Sub

' Maior versão x.y.z em versionsDir que contenha o arquivo do core ("" se nenhuma).
Public Function Loader_FindNewestCoreVersion(versionsDir As String) As String
    Dim fso As Object, fld As Object, sub_ As Object, best As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(versionsDir) Then Exit Function
    Set fld = fso.GetFolder(versionsDir)
    For Each sub_ In fld.SubFolders
        If IsValidVersion(sub_.Name) Then
            If fso.FileExists(sub_.Path & "\" & Loader_CoreFileName()) Then
                If best = "" Then
                    best = sub_.Name
                ElseIf CompareVersions(sub_.Name, best) > 0 Then
                    best = sub_.Name
                End If
            End If
        End If
    Next
    Loader_FindNewestCoreVersion = best
End Function

Public Function Loader_CoreWorkbook() As Workbook
    On Error Resume Next
    Set Loader_CoreWorkbook = Application.Workbooks(Loader_CoreFileName())
    On Error GoTo 0
End Function

Public Sub Loader_CloseCore()
    Dim wb As Workbook
    Set wb = Loader_CoreWorkbook()
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Application.Run "'" & Loader_CoreFileName() & "'!Core_OnUnload"
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

' Abre (ou troca) o core. Retorna False se falhar.
Public Function Loader_OpenCore(path As String) As Boolean
    Dim wb As Workbook
    Set wb = Loader_CoreWorkbook()
    If Not wb Is Nothing Then
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then Loader_OpenCore = True: Exit Function
        Loader_CloseCore
    End If
    On Error GoTo Fail
    Set wb = Application.Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    On Error Resume Next
    Application.Run "'" & Loader_CoreFileName() & "'!Core_RegisterLoader", ThisWorkbook.Name, ThisWorkbook.FullName
    On Error GoTo 0
    Loader_OpenCore = True
    Exit Function
Fail:
    Loader_Log "Loader_OpenCore falhou: " & path & " - " & Err.Description
    Loader_OpenCore = False
End Function

' Encaminha uma chamada ao core. Retorna Empty se o core não estiver carregado.
Public Function Loader_RunCore(procName As String, ParamArray args() As Variant) As Variant
    Dim macro As String
    If Loader_CoreWorkbook() Is Nothing Then Exit Function
    macro = "'" & Loader_CoreFileName() & "'!" & procName
    Select Case UBound(args) - LBound(args) + 1
        Case 0: Loader_RunCore = Application.Run(macro)
        Case 1: Loader_RunCore = Application.Run(macro, args(0))
        Case 2: Loader_RunCore = Application.Run(macro, args(0), args(1))
        Case 3: Loader_RunCore = Application.Run(macro, args(0), args(1), args(2))
    End Select
End Function

Public Function Loader_CurrentCoreVersion() As String
    Dim v As Variant
    v = Loader_RunCore("Core_Version")
    If IsEmpty(v) Then Loader_CurrentCoreVersion = "0.0.0" Else Loader_CurrentCoreVersion = CStr(v)
End Function

Public Sub Loader_Startup()
    On Error GoTo Fail
    Dim ver As String
    Loader_EnsureFolders
    ver = Loader_FindNewestCoreVersion(Loader_VersionsDir())
    If ver = "" Then
        Loader_Log "core não encontrado em " & Loader_VersionsDir()
        If MsgBox("O ThinQQ está instalado, mas o módulo principal (core) não foi encontrado." & vbLf & _
                  "Deseja baixá-lo agora pela fonte de atualização?", vbYesNo + vbQuestion, APP_NAME) = vbYes Then
            Application.OnTime Now + TimeSerial(0, 0, 1), "'" & ThisWorkbook.Name & "'!Updater_CheckInteractive"
        End If
        Exit Sub
    End If
    If Not Loader_OpenCore(Loader_VersionsDir() & "\" & ver & "\" & Loader_CoreFileName()) Then
        MsgBox "Não foi possível carregar o ThinQQ core " & ver & ". Veja o log em " & Loader_Root() & "\logs.", vbExclamation, APP_NAME
        Exit Sub
    End If
    Application.OnTime Now + TimeSerial(0, 0, 8), "'" & ThisWorkbook.Name & "'!Updater_StartupCheck"
    Exit Sub
Fail:
    Loader_Log "Loader_Startup: " & Err.Description
End Sub

Public Sub Loader_Shutdown()
    On Error Resume Next
    Loader_CloseCore
End Sub
```

`Updater_StartupCheck` e `Updater_CheckInteractive` passam a existir na Tarefa 13. Até lá, acrescente ao fim de `modLoaderMain.bas` os stubs abaixo (a Tarefa 13 os remove):

```vb
Public Sub Updater_StartupCheck()
End Sub

Public Sub Updater_CheckInteractive()
End Sub
```

- [ ] **Passo 4: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1 -Target loader`
Expected: todos `ok`.

- [ ] **Passo 5: Verificação manual do boot**

1. Feche o Excel. Run: `pwsh build/build.ps1 -Target all -Configuration Release -Install`
2. Abra o Excel > Arquivo > Opções > Suplementos > Gerenciar: Suplementos do Excel > Ir... > Procurar > `%LOCALAPPDATA%\ThinQQ\ThinQQ.xlam` > OK.
3. Expected: aba **ThinQQ** com o botão "Sobre"; clicar mostra "ThinQQ loader 1.0.0". No VBE (`Alt+F11`), a janela Verificação Imediata com `?Workbooks("ThinQQ.core.xlam").Name` responde `ThinQQ.core.xlam`.

- [ ] **Passo 6: Commit**

```bash
git add src/loader/modules/modLoaderMain.bas tests/vba/loader/TestLoaderMain.bas
git commit -m "feat(loader): boot que localiza e carrega a versão mais nova do core"
```

---

### Tarefa 5: Modelo de dados do gráfico (core)

**Files:**
- Create: `src/core/classes/clsChartData.cls`, `src/core/modules/modChartData.bas`
- Test: `tests/vba/core/TestChartData.bas`

**Interfaces:**
- Produces: `ChartData_FromRange(rng As Range) As clsChartData` (uma célula expande para `CurrentRegion`); `ChartData_ExternalAddress(r As Range) As String` (`='Nome da planilha'!$A$1:$B$2`, apóstrofos dobrados); classe `clsChartData` com `Init(src)`, `Source`, `CategoryCount`, `SeriesCount`, `CategoryCells()`, `SeriesNameCell(j)`, `SeriesValueCells(j)`, `ValueCell(i, j)`, `CategoryText(i)`, `SeriesName(j)`, `Validate() As String` (`""` se válido).

- [ ] **Passo 1: Escrever os testes**

`tests/vba/core/TestChartData.bas`:

```vb
Attribute VB_Name = "TestChartData"
Option Explicit

Private Function Fixture(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    ws.Name = "Dados"
    ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
    ws.Range("A2:D2").Value = Array(2022, 10, 20, 5)
    ws.Range("A3:D3").Value = Array(2023, 12, 22, 6)
    ws.Range("A4:D4").Value = Array(2024, 15, 25, 8)
    ws.Range("A5:D5").Value = Array(2025, 18, 30, 9)
    Set Fixture = ws
End Function

Public Sub Test_ChartData_Shape()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = ChartData_FromRange(ws.Range("A1:D5"))
    AssertEqual 4, d.CategoryCount
    AssertEqual 3, d.SeriesCount
    AssertEqual "Sudeste", d.SeriesName(2)
    AssertEqual "2023", d.CategoryText(2)
    AssertEqual "$A$2:$A$5", d.CategoryCells().Address
    AssertEqual "$C$2:$C$5", d.SeriesValueCells(2).Address
    AssertEqual "$C$1", d.SeriesNameCell(2).Address
    AssertEqual 25, d.ValueCell(3, 2).Value
    AssertEqual "", d.Validate()
    CloseTemp wb
End Sub

Public Sub Test_ChartData_SingleCell_ExpandsToCurrentRegion()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = ChartData_FromRange(ws.Range("B3"))
    AssertEqual "$A$1:$D$5", d.Source.Address
    CloseTemp wb
End Sub

Public Sub Test_ChartData_ExternalAddress_EscapesQuote()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Name = "Plano 'A' 2026"
    AssertEqual "='Plano ''A'' 2026'!$B$2:$B$5", ChartData_ExternalAddress(ws.Range("B2:B5"))
    CloseTemp wb
End Sub

Public Sub Test_ChartData_Validate_TooSmall()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = ChartData_FromRange(ws.Range("A1:A5"))
    AssertTrue Len(d.Validate()) > 0, "uma coluna só deveria ser inválido"
    Set d = ChartData_FromRange(ws.Range("A1:D1"))
    AssertTrue Len(d.Validate()) > 0, "uma linha só deveria ser inválido"
    CloseTemp wb
End Sub

Public Sub Test_ChartData_Validate_MultiArea()
    Dim wb As Workbook, ws As Worksheet, d As clsChartData
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    Set d = ChartData_FromRange(ws.Range("A1:B5,D1:D5"))
    AssertTrue Len(d.Validate()) > 0
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: `not ok` em `TestChartData.*`.

- [ ] **Passo 3: Implementar a classe e o módulo**

`src/core/classes/clsChartData.cls`:

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Tabela de entrada: linha 1 = cabeçalho (canto + nomes das séries); coluna 1 = categorias.
Private mSource As Range

Public Sub Init(src As Range)
    Set mSource = src
End Sub

Public Property Get Source() As Range
    Set Source = mSource
End Property

Public Property Get CategoryCount() As Long
    CategoryCount = mSource.Rows.Count - 1
End Property

Public Property Get SeriesCount() As Long
    SeriesCount = mSource.Columns.Count - 1
End Property

Public Function CategoryCells() As Range
    Set CategoryCells = mSource.Offset(1, 0).Resize(CategoryCount, 1)
End Function

Public Function SeriesNameCell(j As Long) As Range
    Set SeriesNameCell = mSource.Cells(1, j + 1)
End Function

Public Function SeriesValueCells(j As Long) As Range
    Set SeriesValueCells = mSource.Offset(1, j).Resize(CategoryCount, 1)
End Function

Public Function ValueCell(i As Long, j As Long) As Range
    Set ValueCell = mSource.Cells(i + 1, j + 1)
End Function

Public Function CategoryText(i As Long) As String
    CategoryText = CStr(mSource.Cells(i + 1, 1).Text)
End Function

Public Function SeriesName(j As Long) As String
    SeriesName = CStr(SeriesNameCell(j).Text)
End Function

Public Function Validate() As String
    If mSource Is Nothing Then Validate = "Selecione a tabela de dados.": Exit Function
    If mSource.Areas.Count > 1 Then Validate = "Selecione uma única área contígua.": Exit Function
    If mSource.Rows.Count < 2 Or mSource.Columns.Count < 2 Then
        Validate = "Selecione a tabela com a linha de cabeçalho (nomes das séries) e a coluna de categorias (mínimo 2 linhas x 2 colunas)."
        Exit Function
    End If
    If mSource.Worksheet.Name = CALC_SHEET Then Validate = "Essa é a planilha interna do ThinQQ.": Exit Function
    Validate = ""
End Function
```

`src/core/modules/modChartData.bas`:

```vb
Attribute VB_Name = "modChartData"
Option Explicit

Public Function ChartData_FromRange(rng As Range) As clsChartData
    Dim src As Range, d As clsChartData
    Set src = rng
    If src.Cells.CountLarge = 1 Then Set src = src.CurrentRegion
    Set d = New clsChartData
    d.Init src
    Set ChartData_FromRange = d
End Function

' ='Nome da planilha'!$A$1:$B$2 (apóstrofos dobrados), aceito por Series.Values/Name/XValues e por Range.Formula.
Public Function ChartData_ExternalAddress(r As Range) As String
    ChartData_ExternalAddress = "='" & Replace(r.Worksheet.Name, "'", "''") & "'!" & r.Address(True, True)
End Function
```

- [ ] **Passo 4: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: todos `ok`.

- [ ] **Passo 5: Commit**

```bash
git add src/core/classes/clsChartData.cls src/core/modules/modChartData.bas tests/vba/core/TestChartData.bas
git commit -m "feat(core): modelo de dados da tabela selecionada com validação"
```

---

### Tarefa 6: Planilha de cálculo oculta, blocos e registro de gráficos (core)

**Files:**
- Create: `src/core/modules/modCalcSheet.bas`, `src/core/modules/modChartRegistry.bas`, `src/core/modules/modChartUtil.bas`
- Test: `tests/vba/core/TestCalcSheet.bas`

**Interfaces:**
- Produces (modCalcSheet): `CalcSheet_Get(wb, Optional createIfMissing = True) As Worksheet` (muito oculta); `CalcSheet_AllocateBlock(wb, id, chartType, sourceAddr, bodyRows, bodyCols) As Range` (bloco = linha meta + corpo; cria nome oculto `tqq_<id>`); `CalcSheet_FindBlock(wb, id) As Range`; `CalcSheet_DeleteBlock(wb, id)`; `CalcSheet_Cleanup(wb) As Long` (remove blocos sem gráfico); `Block_Meta(blk) As Range`; `Block_Body(blk) As Range`.
- Produces (modChartRegistry): `Registry_NewId()`, `Registry_IdFromChart(cht)`, `Registry_IsThinQQChart(cht)`, `Registry_Workbook(cht)`, `Registry_Block(cht)`, `Registry_ChartType(cht)`, `Registry_FindChartObject(wb, id) As ChartObject`, `Registry_GetOption(cht, key, Optional default) As String`, `Registry_SetOption(cht, key, value)`.
- Produces (modChartUtil): `Util_EnsureDir(path)`, `Util_FileName(path)`, `Util_ToDouble(v) As Double`, `Util_ColumnToDoubles(rng) As Double()`, `Chart_FindSeries(cht, name) As Series`.
- Meta do bloco: A=id, B=tipo (`stacked|clustered|stacked100|line|waterfall`), C=endereço fonte **sem** `=` inicial (gravado como fórmula `="'Plan'!$A$1"` para preservar o apóstrofo), D=opções `chave=valor;...`.

- [ ] **Passo 1: Escrever os testes**

`tests/vba/core/TestCalcSheet.bas`:

```vb
Attribute VB_Name = "TestCalcSheet"
Option Explicit

Public Sub Test_CalcSheet_Get_CreatesVeryHidden()
    Dim wb As Workbook, ws As Worksheet, active As String
    Set wb = NewTempWorkbook()
    active = wb.ActiveSheet.Name
    Set ws = CalcSheet_Get(wb)
    AssertEqual CALC_SHEET, ws.Name
    AssertEqual xlSheetVeryHidden, ws.Visible
    AssertEqual active, wb.ActiveSheet.Name, "planilha ativa não deve mudar"
    AssertTrue CalcSheet_Get(wb) Is ws, "segunda chamada reutiliza"
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_AllocateAndFind()
    Dim wb As Workbook, blk As Range, found As Range
    Set wb = NewTempWorkbook()
    Set blk = CalcSheet_AllocateBlock(wb, "ab12cd", "stacked", "'Dados'!$A$1:$D$5", 5, 3)
    AssertEqual 6, blk.Rows.Count
    AssertEqual "ab12cd", Block_Meta(blk).Cells(1, 1).Value
    AssertEqual "stacked", Block_Meta(blk).Cells(1, 2).Value
    AssertEqual "'Dados'!$A$1:$D$5", Block_Meta(blk).Cells(1, 3).Value
    AssertEqual 5, Block_Body(blk).Rows.Count
    Set found = CalcSheet_FindBlock(wb, "ab12cd")
    AssertEqual blk.Address, found.Address
    AssertTrue CalcSheet_FindBlock(wb, "zzzzzz") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_TwoBlocks_DoNotOverlap()
    Dim wb As Workbook, b1 As Range, b2 As Range
    Set wb = NewTempWorkbook()
    Set b1 = CalcSheet_AllocateBlock(wb, "aaaaaa", "stacked", "x", 4, 3)
    Set b2 = CalcSheet_AllocateBlock(wb, "bbbbbb", "waterfall", "y", 6, 18)
    AssertTrue b2.Row > b1.Row + b1.Rows.Count, "segundo bloco deve ficar abaixo do primeiro"
    AssertTrue Application.Intersect(b1, b2) Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_DeleteBlock()
    Dim wb As Workbook
    Set wb = NewTempWorkbook()
    CalcSheet_AllocateBlock wb, "aaaaaa", "stacked", "x", 4, 3
    CalcSheet_DeleteBlock wb, "aaaaaa"
    AssertTrue CalcSheet_FindBlock(wb, "aaaaaa") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_CalcSheet_Cleanup_RemovesOrphan()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    CalcSheet_AllocateBlock wb, "orphan", "stacked", "x", 4, 3
    CalcSheet_AllocateBlock wb, "alive1", "stacked", "x", 4, 3
    Set co = ws.ChartObjects.Add(10, 10, 100, 100)
    co.Name = ID_PREFIX & "alive1"
    AssertEqual 1, CalcSheet_Cleanup(wb)
    AssertTrue CalcSheet_FindBlock(wb, "orphan") Is Nothing
    AssertTrue Not CalcSheet_FindBlock(wb, "alive1") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_Registry_IdAndType()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    CalcSheet_AllocateBlock wb, "k1k1k1", "waterfall", "x", 4, 18
    Set co = ws.ChartObjects.Add(10, 10, 100, 100)
    co.Name = ID_PREFIX & "k1k1k1"
    AssertEqual "k1k1k1", Registry_IdFromChart(co.Chart)
    AssertTrue Registry_IsThinQQChart(co.Chart)
    AssertEqual "waterfall", Registry_ChartType(co.Chart)
    AssertTrue Registry_Workbook(co.Chart) Is wb
    AssertEqual co.Name, Registry_FindChartObject(wb, "k1k1k1").Name
    Set co = ws.ChartObjects.Add(10, 150, 100, 100)
    AssertFalse Registry_IsThinQQChart(co.Chart), "gráfico comum não é ThinQQ"
    CloseTemp wb
End Sub

Public Sub Test_Registry_Options_RoundTrip()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    CalcSheet_AllocateBlock wb, "opt123", "stacked", "x", 4, 3
    Set co = ws.ChartObjects.Add(10, 10, 100, 100)
    co.Name = ID_PREFIX & "opt123"
    AssertEqual "int", Registry_GetOption(co.Chart, "numfmt", "int")
    Registry_SetOption co.Chart, "numfmt", "dec1"
    Registry_SetOption co.Chart, "totals", "1"
    Registry_SetOption co.Chart, "numfmt", "pct0"
    AssertEqual "pct0", Registry_GetOption(co.Chart, "numfmt")
    AssertEqual "1", Registry_GetOption(co.Chart, "totals")
    AssertEqual "numfmt=pct0;totals=1", Registry_Block(co.Chart).Cells(1, 4).Value
    CloseTemp wb
End Sub

Public Sub Test_Registry_NewId_Format()
    Dim id As String, i As Long
    id = Registry_NewId()
    AssertEqual 6, Len(id)
    For i = 1 To 6
        AssertTrue InStr("abcdefghijklmnopqrstuvwxyz0123456789", Mid$(id, i, 1)) > 0
    Next
End Sub

Public Sub Test_Util_ColumnToDoubles()
    Dim wb As Workbook, ws As Worksheet, v() As Double
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:A3").Value = Application.Transpose(Array(1.5, "x", 3))
    v = Util_ColumnToDoubles(ws.Range("A1:A3"))
    AssertEqual 1, LBound(v): AssertEqual 3, UBound(v)
    AssertAlmost 1.5, v(1): AssertAlmost 0, v(2): AssertAlmost 3, v(3)
    v = Util_ColumnToDoubles(ws.Range("A3"))
    AssertEqual 1, UBound(v)
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: `not ok` em `TestCalcSheet.*`.

- [ ] **Passo 3: Implementar `modChartUtil.bas`**

```vb
Attribute VB_Name = "modChartUtil"
Option Explicit

Public Sub Util_EnsureDir(path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub
    If Not fso.FolderExists(fso.GetParentFolderName(path)) Then Util_EnsureDir fso.GetParentFolderName(path)
    fso.CreateFolder path
End Sub

Public Function Util_FileName(path As String) As String
    Util_FileName = Mid$(path, InStrRev(path, "\") + 1)
End Function

Public Function Util_ToDouble(v As Variant) As Double
    If IsNumeric(v) And Not IsEmpty(v) Then
        If VarType(v) <> vbString Then Util_ToDouble = CDbl(v) Else Util_ToDouble = Val(v)
    End If
End Function

' Coluna de células -> array 1..n de Double (texto/vazio = 0).
Public Function Util_ColumnToDoubles(rng As Range) As Double()
    Dim out() As Double, i As Long, n As Long
    n = rng.Rows.Count
    ReDim out(1 To n)
    For i = 1 To n
        out(i) = Util_ToDouble(rng.Cells(i, 1).Value)
    Next
    Util_ColumnToDoubles = out
End Function

Public Function Chart_FindSeries(cht As Chart, name As String) As Series
    Dim s As Series
    For Each s In cht.SeriesCollection
        If s.Name = name Then Set Chart_FindSeries = s: Exit Function
    Next
End Function
```

- [ ] **Passo 4: Implementar `modCalcSheet.bas`**

```vb
Attribute VB_Name = "modCalcSheet"
Option Explicit
' Planilha muito oculta "_ThinQQ". Cada gráfico tem um bloco: linha 1 = meta
' (A=id, B=tipo, C=endereço da fonte sem "=", D=opções "chave=valor;..."), demais linhas = corpo.
' O nome definido oculto "tqq_<id>" aponta para o bloco inteiro.

Public Function CalcSheet_Get(wb As Workbook, Optional createIfMissing As Boolean = True) As Worksheet
    Dim ws As Worksheet, active As Object, su As Boolean
    On Error Resume Next
    Set ws = wb.Worksheets(CALC_SHEET)
    On Error GoTo 0
    If ws Is Nothing And createIfMissing Then
        Set active = wb.ActiveSheet
        su = Application.ScreenUpdating
        Application.ScreenUpdating = False
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = CALC_SHEET
        ws.Range("A1").Value = "ThinQQ - planilha interna de cálculo. Não editar."
        ws.Visible = xlSheetVeryHidden
        If Not active Is Nothing Then active.Activate
        Application.ScreenUpdating = su
    End If
    Set CalcSheet_Get = ws
End Function

Public Function CalcSheet_AllocateBlock(wb As Workbook, id As String, chartType As String, sourceAddr As String, bodyRows As Long, bodyCols As Long) As Range
    Dim ws As Worksheet, r As Long, blk As Range
    Set ws = CalcSheet_Get(wb)
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 3
    Set blk = ws.Cells(r, 1).Resize(bodyRows + 1, IIf(bodyCols > 4, bodyCols, 4))
    blk.Cells(1, 1).NumberFormat = "@": blk.Cells(1, 1).Value = id
    blk.Cells(1, 2).Value = chartType
    ' Endereço como fórmula de texto (="...") para preservar o apóstrofo inicial de 'Nome'!$A$1
    blk.Cells(1, 3).Formula = "=""" & Replace(sourceAddr, """", "") & """"
    blk.Cells(1, 4).NumberFormat = "@": blk.Cells(1, 4).Value = ""
    blk.Rows(1).Font.Bold = True
    wb.Names.Add Name:=ID_PREFIX & id, RefersTo:="='" & CALC_SHEET & "'!" & blk.Address(True, True), Visible:=False
    Set CalcSheet_AllocateBlock = blk
End Function

Public Function CalcSheet_FindBlock(wb As Workbook, id As String) As Range
    On Error Resume Next
    Set CalcSheet_FindBlock = wb.Names(ID_PREFIX & id).RefersToRange
    On Error GoTo 0
End Function

Public Sub CalcSheet_DeleteBlock(wb As Workbook, id As String)
    Dim blk As Range
    Set blk = CalcSheet_FindBlock(wb, id)
    If blk Is Nothing Then Exit Sub
    blk.Clear
    wb.Names(ID_PREFIX & id).Delete
End Sub

Public Function Block_Meta(blk As Range) As Range
    Set Block_Meta = blk.Rows(1)
End Function

Public Function Block_Body(blk As Range) As Range
    Set Block_Body = blk.Offset(1, 0).Resize(blk.Rows.Count - 1, blk.Columns.Count)
End Function

' Remove blocos cujo gráfico não existe mais. Retorna quantos removeu.
Public Function CalcSheet_Cleanup(wb As Workbook) As Long
    Dim i As Long, nm As Name, id As String, removed As Long
    For i = wb.Names.Count To 1 Step -1
        Set nm = wb.Names(i)
        If Left$(nm.Name, Len(ID_PREFIX)) = ID_PREFIX Then
            id = Mid$(nm.Name, Len(ID_PREFIX) + 1)
            If Registry_FindChartObject(wb, id) Is Nothing Then
                CalcSheet_DeleteBlock wb, id
                removed = removed + 1
            End If
        End If
    Next
    CalcSheet_Cleanup = removed
End Function
```

- [ ] **Passo 5: Implementar `modChartRegistry.bas`**

```vb
Attribute VB_Name = "modChartRegistry"
Option Explicit

Public Function Registry_NewId() As String
    Dim chars As String, i As Long, s As String
    chars = "abcdefghijklmnopqrstuvwxyz0123456789"
    Randomize
    For i = 1 To 6
        s = s & Mid$(chars, Int(Rnd * Len(chars)) + 1, 1)
    Next
    Registry_NewId = s
End Function

Public Function Registry_IdFromChart(cht As Chart) As String
    Dim nm As String
    If TypeName(cht.Parent) <> "ChartObject" Then Exit Function
    nm = cht.Parent.Name
    If Left$(nm, Len(ID_PREFIX)) = ID_PREFIX Then Registry_IdFromChart = Mid$(nm, Len(ID_PREFIX) + 1)
End Function

Public Function Registry_IsThinQQChart(cht As Chart) As Boolean
    Registry_IsThinQQChart = Len(Registry_IdFromChart(cht)) > 0
End Function

Public Function Registry_Workbook(cht As Chart) As Workbook
    Set Registry_Workbook = cht.Parent.Parent.Parent   ' ChartObject -> Worksheet -> Workbook
End Function

Public Function Registry_Block(cht As Chart) As Range
    Dim id As String
    id = Registry_IdFromChart(cht)
    If Len(id) > 0 Then Set Registry_Block = CalcSheet_FindBlock(Registry_Workbook(cht), id)
End Function

Public Function Registry_ChartType(cht As Chart) As String
    Dim blk As Range
    Set blk = Registry_Block(cht)
    If Not blk Is Nothing Then Registry_ChartType = CStr(blk.Cells(1, 2).Value)
End Function

Public Function Registry_FindChartObject(wb As Workbook, id As String) As ChartObject
    Dim ws As Worksheet, co As ChartObject
    For Each ws In wb.Worksheets
        Set co = Nothing
        On Error Resume Next
        Set co = ws.ChartObjects(ID_PREFIX & id)
        On Error GoTo 0
        If Not co Is Nothing Then Set Registry_FindChartObject = co: Exit Function
    Next
End Function

Public Function Registry_GetOption(cht As Chart, key As String, Optional defaultValue As String = "") As String
    Dim blk As Range, parts() As String, i As Long, kv() As String
    Registry_GetOption = defaultValue
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Function
    parts = Split(CStr(blk.Cells(1, 4).Value), ";")
    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) = 1 Then
            If kv(0) = key Then Registry_GetOption = kv(1): Exit Function
        End If
    Next
End Function

Public Sub Registry_SetOption(cht As Chart, key As String, value As String)
    Dim blk As Range, parts() As String, i As Long, kv() As String, out As String, found As Boolean
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    parts = Split(CStr(blk.Cells(1, 4).Value), ";")
    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) = 1 Then
            If kv(0) = key Then parts(i) = key & "=" & value: found = True
            out = out & IIf(Len(out) > 0, ";", "") & parts(i)
        End If
    Next
    If Not found Then out = out & IIf(Len(out) > 0, ";", "") & key & "=" & value
    blk.Cells(1, 4).Value = out
End Sub
```

- [ ] **Passo 6: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: todos `ok`.

- [ ] **Passo 7: Commit**

```bash
git add src/core/modules/modCalcSheet.bas src/core/modules/modChartRegistry.bas src/core/modules/modChartUtil.bas tests/vba/core/TestCalcSheet.bas
git commit -m "feat(core): planilha de cálculo oculta, blocos por gráfico e registro de opções"
```

---

### Tarefa 7: Estilo base, paletas e formatos numéricos (core)

**Files:**
- Create: `src/core/modules/modChartStyle.bas`
- Test: `tests/vba/core/TestChartStyle.bas`

**Interfaces:**
- Produces: `Style_Palette(name) As Variant` (6 cores; `"padrao"`, `"cinza"`, `"tema"`; valores negativos = cor do tema accent1..6); `Style_SetColor(cf As ColorFormat, colorValue)`; `Style_IsHelperSeries(s) As Boolean` (nome começa com `tqq_`); `Style_ApplyBase(cht)` (chamar **após** adicionar séries); `Style_ApplySeriesColors(cht, paletteName)`; `Style_NumberFormat(key) As String` (`int|dec1|dec2|pct0|pct1|mil|mi`); `Style_ApplyNumberFormat(cht, fmt)`; `Style_ContrastFont(rgbValue) As Long`.

- [ ] **Passo 1: Escrever os testes**

`tests/vba/core/TestChartStyle.bas`:

```vb
Attribute VB_Name = "TestChartStyle"
Option Explicit

Private Function MakeChart(wb As Workbook) As Chart
    Dim ws As Worksheet, co As ChartObject, s As Series
    Set ws = wb.Worksheets(1)
    ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Cat", "S1", "S2")
    ws.Range("A2:C2").Value = Array("a", 1, 2)
    ws.Range("A3:C3").Value = Array("b", 3, 4)
    Set co = ws.ChartObjects.Add(100, 10, 300, 200)
    Set s = co.Chart.SeriesCollection.NewSeries: s.Name = "=Dados!$B$1": s.Values = "=Dados!$B$2:$B$3"
    Set s = co.Chart.SeriesCollection.NewSeries: s.Name = "tqq_total": s.Values = "=Dados!$C$2:$C$3"
    co.Chart.ChartType = xlColumnStacked
    Set MakeChart = co.Chart
End Function

Public Sub Test_Style_Palettes_HaveSixColors()
    AssertEqual 6, UBound(Style_Palette("padrao")) - LBound(Style_Palette("padrao")) + 1
    AssertEqual 6, UBound(Style_Palette("cinza")) - LBound(Style_Palette("cinza")) + 1
    AssertEqual 6, UBound(Style_Palette("tema")) - LBound(Style_Palette("tema")) + 1
    AssertEqual RGB(31, 56, 100), Style_Palette("padrao")(0)
    AssertEqual RGB(31, 56, 100), Style_Palette("desconhecida")(0), "nome desconhecido cai no padrão"
End Sub

Public Sub Test_Style_NumberFormats()
    AssertEqual "#,##0;-#,##0;", Style_NumberFormat("int")
    AssertEqual "#,##0.0;-#,##0.0;", Style_NumberFormat("dec1")
    AssertEqual "0.0%;-0.0%;", Style_NumberFormat("pct1")
    AssertEqual "#,##0,;-#,##0,;", Style_NumberFormat("mil")
    AssertEqual "#,##0;-#,##0;", Style_NumberFormat("qualquer")
End Sub

Public Sub Test_Style_ContrastFont()
    AssertEqual RGB(255, 255, 255), Style_ContrastFont(RGB(31, 56, 100))
    AssertEqual RGB(64, 64, 64), Style_ContrastFont(RGB(217, 217, 217))
End Sub

Public Sub Test_Style_ApplyBase_HidesAxisAndGrid()
    Dim wb As Workbook, cht As Chart
    Set wb = NewTempWorkbook(): Set cht = MakeChart(wb)
    Style_ApplyBase cht
    AssertFalse cht.Axes(xlValue).HasMajorGridlines
    AssertEqual xlTickLabelPositionNone, cht.Axes(xlValue).TickLabelPosition
    AssertEqual msoFalse, cht.Axes(xlValue).Format.Line.Visible
    AssertEqual xlTickMarkNone, cht.Axes(xlCategory).MajorTickMark
    AssertFalse cht.HasTitle
    AssertAlmost 0, cht.Axes(xlValue).MinimumScale + 0   ' eixo ainda acessível (não removido)
    CloseTemp wb
End Sub

Public Sub Test_Style_SeriesColors_SkipHelpers()
    Dim wb As Workbook, cht As Chart
    Set wb = NewTempWorkbook(): Set cht = MakeChart(wb)
    Style_ApplySeriesColors cht, "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(1).Format.Fill.ForeColor.RGB
    AssertTrue Style_IsHelperSeries(cht.SeriesCollection(2))
    AssertFalse Style_IsHelperSeries(cht.SeriesCollection(1))
    CloseTemp wb
End Sub

Public Sub Test_Style_ApplyNumberFormat_OnLabels()
    Dim wb As Workbook, cht As Chart
    Set wb = NewTempWorkbook(): Set cht = MakeChart(wb)
    cht.SeriesCollection(1).HasDataLabels = True
    Style_ApplyNumberFormat cht, Style_NumberFormat("dec2")
    AssertEqual "#,##0.00;-#,##0.00;", cht.SeriesCollection(1).DataLabels.NumberFormat
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: `not ok` em `TestChartStyle.*`.

- [ ] **Passo 3: Implementar `modChartStyle.bas`**

```vb
Attribute VB_Name = "modChartStyle"
Option Explicit
' Aparência padrão "limpa": sem grade, eixo de valores invisível (mas presente, para a geometria),
' sem bordas, fonte 9 pt. Paletas de 6 cores.

Public Function Style_Palette(name As String) As Variant
    Select Case LCase$(name)
        Case "cinza"
            Style_Palette = Array(RGB(64, 64, 64), RGB(127, 127, 127), RGB(166, 166, 166), RGB(191, 191, 191), RGB(217, 217, 217), RGB(89, 89, 89))
        Case "tema"
            Style_Palette = Array(-1, -2, -3, -4, -5, -6)    ' accent1..accent6 do tema
        Case Else
            Style_Palette = Array(RGB(31, 56, 100), RGB(68, 114, 196), RGB(157, 195, 230), RGB(165, 165, 165), RGB(0, 150, 136), RGB(237, 162, 36))
    End Select
End Function

Public Sub Style_SetColor(cf As ColorFormat, colorValue As Variant)
    If colorValue < 0 Then
        cf.ObjectThemeColor = msoThemeColorAccent1 + (-colorValue - 1)
    Else
        cf.RGB = CLng(colorValue)
    End If
End Sub

Public Function Style_IsHelperSeries(s As Series) As Boolean
    Style_IsHelperSeries = (Left$(s.Name, Len(HELPER_PREFIX)) = HELPER_PREFIX)
End Function

' Branco para fundos escuros, cinza-escuro para claros (luminância relativa simples).
Public Function Style_ContrastFont(rgbValue As Long) As Long
    Dim r As Long, g As Long, b As Long, lum As Double
    r = rgbValue And &HFF: g = (rgbValue \ &H100) And &HFF: b = (rgbValue \ &H10000) And &HFF
    lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255
    If lum < 0.55 Then Style_ContrastFont = RGB(255, 255, 255) Else Style_ContrastFont = RGB(64, 64, 64)
End Function

Public Sub Style_ApplyBase(cht As Chart)
    With cht
        .HasTitle = False
        .ChartArea.Format.Line.Visible = msoFalse
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Font.Size = 9
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
        .HasAxis(xlValue) = True
        With .Axes(xlValue)
            .HasMajorGridlines = False
            .HasMinorGridlines = False
            .MajorTickMark = xlTickMarkNone
            .MinorTickMark = xlTickMarkNone
            .TickLabelPosition = xlTickLabelPositionNone
            .Format.Line.Visible = msoFalse
        End With
        With .Axes(xlCategory)
            .MajorTickMark = xlTickMarkNone
            .MinorTickMark = xlTickMarkNone
            .Format.Line.ForeColor.RGB = RGB(166, 166, 166)
            .TickLabels.Font.Size = 9
            .TickLabels.Font.Color = RGB(64, 64, 64)
        End With
    End With
End Sub

Public Sub Style_ApplySeriesColors(cht As Chart, paletteName As String)
    Dim pal As Variant, s As Series, k As Long, c As Variant
    pal = Style_Palette(paletteName)
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            c = pal(k Mod 6)
            If s.ChartType = xlLine Or s.ChartType = xlLineMarkers Then
                Style_SetColor s.Format.Line.ForeColor, c
                s.Format.Line.Weight = 2
                s.MarkerStyle = xlMarkerStyleCircle
                s.MarkerSize = 5
                Style_SetColor s.Format.Fill.ForeColor, c
            Else
                Style_SetColor s.Format.Fill.ForeColor, c
                s.Format.Line.Visible = msoFalse
                If s.HasDataLabels And c >= 0 Then s.DataLabels.Font.Color = Style_ContrastFont(CLng(c))
            End If
            k = k + 1
        End If
    Next
End Sub

Public Function Style_NumberFormat(key As String) As String
    Select Case key
        Case "dec1": Style_NumberFormat = "#,##0.0;-#,##0.0;"
        Case "dec2": Style_NumberFormat = "#,##0.00;-#,##0.00;"
        Case "pct0": Style_NumberFormat = "0%;-0%;"
        Case "pct1": Style_NumberFormat = "0.0%;-0.0%;"
        Case "mil":  Style_NumberFormat = "#,##0,;-#,##0,;"
        Case "mi":   Style_NumberFormat = "#,##0.0,,;-#,##0.0,,;"
        Case Else:   Style_NumberFormat = "#,##0;-#,##0;"
    End Select
End Function

Public Sub Style_ApplyNumberFormat(cht As Chart, fmt As String)
    Dim s As Series
    For Each s In cht.SeriesCollection
        If s.HasDataLabels Then s.DataLabels.NumberFormat = fmt
    Next
End Sub
```

- [ ] **Passo 4: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: todos `ok`. Se `Test_Style_ApplyBase_HidesAxisAndGrid` falhar em `TickLabelPosition`, confirme que `.HasAxis(xlValue) = True` precede o `With .Axes(xlValue)`.

- [ ] **Passo 5: Commit**

```bash
git add src/core/modules/modChartStyle.bas tests/vba/core/TestChartStyle.bas
git commit -m "feat(core): estilo base limpo, paletas e formatos numéricos"
```

---
### Tarefa 8: Colunas empilhadas, agrupadas, 100% e linha, com totais, rótulos e conectores (core)

**Files:**
- Create: `src/core/modules/modLabels.bas`, `src/core/modules/modChartStacked.bas`
- Test: `tests/vba/core/TestStacked.bas`

**Interfaces:**
- Produces (modLabels): `Labels_LinkPoints(s As Series, valueCells As Range, numFmt As String)` (rótulo de cada ponto ligado por fórmula à célula; acima se ≥ 0, abaixo se < 0); `Labels_RefreshPositions(cht)` (reaplica acima/abaixo nas séries `tqq_total`/`tqq_lbl`).
- Produces (modChartStacked): `Stacked_BuildBlock(wb, id, data, kind) As Range` (corpo: linha 1 cabeçalhos `Categoria|Total|TotalPos`, linhas 2..n+1); `Stacked_Insert(ws, data, kind) As ChartObject` (`kind` ∈ `stacked|clustered|stacked100|line`); `Stacked_SetTotals(cht, turnOn)`; `Stacked_SetLabels(cht, turnOn)`; `Stacked_SetConnectors(cht, turnOn)`.
- Consumes: Tarefas 5, 6, 7.
- Opções gravadas no bloco: `numfmt`, `palette`, `gap`, `totals`, `labels`, `connectors`.

- [ ] **Passo 1: Escrever os testes**

`tests/vba/core/TestStacked.bas`:

```vb
Attribute VB_Name = "TestStacked"
Option Explicit

Private Function Fixture(wb As Workbook, Optional sheetName As String = "Dados") As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    ws.Name = sheetName
    ws.Range("A1:D1").Value = Array("Ano", "Sul", "Sudeste", "Nordeste")
    ws.Range("A2:D2").Value = Array(2022, 10, 20, 5)
    ws.Range("A3:D3").Value = Array(2023, 12, 22, 6)
    ws.Range("A4:D4").Value = Array(2024, 15, 25, 8)
    ws.Range("A5:D5").Value = Array(2025, 18, 30, 9)
    Set Fixture = ChartData_FromRange(ws.Range("A1:D5"))
End Function

Public Sub Test_Stacked_Insert_CreatesSeriesAndTotals()
    Dim wb As Workbook, d As clsChartData, co As ChartObject, cht As Chart, tot As Series, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set co = Stacked_Insert(d.Source.Worksheet, d, "stacked")
    Set cht = co.Chart
    AssertEqual ID_PREFIX, Left$(co.Name, 4)
    AssertEqual "stacked", Registry_ChartType(cht)
    AssertEqual 4, cht.SeriesCollection.Count
    AssertEqual "Sul", cht.SeriesCollection(1).Name
    AssertEqual xlColumnStacked, cht.SeriesCollection(1).ChartType
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertTrue Not tot Is Nothing, "série de totais ausente"
    AssertEqual xlLine, tot.ChartType
    AssertEqual msoFalse, tot.Format.Line.Visible
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual 35, body.Cells(2, 2).Value         ' 10+20+5
    AssertEqual 57, body.Cells(5, 2).Value         ' 18+30+9
    AssertTrue InStr(tot.Points(1).DataLabel.Formula, CALC_SHEET) > 0, "rótulo do total deve estar ligado ao bloco"
    AssertEqual xlLabelPositionAbove, tot.Points(1).DataLabel.Position
    AssertEqual "1", Registry_GetOption(cht, "totals")
    AssertEqual "1", Registry_GetOption(cht, "labels")
    AssertEqual "0", Registry_GetOption(cht, "connectors")
    AssertEqual "int", Registry_GetOption(cht, "numfmt")
    CloseTemp wb
End Sub

Public Sub Test_Stacked_SegmentLabels_CenteredWithFormat()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    AssertTrue cht.SeriesCollection(1).HasDataLabels
    AssertEqual xlLabelPositionCenter, cht.SeriesCollection(1).DataLabels.Position
    AssertEqual "#,##0;-#,##0;", cht.SeriesCollection(1).DataLabels.NumberFormat
    Stacked_SetLabels cht, False
    AssertFalse cht.SeriesCollection(1).HasDataLabels
    AssertEqual "0", Registry_GetOption(cht, "labels")
    CloseTemp wb
End Sub

Public Sub Test_Stacked_SetTotals_Off_RemovesSeries_On_Restores()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Stacked_SetTotals cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual 3, cht.SeriesCollection.Count
    Stacked_SetTotals cht, True
    AssertEqual 4, cht.SeriesCollection.Count
    CloseTemp wb
End Sub

Public Sub Test_Stacked_SetConnectors()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Stacked_SetConnectors cht, True
    AssertTrue cht.ChartGroups(1).HasSeriesLines
    AssertEqual "1", Registry_GetOption(cht, "connectors")
    Stacked_SetConnectors cht, False
    AssertFalse cht.ChartGroups(1).HasSeriesLines
    CloseTemp wb
End Sub

Public Sub Test_Stacked100_TotalPosIsOne()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked100").Chart
    AssertEqual xlColumnStacked100, cht.SeriesCollection(1).ChartType
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual 1, body.Cells(2, 3).Value
    AssertEqual 35, body.Cells(2, 2).Value
    AssertTrue Not Chart_FindSeries(cht, "tqq_total") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_Clustered_LabelsOutsideEnd_NoTotals()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "clustered").Chart
    AssertEqual xlColumnClustered, cht.SeriesCollection(1).ChartType
    AssertEqual xlLabelPositionOutsideEnd, cht.SeriesCollection(1).DataLabels.Position
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertEqual "80", Registry_GetOption(cht, "gap")
    CloseTemp wb
End Sub

Public Sub Test_Line_Insert()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "line").Chart
    AssertEqual xlLineMarkers, cht.SeriesCollection(1).ChartType
    AssertEqual xlLabelPositionAbove, cht.SeriesCollection(1).DataLabels.Position
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_Stacked_TextCellTreatedAsZero()
    Dim wb As Workbook, d As clsChartData, cht As Chart, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("C3").Value = "n/d"
    d.Source.Worksheet.Range("D4").ClearContents
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Set body = Block_Body(Registry_Block(cht))
    AssertEqual 18, body.Cells(3, 2).Value       ' 12 + "n/d" + 6
    AssertEqual 40, body.Cells(4, 2).Value       ' 15 + 25 + vazio
    CloseTemp wb
End Sub

Public Sub Test_Stacked_NegativeTotal_LabelBelow()
    Dim wb As Workbook, d As clsChartData, cht As Chart, tot As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    d.Source.Worksheet.Range("B2:D2").Value = Array(-10, -20, -5)
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    Set tot = Chart_FindSeries(cht, "tqq_total")
    AssertEqual xlLabelPositionBelow, tot.Points(1).DataLabel.Position
    AssertEqual xlLabelPositionAbove, tot.Points(2).DataLabel.Position
    d.Source.Worksheet.Range("B2:D2").Value = Array(10, 20, 5)
    Labels_RefreshPositions cht
    AssertEqual xlLabelPositionAbove, tot.Points(1).DataLabel.Position
    CloseTemp wb
End Sub

Public Sub Test_Stacked_Insert_SheetNameWithQuote()
    Dim wb As Workbook, d As clsChartData, cht As Chart
    Set wb = NewTempWorkbook(): Set d = Fixture(wb, "Plano 'A' 2026")
    Set cht = Stacked_Insert(d.Source.Worksheet, d, "stacked").Chart
    AssertEqual "Sul", cht.SeriesCollection(1).Name
    AssertEqual 35, Block_Body(Registry_Block(cht)).Cells(2, 2).Value
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: `not ok` em `TestStacked.*`.

- [ ] **Passo 3: Implementar `modLabels.bas`**

```vb
Attribute VB_Name = "modLabels"
Option Explicit

' Liga o rótulo de cada ponto à célula correspondente (rótulo "vivo") e posiciona pelo sinal.
Public Sub Labels_LinkPoints(s As Series, valueCells As Range, numFmt As String)
    Dim i As Long, n As Long, sheetRef As String, v As Variant
    n = valueCells.Rows.Count
    s.HasDataLabels = True
    sheetRef = "='" & Replace(valueCells.Worksheet.Name, "'", "''") & "'!"
    For i = 1 To n
        With s.Points(i)
            .HasDataLabel = True
            .DataLabel.Formula = sheetRef & valueCells.Cells(i, 1).Address(True, True)
            .DataLabel.NumberFormat = numFmt
            .DataLabel.Font.Size = 9
            .DataLabel.Font.Color = RGB(64, 64, 64)
            v = valueCells.Cells(i, 1).Value
            .DataLabel.Position = IIf(Util_ToDouble(v) < 0, xlLabelPositionBelow, xlLabelPositionAbove)
        End With
    Next
End Sub

' Reaplica acima/abaixo conforme o sinal atual das séries auxiliares de rótulo.
Public Sub Labels_RefreshPositions(cht As Chart)
    Dim blk As Range, body As Range, s As Series, col As Long, i As Long, n As Long
    Set blk = Registry_Block(cht)
    If blk Is Nothing Then Exit Sub
    Set body = Block_Body(blk)
    n = body.Rows.Count - 1
    If Registry_ChartType(cht) = "waterfall" Then
        Set s = Chart_FindSeries(cht, "tqq_lbl"): col = 17
    Else
        Set s = Chart_FindSeries(cht, "tqq_total"): col = 2
    End If
    If s Is Nothing Then Exit Sub
    If Not s.HasDataLabels Then Exit Sub
    For i = 1 To n
        If i <= s.Points.Count Then
            s.Points(i).DataLabel.Position = IIf(Util_ToDouble(body.Cells(i + 1, col).Value) < 0, xlLabelPositionBelow, xlLabelPositionAbove)
        End If
    Next
End Sub
```

- [ ] **Passo 4: Implementar `modChartStacked.bas`**

```vb
Attribute VB_Name = "modChartStacked"
Option Explicit
' Colunas empilhadas / agrupadas / 100% / linha.
' Bloco (corpo): linha 1 = cabeçalhos; colunas: 1 Categoria | 2 Total | 3 TotalPos (=Total, ou 1 no 100%).

Public Function Stacked_BuildBlock(wb As Workbook, id As String, data As clsChartData, kind As String) As Range
    Dim blk As Range, body As Range, i As Long, rowCells As Range
    Set blk = CalcSheet_AllocateBlock(wb, id, kind, Mid$(ChartData_ExternalAddress(data.Source), 2), data.CategoryCount + 1, 3)
    Set body = Block_Body(blk)
    body.Cells(1, 1).Value = "Categoria": body.Cells(1, 2).Value = "Total": body.Cells(1, 3).Value = "TotalPos"
    For i = 1 To data.CategoryCount
        body.Cells(i + 1, 1).Formula = ChartData_ExternalAddress(data.Source.Cells(i + 1, 1))
        Set rowCells = data.Source.Cells(i + 1, 2).Resize(1, data.SeriesCount)
        body.Cells(i + 1, 2).Formula = "=SUM(" & Mid$(ChartData_ExternalAddress(rowCells), 2) & ")"
        If kind = "stacked100" Then
            body.Cells(i + 1, 3).Formula = "=1"
        Else
            body.Cells(i + 1, 3).Formula = "=" & body.Cells(i + 1, 2).Address(False, False)
        End If
    Next
    Set Stacked_BuildBlock = blk
End Function

Public Function Stacked_Insert(ws As Worksheet, data As clsChartData, kind As String) As ChartObject
    Dim id As String, co As ChartObject, cht As Chart, s As Series, j As Long, anchor As Range
    id = Registry_NewId()
    Stacked_BuildBlock ws.Parent, id, data, kind
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, anchor.Top, 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    For j = 1 To data.SeriesCount
        Set s = cht.SeriesCollection.NewSeries
        s.Name = ChartData_ExternalAddress(data.SeriesNameCell(j))
        s.Values = ChartData_ExternalAddress(data.SeriesValueCells(j))
        s.XValues = ChartData_ExternalAddress(data.CategoryCells())
    Next
    Select Case kind
        Case "clustered": cht.ChartType = xlColumnClustered
        Case "stacked100": cht.ChartType = xlColumnStacked100
        Case "line": cht.ChartType = xlLineMarkers
        Case Else: cht.ChartType = xlColumnStacked
    End Select
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlZero
    Style_ApplyBase cht
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionRight
    cht.Legend.Font.Size = 9
    Registry_SetOption cht, "numfmt", "int"
    Registry_SetOption cht, "palette", "padrao"
    If kind <> "line" Then
        cht.ChartGroups(1).GapWidth = IIf(kind = "clustered", 80, 50)
        Registry_SetOption cht, "gap", IIf(kind = "clustered", "80", "50")
    End If
    Stacked_SetLabels cht, True
    Style_ApplySeriesColors cht, "padrao"
    Stacked_SetTotals cht, (kind = "stacked" Or kind = "stacked100")
    Stacked_SetConnectors cht, False
    Set Stacked_Insert = co
End Function

Public Sub Stacked_SetTotals(cht As Chart, turnOn As Boolean)
    Dim s As Series, body As Range, n As Long, fmt As String
    Set s = Chart_FindSeries(cht, "tqq_total")
    If Not turnOn Then
        If Not s Is Nothing Then s.Delete
        Registry_SetOption cht, "totals", "0"
        Exit Sub
    End If
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    If s Is Nothing Then
        Set s = cht.SeriesCollection.NewSeries
        s.Name = "tqq_total"
        s.Values = "='" & CALC_SHEET & "'!" & body.Cells(2, 3).Resize(n, 1).Address(True, True)
        s.ChartType = xlLine
        s.Format.Line.Visible = msoFalse
        s.MarkerStyle = xlMarkerStyleNone
        If cht.HasLegend Then
            On Error Resume Next
            cht.Legend.LegendEntries(cht.Legend.LegendEntries.Count).Delete
            On Error GoTo 0
        End If
    End If
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    Labels_LinkPoints s, body.Cells(2, 2).Resize(n, 1), fmt
    s.DataLabels.Font.Bold = True
    Registry_SetOption cht, "totals", "1"
End Sub

Public Sub Stacked_SetLabels(cht As Chart, turnOn As Boolean)
    Dim s As Series, fmt As String
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    For Each s In cht.SeriesCollection
        If Not Style_IsHelperSeries(s) Then
            s.HasDataLabels = turnOn
            If turnOn Then
                With s.DataLabels
                    .ShowValue = True: .ShowSeriesName = False: .ShowCategoryName = False
                    .NumberFormat = fmt
                    .Font.Size = 9
                    If s.ChartType = xlLine Or s.ChartType = xlLineMarkers Then
                        .Position = xlLabelPositionAbove
                        .Font.Color = RGB(64, 64, 64)
                    ElseIf s.ChartType = xlColumnClustered Then
                        .Position = xlLabelPositionOutsideEnd
                        .Font.Color = RGB(64, 64, 64)
                    Else
                        .Position = xlLabelPositionCenter
                        .Font.Color = Style_ContrastFont(s.Format.Fill.ForeColor.RGB)
                    End If
                End With
            End If
        End If
    Next
    Registry_SetOption cht, "labels", IIf(turnOn, "1", "0")
End Sub

' Linhas de série (conectores entre segmentos). Grupos sem suporte (linha/agrupadas) são ignorados.
Public Sub Stacked_SetConnectors(cht As Chart, turnOn As Boolean)
    Dim g As ChartGroup
    On Error Resume Next
    For Each g In cht.ChartGroups
        g.HasSeriesLines = turnOn
        If turnOn Then
            g.SeriesLines.Format.Line.ForeColor.RGB = RGB(166, 166, 166)
            g.SeriesLines.Format.Line.Weight = 0.75
        End If
    Next
    On Error GoTo 0
    Registry_SetOption cht, "connectors", IIf(turnOn, "1", "0")
End Sub
```

- [ ] **Passo 5: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: todos `ok`. Se `DataLabel.Formula` lançar erro 1004 nesta versão do Excel, substitua em `Labels_LinkPoints` o laço por uma única chamada por série: `s.DataLabels.Format.TextFrame2.TextRange.InsertChartField msoChartFieldRange, "='" & nome & "'!" & valueCells.Address(True, True), 0` seguida de `s.DataLabels.ShowRange = True: s.DataLabels.ShowValue = False`, mantendo o laço só para `Position`.

- [ ] **Passo 6: Commit**

```bash
git add src/core/modules/modLabels.bas src/core/modules/modChartStacked.bas tests/vba/core/TestStacked.bas
git commit -m "feat(core): colunas empilhadas/agrupadas/100%/linha com totais, rótulos vivos e conectores"
```

---

### Tarefa 9: Cascata (waterfall) que cruza o zero (core)

**Files:**
- Create: `src/core/modules/modChartWaterfall.bas`
- Test: `tests/vba/core/TestWaterfall.bas`

**Interfaces:**
- Produces: constantes `WF_COL_CAT=1`, `WF_COL_END=6`, `WF_COL_BASE=9`, `WF_COL_LABELVAL=17`, `WF_COL_LABELPOS=18`; `Waterfall_IsTotalMarker(v) As Boolean` (`e`, `t`, `=`, `total`, `subtotal`, sem distinção de maiúsculas); `Waterfall_BuildBlock(wb, id, data) As Range` (corpo com 18 colunas, ver spec §4.3); `Waterfall_Insert(ws, data) As ChartObject` (séries `tqq_base, tqq_uppos, tqq_upneg, tqq_downpos, tqq_downneg, tqq_total` + `tqq_lbl`); `Waterfall_SetLabels(cht, turnOn)`; `Waterfall_ApplyColors(cht, paletteName)`; `Waterfall_EndValues(cht) As Double()`.
- Consumes: Tarefas 5–8. Usa apenas a **primeira** série de valores da tabela.

- [ ] **Passo 1: Escrever os testes**

`tests/vba/core/TestWaterfall.bas`:

```vb
Attribute VB_Name = "TestWaterfall"
Option Explicit

Private Function Fixture(wb As Workbook) As clsChartData
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    ws.Name = "Ponte"
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Início", 100)
    ws.Range("A3:B3").Value = Array("A", -150)
    ws.Range("A4:B4").Value = Array("B", 30)
    ws.Range("A5:B5").Value = Array("Subtotal", "e")
    ws.Range("A6:B6").Value = Array("C", 70)
    ws.Range("A7:B7").Value = Array("Fim", "E")
    Set Fixture = ChartData_FromRange(ws.Range("A1:B7"))
End Function

Public Sub Test_Waterfall_IsTotalMarker()
    AssertTrue Waterfall_IsTotalMarker("e")
    AssertTrue Waterfall_IsTotalMarker("E")
    AssertTrue Waterfall_IsTotalMarker(" total ")
    AssertTrue Waterfall_IsTotalMarker("=")
    AssertFalse Waterfall_IsTotalMarker(12)
    AssertFalse Waterfall_IsTotalMarker("")
    AssertFalse Waterfall_IsTotalMarker("abc")
End Sub

Public Sub Test_Waterfall_Block_CrossingZero()
    Dim wb As Workbook, d As clsChartData, body As Range, r As Long
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0001", d))
    ' End: 100, -50, -20, -20, 50, 50
    AssertEqual 100, body.Cells(2, WF_COL_END).Value
    AssertEqual -50, body.Cells(3, WF_COL_END).Value
    AssertEqual -20, body.Cells(4, WF_COL_END).Value
    AssertEqual -20, body.Cells(5, WF_COL_END).Value
    AssertEqual 50, body.Cells(6, WF_COL_END).Value
    AssertEqual 50, body.Cells(7, WF_COL_END).Value
    ' Linha A (100 -> -50, desce cruzando o zero): Base 0, DownPos 100, DownNeg -50
    r = 3
    AssertEqual 0, body.Cells(r, WF_COL_BASE).Value
    AssertEqual 100, body.Cells(r, 14).Value
    AssertEqual -50, body.Cells(r, 15).Value
    AssertEqual 0, body.Cells(r, 12).Value
    ' Linha B (-50 -> -20, sobe no negativo): Base -20, UpNeg -30
    r = 4
    AssertEqual -20, body.Cells(r, WF_COL_BASE).Value
    AssertEqual -30, body.Cells(r, 13).Value
    AssertEqual 0, body.Cells(r, 12).Value
    ' Subtotal (e): Total -20, rótulo -20 posicionado em Lo (-20)
    r = 5
    AssertEqual True, body.Cells(r, 3).Value
    AssertEqual -20, body.Cells(r, 16).Value
    AssertEqual -20, body.Cells(r, WF_COL_LABELVAL).Value
    AssertEqual -20, body.Cells(r, WF_COL_LABELPOS).Value
    ' Linha C (-20 -> 50, sobe cruzando o zero): Base 0, UpPos 50, UpNeg -20; rótulo 70 em Hi (50)
    r = 6
    AssertEqual 0, body.Cells(r, WF_COL_BASE).Value
    AssertEqual 50, body.Cells(r, 12).Value
    AssertEqual -20, body.Cells(r, 13).Value
    AssertEqual 70, body.Cells(r, WF_COL_LABELVAL).Value
    AssertEqual 50, body.Cells(r, WF_COL_LABELPOS).Value
    ' Início: Base 0, UpPos 100
    AssertEqual 0, body.Cells(2, WF_COL_BASE).Value
    AssertEqual 100, body.Cells(2, 12).Value
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Block_LiveUpdate()
    Dim wb As Workbook, d As clsChartData, body As Range
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set body = Block_Body(Waterfall_BuildBlock(wb, "wf0002", d))
    d.Source.Worksheet.Range("B3").Value = -10
    AssertEqual 190, body.Cells(7, WF_COL_END).Value     ' 100-10+30+70
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_Insert_Series()
    Dim wb As Workbook, d As clsChartData, cht As Chart, lbl As Series
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    AssertEqual "waterfall", Registry_ChartType(cht)
    AssertEqual 7, cht.SeriesCollection.Count
    AssertEqual "tqq_base", cht.SeriesCollection(1).Name
    AssertEqual xlColumnStacked, cht.SeriesCollection(1).ChartType
    AssertEqual msoFalse, cht.SeriesCollection(1).Format.Fill.Visible
    AssertFalse cht.HasLegend
    Set lbl = Chart_FindSeries(cht, "tqq_lbl")
    AssertEqual xlLine, lbl.ChartType
    AssertTrue InStr(lbl.Points(3).DataLabel.Formula, CALC_SHEET) > 0
    AssertEqual xlLabelPositionBelow, lbl.Points(4).DataLabel.Position   ' subtotal -20
    AssertEqual xlLabelPositionAbove, lbl.Points(5).DataLabel.Position   ' +70
    AssertEqual "1", Registry_GetOption(cht, "labels")
    CloseTemp wb
End Sub

Public Sub Test_Waterfall_EndValues_And_LabelsOff()
    Dim wb As Workbook, d As clsChartData, cht As Chart, v() As Double
    Set wb = NewTempWorkbook(): Set d = Fixture(wb)
    Set cht = Waterfall_Insert(d.Source.Worksheet, d).Chart
    v = Waterfall_EndValues(cht)
    AssertEqual 6, UBound(v)
    AssertAlmost -50, v(2)
    AssertAlmost 50, v(6)
    Waterfall_SetLabels cht, False
    AssertTrue Chart_FindSeries(cht, "tqq_lbl") Is Nothing
    AssertEqual 6, cht.SeriesCollection.Count
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: `not ok` em `TestWaterfall.*`.

- [ ] **Passo 3: Implementar `modChartWaterfall.bas`**

```vb
Attribute VB_Name = "modChartWaterfall"
Option Explicit
' Bloco (corpo): linha 1 = cabeçalhos; 18 colunas:
'  1 Cat | 2 Src | 3 IsTotal | 4 Delta | 5 Start | 6 End | 7 Lo | 8 Hi | 9 Base | 10 PiecePos | 11 PieceNeg
' 12 UpPos | 13 UpNeg | 14 DownPos | 15 DownNeg | 16 Total | 17 LabelVal | 18 LabelPos
Public Const WF_COL_CAT As Long = 1
Public Const WF_COL_END As Long = 6
Public Const WF_COL_BASE As Long = 9
Public Const WF_COL_LABELVAL As Long = 17
Public Const WF_COL_LABELPOS As Long = 18

Public Function Waterfall_IsTotalMarker(v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString Then Exit Function
    Select Case LCase$(Trim$(CStr(v)))
        Case "e", "t", "=", "total", "subtotal": Waterfall_IsTotalMarker = True
    End Select
End Function

Public Function Waterfall_BuildBlock(wb As Workbook, id As String, data As clsChartData) As Range
    Dim blk As Range, body As Range, n As Long, i As Long, hdr As Variant, c As Long
    n = data.CategoryCount
    Set blk = CalcSheet_AllocateBlock(wb, id, "waterfall", Mid$(ChartData_ExternalAddress(data.Source), 2), n + 1, 18)
    Set body = Block_Body(blk)
    hdr = Array("Cat", "Src", "IsTotal", "Delta", "Start", "End", "Lo", "Hi", "Base", "PiecePos", "PieceNeg", _
                "UpPos", "UpNeg", "DownPos", "DownNeg", "Total", "LabelVal", "LabelPos")
    For c = 0 To 17
        body.Cells(1, c + 1).Value = hdr(c)
    Next
    For i = 1 To n
        body.Cells(i + 1, 1).Formula = ChartData_ExternalAddress(data.Source.Cells(i + 1, 1))
        body.Cells(i + 1, 2).Formula = ChartData_ExternalAddress(data.ValueCell(i, 1))
        body.Cells(i + 1, 5).FormulaR1C1 = IIf(i = 1, "=0", "=R[-1]C[1]")
    Next
    With body.Offset(1, 0).Resize(n, 1)
        .Offset(0, 2).FormulaR1C1 = "=IF(ISNUMBER(RC[-1]),FALSE,OR(LOWER(TRIM(RC[-1]&""""))={""e"",""t"",""="",""total"",""subtotal""}))"
        .Offset(0, 3).FormulaR1C1 = "=IF(RC[-1],0,IF(ISNUMBER(RC[-2]),RC[-2],0))"
        .Offset(0, 5).FormulaR1C1 = "=RC[-1]+RC[-2]"
        .Offset(0, 6).FormulaR1C1 = "=IF(RC[-4],MIN(0,RC[-1]),MIN(RC[-2],RC[-1]))"
        .Offset(0, 7).FormulaR1C1 = "=IF(RC[-5],MAX(0,RC[-2]),MAX(RC[-3],RC[-2]))"
        .Offset(0, 8).FormulaR1C1 = "=MAX(0,RC[-2])+MIN(0,RC[-1])"
        .Offset(0, 9).FormulaR1C1 = "=MAX(0,RC[-2])-MAX(0,RC[-3])"
        .Offset(0, 10).FormulaR1C1 = "=MIN(0,RC[-4])-MIN(0,RC[-3])"
        .Offset(0, 11).FormulaR1C1 = "=IF(AND(NOT(RC[-9]),RC[-8]>=0),RC[-2],0)"
        .Offset(0, 12).FormulaR1C1 = "=IF(AND(NOT(RC[-10]),RC[-9]>=0),RC[-2],0)"
        .Offset(0, 13).FormulaR1C1 = "=IF(AND(NOT(RC[-11]),RC[-10]<0),RC[-4],0)"
        .Offset(0, 14).FormulaR1C1 = "=IF(AND(NOT(RC[-12]),RC[-11]<0),RC[-4],0)"
        .Offset(0, 15).FormulaR1C1 = "=IF(RC[-13],RC[-10],0)"
        .Offset(0, 16).FormulaR1C1 = "=IF(RC[-14],RC[-11],RC[-13])"
        .Offset(0, 17).FormulaR1C1 = "=IF(RC[-1]>=0,RC[-10],RC[-11])"
    End With
    Set Waterfall_BuildBlock = blk
End Function

Public Function Waterfall_Insert(ws As Worksheet, data As clsChartData) As ChartObject
    Dim id As String, body As Range, n As Long, co As ChartObject, cht As Chart, anchor As Range
    Dim names As Variant, cols As Variant, k As Long, s As Series, sh As String
    id = Registry_NewId()
    Set body = Block_Body(Waterfall_BuildBlock(ws.Parent, id, data))
    n = data.CategoryCount
    sh = "='" & CALC_SHEET & "'!"
    Set anchor = data.Source.Cells(1, data.Source.Columns.Count + 2)
    Set co = ws.ChartObjects.Add(anchor.Left, anchor.Top, 460, 280)
    co.Name = ID_PREFIX & id
    Set cht = co.Chart
    Do While cht.SeriesCollection.Count > 0
        cht.SeriesCollection(1).Delete
    Loop
    names = Array("tqq_base", "tqq_uppos", "tqq_upneg", "tqq_downpos", "tqq_downneg", "tqq_total")
    cols = Array(9, 12, 13, 14, 15, 16)
    For k = 0 To 5
        Set s = cht.SeriesCollection.NewSeries
        s.Name = names(k)
        s.Values = sh & body.Cells(2, cols(k)).Resize(n, 1).Address(True, True)
        s.XValues = ChartData_ExternalAddress(data.CategoryCells())
    Next
    cht.ChartType = xlColumnStacked
    cht.PlotVisibleOnly = False
    cht.DisplayBlanksAs = xlZero
    cht.ChartGroups(1).GapWidth = 50
    Style_ApplyBase cht
    cht.HasLegend = False
    Waterfall_ApplyColors cht, "padrao"
    Registry_SetOption cht, "numfmt", "int"
    Registry_SetOption cht, "palette", "padrao"
    Registry_SetOption cht, "gap", "50"
    Registry_SetOption cht, "totals", "0"
    Registry_SetOption cht, "connectors", "0"
    Waterfall_SetLabels cht, True
    Set Waterfall_Insert = co
End Function

Public Sub Waterfall_ApplyColors(cht As Chart, paletteName As String)
    Dim pal As Variant, s As Series
    pal = Style_Palette(paletteName)
    For Each s In cht.SeriesCollection
        Select Case s.Name
            Case "tqq_base"
                s.Format.Fill.Visible = msoFalse
                s.Format.Line.Visible = msoFalse
            Case "tqq_uppos", "tqq_upneg"
                Style_SetColor s.Format.Fill.ForeColor, pal(1): s.Format.Line.Visible = msoFalse
            Case "tqq_downpos", "tqq_downneg"
                Style_SetColor s.Format.Fill.ForeColor, pal(2): s.Format.Line.Visible = msoFalse
            Case "tqq_total"
                Style_SetColor s.Format.Fill.ForeColor, pal(0): s.Format.Line.Visible = msoFalse
        End Select
    Next
End Sub

Public Sub Waterfall_SetLabels(cht As Chart, turnOn As Boolean)
    Dim s As Series, body As Range, n As Long, sh As String, fmt As String
    Set s = Chart_FindSeries(cht, "tqq_lbl")
    If Not turnOn Then
        If Not s Is Nothing Then s.Delete
        Registry_SetOption cht, "labels", "0"
        Exit Sub
    End If
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    sh = "='" & CALC_SHEET & "'!"
    If s Is Nothing Then
        Set s = cht.SeriesCollection.NewSeries
        s.Name = "tqq_lbl"
        s.Values = sh & body.Cells(2, WF_COL_LABELPOS).Resize(n, 1).Address(True, True)
        s.ChartType = xlLine
        s.Format.Line.Visible = msoFalse
        s.MarkerStyle = xlMarkerStyleNone
    End If
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    Labels_LinkPoints s, body.Cells(2, WF_COL_LABELVAL).Resize(n, 1), fmt
    Registry_SetOption cht, "labels", "1"
End Sub

Public Function Waterfall_EndValues(cht As Chart) As Double()
    Dim body As Range, n As Long
    Set body = Block_Body(Registry_Block(cht))
    n = body.Rows.Count - 1
    Waterfall_EndValues = Util_ColumnToDoubles(body.Cells(2, WF_COL_END).Resize(n, 1))
End Function
```

- [ ] **Passo 4: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: todos `ok`.

- [ ] **Passo 5: Commit**

```bash
git add src/core/modules/modChartWaterfall.bas tests/vba/core/TestWaterfall.bas
git commit -m "feat(core): gráfico de cascata com subtotais 'e', cruzamento do zero e rótulos vivos"
```

---

### Tarefa 10: Geometria e anotações (CAGR, diferença, conectores da cascata) (core)

**Files:**
- Create: `src/core/modules/modGeometry.bas`, `src/core/modules/modAnnotations.bas`
- Modify: `src/core/modules/modChartWaterfall.bas` (chamar `Ann_SetWaterfallConnectors cht, True` em `Waterfall_Insert`)
- Test: `tests/vba/core/TestGeometry.bas`, `tests/vba/core/TestAnnotations.bas`

**Interfaces:**
- Produces (modGeometry): puras `Geom_SlotWidth(insideWidth, nCats)`, `Geom_CategoryCenterX(insideLeft, insideWidth, nCats, idx)`, `Geom_BarWidth(slotWidth, gapWidthPct)`, `Geom_ValueToY(insideTop, insideHeight, axisMin, axisMax, v)`, `Geom_Periods(catFrom, catTo, idxFrom, idxTo) As Double`; ligadas ao gráfico `Geom_ChartCategoryCount(cht)`, `Geom_ChartCategoryCenterX(cht, idx)`, `Geom_ChartValueToY(cht, v)`, `Geom_ChartBarHalfWidth(cht)`.
- Produces (modAnnotations): `Ann_TryCagr(startVal, endVal, periods, ByRef rate) As Boolean`, `Ann_ParseParam(altText, key) As String`, `Ann_GetCategoryData(cht, ByRef values() As Double, ByRef tops() As Double)`, `Ann_AddCagr(cht, fromIdx, toIdx) As Shape`, `Ann_AddDifference(cht, fromIdx, toIdx) As Shape`, `Ann_SetWaterfallConnectors(cht, turnOn)`, `Ann_RemoveAll(cht, Optional includeConnectors = False)`, `Ann_Refresh(cht)`.
- Formas: nome `tqq_ann_<kind>_<n>` (`cagr`, `diff`, `conn`), `AlternativeText = "kind=…;from=…;to=…"`; rótulo `…_lbl`, linha de nível `…_lvl`.
- Em `clustered` e `line` as setas usam a **primeira série**; em `stacked`/`stacked100` usam a soma; na cascata usam `End`.

- [ ] **Passo 1: Escrever os testes de geometria**

`tests/vba/core/TestGeometry.bas`:

```vb
Attribute VB_Name = "TestGeometry"
Option Explicit

Public Sub Test_Geom_CategoryCenterX()
    AssertAlmost 110, Geom_CategoryCenterX(100, 400, 4, 1)
    AssertAlmost 390, Geom_CategoryCenterX(100, 400, 4, 4)
End Sub

Public Sub Test_Geom_BarWidth()
    AssertAlmost 100, Geom_SlotWidth(400, 4)
    AssertAlmost 66.6666667, Geom_BarWidth(100, 50)
    AssertAlmost 100, Geom_BarWidth(100, 0)
End Sub

Public Sub Test_Geom_ValueToY()
    AssertAlmost 50, Geom_ValueToY(50, 200, 0, 100, 100)        ' topo
    AssertAlmost 250, Geom_ValueToY(50, 200, 0, 100, 0)         ' base
    AssertAlmost 150, Geom_ValueToY(50, 200, -100, 100, 0)      ' zero no meio
    AssertAlmost 250, Geom_ValueToY(50, 200, 5, 5, 5)           ' escala degenerada
End Sub

Public Sub Test_Geom_Periods()
    AssertAlmost 3, Geom_Periods(2022, 2025, 1, 4)
    AssertAlmost 3, Geom_Periods("2022", "2025", 1, 4)
    AssertAlmost 2, Geom_Periods("T1", "T3", 1, 3)
    AssertAlmost 2, Geom_Periods(DateSerial(2020, 1, 1), DateSerial(2022, 1, 1), 1, 2), 0.01
    AssertAlmost 3, Geom_Periods(10, 40, 1, 4)                  ' números fora da faixa de anos: usa índices
End Sub
```

- [ ] **Passo 2: Escrever os testes de anotações**

`tests/vba/core/TestAnnotations.bas`:

```vb
Attribute VB_Name = "TestAnnotations"
Option Explicit

Private Function StackedFixture(wb As Workbook) As Chart
    Dim ws As Worksheet, d As clsChartData
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 100, 50)
    ws.Range("A3:C3").Value = Array(2023, 110, 60)
    ws.Range("A4:C4").Value = Array(2024, 120, 70)
    ws.Range("A5:C5").Value = Array(2025, 150, 90)
    Set d = ChartData_FromRange(ws.Range("A1:C5"))
    Set StackedFixture = Stacked_Insert(ws, d, "stacked").Chart
End Function

Private Function WaterfallFixture(wb As Workbook) As Chart
    Dim ws As Worksheet, d As clsChartData
    Set ws = wb.Worksheets(1): ws.Name = "Ponte"
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Início", 100)
    ws.Range("A3:B3").Value = Array("A", -30)
    ws.Range("A4:B4").Value = Array("B", 50)
    ws.Range("A5:B5").Value = Array("Fim", "e")
    Set d = ChartData_FromRange(ws.Range("A1:B5"))
    Set WaterfallFixture = Waterfall_Insert(ws, d).Chart
End Function

Public Sub Test_Ann_TryCagr()
    Dim r As Double
    AssertTrue Ann_TryCagr(100, 200, 2, r)
    AssertAlmost 0.414213562, r
    AssertFalse Ann_TryCagr(0, 200, 2, r)
    AssertFalse Ann_TryCagr(100, -50, 2, r)
    AssertFalse Ann_TryCagr(100, 200, 0, r)
End Sub

Public Sub Test_Ann_ParseParam()
    AssertEqual "cagr", Ann_ParseParam("kind=cagr;from=1;to=4", "kind")
    AssertEqual "4", Ann_ParseParam("kind=cagr;from=1;to=4", "to")
    AssertEqual "", Ann_ParseParam("kind=cagr;from=1;to=4", "x")
End Sub

Public Sub Test_Ann_CategoryData_Stacked()
    Dim wb As Workbook, cht As Chart, v() As Double, t() As Double
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    Ann_GetCategoryData cht, v, t
    AssertEqual 4, UBound(v)
    AssertAlmost 150, v(1): AssertAlmost 240, v(4)
    AssertAlmost 150, t(1)
    CloseTemp wb
End Sub

Public Sub Test_Ann_CategoryData_Waterfall()
    Dim wb As Workbook, cht As Chart, v() As Double, t() As Double
    Set wb = NewTempWorkbook(): Set cht = WaterfallFixture(wb)
    Ann_GetCategoryData cht, v, t
    AssertAlmost 100, v(1): AssertAlmost 70, v(2): AssertAlmost 120, v(3): AssertAlmost 120, v(4)
    CloseTemp wb
End Sub

Public Sub Test_Ann_AddCagr_CreatesShapesAndLabel()
    Dim wb As Workbook, cht As Chart, ln As Shape, lbl As Shape
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    Set ln = Ann_AddCagr(cht, 1, 4)
    AssertEqual "tqq_ann_cagr_1", ln.Name
    AssertEqual "kind=cagr;from=1;to=4", ln.AlternativeText
    Set lbl = cht.Shapes("tqq_ann_cagr_1_lbl")
    AssertEqual "CAGR +17,0%", lbl.TextFrame2.TextRange.Text      ' (240/150)^(1/3)-1 = 16,96%
    AssertAlmost Geom_ChartCategoryCenterX(cht, 1), ln.Left, 0.5
    AssertAlmost Geom_ChartCategoryCenterX(cht, 4) - Geom_ChartCategoryCenterX(cht, 1), ln.Width, 0.5
    Set ln = Ann_AddCagr(cht, 2, 3)
    AssertEqual "tqq_ann_cagr_2", ln.Name
    CloseTemp wb
End Sub

Public Sub Test_Ann_AddDifference_LabelText()
    Dim wb As Workbook, cht As Chart, ln As Shape
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    Set ln = Ann_AddDifference(cht, 1, 4)
    AssertEqual "+90 (+60%)", cht.Shapes("tqq_ann_diff_1_lbl").TextFrame2.TextRange.Text
    AssertTrue Not cht.Shapes("tqq_ann_diff_1_lvl") Is Nothing
    AssertAlmost 0, ln.Width, 0.01, "seta de diferença é vertical"
    CloseTemp wb
End Sub

Public Sub Test_Ann_Refresh_FollowsData()
    Dim wb As Workbook, cht As Chart, lbl As Shape
    Set wb = NewTempWorkbook(): Set cht = StackedFixture(wb)
    Ann_AddCagr cht, 1, 4
    wb.Worksheets("Dados").Range("B5").Value = 250          ' total 2025 = 340
    Ann_Refresh cht
    Set lbl = cht.Shapes("tqq_ann_cagr_1_lbl")
    AssertEqual "CAGR +31,4%", lbl.TextFrame2.TextRange.Text   ' (340/150)^(1/3)-1
    CloseTemp wb
End Sub

Public Sub Test_Ann_WaterfallConnectors()
    Dim wb As Workbook, cht As Chart, i As Long, n As Long
    Set wb = NewTempWorkbook(): Set cht = WaterfallFixture(wb)
    AssertEqual "1", Registry_GetOption(cht, "connectors")
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, 13) = "tqq_ann_conn_" Then n = n + 1
    Next
    AssertEqual 3, n
    AssertAlmost Geom_ChartValueToY(cht, 100), cht.Shapes("tqq_ann_conn_1").Top, 0.5
    Ann_SetWaterfallConnectors cht, False
    n = 0
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, 13) = "tqq_ann_conn_" Then n = n + 1
    Next
    AssertEqual 0, n
    CloseTemp wb
End Sub

Public Sub Test_Ann_RemoveAll_KeepsConnectors()
    Dim wb As Workbook, cht As Chart, i As Long, conn As Long, others As Long
    Set wb = NewTempWorkbook(): Set cht = WaterfallFixture(wb)
    Ann_AddCagr cht, 1, 4
    Ann_AddDifference cht, 1, 4
    Ann_RemoveAll cht
    For i = 1 To cht.Shapes.Count
        If Left$(cht.Shapes(i).Name, 13) = "tqq_ann_conn_" Then conn = conn + 1 Else others = others + 1
    Next
    AssertEqual 3, conn
    AssertEqual 0, others
    Ann_RemoveAll cht, True
    AssertEqual 0, cht.Shapes.Count
    CloseTemp wb
End Sub
```

- [ ] **Passo 3: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: `not ok` em `TestGeometry.*` e `TestAnnotations.*`.

- [ ] **Passo 4: Implementar `modGeometry.bas`**

```vb
Attribute VB_Name = "modGeometry"
Option Explicit
' Converte categoria/valor em coordenadas (pontos) dentro do gráfico, a partir da área de plotagem
' interna e da escala do eixo de valores. Funções puras primeiro (testáveis sem gráfico).

Public Function Geom_SlotWidth(insideWidth As Double, nCats As Long) As Double
    Geom_SlotWidth = insideWidth / nCats
End Function

Public Function Geom_CategoryCenterX(insideLeft As Double, insideWidth As Double, nCats As Long, idx As Long) As Double
    Geom_CategoryCenterX = insideLeft + (idx - 0.5) * insideWidth / nCats
End Function

Public Function Geom_BarWidth(slotWidth As Double, gapWidthPct As Double) As Double
    Geom_BarWidth = slotWidth / (1 + gapWidthPct / 100)
End Function

Public Function Geom_ValueToY(insideTop As Double, insideHeight As Double, axisMin As Double, axisMax As Double, v As Double) As Double
    If axisMax = axisMin Then
        Geom_ValueToY = insideTop + insideHeight
    Else
        Geom_ValueToY = insideTop + insideHeight * (axisMax - v) / (axisMax - axisMin)
    End If
End Function

' Períodos entre duas categorias: anos se forem datas ou anos (1900..2200); senão diferença de índices.
Public Function Geom_Periods(catFrom As Variant, catTo As Variant, idxFrom As Long, idxTo As Long) As Double
    Dim p As Double
    If VarType(catFrom) = vbDate And VarType(catTo) = vbDate Then
        p = (CDate(catTo) - CDate(catFrom)) / 365.25
    ElseIf IsNumeric(catFrom) And IsNumeric(catTo) And Not IsEmpty(catFrom) And Not IsEmpty(catTo) Then
        If CDbl(catFrom) >= 1900 And CDbl(catFrom) <= 2200 And CDbl(catTo) >= 1900 And CDbl(catTo) <= 2200 Then
            p = CDbl(catTo) - CDbl(catFrom)
        End If
    End If
    If p <= 0 Then p = idxTo - idxFrom
    Geom_Periods = p
End Function

Public Function Geom_ChartCategoryCount(cht As Chart) As Long
    If cht.SeriesCollection.Count = 0 Then Exit Function
    Geom_ChartCategoryCount = cht.SeriesCollection(1).Points.Count
End Function

Public Function Geom_ChartCategoryCenterX(cht As Chart, idx As Long) As Double
    With cht.PlotArea
        Geom_ChartCategoryCenterX = Geom_CategoryCenterX(.InsideLeft, .InsideWidth, Geom_ChartCategoryCount(cht), idx)
    End With
End Function

Public Function Geom_ChartValueToY(cht As Chart, v As Double) As Double
    Dim ax As Axis
    Set ax = cht.Axes(xlValue)
    With cht.PlotArea
        Geom_ChartValueToY = Geom_ValueToY(.InsideTop, .InsideHeight, ax.MinimumScale, ax.MaximumScale, v)
    End With
End Function

Public Function Geom_ChartBarHalfWidth(cht As Chart) As Double
    Dim gap As Double
    gap = 50
    On Error Resume Next
    gap = cht.ChartGroups(1).GapWidth
    On Error GoTo 0
    Geom_ChartBarHalfWidth = Geom_BarWidth(Geom_SlotWidth(cht.PlotArea.InsideWidth, Geom_ChartCategoryCount(cht)), gap) / 2
End Function
```

- [ ] **Passo 5: Implementar `modAnnotations.bas`**

```vb
Attribute VB_Name = "modAnnotations"
Option Explicit
' Anotações = formas dentro do gráfico (movem com ele). Parâmetros no AlternativeText.
Private mRefreshing As Boolean

Public Function Ann_TryCagr(startVal As Double, endVal As Double, periods As Double, ByRef rate As Double) As Boolean
    rate = 0
    If periods <= 0 Or startVal = 0 Then Exit Function
    If (startVal < 0) <> (endVal < 0) Or endVal = 0 Then Exit Function
    rate = (endVal / startVal) ^ (1 / periods) - 1
    Ann_TryCagr = True
End Function

Public Function Ann_ParseParam(altText As String, key As String) As String
    Dim parts() As String, i As Long, kv() As String
    parts = Split(altText, ";")
    For i = LBound(parts) To UBound(parts)
        kv = Split(parts(i), "=")
        If UBound(kv) = 1 Then
            If kv(0) = key Then Ann_ParseParam = kv(1): Exit Function
        End If
    Next
End Function

' values = número usado nos rótulos (total / End / 1ª série); tops = valor plotado onde a seta encosta.
Public Sub Ann_GetCategoryData(cht As Chart, ByRef values() As Double, ByRef tops() As Double)
    Dim kind As String, n As Long, i As Long, s As Series, vals As Variant, v As Double
    kind = Registry_ChartType(cht)
    n = Geom_ChartCategoryCount(cht)
    ReDim values(1 To n): ReDim tops(1 To n)
    Select Case kind
        Case "waterfall"
            values = Waterfall_EndValues(cht)
            tops = values
        Case "clustered", "line"
            vals = cht.SeriesCollection(1).Values
            For i = 1 To n
                v = Util_ToDouble(vals(i)): values(i) = v: tops(i) = v
            Next
        Case Else
            For Each s In cht.SeriesCollection
                If Not Style_IsHelperSeries(s) Then
                    vals = s.Values
                    For i = 1 To n
                        v = Util_ToDouble(vals(i))
                        values(i) = values(i) + v
                        If v > 0 Then tops(i) = tops(i) + v
                    Next
                End If
            Next
            If kind = "stacked100" Then
                For i = 1 To n: tops(i) = 1: Next
            End If
    End Select
End Sub

Public Function Ann_AddCagr(cht As Chart, fromIdx As Long, toIdx As Long) As Shape
    Dim nm As String, ln As Shape
    nm = Ann_NewName(cht, "cagr")
    Set ln = cht.Shapes.AddLine(0, 0, 10, 10)
    ln.Name = nm
    ln.AlternativeText = "kind=cagr;from=" & fromIdx & ";to=" & toIdx
    With ln.Line
        .ForeColor.RGB = RGB(89, 89, 89): .Weight = 1.25
        .EndArrowheadStyle = msoArrowheadTriangle
        .EndArrowheadLength = msoArrowheadShort: .EndArrowheadWidth = msoArrowheadNarrow
    End With
    Ann_AddLabel cht, nm
    Ann_Refresh cht
    Set Ann_AddCagr = ln
End Function

Public Function Ann_AddDifference(cht As Chart, fromIdx As Long, toIdx As Long) As Shape
    Dim nm As String, ln As Shape, lvl As Shape
    nm = Ann_NewName(cht, "diff")
    Set ln = cht.Shapes.AddLine(0, 0, 0, 10)
    ln.Name = nm
    ln.AlternativeText = "kind=diff;from=" & fromIdx & ";to=" & toIdx
    With ln.Line
        .ForeColor.RGB = RGB(89, 89, 89): .Weight = 1.25
        .BeginArrowheadStyle = msoArrowheadTriangle: .EndArrowheadStyle = msoArrowheadTriangle
        .BeginArrowheadLength = msoArrowheadShort: .EndArrowheadLength = msoArrowheadShort
    End With
    Set lvl = cht.Shapes.AddLine(0, 0, 10, 0)
    lvl.Name = nm & "_lvl"
    lvl.AlternativeText = "kind=level;owner=" & nm
    lvl.Line.ForeColor.RGB = RGB(166, 166, 166): lvl.Line.Weight = 0.75: lvl.Line.DashStyle = msoLineDash
    Ann_AddLabel cht, nm
    Ann_Refresh cht
    Set Ann_AddDifference = ln
End Function

Public Sub Ann_SetWaterfallConnectors(cht As Chart, turnOn As Boolean)
    Dim n As Long, i As Long, shp As Shape
    For i = cht.Shapes.Count To 1 Step -1
        If Left$(cht.Shapes(i).Name, Len(ANN_PREFIX & "conn_")) = ANN_PREFIX & "conn_" Then cht.Shapes(i).Delete
    Next
    Registry_SetOption cht, "connectors", IIf(turnOn, "1", "0")
    If Not turnOn Then Exit Sub
    n = Geom_ChartCategoryCount(cht)
    For i = 1 To n - 1
        Set shp = cht.Shapes.AddLine(0, 0, 10, 0)
        shp.Name = ANN_PREFIX & "conn_" & i
        shp.AlternativeText = "kind=conn;from=" & i & ";to=" & (i + 1)
        shp.Line.ForeColor.RGB = RGB(127, 127, 127): shp.Line.Weight = 0.75
    Next
    Ann_Refresh cht
End Sub

Public Sub Ann_RemoveAll(cht As Chart, Optional includeConnectors As Boolean = False)
    Dim i As Long, nm As String
    For i = cht.Shapes.Count To 1 Step -1
        nm = cht.Shapes(i).Name
        If Left$(nm, Len(ANN_PREFIX)) = ANN_PREFIX Then
            If includeConnectors Or Left$(nm, Len(ANN_PREFIX & "conn_")) <> ANN_PREFIX & "conn_" Then cht.Shapes(i).Delete
        End If
    Next
End Sub

' Reposiciona todas as anotações e reescreve os textos. Seguro para chamar em eventos.
Public Sub Ann_Refresh(cht As Chart)
    Dim values() As Double, tops() As Double, n As Long, i As Long, shp As Shape, kind As String
    Dim fmt As String, fromIdx As Long, toIdx As Long
    If mRefreshing Then Exit Sub
    mRefreshing = True
    On Error GoTo Done
    n = Geom_ChartCategoryCount(cht)
    If n = 0 Then GoTo Done
    Ann_GetCategoryData cht, values, tops
    fmt = Style_NumberFormat(Registry_GetOption(cht, "numfmt", "int"))
    For i = 1 To cht.Shapes.Count
        Set shp = cht.Shapes(i)
        If Left$(shp.Name, Len(ANN_PREFIX)) = ANN_PREFIX Then
            kind = Ann_ParseParam(shp.AlternativeText, "kind")
            fromIdx = Val(Ann_ParseParam(shp.AlternativeText, "from"))
            toIdx = Val(Ann_ParseParam(shp.AlternativeText, "to"))
            If fromIdx >= 1 And toIdx <= n And fromIdx < toIdx Then
                Select Case kind
                    Case "cagr": Ann_LayoutCagr cht, shp, fromIdx, toIdx, values, tops, fmt
                    Case "diff": Ann_LayoutDiff cht, shp, fromIdx, toIdx, values, tops, fmt
                    Case "conn": Ann_LayoutConn cht, shp, fromIdx, toIdx, tops
                End Select
            End If
        End If
    Next
Done:
    mRefreshing = False
End Sub

' ---------- privados ----------

Private Function Ann_NewName(cht As Chart, kind As String) As String
    Dim k As Long, nm As String
    Do
        k = k + 1
        nm = ANN_PREFIX & kind & "_" & k
    Loop While Not Ann_Shape(cht, nm) Is Nothing
    Ann_NewName = nm
End Function

Private Function Ann_Shape(cht As Chart, nm As String) As Shape
    On Error Resume Next
    Set Ann_Shape = cht.Shapes(nm)
    On Error GoTo 0
End Function

Private Sub Ann_AddLabel(cht As Chart, ownerName As String)
    Dim lbl As Shape
    Set lbl = cht.Shapes.AddTextbox(msoTextOrientationHorizontal, 0, 0, 70, 16)
    lbl.Name = ownerName & "_lbl"
    lbl.AlternativeText = "kind=label;owner=" & ownerName
    With lbl.TextFrame2
        .WordWrap = msoFalse
        .AutoSize = msoAutoSizeShapeToFitText
        .MarginLeft = 3: .MarginRight = 3: .MarginTop = 1: .MarginBottom = 1
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(64, 64, 64)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    lbl.Fill.Visible = msoTrue
    lbl.Fill.ForeColor.RGB = RGB(255, 255, 255)
    lbl.Line.Visible = msoTrue
    lbl.Line.ForeColor.RGB = RGB(166, 166, 166)
    lbl.Line.Weight = 0.5
End Sub

' Move uma linha para ir de (x1,y1) a (x2,y2), ajustando os flips para manter a ponta da seta no fim.
Private Sub Ann_MoveLine(ln As Shape, x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    ln.Left = IIf(x1 < x2, x1, x2)
    ln.Top = IIf(y1 < y2, y1, y2)
    ln.Width = Abs(x2 - x1)
    ln.Height = Abs(y2 - y1)
    If (ln.HorizontalFlip = msoTrue) <> (x2 < x1) Then ln.Flip msoFlipHorizontal
    If (ln.VerticalFlip = msoTrue) <> (y2 < y1) Then ln.Flip msoFlipVertical
End Sub

Private Function Ann_SignedFormat(fmt As String) As String
    Dim sec As String
    sec = Split(fmt & ";", ";")(0)
    Ann_SignedFormat = "+" & sec & ";-" & sec & ";0"
End Function

Private Sub Ann_LayoutCagr(cht As Chart, ln As Shape, fromIdx As Long, toIdx As Long, values() As Double, tops() As Double, fmt As String)
    Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double, lbl As Shape, rate As Double, txt As String
    Dim body As Range, periods As Double
    x1 = Geom_ChartCategoryCenterX(cht, fromIdx): x2 = Geom_ChartCategoryCenterX(cht, toIdx)
    y1 = Geom_ChartValueToY(cht, tops(fromIdx)) - 10
    y2 = Geom_ChartValueToY(cht, tops(toIdx)) - 10
    Ann_MoveLine ln, x1, y1, x2, y2
    Set body = Block_Body(Registry_Block(cht))
    periods = Geom_Periods(body.Cells(fromIdx + 1, 1).Value, body.Cells(toIdx + 1, 1).Value, fromIdx, toIdx)
    If Ann_TryCagr(values(fromIdx), values(toIdx), periods, rate) Then
        txt = "CAGR " & Format$(rate, "+0.0%;-0.0%;0.0%")
    Else
        txt = "CAGR n/d"
    End If
    Set lbl = Ann_Shape(cht, ln.Name & "_lbl")
    If lbl Is Nothing Then Exit Sub
    lbl.TextFrame2.TextRange.Text = txt
    lbl.Left = (x1 + x2) / 2 - lbl.Width / 2
    lbl.Top = IIf(y1 < y2, y1, y2) - lbl.Height - 4
End Sub

Private Sub Ann_LayoutDiff(cht As Chart, ln As Shape, fromIdx As Long, toIdx As Long, values() As Double, tops() As Double, fmt As String)
    Dim xFrom As Double, xTo As Double, yFrom As Double, yTo As Double, xArrow As Double, half As Double
    Dim lvl As Shape, lbl As Shape, diff As Double, txt As String
    half = Geom_ChartBarHalfWidth(cht)
    xFrom = Geom_ChartCategoryCenterX(cht, fromIdx): xTo = Geom_ChartCategoryCenterX(cht, toIdx)
    yFrom = Geom_ChartValueToY(cht, tops(fromIdx)): yTo = Geom_ChartValueToY(cht, tops(toIdx))
    xArrow = xTo + half + 10
    Set lvl = Ann_Shape(cht, ln.Name & "_lvl")
    If Not lvl Is Nothing Then Ann_MoveLine lvl, xFrom + half, yFrom, xArrow + 4, yFrom
    Ann_MoveLine ln, xArrow, yFrom, xArrow, yTo
    diff = values(toIdx) - values(fromIdx)
    txt = Format$(diff, Ann_SignedFormat(fmt))
    If values(fromIdx) <> 0 Then txt = txt & " (" & Format$(diff / Abs(values(fromIdx)), "+0%;-0%;0%") & ")"
    Set lbl = Ann_Shape(cht, ln.Name & "_lbl")
    If lbl Is Nothing Then Exit Sub
    lbl.TextFrame2.TextRange.Text = txt
    lbl.Left = xArrow + 6
    lbl.Top = (yFrom + yTo) / 2 - lbl.Height / 2
End Sub

Private Sub Ann_LayoutConn(cht As Chart, shp As Shape, i As Long, j As Long, tops() As Double)
    Dim half As Double, y As Double
    half = Geom_ChartBarHalfWidth(cht)
    y = Geom_ChartValueToY(cht, tops(i))
    Ann_MoveLine shp, Geom_ChartCategoryCenterX(cht, i) + half, y, Geom_ChartCategoryCenterX(cht, j) - half, y
End Sub
```

- [ ] **Passo 6: Ligar os conectores na inserção da cascata**

Em `src/core/modules/modChartWaterfall.bas`, dentro de `Waterfall_Insert`, substitua a linha `Registry_SetOption cht, "connectors", "0"` por:

```vb
    Ann_SetWaterfallConnectors cht, True
```

e mova a chamada `Waterfall_SetLabels cht, True` para **antes** dela (os rótulos precisam existir antes do primeiro `Ann_Refresh`). O fim de `Waterfall_Insert` fica:

```vb
    Registry_SetOption cht, "totals", "0"
    Waterfall_SetLabels cht, True
    Ann_SetWaterfallConnectors cht, True
    Set Waterfall_Insert = co
End Function
```

- [ ] **Passo 7: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: todos `ok`. Os textos esperados usam vírgula decimal (`+17,0%`) porque o Excel desta máquina é pt-BR; em máquina en-US ajuste os testes para ponto.

- [ ] **Passo 8: Commit**

```bash
git add src/core/modules/modGeometry.bas src/core/modules/modAnnotations.bas src/core/modules/modChartWaterfall.bas tests/vba/core/TestGeometry.bas tests/vba/core/TestAnnotations.bas
git commit -m "feat(core): setas de CAGR e diferença e conectores da cascata com reposicionamento"
```

---
### Tarefa 11: Eventos de gráfico, API pública do core e formulário de anotação (core)

**Files:**
- Create: `src/core/classes/clsChartWatcher.cls`, `src/core/classes/clsAppEvents.cls`, `src/core/modules/modWatchers.bas`, `src/core/modules/modErrors.bas`, `src/core/forms/frmAnnotation.json`, `src/core/forms/frmAnnotation.vba`
- Modify: `src/core/modules/modCoreApi.bas` (substituir inteiro), `src/core/ThisWorkbook.vba` (substituir inteiro)
- Test: `tests/vba/core/TestCoreApi.bas`

**Interfaces:**
- Produces (modCoreApi, chamadas pelo loader via `Application.Run`): `Core_Ping`, `Core_Version`, `Core_OnLoaded()`, `Core_OnUnload()`, `Core_RegisterLoader(loaderName, loaderFullPath)`, `Core_ActiveChart() As Chart` (só gráficos ThinQQ), `Core_HasChart() As Boolean`, `Core_Insert(kind)`, `Core_GetToggle(key) As Boolean`, `Core_SetToggle(key, value)`, `Core_GetOption(key) As String`, `Core_SetOption(key, value)`, `Core_AddAnnotation(kind, Optional fromIdx = 0, Optional toIdx = 0)` (sem índices abre `frmAnnotation`), `Core_RemoveAnnotations()`, `Core_Refresh()`, `Core_ChangeSource()`, `Core_SafeRefresh(cht)`, `Core_LoaderPath() As String`.
- Produces (modWatchers): `Watchers_Start()`, `Watchers_Stop()`, `Watchers_SetLoader(name)`, `Watchers_Rehook()`, `Watchers_Count() As Long`, `Watchers_NotifySelection()` (invalida a faixa via `Loader_InvalidateRibbon` quando o estado muda).
- Produces (modErrors): `Errors_Log(msg)`, `Errors_Handle(context)` (loga e mostra MsgBox; usar em `Fail:`).
- `frmAnnotation`: `Prepare(categories As Variant, defaultFrom, defaultTo)`, `Result` (`"ok"|"cancel"`), `FromIdx`, `ToIdx`.
- Chaves de toggle: `totals`, `labels`, `connectors`. Chaves de opção: `numfmt`, `palette`, `gap`.

- [ ] **Passo 1: Escrever os testes**

`tests/vba/core/TestCoreApi.bas`:

```vb
Attribute VB_Name = "TestCoreApi"
Option Explicit

Private Function Fixture(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1): ws.Name = "Dados"
    ws.Range("A1:C1").Value = Array("Ano", "Sul", "Norte")
    ws.Range("A2:C2").Value = Array(2022, 100, 50)
    ws.Range("A3:C3").Value = Array(2023, 110, 60)
    ws.Range("A4:C4").Value = Array(2024, 120, 70)
    ws.Range("A5:C5").Value = Array(2025, 150, 90)
    Set Fixture = ws
End Function

Public Sub Test_Core_HasChart_FalseOnCells()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1").Select
    AssertFalse Core_HasChart()
    AssertTrue Core_ActiveChart() Is Nothing
    CloseTemp wb
End Sub

Public Sub Test_Core_Insert_SelectsChart_And_Hooks()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_OnLoaded
    Core_Insert "stacked"
    AssertTrue Core_HasChart(), "o gráfico inserido deve ficar selecionado"
    AssertEqual "stacked", Registry_ChartType(Core_ActiveChart())
    AssertEqual 1, Watchers_Count()
    AssertTrue Core_GetToggle("totals")
    AssertTrue Core_GetToggle("labels")
    AssertFalse Core_GetToggle("connectors")
    AssertEqual "int", Core_GetOption("numfmt")
    Core_OnUnload
    CloseTemp wb
End Sub

Public Sub Test_Core_Toggles_And_Options()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    Set cht = Core_ActiveChart()
    Core_SetToggle "totals", False
    AssertTrue Chart_FindSeries(cht, "tqq_total") Is Nothing
    AssertFalse Core_GetToggle("totals")
    Core_SetToggle "connectors", True
    AssertTrue cht.ChartGroups(1).HasSeriesLines
    Core_SetOption "numfmt", "dec1"
    AssertEqual "#,##0.0;-#,##0.0;", cht.SeriesCollection(1).DataLabels.NumberFormat
    Core_SetOption "gap", "120"
    AssertEqual 120, cht.ChartGroups(1).GapWidth
    Core_SetOption "palette", "cinza"
    AssertEqual RGB(64, 64, 64), cht.SeriesCollection(1).Format.Fill.ForeColor.RGB
    CloseTemp wb
End Sub

Public Sub Test_Core_Insert_Waterfall_And_Annotations()
    Dim wb As Workbook, ws As Worksheet, cht As Chart
    Set wb = NewTempWorkbook(): Set ws = wb.Worksheets(1)
    ws.Range("A1:B1").Value = Array("Etapa", "Valor")
    ws.Range("A2:B2").Value = Array("Início", 100)
    ws.Range("A3:B3").Value = Array("A", -30)
    ws.Range("A4:B4").Value = Array("Fim", "e")
    ws.Range("A1").Select
    Core_Insert "waterfall"
    Set cht = Core_ActiveChart()
    AssertEqual "waterfall", Registry_ChartType(cht)
    AssertTrue Core_GetToggle("connectors")
    Core_AddAnnotation "cagr", 1, 3
    Core_AddAnnotation "diff", 1, 3
    AssertTrue Not cht.Shapes("tqq_ann_cagr_1") Is Nothing
    AssertTrue Not cht.Shapes("tqq_ann_diff_1") Is Nothing
    Core_RemoveAnnotations
    AssertEqual 2, cht.Shapes.Count, "só os 2 conectores devem restar"
    Core_SetToggle "connectors", False
    AssertEqual 0, cht.Shapes.Count
    CloseTemp wb
End Sub

Public Sub Test_Core_Insert_InvalidSelection_NoChart()
    Dim wb As Workbook, ws As Worksheet, before As Long
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    before = ws.ChartObjects.Count
    ws.Range("A1:A5").Select
    Application.DisplayAlerts = False
    Core_Insert "stacked"       ' mostra MsgBox apenas quando interativo; em testes Core_Insert não deve travar
    AssertEqual before, ws.ChartObjects.Count
    CloseTemp wb
End Sub

Public Sub Test_Core_Refresh_CleansOrphans()
    Dim wb As Workbook, ws As Worksheet
    Set wb = NewTempWorkbook(): Set ws = Fixture(wb)
    ws.Range("A1:C5").Select
    Core_Insert "stacked"
    ws.ChartObjects(1).Delete
    AssertTrue Not CalcSheet_FindBlock(wb, Mid$(wb.Names(1).Name, 5)) Is Nothing
    Core_Refresh
    AssertEqual 0, wb.Names.Count
    CloseTemp wb
End Sub
```

Observação: `Core_Insert` com seleção inválida exibe uma `MsgBox`, que travaria o Excel invisível. Para que o teste passe sem interação, `modCoreApi` usa `Core_Notify(msg)`, que **não** mostra caixa de mensagem quando `BUILD_CONFIG = "Test"` (apenas registra no log).

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: `not ok` em `TestCoreApi.*`.

- [ ] **Passo 3: Implementar `modErrors.bas`**

```vb
Attribute VB_Name = "modErrors"
Option Explicit

Public Sub Errors_Log(msg As String)
    On Error Resume Next
    Dim f As Integer, dir_ As String
    dir_ = Environ$("LOCALAPPDATA") & "\ThinQQ\logs"
    Util_EnsureDir dir_
    f = FreeFile
    Open dir_ & "\thinqq.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & "core" & vbTab & msg
    Close #f
End Sub

' Usar em rótulos Fail: de funções públicas. Loga e avisa o usuário (sem caixa em build de teste).
Public Sub Errors_Handle(context As String)
    Dim num As Long, desc As String
    num = Err.Number: desc = Err.Description
    Errors_Log context & vbTab & num & vbTab & desc
    Core_Notify "Ocorreu um erro no ThinQQ (" & context & "):" & vbLf & desc & vbLf & vbLf & _
                "Detalhes em %LOCALAPPDATA%\ThinQQ\logs\thinqq.log", vbExclamation
End Sub
```

- [ ] **Passo 4: Implementar as classes de eventos e `modWatchers.bas`**

`src/core/classes/clsChartWatcher.cls`:

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsChartWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Um observador por gráfico ThinQQ da planilha ativa.
Public WithEvents Cht As Chart

Private Sub Cht_Activate()
    Watchers_NotifySelection
End Sub

Private Sub Cht_Deactivate()
    Watchers_NotifySelection
End Sub

Private Sub Cht_Calculate()
    Core_SafeRefresh Cht
End Sub

Private Sub Cht_Resize()
    Core_SafeRefresh Cht
End Sub
```

`src/core/classes/clsAppEvents.cls`:

```vb
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents App As Application

Private Sub App_SheetActivate(ByVal Sh As Object)
    Watchers_Rehook
End Sub

Private Sub App_WorkbookActivate(ByVal Wb As Workbook)
    Watchers_Rehook
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    Watchers_Rehook
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Watchers_NotifySelection
End Sub
```

`src/core/modules/modWatchers.bas`:

```vb
Attribute VB_Name = "modWatchers"
Option Explicit
Private mWatchers As Collection
Private mAppEvents As clsAppEvents
Private mLastHasChart As Boolean
Private mLoaderName As String

Public Sub Watchers_Start()
    Set mAppEvents = New clsAppEvents
    Set mAppEvents.App = Application
    Watchers_Rehook
End Sub

Public Sub Watchers_Stop()
    Set mAppEvents = Nothing
    Set mWatchers = Nothing
End Sub

Public Sub Watchers_SetLoader(name As String)
    mLoaderName = name
End Sub

' Reassocia observadores a todos os gráficos ThinQQ da planilha ativa.
Public Sub Watchers_Rehook()
    Dim ws As Worksheet, co As ChartObject, w As clsChartWatcher
    Set mWatchers = New Collection
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    For Each co In ws.ChartObjects
        If Left$(co.Name, Len(ID_PREFIX)) = ID_PREFIX Then
            Set w = New clsChartWatcher
            Set w.Cht = co.Chart
            mWatchers.Add w
        End If
    Next
End Sub

Public Function Watchers_Count() As Long
    If mWatchers Is Nothing Then Exit Function
    Watchers_Count = mWatchers.Count
End Function

' Invalida a faixa quando entra/sai de um gráfico ThinQQ (ou troca de gráfico).
Public Sub Watchers_NotifySelection()
    Dim hasChart As Boolean
    hasChart = Core_HasChart()
    If hasChart Or hasChart <> mLastHasChart Then
        mLastHasChart = hasChart
        If Len(mLoaderName) > 0 Then
            On Error Resume Next
            Application.Run "'" & mLoaderName & "'!Loader_InvalidateRibbon"
            On Error GoTo 0
        End If
    End If
End Sub
```

- [ ] **Passo 5: Criar `frmAnnotation`**

`src/core/forms/frmAnnotation.json`:

```json
{
  "name": "frmAnnotation",
  "caption": "Adicionar seta",
  "width": 320,
  "height": 170,
  "controls": [
    { "type": "Label", "name": "lblFrom", "Left": 16, "Top": 16, "Width": 120, "Height": 14, "Caption": "Da categoria:", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "ComboBox", "name": "cboFrom", "Left": 140, "Top": 14, "Width": 160, "Height": 20, "Style": 2, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "Label", "name": "lblTo", "Left": 16, "Top": 46, "Width": 120, "Height": 14, "Caption": "Até a categoria:", "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "ComboBox", "name": "cboTo", "Left": 140, "Top": 44, "Width": 160, "Height": 20, "Style": 2, "fontName": "Segoe UI", "fontSize": 9 },
    { "type": "CommandButton", "name": "btnOk", "Left": 140, "Top": 96, "Width": 76, "Height": 26, "Caption": "OK", "Default": true, "fontName": "Segoe UI", "fontSize": 9, "fontBold": true },
    { "type": "CommandButton", "name": "btnCancel", "Left": 224, "Top": 96, "Width": 76, "Height": 26, "Caption": "Cancelar", "Cancel": true, "fontName": "Segoe UI", "fontSize": 9 }
  ]
}
```

`src/core/forms/frmAnnotation.vba`:

```vb
Option Explicit
Public Result As String
Public FromIdx As Long
Public ToIdx As Long

Public Sub Prepare(categories As Variant, defaultFrom As Long, defaultTo As Long)
    Dim i As Long
    cboFrom.Clear: cboTo.Clear
    For i = LBound(categories) To UBound(categories)
        cboFrom.AddItem CStr(categories(i))
        cboTo.AddItem CStr(categories(i))
    Next
    cboFrom.ListIndex = defaultFrom - 1
    cboTo.ListIndex = defaultTo - 1
    Result = "cancel"
End Sub

Private Sub btnOk_Click()
    If cboFrom.ListIndex < 0 Or cboTo.ListIndex < 0 Then Exit Sub
    If cboFrom.ListIndex >= cboTo.ListIndex Then
        MsgBox "A categoria final deve vir depois da inicial.", vbExclamation, "ThinQQ"
        Exit Sub
    End If
    FromIdx = cboFrom.ListIndex + 1
    ToIdx = cboTo.ListIndex + 1
    Result = "ok"
    Me.Hide
End Sub

Private Sub btnCancel_Click()
    Result = "cancel"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then Cancel = True: Result = "cancel": Me.Hide
End Sub
```

- [ ] **Passo 6: Substituir `modCoreApi.bas` pela API completa**

```vb
Attribute VB_Name = "modCoreApi"
Option Explicit
' API pública chamada pelo loader via Application.Run. Nunca propaga erro.
Private mLoaderPath As String

Public Function Core_Ping() As String
    Core_Ping = "pong"
End Function

Public Function Core_Version() As String
    Core_Version = APP_VERSION
End Function

Public Function Core_LoaderPath() As String
    Core_LoaderPath = mLoaderPath
End Function

Public Sub Core_OnLoaded()
    On Error Resume Next
    Watchers_Start
End Sub

Public Sub Core_OnUnload()
    On Error Resume Next
    Watchers_Stop
End Sub

Public Sub Core_RegisterLoader(loaderName As String, loaderFullPath As String)
    Watchers_SetLoader loaderName
    mLoaderPath = loaderFullPath
End Sub

' Mensagem ao usuário (suprimida em build de teste para não travar o Excel invisível).
Public Sub Core_Notify(msg As String, Optional style As VbMsgBoxStyle = vbInformation)
    Errors_Log "notify" & vbTab & Replace(msg, vbLf, " | ")
    If BUILD_CONFIG = "Test" Then Exit Sub
    MsgBox msg, style, APP_NAME
End Sub

Public Function Core_ActiveChart() As Chart
    Dim c As Chart
    On Error Resume Next
    Set c = Application.ActiveChart
    On Error GoTo 0
    If c Is Nothing Then Exit Function
    If Registry_IsThinQQChart(c) Then Set Core_ActiveChart = c
End Function

Public Function Core_HasChart() As Boolean
    Core_HasChart = Not Core_ActiveChart() Is Nothing
End Function

Public Sub Core_Insert(kind As String)
    On Error GoTo Fail
    Dim data As clsChartData, msg As String, co As ChartObject, ws As Worksheet
    If TypeName(Selection) <> "Range" Then
        Core_Notify "Selecione a tabela de dados (cabeçalho + categorias) antes de inserir o gráfico."
        Exit Sub
    End If
    Set data = ChartData_FromRange(Selection)
    msg = data.Validate()
    If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Sub
    Set ws = data.Source.Worksheet
    Application.ScreenUpdating = False
    CalcSheet_Cleanup ws.Parent
    If kind = "waterfall" Then
        Set co = Waterfall_Insert(ws, data)
    Else
        Set co = Stacked_Insert(ws, data, kind)
    End If
    Application.ScreenUpdating = True
    Watchers_Rehook
    co.Select
    Core_SafeRefresh co.Chart
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Errors_Handle "Core_Insert(" & kind & ")"
End Sub

Public Function Core_GetToggle(key As String) As Boolean
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Function
    Core_GetToggle = (Registry_GetOption(c, key, "0") = "1")
End Function

Public Sub Core_SetToggle(key As String, value As Boolean)
    On Error GoTo Fail
    Dim c As Chart, kind As String
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    kind = Registry_ChartType(c)
    Select Case key
        Case "totals"
            If kind <> "waterfall" Then Stacked_SetTotals c, value
        Case "labels"
            If kind = "waterfall" Then Waterfall_SetLabels c, value Else Stacked_SetLabels c, value
        Case "connectors"
            If kind = "waterfall" Then Ann_SetWaterfallConnectors c, value Else Stacked_SetConnectors c, value
    End Select
    Core_SafeRefresh c
    Exit Sub
Fail:
    Errors_Handle "Core_SetToggle(" & key & ")"
End Sub

Public Function Core_GetOption(key As String) As String
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Function
    Core_GetOption = Registry_GetOption(c, key, "")
End Function

Public Sub Core_SetOption(key As String, value As String)
    On Error GoTo Fail
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    Registry_SetOption c, key, value
    Select Case key
        Case "numfmt"
            Style_ApplyNumberFormat c, Style_NumberFormat(value)
        Case "palette"
            If Registry_ChartType(c) = "waterfall" Then Waterfall_ApplyColors c, value Else Style_ApplySeriesColors c, value
        Case "gap"
            On Error Resume Next
            c.ChartGroups(1).GapWidth = CLng(value)
            On Error GoTo Fail
    End Select
    Core_SafeRefresh c
    Exit Sub
Fail:
    Errors_Handle "Core_SetOption(" & key & ")"
End Sub

Public Sub Core_AddAnnotation(kind As String, Optional fromIdx As Long = 0, Optional toIdx As Long = 0)
    On Error GoTo Fail
    Dim c As Chart, n As Long
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    n = Geom_ChartCategoryCount(c)
    If n < 2 Then Core_Notify "O gráfico precisa de pelo menos 2 categorias.": Exit Sub
    If fromIdx < 1 Or toIdx > n Or fromIdx >= toIdx Then
        If Not Core_AskCategories(c, n, fromIdx, toIdx) Then Exit Sub
    End If
    If kind = "cagr" Then Ann_AddCagr c, fromIdx, toIdx Else Ann_AddDifference c, fromIdx, toIdx
    Exit Sub
Fail:
    Errors_Handle "Core_AddAnnotation(" & kind & ")"
End Sub

Private Function Core_AskCategories(c As Chart, n As Long, ByRef fromIdx As Long, ByRef toIdx As Long) As Boolean
    Dim frm As frmAnnotation, cats() As String, body As Range, i As Long
    Set body = Block_Body(Registry_Block(c))
    ReDim cats(1 To n)
    For i = 1 To n
        cats(i) = CStr(body.Cells(i + 1, 1).Text)
    Next
    Set frm = New frmAnnotation
    frm.Prepare cats, 1, n
    frm.Show
    If frm.Result = "ok" Then
        fromIdx = frm.FromIdx: toIdx = frm.ToIdx
        Core_AskCategories = True
    End If
    Unload frm
End Function

Public Sub Core_RemoveAnnotations()
    On Error GoTo Fail
    Dim c As Chart
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    Ann_RemoveAll c, False
    Exit Sub
Fail:
    Errors_Handle "Core_RemoveAnnotations"
End Sub

Public Sub Core_SafeRefresh(cht As Chart)
    On Error Resume Next
    Labels_RefreshPositions cht
    Ann_Refresh cht
End Sub

' Re-layout do gráfico ativo (ou de todos da planilha) e limpeza de blocos órfãos.
Public Sub Core_Refresh()
    On Error GoTo Fail
    Dim c As Chart, co As ChartObject, ws As Worksheet
    Set c = Core_ActiveChart()
    If Not c Is Nothing Then
        Core_SafeRefresh c
        CalcSheet_Cleanup Registry_Workbook(c)
    Else
        Set ws = ActiveSheet
        For Each co In ws.ChartObjects
            If Left$(co.Name, Len(ID_PREFIX)) = ID_PREFIX Then Core_SafeRefresh co.Chart
        Next
        CalcSheet_Cleanup ws.Parent
    End If
    Watchers_Rehook
    Exit Sub
Fail:
    Errors_Handle "Core_Refresh"
End Sub

' Troca a fonte de dados: recria o gráfico no mesmo lugar, com as mesmas opções.
Public Sub Core_ChangeSource()
    On Error GoTo Fail
    Dim c As Chart, co As ChartObject, rng As Range, data As clsChartData, msg As String
    Dim kind As String, l As Double, t As Double, w As Double, h As Double, opts As String, ws As Worksheet, wb As Workbook
    Set c = Core_ActiveChart()
    If c Is Nothing Then Exit Sub
    On Error Resume Next
    Set rng = Application.InputBox("Selecione a nova tabela (cabeçalho + categorias):", APP_NAME, Type:=8)
    On Error GoTo Fail
    If rng Is Nothing Then Exit Sub
    Set data = ChartData_FromRange(rng)
    msg = data.Validate()
    If Len(msg) > 0 Then Core_Notify msg, vbExclamation: Exit Sub
    Set co = c.Parent
    kind = Registry_ChartType(c)
    l = co.Left: t = co.Top: w = co.Width: h = co.Height
    opts = CStr(Registry_Block(c).Cells(1, 4).Value)
    Set ws = co.Parent: Set wb = ws.Parent
    co.Delete
    CalcSheet_Cleanup wb
    If kind = "waterfall" Then Set co = Waterfall_Insert(ws, data) Else Set co = Stacked_Insert(ws, data, kind)
    co.Left = l: co.Top = t: co.Width = w: co.Height = h
    Registry_Block(co.Chart).Cells(1, 4).Value = opts
    co.Select
    Core_SetOption "numfmt", Registry_GetOption(co.Chart, "numfmt", "int")
    Core_SetOption "palette", Registry_GetOption(co.Chart, "palette", "padrao")
    Core_SetToggle "totals", Registry_GetOption(co.Chart, "totals", "0") = "1"
    Core_SetToggle "labels", Registry_GetOption(co.Chart, "labels", "1") = "1"
    Core_SetToggle "connectors", Registry_GetOption(co.Chart, "connectors", "0") = "1"
    Watchers_Rehook
    Exit Sub
Fail:
    Errors_Handle "Core_ChangeSource"
End Sub
```

`src/core/ThisWorkbook.vba`:

```vb
Option Explicit

Private Sub Workbook_Open()
    Core_OnLoaded
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Core_OnUnload
End Sub
```

- [ ] **Passo 7: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: todos `ok`. Se `Test_Core_Insert_SelectsChart_And_Hooks` falhar em `Core_HasChart` por `ActiveChart` vazio no Excel invisível, troque `co.Select` por `co.Activate` em `Core_Insert`.

- [ ] **Passo 8: Commit**

```bash
git add src/core tests/vba/core/TestCoreApi.bas
git commit -m "feat(core): API pública, observadores de eventos e formulário de anotação"
```

---

### Tarefa 12: Faixa de opções completa (loader)

**Files:**
- Modify: `src/loader/ribbon/customUI14.xml` (substituir inteiro), `src/loader/modules/modRibbon.bas` (substituir inteiro)
- Test: `tests/vba/loader/TestRibbon.bas`

**Interfaces:**
- Produces: callbacks `Rx_OnLoad`, `Rx_OnInsert` (tag = kind), `Rx_GetEnabled`, `Rx_OnToggle`/`Rx_GetPressed` (tag = chave), `Rx_OnDropdown`/`Rx_GetSelectedId` (tag = chave; id do item = valor), `Rx_OnAnnotation` (tag = `cagr|diff`), `Rx_OnRemoveAnnotations`, `Rx_OnRefresh`, `Rx_OnChangeSource`, `Rx_OnSamples`, `Rx_OnTutorial`, `Rx_OnCheckUpdates`, `Rx_OnSettings`, `Rx_OnAbout`; `Loader_InvalidateRibbon()`; `Loader_TutorialPath() As String`.
- Consumes: `Loader_RunCore` (Tarefa 4), `Core_*` (Tarefa 11), `Updater_CheckInteractive` (stub da Tarefa 4, real na Tarefa 13), `Core_InsertSamples` (Tarefa 15; até lá o botão avisa).

- [ ] **Passo 1: Escrever o teste**

`tests/vba/loader/TestRibbon.bas`:

```vb
Attribute VB_Name = "TestRibbon"
Option Explicit

Public Sub Test_Ribbon_Xml_HasAllCallbacks()
    Dim xml As String, names As Variant, i As Long
    xml = Transport_ReadTextFile(RepoPath() & "\src\loader\ribbon\customUI14.xml")
    names = Array("Rx_OnLoad", "Rx_OnInsert", "Rx_GetEnabled", "Rx_OnToggle", "Rx_GetPressed", "Rx_OnDropdown", _
                  "Rx_GetSelectedId", "Rx_OnAnnotation", "Rx_OnRemoveAnnotations", "Rx_OnRefresh", "Rx_OnChangeSource", _
                  "Rx_OnSamples", "Rx_OnTutorial", "Rx_OnCheckUpdates", "Rx_OnSettings", "Rx_OnAbout")
    For i = LBound(names) To UBound(names)
        AssertTrue InStr(xml, names(i)) > 0, "callback ausente no XML: " & names(i)
    Next
    AssertTrue InStr(xml, "tag=""waterfall""") > 0
    AssertTrue InStr(xml, "tag=""numfmt""") > 0
End Sub

Public Sub Test_Ribbon_CoreState_WithoutChart()
    AssertEqual False, CBool(Loader_RunCore("Core_HasChart"))
End Sub

Public Sub Test_Ribbon_TutorialPath_FallsBackToUrl()
    Dim p As String
    Loader_SetRootOverride TempDir()
    p = Loader_TutorialPath()
    Loader_SetRootOverride ""
    AssertTrue Transport_IsHttp(p), "sem arquivo local deve usar a URL"
End Sub
```

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target loader`
Expected: `not ok` em `TestRibbon.*`.

- [ ] **Passo 3: Escrever o `customUI14.xml` completo**

```xml
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Rx_OnLoad">
  <ribbon>
    <tabs>
      <tab id="tqqTab" label="ThinQQ" insertAfterMso="TabInsert">

        <group id="grpInsert" label="Inserir gráfico">
          <button id="btnStacked" label="Colunas empilhadas" size="large" imageMso="ChartTypeColumnInsertGallery"
                  tag="stacked" onAction="Rx_OnInsert"
                  screentip="Colunas empilhadas" supertip="Selecione a tabela (linha de cabeçalho + coluna de categorias) e clique. Totais e rótulos entram automaticamente."/>
          <button id="btnClustered" label="Colunas agrupadas" size="large" imageMso="ChartTypeColumnInsertGallery"
                  tag="clustered" onAction="Rx_OnInsert" screentip="Colunas agrupadas"/>
          <button id="btnStacked100" label="Colunas 100%" size="large" imageMso="ChartTypeColumnInsertGallery"
                  tag="stacked100" onAction="Rx_OnInsert" screentip="Colunas 100% empilhadas"/>
          <button id="btnWaterfall" label="Cascata" size="large" imageMso="ChartTypeOtherInsertGallery"
                  tag="waterfall" onAction="Rx_OnInsert"
                  screentip="Cascata (ponte)" supertip="Tabela com 2 colunas: etapa e valor. Escreva 'e' no valor para uma barra de total/subtotal calculada."/>
          <button id="btnLine" label="Linha" size="large" imageMso="ChartTypeLineInsertGallery"
                  tag="line" onAction="Rx_OnInsert" screentip="Linha com rótulos"/>
        </group>

        <group id="grpChart" label="Gráfico selecionado">
          <toggleButton id="tglTotals" label="Totais" imageMso="AutoSum" tag="totals"
                        onAction="Rx_OnToggle" getPressed="Rx_GetPressed" getEnabled="Rx_GetEnabled"
                        screentip="Totais acima das colunas"/>
          <toggleButton id="tglLabels" label="Rótulos" imageMso="TextBoxInsert" tag="labels"
                        onAction="Rx_OnToggle" getPressed="Rx_GetPressed" getEnabled="Rx_GetEnabled"
                        screentip="Rótulos de valor"/>
          <toggleButton id="tglConnectors" label="Conectores" imageMso="ShapeStraightConnector" tag="connectors"
                        onAction="Rx_OnToggle" getPressed="Rx_GetPressed" getEnabled="Rx_GetEnabled"
                        screentip="Linhas conectando as colunas"/>
          <separator id="sep1"/>
          <dropDown id="ddFormat" label="Formato" tag="numfmt" onAction="Rx_OnDropdown"
                    getSelectedItemID="Rx_GetSelectedId" getEnabled="Rx_GetEnabled" screentip="Formato dos números">
            <item id="int"  label="1.234"/>
            <item id="dec1" label="1.234,5"/>
            <item id="dec2" label="1.234,56"/>
            <item id="pct0" label="12%"/>
            <item id="pct1" label="12,3%"/>
            <item id="mil"  label="Milhares (1.234 → 1)"/>
            <item id="mi"   label="Milhões (1,2)"/>
          </dropDown>
          <dropDown id="ddPalette" label="Paleta" tag="palette" onAction="Rx_OnDropdown"
                    getSelectedItemID="Rx_GetSelectedId" getEnabled="Rx_GetEnabled" screentip="Cores das séries">
            <item id="padrao" label="Padrão ThinQQ"/>
            <item id="cinza"  label="Tons de cinza"/>
            <item id="tema"   label="Cores do tema"/>
          </dropDown>
          <dropDown id="ddGap" label="Espaço" tag="gap" onAction="Rx_OnDropdown"
                    getSelectedItemID="Rx_GetSelectedId" getEnabled="Rx_GetEnabled" screentip="Espaço entre colunas">
            <item id="30"  label="Estreito (30%)"/>
            <item id="50"  label="Normal (50%)"/>
            <item id="80"  label="Largo (80%)"/>
            <item id="120" label="Muito largo (120%)"/>
          </dropDown>
          <separator id="sep2"/>
          <button id="btnCagr" label="Seta CAGR" imageMso="ArrowsMore" tag="cagr" onAction="Rx_OnAnnotation"
                  getEnabled="Rx_GetEnabled" screentip="Seta de crescimento anual composto entre duas categorias"/>
          <button id="btnDiff" label="Seta diferença" imageMso="ArrowsMore" tag="diff" onAction="Rx_OnAnnotation"
                  getEnabled="Rx_GetEnabled" screentip="Seta com a diferença absoluta e relativa entre dois totais"/>
          <button id="btnRemoveAnn" label="Remover setas" imageMso="Delete" onAction="Rx_OnRemoveAnnotations"
                  getEnabled="Rx_GetEnabled"/>
          <separator id="sep3"/>
          <button id="btnRefresh" label="Atualizar" imageMso="Refresh" onAction="Rx_OnRefresh"
                  screentip="Reposiciona setas e rótulos e limpa dados internos sem uso"/>
          <button id="btnChangeSource" label="Alterar dados" imageMso="TableInsert" onAction="Rx_OnChangeSource"
                  getEnabled="Rx_GetEnabled" screentip="Trocar a tabela de origem do gráfico"/>
        </group>

        <group id="grpAbout" label="ThinQQ">
          <button id="btnSamples" label="Exemplos" size="large" imageMso="FileNew" onAction="Rx_OnSamples"
                  screentip="Abre uma pasta com tabelas e gráficos de exemplo"/>
          <button id="btnTutorial" label="Tutorial" size="large" imageMso="Help" onAction="Rx_OnTutorial"/>
          <button id="btnCheckUpdates" label="Verificar atualizações" imageMso="Refresh" onAction="Rx_OnCheckUpdates"/>
          <button id="btnSettings" label="Configurações" imageMso="PropertySheet" onAction="Rx_OnSettings"
                  screentip="Fonte de atualização (URL ou pasta)"/>
          <button id="btnAbout" label="Sobre" imageMso="Info" onAction="Rx_OnAbout"/>
        </group>

      </tab>
    </tabs>
  </ribbon>
</customUI>
```

Se algum ícone aparecer em branco, troque o `imageMso` por outro nome válido da galeria de ícones do Office (a faixa continua funcionando).

- [ ] **Passo 4: Escrever o `modRibbon.bas` completo**

```vb
Attribute VB_Name = "modRibbon"
Option Explicit
Private mRibbon As IRibbonUI
Public Const TUTORIAL_URL As String = "https://github.com/israeldias-hub/ThinQQ/blob/main/docs/tutorial/tutorial.html"

Public Sub Rx_OnLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
End Sub

Public Sub Loader_InvalidateRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

Private Function CoreOrWarn() As Boolean
    If Loader_CoreWorkbook() Is Nothing Then
        MsgBox "O módulo principal do ThinQQ não está carregado. Use 'Verificar atualizações' ou reinstale.", vbExclamation, APP_NAME
        Exit Function
    End If
    CoreOrWarn = True
End Function

Public Sub Rx_OnInsert(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_Insert", control.Tag
End Sub

Public Sub Rx_GetEnabled(control As IRibbonControl, ByRef enabled)
    On Error Resume Next
    enabled = False
    If Not Loader_CoreWorkbook() Is Nothing Then enabled = CBool(Loader_RunCore("Core_HasChart"))
End Sub

Public Sub Rx_OnToggle(control As IRibbonControl, pressed As Boolean)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_SetToggle", control.Tag, pressed
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_GetPressed(control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    pressed = False
    If Not Loader_CoreWorkbook() Is Nothing Then pressed = CBool(Loader_RunCore("Core_GetToggle", control.Tag))
End Sub

Public Sub Rx_OnDropdown(control As IRibbonControl, id As String, index As Integer)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_SetOption", control.Tag, id
End Sub

Public Sub Rx_GetSelectedId(control As IRibbonControl, ByRef itemId)
    On Error Resume Next
    Dim v As Variant
    v = Loader_RunCore("Core_GetOption", control.Tag)
    If IsEmpty(v) Or Len(CStr(v)) = 0 Then
        Select Case control.Tag
            Case "numfmt": itemId = "int"
            Case "palette": itemId = "padrao"
            Case "gap": itemId = "50"
        End Select
    Else
        itemId = CStr(v)
    End If
End Sub

Public Sub Rx_OnAnnotation(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_AddAnnotation", control.Tag
End Sub

Public Sub Rx_OnRemoveAnnotations(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_RemoveAnnotations"
End Sub

Public Sub Rx_OnRefresh(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_Refresh"
    Loader_InvalidateRibbon
End Sub

Public Sub Rx_OnChangeSource(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_ChangeSource"
End Sub

Public Sub Rx_OnSamples(control As IRibbonControl)
    On Error Resume Next
    If CoreOrWarn() Then Loader_RunCore "Core_InsertSamples"
End Sub

' tutorial.html da versão instalada; se não existir, a URL pública.
Public Function Loader_TutorialPath() As String
    Dim ver As String, p As String
    ver = Loader_FindNewestCoreVersion(Loader_VersionsDir())
    If Len(ver) > 0 Then
        p = Loader_VersionsDir() & "\" & ver & "\tutorial.html"
        If Len(Dir(p)) > 0 Then Loader_TutorialPath = p: Exit Function
    End If
    Loader_TutorialPath = TUTORIAL_URL
End Function

Public Sub Rx_OnTutorial(control As IRibbonControl)
    On Error GoTo Fail
    ThisWorkbook.FollowHyperlink Loader_TutorialPath()
    Exit Sub
Fail:
    MsgBox "Não foi possível abrir o tutorial: " & Err.Description, vbExclamation, APP_NAME
End Sub

Public Sub Rx_OnCheckUpdates(control As IRibbonControl)
    On Error Resume Next
    Updater_CheckInteractive
End Sub

Public Sub Rx_OnSettings(control As IRibbonControl)
    Dim cur As String, nv As Variant
    cur = Settings_UpdateSource()
    nv = Application.InputBox("Fonte de atualização (URL do manifest.json ou caminho de pasta/arquivo):", APP_NAME, cur, Type:=2)
    If VarType(nv) = vbBoolean Then Exit Sub            ' cancelado
    If Len(Trim$(CStr(nv))) = 0 Then
        Settings_Delete KEY_UPDATE_SOURCE
    Else
        Settings_Set KEY_UPDATE_SOURCE, Trim$(CStr(nv))
    End If
    MsgBox "Fonte de atualização: " & Settings_UpdateSource(), vbInformation, APP_NAME
End Sub

Public Sub Rx_OnAbout(control As IRibbonControl)
    MsgBox "ThinQQ" & vbLf & "Loader " & APP_VERSION & "  |  Core " & Loader_CurrentCoreVersion() & vbLf & vbLf & _
           "Gráficos de apresentação para Excel." & vbLf & "Pasta: " & Loader_Root(), vbInformation, APP_NAME
End Sub
```

- [ ] **Passo 5: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1 -Target loader`
Expected: todos `ok`.

- [ ] **Passo 6: Verificação manual da faixa (checklist)**

Feche o Excel; `pwsh build/build.ps1 -Target all -Configuration Release -Install`; abra o Excel com uma tabela de 4 linhas × 3 colunas com cabeçalho.

1. Com células selecionadas, o grupo "Gráfico selecionado" aparece desabilitado.
2. "Colunas empilhadas" insere o gráfico à direita da tabela, selecionado; o grupo fica habilitado e "Totais" e "Rótulos" aparecem pressionados.
3. Alternar "Totais" remove/recoloca os números acima das colunas; "Conectores" liga as linhas de série.
4. "Formato" → `1.234,5` muda os rótulos; "Paleta" → "Tons de cinza" recolore; "Espaço" → "Largo".
5. "Seta CAGR" abre o formulário com as categorias (1ª e última pré-selecionadas); OK desenha a seta e o rótulo `CAGR +x,x%`. Alterar um valor da tabela move a seta e atualiza o texto.
6. "Seta diferença" desenha a seta vertical com `+n (+p%)`; "Remover setas" apaga as duas.
7. Redimensionar o gráfico reposiciona as setas.
8. Tabela de 2 colunas com um `e`: "Cascata" desenha barras, rótulos e conectores; "Conectores" liga/desliga as linhas.
9. Clicar numa célula desabilita o grupo de novo. "Sobre" mostra as versões.

- [ ] **Passo 7: Commit**

```bash
git add src/loader/ribbon/customUI14.xml src/loader/modules/modRibbon.bas tests/vba/loader/TestRibbon.bas
git commit -m "feat(loader): faixa de opções ThinQQ completa com controles contextuais"
```

---

### Tarefa 13: Atualizador com manifesto, formulário e instalação a quente (loader + core)

**Files:**
- Create: `src/loader/modules/modUpdater.bas`
- Modify: `src/loader/modules/modLoaderMain.bas` (remover os stubs `Updater_StartupCheck`/`Updater_CheckInteractive`), `src/core/modules/modCoreApi.bas` (adicionar `Core_ApplyPendingLoader`)
- Test: `tests/vba/loader/TestUpdater.bas`

**Interfaces:**
- Produces (loader): `Updater_ShouldCheckNow(lastCheckText, nowValue, intervalHours) As Boolean`; `Updater_Evaluate(json, installedCore, installedLoader) As String` (`none|core|loader|both`); `Updater_StartupCheck()` (silencioso, 1×/intervalo); `Updater_CheckInteractive()`; `Updater_Check(silent)`; `Updater_Apply(json, what, Optional quiet = False) As Boolean`.
- Produces (core): `Core_ApplyPendingLoader()` (substitui o arquivo do loader registrado em `PendingLoaderPath` e o recarrega).
- Consumes: Tarefas 2, 3, 4, `frmUpdate` (Tarefa 1), `Core_LoaderPath` (Tarefa 11).
- Manifesto plano: `coreVersion`, `coreUrl`, `coreSha256`, `loaderVersion`, `loaderUrl`, `loaderSha256`, `tutorialUrl`, `publishedAt`, `notes` (spec §4.1).

- [ ] **Passo 1: Escrever os testes**

`tests/vba/loader/TestUpdater.bas`:

```vb
Attribute VB_Name = "TestUpdater"
Option Explicit

Private Function HttpBase() As String
    HttpBase = Transport_ReadTextFile(RepoPath() & "\tests\results\http-base.txt")
End Function

Private Function LocalManifest(coreUrl As String, sha As String) As String
    LocalManifest = "{""schemaVersion"":1,""coreVersion"":""9.9.9"",""coreUrl"":""" & Replace(coreUrl, "\", "\\") & _
        """,""coreSha256"":""" & sha & """,""loaderVersion"":""0.0.1"",""loaderUrl"":"""",""loaderSha256"":"""",""tutorialUrl"":"""",""publishedAt"":""2099-01-01"",""notes"":""x""}"
End Function

Public Sub Test_Updater_ShouldCheckNow()
    AssertTrue Updater_ShouldCheckNow("", Now, 24)
    AssertTrue Updater_ShouldCheckNow("lixo", Now, 24)
    AssertTrue Updater_ShouldCheckNow("2026-01-01 10:00:00", DateSerial(2026, 1, 2) + TimeSerial(10, 0, 1), 24)
    AssertFalse Updater_ShouldCheckNow("2026-01-01 10:00:00", DateSerial(2026, 1, 1) + TimeSerial(20, 0, 0), 24)
End Sub

Public Sub Test_Updater_Evaluate()
    Dim j As String
    j = "{""coreVersion"":""1.2.0"",""loaderVersion"":""1.0.0""}"
    AssertEqual "core", Updater_Evaluate(j, "1.1.0", "1.0.0")
    AssertEqual "none", Updater_Evaluate(j, "1.2.0", "1.0.0")
    AssertEqual "none", Updater_Evaluate(j, "1.3.0", "1.0.0")
    AssertEqual "both", Updater_Evaluate(j, "1.1.0", "0.9.0")
    AssertEqual "loader", Updater_Evaluate(j, "1.2.0", "0.9.0")
End Sub

Public Sub Test_Updater_Evaluate_MalformedJson_IsNone()
    AssertEqual "none", Updater_Evaluate("<html>erro</html>", "1.0.0", "1.0.0")
    AssertEqual "none", Updater_Evaluate("", "1.0.0", "1.0.0")
End Sub

Public Sub Test_Updater_Apply_FromLocalPath_InstallsCore()
    Dim root As String, src As String, sha As String, ok As Boolean, dest As String
    root = TempDir()
    Loader_SetRootOverride root
    src = RepoPath() & "\tests\fixtures\release\ThinQQ.core.test.xlam"
    sha = Transport_Sha256File(src)
    ok = Updater_Apply(LocalManifest(src, sha), "core", True)
    Loader_SetRootOverride ""
    AssertTrue ok, "Updater_Apply deveria retornar True"
    dest = root & "\versions\9.9.9\" & Loader_CoreFileName()
    AssertTrue Len(Dir(dest)) > 0, "core não foi copiado para versions\9.9.9"
    AssertEqual "9.9.9", Loader_FindNewestCoreVersion(root & "\versions")
    AssertEqual dest, Loader_CoreWorkbook().FullName
    AssertEqual "pong", Loader_RunCore("Core_Ping")
End Sub

Public Sub Test_Updater_Apply_BadSha_DoesNotInstall()
    Dim root As String, src As String, ok As Boolean
    root = TempDir()
    Loader_SetRootOverride root
    src = RepoPath() & "\tests\fixtures\release\ThinQQ.core.test.xlam"
    ok = Updater_Apply(LocalManifest(src, "0000000000000000000000000000000000000000000000000000000000000000"), "core", True)
    Loader_SetRootOverride ""
    AssertFalse ok
    AssertEqual "", Loader_FindNewestCoreVersion(root & "\versions")
End Sub

Public Sub Test_Updater_Apply_FromHttp()
    Dim root As String, json As String, ok As Boolean
    root = TempDir()
    json = Transport_FetchText(HttpBase() & "release/manifest.json")
    AssertEqual "9.9.9", JsonGetString(json, "coreVersion")
    Loader_SetRootOverride root
    ok = Updater_Apply(json, "core", True)
    Loader_SetRootOverride ""
    AssertTrue ok
    AssertTrue Len(Dir(root & "\versions\9.9.9\" & Loader_CoreFileName())) > 0
End Sub

Public Sub Test_Updater_Apply_LoaderPart_SchedulesPending()
    Dim root As String, src As String, sha As String, json As String, ok As Boolean
    root = TempDir()
    src = RepoPath() & "\tests\fixtures\http\hello.txt"
    sha = Transport_Sha256File(src)
    json = "{""loaderVersion"":""9.9.9"",""loaderUrl"":""" & Replace(src, "\", "\\") & """,""loaderSha256"":""" & sha & """}"
    Loader_SetRootOverride root
    Settings_Delete KEY_PENDING_LOADER
    ok = Updater_Apply(json, "loader", True)
    Loader_SetRootOverride ""
    AssertTrue ok
    AssertEqual root & "\tmp\ThinQQ.xlam", Settings_Get(KEY_PENDING_LOADER)
    Settings_Delete KEY_PENDING_LOADER        ' o Core_ApplyPendingLoader agendado vira no-op
End Sub
```

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target loader`
Expected: `not ok` em `TestUpdater.*`.

- [ ] **Passo 3: Implementar `modUpdater.bas`**

```vb
Attribute VB_Name = "modUpdater"
Option Explicit
' Verifica o manifesto (HTTP ou pasta), avisa o usuário e instala core/loader.

Public Function Updater_ShouldCheckNow(lastCheckText As String, nowValue As Date, intervalHours As Double) As Boolean
    If Len(lastCheckText) = 0 Then Updater_ShouldCheckNow = True: Exit Function
    If Not IsDate(lastCheckText) Then Updater_ShouldCheckNow = True: Exit Function
    Updater_ShouldCheckNow = ((nowValue - CDate(lastCheckText)) * 24 >= intervalHours)
End Function

Public Function Updater_Evaluate(json As String, installedCore As String, installedLoader As String) As String
    Dim needCore As Boolean, needLoader As Boolean
    needCore = CompareVersions(JsonGetString(json, "coreVersion"), installedCore) > 0
    needLoader = CompareVersions(JsonGetString(json, "loaderVersion"), installedLoader) > 0
    If needCore And needLoader Then
        Updater_Evaluate = "both"
    ElseIf needCore Then
        Updater_Evaluate = "core"
    ElseIf needLoader Then
        Updater_Evaluate = "loader"
    Else
        Updater_Evaluate = "none"
    End If
End Function

Public Sub Updater_StartupCheck()
    On Error Resume Next
    If Not Updater_ShouldCheckNow(Settings_Get(KEY_LAST_CHECK), Now, Val(Settings_Get(KEY_INTERVAL, "24"))) Then Exit Sub
    Updater_Check True
End Sub

Public Sub Updater_CheckInteractive()
    Updater_Check False
End Sub

Public Sub Updater_Check(silent As Boolean)
    On Error GoTo Fail
    Dim json As String, what As String, coreVer As String, installedCore As String, frm As frmUpdate
    json = Transport_FetchText(Settings_UpdateSource())
    Settings_Set KEY_LAST_CHECK, Format$(Now, "yyyy-mm-dd hh:nn:ss")
    installedCore = Loader_CurrentCoreVersion()
    what = Updater_Evaluate(json, installedCore, APP_VERSION)
    coreVer = JsonGetString(json, "coreVersion")
    If what = "none" Then
        If Not silent Then MsgBox "O ThinQQ já está na versão mais recente (" & installedCore & ").", vbInformation, APP_NAME
        Exit Sub
    End If
    If silent And Settings_Get(KEY_SKIPPED) = coreVer Then Exit Sub
    Set frm = New frmUpdate
    frm.Prepare "Nova versão do ThinQQ disponível", _
                "Instalada: " & installedCore & "   -   Nova: " & coreVer & "  (" & JsonGetString(json, "publishedAt") & ")", _
                JsonGetString(json, "notes"), silent
    frm.Show
    Select Case frm.Result
        Case "update": Updater_Apply json, what
        Case "skip": Settings_Set KEY_SKIPPED, coreVer
    End Select
    Unload frm
    Exit Sub
Fail:
    Loader_Log "Updater_Check: " & Err.Description
    If Not silent Then MsgBox "Não foi possível verificar atualizações." & vbLf & Err.Description & vbLf & vbLf & _
                              "Fonte: " & Settings_UpdateSource(), vbExclamation, APP_NAME
End Sub

' Baixa e instala. what = core | loader | both. Retorna True se tudo deu certo.
Public Function Updater_Apply(json As String, what As String, Optional quiet As Boolean = False) As Boolean
    On Error GoTo Fail
    Dim coreVer As String, tmp As String, dest As String, destDir As String, sha As String, url As String
    Loader_EnsureFolders
    Application.StatusBar = "ThinQQ: baixando atualização..."
    If what = "core" Or what = "both" Then
        coreVer = JsonGetString(json, "coreVersion")
        If Not IsValidVersion(coreVer) Then Err.Raise vbObjectError + 2000, "Updater", "Manifesto sem coreVersion válida."
        url = JsonGetString(json, "coreUrl")
        tmp = Loader_TmpDir() & "\" & Loader_CoreFileName()
        If Len(Dir(tmp)) > 0 Then Kill tmp
        Transport_DownloadFile url, tmp
        sha = LCase$(JsonGetString(json, "coreSha256"))
        If Len(sha) > 0 Then
            If Transport_Sha256File(tmp) <> sha Then Err.Raise vbObjectError + 2001, "Updater", "Arquivo baixado corrompido (SHA-256 não confere)."
        End If
        destDir = Loader_VersionsDir() & "\" & coreVer
        If Len(Dir(destDir, vbDirectory)) = 0 Then MkDir destDir
        dest = destDir & "\" & Loader_CoreFileName()
        If Len(Dir(dest)) > 0 Then Loader_CloseCore: Kill dest
        FileCopy tmp, dest
        Kill tmp
        url = JsonGetString(json, "tutorialUrl")
        If Len(url) > 0 Then
            On Error Resume Next
            Transport_DownloadFile url, destDir & "\tutorial.html"
            On Error GoTo Fail
        End If
        If Not Loader_OpenCore(dest) Then Err.Raise vbObjectError + 2003, "Updater", "Core instalado, mas não pôde ser carregado."
        Settings_Set KEY_LAST_NOTIFIED_CORE, coreVer
        Settings_Delete KEY_SKIPPED
        Loader_InvalidateRibbon
    End If
    If what = "loader" Or what = "both" Then
        url = JsonGetString(json, "loaderUrl")
        tmp = Loader_TmpDir() & "\ThinQQ.xlam"
        If Len(Dir(tmp)) > 0 Then Kill tmp
        Transport_DownloadFile url, tmp
        sha = LCase$(JsonGetString(json, "loaderSha256"))
        If Len(sha) > 0 Then
            If Transport_Sha256File(tmp) <> sha Then Err.Raise vbObjectError + 2002, "Updater", "Arquivo baixado corrompido (SHA-256 não confere)."
        End If
        Settings_Set KEY_PENDING_LOADER, tmp
        ' A troca do loader é feita pelo core, fora da pilha de chamadas do loader.
        Application.OnTime Now, "'" & Loader_CoreFileName() & "'!Core_ApplyPendingLoader"
    End If
    Application.StatusBar = False
    If Not quiet Then MsgBox "ThinQQ atualizado com sucesso." & IIf(Len(coreVer) > 0, " Versão " & coreVer & ".", ""), vbInformation, APP_NAME
    Updater_Apply = True
    Exit Function
Fail:
    Application.StatusBar = False
    Loader_Log "Updater_Apply: " & Err.Description
    If Not quiet Then MsgBox "Falha ao atualizar: " & Err.Description, vbExclamation, APP_NAME
    Updater_Apply = False
End Function
```

- [ ] **Passo 4: Remover os stubs do loader e adicionar `Core_ApplyPendingLoader` ao core**

Em `src/loader/modules/modLoaderMain.bas`, apague os dois stubs `Updater_StartupCheck` e `Updater_CheckInteractive` adicionados na Tarefa 4.

Acrescente ao fim de `src/core/modules/modCoreApi.bas`:

```vb
' Substitui o arquivo do loader (registrado em PendingLoaderPath) e o recarrega.
' Executado via OnTime a partir do core, para que nenhum código do loader esteja na pilha.
Public Sub Core_ApplyPendingLoader()
    On Error GoTo Fail
    Dim pending As String, target As String, ai As AddIn, found As Boolean
    pending = GetSetting("ThinQQ", "Updater", "PendingLoaderPath", "")
    target = mLoaderPath
    If Len(pending) = 0 Or Len(target) = 0 Then Exit Sub
    If Len(Dir(pending)) = 0 Then SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", "": Exit Sub
    For Each ai In Application.AddIns
        If StrComp(ai.FullName, target, vbTextCompare) = 0 Then ai.Installed = False: found = True
    Next
    On Error Resume Next
    Application.Workbooks(Util_FileName(target)).Close SaveChanges:=False
    On Error GoTo Fail
    FileCopy pending, target
    Kill pending
    SaveSetting "ThinQQ", "Updater", "PendingLoaderPath", ""
    If found Then
        For Each ai In Application.AddIns
            If StrComp(ai.FullName, target, vbTextCompare) = 0 Then ai.Installed = True
        Next
    Else
        Application.Workbooks.Open target
    End If
    Errors_Log "loader substituído: " & target
    Exit Sub
Fail:
    Errors_Handle "Core_ApplyPendingLoader"
End Sub
```

- [ ] **Passo 5: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1`
Expected: todos `ok` em core e loader.

- [ ] **Passo 6: Verificação manual do ciclo de atualização**

1. `pwsh build/build.ps1 -Target all -Configuration Release -Install` (Excel fechado); abra o Excel.
2. Crie uma pasta `C:\temp\tqq-release` e copie para lá `dist\ThinQQ.core.xlam` e um `manifest.json` com `coreVersion` `1.0.1`, `coreUrl` `C:\\temp\\tqq-release\\ThinQQ.core.xlam`, `coreSha256` = `(Get-FileHash C:\temp\tqq-release\ThinQQ.core.xlam -Algorithm SHA256).Hash.ToLower()`, `loaderVersion` `1.0.0`, `notes` `"• teste"`.
3. ThinQQ > Configurações → `C:\temp\tqq-release\manifest.json`. ThinQQ > Verificar atualizações → formulário com "Instalada: 1.0.0 - Nova: 1.0.1"; "Atualizar agora" → mensagem de sucesso; "Sobre" mostra Core 1.0.1 sem reiniciar o Excel; existe `%LOCALAPPDATA%\ThinQQ\versions\1.0.1\`.
4. Verificar de novo → "já está na versão mais recente".
5. Mude `coreVersion` para `1.0.2` e o SHA para `abc` → "Falha ao atualizar: … SHA-256" e nenhuma pasta `1.0.2`.

- [ ] **Passo 7: Commit**

```bash
git add src/loader/modules/modUpdater.bas src/loader/modules/modLoaderMain.bas src/core/modules/modCoreApi.bas tests/vba/loader/TestUpdater.bas
git commit -m "feat: atualizador com manifesto, verificação de SHA-256 e troca a quente do core"
```

---
### Tarefa 14: Instalador, desinstalador e script de release

**Files:**
- Create: `installer/install.ps1`, `installer/uninstall.ps1`, `installer/Instalar-ThinQQ.cmd`, `installer/LEIA-ME.txt`, `build/release.ps1`, `release/notes.txt`
- Test: `tests/ps/Test-Installer.ps1`

**Interfaces:**
- `install.ps1 [-Root <pasta>] [-RegistryRoot HKCU:\Software] [-UpdateSource <url|pasta>] [-PackageDir <pasta>] [-SkipExcelCheck]`: copia `ThinQQ.xlam` para `Root`, `ThinQQ.core.xlam` e `tutorial.html` para `Root\versions\<coreVersion>`, remove MOTW, cria Local Confiável, registra `OPEN`/`OPENn` = `/R "Root\ThinQQ.xlam"`, grava `UpdateSource` (parâmetro ou `manifestUrl` do manifesto).
- `uninstall.ps1 [-Root] [-RegistryRoot] [-KeepFiles]`: remove registro, local confiável, configurações e (opcionalmente) arquivos.
- `release.ps1 -Version x.y.z -BaseUrl <url|pasta> [-LoaderVersion x.y.z] [-NotesFile release/notes.txt] [-PublishTo <pasta>]`: atualiza `version.json`, build Release, gera `dist/manifest.json` (com `manifestUrl`), copia tutorial/instalador para `dist/`, cria `dist/ThinQQ-x.y.z.zip`, atualiza `release/manifest.json`.
- Compatível com Windows PowerShell 5.1 (instalador) — sem `??`, sem ternário.

- [ ] **Passo 1: Escrever o teste do instalador**

`tests/ps/Test-Installer.ps1`:

```powershell
# Executa install.ps1 / uninstall.ps1 contra pastas e chave de registro temporárias.
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$tmp  = Join-Path $env:TEMP ("tqq_inst_" + [guid]::NewGuid().ToString('N'))
$pkg  = Join-Path $tmp 'pkg'; $root = Join-Path $tmp 'root'
$reg  = 'HKCU:\Software\ThinQQ_Test_' + [guid]::NewGuid().ToString('N').Substring(0, 8)
New-Item -ItemType Directory -Force $pkg | Out-Null
Set-Content (Join-Path $pkg 'ThinQQ.xlam') 'loader'
Set-Content (Join-Path $pkg 'ThinQQ.core.xlam') 'core'
Set-Content (Join-Path $pkg 'tutorial.html') '<html></html>'
@{ coreVersion = '1.2.3'; loaderVersion = '1.0.0'; manifestUrl = 'https://exemplo/manifest.json' } | ConvertTo-Json | Set-Content (Join-Path $pkg 'manifest.json')
function Assert($cond, $msg) { if (-not $cond) { throw "FALHA: $msg" } }
try {
    & (Join-Path $repo 'installer\install.ps1') -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck
    Assert (Test-Path (Join-Path $root 'ThinQQ.xlam')) 'loader copiado'
    Assert (Test-Path (Join-Path $root 'versions\1.2.3\ThinQQ.core.xlam')) 'core copiado na pasta da versão'
    Assert (Test-Path (Join-Path $root 'versions\1.2.3\tutorial.html')) 'tutorial copiado'
    $opt = Get-ItemProperty (Join-Path $reg 'Microsoft\Office\16.0\Excel\Options')
    Assert ($opt.OPEN -eq ('/R "' + (Join-Path $root 'ThinQQ.xlam') + '"')) "valor OPEN: $($opt.OPEN)"
    $tl = Get-ItemProperty (Join-Path $reg 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ')
    Assert ($tl.AllowSubfolders -eq 1) 'local confiável com subpastas'
    $up = Get-ItemProperty (Join-Path $reg 'VB and VBA Program Settings\ThinQQ\Updater')
    Assert ($up.UpdateSource -eq 'https://exemplo/manifest.json') 'UpdateSource do manifesto'
    # Segunda instalação não duplica OPEN
    & (Join-Path $repo 'installer\install.ps1') -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck
    $opt = Get-ItemProperty (Join-Path $reg 'Microsoft\Office\16.0\Excel\Options')
    Assert (-not ($opt.PSObject.Properties.Name -contains 'OPEN1')) 'OPEN não duplicado'
    # UpdateSource explícito prevalece
    & (Join-Path $repo 'installer\install.ps1') -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck -UpdateSource '\\srv\tqq\manifest.json'
    Assert ((Get-ItemProperty (Join-Path $reg 'VB and VBA Program Settings\ThinQQ\Updater')).UpdateSource -eq '\\srv\tqq\manifest.json') 'UpdateSource explícito'
    & (Join-Path $repo 'installer\uninstall.ps1') -Root $root -RegistryRoot $reg -SkipExcelCheck
    Assert (-not (Test-Path (Join-Path $reg 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ'))) 'local confiável removido'
    Assert (-not ((Get-ItemProperty (Join-Path $reg 'Microsoft\Office\16.0\Excel\Options')).PSObject.Properties.Name -contains 'OPEN')) 'OPEN removido'
    Assert (-not (Test-Path $root)) 'arquivos removidos'
    Write-Host 'Test-Installer: OK' -ForegroundColor Green
} finally {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $reg -Recurse -Force -ErrorAction SilentlyContinue
}
```

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh tests/ps/Test-Installer.ps1`
Expected: erro "não foi possível localizar … install.ps1".

- [ ] **Passo 3: Escrever `installer/install.ps1`**

```powershell
# Instala o ThinQQ para o usuário atual (sem administrador). Compatível com Windows PowerShell 5.1.
param(
    [string]$Root = (Join-Path $env:LOCALAPPDATA 'ThinQQ'),
    [string]$RegistryRoot = 'HKCU:\Software',
    [string]$UpdateSource = '',
    [string]$PackageDir = $PSScriptRoot,
    [switch]$SkipExcelCheck
)
$ErrorActionPreference = 'Stop'
if (-not $SkipExcelCheck) {
    if (Get-Process excel -ErrorAction SilentlyContinue) { throw 'Feche o Excel antes de instalar o ThinQQ.' }
}
$manifestPath = Join-Path $PackageDir 'manifest.json'
if (-not (Test-Path $manifestPath)) { throw "manifest.json não encontrado em $PackageDir" }
$manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
$coreVer = $manifest.coreVersion
$verDir = Join-Path $Root "versions\$coreVer"
foreach ($d in @($verDir, (Join-Path $Root 'tmp'), (Join-Path $Root 'logs'))) { New-Item -ItemType Directory -Force $d | Out-Null }

Copy-Item (Join-Path $PackageDir 'ThinQQ.xlam') $Root -Force
Copy-Item (Join-Path $PackageDir 'ThinQQ.core.xlam') $verDir -Force
if (Test-Path (Join-Path $PackageDir 'tutorial.html')) { Copy-Item (Join-Path $PackageDir 'tutorial.html') $verDir -Force }
Get-ChildItem $Root -Recurse -File | Unblock-File -ErrorAction SilentlyContinue

# Local confiável (macros permitidas sem aviso)
$tl = Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ'
New-Item -Path $tl -Force | Out-Null
Set-ItemProperty -Path $tl -Name Path -Value ($Root.TrimEnd('\') + '\')
Set-ItemProperty -Path $tl -Name AllowSubfolders -Value 1 -Type DWord
Set-ItemProperty -Path $tl -Name Description -Value 'ThinQQ add-in'

# Registro do suplemento: OPEN, OPEN1, OPEN2... (sem lacunas)
$opt = Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Options'
New-Item -Path $opt -Force | Out-Null
$loaderPath = Join-Path $Root 'ThinQQ.xlam'
$existing = (Get-ItemProperty $opt).PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' }
$already = $existing | Where-Object { $_.Value -like '*ThinQQ.xlam*' }
if (-not $already) {
    $n = 0
    $names = @($existing | ForEach-Object { $_.Name })
    while ($true) {
        $candidate = 'OPEN'
        if ($n -gt 0) { $candidate = "OPEN$n" }
        if ($names -notcontains $candidate) { break }
        $n++
    }
    Set-ItemProperty -Path $opt -Name $candidate -Value ('/R "' + $loaderPath + '"')
}

# Fonte de atualização
$source = $UpdateSource
if (-not $source -and $manifest.PSObject.Properties['manifestUrl']) { $source = $manifest.manifestUrl }
if ($source) {
    $up = Join-Path $RegistryRoot 'VB and VBA Program Settings\ThinQQ\Updater'
    New-Item -Path $up -Force | Out-Null
    Set-ItemProperty -Path $up -Name UpdateSource -Value $source
}
Write-Host "ThinQQ $coreVer instalado em $Root. Abra o Excel e procure a aba 'ThinQQ'."
```

- [ ] **Passo 4: Escrever `installer/uninstall.ps1`, `Instalar-ThinQQ.cmd` e `LEIA-ME.txt`**

`installer/uninstall.ps1`:

```powershell
param(
    [string]$Root = (Join-Path $env:LOCALAPPDATA 'ThinQQ'),
    [string]$RegistryRoot = 'HKCU:\Software',
    [switch]$KeepFiles,
    [switch]$SkipExcelCheck
)
$ErrorActionPreference = 'Stop'
if (-not $SkipExcelCheck) {
    if (Get-Process excel -ErrorAction SilentlyContinue) { throw 'Feche o Excel antes de desinstalar o ThinQQ.' }
}
$opt = Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Options'
if (Test-Path $opt) {
    $props = (Get-ItemProperty $opt).PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' } | Sort-Object { if ($_.Name -eq 'OPEN') { 0 } else { [int]$_.Name.Substring(4) } }
    $values = @($props | Where-Object { $_.Value -notlike '*ThinQQ.xlam*' } | ForEach-Object { $_.Value })
    foreach ($p in $props) { Remove-ItemProperty -Path $opt -Name $p.Name }
    for ($i = 0; $i -lt $values.Count; $i++) {
        $name = 'OPEN'
        if ($i -gt 0) { $name = "OPEN$i" }
        Set-ItemProperty -Path $opt -Name $name -Value $values[$i]
    }
}
Remove-Item (Join-Path $RegistryRoot 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ') -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $RegistryRoot 'VB and VBA Program Settings\ThinQQ') -Recurse -Force -ErrorAction SilentlyContinue
if (-not $KeepFiles -and (Test-Path $Root)) { Remove-Item $Root -Recurse -Force }
Write-Host 'ThinQQ desinstalado.'
```

`installer/Instalar-ThinQQ.cmd`:

```bat
@echo off
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
echo Instalando o ThinQQ...
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
if errorlevel 1 (
  echo.
  echo Falha na instalacao. Feche o Excel e tente de novo, ou siga a instalacao manual do LEIA-ME.txt.
  pause
  exit /b 1
)
echo.
echo Pronto! Abra o Excel e procure a aba ThinQQ.
pause
```

`installer/LEIA-ME.txt`:

```
ThinQQ - gráficos de apresentação para Excel
============================================

INSTALAÇÃO AUTOMÁTICA
1. Feche o Excel.
2. Dê dois cliques em "Instalar-ThinQQ.cmd". Se o Windows perguntar, clique em "Executar".
3. Abra o Excel: a aba "ThinQQ" aparece na faixa de opções.

INSTALAÇÃO MANUAL (se o script for bloqueado)
1. Crie a pasta %LOCALAPPDATA%\ThinQQ (cole esse caminho na barra de endereços do Explorador).
2. Copie ThinQQ.xlam para essa pasta e ThinQQ.core.xlam para %LOCALAPPDATA%\ThinQQ\versions\<versão>\
   (a versão está no manifest.json, ex.: versions\1.0.0\). Copie tutorial.html para a mesma pasta da versão.
3. Em cada arquivo .xlam: botão direito > Propriedades > marque "Desbloquear" > OK.
4. Excel > Arquivo > Opções > Central de Confiabilidade > Configurações > Locais Confiáveis >
   Adicionar novo local: %LOCALAPPDATA%\ThinQQ, marcando "Subpastas também são confiáveis".
5. Excel > Arquivo > Opções > Suplementos > Gerenciar: Suplementos do Excel > Ir... > Procurar >
   escolha %LOCALAPPDATA%\ThinQQ\ThinQQ.xlam > OK.

ATUALIZAÇÕES
O ThinQQ verifica novas versões ao abrir o Excel (1x por dia) e avisa. Use ThinQQ > Verificar
atualizações para forçar. A fonte pode ser alterada em ThinQQ > Configurações.

DESINSTALAÇÃO
Feche o Excel e execute: powershell -ExecutionPolicy Bypass -File uninstall.ps1
```

- [ ] **Passo 5: Rodar o teste do instalador**

Run: `pwsh tests/ps/Test-Installer.ps1` e também `powershell -NoProfile -ExecutionPolicy Bypass -File tests/ps/Test-Installer.ps1` (valida compatibilidade com 5.1)
Expected: `Test-Installer: OK` nos dois.

- [ ] **Passo 6: Escrever `build/release.ps1` e `release/notes.txt`**

`release/notes.txt`:

```
• Primeira versão: colunas empilhadas, agrupadas, 100% e linha; cascata com subtotais "e"; totais, rótulos vivos, conectores; setas de CAGR e diferença; paletas e formatos; exemplos e tutorial; atualização automática.
```

`build/release.ps1`:

```powershell
<#
.SYNOPSIS  Gera um release: version.json, build Release, manifest.json, zip e (opcional) publicação em pasta.
.EXAMPLE   pwsh build/release.ps1 -Version 1.0.0 -BaseUrl https://github.com/israeldias-hub/ThinQQ/releases/download/v1.0.0
.EXAMPLE   pwsh build/release.ps1 -Version 1.0.0 -BaseUrl '\\servidor\ThinQQ' -PublishTo '\\servidor\ThinQQ'
#>
param(
    [Parameter(Mandatory)][string]$Version,
    [string]$LoaderVersion,
    [Parameter(Mandatory)][string]$BaseUrl,
    [string]$NotesFile = (Join-Path $PSScriptRoot '..\release\notes.txt'),
    [string]$PublishTo
)
. "$PSScriptRoot\lib\Build-Common.ps1"
if ($Version -notmatch '^\d+\.\d+\.\d+$') { throw 'Versão deve ser x.y.z' }
$versionPath = Join-Path $RepoRoot 'version.json'
$versions = Get-Content $versionPath -Raw | ConvertFrom-Json
$versions.core = $Version
if ($LoaderVersion) { $versions.loader = $LoaderVersion }
$versions | ConvertTo-Json | Set-Content $versionPath -Encoding utf8

& "$PSScriptRoot\build.ps1" -Target all -Configuration Release

$sep = '\'
if ($BaseUrl -match '^https?://') { $sep = '/' }
$base = $BaseUrl.TrimEnd('/', '\') + $sep
$notes = ''
if (Test-Path $NotesFile) { $notes = (Get-Content $NotesFile -Raw).Trim() }
$manifest = [ordered]@{
    schemaVersion = 1
    coreVersion   = $versions.core
    coreUrl       = $base + 'ThinQQ.core.xlam'
    coreSha256    = (Get-FileHash (Join-Path $DistDir 'ThinQQ.core.xlam') -Algorithm SHA256).Hash.ToLower()
    loaderVersion = $versions.loader
    loaderUrl     = $base + 'ThinQQ.xlam'
    loaderSha256  = (Get-FileHash (Join-Path $DistDir 'ThinQQ.xlam') -Algorithm SHA256).Hash.ToLower()
    tutorialUrl   = $base + 'tutorial.html'
    manifestUrl   = $base + 'manifest.json'
    publishedAt   = (Get-Date -Format 'yyyy-MM-dd')
    notes         = $notes
}
$manifest | ConvertTo-Json | Set-Content (Join-Path $DistDir 'manifest.json') -Encoding utf8
Copy-Item (Join-Path $RepoRoot 'docs\tutorial\tutorial.html') $DistDir -Force
Copy-Item (Join-Path $RepoRoot 'installer\*') $DistDir -Force
New-Item -ItemType Directory -Force (Join-Path $RepoRoot 'release') | Out-Null
Copy-Item (Join-Path $DistDir 'manifest.json') (Join-Path $RepoRoot 'release\manifest.json') -Force

$files = @('ThinQQ.xlam', 'ThinQQ.core.xlam', 'manifest.json', 'tutorial.html', 'install.ps1', 'uninstall.ps1', 'Instalar-ThinQQ.cmd', 'LEIA-ME.txt') | ForEach-Object { Join-Path $DistDir $_ }
$zip = Join-Path $DistDir "ThinQQ-$($versions.core).zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path $files -DestinationPath $zip
if ($PublishTo) {
    New-Item -ItemType Directory -Force $PublishTo | Out-Null
    Copy-Item $files $PublishTo -Force
    Copy-Item $zip $PublishTo -Force
    Write-Host "Publicado em $PublishTo"
}
Write-Host "Release pronto: $zip"
Write-Host "Publique os arquivos de dist/ no endereço base $BaseUrl e faça commit de release/manifest.json e version.json."
```

- [ ] **Passo 7: Rodar um release de teste**

Run: `pwsh build/release.ps1 -Version 1.0.0 -BaseUrl C:\temp\tqq-release -PublishTo C:\temp\tqq-release`
Expected: `dist\ThinQQ-1.0.0.zip` criado, `release\manifest.json` com `coreSha256` de 64 hex e `manifestUrl` = `C:\temp\tqq-release\manifest.json`; pasta `C:\temp\tqq-release` com 8 arquivos + zip. (`docs/tutorial/tutorial.html` precisa existir: se ainda não existir, crie um arquivo com `<html><body>ThinQQ</body></html>` e substitua na Tarefa 15.)

Teste ponta a ponta do instalador real: feche o Excel, rode `C:\temp\tqq-release\Instalar-ThinQQ.cmd`, abra o Excel → aba ThinQQ presente; ThinQQ > Sobre → Core 1.0.0.

- [ ] **Passo 8: Commit**

```bash
git add installer build/release.ps1 release/notes.txt release/manifest.json version.json tests/ps/Test-Installer.ps1
git commit -m "build: instalador por usuário, desinstalador e script de release com manifesto"
```

---

### Tarefa 15: Exemplos, tutorial, documentação e release 1.0.0

**Files:**
- Create: `src/core/modules/modSamples.bas`, `docs/tutorial/tutorial.html`, `CHANGELOG.md`
- Modify: `src/core/modules/modCoreApi.bas` (adicionar `Core_InsertSamples`), `README.md` (substituir inteiro)
- Test: `tests/vba/core/TestSamples.bas`

**Interfaces:**
- Produces: `Samples_Create() As Workbook` (pasta nova "ThinQQ - Exemplos" com 4 tabelas e 4 gráficos prontos, incluindo uma seta CAGR e uma de diferença); `Core_InsertSamples()`.

- [ ] **Passo 1: Escrever o teste**

`tests/vba/core/TestSamples.bas`:

```vb
Attribute VB_Name = "TestSamples"
Option Explicit

Public Sub Test_Samples_Create()
    Dim wb As Workbook, ws As Worksheet, co As ChartObject, kinds As String
    Set wb = Samples_Create()
    Set ws = wb.Worksheets("Exemplos")
    AssertEqual 4, ws.ChartObjects.Count
    For Each co In ws.ChartObjects
        AssertTrue Registry_IsThinQQChart(co.Chart)
        kinds = kinds & Registry_ChartType(co.Chart) & ";"
    Next
    AssertTrue InStr(kinds, "stacked;") > 0 And InStr(kinds, "waterfall;") > 0 And InStr(kinds, "clustered;") > 0 And InStr(kinds, "line;") > 0, kinds
    AssertTrue Not ws.ChartObjects(1).Chart.Shapes("tqq_ann_cagr_1") Is Nothing
    AssertTrue Not ws.ChartObjects(1).Chart.Shapes("tqq_ann_diff_1") Is Nothing
    AssertEqual "e", LCase$(ws.Range("B27").Value)
    CloseTemp wb
End Sub
```

- [ ] **Passo 2: Rodar para ver falhar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: `not ok` em `TestSamples.Test_Samples_Create`.

- [ ] **Passo 3: Implementar `modSamples.bas` e `Core_InsertSamples`**

`src/core/modules/modSamples.bas`:

```vb
Attribute VB_Name = "modSamples"
Option Explicit

Private Sub WriteTable(topLeft As Range, rows As Variant)
    Dim i As Long, j As Long, r As Variant
    For i = LBound(rows) To UBound(rows)
        r = rows(i)
        For j = LBound(r) To UBound(r)
            topLeft.Offset(i - LBound(rows), j - LBound(r)).Value = r(j)
        Next
    Next
    topLeft.Resize(1, UBound(rows(LBound(rows))) - LBound(rows(LBound(rows))) + 1).Font.Bold = True
End Sub

Public Function Samples_Create() As Workbook
    Dim wb As Workbook, ws As Worksheet, d As clsChartData, co As ChartObject
    Set wb = Application.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Exemplos"
    ws.Columns("A").ColumnWidth = 16

    ws.Range("A1").Value = "1) Colunas empilhadas com totais, CAGR e diferença (R$ mi)"
    WriteTable ws.Range("A2"), Array(Array("Ano", "Sul", "Sudeste", "Nordeste"), Array(2022, 120, 210, 65), _
                                     Array(2023, 131, 228, 72), Array(2024, 140, 251, 80), Array(2025, 158, 276, 91))
    Set d = ChartData_FromRange(ws.Range("A2:D6"))
    Set co = Stacked_Insert(ws, d, "stacked")
    Ann_AddCagr co.Chart, 1, 4
    Ann_AddDifference co.Chart, 1, 4

    ws.Range("A22").Value = "2) Cascata: escreva 'e' para uma barra de total calculada"
    WriteTable ws.Range("A23"), Array(Array("Etapa", "Valor"), Array("Receita", 1200), Array("Custos", -700), _
                                      Array("Despesas", -250), Array("EBITDA", "e"), Array("Impostos", -80), Array("Lucro", "e"))
    Set d = ChartData_FromRange(ws.Range("A23:B29"))
    Waterfall_Insert ws, d

    ws.Range("A43").Value = "3) Colunas agrupadas"
    WriteTable ws.Range("A44"), Array(Array("Trimestre", "Loja", "Online"), Array("T1", 420, 180), _
                                      Array("T2", 455, 230), Array("T3", 430, 290), Array("T4", 510, 360))
    Set d = ChartData_FromRange(ws.Range("A44:C48"))
    Stacked_Insert ws, d, "clustered"

    ws.Range("A64").Value = "4) Linha"
    WriteTable ws.Range("A65"), Array(Array("Mês", "Clientes ativos (mil)"), Array("jan", 48), Array("fev", 51), _
                                      Array("mar", 55), Array("abr", 54), Array("mai", 60), Array("jun", 66))
    Set d = ChartData_FromRange(ws.Range("A65:B71"))
    Stacked_Insert ws, d, "line"

    ws.Range("A1").Select
    Set Samples_Create = wb
End Function
```

Acrescente ao fim de `src/core/modules/modCoreApi.bas`:

```vb
Public Sub Core_InsertSamples()
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Samples_Create
    Application.ScreenUpdating = True
    Watchers_Rehook
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    Errors_Handle "Core_InsertSamples"
End Sub
```

- [ ] **Passo 4: Rodar para ver passar**

Run: `pwsh build/run-tests.ps1 -Target core`
Expected: todos `ok`. (O teste espera `e` em `B27`: linha 23 cabeçalho, 24 Receita, 25 Custos, 26 Despesas, 27 EBITDA.)

- [ ] **Passo 5: Escrever o tutorial**

`docs/tutorial/tutorial.html`:

```html
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>ThinQQ — Tutorial</title>
<style>
  :root { --ink:#1f2937; --muted:#6b7280; --accent:#1f3864; --accent2:#4472c4; --bg:#ffffff; --card:#f3f6fb; --line:#e5e7eb; }
  * { box-sizing:border-box; }
  body { margin:0; font-family:"Segoe UI", system-ui, sans-serif; color:var(--ink); background:var(--bg); line-height:1.55; }
  header { background:var(--accent); color:#fff; padding:40px 24px; }
  header h1 { margin:0 0 8px; font-size:32px; font-weight:600; }
  header p { margin:0; opacity:.9; }
  main { max-width:860px; margin:0 auto; padding:24px; }
  nav { background:var(--card); border:1px solid var(--line); border-radius:8px; padding:16px 20px; margin-bottom:28px; }
  nav a { display:inline-block; margin:4px 12px 4px 0; color:var(--accent2); text-decoration:none; }
  h2 { font-size:22px; margin:36px 0 12px; color:var(--accent); border-bottom:2px solid var(--line); padding-bottom:6px; }
  h3 { font-size:17px; margin:22px 0 8px; }
  ol, ul { padding-left:22px; }
  li { margin:4px 0; }
  code, kbd { background:var(--card); border:1px solid var(--line); border-radius:4px; padding:1px 6px; font-size:.93em; }
  table { border-collapse:collapse; margin:12px 0; font-size:.95em; }
  th, td { border:1px solid var(--line); padding:6px 12px; text-align:left; }
  th { background:var(--card); }
  .tip { background:#fff8e6; border-left:4px solid #eda224; padding:10px 14px; border-radius:4px; margin:14px 0; }
  .ribbon { display:grid; grid-template-columns:repeat(auto-fit,minmax(220px,1fr)); gap:12px; margin:12px 0; }
  .ribbon div { background:var(--card); border:1px solid var(--line); border-radius:8px; padding:12px 14px; }
  .ribbon b { display:block; color:var(--accent); margin-bottom:6px; }
  footer { color:var(--muted); font-size:.9em; text-align:center; padding:32px 0; }
</style>
</head>
<body>
<header>
  <h1>ThinQQ</h1>
  <p>Gráficos de apresentação no Excel, em poucos cliques — tutorial de uso</p>
</header>
<main>
<nav>
  <a href="#instalar">Instalação</a><a href="#primeiro">Primeiro gráfico</a><a href="#cascata">Cascata</a>
  <a href="#opcoes">Opções do gráfico</a><a href="#setas">Setas CAGR e diferença</a><a href="#atualizar">Atualizações</a>
  <a href="#dicas">Dicas</a><a href="#problemas">Problemas comuns</a>
</nav>

<h2 id="instalar">1. Instalação</h2>
<ol>
  <li>Feche o Excel.</li>
  <li>Descompacte o arquivo <code>ThinQQ-x.y.z.zip</code> e dê dois cliques em <code>Instalar-ThinQQ.cmd</code>.</li>
  <li>Abra o Excel: a aba <b>ThinQQ</b> aparece na faixa de opções.</li>
</ol>
<p>Se o script for bloqueado pela política da empresa, siga a instalação manual descrita no <code>LEIA-ME.txt</code> (copiar os arquivos, desbloquear, local confiável e registrar o suplemento).</p>

<h2 id="primeiro">2. Primeiro gráfico em 3 passos</h2>
<ol>
  <li>Monte a tabela com <b>uma linha de cabeçalho</b> (nomes das séries) e <b>uma coluna de categorias</b>:
    <table><tr><th>Ano</th><th>Sul</th><th>Sudeste</th><th>Nordeste</th></tr>
    <tr><td>2022</td><td>120</td><td>210</td><td>65</td></tr><tr><td>2023</td><td>131</td><td>228</td><td>72</td></tr>
    <tr><td>2024</td><td>140</td><td>251</td><td>80</td></tr><tr><td>2025</td><td>158</td><td>276</td><td>91</td></tr></table></li>
  <li>Clique em qualquer célula da tabela (ou selecione-a inteira).</li>
  <li>Na aba ThinQQ, clique em <b>Colunas empilhadas</b>. O gráfico aparece à direita, já com totais, rótulos e paleta limpa.</li>
</ol>
<div class="tip">O gráfico fica <b>ligado à tabela</b>: altere um número e tudo se atualiza — inclusive totais, setas e conectores. Quem receber o arquivo não precisa ter o ThinQQ instalado.</div>

<h2 id="cascata">3. Cascata (ponte)</h2>
<p>Use uma tabela de <b>duas colunas</b>: etapa e valor. Valores positivos sobem, negativos descem. Para uma barra de <b>total ou subtotal calculada automaticamente</b>, escreva <code>e</code> (ou <code>t</code>, <code>total</code>, <code>subtotal</code>) no lugar do valor:</p>
<table><tr><th>Etapa</th><th>Valor</th></tr><tr><td>Receita</td><td>1200</td></tr><tr><td>Custos</td><td>-700</td></tr>
<tr><td>Despesas</td><td>-250</td></tr><tr><td>EBITDA</td><td>e</td></tr><tr><td>Impostos</td><td>-80</td></tr><tr><td>Lucro</td><td>e</td></tr></table>
<p>Selecione a tabela e clique em <b>Cascata</b>. Barras que cruzam o zero são desenhadas corretamente; os conectores ligam o fim de cada barra ao início da próxima.</p>

<h2 id="opcoes">4. Opções do gráfico selecionado</h2>
<p>Clique no gráfico: o grupo <b>Gráfico selecionado</b> fica ativo e mostra o estado atual.</p>
<div class="ribbon">
  <div><b>Totais</b>Mostra/oculta os totais acima das colunas (empilhadas e 100%).</div>
  <div><b>Rótulos</b>Mostra/oculta os valores nos segmentos, pontos ou barras.</div>
  <div><b>Conectores</b>Linhas entre as colunas (empilhadas) ou entre as barras (cascata).</div>
  <div><b>Formato</b>Inteiro, 1 ou 2 decimais, percentual, milhares ou milhões.</div>
  <div><b>Paleta</b>Padrão ThinQQ, tons de cinza ou cores do tema da pasta.</div>
  <div><b>Espaço</b>Largura do vão entre colunas.</div>
  <div><b>Atualizar</b>Reposiciona setas e rótulos e limpa dados internos sem uso.</div>
  <div><b>Alterar dados</b>Troca a tabela de origem mantendo posição e opções.</div>
</div>

<h2 id="setas">5. Setas de CAGR e de diferença</h2>
<ol>
  <li>Selecione o gráfico e clique em <b>Seta CAGR</b> ou <b>Seta diferença</b>.</li>
  <li>Escolha a categoria inicial e a final (por padrão, primeira e última) e confirme.</li>
</ol>
<ul>
  <li><b>CAGR</b>: taxa composta anual entre os totais das duas categorias. Se as categorias forem anos (ex.: 2022…2025) ou datas, o número de anos é calculado por elas; senão, cada categoria conta como um período.</li>
  <li><b>Diferença</b>: seta vertical ao lado da coluna final com a diferença absoluta e relativa (ex.: <code>+90 (+60%)</code>) e uma linha tracejada no nível inicial.</li>
  <li>Em gráficos agrupados e de linha, as setas usam a <b>primeira série</b>.</li>
  <li><b>Remover setas</b> apaga as setas (os conectores da cascata ficam).</li>
</ul>
<div class="tip">As setas acompanham os dados e o tamanho do gráfico. Se mover o gráfico para outra planilha, clique em <b>Atualizar</b>.</div>

<h2 id="atualizar">6. Atualizações</h2>
<p>Ao abrir o Excel (no máximo uma vez por dia) o ThinQQ consulta a fonte de atualização. Havendo versão nova, mostra as novidades e os botões <b>Atualizar agora</b>, <b>Depois</b> e <b>Pular esta versão</b>. A atualização é baixada, conferida (SHA-256) e aplicada <b>sem reiniciar o Excel</b>.</p>
<ul>
  <li><b>Verificar atualizações</b>: consulta imediatamente.</li>
  <li><b>Configurações</b>: muda a fonte (URL do <code>manifest.json</code> ou caminho de pasta/arquivo, inclusive pasta sincronizada do OneDrive/Teams, com <code>%USERPROFILE%</code>).</li>
  <li><b>Sobre</b>: versões instaladas e pasta de instalação.</li>
</ul>

<h2 id="dicas">7. Dicas</h2>
<ul>
  <li>Clique em <b>Exemplos</b> para abrir uma pasta com as quatro tabelas deste tutorial e os gráficos prontos.</li>
  <li>Células vazias ou com texto no meio dos valores contam como zero.</li>
  <li>Os gráficos são gráficos Excel comuns: você pode redimensionar, mover, mudar fonte e cores manualmente.</li>
  <li>Os dados auxiliares ficam numa planilha oculta <code>_ThinQQ</code>; não a edite. <b>Atualizar</b> remove blocos de gráficos apagados.</li>
  <li>Precisa colar no PowerPoint? Copie o gráfico e cole como imagem (ou mantenha o vínculo com o Excel).</li>
</ul>

<h2 id="problemas">8. Problemas comuns</h2>
<table>
<tr><th>Sintoma</th><th>O que fazer</th></tr>
<tr><td>A aba ThinQQ não aparece</td><td>Confirme que o Excel foi reaberto após a instalação; em Arquivo &gt; Opções &gt; Suplementos &gt; Ir… marque "ThinQQ". Se os botões aparecem mas nada acontece, use <b>Verificar atualizações</b> para baixar o core.</td></tr>
<tr><td>"As macros foram desabilitadas"</td><td>Os arquivos precisam estar desbloqueados e em um Local Confiável (veja LEIA-ME.txt). Se a política da empresa bloquear locais confiáveis, peça à TI para assinar o suplemento.</td></tr>
<tr><td>Botão cinza</td><td>O grupo "Gráfico selecionado" só ativa com um gráfico criado pelo ThinQQ selecionado.</td></tr>
<tr><td>Setas fora do lugar</td><td>Clique em <b>Atualizar</b>. Isso recalcula as posições pelo tamanho atual do gráfico.</td></tr>
<tr><td>Erro ao atualizar</td><td>Verifique a fonte em <b>Configurações</b> e o acesso à rede/pasta. Detalhes em <code>%LOCALAPPDATA%\ThinQQ\logs\thinqq.log</code>.</td></tr>
</table>
</main>
<footer>ThinQQ — tutorial da versão 1.0</footer>
</body>
</html>
```

- [ ] **Passo 6: Escrever `README.md` e `CHANGELOG.md`**

`README.md`:

```markdown
# ThinQQ

Suplemento Excel (VBA) para gráficos de apresentação no estilo think-cell: colunas empilhadas,
agrupadas, 100%, linha e cascata, com totais, rótulos vivos, conectores e setas de CAGR/diferença.
Os gráficos gerados são gráficos Excel comuns ligados aos dados — quem recebe o arquivo não precisa do ThinQQ.

- Design: `docs/superpowers/specs/2026-10-10-thinqq-design.md`
- Plano da Fase 1: `docs/superpowers/plans/2026-10-10-thinqq-fase1.md`
- Tutorial do usuário: `docs/tutorial/tutorial.html`

## Arquitetura
- `ThinQQ.xlam` (loader): faixa de opções, boot, atualizador, configurações.
- `ThinQQ.core.xlam` (core): motor de gráficos. Carregado de `%LOCALAPPDATA%\ThinQQ\versions\<x.y.z>\`.
- Fonte em texto em `src/`; os `.xlam` são montados pelo build (COM) e não ficam no Git.

## Pré-requisitos de desenvolvimento
- Excel 2016+ para Windows; PowerShell 7 (`pwsh`); Git; Node.js (só para os testes HTTP).
- Excel > Arquivo > Opções > Central de Confiabilidade > Configurações de Macro > marcar
  "Confiar no acesso ao modelo de objeto do projeto VBA".

## Comandos
```powershell
pwsh build/build.ps1 -Target all -Configuration Release          # monta dist/
pwsh build/build.ps1 -Target all -Configuration Release -Install # + instala em %LOCALAPPDATA%\ThinQQ (Excel fechado)
pwsh build/run-tests.ps1                                          # testes VBA (core e loader) no Excel invisível
pwsh tests/ps/Test-Installer.ps1                                  # testes do instalador
pwsh build/release.ps1 -Version 1.1.0 -BaseUrl <url-ou-pasta> [-PublishTo <pasta>]
```

## Publicar uma versão
1. Atualize `release/notes.txt` e `CHANGELOG.md`.
2. `pwsh build/release.ps1 -Version x.y.z -BaseUrl https://github.com/israeldias-hub/ThinQQ/releases/download/vx.y.z`
   (ou `-BaseUrl \\servidor\ThinQQ -PublishTo \\servidor\ThinQQ` para pasta de rede/OneDrive).
3. Para GitHub: crie a release `vx.y.z` com os arquivos de `dist/` como anexos; faça commit de
   `version.json` e `release/manifest.json` (a fonte padrão aponta para `release/manifest.json` no `main`).
4. Os usuários recebem o aviso na próxima abertura do Excel.

Troque `israeldias-hub` em `src/loader/modules/modSettings.bas` (`DEFAULT_UPDATE_SOURCE`) e em
`src/loader/modules/modRibbon.bas` (`TUTORIAL_URL`) pelo repositório real antes do primeiro release.

## Estrutura
```
build/      scripts de build, testes e release
src/        fontes VBA (loader e core), formulários (JSON + .vba), ribbon XML
tests/      testes VBA (TAP), servidor de fixtures, teste do instalador
installer/  instalador/desinstalador por usuário
release/    manifest.json publicado e notas
docs/       spec, plano e tutorial
```
```

`CHANGELOG.md`:

```markdown
# Changelog

## 1.0.0 — 2026-10-10
- Gráficos: colunas empilhadas, agrupadas, 100% e linha; cascata com subtotais `e` e cruzamento do zero.
- Totais, rótulos ligados às células, conectores (linhas de série e conectores da cascata).
- Setas de CAGR e de diferença, reposicionadas automaticamente ao mudar dados ou tamanho.
- Paletas (Padrão, Cinza, Tema), formatos numéricos e espaço entre colunas.
- Faixa de opções ThinQQ com controles contextuais; pasta de exemplos; tutorial.
- Instalador por usuário; atualização por manifesto (HTTP ou pasta) com conferência SHA-256 e troca a quente.
```

- [ ] **Passo 7: Verificação final e release**

1. Run: `pwsh build/run-tests.ps1` → `Todos os testes passaram`.
2. Run: `pwsh tests/ps/Test-Installer.ps1` → `Test-Installer: OK`.
3. Feche o Excel; `pwsh build/release.ps1 -Version 1.0.0 -BaseUrl C:\temp\tqq-release -PublishTo C:\temp\tqq-release`; instale pelo `Instalar-ThinQQ.cmd` da pasta publicada; abra o Excel.
4. Checklist: ThinQQ > Exemplos abre a pasta com 4 gráficos (o primeiro com setas CAGR e diferença); Tutorial abre o `tutorial.html` da versão instalada no navegador; Sobre mostra Loader 1.0.0 / Core 1.0.0; a checklist da Tarefa 12 passa inteira.

- [ ] **Passo 8: Commit e tag**

```bash
git add -A
git commit -m "feat: exemplos, tutorial e documentação da versão 1.0.0"
git tag -a v1.0.0 -m "ThinQQ 1.0.0"
```

---

## Checklist de QA manual antes de distribuir

1. Instalação limpa em outra máquina (ou outro usuário Windows) só com o zip: aba aparece, Exemplos funciona.
2. Arquivo salvo com gráficos ThinQQ aberto em um Excel **sem** o ThinQQ: gráficos intactos, sem avisos de macro (o arquivo do usuário não tem macros).
3. Tabela com nome de planilha com espaço e apóstrofo; tabela com 2 categorias; tabela com 12 categorias.
4. Cascata totalmente negativa e cascata começando com `e`.
5. Atualização por pasta de rede e por URL HTTP; manifesto com SHA errado; fonte inacessível (mensagem amigável ao verificar manualmente, silêncio no início).
6. Fechar e reabrir o Excel: observadores reassociados (seta acompanha mudança de dados sem clicar em Atualizar).
