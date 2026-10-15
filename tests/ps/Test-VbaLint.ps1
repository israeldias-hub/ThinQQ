# Prova que o lint VBA pega continuacao "_" invalida e que os fontes reais passam.
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
. (Join-Path $RepoRoot 'build\lib\Vba-Lint.ps1')
$script:fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }

$bad1 = @('Public Sub A()', '    x = "a" & _           "b"', 'End Sub')
Check ((Test-VbaSourceLines $bad1 'bad1.bas').Count -eq 1) 'pega _ seguido de texto na mesma linha'
$bad2 = @('Public Sub A()', '    x = "a" & _', '', 'End Sub')
Check ((Test-VbaSourceLines $bad2 'bad2.bas').Count -eq 1) 'pega _ seguido de linha em branco'
$bad3 = @('    x = "a" & _')
Check ((Test-VbaSourceLines $bad3 'bad3.bas').Count -eq 1) 'pega _ no fim do arquivo'
$good = @('    x = "a" & _', '        "b"', "    y = ""texto com  _ dentro"" ' comentario com _ e mais", '    z = "aspas "" _ x"', '    my_var = 1')
Check ((Test-VbaSourceLines $good 'good.bas').Count -eq 0) 'ignora _ em string, comentario e identificador'
$arrow = [string][char]0x2192
Check ((Test-VbaCp1252Lines @("' seta $arrow aqui") 'arrow.bas').Count -eq 1) 'M13: pega caractere fora do cp1252 (seta)'
$okChars = -join ([char[]](0xE7, 0xE3, 0xE9, 0x201C, 0x201D, 0x2014, 0x20AC))
Check ((Test-VbaCp1252Lines @("' acentos e pontuacao: $okChars") 'ok.bas').Count -eq 0) 'M13: aceita acentos e pontuacao tipografica do cp1252'
$tw = Join-Path ([IO.Path]::GetTempPath()) ('tqq_twin_' + [guid]::NewGuid().ToString('N'))
foreach ($d in 'src\loader\modules', 'src\core\modules', 'tests\vba\loader', 'tests\vba\core') { New-Item -ItemType Directory -Force (Join-Path $tw $d) | Out-Null }
foreach ($p in 'src\loader\modules\modSha256.bas', 'src\core\modules\modSha256.bas', 'tests\vba\loader\TestSha256.bas', 'tests\vba\core\TestSha256.bas') { Set-Content (Join-Path $tw $p) 'x' }
Check ((Test-VbaTwinFiles $tw).Count -eq 0) 'gemeos identicos passam'
Set-Content (Join-Path $tw 'src\core\modules\modSha256.bas') 'y'
Check ((Test-VbaTwinFiles $tw).Count -eq 1) 'gemeos divergentes falham'
Remove-Item -LiteralPath $tw -Recurse -Force
# Registro do atualizador: nenhum fonte VBA (core, loader e testes) escreve o nome fixo "ThinQQ" no GetSetting/SaveSetting/DeleteSetting;
# todos passam por Settings_AppName (loader) / Util_UpdaterRegApp (core), que no build Test devolvem "ThinQQ.Test".
$regLit = @()
foreach ($d in 'src\core', 'src\loader', 'tests\vba\core', 'tests\vba\loader') {
    Get-ChildItem -LiteralPath (Join-Path $RepoRoot $d) -Recurse -File -Include '*.bas', '*.cls', '*.vba' | ForEach-Object {
        $f = $_.FullName
        Select-String -LiteralPath $f -Pattern '(GetSetting|SaveSetting|DeleteSetting)\s*\(?\s*"ThinQQ"' | ForEach-Object { $regLit += "$($f):$($_.LineNumber)" }
    }
}
$regLit | ForEach-Object { Write-Host "  nome fixo do registro: $_" }
Check ($regLit.Count -eq 0) 'registro do atualizador sem nome fixo "ThinQQ" (build Test usa ThinQQ.Test)'
# Nenhum fonte VBA (core, loader e testes) inicia processo filho (powershell/cmd/wscript..., Shell(), .Exec): o antivirus/EDR da maquina
# trata "Excel que salvou/vinculou uma pasta e depois roda powershell.exe" como documento malicioso e encerra o Excel (a suite completa do
# core morria com 0x800706BE no Test_Wfp_Connectors_SavedXml_HoldsErrBarsOnBlock). Linhas de comentario sao ignoradas.
# Shell() e VBA.Shell(); .Run com um executavel literal (.exe/.bat/.cmd/.ps1/.vbs/.js: WScript.Shell.Run); Application.Run de macro passa.
$spawnRx = '(?i)("\s*(powershell|pwsh|cmd|wscript|cscript|mshta|rundll32|regsvr32)(\.exe)?(\s|"))|((^|[^\w])(VBA\.)?Shell\s*\()|(\.Exec\s*[("])|(\.Run\s*\(?\s*"+[^"]*\.(exe|bat|cmd|ps1|vbs|js)\b)'
function Find-VbaSpawn([string[]]$Lines, [string]$Name) {
    $hits = @()
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $l = $Lines[$i].TrimStart()
        if ($l.StartsWith("'")) { continue }
        if ($l -match $spawnRx) { $hits += "$($Name):$($i + 1)" }
    }
    return , $hits
}
Check ((Find-VbaSpawn @('    sh.Run "powershell.exe -NoProfile -File " & q & p & q, 0, True') 'x.bas').Count -eq 1) 'pega processo filho (powershell)'
Check ((Find-VbaSpawn @('    x = Shell("cmd /c dir")', '    CreateObject("WScript.Shell").Exec("x")') 'y.bas').Count -eq 2) 'pega Shell() e .Exec'
Check ((Find-VbaSpawn @("    ' le o XML sem PowerShell", '    v = CreateObject("WScript.Shell").RegRead(k)', '    Application.Run "x"', '    Set sh = CreateObject("Shell.Application")') 'z.bas').Count -eq 0) 'ignora comentario, RegRead, Application.Run e Shell.Application'
Check ((Find-VbaSpawn @('    pid = VBA.Shell("notepad.exe " & f, vbHide)', '    VBA.Shell ("x")') 'v.bas').Count -eq 2) 'pega VBA.Shell('
Check ((Find-VbaSpawn @('    sh.Run "C:\ferramentas\atualiza.exe", 0, True', '    sh.Run("setup.bat")', '    o.Run "x.ps1 -a"', '    sh.Run """C:\a b\c.vbs"""') 'w.bas').Count -eq 4) 'pega .Run com executavel literal (.exe/.bat/.ps1/.vbs)'
Check ((Find-VbaSpawn @('    Application.Run "''" & wb.Name & "''!modX.Proc"', '    Application.Run "Pasta1.xlsm!Macro", 1', '    xl.Run("ThinQQ.core.xlam!Core_Refresh")') 'u.bas').Count -eq 0) 'ignora Application.Run de macro (sem executavel)'
$spawn = @()
foreach ($d in 'src\core', 'src\loader', 'src\ppt', 'tests\vba') {
    Get-ChildItem -LiteralPath (Join-Path $RepoRoot $d) -Recurse -File -Include '*.bas', '*.cls', '*.vba' | ForEach-Object {
        $spawn += Find-VbaSpawn @(Get-Content -LiteralPath $_.FullName -Encoding utf8) $_.FullName
    }
}
$spawn | ForEach-Object { Write-Host "  processo filho no VBA: $_" }
Check ($spawn.Count -eq 0) 'nenhum fonte VBA inicia processo filho (o EDR encerra o Excel da suite)'
$real = Test-VbaLint $RepoRoot
$real | ForEach-Object { Write-Host $_ }
Check ($real.Count -eq 0) 'fontes reais sem erros de continuacao de linha'
if ($script:fail -gt 0) { Write-Host "FALHAS vba-lint: $script:fail"; exit 1 }
