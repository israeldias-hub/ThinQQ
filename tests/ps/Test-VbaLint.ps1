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
$real = Test-VbaLint $RepoRoot
$real | ForEach-Object { Write-Host $_ }
Check ($real.Count -eq 0) 'fontes reais sem erros de continuacao de linha'
if ($script:fail -gt 0) { Write-Host "FALHAS vba-lint: $script:fail"; exit 1 }
