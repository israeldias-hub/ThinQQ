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
$real = Test-VbaLint $RepoRoot
$real | ForEach-Object { Write-Host $_ }
Check ($real.Count -eq 0) 'fontes reais sem erros de continuacao de linha'
if ($script:fail -gt 0) { Write-Host "FALHAS vba-lint: $script:fail"; exit 1 }
