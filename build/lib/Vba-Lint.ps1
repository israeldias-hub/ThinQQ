# Lint estatico de fontes VBA: detecta continuacao de linha "_" seguida de texto (erro de sintaxe que o VBA so
# acusa ao compilar o procedimento, e que os testes nao veem quando o procedimento nunca e chamado).
function Get-VbaLintFiles([string]$Root) {
    $dirs = @('src', 'tests\vba')
    foreach ($d in $dirs) {
        $p = Join-Path $Root $d
        if (Test-Path -LiteralPath $p) {
            Get-ChildItem -LiteralPath $p -Recurse -File | Where-Object { $_.Extension -in '.bas', '.cls', '.vba' }
        }
    }
}

# Retorna lista de strings "arquivo:linha: mensagem". Tokenizador ciente de literais ("" como escape) e comentarios (').
function Test-VbaSourceLines([string[]]$Lines, [string]$Name) {
    $issues = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $ln = $Lines[$i]
        $inStr = $false
        for ($c = 0; $c -lt $ln.Length; $c++) {
            $ch = $ln[$c]
            if ($inStr) { if ($ch -eq '"') { $inStr = $false }; continue }
            if ($ch -eq '"') { $inStr = $true; continue }
            if ($ch -eq "'") { break }
            if ($ch -eq '_' -and $c -gt 0 -and ($ln[$c - 1] -eq ' ' -or $ln[$c - 1] -eq "`t")) {
                $rest = $ln.Substring($c + 1)
                if ($rest.Trim().Length -gt 0) {
                    $issues.Add("${Name}:$($i + 1): continuação de linha '_' seguida de texto na mesma linha")
                } else {
                    $next = if ($i + 1 -lt $Lines.Count) { $Lines[$i + 1] } else { '' }
                    if ($next.Trim().Length -eq 0) { $issues.Add("${Name}:$($i + 1): continuação de linha '_' seguida de linha em branco ou fim do arquivo") }
                }
                break
            }
        }
    }
    return , $issues
}

function Test-VbaLint([string]$Root) {
    $all = New-Object System.Collections.Generic.List[string]
    foreach ($f in (Get-VbaLintFiles $Root)) {
        $lines = [System.IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8) -split "\r?\n"
        $rel = $f.FullName.Substring($Root.Length).TrimStart('\')
        foreach ($m in (Test-VbaSourceLines $lines $rel)) { $all.Add($m) }
    }
    return , $all
}
