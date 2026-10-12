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
        foreach ($m in (Test-VbaCp1252Lines $lines $rel)) { $all.Add($m) }
    }
    foreach ($m in (Test-VbaTwinFiles $Root)) { $all.Add($m) }
    return , $all
}

# M13: o build converte os fontes VBA para cp1252; caracteres fora dele viram '?' em silencio (ex.: setas, aspas
# tipograficas fora do cp1252, emojis). Retorna "arquivo:linha: ..." para cada linha com tal caractere.
function Test-VbaCp1252Lines([string[]]$Lines, [string]$Name) {
    $issues = New-Object System.Collections.Generic.List[string]
    $enc = [System.Text.Encoding]::GetEncoding(1252)
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $ln = $Lines[$i]
        for ($c = 0; $c -lt $ln.Length; $c++) {
            $ch = $ln[$c]
            if ([int]$ch -lt 128) { continue }
            $back = $enc.GetString($enc.GetBytes([string]$ch))
            if ($back -ne [string]$ch) {
                $issues.Add(("{0}:{1}: caractere fora do cp1252 (U+{2:X4}) viraria '?' no build" -f $Name, ($i + 1), [int]$ch))
                break
            }
        }
    }
    return , $issues
}

# Pares de arquivos que precisam ser identicos (o core nao pode depender do loader, entao SHA-256 e seus testes existem nos dois).
function Test-VbaTwinFiles([string]$Root) {
    $issues = New-Object System.Collections.Generic.List[string]
    $pairs = @(
        @('src\loader\modules\modSha256.bas', 'src\core\modules\modSha256.bas'),
        @('tests\vba\loader\TestSha256.bas', 'tests\vba\core\TestSha256.bas'))
    foreach ($p in $pairs) {
        $a = Join-Path $Root $p[0]; $b = Join-Path $Root $p[1]
        if (-not (Test-Path -LiteralPath $a) -or -not (Test-Path -LiteralPath $b)) { $issues.Add("$($p[0]) / $($p[1]): arquivo gemeo ausente"); continue }
        $ta = ([System.IO.File]::ReadAllText($a, [System.Text.Encoding]::UTF8) -replace "\r\n", "`n").TrimEnd()
        $tb = ([System.IO.File]::ReadAllText($b, [System.Text.Encoding]::UTF8) -replace "\r\n", "`n").TrimEnd()
        if ($ta -ne $tb) { $issues.Add("$($p[0]) e $($p[1]) divergem (devem ser identicos)") }
    }
    return , $issues
}
