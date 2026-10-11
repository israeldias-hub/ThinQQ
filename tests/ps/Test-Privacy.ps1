# Privacidade dos arquivos versionados: o repositorio (e o GitHub Pages, que publica tudo sob docs/) nao pode carregar identificadores locais do
# autor: id de maquina/usuario, caminho de perfil do Windows com nome de usuario, dominio de e-mail corporativo, pasta de rascunho do Claude, PID de processo.
# Placeholders aceitos: C:\Users\<usuario>\..., %USERPROFILE%\..., <repo>.
# Os padroes abaixo sao montados em pedacos de proposito: assim este arquivo nao casa com os proprios padroes (a lista de excecoes sao os
# pedacos, nao um caminho de arquivo). Em um repositorio Git o conjunto de arquivos vem de `git ls-files` (versionados e novos nao ignorados);
# numa copia sem .git (a copia de teste em C:\temp) varre a arvore, sem dist, tests\results, tests\fixtures\release, .git e .superpowers.
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$ErrorActionPreference = 'Stop'
$script:fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }

$bs = [string][char]92
# nome -> expressao regular (montadas por concatenacao)
$patterns = [ordered]@{
    'id de maquina/usuario'        = ('14' + '3272')
    'caminho de perfil do Windows' = ('(?i)[a-z]:[\\/]' + 'Us' + 'ers[\\/](?![<%$\{]|\.\.\.)[^\\/\s''"`<>]+')
    'e-mail corporativo'           = ('(?i)quero-?' + 'quero\.[a-z]+')
    'e-mail do autor'              = ('(?i)israel\.' + 'dias@')
    'pasta de rascunho do Claude'  = ('(?i)AppData[\\/]Local[\\/]Temp[\\/]' + 'claude')
    'PID de processo'              = ('(?i)\bPI' + 'Ds?\b[^a-z\n]{0,6}\d{3,}')
}

# Linhas (1-based) de um texto que casam com algum padrao: devolve "linha: nome do padrao".
function Find-PrivacyHits([string]$text) {
    $hits = @()
    $n = 0
    foreach ($line in ($text -split "`n")) {
        $n++
        foreach ($k in $patterns.Keys) {
            if ($line -match $patterns[$k]) { $hits += "$n`: $k" }
        }
    }
    return $hits
}

# ---- o detector pega o que deve e deixa passar os placeholders
$userPath = 'C:' + $bs + 'Us' + 'ers' + $bs
Check (@(Find-PrivacyHits ($userPath + 'joao' + $bs + 'Desktop')).Count -eq 1) 'pega caminho de perfil com nome de usuario'
Check (@(Find-PrivacyHits ('c:/Us' + 'ers/maria/x')).Count -eq 1) 'pega caminho de perfil com barra normal'
Check (@(Find-PrivacyHits ($userPath + '<usuario>' + $bs + 'x')).Count -eq 0) 'aceita o placeholder <usuario>'
Check (@(Find-PrivacyHits ($userPath + '<usu' + [char]0xE1 + 'rio>' + $bs + 'x')).Count -eq 0) 'aceita o placeholder com acento'
Check (@(Find-PrivacyHits ('%USERPROFILE%' + $bs + 'Desktop')).Count -eq 0) 'aceita %USERPROFILE%'
Check (@(Find-PrivacyHits ('<repo>' + $bs + 'src')).Count -eq 0) 'aceita <repo>'
Check (@(Find-PrivacyHits ('14' + '3272')).Count -eq 1) 'pega o id de maquina'
Check (@(Find-PrivacyHits ('alguem@quero-' + 'quero.com.br')).Count -eq 1) 'pega o dominio corporativo (com hifen)'
Check (@(Find-PrivacyHits ('alguem@quero' + 'quero.local')).Count -eq 1) 'pega o dominio corporativo (sem hifen)'
Check (@(Find-PrivacyHits ('israel.' + 'dias@example.org')).Count -eq 1) 'pega o e-mail do autor'
Check (@(Find-PrivacyHits ('israeldias-hub/ThinQQ')).Count -eq 0) 'a conta publica do GitHub (israeldias-hub) e permitida'
Check (@(Find-PrivacyHits ('os PI' + 'Ds **1234, 56789**')).Count -eq 1) 'pega PIDs de processo'
Check (@(Find-PrivacyHits ('o PI' + 'D e o numero do processo')).Count -eq 0) 'a palavra PID sem numero passa'

# ---- conjunto de arquivos
$files = @()
$hasGit = (Test-Path (Join-Path $RepoRoot '.git')) -and (Get-Command git -ErrorAction SilentlyContinue)
if ($hasGit) {
    $out = & git -C $RepoRoot -c core.quotepath=off ls-files --cached --others --exclude-standard
    if ($LASTEXITCODE -eq 0) { $files = @($out | Where-Object { $_ } | ForEach-Object { Join-Path $RepoRoot ($_ -replace '/', $bs) }) } else { $hasGit = $false }
}
if (-not $hasGit) {
    $skip = '\\(\.git|\.superpowers|dist|node_modules)\\|\\tests\\results\\|\\tests\\fixtures\\release\\'
    $files = @(Get-ChildItem -LiteralPath $RepoRoot -Recurse -File -Force | Where-Object {
        ($_.FullName.Substring($RepoRoot.Length) + $bs) -notmatch $skip -and $_.Name -notlike '*.tmp' -and $_.Name -notlike '~$*'
    } | ForEach-Object { $_.FullName })
}
Check ($files.Count -ge 50) "arquivos a verificar ($($files.Count)): $(if ($hasGit) { 'git ls-files' } else { 'varredura da arvore (sem .git)' })"

# ---- varredura (so texto: arquivo com byte NUL e binario)
$utf8 = New-Object Text.UTF8Encoding($false)
$bad = 0
foreach ($f in $files) {
    if (-not (Test-Path -LiteralPath $f -PathType Leaf)) { continue }
    $bytes = [IO.File]::ReadAllBytes($f)
    if ([Array]::IndexOf($bytes, [byte]0) -ge 0) { continue }
    $hits = @(Find-PrivacyHits ($utf8.GetString($bytes)))
    if ($hits.Count -gt 0) {
        $bad++
        $rel = $f.Substring($RepoRoot.Length).TrimStart($bs, '/')
        Check $false "$rel : $($hits -join '; ')"
    }
}
Check ($bad -eq 0) 'nenhum arquivo versionado tem id de maquina, caminho de perfil, e-mail corporativo ou pasta de rascunho do Claude'
if ($script:fail -gt 0) { Write-Host "FALHAS privacidade: $script:fail"; exit 1 }
