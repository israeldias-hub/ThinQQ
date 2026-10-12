# Privacidade dos arquivos versionados: o repositorio (e o GitHub Pages, que publica tudo sob docs/) nao pode carregar identificadores locais do
# autor: id de maquina/usuario, caminho de perfil do Windows com nome de usuario, dominio de e-mail corporativo, pasta de rascunho do Claude, PID de processo.
# Placeholders aceitos: C:\Users\<usuario>\..., %USERPROFILE%\..., <repo>.
# Os padroes abaixo sao montados em pedacos de proposito: assim este arquivo nao casa com os proprios padroes (a lista de excecoes sao os
# pedacos, nao um caminho de arquivo). Em um repositorio Git o conjunto de arquivos vem de `git ls-files -z` (versionados e novos nao
# ignorados), lido em UTF-8 explicito (nomes com acento chegam intactos no Windows PowerShell 5.1 e no pwsh 7); um caminho listado que nao
# existe no disco conta como FALHA (nunca e pulado em silencio). Numa copia sem .git (a copia de teste em C:\temp) varre a arvore, sem dist,
# tests\results, tests\fixtures\release, .git e .superpowers. Detalhes da maquina (build do Office e do Windows) tambem nao entram.
param([string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)))
$ErrorActionPreference = 'Stop'
$script:fail = 0
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m"; $script:fail++ } }

$bs = [string][char]92
# nome -> expressao regular (montadas por concatenacao)
$patterns = [ordered]@{
    'id de maquina/usuario'        = ('(?<!\d)14' + '3272(?!\d)')
    'caminho de perfil do Windows' = ('(?i)[a-z]:[\\/]+' + 'Us' + 'ers[\\/]+(?![<%$\{]|\.\.\.)[^\\/\s''"`<>]+')
    'e-mail corporativo'           = ('(?i)quero-?' + 'quero\.[a-z]+')
    'e-mail do autor'              = ('(?i)israel\.' + 'dias@')
    'pasta de rascunho do Claude'  = ('(?i)AppData[\\/]Local[\\/]Temp[\\/]' + 'claude')
    'PID de processo'              = ('(?i)\bPI' + 'Ds?\b[^a-z\n]{0,6}\d{3,}')
    'build do Office'              = ('\b16\.0\.' + '\d{4,}')
    'build do Windows'             = ('\b10\.0\.' + '\d{5}')
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
Check (@(Find-PrivacyHits ('c:' + $bs + $bs + 'Us' + 'ers' + $bs + $bs + 'ana' + $bs + $bs + 'x')).Count -eq 1) 'pega caminho de perfil com barras duplicadas (JSON)'
Check (@(Find-PrivacyHits ('c://Us' + 'ers//ana/x')).Count -eq 1) 'pega caminho de perfil com barras normais duplicadas'
Check (@(Find-PrivacyHits ('14' + '3272')).Count -eq 1) 'pega o id de maquina'
Check (@(Find-PrivacyHits ('u14' + '3272@x')).Count -eq 1) 'pega o id de maquina colado a letras'
Check (@(Find-PrivacyHits ('9' + '14' + '3272' + '1')).Count -eq 0) 'o id dentro de um numero maior nao conta'
Check (@(Find-PrivacyHits ('Excel 16.0.' + '19127')).Count -eq 1) 'pega o build do Office'
Check (@(Find-PrivacyHits ('Windows 10.0.' + '26200')).Count -eq 1) 'pega o build do Windows'
Check (@(Find-PrivacyHits 'Office 16.0 (2016 ou mais novo)').Count -eq 0) 'a versao 16.0 sem build passa'
Check (@(Find-PrivacyHits ('alguem@quero-' + 'quero.com.br')).Count -eq 1) 'pega o dominio corporativo (com hifen)'
Check (@(Find-PrivacyHits ('alguem@quero' + 'quero.local')).Count -eq 1) 'pega o dominio corporativo (sem hifen)'
Check (@(Find-PrivacyHits ('israel.' + 'dias@example.org')).Count -eq 1) 'pega o e-mail do autor'
Check (@(Find-PrivacyHits ('israeldias-hub/ThinQQ')).Count -eq 0) 'a conta publica do GitHub (israeldias-hub) e permitida'
Check (@(Find-PrivacyHits ('os PI' + 'Ds **1234, 56789**')).Count -eq 1) 'pega PIDs de processo'
Check (@(Find-PrivacyHits ('o PI' + 'D e o numero do processo')).Count -eq 0) 'a palavra PID sem numero passa'

# Arquivos versionados e novos nao ignorados de um repositorio Git: `git ls-files -z` (separados por NUL, sem aspas nem escapes) lido em UTF-8
# explicito por System.Diagnostics.Process (a saida de um comando nativo e decodificada pela pagina de codigo do console no Windows
# PowerShell 5.1: um nome com acento chegava trocado e o arquivo era pulado). $null se o git falhar. A raiz vai para o git sem a barra final
# (uma barra invertida antes da aspa de fechamento a escaparia e o git falharia); o stderr e lido em paralelo (um stderr grande nao trava o pipe).
function Get-GitFiles([string]$root) {
    $arg = $root.TrimEnd([char]92, [char]47)
    if ($arg -match '^[A-Za-z]:$') { $arg += '/' }                       # raiz de unidade: C:/
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = 'git'
    $psi.Arguments = '-C "' + $arg + '" -c core.quotepath=off ls-files -z --cached --others --exclude-standard'
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = New-Object Text.UTF8Encoding($false)
    $p = [Diagnostics.Process]::Start($psi)
    $errTask = $p.StandardError.ReadToEndAsync()
    $out = $p.StandardOutput.ReadToEnd()
    $p.WaitForExit()
    $null = $errTask.Result
    if ($p.ExitCode -ne 0) { return $null }
    return @($out.Split([char]0) | Where-Object { $_ } | ForEach-Object { Join-Path $root ($_ -replace '/', $bs) })
}

$utf8 = New-Object Text.UTF8Encoding($false)
$hasGitCmd = [bool](Get-Command git -ErrorAction SilentlyContinue)

# ---- autoteste da listagem: um arquivo de nome nao ASCII num repositorio Git temporario proprio (apagado no fim)
if ($hasGitCmd) {
    $tmp = Join-Path ([IO.Path]::GetTempPath()) ('thinqq_priv_' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        $null = & git -C $tmp init -q 2>&1
        $odd = 'relat' + [char]0xF3 + 'rio a' + [char]0xE7 + [char]0xE3 + 'o.txt'
        [IO.File]::WriteAllText((Join-Path $tmp $odd), ('c:/Us' + 'ers/maria/x'), $utf8)
        $listed = Get-GitFiles $tmp
        Check ($null -ne $listed -and @($listed | Where-Object { (Split-Path $_ -Leaf) -ceq $odd }).Count -eq 1) 'git ls-files -z em UTF-8 lista o arquivo de nome nao ASCII intacto'
        Check (@($listed | Where-Object { -not (Test-Path -LiteralPath $_ -PathType Leaf) }).Count -eq 0) 'todo caminho listado existe no disco'
        Check (@(Find-PrivacyHits ([IO.File]::ReadAllText((Join-Path $tmp $odd), $utf8))).Count -eq 1) 'o conteudo do arquivo de nome nao ASCII e varrido'
        $listed2 = Get-GitFiles ($tmp + $bs)
        Check ($null -ne $listed2 -and @($listed2 | Where-Object { (Split-Path $_ -Leaf) -ceq $odd }).Count -eq 1) 'raiz terminada em barra invertida: o git lista (nao cai na varredura da arvore)'
    } finally {
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }
} else {
    Write-Host 'skip - autoteste da listagem do git (git ausente)'
}

# ---- conjunto de arquivos
$files = @()
$hasGit = (Test-Path (Join-Path $RepoRoot '.git')) -and $hasGitCmd
$gitFailed = $false
if ($hasGit) {
    $files = Get-GitFiles $RepoRoot
    if ($null -eq $files) { $files = @(); $hasGit = $false; $gitFailed = $true }
}
if (-not $hasGit) {
    $skip = '\\(\.git|\.superpowers|dist|node_modules)\\|\\tests\\results\\|\\tests\\fixtures\\release\\'
    $files = @(Get-ChildItem -LiteralPath $RepoRoot -Recurse -File -Force | Where-Object {
        ($_.FullName.Substring($RepoRoot.Length) + $bs) -notmatch $skip -and $_.Name -notlike '*.tmp' -and $_.Name -notlike '~$*'
    } | ForEach-Object { $_.FullName })
}
Check ($files.Count -ge 50) "arquivos a verificar ($($files.Count)): $(if ($hasGit) { 'git ls-files' } elseif ($gitFailed) { 'varredura da arvore (o git ls-files FALHOU)' } else { 'varredura da arvore (sem .git)' })"

# ---- varredura (so texto: arquivo com byte NUL e binario); caminho listado que nao existe no disco e falha, nunca pulado
$bad = 0
foreach ($f in $files) {
    if (-not (Test-Path -LiteralPath $f -PathType Leaf)) {
        $bad++
        Check $false "arquivo listado nao encontrado no disco (nao varrido): $($f.Substring($RepoRoot.Length).TrimStart($bs, '/'))"
        continue
    }
    $bytes = [IO.File]::ReadAllBytes($f)
    if ([Array]::IndexOf($bytes, [byte]0) -ge 0) { continue }
    $hits = @(Find-PrivacyHits ($utf8.GetString($bytes)))
    if ($hits.Count -gt 0) {
        $bad++
        $rel = $f.Substring($RepoRoot.Length).TrimStart($bs, '/')
        Check $false "$rel : $($hits -join '; ')"
    }
}
Check ($bad -eq 0) 'nenhum arquivo versionado tem id de maquina, caminho de perfil, e-mail corporativo, pasta de rascunho do Claude, PID ou build da maquina'
if ($script:fail -gt 0) { Write-Host "FALHAS privacidade: $script:fail"; exit 1 }
