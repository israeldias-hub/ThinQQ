# Funcoes do publicar-release (tools/Publish-Release.ps1, carregadas com dot-source): versao (semver), CHANGELOG, version.json, estado da
# preparacao (state.json), git, API do GitHub com transporte substituivel e as duas fases (Prepare e Publish).
# Compativel com Windows PowerShell 5.1 e pwsh 7. Texto ASCII de proposito (o 5.1 le .ps1 sem BOM como ANSI); os acentos do CHANGELOG
# sao montados com [char].
#
# SEGURANCA DO TOKEN: o token do GitHub vem de `git credential fill` (stdin, nunca linha de comando) e fica so na memoria do processo, dentro
# de um PSCredential (SecureString). Ele so vira texto dentro de Invoke-GitHubApi, no cabecalho Authorization de uma chamada, e e removido
# ao fim dela. Nunca e gravado (state.json, log, arquivo temporario), nunca e impresso e toda mensagem de erro passa por Protect-RpSecret.
# Chamadas so para https://api.github.com e https://uploads.github.com (redirecionamentos sao recusados, nao seguidos).

$script:RpUtf8 = New-Object Text.UTF8Encoding($false)
$script:RpUnreleased = '## [N' + [char]0xE3 + 'o lan' + [char]0xE7 + 'ado]'
$script:RpAllowedHosts = @('api.github.com', 'uploads.github.com')
$script:RpApiVersion = '2022-11-28'
$script:RpUserAgent = 'ThinQQ-Publish-Release'
$script:RpTestAllowLoopback = $false
# Arquivos que a etapa Preparar grava e commita (conferidos num merge por squash/rebase)
$script:RpVersionFiles = @('version.json', 'CHANGELOG.md', 'release/manifest.json', 'release/notes.txt', 'docs/tutorial/tutorial.html', 'installer/LEIA-ME.txt')

# ---------------------------------------------------------------------------------------------------------------------------------------
# Erros: Kind 'precondicao' (o script sai com 2: nada foi alterado) ou 'falha' (sai com 1). HttpStatus quando veio da API.
function New-RpError([string]$Message, [string]$Kind = 'falha', [int]$HttpStatus = 0) {
    $ex = New-Object System.Exception($Message)
    $ex.Data['RpKind'] = $Kind
    if ($HttpStatus -gt 0) { $ex.Data['HttpStatus'] = $HttpStatus }
    return $ex
}

function Get-RpErrorKind($ErrorRecord) {
    $ex = $ErrorRecord
    if ($ErrorRecord -is [Management.Automation.ErrorRecord]) { $ex = $ErrorRecord.Exception }
    if ($ex -and $ex.Data.Contains('RpKind')) { return [string]$ex.Data['RpKind'] }
    return 'falha'
}

function Get-RpErrorStatus($ErrorRecord) {
    $ex = $ErrorRecord
    if ($ErrorRecord -is [Management.Automation.ErrorRecord]) { $ex = $ErrorRecord.Exception }
    if ($ex -and $ex.Data.Contains('HttpStatus')) { return [int]$ex.Data['HttpStatus'] }
    return 0
}

# Troca o segredo (e qualquer coisa com cara de token) por *** num texto que vai para a tela ou para uma excecao.
function Protect-RpSecret([string]$Text, [string]$Secret = '') {
    if ($null -eq $Text) { return '' }
    $t = $Text
    if ($Secret) { $t = $t.Replace($Secret, '***') }
    $t = [regex]::Replace($t, '(?i)(bearer)\s+\S+', '$1 ***')
    $t = [regex]::Replace($t, '(?i)(basic)\s+[A-Za-z0-9+/=]{12,}', '$1 ***')
    $t = [regex]::Replace($t, '\bgh[pousr]_[A-Za-z0-9]{16,}\b', '***')
    $t = [regex]::Replace($t, '\bgithub_pat_[A-Za-z0-9_]{16,}\b', '***')
    return $t
}

# ---------------------------------------------------------------------------------------------------------------------------------------
# Versao (semver 2.0: x.y.z com pre-lancamento opcional, sem metadados de build)
function Test-SemVer([string]$Version) {
    if (-not $Version) { return $false }
    return [bool]($Version -cmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)(-(0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*)(\.(0|[1-9]\d*|\d*[A-Za-z-][0-9A-Za-z-]*))*)?$')
}

# So x.y.z (o build/release.ps1 e o version.json usam esse formato)
function Test-PlainVersion([string]$Version) {
    if (-not $Version) { return $false }
    return [bool]($Version -cmatch '^(0|[1-9]\d*)\.(0|[1-9]\d*)\.(0|[1-9]\d*)$')
}

# -1 se A < B, 0 se iguais, 1 se A > B (precedencia do semver: pre-lancamento < lancamento; identificadores numericos < alfanumericos)
function Compare-SemVer([string]$A, [string]$B) {
    if (-not (Test-SemVer $A)) { throw "Versao invalida: '$A'" }
    if (-not (Test-SemVer $B)) { throw "Versao invalida: '$B'" }
    $pa = @($A -split '-', 2); $pb = @($B -split '-', 2)
    $na = $pa[0].Split('.'); $nb = $pb[0].Split('.')
    for ($i = 0; $i -lt 3; $i++) {
        $x = [decimal]$na[$i]; $y = [decimal]$nb[$i]
        if ($x -lt $y) { return -1 }
        if ($x -gt $y) { return 1 }
    }
    $hasA = $pa.Count -gt 1; $hasB = $pb.Count -gt 1
    if (-not $hasA -and -not $hasB) { return 0 }
    if (-not $hasA) { return 1 }
    if (-not $hasB) { return -1 }
    $ia = $pa[1].Split('.'); $ib = $pb[1].Split('.')
    $n = [Math]::Min($ia.Count, $ib.Count)
    for ($i = 0; $i -lt $n; $i++) {
        $u = $ia[$i]; $v = $ib[$i]
        $uNum = $u -cmatch '^\d+$'; $vNum = $v -cmatch '^\d+$'
        if ($uNum -and $vNum) {
            $x = [decimal]$u; $y = [decimal]$v
            if ($x -lt $y) { return -1 }
            if ($x -gt $y) { return 1 }
        } elseif ($uNum) { return -1 }
        elseif ($vNum) { return 1 }
        else {
            $c = [string]::CompareOrdinal($u, $v)
            if ($c -lt 0) { return -1 }
            if ($c -gt 0) { return 1 }
        }
    }
    if ($ia.Count -lt $ib.Count) { return -1 }
    if ($ia.Count -gt $ib.Count) { return 1 }
    return 0
}

# ---------------------------------------------------------------------------------------------------------------------------------------
# Texto: le/grava UTF-8 sem BOM e preserva o fim de linha que o arquivo ja usa
function Read-RpText([string]$Path) { return [IO.File]::ReadAllText($Path, $script:RpUtf8) }
function Write-RpText([string]$Path, [string]$Text) { [IO.File]::WriteAllText($Path, $Text, $script:RpUtf8) }
function Get-TextNewline([string]$Text) { if ($Text -and $Text.Contains("`r`n")) { return "`r`n" } return "`n" }

# CRLF -> LF e sem espacos nas pontas (compara notas e manifestos sem depender do core.autocrlf)
function Get-NormalizedText([string]$Text) {
    if ($null -eq $Text) { return '' }
    $t = $Text
    if ($t.Length -gt 0 -and $t[0] -eq [char]0xFEFF) { $t = $t.Substring(1) }
    return $t.Replace("`r`n", "`n").Trim()
}

function Get-BytesSha256([byte[]]$Bytes) {
    $sha = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($Bytes)).Replace('-', '').ToLowerInvariant()) } finally { $sha.Dispose() }
}

function Get-TextSha256([string]$Text) { return (Get-BytesSha256 ($script:RpUtf8.GetBytes((Get-NormalizedText $Text)))) }

# .NET direto (sem Get-FileHash: num Windows PowerShell 5.1 aberto a partir do pwsh 7, o PSModulePath herdado pode esconder o cmdlet)
function Get-FileSha256([string]$Path) {
    $sha = [Security.Cryptography.SHA256]::Create()
    $fs = [IO.File]::OpenRead($Path)
    try { return ([BitConverter]::ToString($sha.ComputeHash($fs)).Replace('-', '').ToLowerInvariant()) } finally { $fs.Dispose(); $sha.Dispose() }
}

# ---------------------------------------------------------------------------------------------------------------------------------------
# CHANGELOG: entradas ("- ...") da secao [Nao lancado]; $null se a secao nao existe
function Get-UnreleasedEntries([string]$Text) {
    $lines = $Text -split "\r?\n"
    $start = -1
    for ($i = 0; $i -lt $lines.Count; $i++) { if ($lines[$i].TrimEnd() -ceq $script:RpUnreleased) { $start = $i; break } }
    if ($start -lt 0) { return $null }
    $entries = New-Object System.Collections.Generic.List[string]
    for ($i = $start + 1; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^## \[') { break }
        if ($lines[$i] -match '^\s*[-*]\s+\S') { $entries.Add($lines[$i].Trim()) }
    }
    return ,$entries.ToArray()
}

# A secao [Nao lancado] vira "## [x.y.z] - AAAA-MM-DD" (com as entradas dela) e uma secao [Nao lancado] vazia nova fica no topo.
function Update-ChangelogText([string]$Text, [string]$Version, [string]$Date) {
    if (-not (Test-SemVer $Version)) { throw "Versao invalida: '$Version'" }
    if ($Date -notmatch '^\d{4}-\d{2}-\d{2}$') { throw "Data invalida: '$Date' (use AAAA-MM-DD)" }
    $m = [regex]::Match($Text, '(?m)^' + [regex]::Escape($script:RpUnreleased) + '[ \t]*(?=\r?$)')
    if (-not $m.Success) { throw 'O CHANGELOG.md nao tem a secao "' + $script:RpUnreleased + '".' }
    if ($Text -match ('(?m)^## \[' + [regex]::Escape($Version) + '\]')) { throw "O CHANGELOG.md ja tem a secao [$Version]." }
    $nl = Get-TextNewline $Text
    $insert = $script:RpUnreleased + $nl + $nl + "## [$Version] - $Date"
    return $Text.Substring(0, $m.Index) + $insert + $Text.Substring($m.Index + $m.Length)
}

# version.json: loader e core recebem a versao nova; o formato do arquivo (uma linha) e mantido
function Update-VersionJsonText([string]$Text, [string]$Version) {
    if (-not (Test-PlainVersion $Version)) { throw "Versao invalida para o version.json: '$Version' (use x.y.z)" }
    $obj = $Text | ConvertFrom-Json
    foreach ($k in 'loader', 'core') {
        if (-not ($obj.PSObject.Properties.Name -contains $k)) { throw "version.json sem o campo '$k'." }
    }
    $t = $Text
    foreach ($k in 'loader', 'core') {
        $t = [regex]::Replace($t, '("' + $k + '"\s*:\s*")[^"]*(")', ('${1}' + $Version + '${2}'))
    }
    $back = $t | ConvertFrom-Json
    if ($back.loader -ne $Version -or $back.core -ne $Version) { throw 'Nao foi possivel atualizar o version.json.' }
    return $t
}

# Rodape do tutorial ("<p>ThinQQ x.y.z. ...") e o exemplo de pasta do LEIA-ME ("versions\x.y.z\"): conferidos pelo Test-Docs contra o
# version.json. Devolvem o texto igual quando nao acham o padrao (quem chama avisa).
function Update-TutorialVersionText([string]$Text, [string]$Version) {
    return [regex]::Replace($Text, '(<p>ThinQQ )\d+\.\d+\.\d+(\. )', ('${1}' + $Version + '${2}'))
}

function Update-LeiaMeVersionText([string]$Text, [string]$Version) {
    return [regex]::Replace($Text, '(versions\\)\d+\.\d+\.\d+(\\)', ('${1}' + $Version + '${2}'))
}

# Corpo da release: o texto de release/notes.txt com LF e sem espacos nas pontas
function New-ReleaseBody([string]$NotesText) {
    $b = Get-NormalizedText $NotesText
    if (-not $b) { throw 'release/notes.txt esta vazio: escreva as notas da versao.' }
    return $b
}

# ---------------------------------------------------------------------------------------------------------------------------------------
# Assets
function Get-ReleaseAssetNames([string]$Version) {
    return @("ThinQQ-$Version.zip", 'ThinQQ.xlam', 'ThinQQ.core.xlam', 'manifest.json', 'tutorial.html')
}

function Get-AssetContentType([string]$Name) {
    switch -regex ($Name.ToLowerInvariant()) {
        '\.zip$' { return 'application/zip' }
        '\.xlam$' { return 'application/vnd.ms-excel.addin.macroEnabled.12' }
        '\.json$' { return 'application/json' }
        '\.html?$' { return 'text/html' }
        default { return 'application/octet-stream' }
    }
}

function Get-CompareUrl([string]$Repo, [string]$Version) { return "https://github.com/$Repo/compare/main...release/v$Version" + '?expand=1' }

function Get-ReleaseBaseUrl([string]$Repo, [string]$Version) { return "https://github.com/$Repo/releases/download/v$Version" }

function Test-RepoName([string]$Repo) { return [bool]($Repo -match '^[A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+$') }

# ---------------------------------------------------------------------------------------------------------------------------------------
# state.json: versao, SHA-256 e tamanho dos 5 arquivos, commit de preparacao e a main de origem. Nunca guarda segredo.
function New-ReleaseState([string]$Version, [string]$Repo, [string]$StateDir, [string]$PrepareCommit, [string]$SourceMain, [string]$CreatedAt = '') {
    if (-not $CreatedAt) { $CreatedAt = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ssK') }
    $files = @()
    foreach ($n in (Get-ReleaseAssetNames $Version)) {
        $p = Join-Path $StateDir $n
        if (-not (Test-Path -LiteralPath $p)) { throw "Arquivo ausente para o state.json: $p" }
        $files += [ordered]@{ name = $n; size = (Get-Item -LiteralPath $p).Length; sha256 = (Get-FileSha256 $p) }
    }
    return [ordered]@{
        schemaVersion = 1
        version       = $Version
        repo          = $Repo
        prepareCommit = $PrepareCommit
        sourceMain    = $SourceMain
        createdAt     = $CreatedAt
        files         = $files
    }
}

function Write-ReleaseState($State, [string]$StateDir) {
    Write-RpText (Join-Path $StateDir 'state.json') ((ConvertTo-Json -InputObject $State -Depth 6) + "`n")
}

# Le e confere o state.json: a versao, o commit e os 5 arquivos (tamanho e SHA-256). Recusa (precondicao) com a lista do que nao confere.
function Read-ReleaseState([string]$StateDir, [string]$Version) {
    $path = Join-Path $StateDir 'state.json'
    if (-not (Test-Path -LiteralPath $path)) {
        throw (New-RpError "Nao encontrei $path. Rode antes a etapa Preparar (publicar-release.bat $Version) nesta maquina." 'precondicao')
    }
    try { $st = (Read-RpText $path) | ConvertFrom-Json } catch { throw (New-RpError "state.json invalido em ${path}: $($_.Exception.Message)" 'precondicao') }
    if ($st.version -cne $Version) { throw (New-RpError "O state.json e da versao '$($st.version)', nao de $Version." 'precondicao') }
    if ([string]$st.prepareCommit -notmatch '^[0-9a-f]{40}$') { throw (New-RpError 'state.json sem o commit de preparacao (prepareCommit).' 'precondicao') }
    $problems = New-Object System.Collections.Generic.List[string]
    $byName = @{}
    foreach ($f in @($st.files)) { if ($f -and $f.name) { $byName[[string]$f.name] = $f } }
    foreach ($n in (Get-ReleaseAssetNames $Version)) {
        $p = Join-Path $StateDir $n
        if (-not $byName.ContainsKey($n)) { $problems.Add("$n nao esta no state.json"); continue }
        if (-not (Test-Path -LiteralPath $p)) { $problems.Add("$n ausente em $StateDir"); continue }
        $len = (Get-Item -LiteralPath $p).Length
        if ($len -ne [long]$byName[$n].size) { $problems.Add("$n tem $len bytes, o state.json diz $($byName[$n].size)") ; continue }
        $sha = Get-FileSha256 $p
        if ($sha -cne [string]$byName[$n].sha256) { $problems.Add("$n mudou depois da preparacao (SHA-256 $sha, o state.json diz $($byName[$n].sha256))") }
    }
    if ($problems.Count -gt 0) {
        throw (New-RpError ("Os arquivos da preparacao nao conferem com o state.json: " + ($problems -join '; ') + '. Rode a etapa Preparar de novo.') 'precondicao')
    }
    return $st
}

# ---------------------------------------------------------------------------------------------------------------------------------------
# Processos (git, pwsh): System.Diagnostics.Process, saida em UTF-8, stderr lido em paralelo, stdin opcional (o token nunca vai em argumento)
function ConvertTo-RpArgument([string]$Arg) {
    if ($null -eq $Arg -or $Arg.Length -eq 0) { return '""' }
    if ($Arg -notmatch '[\s"]') { return $Arg }
    $sb = New-Object Text.StringBuilder
    [void]$sb.Append('"')
    $bs = 0
    foreach ($ch in $Arg.ToCharArray()) {
        if ($ch -eq [char]92) { $bs++; continue }
        if ($ch -eq '"') { [void]$sb.Append(([string][char]92) * ($bs * 2 + 1)); [void]$sb.Append('"'); $bs = 0; continue }
        if ($bs -gt 0) { [void]$sb.Append(([string][char]92) * $bs); $bs = 0 }
        [void]$sb.Append($ch)
    }
    if ($bs -gt 0) { [void]$sb.Append(([string][char]92) * ($bs * 2)) }
    [void]$sb.Append('"')
    return $sb.ToString()
}

function Invoke-RpProcess {
    param([string]$FilePath, [string[]]$Arguments = @(), [string]$WorkingDirectory = '', [string]$InputText = $null, [hashtable]$Environment = @{})
    $psi = New-Object Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = (@($Arguments | ForEach-Object { ConvertTo-RpArgument $_ }) -join ' ')
    if ($WorkingDirectory) { $psi.WorkingDirectory = $WorkingDirectory }
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardInput = $true
    $psi.StandardOutputEncoding = $script:RpUtf8
    $psi.StandardErrorEncoding = $script:RpUtf8
    foreach ($k in $Environment.Keys) { $psi.EnvironmentVariables[[string]$k] = [string]$Environment[$k] }
    # O Process cria o escritor do stdin com a codificacao de entrada do console e, com AutoFlush, grava logo o preambulo: num console UTF-8
    # (chcp 65001, terminal pwsh) o processo filho receberia um BOM antes de "protocol=" e o git recusaria a credencial. pwsh 7: codificacao
    # explicita sem BOM; Windows PowerShell 5.1 (sem essa propriedade): troca a codificacao de entrada do console so durante o Start.
    $restoreEnc = $null
    if ($psi.PSObject.Properties.Name -contains 'StandardInputEncoding') { $psi.StandardInputEncoding = $script:RpUtf8 }
    else {
        try {
            $cur = [Console]::InputEncoding
            if ($cur.GetPreamble().Length -gt 0) { $restoreEnc = $cur; [Console]::InputEncoding = $script:RpUtf8 }
        } catch { $restoreEnc = $null }
    }
    try { $p = [Diagnostics.Process]::Start($psi) }
    finally { if ($restoreEnc) { try { [Console]::InputEncoding = $restoreEnc } catch { } } }
    try {
        $errTask = $p.StandardError.ReadToEndAsync()
        # bytes UTF-8 SEM BOM direto no stream: o StreamWriter do stdin pode gravar um BOM (Console.InputEncoding UTF-8 no Windows
        # PowerShell 5.1), e o git leria a primeira chave como "<BOM>protocol"
        if ($InputText) {
            $inBytes = $script:RpUtf8.GetBytes($InputText)
            $p.StandardInput.BaseStream.Write($inBytes, 0, $inBytes.Length)
            $p.StandardInput.BaseStream.Flush()
            [Array]::Clear($inBytes, 0, $inBytes.Length)
        }
        $p.StandardInput.Close()
        $out = $p.StandardOutput.ReadToEnd()
        $p.WaitForExit()
        return [pscustomobject]@{ Code = $p.ExitCode; Output = $out; Error = $errTask.Result }
    } finally { $p.Dispose() }
}

$script:RpGitExe = $null
function Get-RpGitExe {
    if (-not $script:RpGitExe) {
        $c = Get-Command git -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $c) { throw (New-RpError 'O Git nao foi encontrado. Instale o Git for Windows e tente de novo.' 'precondicao') }
        $script:RpGitExe = $c.Source
    }
    return $script:RpGitExe
}

function Invoke-RpGit([string]$RepoRoot, [string[]]$Arguments, [string]$InputText = $null, [hashtable]$Environment = @{}) {
    return (Invoke-RpProcess (Get-RpGitExe) (@('-C', $RepoRoot) + $Arguments) -InputText $InputText -Environment $Environment)
}

# git que precisa dar certo: devolve a saida sem espacos nas pontas ou lanca erro com o stderr
function Invoke-RpGitChecked([string]$RepoRoot, [string[]]$Arguments, [string]$What, [string]$Kind = 'falha') {
    $r = Invoke-RpGit $RepoRoot $Arguments
    if ($r.Code -ne 0) {
        $detail = (($r.Error + ' ' + $r.Output).Trim())
        throw (New-RpError ("$What falhou (git $($Arguments -join ' '), codigo $($r.Code)): " + (Protect-RpSecret $detail)) $Kind)
    }
    return $r.Output.Trim()
}

function Get-RpLastTag([string]$RepoRoot) {
    $r = Invoke-RpGit $RepoRoot @('describe', '--tags', '--abbrev=0', '--match', 'v*', 'HEAD')
    if ($r.Code -ne 0) { return $null }
    return $r.Output.Trim()
}

# Caminho relativo ao repositorio com '/', para `git show <ref>:<caminho>`
function Get-RpRepoRelativePath([string]$RepoRoot, [string]$Path) {
    $full = [IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($Path)) { $Path } else { Join-Path $RepoRoot $Path }))
    $root = [IO.Path]::GetFullPath($RepoRoot).TrimEnd([char]92, [char]47) + [char]92
    if (-not $full.StartsWith($root, [StringComparison]::OrdinalIgnoreCase)) { return $null }
    return $full.Substring($root.Length).Replace([char]92, [char]47)
}

# ---------------------------------------------------------------------------------------------------------------------------------------
# Login do GitHub pelo Git (Git Credential Manager): `git credential fill` com protocol/host/username no stdin, sem prompt (falha em vez de
# travar). Devolve um PSCredential (a senha/token numa SecureString); erros citam so as CHAVES recebidas, nunca os valores.
function Get-RepoCredentialUser([string]$RemoteUrl, [string]$Repo) {
    $m = [regex]::Match([string]$RemoteUrl, '^https://([^@/:]+)(:[^@/]*)?@github\.com/', 'IgnoreCase')
    if ($m.Success) { return $m.Groups[1].Value }
    return $Repo.Split('/')[0]
}

function Get-GitCredentialRequest([string]$UserName) {
    return "protocol=https`nhost=github.com`nusername=$UserName`n`n"
}

function ConvertFrom-GitCredentialOutput([string]$Output) {
    $user = $null; $pass = $null; $keys = New-Object System.Collections.Generic.List[string]
    foreach ($line in ([string]$Output -split "\r?\n")) {
        $i = $line.IndexOf('=')
        if ($i -le 0) { continue }
        $k = $line.Substring(0, $i); $v = $line.Substring($i + 1)
        if (-not $keys.Contains($k)) { $keys.Add($k) }
        if ($k -ceq 'username') { $user = $v }
        elseif ($k -ceq 'password') { $pass = $v }
    }
    if (-not $pass) {
        throw (New-RpError ("O git credential fill nao devolveu um token do GitHub (chaves recebidas: " + $(if ($keys.Count) { $keys -join ', ' } else { 'nenhuma' }) + ').') 'precondicao')
    }
    $sec = New-Object Security.SecureString
    foreach ($ch in $pass.ToCharArray()) { $sec.AppendChar($ch) }
    $sec.MakeReadOnly()
    $pass = $null
    if (-not $user) { $user = 'x-access-token' }
    return (New-Object Management.Automation.PSCredential($user, $sec))
}

# Runner (substituivel pelo teste): recebe o texto do stdin e devolve @{ Code; Output; Error }
function Get-GitHubCredential([string]$RepoRoot, [string]$Repo, [scriptblock]$Runner = $null) {
    $remote = ''
    $r = Invoke-RpGit $RepoRoot @('remote', 'get-url', 'origin')
    if ($r.Code -eq 0) { $remote = $r.Output.Trim() }
    $request = Get-GitCredentialRequest (Get-RepoCredentialUser $remote $Repo)
    if ($Runner) { $res = & $Runner $request }
    else { $res = Invoke-RpGit $RepoRoot @('credential', 'fill') -InputText $request -Environment @{ GIT_TERMINAL_PROMPT = '0'; GCM_INTERACTIVE = 'never' } }
    try {
        if ($res.Code -ne 0) {
            throw (New-RpError ("Nao foi possivel obter o login do GitHub pelo Git (git credential fill, codigo $($res.Code)). Faca login uma vez (por exemplo, " +
                "um git fetch neste repositorio, que abre a janela do Git Credential Manager) e rode de novo. " + (Protect-RpSecret ([string]$res.Error).Trim())) 'precondicao')
        }
        return (ConvertFrom-GitCredentialOutput ([string]$res.Output))
    } finally { $res = $null }
}

# Registros de erro: o Invoke-WebRequest do pwsh 7 poe a requisicao (com o cabecalho Authorization) dentro da excecao e do TargetObject do
# ErrorRecord, e o PowerShell guarda em $Error ate os erros tratados num catch. Get-RpErrorMark tira uma foto de $Error e
# Remove-RpNewErrors apaga tudo o que entrou depois dela (comparacao por referencia).
function Get-RpErrorMark {
    $set = New-Object 'System.Collections.Generic.HashSet[object]'
    foreach ($e in @($global:Error)) { [void]$set.Add($e) }
    return ,$set
}

function Remove-RpNewErrors($Mark) {
    foreach ($e in @($global:Error)) { if (-not $Mark.Contains($e)) { $global:Error.Remove($e) } }
}

# ---------------------------------------------------------------------------------------------------------------------------------------
# API do GitHub. O transporte e um scriptblock param($Request) -> @{ Status = int; Content = byte[] ou string }, onde $Request e
# @{ Method; Uri; Headers; Body (byte[] ou $null); ContentType; UserAgent }. O padrao usa Invoke-WebRequest sem seguir redirecionamentos.
# Status HTTP de erro NUNCA viram excecao: pwsh 7 usa -SkipHttpErrorCheck; no 5.1 a WebException e lida e descartada. Uma falha de rede vira
# uma Exception NOVA so com a mensagem (sem o objeto da requisicao), e os registros que o Invoke-WebRequest deixou em $Error sao apagados.
$script:RpDefaultTransport = {
    param($Request)
    $isCore = $PSVersionTable.PSVersion.Major -ge 6
    if (-not $isCore) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    }
    $p = @{ Method = $Request.Method; Uri = $Request.Uri; Headers = $Request.Headers; UseBasicParsing = $true; MaximumRedirection = 0
        TimeoutSec = 600; UserAgent = $Request.UserAgent; ErrorAction = 'Stop' }
    if ($isCore) { $p.SkipHttpErrorCheck = $true }
    if ($null -ne $Request.Body) { $p.Body = [byte[]]$Request.Body }
    if ($Request.ContentType) { $p.ContentType = $Request.ContentType }
    $secret = ''
    if ($Request.Headers -and $Request.Headers.ContainsKey('Authorization') -and ([string]$Request.Headers['Authorization']) -match '^\S+\s+(.+)$') { $secret = $Matches[1] }
    $mark = Get-RpErrorMark
    $failure = $null
    try {
        try {
            $r = Invoke-WebRequest @p
            return @{ Status = [int]$r.StatusCode; Content = $r.RawContentStream.ToArray() }
        } catch {
            $resp = $null
            if (-not $isCore -and $_.Exception -is [Net.WebException]) { $resp = $_.Exception.Response }
            if ($resp -is [Net.HttpWebResponse]) {
                $text = ''
                try { $sr = New-Object IO.StreamReader($resp.GetResponseStream()); $text = $sr.ReadToEnd(); $sr.Dispose() } catch { $text = '' }
                $status = [int]$resp.StatusCode
                try { $resp.Close() } catch { }
                return @{ Status = $status; Content = $text }
            }
            $failure = Protect-RpSecret ([string]$_.Exception.Message) $secret
        }
    } finally {
        Remove-RpNewErrors $mark
        $secret = $null
    }
    throw (New-Object System.Exception($failure))
}

function Get-RpContentBytes($Content) {
    if ($null -eq $Content) { return ,([byte[]]@()) }
    if ($Content -is [byte[]]) { return ,$Content }
    if ($Content -is [string]) { return ,($script:RpUtf8.GetBytes($Content)) }
    return ,([byte[]]@($Content))
}

function Get-RpApiHint([int]$Status) {
    switch ($Status) {
        401 { return 'O login do GitHub guardado no Git foi recusado (expirado ou revogado): faca login de novo (git fetch abre a janela do Git Credential Manager).' }
        403 { return 'Sem permissao (ou limite de chamadas da API): a conta precisa poder escrever no repositorio e criar releases.' }
        404 { return 'Nao encontrado: confira o -Repo e se a conta tem acesso ao repositorio.' }
        422 { return 'O GitHub recusou os dados (por exemplo, tag ou asset ja existente).' }
        default { return '' }
    }
}

function Invoke-GitHubApi {
    param(
        [Parameter(Mandatory)][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [Management.Automation.PSCredential]$Credential = $null,
        $Body = $null,
        [byte[]]$BodyBytes = $null,
        [string]$ContentType = '',
        [string]$Accept = 'application/vnd.github+json',
        [switch]$Raw,
        [scriptblock]$Transport = $null
    )
    if (-not $Transport) { $Transport = $script:RpDefaultTransport }
    $u = $null
    $okHost = [uri]::TryCreate($Uri, [UriKind]::Absolute, [ref]$u) -and $u.Scheme -eq 'https' -and $script:RpAllowedHosts -contains $u.Host.ToLowerInvariant()
    # chave SO de teste (Test-PublishRelease, servidor HttpListener local): nunca vem de parametro, ambiente ou linha de comando
    if (-not $okHost -and $script:RpTestAllowLoopback -eq $true -and $u -and $u.Scheme -eq 'http' -and $u.Host -eq '127.0.0.1') { $okHost = $true }
    if (-not $okHost) {
        throw (New-RpError "Chamada recusada: so https://api.github.com e https://uploads.github.com sao permitidos ($Method $Uri)." 'falha')
    }
    $where = "$Method $($u.AbsolutePath)"
    $headers = @{ 'Accept' = $Accept; 'X-GitHub-Api-Version' = $script:RpApiVersion }
    $bytes = $null
    if ($null -ne $Body) {
        $bytes = $script:RpUtf8.GetBytes((ConvertTo-Json -InputObject $Body -Depth 10 -Compress))
        $ContentType = 'application/json; charset=utf-8'
    } elseif ($null -ne $BodyBytes) { $bytes = $BodyBytes }
    $secret = ''
    $resp = $null
    try {
        if ($Credential) { $secret = $Credential.GetNetworkCredential().Password; $headers['Authorization'] = 'Bearer ' + $secret }
        $req = @{ Method = $Method; Uri = $u.AbsoluteUri; Headers = $headers; Body = $bytes; ContentType = $ContentType; UserAgent = $script:RpUserAgent }
        $mark = Get-RpErrorMark
        $netFail = $null
        try { $resp = & $Transport $req }
        catch { $netFail = Protect-RpSecret ([string]$_.Exception.Message) $secret }
        finally { $headers.Remove('Authorization'); $req = $null; Remove-RpNewErrors $mark }
        if ($null -ne $netFail) { throw (New-RpError ("Falha de rede ao chamar o GitHub ($where): " + $netFail) 'falha') }
        if ($null -eq $resp) { throw (New-RpError "Sem resposta do GitHub ($where)." 'falha') }
        $status = [int]$resp.Status
        $content = Get-RpContentBytes $resp.Content
        if ($status -ge 200 -and $status -lt 300) {
            if ($Raw) { return ,$content }
            $text = $script:RpUtf8.GetString($content)
            if (-not $text.Trim()) { return $null }
            return ($text | ConvertFrom-Json)
        }
        $apiMsg = ''
        $text = $script:RpUtf8.GetString($content)
        try {
            $j = $text | ConvertFrom-Json
            if ($j.PSObject.Properties.Name -contains 'message') { $apiMsg = [string]$j.message }
            if ($j.PSObject.Properties.Name -contains 'errors') {
                $det = @($j.errors | ForEach-Object { if ($_ -is [string]) { $_ } else { (@($_.resource, $_.field, $_.code, $_.message) | Where-Object { $_ }) -join ' ' } })
                if ($det.Count -gt 0) { $apiMsg += ' (' + ($det -join '; ') + ')' }
            }
        } catch { $apiMsg = $text }
        if ($apiMsg.Length -gt 400) { $apiMsg = $apiMsg.Substring(0, 400) + '...' }
        if ($status -ge 300 -and $status -lt 400) { $apiMsg = 'redirecionamento para outro endereco (nao seguido: so api.github.com e uploads.github.com sao permitidos)' }
        $msg = "GitHub respondeu $status em $where" + $(if ($apiMsg) { ": $apiMsg" } else { '' })
        $hint = Get-RpApiHint $status
        if ($hint) { $msg += ". $hint" }
        throw (New-RpError (Protect-RpSecret $msg $secret) 'falha' $status)
    } finally { $secret = $null }
}

# ---------------------------------------------------------------------------------------------------------------------------------------
# Release no GitHub: rascunho -> 5 assets -> conferencia (tamanho e SHA-256) -> publica (sem -Draft).
# Retomada: um RASCUNHO anterior da mesma tag com o nome "ThinQQ <versao>" (deixado por uma execucao que falhou) e APAGADO e recriado do zero
# (assets enviados de novo); rascunho com outro nome, release publicada ou tag existente => recusa.
# Todas as releases (rascunhos inclusive), pagina a pagina (100 por pagina, ate 50 paginas)
function Get-RpAllReleases([string]$Api, [hashtable]$Call) {
    $all = New-Object System.Collections.Generic.List[object]
    for ($page = 1; $page -le 50; $page++) {
        $chunk = @(Invoke-GitHubApi -Method GET -Uri "$Api/releases?per_page=100&page=$page" @Call | Where-Object { $_ })
        foreach ($r in $chunk) { $all.Add($r) }
        if ($chunk.Count -lt 100) { break }
    }
    return $all.ToArray()   # enumerado de proposito: quem chama filtra no pipeline
}

function Publish-GitHubRelease {
    param(
        [Parameter(Mandatory)][string]$Repo,
        [Parameter(Mandatory)][string]$Version,
        [Parameter(Mandatory)][string]$StateDir,
        [Parameter(Mandatory)]$State,
        [Parameter(Mandatory)][string]$Body,
        [Parameter(Mandatory)][string]$TargetSha,
        [switch]$Draft,
        [Management.Automation.PSCredential]$Credential = $null,
        [scriptblock]$Transport = $null,
        [scriptblock]$Log = $null
    )
    if (-not $Log) { $Log = { param($m) Write-Host $m } }
    $api = "https://api.github.com/repos/$Repo"
    $tag = "v$Version"
    $name = "ThinQQ $Version"
    $call = @{ Credential = $Credential; Transport = $Transport }

    # 1. a tag e a release nao podem existir
    $tagExists = $true
    try { [void](Invoke-GitHubApi -Method GET -Uri "$api/git/ref/tags/$tag" @call) }
    catch { if ((Get-RpErrorStatus $_) -eq 404) { $tagExists = $false } else { throw } }
    if ($tagExists) { throw (New-RpError "A tag $tag ja existe no GitHub: escolha outra versao (nada foi alterado)." 'precondicao') }
    $existing = @(Get-RpAllReleases $api $call | Where-Object { $_ -and $_.tag_name -ceq $tag })
    foreach ($e in $existing) {
        if (-not $e.draft) { throw (New-RpError "A release $tag ja existe (publicada): $($e.html_url). Nada foi alterado." 'precondicao') }
        if ($e.name -cne $name) { throw (New-RpError "Ja existe um rascunho de release com a tag $tag e outro nome ('$($e.name)'): apague-o ou renomeie-o no GitHub antes ($($e.html_url))." 'precondicao') }
        if ((Get-NormalizedText ([string]$e.body)) -cne (Get-NormalizedText $Body)) {
            throw (New-RpError ("O rascunho '$name' ($($e.html_url)) foi editado no GitHub (o texto nao e o de release/notes.txt): ele NAO foi apagado. " +
                "Publique-o pela pagina dele, ou apague-o la e rode de novo.") 'precondicao')
        }
    }
    foreach ($e in $existing) {
        & $Log "Rascunho anterior de $tag encontrado (id $($e.id), de uma execucao interrompida): apagando para recriar."
        [void](Invoke-GitHubApi -Method DELETE -Uri "$api/releases/$($e.id)" @call)
    }

    # 2. rascunho
    $payload = [ordered]@{ tag_name = $tag; target_commitish = $TargetSha; name = $name; body = $Body; draft = $true; prerelease = $false }
    $rel = Invoke-GitHubApi -Method POST -Uri "$api/releases" -Body $payload @call
    if (-not $rel -or -not $rel.id -or -not $rel.upload_url) { throw (New-RpError 'O GitHub nao devolveu o id e o upload_url do rascunho.' 'falha') }
    & $Log "Rascunho criado: $($rel.html_url)"
    $resume = "O rascunho ficou em $($rel.html_url) (nada foi publicado). Para retomar, rode de novo: publicar-release.bat $Version publicar (o rascunho e apagado e recriado)."

    # 3. assets
    $uploadBase = ([string]$rel.upload_url) -replace '\{.*\}$', ''
    foreach ($n in (Get-ReleaseAssetNames $Version)) {
        $bytes = [IO.File]::ReadAllBytes((Join-Path $StateDir $n))
        try {
            [void](Invoke-GitHubApi -Method POST -Uri ($uploadBase + '?name=' + [uri]::EscapeDataString($n)) -BodyBytes $bytes -ContentType (Get-AssetContentType $n) @call)
        } catch { throw (New-RpError ("O envio de $n falhou: " + $_.Exception.Message + " $resume") 'falha' (Get-RpErrorStatus $_)) }
        & $Log ("  enviado: $n ({0:N0} bytes)" -f $bytes.Length)
    }

    # 4. conferencia
    $expected = @{}
    foreach ($f in @($State.files)) { $expected[[string]$f.name] = $f }
    $assets = @(Invoke-GitHubApi -Method GET -Uri "$api/releases/$($rel.id)/assets?per_page=100" @call | Where-Object { $_ })
    $problems = New-Object System.Collections.Generic.List[string]
    $checked = New-Object System.Collections.Generic.List[object]
    foreach ($n in (Get-ReleaseAssetNames $Version)) {
        $match = @($assets | Where-Object { $_.name -ceq $n })
        if ($match.Count -ne 1) { $problems.Add("$n aparece $($match.Count) vez(es) na release"); continue }
        $a = $match[0]
        if ([long]$a.size -ne [long]$expected[$n].size) { $problems.Add("$n tem $($a.size) bytes na release, o esperado e $($expected[$n].size)"); continue }
        $digest = ''
        if ($a.PSObject.Properties.Name -contains 'digest' -and $a.digest) { $digest = [string]$a.digest }
        $how = 'digest'
        if ($digest -match '^sha256:([0-9a-fA-F]{64})$') { $sha = $Matches[1].ToLowerInvariant() }
        else {
            $how = 'download'
            try {
                $got = Invoke-GitHubApi -Method GET -Uri "$api/releases/assets/$($a.id)" -Accept 'application/octet-stream' -Raw @call
                $sha = Get-BytesSha256 $got
            } catch { $problems.Add("$n sem digest e o download para conferir falhou ($($_.Exception.Message))"); continue }
        }
        if ($sha -cne [string]$expected[$n].sha256) { $problems.Add("$n tem SHA-256 $sha na release, o esperado e $($expected[$n].sha256)"); continue }
        $checked.Add([pscustomobject]@{ Name = $n; Size = [long]$a.size; Sha256 = $sha; Method = $how })
    }
    foreach ($a in $assets) { if ((Get-ReleaseAssetNames $Version) -cnotcontains [string]$a.name) { $problems.Add("asset inesperado na release: $($a.name)") } }
    if ($problems.Count -gt 0) {
        throw (New-RpError ("A conferencia dos arquivos enviados falhou: " + ($problems -join '; ') + ". $resume") 'falha')
    }
    & $Log ("Conferencia: os 5 arquivos batem com o state.json (tamanho e SHA-256; " + (($checked | ForEach-Object { "$($_.Name) por $($_.Method)" }) -join ', ') + ').')

    # 5. publicar
    $htmlUrl = [string]$rel.html_url
    if (-not $Draft) {
        $pub = Invoke-GitHubApi -Method PATCH -Uri "$api/releases/$($rel.id)" -Body ([ordered]@{ draft = $false; make_latest = 'true' }) @call
        if ($pub -and $pub.html_url) { $htmlUrl = [string]$pub.html_url }
    }
    return [pscustomobject]@{ Id = $rel.id; HtmlUrl = $htmlUrl; Published = (-not $Draft); Assets = $checked.ToArray() }
}

# ---------------------------------------------------------------------------------------------------------------------------------------
# Passos substituiveis da fase Prepare
function Get-RpPwsh {
    $c = Get-Command pwsh -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($c) { return $c.Source }
    return $null
}

# Test-Docs e Test-Privacy (rapidos, sem Excel) num processo proprio; devolve a lista do que falhou
$script:RpDefaultChecks = {
    param([string]$RepoRoot)
    $fails = @()
    $exe = Get-RpPwsh
    if (-not $exe) { $exe = (Get-Process -Id $PID).Path }
    foreach ($t in 'Test-Docs.ps1', 'Test-Privacy.ps1') {
        $r = Invoke-RpProcess $exe @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $RepoRoot "tests\ps\$t"), '-RepoRoot', $RepoRoot) -WorkingDirectory $RepoRoot
        if ($r.Code -ne 0) {
            $bad = @(($r.Output + "`n" + $r.Error) -split "\r?\n" | Where-Object { $_ -match '^not ok|FALHA|ERRO' } | Select-Object -First 15)
            $fails += ("$t (codigo $($r.Code))" + $(if ($bad.Count) { ":`n    " + ($bad -join "`n    ") } else { '' }))
        }
    }
    return $fails
}

# Constroi numa copia em <WorkRoot>\tqq-release-<versao> (robocopy sem .git, .superpowers e dist: o Excel do usuario pode travar dist\),
# roda (opcional) a suite e o build\release.ps1 com o pwsh e copia os 5 arquivos para $Ctx.OutDir. Apaga SO a copia que criou.
$script:RpDefaultBuild = {
    param($Ctx)
    $copy = Join-Path $Ctx.WorkRoot ("tqq-release-" + $Ctx.Version)
    if (Test-Path -LiteralPath $copy) { throw (New-RpError "A pasta $copy ja existe (de outra execucao?): apague-a e rode de novo." 'falha') }
    $pwsh = Get-RpPwsh
    if (-not $pwsh) { throw (New-RpError 'O PowerShell 7 (pwsh) nao foi encontrado: ele e preciso para o build.' 'falha') }
    $created = $false
    try {
        New-Item -ItemType Directory -Force $copy | Out-Null
        $created = $true
        & robocopy $Ctx.RepoRoot $copy /MIR /XD .git .superpowers dist /NFL /NDL /NJH /NJS /NP /R:1 /W:1 | Out-Null
        if ($LASTEXITCODE -ge 8) { throw (New-RpError "A copia para $copy falhou (robocopy codigo $LASTEXITCODE)." 'falha') }
        if ($Ctx.RunTests) {
            Write-Host "Rodando a suite completa na copia (pode levar alguns minutos)..."
            $global:LASTEXITCODE = 0
            & $pwsh -NoProfile -ExecutionPolicy Bypass -File (Join-Path $copy 'build\run-tests.ps1') -Target all -LoaderSwap -TimeoutSec 900
            if ($LASTEXITCODE -ne 0) { throw (New-RpError "A suite de testes falhou na copia (codigo $LASTEXITCODE): nada foi publicado." 'falha') }
        }
        Write-Host "Construindo o release na copia $copy (o build usa um Excel proprio e invisivel)..."
        $global:LASTEXITCODE = 0
        & $pwsh -NoProfile -ExecutionPolicy Bypass -File (Join-Path $copy 'build\release.ps1') -Version $Ctx.Version -BaseUrl $Ctx.BaseUrl
        if ($LASTEXITCODE -ne 0) { throw (New-RpError "O build\release.ps1 falhou (codigo $LASTEXITCODE)." 'falha') }
        foreach ($n in (Get-ReleaseAssetNames $Ctx.Version)) {
            $src = Join-Path $copy "dist\$n"
            if (-not (Test-Path -LiteralPath $src)) { throw (New-RpError "O build nao gerou dist\$n." 'falha') }
            Copy-Item -LiteralPath $src -Destination $Ctx.OutDir -Force
        }
    } finally {
        if ($created) { Remove-Item -LiteralPath $copy -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

$script:RpDefaultConfirm = {
    param([string]$Question)
    $a = Read-Host "$Question (s/N)"
    return ($a -match '^(s|sim|y|yes)$')
}

$script:RpDefaultBrowser = { param([string]$Url) Start-Process $Url }

# ---------------------------------------------------------------------------------------------------------------------------------------
# Fase Prepare
function Invoke-ReleasePrepare {
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [Parameter(Mandatory)][string]$Version,
        [string]$Repo = 'israeldias-hub/ThinQQ',
        [string]$NotesFile = 'release/notes.txt',
        [switch]$DryRun,
        [switch]$NoBrowser,
        [switch]$RunTests,
        [scriptblock]$BuildStep = $null,
        [scriptblock]$ChecksStep = $null,
        [scriptblock]$ConfirmStep = $null,
        [scriptblock]$BrowserStep = $null,
        [string]$Today = '',
        [string]$WorkRoot = 'C:\temp'
    )
    if (-not $BuildStep) { $BuildStep = $script:RpDefaultBuild }
    if (-not $ChecksStep) { $ChecksStep = $script:RpDefaultChecks }
    if (-not $ConfirmStep) { $ConfirmStep = $script:RpDefaultConfirm }
    if (-not $BrowserStep) { $BrowserStep = $script:RpDefaultBrowser }
    if (-not $Today) { $Today = Get-Date -Format 'yyyy-MM-dd' }
    $tag = "v$Version"
    $branch = "release/$tag"
    $pre = 'precondicao'
    Write-Host "== ThinQQ $Version - etapa 1 de 2: Preparar ==" -ForegroundColor Cyan
    if ($DryRun) { Write-Host '[ENSAIO] -DryRun: confere as pre-condicoes e mostra o que faria; nao grava, nao constroi, nao envia nada.' -ForegroundColor Yellow }

    # ---- pre-condicoes (nada e alterado ate o fim delas)
    if (-not (Test-SemVer $Version)) { throw (New-RpError "Versao invalida '$Version': use x.y.z (ex.: 1.2.1)." $pre) }
    if (-not (Test-PlainVersion $Version)) { throw (New-RpError "A etapa Preparar so aceita x.y.z (sem sufixo como '-rc1'): o build/release.ps1 e o version.json usam esse formato." $pre) }
    if (-not (Test-RepoName $Repo)) { throw (New-RpError "Repositorio invalido '$Repo' (use dono/nome)." $pre) }
    $status = Invoke-RpGitChecked $RepoRoot @('status', '--porcelain') 'git status' $pre
    if ($status) { throw (New-RpError ("Ha alteracoes nao commitadas (commit ou descarte antes de preparar a versao):`n" + $status) $pre) }
    $cur = Invoke-RpGitChecked $RepoRoot @('rev-parse', '--abbrev-ref', 'HEAD') 'git rev-parse' $pre
    if ($cur -cne 'main') { throw (New-RpError "A branch atual e '$cur': a versao e preparada a partir da main (git checkout main)." $pre) }
    if ($DryRun) { Write-Host '[ENSAIO] git fetch origin (pulado: compara com a origin/main ja baixada)' }
    else {
        Write-Host 'Atualizando a origin (git fetch)...'
        [void](Invoke-RpGitChecked $RepoRoot @('fetch', '--tags', 'origin') 'git fetch origin' $pre)
    }
    $mainSha = Invoke-RpGitChecked $RepoRoot @('rev-parse', 'main') 'git rev-parse main' $pre
    $originSha = Invoke-RpGitChecked $RepoRoot @('rev-parse', '--verify', '--quiet', 'refs/remotes/origin/main') 'origin/main' $pre
    if ($mainSha -cne $originSha) {
        $behind = (Invoke-RpGit $RepoRoot @('merge-base', '--is-ancestor', $mainSha, $originSha)).Code -eq 0
        $ahead = (Invoke-RpGit $RepoRoot @('merge-base', '--is-ancestor', $originSha, $mainSha)).Code -eq 0
        if ($behind) { throw (New-RpError 'A main local esta atrasada em relacao a origin/main: rode git pull --ff-only e tente de novo.' $pre) }
        if ($ahead) { throw (New-RpError 'A main local tem commits que nao estao na origin/main: envie-os por um pull request antes de preparar a versao.' $pre) }
        throw (New-RpError 'A main local e a origin/main divergiram: acerte a main (ela deve ser igual a origin/main) antes de preparar a versao.' $pre)
    }
    $vjPath = Join-Path $RepoRoot 'version.json'
    $vjText = Read-RpText $vjPath
    $vj = $vjText | ConvertFrom-Json
    foreach ($cv in @($vj.core, $vj.loader)) {
        if ((Compare-SemVer $Version $cv) -le 0) { throw (New-RpError "A versao $Version precisa ser maior que a atual do version.json (core $($vj.core), loader $($vj.loader))." $pre) }
    }
    if ((Invoke-RpGit $RepoRoot @('rev-parse', '--verify', '--quiet', "refs/tags/$tag")).Code -eq 0) { throw (New-RpError "A tag $tag ja existe neste repositorio." $pre) }
    foreach ($ref in @("refs/heads/$branch", "refs/remotes/origin/$branch")) {
        if ((Invoke-RpGit $RepoRoot @('rev-parse', '--verify', '--quiet', $ref)).Code -eq 0) { throw (New-RpError "A branch $branch ja existe ($ref): apague-a ou use outra versao." $pre) }
    }
    $notesRel = Get-RpRepoRelativePath $RepoRoot $NotesFile
    if (-not $notesRel) { throw (New-RpError "O arquivo de notas precisa estar dentro do repositorio: $NotesFile" $pre) }
    $notesPath = Join-Path $RepoRoot $notesRel
    if (-not (Test-Path -LiteralPath $notesPath)) { throw (New-RpError "Arquivo de notas nao encontrado: $notesRel" $pre) }
    $notesText = Read-RpText $notesPath
    if (-not (Get-NormalizedText $notesText)) { throw (New-RpError "$notesRel esta vazio: escreva as notas da versao $Version." $pre) }
    $lastTag = Get-RpLastTag $RepoRoot
    if ($lastTag) {
        $old = Invoke-RpGit $RepoRoot @('show', "${lastTag}:$notesRel")
        if ($old.Code -eq 0 -and (Get-TextSha256 $old.Output) -ceq (Get-TextSha256 $notesText)) {
            throw (New-RpError "$notesRel nao mudou desde a ultima versao ($lastTag): escreva as notas da versao $Version (de 6 a 8 linhas com o marcador), faca o commit na main e rode de novo." $pre)
        }
        Write-Host "Notas: $notesRel mudou desde $lastTag."
    } else { Write-Host "Notas: nenhuma tag v* anterior; $notesRel nao foi comparado." }
    $clPath = Join-Path $RepoRoot 'CHANGELOG.md'
    $clText = Read-RpText $clPath
    $entries = Get-UnreleasedEntries $clText
    if ($null -eq $entries) { throw (New-RpError ('O CHANGELOG.md nao tem a secao "' + $script:RpUnreleased + '".') $pre) }
    if ($entries.Count -eq 0) {
        Write-Host ('AVISO: a secao "' + $script:RpUnreleased + '" do CHANGELOG.md esta vazia (nenhuma entrada vai para a versao ' + $Version + ').') -ForegroundColor Yellow
        if (-not $DryRun) {
            if (-not (& $ConfirmStep 'Continuar mesmo assim?')) { throw (New-RpError 'Cancelado: escreva as entradas do CHANGELOG.md, faca o commit na main e rode de novo.' $pre) }
        }
    } else { Write-Host "CHANGELOG: $($entries.Count) entrada(s) em [Nao lancado]." }
    Write-Host 'Conferindo a documentacao e a privacidade (Test-Docs, Test-Privacy)...'
    $fails = @(& $ChecksStep $RepoRoot | Where-Object { $_ })
    if ($fails.Count -gt 0) { throw (New-RpError ("As verificacoes falharam (corrija na main e rode de novo):`n  " + ($fails -join "`n  ")) $pre) }
    $outDir = Join-Path $RepoRoot "dist\release-$Version"
    if (Test-Path -LiteralPath $outDir) { throw (New-RpError "A pasta $outDir ja existe (de uma preparacao anterior): apague-a e rode de novo." $pre) }
    $baseUrl = Get-ReleaseBaseUrl $Repo $Version
    $compare = Get-CompareUrl $Repo $Version
    $changedFiles = @('version.json', 'CHANGELOG.md', 'docs/tutorial/tutorial.html', 'installer/LEIA-ME.txt', 'release/manifest.json')
    Write-Host 'Pre-condicoes ok.' -ForegroundColor Green

    if ($DryRun) {
        Write-Host "[ENSAIO] criaria a branch $branch a partir da main ($mainSha)"
        Write-Host "[ENSAIO] gravaria version.json (loader e core $Version), CHANGELOG.md ([$Version] - $Today e [Nao lancado] nova), rodape do tutorial e exemplo do LEIA-ME"
        Write-Host ("[ENSAIO] construiria numa copia em " + (Join-Path $WorkRoot "tqq-release-$Version") + " com build\release.ps1 -Version $Version -BaseUrl $baseUrl" + $(if ($RunTests) { ' (depois da suite completa)' } else { '' }))
        Write-Host "[ENSAIO] copiaria $((Get-ReleaseAssetNames $Version) -join ', ') e state.json para $outDir"
        Write-Host "[ENSAIO] copiaria manifest.json para release/manifest.json e faria o commit 'release: ThinQQ $Version'"
        Write-Host "[ENSAIO] git push -u origin $branch e abriria $compare"
        return [pscustomobject]@{ Version = $Version; Branch = $branch; DryRun = $true; CompareUrl = $compare; StateDir = $outDir }
    }

    # ---- execucao
    $committed = $false
    $branchCreated = $false
    $outCreated = $false
    $distCreated = $false
    $distDir = ''
    try {
        [void](Invoke-RpGitChecked $RepoRoot @('checkout', '-b', $branch) "Criar a branch $branch")
        $branchCreated = $true
        Write-RpText $vjPath (Update-VersionJsonText $vjText $Version)
        Write-RpText $clPath (Update-ChangelogText $clText $Version $Today)
        foreach ($pair in @(@('docs\tutorial\tutorial.html', 'Update-TutorialVersionText'), @('installer\LEIA-ME.txt', 'Update-LeiaMeVersionText'))) {
            $p = Join-Path $RepoRoot $pair[0]
            if (-not (Test-Path -LiteralPath $p)) { Write-Host "AVISO: $($pair[0]) nao existe (versao nao atualizada nele)." -ForegroundColor Yellow; continue }
            $t = Read-RpText $p
            $fn = $pair[1]
            $n = & $fn $t $Version
            if ($n -ceq $t) { Write-Host "AVISO: nao achei a versao para trocar em $($pair[0])." -ForegroundColor Yellow } else { Write-RpText $p $n }
        }
        Write-Host "version.json, CHANGELOG.md, tutorial e LEIA-ME atualizados para $Version."
        $distDir = Split-Path $outDir
        $distCreated = -not (Test-Path -LiteralPath $distDir)
        New-Item -ItemType Directory -Force $outDir | Out-Null
        $outCreated = $true
        & $BuildStep @{ RepoRoot = $RepoRoot; Version = $Version; BaseUrl = $baseUrl; OutDir = $outDir; RunTests = [bool]$RunTests; WorkRoot = $WorkRoot } | Out-Host
        foreach ($n in (Get-ReleaseAssetNames $Version)) {
            if (-not (Test-Path -LiteralPath (Join-Path $outDir $n))) { throw (New-RpError "O build nao entregou $n em $outDir." 'falha') }
        }
        Copy-Item -LiteralPath (Join-Path $outDir 'manifest.json') -Destination (Join-Path $RepoRoot 'release\manifest.json') -Force
        $toAdd = @(@($changedFiles + @($notesRel)) | Where-Object { Test-Path -LiteralPath (Join-Path $RepoRoot $_) } | Select-Object -Unique)
        [void](Invoke-RpGitChecked $RepoRoot (@('add', '--') + $toAdd) 'git add')
        [void](Invoke-RpGitChecked $RepoRoot @('commit', '-q', '-m', "release: ThinQQ $Version") 'git commit')
        $committed = $true
        $prepSha = Invoke-RpGitChecked $RepoRoot @('rev-parse', 'HEAD') 'git rev-parse HEAD'
        Write-ReleaseState (New-ReleaseState $Version $Repo $outDir $prepSha $mainSha) $outDir
        Write-Host "Arquivos e state.json em $outDir"
    } catch {
        $err = $_
        if (-not $committed) {
            Write-Host 'Falhou antes do commit: desfazendo (arquivos, branch e pasta da preparacao)...' -ForegroundColor Yellow
            if ($branchCreated) {
                $tracked = @($changedFiles | Where-Object { (Invoke-RpGit $RepoRoot @('cat-file', '-e', "HEAD:$_")).Code -eq 0 })
                if ($tracked.Count -gt 0) { [void](Invoke-RpGit $RepoRoot (@('checkout', 'HEAD', '--') + $tracked)) }
                [void](Invoke-RpGit $RepoRoot @('checkout', 'main'))
                [void](Invoke-RpGit $RepoRoot @('branch', '-D', $branch))
            }
            if ($outCreated) {
                Remove-Item -LiteralPath $outDir -Recurse -Force -ErrorAction SilentlyContinue
                if ($distCreated -and @(Get-ChildItem -LiteralPath $distDir -Force -ErrorAction SilentlyContinue).Count -eq 0) { Remove-Item -LiteralPath $distDir -Force -ErrorAction SilentlyContinue }
            }
        }
        throw $err
    }

    Write-Host "Enviando a branch $branch (git push)..."
    $push = Invoke-RpGit $RepoRoot @('push', '-u', 'origin', $branch)
    if ($push.Code -ne 0) {
        throw (New-RpError ("O commit 'release: ThinQQ $Version' esta pronto na branch $branch, mas o envio falhou: " + (Protect-RpSecret ($push.Error.Trim())) +
            "`nConfira a rede e o login do Git e rode: git push -u origin $branch ; depois abra $compare") 'falha')
    }
    if ($NoBrowser) { Write-Host "Abra o pull request: $compare" }
    else { Write-Host "Abrindo o pull request no navegador: $compare"; & $BrowserStep $compare }
    Write-Host ''
    Write-Host "Proximo passo: crie e mescle o pull request (Create a merge commit) e depois rode: publicar-release.bat $Version publicar" -ForegroundColor Green
    return [pscustomobject]@{ Version = $Version; Branch = $branch; DryRun = $false; CompareUrl = $compare; StateDir = $outDir; PrepareCommit = $prepSha; SourceMain = $mainSha }
}

# ---------------------------------------------------------------------------------------------------------------------------------------
# Fase Publish
function Invoke-ReleasePublish {
    param(
        [Parameter(Mandatory)][string]$RepoRoot,
        [Parameter(Mandatory)][string]$Version,
        [string]$Repo = 'israeldias-hub/ThinQQ',
        [string]$NotesFile = 'release/notes.txt',
        [string]$StateDir = '',
        [switch]$Draft,
        [switch]$DryRun,
        [switch]$NoBrowser,
        [switch]$AllowMainChanges,
        [scriptblock]$Transport = $null,
        [scriptblock]$CredentialRunner = $null,
        [scriptblock]$BrowserStep = $null
    )
    if (-not $BrowserStep) { $BrowserStep = $script:RpDefaultBrowser }
    if (-not $StateDir) { $StateDir = Join-Path $RepoRoot "dist\release-$Version" }
    $pre = 'precondicao'
    $tag = "v$Version"
    Write-Host "== ThinQQ $Version - etapa 2 de 2: Publicar ==" -ForegroundColor Cyan
    if ($DryRun) { Write-Host '[ENSAIO] -DryRun: confere as pre-condicoes locais e mostra o que faria; nao le o token nem chama a API.' -ForegroundColor Yellow }
    if (-not (Test-SemVer $Version)) { throw (New-RpError "Versao invalida '$Version': use x.y.z (ex.: 1.2.1)." $pre) }
    if (-not (Test-RepoName $Repo)) { throw (New-RpError "Repositorio invalido '$Repo' (use dono/nome)." $pre) }

    $state = Read-ReleaseState $StateDir $Version
    if ([string]$state.repo -cne $Repo) {
        throw (New-RpError "O state.json de $StateDir e do repositorio '$($state.repo)', nao de '$Repo': confira o -Repo (nada foi alterado)." $pre)
    }
    Write-Host "Arquivos da preparacao conferidos ($StateDir)."

    if ($DryRun) { Write-Host '[ENSAIO] git fetch origin (pulado: usa a origin/main ja baixada)' }
    else { [void](Invoke-RpGitChecked $RepoRoot @('fetch', 'origin') 'git fetch origin' $pre) }
    $originSha = Invoke-RpGitChecked $RepoRoot @('rev-parse', '--verify', '--quiet', 'refs/remotes/origin/main') 'origin/main' $pre
    $prep = [string]$state.prepareCommit
    $short = $prep.Substring(0, 7)
    if ((Invoke-RpGit $RepoRoot @('cat-file', '-e', "$prep^{commit}")).Code -ne 0) {
        throw (New-RpError "Nao achei o commit de preparacao $prep neste repositorio: rode git fetch e confira a branch release/$tag." $pre)
    }
    $remoteManifest = Invoke-RpGit $RepoRoot @('show', "${originSha}:release/manifest.json")
    $localManifest = Read-RpText (Join-Path $StateDir 'manifest.json')
    $manifestOk = ($remoteManifest.Code -eq 0) -and ((Get-TextSha256 $remoteManifest.Output) -ceq (Get-TextSha256 $localManifest))
    $anc = Invoke-RpGit $RepoRoot @('merge-base', '--is-ancestor', $prep, $originSha)
    if ($anc.Code -eq 1) {
        if (-not $manifestOk) {
            throw (New-RpError "O commit de preparacao ($short) ainda nao esta na origin/main: mescle o pull request da branch release/$tag (com 'Create a merge commit') e rode de novo." $pre)
        }
        # mesclado por 'Squash and merge' ou 'Rebase and merge': o commit muda, o conteudo dos arquivos da versao tem de ser o mesmo
        $same = (Invoke-RpGit $RepoRoot (@('diff', '--quiet', $prep, $originSha, '--') + $script:RpVersionFiles)).Code -eq 0
        if (-not $same) {
            throw (New-RpError ("O pull request da branch release/$tag parece ter sido mesclado por 'Squash and merge' ou 'Rebase and merge' (o commit de preparacao $short nao esta na origin/main), " +
                "e os arquivos da versao na origin/main (" + ($script:RpVersionFiles -join ', ') + ") nao sao os do commit de preparacao. Nada foi publicado. " +
                "Publique pela 'Alternativa manual' do README (secao Publicar uma versao) usando os arquivos de $StateDir, ou prepare uma versao nova; " +
                "use sempre 'Create a merge commit' no pull request da versao.") $pre)
        }
        Write-Host "O pull request foi mesclado por squash/rebase: o commit de preparacao ($short) nao esta na origin/main, mas os arquivos da versao sao identicos; seguindo." -ForegroundColor Yellow
    } elseif ($anc.Code -ne 0) {
        throw (New-RpError "Nao foi possivel comparar o commit de preparacao $prep com a origin/main (git merge-base, codigo $($anc.Code))." $pre)
    }
    if ($remoteManifest.Code -ne 0) { throw (New-RpError 'A origin/main nao tem release/manifest.json.' $pre) }
    if (-not $manifestOk) {
        throw (New-RpError "O release/manifest.json da origin/main nao e o manifest.json desta preparacao ($StateDir): os assets nao casariam com o manifesto que os usuarios leem. Nada foi publicado." $pre)
    }
    # a tag vai para a origin/main: o codigo dela tem de ser o que gerou os arquivos (um PR mesclado depois do da versao muda isso)
    $diffAll = Invoke-RpGit $RepoRoot @('diff', '--name-only', $prep, $originSha)
    if ($diffAll.Code -ne 0) { throw (New-RpError "Nao foi possivel comparar o commit de preparacao com a origin/main (git diff, codigo $($diffAll.Code))." $pre) }
    $changedSince = @(([string]$diffAll.Output) -split "\r?\n" | Where-Object { $_.Trim() })
    if ($changedSince.Count -gt 0) {
        $list = (@($changedSince | Select-Object -First 10) -join ', ') + $(if ($changedSince.Count -gt 10) { ', ...' } else { '' })
        if (-not $AllowMainChanges) {
            throw (New-RpError ("A origin/main mudou depois da preparacao ($($changedSince.Count) arquivo(s): $list): a tag $tag apontaria para um codigo diferente do que gerou os arquivos. " +
                "Prepare uma versao nova, ou rode com -AllowMainChanges para publicar assim mesmo (a tag fica na origin/main atual).") $pre)
        }
        Write-Host "AVISO (-AllowMainChanges): a origin/main mudou depois da preparacao ($list); a tag $tag vai para a origin/main atual, os arquivos sao os da preparacao." -ForegroundColor Yellow
    }
    Write-Host "origin/main ($($originSha.Substring(0, 7))) contem a preparacao e o manifesto confere."

    # texto da release: o release/notes.txt da origin/main (o que foi revisado e mesclado), nao o da arvore de trabalho; um -NotesFile fora do
    # repositorio (ensaio) e lido do disco
    $notesRel = Get-RpRepoRelativePath $RepoRoot $NotesFile
    if ($notesRel) {
        $nshow = Invoke-RpGit $RepoRoot @('show', "${originSha}:$notesRel")
        if ($nshow.Code -ne 0) { throw (New-RpError "A origin/main nao tem $notesRel." $pre) }
        $body = New-ReleaseBody $nshow.Output
        $notesSource = "origin/main:$notesRel"
    } else {
        if (-not (Test-Path -LiteralPath $NotesFile)) { throw (New-RpError "Arquivo de notas nao encontrado: $NotesFile" $pre) }
        $body = New-ReleaseBody (Read-RpText $NotesFile)
        $notesSource = $NotesFile
    }

    if ($DryRun) {
        Write-Host "[ENSAIO] conferiria na API que a tag $tag e a release nao existem"
        Write-Host "[ENSAIO] criaria o rascunho 'ThinQQ $Version' (tag $tag, alvo $originSha) com o texto de $notesSource"
        foreach ($f in @($state.files)) { Write-Host ("[ENSAIO] enviaria {0} ({1:N0} bytes, {2})" -f $f.name, [long]$f.size, (Get-AssetContentType $f.name)) }
        Write-Host '[ENSAIO] conferiria tamanho e SHA-256 de cada asset'
        if ($Draft) { Write-Host '[ENSAIO] deixaria a release em RASCUNHO (-Draft)' } else { Write-Host '[ENSAIO] publicaria a release (draft=false, make_latest=true)' }
        return [pscustomobject]@{ Version = $Version; DryRun = $true; Published = $false; HtmlUrl = '' }
    }

    $cred = Get-GitHubCredential $RepoRoot $Repo $CredentialRunner
    try {
        $res = Publish-GitHubRelease -Repo $Repo -Version $Version -StateDir $StateDir -State $state -Body $body -TargetSha $originSha -Draft:$Draft -Credential $cred -Transport $Transport
    } finally {
        if ($cred) { $cred.Password.Dispose() }
        $cred = $null
    }
    Write-Host ''
    if ($res.Published) { Write-Host "Release publicada: $($res.HtmlUrl)" -ForegroundColor Green }
    else {
        Write-Host "Release em RASCUNHO (-Draft; nada foi publicado): $($res.HtmlUrl)" -ForegroundColor Green
        Write-Host 'Para publicar, use o botao da pagina do rascunho. Se for editar o texto do rascunho no GitHub, publique-o por la: rodar esta etapa de novo recusaria (o texto deixa de ser o de release/notes.txt).'
    }
    if (-not $NoBrowser) { & $BrowserStep $res.HtmlUrl }
    Write-Host 'Falta a verificacao manual: baixe o ZIP pela pagina da release, extraia a pasta INTEIRA, rode Instalar-ThinQQ.cmd e confira a aba ThinQQ no Excel.'
    return [pscustomobject]@{ Version = $Version; DryRun = $false; Published = $res.Published; HtmlUrl = $res.HtmlUrl; Id = $res.Id; Assets = $res.Assets }
}
