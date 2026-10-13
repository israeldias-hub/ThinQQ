# Testa o publicar-release (tools/Publish-Release.ps1, tools/lib/Release-Publish.ps1 e publicar-release.bat) SEM rede e SEM Excel:
# funcoes puras (semver, CHANGELOG, version.json, corpo da release, saida do git credential fill, Content-Type), a fase Publish contra um
# backend HTTP FALSO (transporte injetado em Invoke-GitHubApi), a fase Prepare num repositorio git DESCARTAVEL com um remoto bare local
# (o passo de build e substituido: nada de Excel), -DryRun (nao grava nem chama o transporte) e o token: um valor SENTINELA nunca aparece na
# saida capturada de nenhum cenario nem em arquivo algum da pasta temporaria.
# Roda no pwsh 7 e no Windows PowerShell 5.1 (o run-tests o chama nos dois, num processo proprio): termina sempre com exit 0 ou 1 e apaga so
# a pasta tqq_pub_* que criou em %TEMP%.
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$script:fail = 0
$script:allOutput = New-Object Text.StringBuilder
function Check([bool]$c, [string]$m) { if ($c) { Write-Host "ok - $m" } else { Write-Host "not ok - $m" -ForegroundColor Red; $script:fail++ } }
$tmp = Join-Path $env:TEMP ('tqq_pub_' + [guid]::NewGuid().ToString('N'))
$exitCode = 1
$sentinel = 'SENTINELA' + [guid]::NewGuid().ToString('N')
$R = 'israeldias-hub/ThinQQ'
$unrel = '## [N' + [char]0xE3 + 'o lan' + [char]0xE7 + 'ado]'

# Roda um cenario capturando TODA a saida (pipeline, Write-Host, avisos, erros) e a excecao, se houver
function Invoke-Scenario([scriptblock]$Block) {
    $script:scnError = $null
    $script:scnResult = $null
    $captured = (& { try { $script:scnResult = & $Block } catch { $script:scnError = $_ } } *>&1 | Out-String -Width 4096)
    $kind = ''; $msg = ''; $status = 0
    if ($script:scnError) {
        $kind = Get-RpErrorKind $script:scnError
        $msg = $script:scnError.Exception.Message
        $status = Get-RpErrorStatus $script:scnError
        $captured += "`nEXCECAO: $msg"
    }
    [void]$script:allOutput.Append($captured)
    return [pscustomobject]@{ Output = $captured; Failed = [bool]$script:scnError; Kind = $kind; Message = $msg; Status = $status; Result = $script:scnResult }
}

# ---- backend falso do GitHub
$script:gh = $null
function Reset-FakeGitHub {
    $script:gh = @{ Releases = New-Object Collections.ArrayList; Tags = @(); Requests = New-Object Collections.ArrayList; NextId = 100; NextAsset = 500
        FailUploadOn = ''; DigestMode = 'ok'; ForceStatus = @{}; ForceMessage = @{}; Blobs = @{} }
}
function New-FakeResponse([int]$Status, $Obj) {
    $c = ''
    if ($null -ne $Obj) { $c = ConvertTo-Json -InputObject $Obj -Depth 8 }
    return @{ Status = $Status; Content = $c }
}
function Get-FakeRelease($Id) { return @($script:gh.Releases | Where-Object { $_.id -eq [int]$Id }) | Select-Object -First 1 }
$fakeTransport = {
    param($Request)
    $u = [uri]$Request.Uri
    $auth = ''
    if ($Request.Headers.ContainsKey('Authorization')) { $auth = [string]$Request.Headers['Authorization'] }
    $bodyText = ''
    if ($null -ne $Request.Body -and $Request.ContentType -like 'application/json*') { $bodyText = [Text.Encoding]::UTF8.GetString([byte[]]$Request.Body) }
    $bodyLen = 0
    if ($null -ne $Request.Body) { $bodyLen = ([byte[]]$Request.Body).Length }
    [void]$script:gh.Requests.Add([pscustomobject]@{ Method = $Request.Method; Uri = $Request.Uri; Host = $u.Host; Path = $u.AbsolutePath; Query = $u.Query
            Auth = $auth; Accept = [string]$Request.Headers['Accept']; ContentType = $Request.ContentType; BodyText = $bodyText; BodyLength = $bodyLen })
    $key = "$($Request.Method) $($u.AbsolutePath)"
    if ($script:gh.ForceStatus.ContainsKey($key)) { return (New-FakeResponse $script:gh.ForceStatus[$key] @{ message = $script:gh.ForceMessage[$key] }) }
    $p = '/repos/israeldias-hub/ThinQQ'
    if ($key -match "^GET $p/git/ref/tags/(.+)$") {
        if ($script:gh.Tags -contains $Matches[1]) { return (New-FakeResponse 200 @{ ref = "refs/tags/$($Matches[1])" }) }
        return (New-FakeResponse 404 @{ message = 'Not Found' })
    }
    if ($key -eq "GET $p/releases") {
        $pp = 30; $pg = 1
        if ($u.Query -match 'per_page=(\d+)') { $pp = [int]$Matches[1] }
        if ($u.Query -match '[?&]page=(\d+)') { $pg = [int]$Matches[1] }
        $all = @($script:gh.Releases)
        $slice = @($all | Select-Object -Skip (($pg - 1) * $pp) -First $pp)
        return (New-FakeResponse 200 $slice)
    }
    if ($key -eq "POST $p/releases") {
        $b = $bodyText | ConvertFrom-Json
        $id = $script:gh.NextId; $script:gh.NextId++
        $rel = [ordered]@{ id = $id; tag_name = $b.tag_name; target_commitish = $b.target_commitish; name = $b.name; body = $b.body; draft = [bool]$b.draft
            prerelease = [bool]$b.prerelease; html_url = "https://github.com/$R/releases/tag/untagged-$id"
            upload_url = "https://uploads.github.com/repos/$R/releases/$id/assets{?name,label}"; assets = (New-Object Collections.ArrayList) }
        [void]$script:gh.Releases.Add($rel)
        return (New-FakeResponse 201 $rel)
    }
    if ($u.Host -eq 'uploads.github.com' -and $key -match "^POST $p/releases/(\d+)/assets$") {
        $rel = Get-FakeRelease $Matches[1]
        $name = [uri]::UnescapeDataString(($u.Query -replace '^\?name=', ''))
        if ($script:gh.FailUploadOn -ceq $name) { return (New-FakeResponse 500 @{ message = 'Server Error' }) }
        $bytes = [byte[]]$Request.Body
        $aid = $script:gh.NextAsset; $script:gh.NextAsset++
        $sha = Get-BytesSha256 $bytes
        $digest = 'sha256:' + $sha
        if ($script:gh.DigestMode -eq 'bad' -and $name -eq 'ThinQQ.core.xlam') { $digest = 'sha256:' + ('0' * 64) }
        if ($script:gh.DigestMode -like 'none*') { $digest = $null }
        $stored = $bytes
        if ($script:gh.DigestMode -eq 'none-bad' -and $name -eq 'tutorial.html') { $stored = [byte[]]($bytes.Clone()); $stored[0] = [byte](($stored[0] + 1) % 256) }
        $script:gh.Blobs[$aid] = $stored
        $asset = [ordered]@{ id = $aid; name = $name; size = $bytes.Length; content_type = $Request.ContentType; state = 'uploaded'; digest = $digest }
        [void]$rel.assets.Add($asset)
        return (New-FakeResponse 201 $asset)
    }
    if ($key -match "^GET $p/releases/(\d+)/assets$") { return (New-FakeResponse 200 @((Get-FakeRelease $Matches[1]).assets)) }
    if ($key -match "^GET $p/releases/assets/(\d+)$") { return @{ Status = 200; Content = [byte[]]$script:gh.Blobs[[int]$Matches[1]] } }
    if ($key -match "^PATCH $p/releases/(\d+)$") {
        $rel = Get-FakeRelease $Matches[1]
        $b = $bodyText | ConvertFrom-Json
        $rel.draft = [bool]$b.draft
        if (-not $rel.draft) { $rel.html_url = "https://github.com/$R/releases/tag/$($rel.tag_name)" }
        return (New-FakeResponse 200 $rel)
    }
    if ($key -match "^DELETE $p/releases/(\d+)$") {
        $rel = Get-FakeRelease $Matches[1]
        $script:gh.Releases.Remove($rel)
        return @{ Status = 204; Content = '' }
    }
    return (New-FakeResponse 404 @{ message = "rota falsa inexistente: $key" })
}
function Get-FakeRequests([string]$Method, [string]$PathRegex) { return ,@($script:gh.Requests | Where-Object { $_.Method -eq $Method -and $_.Path -match $PathRegex }) }

try {
    . (Join-Path $repo 'tools\lib\Release-Publish.ps1')
    New-Item -ItemType Directory -Force $tmp | Out-Null

    # ================================================================ funcoes puras
    foreach ($v in '0.0.0', '1.2.3', '10.20.30', '1.0.0-rc.1', '0.0.0-ensaio', '1.0.0-alpha-1') { Check (Test-SemVer $v) "semver valida: $v" }
    foreach ($v in '', '1.2', '1.2.3.4', 'v1.2.3', '01.2.3', '1.02.3', '1.2.3-', '1.2.3-01', '1.2.3+build', ' 1.2.3', 'abc') { Check (-not (Test-SemVer $v)) "semver invalida: '$v'" }
    Check ((Test-PlainVersion '1.2.1') -and -not (Test-PlainVersion '1.2.1-rc1')) 'Test-PlainVersion so aceita x.y.z'
    Check (-not (Test-SemVer "1.2.3`n") -and -not (Test-PlainVersion "1.2.3`n") -and -not (Test-SemVer "1.2.3-rc`n")) 'versao com quebra de linha no fim: invalida (\z, nao $)'
    foreach ($rn in 'israeldias-hub/ThinQQ', 'a.b/c-d_e', 'A/b.c') { Check (Test-RepoName $rn) "repositorio valido: $rn" }
    foreach ($rn in '../..', 'a/..', '../b', './b', 'a/.', '.a/b', 'a./b', 'a/.b', 'a/b.', "a/b`n", 'a/b/c', 'a', '/b', 'a/', 'a b/c', '') { Check (-not (Test-RepoName $rn)) "repositorio invalido: '$($rn.Replace("`n", '\n'))'" }
    $cmp = @(
        @('1.2.1', '1.2.0', 1), @('1.2.0', '1.2.0', 0), @('1.1.9', '1.2.0', -1), @('1.10.0', '1.9.9', 1), @('2.0.0', '1.99.99', 1),
        @('1.0.0', '1.0.0-rc.1', 1), @('1.0.0-rc.2', '1.0.0-rc.1', 1), @('1.0.0-alpha', '1.0.0-alpha.1', -1),
        @('1.0.0-alpha.beta', '1.0.0-alpha.1', 1), @('1.0.0-beta.11', '1.0.0-beta.2', 1), @('1.0.0-rc.1', '1.0.0-beta', 1))
    foreach ($c in $cmp) { Check ((Compare-SemVer $c[0] $c[1]) -eq $c[2]) "Compare-SemVer $($c[0]) vs $($c[1]) = $($c[2])" }
    $threw = $false; try { [void](Compare-SemVer '1.2' '1.2.0') } catch { $threw = $true }
    Check $threw 'Compare-SemVer recusa versao invalida'

    $clLf = "# Changelog`n`n$unrel`n`n### Adicionado`n- Uma coisa.`n- Outra.`n`n## [1.2.0] - 2026-10-12`n`n### Adicionado`n- Antiga.`n"
    $e = Get-UnreleasedEntries $clLf
    Check ($e.Count -eq 2) "Get-UnreleasedEntries conta as entradas de [Nao lancado] ($($e.Count))"
    Check ((Get-UnreleasedEntries "# C`n`n$unrel`n`n## [1.2.0] - 2026-10-12`n- velha`n").Count -eq 0) 'Get-UnreleasedEntries: secao vazia => 0 (nao conta a versao de baixo)'
    Check ($null -eq (Get-UnreleasedEntries "# C`n## [1.2.0] - x`n")) 'Get-UnreleasedEntries: sem a secao => $null'
    $newCl = Update-ChangelogText $clLf '1.2.1' '2026-10-13'
    Check ($newCl.Contains("$unrel`n`n## [1.2.1] - 2026-10-13`n`n### Adicionado`n- Uma coisa.")) 'CHANGELOG: [Nao lancado] vira [1.2.1] - data e uma [Nao lancado] vazia fica no topo'
    Check ((Get-UnreleasedEntries $newCl).Count -eq 0) 'CHANGELOG atualizado: [Nao lancado] nova sem entradas'
    Check ($newCl.IndexOf($unrel) -lt $newCl.IndexOf('## [1.2.1]') -and $newCl.IndexOf('## [1.2.1]') -lt $newCl.IndexOf('## [1.2.0]')) 'CHANGELOG: ordem [Nao lancado], [1.2.1], [1.2.0]'
    $newCrlf = Update-ChangelogText ($clLf.Replace("`n", "`r`n")) '1.2.1' '2026-10-13'
    Check ($newCrlf.Contains("$unrel`r`n`r`n## [1.2.1] - 2026-10-13`r`n") -and -not ($newCrlf -match "[^\r]`n")) 'CHANGELOG com CRLF continua so com CRLF'
    $threw = $false; try { [void](Update-ChangelogText "# C`n## [1.2.0] - x`n" '1.2.1' '2026-10-13') } catch { $threw = $true }
    Check $threw 'CHANGELOG sem [Nao lancado]: erro'
    $threw = $false; try { [void](Update-ChangelogText $newCl '1.2.1' '2026-10-13') } catch { $threw = $true }
    Check $threw 'CHANGELOG que ja tem a secao da versao: erro'

    $vjNew = Update-VersionJsonText ('{ "loader": "1.2.0", "core": "1.2.0" }' + "`n") '1.2.1'
    Check ($vjNew -ceq ('{ "loader": "1.2.1", "core": "1.2.1" }' + "`n")) "version.json: loader e core trocados no mesmo formato ($vjNew)"
    $threw = $false; try { [void](Update-VersionJsonText '{ "core": "1.2.0" }' '1.2.1') } catch { $threw = $true }
    Check $threw 'version.json sem loader: erro'
    Check ((Update-TutorialVersionText '<p>ThinQQ 1.2.0. Duvidas: x</p>' '1.2.1') -ceq '<p>ThinQQ 1.2.1. Duvidas: x</p>') 'rodape do tutorial'
    Check ((Update-LeiaMeVersionText 'ex.: versions\1.2.0\). Copie' '1.2.1') -ceq 'ex.: versions\1.2.1\). Copie') 'exemplo de pasta do LEIA-ME'

    Check ((New-ReleaseBody ("`r`n" + [char]0x2022 + " linha 1`r`n" + [char]0x2022 + " linha 2`r`n`r`n")) -ceq ([char]0x2022 + " linha 1`n" + [char]0x2022 + ' linha 2')) 'corpo da release: notes com LF e sem espacos nas pontas'
    $threw = $false; try { [void](New-ReleaseBody "  `r`n ") } catch { $threw = $true }
    Check $threw 'corpo da release vazio: erro'
    Check ((Get-TextSha256 "a`r`nb`r`n") -ceq (Get-TextSha256 "a`nb")) 'Get-TextSha256 ignora CRLF/LF e espacos nas pontas'

    $ct = @{ 'ThinQQ-1.2.1.zip' = 'application/zip'; 'ThinQQ.xlam' = 'application/vnd.ms-excel.addin.macroEnabled.12'; 'ThinQQ.core.xlam' = 'application/vnd.ms-excel.addin.macroEnabled.12'
        'manifest.json' = 'application/json'; 'tutorial.html' = 'text/html'; 'outro.bin' = 'application/octet-stream' }
    foreach ($k in $ct.Keys) { Check ((Get-AssetContentType $k) -ceq $ct[$k]) "Content-Type de $k" }
    $names = Get-ReleaseAssetNames '1.2.1'
    Check (($names -join '|') -ceq 'ThinQQ-1.2.1.zip|ThinQQ.xlam|ThinQQ.core.xlam|manifest.json|tutorial.html') 'nomes EXATOS dos 5 assets'
    $cu = Get-CompareUrl $R '1.2.1'
    Check ($cu.StartsWith('https://github.com/israeldias-hub/ThinQQ/compare/main...release/v1.2.1?expand=1&title=')) "URL do pull request ($cu)"
    $cuq = @{}
    foreach ($kv in ($cu.Split('?')[1].Split('&'))) { $parts = $kv.Split('=', 2); $cuq[$parts[0]] = [uri]::UnescapeDataString($parts[1]) }
    Check ($cuq['title'] -ceq 'release: ThinQQ 1.2.1' -and $cuq['body'] -match 'IMEDIATAMENTE' -and $cuq['body'] -match 'publicar-release\.bat 1\.2\.1 publicar' -and $cuq['body'] -match 'Create a merge commit') 'URL do pull request preenche titulo e descricao (codificados)'
    Check ($cu -notmatch '[\s#]' -and $cu.Length -lt 2000) 'URL do pull request sem espacos, sem # e curta'
    Check ((Get-ReleaseBaseUrl $R '1.2.1') -ceq 'https://github.com/israeldias-hub/ThinQQ/releases/download/v1.2.1') 'BaseUrl do release.ps1'

    # git credential fill: parsing sem vazar
    $cred = ConvertFrom-GitCredentialOutput "protocol=https`nhost=github.com`nusername=israeldias-hub`npassword=$sentinel`n"
    Check ($cred.UserName -ceq 'israeldias-hub' -and $cred.GetNetworkCredential().Password -ceq $sentinel) 'credencial: usuario e token lidos'
    Check (-not (($cred | Out-String) + ($cred | Format-List * | Out-String)).Contains($sentinel)) 'credencial impressa nao mostra o token'
    $s = Invoke-Scenario { ConvertFrom-GitCredentialOutput "protocol=https`nhost=github.com`nusername=$sentinel`n" }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and $s.Message -match 'username' -and -not $s.Message.Contains($sentinel)) "sem password: erro cita so as chaves ($($s.Message))"
    Check ((Get-RepoCredentialUser 'https://fulano@github.com/israeldias-hub/ThinQQ.git' $R) -ceq 'fulano') 'usuario do remoto (https://usuario@github.com)'
    Check ((Get-RepoCredentialUser 'C:\outro\repo.git' $R) -ceq 'israeldias-hub') 'sem usuario no remoto: dono do repositorio'
    Check ((Get-GitCredentialRequest 'israeldias-hub') -ceq "protocol=https`nhost=github.com`nusername=israeldias-hub`n`n") 'pedido ao git credential fill'
    # stdin de um processo real (como o do git credential fill) chega sem BOM: git hash-object --stdin de "abc`n" = SHA-1 de "blob 4`0abc`n"
    $sha1 = [Security.Cryptography.SHA1]::Create()
    $expectedBlob = [BitConverter]::ToString($sha1.ComputeHash([Text.Encoding]::ASCII.GetBytes("blob 4" + [char]0 + "abc`n"))).Replace('-', '').ToLowerInvariant()
    $sha1.Dispose()
    # o defeito so aparece com a entrada do console em UTF-8 COM preambulo (terminal UTF-8): o teste a liga ele mesmo, para reproduzir em
    # qualquer ambiente, e a devolve no fim
    $prevInEnc = $null
    try { $prevInEnc = [Console]::InputEncoding; [Console]::InputEncoding = New-Object Text.UTF8Encoding($true) } catch { Write-Host "(nao deu para trocar a codificacao do console: $($_.Exception.Message))" }
    try {
        $hb = Invoke-RpProcess (Get-RpGitExe) @('hash-object', '--stdin') -InputText "abc`n"
        Check ($hb.Code -eq 0 -and $hb.Output.Trim() -ceq $expectedBlob) "stdin do processo sem BOM, mesmo com o console em UTF-8 com preambulo (git hash-object: $($hb.Output.Trim()))"
        Check ($null -eq $prevInEnc -or [Console]::InputEncoding.GetPreamble().Length -eq 3) 'Invoke-RpProcess devolve a codificacao do console como estava'
    } finally { if ($prevInEnc) { try { [Console]::InputEncoding = $prevInEnc } catch { } } }
    # SHA-256 sem o Get-FileHash (num 5.1 aberto a partir do pwsh 7 o cmdlet pode faltar): com ele mascarado, Get-FileSha256 continua certo
    $hf = Join-Path $tmp 'hash.txt'; [IO.File]::WriteAllBytes($hf, [Text.Encoding]::ASCII.GetBytes('abc'))
    function Get-FileHash { throw 'Get-FileHash mascarado pelo teste' }
    try { $hv = Get-FileSha256 $hf } catch { $hv = "erro: $($_.Exception.Message)" } finally { Remove-Item Function:\Get-FileHash }
    Check ($hv -ceq 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad') "Get-FileSha256 nao depende do Get-FileHash ($hv)"
    $libText = [IO.File]::ReadAllText((Join-Path $repo 'tools\lib\Release-Publish.ps1'))
    Check ($libText -match "GIT_TERMINAL_PROMPT = '0'" -and $libText -match "GCM_INTERACTIVE = 'never'") 'git credential fill sem prompt (GIT_TERMINAL_PROMPT=0, GCM_INTERACTIVE=never)'
    Check ($libText -match 'SecurityProtocolType\]::Tls12') 'TLS 1.2 forcado no Windows PowerShell 5.1'
    Check ($libText -match 'robocopy \S+ \S+ /MIR /XJ ') 'copia do build com robocopy /XJ (nao segue juncoes)'
    Check ((Protect-RpSecret "x Bearer $sentinel y $sentinel" $sentinel) -notmatch $sentinel) 'Protect-RpSecret troca o segredo'
    Check ((Protect-RpSecret 'Authorization: Bearer ghp_abcdefghijklmnopqrstuvwxyz0123' '') -notmatch 'ghp_') 'Protect-RpSecret troca um token com cara de token'

    # ================================================================ backend falso: Publish-GitHubRelease
    $sd = Join-Path $tmp 'state-1.2.1'
    New-Item -ItemType Directory -Force $sd | Out-Null
    foreach ($n in $names) { [IO.File]::WriteAllText((Join-Path $sd $n), "conteudo ficticio de $n`n") }
    $state = New-ReleaseState '1.2.1' $R $sd ('a' * 40) ('b' * 40)
    Write-ReleaseState $state $sd
    $stRead = Read-ReleaseState $sd '1.2.1'
    Check ($stRead.version -ceq '1.2.1' -and @($stRead.files).Count -eq 5) 'state.json: ida e volta'
    $tgt = 'c' * 40
    $pubArgs = @{ Repo = $R; Version = '1.2.1'; StateDir = $sd; State = $stRead; Body = 'notas da versao'; TargetSha = $tgt; Credential = $cred; Transport = $fakeTransport }

    Reset-FakeGitHub
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs -Draft }
    Check (-not $s.Failed) "rascunho (-Draft): sem erro ($($s.Message))"
    $post = Get-FakeRequests 'POST' '/releases$'
    Check ($post.Count -eq 1) 'rascunho: um POST /releases'
    $pb = $post[0].BodyText | ConvertFrom-Json
    Check ($pb.tag_name -ceq 'v1.2.1' -and $pb.target_commitish -ceq $tgt -and $pb.name -ceq 'ThinQQ 1.2.1' -and $pb.body -ceq 'notas da versao' -and $pb.draft -eq $true -and $pb.prerelease -eq $false) 'rascunho: tag, alvo, nome, corpo, draft=true, prerelease=false'
    $up = Get-FakeRequests 'POST' '/assets$'
    Check ($up.Count -eq 5 -and @($up | Where-Object { $_.Host -ne 'uploads.github.com' }).Count -eq 0) 'os 5 assets vao para uploads.github.com'
    Check ((@($up | ForEach-Object { [uri]::UnescapeDataString(($_.Query -replace '^\?name=', '')) }) -join '|') -ceq ($names -join '|')) 'assets com os nomes exatos'
    Check ((@($up | ForEach-Object { $_.ContentType }) -join '|') -ceq ((@($names | ForEach-Object { Get-AssetContentType $_ })) -join '|')) 'Content-Type de cada asset'
    Check ((Get-FakeRequests 'PATCH' '.').Count -eq 0) '-Draft: nada e publicado (sem PATCH)'
    Check ($script:gh.Releases[0].draft -eq $true -and $s.Result.Published -eq $false) '-Draft: a release continua rascunho'
    Check (@($s.Result.Assets | Where-Object { $_.Method -eq 'digest' }).Count -eq 5) 'conferencia pelo campo digest da API'
    Check (@($script:gh.Requests | Where-Object { $_.Auth -cne "Bearer $sentinel" }).Count -eq 0) 'toda chamada leva Authorization: Bearer <token> no cabecalho'
    Check (@($script:gh.Requests | Where-Object { $RpAllowedHosts -notcontains $_.Host }).Count -eq 0) 'so api.github.com e uploads.github.com'
    Check (@($script:gh.Requests | Where-Object { $_.Uri.Contains($sentinel) -or $_.BodyText.Contains($sentinel) }).Count -eq 0) 'o token nunca vai na URL nem no corpo'

    Reset-FakeGitHub
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs }
    $patch = Get-FakeRequests 'PATCH' '/releases/\d+$'
    Check (-not $s.Failed -and $patch.Count -eq 1) "publicacao: um PATCH ($($s.Message))"
    $pp = $patch[0].BodyText | ConvertFrom-Json
    Check ($pp.draft -eq $false -and $pp.make_latest -ceq 'true') 'publicacao: draft=false, make_latest=true'
    Check ($s.Result.Published -and $s.Result.HtmlUrl -ceq "https://github.com/$R/releases/tag/v1.2.1" -and $script:gh.Releases[0].draft -eq $false) 'publicacao: release publicada e link final'

    Reset-FakeGitHub; $script:gh.DigestMode = 'bad'
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs }
    Check ($s.Failed -and $s.Kind -eq 'falha' -and $s.Message -match 'conferencia' -and $s.Message -match 'ThinQQ\.core\.xlam' -and $s.Message -match 'rascunho') "digest divergente: falha clara ($($s.Message))"
    Check ((Get-FakeRequests 'PATCH' '.').Count -eq 0 -and $script:gh.Releases[0].draft -eq $true) 'digest divergente: NAO publica e deixa o rascunho'

    Reset-FakeGitHub; $script:gh.DigestMode = 'none'
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs }
    $dl = Get-FakeRequests 'GET' '/releases/assets/\d+$'
    Check (-not $s.Failed -and $dl.Count -eq 5 -and @($dl | Where-Object { $_.Accept -cne 'application/octet-stream' }).Count -eq 0) "sem digest: baixa cada asset pela API (Accept octet-stream) e confere ($($s.Message))"
    Check (@($s.Result.Assets | Where-Object { $_.Method -eq 'download' }).Count -eq 5 -and $s.Result.Published) 'sem digest: conferido por download e publicado'
    Reset-FakeGitHub; $script:gh.DigestMode = 'none-bad'
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs }
    Check ($s.Failed -and $s.Message -match 'tutorial\.html' -and (Get-FakeRequests 'PATCH' '.').Count -eq 0) 'sem digest e conteudo diferente: nao publica'

    Reset-FakeGitHub
    [void]$script:gh.Releases.Add([ordered]@{ id = 7; tag_name = 'v1.2.1'; name = 'ThinQQ 1.2.1'; draft = $false; html_url = "https://github.com/$R/releases/tag/v1.2.1"; assets = @() })
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and $s.Message -match 'ja existe' -and (Get-FakeRequests 'POST' '.').Count -eq 0) "release ja publicada: recusa sem criar nada ($($s.Message))"
    Reset-FakeGitHub; $script:gh.Tags = @('v1.2.1')
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and $s.Message -match 'tag v1\.2\.1' -and (Get-FakeRequests 'POST' '.').Count -eq 0) 'tag ja existente: recusa'
    Reset-FakeGitHub
    [void]$script:gh.Releases.Add([ordered]@{ id = 8; tag_name = 'v1.2.1'; name = 'ThinQQ 1.2.1'; body = "notas da versao`r`n"; draft = $true; html_url = "https://github.com/$R/releases/tag/untagged-8"; assets = @() })
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs -Draft }
    Check (-not $s.Failed -and (Get-FakeRequests 'DELETE' '/releases/8$').Count -eq 1 -and (Get-FakeRequests 'POST' '/releases$').Count -eq 1) "retomada: rascunho anterior 'ThinQQ 1.2.1' apagado e recriado ($($s.Message))"
    Check ($script:gh.Releases.Count -eq 1 -and $script:gh.Releases[0].id -ne 8) 'retomada: sobra so o rascunho novo'
    # rascunho com o nosso nome, mas texto editado na pagina do GitHub: nao apaga (o trabalho de quem editou se perderia)
    Reset-FakeGitHub
    [void]$script:gh.Releases.Add([ordered]@{ id = 8; tag_name = 'v1.2.1'; name = 'ThinQQ 1.2.1'; body = 'texto editado no GitHub'; draft = $true; html_url = "https://github.com/$R/releases/tag/untagged-8"; assets = @() })
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and $s.Message -match 'editado' -and (Get-FakeRequests 'DELETE' '.').Count -eq 0 -and (Get-FakeRequests 'POST' '.').Count -eq 0) "rascunho com texto editado: recusa e nao apaga ($($s.Message))"
    # paginacao: o rascunho anterior esta na 2.a pagina de releases
    Reset-FakeGitHub
    for ($i = 1; $i -le 150; $i++) { [void]$script:gh.Releases.Add([ordered]@{ id = 1000 + $i; tag_name = "v0.0.$i"; name = "ThinQQ 0.0.$i"; body = 'x'; draft = $false; html_url = 'https://github.com/x'; assets = @() }) }
    [void]$script:gh.Releases.Insert(120, [ordered]@{ id = 8; tag_name = 'v1.2.1'; name = 'ThinQQ 1.2.1'; body = 'notas da versao'; draft = $true; html_url = 'https://github.com/x'; assets = @() })
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs -Draft }
    Check (-not $s.Failed -and (Get-FakeRequests 'DELETE' '/releases/8$').Count -eq 1 -and @(Get-FakeRequests 'GET' '/releases$' | Where-Object { $_.Query -match 'page=2' }).Count -ge 1) "paginacao: acha o rascunho anterior na 2.a pagina ($($s.Message))"
    Check (@(Get-FakeRequests 'GET' '/releases$' | Where-Object { $_.Query -notmatch 'per_page=100' }).Count -eq 0) 'paginacao: sempre per_page=100'
    Reset-FakeGitHub
    [void]$script:gh.Releases.Add([ordered]@{ id = 9; tag_name = 'v1.2.1'; name = 'Rascunho manual'; draft = $true; html_url = 'https://github.com/x'; assets = @() })
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and (Get-FakeRequests 'DELETE' '.').Count -eq 0) 'rascunho de outra pessoa (outro nome): recusa e nao apaga'

    foreach ($st in 401, 403, 422) {
        Reset-FakeGitHub
        $script:gh.ForceStatus["POST /repos/$R/releases"] = $st
        $script:gh.ForceMessage["POST /repos/$R/releases"] = "Bad credentials; recebido Authorization: Bearer $sentinel e $sentinel"
        $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs }
        Check ($s.Failed -and $s.Status -eq $st -and $s.Message -match "$st" -and -not $s.Output.Contains($sentinel)) "erro $st da API: mensagem com o codigo e SEM o token ($($s.Message))"
    }
    Check ($s.Message -match 'recusou') 'erro 422: dica em portugues'
    Reset-FakeGitHub
    $script:gh.ForceStatus["POST /repos/$R/releases"] = 401; $script:gh.ForceMessage["POST /repos/$R/releases"] = 'Bad credentials'
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs }
    Check ($s.Message -match 'login') 'erro 401: orienta a refazer o login'

    Reset-FakeGitHub; $script:gh.FailUploadOn = 'ThinQQ.core.xlam'
    $s = Invoke-Scenario { Publish-GitHubRelease @pubArgs }
    Check ($s.Failed -and $s.Message -match 'ThinQQ\.core\.xlam' -and $s.Message -match 'rascunho' -and $s.Message -match 'publicar-release\.bat 1\.2\.1 publicar') "upload que falha no meio: diz o arquivo, o rascunho e como retomar ($($s.Message))"
    Check ((Get-FakeRequests 'POST' '/assets$').Count -eq 3 -and $script:gh.Releases[0].assets.Count -eq 2 -and (Get-FakeRequests 'PATCH' '.').Count -eq 0) 'upload que falha no meio: para no 3.o arquivo e nao publica'

    Reset-FakeGitHub
    foreach ($bad in 'https://example.com/repos/x', 'http://api.github.com/repos/x', 'https://api.github.com.example.com/x', 'https://objects.githubusercontent.com/x') {
        $s = Invoke-Scenario { Invoke-GitHubApi -Method GET -Uri $bad -Credential $cred -Transport $fakeTransport }
        Check ($s.Failed -and $s.Message -match 'recusada') "host/esquema fora da lista recusado: $bad"
    }
    Check ($script:gh.Requests.Count -eq 0) 'host recusado: o transporte nem e chamado'
    $boom = { param($Request) throw ('rede caiu; cabecalho ' + $Request.Headers['Authorization']) }
    $s = Invoke-Scenario { Invoke-GitHubApi -Method GET -Uri "https://api.github.com/repos/$R/releases" -Credential $cred -Transport $boom }
    Check ($s.Failed -and $s.Message -match 'rede' -and -not $s.Output.Contains($sentinel)) 'erro de rede com o cabecalho na mensagem: token removido'
    $redir = { param($Request) return @{ Status = 302; Content = '' } }
    $s = Invoke-Scenario { Invoke-GitHubApi -Method GET -Uri "https://api.github.com/repos/$R/releases/assets/1" -Accept 'application/octet-stream' -Raw -Credential $cred -Transport $redir }
    Check ($s.Failed -and $s.Message -match 'redirecionamento') 'redirecionamento para outro host: recusado, nao seguido'

    # ================================================================ transporte PADRAO (Invoke-WebRequest) contra um HttpListener local
    # No pwsh 7 uma resposta HTTP de erro virava HttpResponseException, cuja RequestMessage guarda "Authorization: Bearer <token>", e o
    # registro ficava em $Error (Get-Error e Format-List * -Force o mostravam). O token (sentinela) nao pode aparecer em $Error, Get-Error,
    # transcricao nem nas mensagens. O host http://127.0.0.1 so e aceito com a chave de teste $script:RpTestAllowLoopback.
    $tlp = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0); $tlp.Start(); $lport = ([Net.IPEndPoint]$tlp.LocalEndpoint).Port; $tlp.Stop()
    $tlp = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0); $tlp.Start(); $closedPort = ([Net.IPEndPoint]$tlp.LocalEndpoint).Port; $tlp.Stop()
    $listener = New-Object Net.HttpListener
    $listener.Prefixes.Add("http://127.0.0.1:$lport/")
    $listener.Start()
    $seen = [hashtable]::Synchronized(@{ Auth = (New-Object Collections.ArrayList) })
    $srv = [powershell]::Create()
    [void]$srv.AddScript({
            param($l, $seen)
            while ($l.IsListening) {
                try { $ctx = $l.GetContext() } catch { break }
                [void]$seen.Auth.Add([string]$ctx.Request.Headers['Authorization'])
                $code = 200
                $path = $ctx.Request.Url.AbsolutePath.Trim('/')
                if ($path -match '^\d{3}$') { $code = [int]$path }
                $b = [Text.Encoding]::UTF8.GetBytes('{"message":"resposta ' + $code + '"}')
                $ctx.Response.StatusCode = $code
                $ctx.Response.ContentType = 'application/json'
                $ctx.Response.ContentLength64 = $b.Length
                $ctx.Response.OutputStream.Write($b, 0, $b.Length)
                $ctx.Response.Close()
            }
        }).AddArgument($listener).AddArgument($seen)
    $srvHandle = $srv.BeginInvoke()
    $transcript = Join-Path $tmp 'transcricao.txt'
    $transcribing = $false
    try {
        try { Start-Transcript -LiteralPath $transcript -Force | Out-Null; $transcribing = $true } catch { Write-Host "(transcricao indisponivel neste host: $($_.Exception.Message))" }
        $global:Error.Clear()
        $s = Invoke-Scenario { Invoke-GitHubApi -Method GET -Uri "http://127.0.0.1:$lport/200" -Credential $cred }
        Check ($s.Failed -and $s.Message -match 'recusada') 'sem a chave de teste, http://127.0.0.1 e recusado'
        foreach ($code in 404, 401, 403, 200) {
            $req = @{ Method = 'GET'; Uri = "http://127.0.0.1:$lport/$code"; Headers = @{ Accept = 'application/json'; Authorization = "Bearer $sentinel" }; Body = $null; ContentType = ''; UserAgent = 'teste' }
            $s = Invoke-Scenario { & $script:RpDefaultTransport $req }
            Check (-not $s.Failed -and [int]$s.Result.Status -eq $code) "transporte padrao: HTTP $code volta como status, sem excecao ($($s.Message))"
        }
        $script:RpTestAllowLoopback = $true
        try {
            foreach ($code in 404, 401) {
                $s = Invoke-Scenario { Invoke-GitHubApi -Method GET -Uri "http://127.0.0.1:$lport/$code" -Credential $cred }
                Check ($s.Failed -and $s.Status -eq $code -and -not $s.Output.Contains($sentinel)) "Invoke-GitHubApi + transporte padrao: $code vira erro com status, sem o token"
            }
            $s = Invoke-Scenario { Invoke-GitHubApi -Method GET -Uri "http://127.0.0.1:$lport/200" -Credential $cred }
            Check (-not $s.Failed -and $s.Result.message -eq 'resposta 200') 'Invoke-GitHubApi + transporte padrao: 200 devolve o JSON'
            $s = Invoke-Scenario { Invoke-GitHubApi -Method GET -Uri "http://127.0.0.1:$closedPort/x" -Credential $cred }
            Check ($s.Failed -and $s.Message -match 'rede' -and -not $s.Output.Contains($sentinel)) "porta fechada: erro de rede sem o token ($($s.Message))"
            $req = @{ Method = 'GET'; Uri = "http://127.0.0.1:$closedPort/x"; Headers = @{ Authorization = "Bearer $sentinel" }; Body = $null; ContentType = ''; UserAgent = 'teste' }
            $s = Invoke-Scenario { & $script:RpDefaultTransport $req }
            Check ($s.Failed -and -not $s.Output.Contains($sentinel)) 'transporte padrao, porta fechada: a excecao nao carrega o token'
        } finally { $script:RpTestAllowLoopback = $false }
        Check (@($seen.Auth | Where-Object { $_ -ceq "Bearer $sentinel" }).Count -ge 6) 'o servidor local recebeu o token no cabecalho (ele estava la)'
        Check ($global:Error.Count -gt 0) "`$Error tem registros para conferir ($($global:Error.Count))"
        $errText = ($global:Error | Format-List * -Force | Out-String) +
            (@($global:Error | ForEach-Object { if ($_.Exception) { $_.Exception | Format-List * -Force | Out-String }; if ($_.TargetObject) { [string]$_.TargetObject } }) -join "`n")
        Check (-not $errText.Contains($sentinel)) 'o token nao aparece em $Error | Format-List * -Force (nem na excecao nem no TargetObject)'
        if (Get-Command Get-Error -ErrorAction SilentlyContinue) {
            $ge = (Get-Error -Newest ([Math]::Max(1, $global:Error.Count)) | Out-String -Width 400)
            Check (-not $ge.Contains($sentinel)) 'o token nao aparece em Get-Error (pwsh 7)'
        }
    } finally {
        if ($transcribing) { try { Stop-Transcript | Out-Null } catch { } }
        $listener.Stop(); $listener.Close()
        [void]$srvHandle.AsyncWaitHandle.WaitOne(5000)
        $srv.Dispose()
    }
    if ($transcribing) { Check (-not ([IO.File]::ReadAllText($transcript)).Contains($sentinel)) 'o token nao aparece na transcricao' }
    $entryText = [IO.File]::ReadAllText((Join-Path $repo 'tools\Publish-Release.ps1')) + [IO.File]::ReadAllText((Join-Path $repo 'publicar-release.bat'))
    $libAssign = @([regex]::Matches([IO.File]::ReadAllText((Join-Path $repo 'tools\lib\Release-Publish.ps1')), 'RpTestAllowLoopback\s*=\s*(\S+)') | ForEach-Object { $_.Groups[1].Value })
    Check ($entryText -notmatch 'RpTestAllowLoopback|127\.0\.0\.1' -and ($libAssign -join ',') -ceq '$false') 'a chave de teste do loopback nao e alcancavel pelo .bat nem pelo script (a lib so a zera)'

    # ================================================================ repositorio git descartavel (remoto bare local)
    $remote = Join-Path $tmp 'origem.git'
    $work = Join-Path $tmp 'trabalho'
    [void](Invoke-RpGitChecked $tmp @('init', '-q', '--bare', '-b', 'main', $remote) 'init bare')
    [void](Invoke-RpGitChecked $tmp @('init', '-q', '-b', 'main', $work) 'init')
    foreach ($kv in @(@('user.name', 'Teste ThinQQ'), @('user.email', 'teste@example.invalid'), @('commit.gpgsign', 'false'), @('tag.gpgsign', 'false'))) {
        [void](Invoke-RpGitChecked $work @('config', $kv[0], $kv[1]) 'config')
    }
    function Write-WorkFile([string]$rel, [string]$text) {
        $p = Join-Path $work $rel
        New-Item -ItemType Directory -Force (Split-Path $p) | Out-Null
        Write-RpText $p $text
    }
    function Invoke-WorkGit([string[]]$a) { return (Invoke-RpGitChecked $work $a 'git do teste') }
    $clWork = "# Changelog`n`n$unrel`n`n### Adicionado`n- Algo novo.`n`n## [1.2.0] - 2026-10-12`n`n### Adicionado`n- Primeira.`n"
    Write-WorkFile 'version.json' ('{ "loader": "1.2.0", "core": "1.2.0" }' + "`n")
    Write-WorkFile 'CHANGELOG.md' $clWork
    Write-WorkFile 'release\notes.txt' ([char]0x2022 + " nota antiga`n")
    Write-WorkFile 'release\manifest.json' ('{ "coreVersion": "1.2.0" }' + "`n")
    Write-WorkFile 'docs\tutorial\tutorial.html' "<html><body><p>ThinQQ 1.2.0. Duvidas: issue.</p></body></html>`n"
    Write-WorkFile 'installer\LEIA-ME.txt' "ex.: versions\1.2.0\). Copie`n"
    $leiaBom = Join-Path $work 'installer\LEIA-ME.txt'
    [IO.File]::WriteAllText($leiaBom, (Read-RpText $leiaBom), (New-Object Text.UTF8Encoding($true)))   # o LEIA-ME real tem BOM (vai no ZIP)
    Write-WorkFile '.gitignore' "dist/`n"
    [void](Invoke-WorkGit @('add', '-A')); [void](Invoke-WorkGit @('commit', '-q', '-m', 'inicial')); [void](Invoke-WorkGit @('tag', 'v1.2.0'))
    [void](Invoke-WorkGit @('remote', 'add', 'origin', $remote)); [void](Invoke-WorkGit @('push', '-q', '-u', 'origin', 'main', '--tags'))

    $script:buildCalls = 0; $script:confirmCalls = 0; $script:browserUrls = @()
    $fakeBuild = {
        param($Ctx)
        $script:buildCalls++
        if ($Ctx.BaseUrl -cne "https://github.com/$R/releases/download/v$($Ctx.Version)") { throw "BaseUrl errada: $($Ctx.BaseUrl)" }
        foreach ($n in (Get-ReleaseAssetNames $Ctx.Version)) {
            $c = "conteudo ficticio de $n"
            if ($n -eq 'manifest.json') { $c = '{ "coreVersion": "' + $Ctx.Version + '" }' }
            [IO.File]::WriteAllText((Join-Path $Ctx.OutDir $n), $c + "`n")
        }
    }
    $okChecks = { param($root) return @() }
    $yes = { param($q) $script:confirmCalls++; return $true }
    $no = { param($q) $script:confirmCalls++; return $false }
    $browser = { param($u) $script:browserUrls += $u }
    $prep = @{ RepoRoot = $work; Version = '1.2.1'; Repo = $R; BuildStep = $fakeBuild; ChecksStep = $okChecks; ConfirmStep = $yes; BrowserStep = $browser; Today = '2026-10-13'; WorkRoot = $tmp }
    function Get-RepoSnapshot {
        $b = (Invoke-RpGit $work @('branch', '--list', '--format=%(refname:short)')).Output.Trim()
        $h = (Invoke-RpGit $work @('rev-parse', 'HEAD')).Output.Trim()
        $st = (Invoke-RpGit $work @('status', '--porcelain')).Output.Trim()
        $cur = (Invoke-RpGit $work @('rev-parse', '--abbrev-ref', 'HEAD')).Output.Trim()
        $rh = (Invoke-RpGit $work @('ls-remote', '--heads', 'origin')).Output.Trim()
        $vjs = Get-TextSha256 (Read-RpText (Join-Path $work 'version.json'))
        $cls = Get-TextSha256 (Read-RpText (Join-Path $work 'CHANGELOG.md'))
        return "$cur|$h|$b|$st|$rh|$vjs|$cls|" + (Test-Path (Join-Path $work 'dist'))
    }
    function Test-Refused($scn, [string]$pattern, [string]$label, [string]$before) {
        Check ($scn.Failed -and $scn.Kind -eq 'precondicao' -and $scn.Message -match $pattern) "Prepare recusa: $label ($($scn.Message))"
        Check ((Get-RepoSnapshot) -ceq $before) "Prepare recusado ($label) nao mudou nada no repositorio"
    }

    $snap = Get-RepoSnapshot
    $s = Invoke-Scenario { Invoke-ReleasePrepare @prep }
    Test-Refused $s 'notes\.txt nao mudou.*v1\.2\.0' 'notas inalteradas desde a ultima tag' $snap
    Write-WorkFile 'release\notes.txt' ([char]0x2022 + " nota nova da 1.2.1`n" + [char]0x2022 + " outra linha`n")
    [void](Invoke-WorkGit @('commit', '-q', '-am', 'notas')); [void](Invoke-WorkGit @('push', '-q', 'origin', 'main'))

    $snap = Get-RepoSnapshot
    Write-WorkFile 'sujo.txt' 'x'
    $snapDirty = Get-RepoSnapshot
    $s = Invoke-Scenario { Invoke-ReleasePrepare @prep }
    Test-Refused $s 'nao commitadas' 'arvore suja' $snapDirty
    Remove-Item -LiteralPath (Join-Path $work 'sujo.txt')
    [void](Invoke-WorkGit @('checkout', '-q', '-b', 'outra'))
    $snapB = Get-RepoSnapshot
    $s = Invoke-Scenario { Invoke-ReleasePrepare @prep }
    Test-Refused $s "branch atual e 'outra'" 'branch errada' $snapB
    [void](Invoke-WorkGit @('checkout', '-q', 'main')); [void](Invoke-WorkGit @('branch', '-q', '-D', 'outra'))
    Write-WorkFile 'extra.txt' "x`n"
    [void](Invoke-WorkGit @('add', 'extra.txt')); [void](Invoke-WorkGit @('commit', '-q', '-m', 'extra')); [void](Invoke-WorkGit @('push', '-q', 'origin', 'main'))
    [void](Invoke-WorkGit @('reset', '-q', '--hard', 'HEAD~1'))
    $snapBehind = Get-RepoSnapshot
    $s = Invoke-Scenario { Invoke-ReleasePrepare @prep }
    Test-Refused $s 'atrasada' 'main atrasada em relacao a origin/main' $snapBehind
    [void](Invoke-WorkGit @('reset', '-q', '--hard', 'origin/main'))
    Write-WorkFile 'local.txt' "x`n"
    [void](Invoke-WorkGit @('add', 'local.txt')); [void](Invoke-WorkGit @('commit', '-q', '-m', 'so local'))
    $snapAhead = Get-RepoSnapshot
    $s = Invoke-Scenario { Invoke-ReleasePrepare @prep }
    Test-Refused $s 'nao estao na origin/main' 'main com commits locais' $snapAhead
    [void](Invoke-WorkGit @('reset', '-q', '--hard', 'origin/main'))
    $snap = Get-RepoSnapshot
    foreach ($bad in @(@('1.2.0', 'maior'), @('1.1.9', 'maior'), @('abc', 'invalida'), @('1.2', 'invalida'), @('1.2.1-rc1', 'x\.y\.z'))) {
        $p2 = $prep.Clone(); $p2.Version = $bad[0]
        $s = Invoke-Scenario { Invoke-ReleasePrepare @p2 }
        Test-Refused $s $bad[1] "versao $($bad[0])" $snap
    }
    $p2 = $prep.Clone(); $p2.ChecksStep = { param($root) return @('Test-Docs.ps1 (codigo 1): not ok - algo') }
    $s = Invoke-Scenario { Invoke-ReleasePrepare @p2 }
    Test-Refused $s 'Test-Docs' 'Test-Docs falhando' $snap
    Check ($script:buildCalls -eq 0) 'nenhuma recusa chegou a construir'
    # [Nao lancado] vazia: aviso e pergunta; "nao" cancela
    Write-WorkFile 'CHANGELOG.md' "# Changelog`n`n$unrel`n`n## [1.2.0] - 2026-10-12`n`n### Adicionado`n- Primeira.`n"
    [void](Invoke-WorkGit @('commit', '-q', '-am', 'changelog vazio')); [void](Invoke-WorkGit @('push', '-q', 'origin', 'main'))
    $snap = Get-RepoSnapshot
    $p2 = $prep.Clone(); $p2.ConfirmStep = $no
    $s = Invoke-Scenario { Invoke-ReleasePrepare @p2 }
    Test-Refused $s 'Cancelado' 'CHANGELOG sem entradas e resposta nao' $snap
    Check ($script:confirmCalls -eq 1 -and $s.Output -match 'AVISO') 'CHANGELOG sem entradas: avisa e pergunta'
    Write-WorkFile 'CHANGELOG.md' $clWork
    [void](Invoke-WorkGit @('commit', '-q', '-am', 'changelog de volta')); [void](Invoke-WorkGit @('push', '-q', 'origin', 'main'))

    # -DryRun: nada gravado, nada construido, nada enviado
    $snap = Get-RepoSnapshot
    $script:confirmCalls = 0
    $s = Invoke-Scenario { Invoke-ReleasePrepare @prep -DryRun }
    Check (-not $s.Failed -and $s.Result.DryRun -and $s.Output -match '\[ENSAIO\]') "Prepare -DryRun: mostra o plano ($($s.Message))"
    Check ((Get-RepoSnapshot) -ceq $snap -and $script:buildCalls -eq 0 -and $script:browserUrls.Count -eq 0) 'Prepare -DryRun: nao grava, nao constroi, nao envia, nao abre o navegador'

    # build que falha: tudo desfeito
    $snap = Get-RepoSnapshot
    $p2 = $prep.Clone(); $p2.BuildStep = { param($Ctx) [IO.File]::WriteAllText((Join-Path $Ctx.OutDir 'parcial.txt'), 'x'); throw 'build quebrou de proposito' }
    $s = Invoke-Scenario { Invoke-ReleasePrepare @p2 }
    Check ($s.Failed -and $s.Message -match 'build quebrou') 'build que falha: erro repassado'
    $after = Get-RepoSnapshot
    Check ($after -ceq $snap -and -not (Test-Path (Join-Path $work 'dist\release-1.2.1'))) ("build que falha: volta para a main, apaga a branch e a pasta, restaura os arquivos" + $(if ($after -cne $snap) { " (antes: $snap; depois: $after)" } else { '' }))

    # arquivo da versao alterado por fora enquanto o build rodava: nada e commitado e a alteracao NAO e descartada
    $snapMain = Invoke-WorkGit @('rev-parse', 'main')
    $p2 = $prep.Clone()
    $p2.BuildStep = { param($Ctx) & $fakeBuild $Ctx; $cp = Join-Path $Ctx.RepoRoot 'CHANGELOG.md'; Write-RpText $cp ((Read-RpText $cp) + "- editado por fora`n") }
    $s = Invoke-Scenario { Invoke-ReleasePrepare @p2 }
    Check ($s.Failed -and $s.Message -match 'CHANGELOG\.md' -and $s.Message -match 'alterado') "arquivo editado durante a preparacao: recusa o commit ($($s.Message))"
    Check ((Invoke-WorkGit @('rev-parse', '--abbrev-ref', 'HEAD')) -ceq 'main' -and (Invoke-WorkGit @('rev-parse', 'main')) -ceq $snapMain -and (Invoke-RpGit $work @('rev-parse', '--verify', '--quiet', 'refs/heads/release/v1.2.1')).Code -ne 0) 'arquivo editado: volta para a main, sem commit e sem a branch'
    Check ((Read-RpText (Join-Path $work 'CHANGELOG.md')) -match 'editado por fora' -and $s.Output -match 'nao foi descartad') 'arquivo editado: a alteracao de quem editou fica (e e dito)'
    Check ((Read-RpText (Join-Path $work 'version.json')) -match '"1\.2\.0"' -and -not (Test-Path (Join-Path $work 'dist\release-1.2.1'))) 'arquivo editado: os outros arquivos voltam e a pasta sai'
    [void](Invoke-WorkGit @('checkout', '--', 'CHANGELOG.md'))
    # falha depois do commit e antes do state.json: orientacao de como recomecar
    $p2 = $prep.Clone()
    $p2.BuildStep = { param($Ctx) & $fakeBuild $Ctx; New-Item -ItemType Directory -Force (Join-Path $Ctx.OutDir 'state.json') | Out-Null }
    $s = Invoke-Scenario { Invoke-ReleasePrepare @p2 }
    Check ($s.Failed -and $s.Output -match 'git checkout main' -and $s.Output -match 'git branch -D release/v1\.2\.1' -and $s.Output -match 'release-1\.2\.1') "falha entre o commit e o state.json: diz como recomecar ($($s.Message))"
    [void](Invoke-WorkGit @('checkout', '-q', 'main')); [void](Invoke-WorkGit @('branch', '-q', '-D', 'release/v1.2.1'))
    Remove-Item -LiteralPath (Join-Path $work 'dist') -Recurse -Force
    Check ((Get-RepoSnapshot) -ceq $snap) 'depois de recomecar como orientado, o repositorio volta ao estado anterior'

    # caminho feliz
    $script:buildCalls = 0
    $mainBefore = (Invoke-WorkGit @('rev-parse', 'main'))
    $s = Invoke-Scenario { Invoke-ReleasePrepare @prep }
    Check (-not $s.Failed) "Prepare: sem erro ($($s.Message))"
    Check ($script:buildCalls -eq 1) 'Prepare: construiu uma vez'
    Check ((Invoke-WorkGit @('rev-parse', '--abbrev-ref', 'HEAD')) -ceq 'release/v1.2.1') 'Prepare: branch release/v1.2.1'
    Check ((Invoke-WorkGit @('log', '-1', '--format=%s')) -ceq 'release: ThinQQ 1.2.1') "Prepare: commit 'release: ThinQQ 1.2.1'"
    Check ((Invoke-WorkGit @('rev-parse', 'HEAD~1')) -ceq $mainBefore) 'Prepare: um commit sobre a main'
    $vjAfter = (Read-RpText (Join-Path $work 'version.json')) | ConvertFrom-Json
    Check ($vjAfter.loader -ceq '1.2.1' -and $vjAfter.core -ceq '1.2.1') 'Prepare: version.json (loader e core)'
    $clAfter = Read-RpText (Join-Path $work 'CHANGELOG.md')
    Check ($clAfter.Replace("`r`n", "`n").Contains("$unrel`n`n## [1.2.1] - 2026-10-13`n`n### Adicionado`n- Algo novo.")) 'Prepare: CHANGELOG com [1.2.1] - data e [Nao lancado] nova'
    Check ((Read-RpText (Join-Path $work 'docs\tutorial\tutorial.html')) -match '<p>ThinQQ 1\.2\.1\. ') 'Prepare: rodape do tutorial'
    Check ((Read-RpText (Join-Path $work 'installer\LEIA-ME.txt')) -match 'versions\\1\.2\.1\\') 'Prepare: exemplo do LEIA-ME'
    $leiaBytes = [IO.File]::ReadAllBytes((Join-Path $work 'installer\LEIA-ME.txt'))
    Check ($leiaBytes.Length -gt 3 -and $leiaBytes[0] -eq 0xEF -and $leiaBytes[1] -eq 0xBB -and $leiaBytes[2] -eq 0xBF) 'Prepare: o LEIA-ME continua com BOM'
    $clBytes = [IO.File]::ReadAllBytes((Join-Path $work 'CHANGELOG.md'))
    Check (-not ($clBytes[0] -eq 0xEF -and $clBytes[1] -eq 0xBB)) 'Prepare: o CHANGELOG continua sem BOM'
    $outDir = Join-Path $work 'dist\release-1.2.1'
    Check ((Get-TextSha256 (Invoke-WorkGit @('show', 'HEAD:release/manifest.json'))) -ceq (Get-TextSha256 (Read-RpText (Join-Path $outDir 'manifest.json')))) 'Prepare: release/manifest.json = manifest.json construido (no commit)'
    $changed = @((Invoke-WorkGit @('diff', '--name-only', 'HEAD~1', 'HEAD')) -split "\r?\n" | Sort-Object)
    Check (($changed -join ',') -ceq 'CHANGELOG.md,docs/tutorial/tutorial.html,installer/LEIA-ME.txt,release/manifest.json,version.json') "Prepare: o commit tem so os arquivos da versao ($($changed -join ','))"
    Check ((Invoke-WorkGit @('status', '--porcelain')) -eq '') 'Prepare: arvore limpa depois do commit (dist ignorado)'
    $stPrep = Read-ReleaseState $outDir '1.2.1'
    Check ($stPrep.prepareCommit -ceq (Invoke-WorkGit @('rev-parse', 'HEAD')) -and $stPrep.sourceMain -ceq $mainBefore -and $stPrep.repo -ceq $R) 'Prepare: state.json com o commit de preparacao e a main de origem'
    Check ((Invoke-RpGit $work @('ls-remote', '--heads', 'origin', 'release/v1.2.1')).Output -match [regex]::Escape($stPrep.prepareCommit)) 'Prepare: branch enviada ao remoto'
    Check ((Invoke-WorkGit @('rev-parse', 'origin/main')) -ceq $mainBefore) 'Prepare: a main do remoto nao foi tocada'
    Check ($script:browserUrls.Count -eq 1 -and $script:browserUrls[0] -ceq (Get-CompareUrl $R '1.2.1')) 'Prepare: abre a pagina do pull request'
    Check ($s.Output -match 'publicar-release\.bat 1\.2\.1 publicar') 'Prepare: imprime o proximo passo'
    Check ($s.Output -match 'IMEDIATAMENTE' -and $s.Output -match 'git checkout main' -and $s.Output -match 'git pull') 'Prepare: pede a etapa publicar logo depois do merge e lembra de voltar para a main'

    # ================================================================ Publish com git
    $script:credCalls = 0
    $credRunner = { param($inputText) $script:credCalls++; $script:credInput = $inputText; return @{ Code = 0; Output = "protocol=https`nhost=github.com`nusername=israeldias-hub`npassword=$sentinel`n"; Error = '' } }
    $pubGit = @{ RepoRoot = $work; Version = '1.2.1'; Repo = $R; Transport = $fakeTransport; CredentialRunner = $credRunner; NoBrowser = $true }
    Reset-FakeGitHub
    $s = Invoke-Scenario { Invoke-ReleasePublish @pubGit -Draft }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and $s.Message -match 'mescle o pull request') "Publish antes do merge: recusa ($($s.Message))"
    Check ($script:credCalls -eq 0 -and $script:gh.Requests.Count -eq 0) 'Publish recusado: nao leu o token nem chamou a API'
    $pubOther = $pubGit.Clone(); $pubOther.Repo = 'outra-conta/ThinQQ'
    $s = Invoke-Scenario { Invoke-ReleasePublish @pubOther -Draft }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and $s.Message -match 'outra-conta/ThinQQ' -and $s.Message -match 'israeldias-hub/ThinQQ') "Publish: -Repo diferente do state.json => recusa ($($s.Message))"
    [void](Invoke-WorkGit @('checkout', '-q', 'main'))
    # merge por squash (ou rebase): o commit de preparacao nao entra na main, mas os arquivos da versao sao os mesmos => aceito
    $mainPre = Invoke-WorkGit @('rev-parse', 'main')
    [void](Invoke-WorkGit @('merge', '-q', '--squash', 'release/v1.2.1')); [void](Invoke-WorkGit @('commit', '-q', '-m', 'release: ThinQQ 1.2.1 (#9)')); [void](Invoke-WorkGit @('push', '-q', 'origin', 'main'))
    Reset-FakeGitHub
    $s = Invoke-Scenario { Invoke-ReleasePublish @pubGit -Draft }
    Check (-not $s.Failed -and $s.Output -match 'squash') "Publish depois de um squash merge com o mesmo conteudo: aceito e dito ($($s.Message))"
    Check (((Get-FakeRequests 'POST' '/releases$')[0].BodyText | ConvertFrom-Json).target_commitish -ceq (Invoke-WorkGit @('rev-parse', 'origin/main'))) 'squash: a tag vai para a origin/main'
    [void](Invoke-WorkGit @('reset', '-q', '--hard', $mainPre)); [void](Invoke-WorkGit @('push', '-q', '-f', 'origin', 'main'))
    # squash com conteudo diferente nos arquivos da versao => recusa explicando squash/rebase e a alternativa manual
    [void](Invoke-WorkGit @('merge', '-q', '--squash', 'release/v1.2.1'))
    Write-WorkFile 'CHANGELOG.md' ((Read-RpText (Join-Path $work 'CHANGELOG.md')) + "- editado no merge`n")
    [void](Invoke-WorkGit @('commit', '-q', '-am', 'release squash editado')); [void](Invoke-WorkGit @('push', '-q', 'origin', 'main'))
    Reset-FakeGitHub
    $s = Invoke-Scenario { Invoke-ReleasePublish @pubGit -Draft }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and $s.Message -match 'squash' -and $s.Message -match 'Alternativa manual' -and $s.Message -match 'release-1\.2\.1' -and $script:gh.Requests.Count -eq 0) "squash com conteudo diferente: recusa especifica ($($s.Message))"
    [void](Invoke-WorkGit @('reset', '-q', '--hard', $mainPre)); [void](Invoke-WorkGit @('push', '-q', '-f', 'origin', 'main'))
    [void](Invoke-WorkGit @('merge', '-q', '--no-ff', '-m', 'Merge pull request #9 from israeldias-hub/release/v1.2.1', 'release/v1.2.1'))
    [void](Invoke-WorkGit @('push', '-q', 'origin', 'main'))
    $origin = Invoke-WorkGit @('rev-parse', 'origin/main')
    $tut = Join-Path $outDir 'tutorial.html'
    $script:credCalls = 0
    $tutBytes = [IO.File]::ReadAllBytes($tut)
    [IO.File]::WriteAllText($tut, 'alterado depois da preparacao')
    $s = Invoke-Scenario { Invoke-ReleasePublish @pubGit -Draft }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and $s.Message -match 'tutorial\.html' -and $script:credCalls -eq 0) 'Publish: arquivo alterado depois da preparacao => recusa'
    [IO.File]::WriteAllBytes($tut, $tutBytes)
    $s = Invoke-Scenario { Invoke-ReleasePublish @pubGit -Draft -StateDir (Join-Path $tmp 'nao-existe') }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and $s.Message -match 'state\.json') 'Publish sem state.json: recusa'
    Write-WorkFile 'release\manifest.json' ('{ "coreVersion": "9.9.9" }' + "`n")
    [void](Invoke-WorkGit @('commit', '-q', '-am', 'manifesto trocado')); [void](Invoke-WorkGit @('push', '-q', 'origin', 'main'))
    $s = Invoke-Scenario { Invoke-ReleasePublish @pubGit -Draft }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and $s.Message -match 'manifest' -and $script:credCalls -eq 0) "Publish: manifesto da origin/main diferente => recusa ($($s.Message))"
    [void](Invoke-WorkGit @('revert', '--no-edit', 'HEAD')); [void](Invoke-WorkGit @('push', '-q', 'origin', 'main'))
    # outro PR mesclado DEPOIS do da versao: a tag apontaria para um codigo diferente do que gerou os arquivos => recusa sem -AllowMainChanges
    Write-WorkFile 'depois.txt' "mudanca depois da versao`n"
    [void](Invoke-WorkGit @('add', 'depois.txt')); [void](Invoke-WorkGit @('commit', '-q', '-m', 'outro PR')); [void](Invoke-WorkGit @('push', '-q', 'origin', 'main'))
    Reset-FakeGitHub
    $s = Invoke-Scenario { Invoke-ReleasePublish @pubGit -Draft }
    Check ($s.Failed -and $s.Kind -eq 'precondicao' -and $s.Message -match 'depois\.txt' -and $s.Message -match 'AllowMainChanges' -and $script:gh.Requests.Count -eq 0) "origin/main mudou depois da preparacao: recusa ($($s.Message))"
    $s = Invoke-Scenario { Invoke-ReleasePublish @pubGit -Draft -AllowMainChanges }
    Check (-not $s.Failed -and $s.Output -match 'AVISO' -and $s.Output -match 'depois\.txt') "com -AllowMainChanges: segue com aviso ($($s.Message))"
    [void](Invoke-WorkGit @('revert', '--no-edit', 'HEAD')); [void](Invoke-WorkGit @('push', '-q', 'origin', 'main'))
    $origin = Invoke-WorkGit @('rev-parse', 'origin/main')

    Reset-FakeGitHub; $script:credCalls = 0
    $s = Invoke-Scenario { Invoke-ReleasePublish @pubGit -Draft -DryRun }
    Check (-not $s.Failed -and $s.Result.DryRun -and $s.Output -match '\[ENSAIO\]') "Publish -DryRun: mostra o plano ($($s.Message))"
    Check ($script:gh.Requests.Count -eq 0 -and $script:credCalls -eq 0) 'Publish -DryRun: nao le o token nem chama o transporte'

    Reset-FakeGitHub; $script:credCalls = 0
    $notesWork = Join-Path (Join-Path $work 'release') 'notes.txt'
    $notesSaved = [IO.File]::ReadAllBytes($notesWork)
    Write-RpText $notesWork "texto NAO commitado da arvore de trabalho`n"
    $s = Invoke-Scenario { Invoke-ReleasePublish @pubGit -Draft }
    [IO.File]::WriteAllBytes($notesWork, $notesSaved)
    Check (-not $s.Failed -and -not $s.Result.Published) "Publish -Draft completo: rascunho criado e conferido ($($s.Message))"
    Check ($script:credCalls -eq 1 -and $script:credInput -ceq (Get-GitCredentialRequest 'israeldias-hub')) 'Publish: token lido uma vez pelo git credential fill (protocol/host/username)'
    $pb = (Get-FakeRequests 'POST' '/releases$')[0].BodyText | ConvertFrom-Json
    Check ($pb.target_commitish -ceq $origin) 'Publish: target_commitish = SHA da origin/main'
    Check ($pb.body -ceq ([char]0x2022 + " nota nova da 1.2.1`n" + [char]0x2022 + ' outra linha')) 'Publish: corpo = release/notes.txt da origin/main (nao o da arvore de trabalho)'
    Check ((Get-FakeRequests 'PATCH' '.').Count -eq 0) 'Publish -Draft: sem PATCH'
    Check ($s.Output -match 'editar o texto do rascunho') 'Publish -Draft: avisa para publicar pela pagina se editar o texto (e nao rodar de novo)'
    Check ($s.Output -match 'Instalar-ThinQQ\.cmd') 'Publish: lembrete da verificacao manual'

    # ================================================================ script e .bat (processos proprios)
    $host1 = (Get-Process -Id $PID).Path
    $script1 = Join-Path $repo 'tools\Publish-Release.ps1'
    $r1 = Invoke-RpProcess $host1 @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script1, '1.2.1', 'publicar', '-DryRun', '-NoBrowser', '-RepoRoot', $work)
    [void]$script:allOutput.Append($r1.Output + $r1.Error)
    Check ($r1.Code -eq 0 -and $r1.Output -match '\[ENSAIO\]') "script: '1.2.1 publicar -DryRun' sai com 0 (codigo $($r1.Code))"
    $r2 = Invoke-RpProcess $host1 @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script1, 'abc', '-RepoRoot', $work)
    Check ($r2.Code -eq 2 -and $r2.Output -match 'invalida') "script: versao invalida sai com 2 (codigo $($r2.Code))"
    $r3 = Invoke-RpProcess $host1 @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $script1)
    Check ($r3.Code -eq 2 -and $r3.Output -match 'Uso') "script: sem versao mostra o uso e sai com 2 (codigo $($r3.Code))"
    $bat = Join-Path $repo 'publicar-release.bat'
    $batBytes = [IO.File]::ReadAllBytes($bat)
    $batText = [Text.Encoding]::ASCII.GetString($batBytes)
    Check (@($batBytes | Where-Object { $_ -gt 127 }).Count -eq 0) 'publicar-release.bat: so ASCII'
    Check ($batText.Contains("`r`n") -and -not ($batText -match "[^\r]`n")) 'publicar-release.bat: CRLF'
    foreach ($t in 'cd /d "%~dp0"', 'powershell -NoProfile -ExecutionPolicy Bypass -File tools\Publish-Release.ps1 %*', 'if not defined NOPAUSE pause', 'if defined DRYRUN', 'exit /b %RC%') {
        Check ($batText.Contains($t)) "publicar-release.bat contem: $t"
    }
    $r4 = Invoke-RpProcess $env:ComSpec @('/c', $bat) -Environment @{ NOPAUSE = '1' }
    Check ($r4.Code -eq 2 -and $r4.Output -match 'Uso') "publicar-release.bat sem argumentos (NOPAUSE): mostra o uso e devolve 2 (codigo $($r4.Code))"

    # ================================================================ o token nunca aparece
    Check (-not $script:allOutput.ToString().Contains($sentinel)) 'o token (sentinela) nao aparece em nenhuma saida capturada'
    $hits = @(Get-ChildItem -LiteralPath $tmp -Recurse -File | Where-Object { [Text.Encoding]::UTF8.GetString([IO.File]::ReadAllBytes($_.FullName)).Contains($sentinel) })
    Check ($hits.Count -eq 0) "o token (sentinela) nao foi gravado em arquivo algum ($($hits.Count))"

    if ($script:fail -eq 0) { Write-Host "Test-PublishRelease: OK (PowerShell $($PSVersionTable.PSVersion))" -ForegroundColor Green; $exitCode = 0 }
    else { Write-Host "FALHAS Test-PublishRelease: $script:fail (PowerShell $($PSVersionTable.PSVersion))" -ForegroundColor Red }
} catch {
    Write-Host ("ERRO inesperado no Test-PublishRelease: " + ([string]$_.Exception.Message).Replace($sentinel, "***")) -ForegroundColor Red
    Write-Host $_.ScriptStackTrace
    $exitCode = 1
} finally {
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}
exit $exitCode
