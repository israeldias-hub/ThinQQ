# Certificado autoassinado por usuario (sem administrador): criar, exportar, confiar como editor, diagnosticar/assinar e REMOVER. Nunca toca
# Cert:\...\Root nem grava em HKLM. Assunto unico por execucao; a limpeza do finally remove pelo ASSUNTO, nas tres lojas em que o certificado
# pode estar (My; CA, onde New-SelfSignedCertificate grava uma copia em "Autoridades de Certificacao Intermediarias"; TrustedPublisher),
# inclusive quando a criacao falha no meio, e confere que nada sobrou. O run-tests o chama num processo proprio: termina sempre com exit 0 ou 1.
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$tmp = Join-Path $env:TEMP ("tqq_sign_" + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $tmp | Out-Null
function Assert($cond, $msg) { if (-not $cond) { throw "FALHA: $msg" } }
function Get-DistHash([string]$p) { if (Test-Path -LiteralPath $p) { return (Get-FileHash -LiteralPath $p).Hash } else { return '' } }
$subject = 'CN=ThinQQ Teste ' + [guid]::NewGuid().ToString('N').Substring(0, 8)
$subjectFail = 'CN=ThinQQ Teste ' + [guid]::NewGuid().ToString('N').Substring(0, 8)   # segundo certificado: caminho de falha do New-ThinQQCert (desfeito por ele)
$subjects = @($subject, $subjectFail)
$exitCode = 1
$dist = Join-Path $repo 'dist\ThinQQ.xlam'
$distHash = Get-DistHash $dist
try {
    # classificacao do resultado do `signtool verify /pa` (funcao pura, sem signtool): uma raiz nao confiavel (certificado autoassinado confiavel so
    # como editor) vira aviso; qualquer outra falha continua falhando. Saidas de exemplo no formato do signtool (hex estavel, texto em ingles)
    . (Join-Path $repo 'tools\lib\SignVerify.ps1')
    Assert ((Get-SignVerifyOutcome 0 'Successfully verified: x.xlam') -eq 'ok') 'verify: codigo 0 => ok'
    Assert ((Get-SignVerifyOutcome 0 'A certificate chain processed, but terminated in a root certificate which is not trusted by the trust provider.') -eq 'ok') 'verify: codigo 0 prevalece sobre o texto'
    Assert ((Get-SignVerifyOutcome 1 "SignTool Error: WinVerifyTrust returned error: 0x800B0109`n`tA certificate chain processed, but terminated in a root certificate which is not trusted by the trust provider.") -eq 'untrusted-root') 'verify: 0x800B0109 => untrusted-root'
    Assert ((Get-SignVerifyOutcome 1 'SignTool Error: A certificate chain processed, but terminated in a root certificate which is not trusted by the trust provider.') -eq 'untrusted-root') 'verify: texto em ingles => untrusted-root'
    Assert ((Get-SignVerifyOutcome 1 'SignTool Error: WinVerifyTrust returned error: 0x80096010 The digital signature of the object did not verify.') -eq 'failure') 'verify: assinatura invalida => failure'
    Assert ((Get-SignVerifyOutcome 1 '') -eq 'failure') 'verify: falha sem texto => failure'
    Assert ((Get-SignVerifyOutcome 2 $null) -eq 'failure') 'verify: saida nula => failure'
    Assert ((Get-SignVerifyOutcome 1 '0x80096010 e tambem 0x800B0109') -eq 'failure') 'verify: raiz nao confiavel misturada com outra falha => failure'
    # Invoke-SignTool com cmd.exe no lugar do signtool: codigo de saida e texto (stdout+stderr) chegam intactos, mesmo com $ErrorActionPreference='Stop'
    $vr = Invoke-SignTool $env:ComSpec @('/c', 'echo SignTool Error: 0x800B0109 & echo erro-no-stderr 1>&2 & exit 1')
    Assert ($vr.Code -eq 1 -and $vr.Output -match '0x800B0109' -and $vr.Output -match 'erro-no-stderr') "Invoke-SignTool: codigo/saida ($($vr.Code))"
    Assert ((Get-SignVerifyOutcome $vr.Code $vr.Output) -eq 'untrusted-root') 'Invoke-SignTool + classificacao de ponta a ponta'
    Assert ((Invoke-SignTool $env:ComSpec @('/c', 'exit 0')).Code -eq 0) 'Invoke-SignTool: codigo 0'
    $r = & (Join-Path $repo 'tools\New-ThinQQCert.ps1') -Subject $subject -OutDir $tmp -Years 1 -PfxPassword 'teste'
    $thumb = $r.Thumbprint
    Assert ($thumb -match '^[0-9A-F]{40}$') "thumbprint: $thumb"
    Assert ((Get-Item "Cert:\CurrentUser\My\$thumb").EnhancedKeyUsageList.ObjectId -contains '1.3.6.1.5.5.7.3.3') 'EKU de assinatura de codigo'
    Assert (-not (Test-Path "Cert:\CurrentUser\CA\$thumb")) 'nenhuma copia em Autoridades Intermediarias (New-ThinQQCert a remove)'
    Assert (Test-Path $r.Cer) '.cer exportado'; Assert (Test-Path $r.Pfx) '.pfx exportado'
    # EphemeralKeySet: abrir o PFX nao grava a chave privada num conteiner do perfil; Dispose ao final
    $pfx = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($r.Pfx, 'teste', [System.Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet)
    try { Assert ($pfx.HasPrivateKey -and $pfx.Thumbprint -eq $thumb) 'pfx com a chave privada' } finally { $pfx.Dispose() }
    Assert (Test-Path "Cert:\CurrentUser\TrustedPublisher\$thumb") 'editor confiavel do usuario'
    $t2 = & (Join-Path $repo 'tools\Trust-ThinQQPublisher.ps1') -CerPath $r.Cer
    Assert ($t2 -eq $thumb) 'Trust-ThinQQPublisher idempotente'
    Assert (-not (Test-Path "Cert:\CurrentUser\Root\$thumb")) 'Root nao foi tocada'
    # diagnostico / assinatura: Sign-ThinQQ so escreve com Write-Host (fluxo de informacao): *>&1 o captura; 2>&1 nao captura nada
    $copy = Join-Path $tmp 'ThinQQ.xlam'
    if (Test-Path $dist) { Copy-Item $dist $copy } else { Set-Content $copy 'x' }
    $global:LASTEXITCODE = 0
    $out = & (Join-Path $repo 'tools\Sign-ThinQQ.ps1') -Files $copy -PfxPath $r.Pfx -PfxPassword 'teste' -WhatIf *>&1 | Out-String
    $code = $global:LASTEXITCODE
    Assert ($code -eq 0 -or $code -eq 2) "Sign-ThinQQ -WhatIf: codigo $code"
    Assert ($out -match 'signtool' -or $out -match 'SIP') "diagnostico cita signtool/SIP: $out"
    if ($code -eq 0 -and (Test-Path $dist)) {
        $global:LASTEXITCODE = 0
        & (Join-Path $repo 'tools\Sign-ThinQQ.ps1') -Files $copy -PfxPath $r.Pfx -PfxPassword 'teste' -SkipVerify *>&1 | Out-Null
        Assert ($global:LASTEXITCODE -eq 0) 'assinatura real da copia (ferramentas presentes)'
        Write-Host 'Test-Signing: assinatura real executada numa copia'
    } else { Write-Host "Test-Signing: ferramentas de assinatura ausentes (codigo $code): so o diagnostico foi verificado" }
    Assert ((Get-DistHash $dist) -eq $distHash) 'dist\ThinQQ.xlam nao mudou (a assinatura real e feita numa copia)'
    # falha DEPOIS de criar o certificado (uma pasta ocupa o nome do .cer, o que faz o Export-Certificate falhar): o New-ThinQQCert deve
    # desfazer sozinho o que gravou nas lojas e nao deixar .pfx/.cer novos para tras
    $failDir = Join-Path $tmp 'falha'
    New-Item -ItemType Directory -Force (Join-Path $failDir 'ThinQQ-publisher.cer') | Out-Null
    $threw = $false
    try { & (Join-Path $repo 'tools\New-ThinQQCert.ps1') -Subject $subjectFail -OutDir $failDir -Years 1 -PfxPassword 'teste' | Out-Null } catch { $threw = $true }
    Assert $threw 'New-ThinQQCert propaga a falha de exportacao'
    foreach ($store in 'My', 'CA', 'TrustedPublisher') {
        Assert (-not @(Get-ChildItem "Cert:\CurrentUser\$store" | Where-Object Subject -eq $subjectFail)) "New-ThinQQCert desfez o certificado da falha em $store"
    }
    Assert (-not (Test-Path (Join-Path $failDir 'ThinQQ-signing.pfx'))) 'nenhum .pfx orfao apos a falha'
    Write-Host "Test-Signing: OK (PowerShell $($PSVersionTable.PSVersion))" -ForegroundColor Green
    $exitCode = 0
} catch {
    Write-Host "not ok - $($_.Exception.Message)" -ForegroundColor Red
} finally {
    foreach ($store in 'My', 'CA', 'TrustedPublisher') {
        Get-ChildItem "Cert:\CurrentUser\$store" -ErrorAction SilentlyContinue | Where-Object { $subjects -contains $_.Subject } | ForEach-Object {
            if ($store -eq 'My') { Remove-Item $_.PSPath -DeleteKey -ErrorAction SilentlyContinue } else { Remove-Item $_.PSPath -ErrorAction SilentlyContinue } }
    }
    # Confere TODAS as lojas do usuario, uma a uma: "Get-ChildItem Cert:\CurrentUser -Recurse" levanta erro terminante em "UserDS" (loja do
    # Active Directory, inacessivel fora do dominio) mesmo com -ErrorAction SilentlyContinue, o que abortaria o finally antes de apagar a pasta.
    $left = @()
    foreach ($st in @(Get-ChildItem Cert:\CurrentUser -ErrorAction SilentlyContinue)) {
        try { $left += @(Get-ChildItem ('Cert:\CurrentUser\' + $st.Name) -ErrorAction Stop | Where-Object { $_.PSObject.Properties['Subject'] -and $subjects -contains $_.Subject } | ForEach-Object { $st.Name }) } catch { }
    }
    if ($left.Count -gt 0) { Write-Host "not ok - certificado de teste sobrou em: $($left -join ', ')" -ForegroundColor Red; $exitCode = 1 }
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
exit $exitCode
