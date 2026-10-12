# Funcoes de apoio de tools/Sign-ThinQQ.ps1, testadas por tests/ps/Test-Signing.ps1 sem precisar do signtool (Get-SignVerifyOutcome e pura;
# Invoke-SignTool e exercitada com cmd.exe no lugar do signtool).
# Get-SignVerifyOutcome classifica o resultado de `signtool verify /pa`:
#   'ok'             codigo 0
#   'untrusted-root' a cadeia termina numa raiz nao confiavel (CERT_E_UNTRUSTEDROOT, 0x800B0109): caso normal de um certificado autoassinado
#                    confiavel so como Editor Confiavel; a assinatura foi feita, so a verificacao /pa nao fecha ate a raiz ser confiada
#                    (New-ThinQQCert -TrustRoot) ou o certificado vir de uma CA corporativa
#   'failure'        qualquer outro codigo diferente de 0 (assinatura invalida, arquivo ilegivel, saida desconhecida...)
# O reconhecimento usa o HRESULT (estavel em qualquer idioma) e o texto em ingles. Se o signtool localizado nao imprimir nenhum dos dois, o
# resultado e 'failure' (falha segura). Se houver outro HRESULT de erro junto do 0x800B0109, tambem e 'failure'.
function Get-SignVerifyOutcome([int]$ExitCode, [string]$Output) {
    if ($ExitCode -eq 0) { return 'ok' }
    if (-not $Output) { return 'failure' }
    $untrusted = ($Output -match '0x800B0109') -or ($Output -match 'CERT_E_UNTRUSTEDROOT') -or ($Output -match 'root certificate which is not trusted')
    if (-not $untrusted) { return 'failure' }
    $others = @([regex]::Matches($Output, '0x8[0-9A-Fa-f]{7}') | ForEach-Object { $_.Value.ToUpper() } | Where-Object { $_ -ne '0X800B0109' })
    if ($others.Count -gt 0) { return 'failure' }
    return 'untrusted-root'
}

# Roda o signtool: a saida vai para o host (Out-Host), nao para o fluxo de sucesso do script (quem captura a saida do script nao recebe o ruido
# do signtool), e tambem e devolvida como texto para a classificacao. Com $ErrorActionPreference='Stop', o 2>&1 de um programa nativo
# transformaria a primeira linha de stderr em erro terminante no Windows PowerShell 5.1: dai o 'Continue' local.
function Invoke-SignTool([string]$Exe, [string[]]$ToolArgs) {
    $global:LASTEXITCODE = 0
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { $text = (& $Exe @ToolArgs 2>&1 | ForEach-Object { if ($_ -is [System.Management.Automation.ErrorRecord]) { $_.Exception.Message } else { $_ } } | Out-String); $code = $global:LASTEXITCODE } finally { $ErrorActionPreference = $prev }
    $text | Out-Host
    [pscustomobject]@{ Code = $code; Output = $text }
}
