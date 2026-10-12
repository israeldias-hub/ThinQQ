<#
.SYNOPSIS  Cria um certificado autoassinado de assinatura de codigo para o ThinQQ (por usuario, sem administrador), exporta .cer e .pfx
           e instala o .cer como Editor Confiavel do usuario. Com -TrustRoot tambem o instala em Autoridades Raiz Confiaveis do usuario
           (o Windows pede confirmacao numa janela). Sem a raiz, o Office mostra o editor como "nao verificado".
.EXAMPLE   pwsh tools/New-ThinQQCert.ps1 -Subject "CN=Minha Empresa - ThinQQ" -TrustRoot
#>
param(
    [string]$Subject = 'CN=ThinQQ (autoassinado)',
    [string]$OutDir = (Join-Path $PSScriptRoot '..\dist\cert'),
    [int]$Years = 3,
    [string]$PfxPassword = '',
    [switch]$TrustRoot,
    [switch]$NoTrustPublisher
)
$ErrorActionPreference = 'Stop'
New-Item -ItemType Directory -Force $OutDir | Out-Null
$cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $Subject -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
    -KeyExportPolicy Exportable -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears($Years) -FriendlyName 'ThinQQ assinatura de codigo'
# o cmdlet tambem grava uma copia em "Autoridades de Certificacao Intermediarias" (Cert:\CurrentUser\CA); ela nao serve para nada aqui
Get-ChildItem Cert:\CurrentUser\CA | Where-Object Thumbprint -eq $cert.Thumbprint | Remove-Item
$cer = Join-Path $OutDir 'ThinQQ-publisher.cer'
$pfx = Join-Path $OutDir 'ThinQQ-signing.pfx'
$newFiles = @($cer, $pfx) | Where-Object { -not (Test-Path -LiteralPath $_) }   # so desfaz arquivos que esta execucao cria
try {
    Export-Certificate -Cert $cert -FilePath $cer -Type CERT | Out-Null
    if (-not $PfxPassword) {
        $PfxPassword = -join ((48..57 + 65..90 + 97..122) | Get-Random -Count 16 | ForEach-Object { [char]$_ })
        Write-Host "Senha do PFX gerada (guarde-a; nao e gravada em lugar nenhum): $PfxPassword"
    }
    Export-PfxCertificate -Cert $cert -FilePath $pfx -Password (ConvertTo-SecureString $PfxPassword -AsPlainText -Force) | Out-Null
    if (-not $NoTrustPublisher) { Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null }
    if ($TrustRoot) {
        Write-Host 'Instalando o certificado em Autoridades Raiz Confiaveis do usuario: o Windows vai pedir confirmacao.' -ForegroundColor Yellow
        Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
    }
} catch {
    # falha depois de criar o certificado: desfaz o que esta execucao fez (nao deixa certificado orfao nas lojas do usuario) e repassa o erro
    foreach ($st in 'My', 'CA', 'TrustedPublisher') {
        Get-ChildItem "Cert:\CurrentUser\$st" -ErrorAction SilentlyContinue | Where-Object Thumbprint -eq $cert.Thumbprint | ForEach-Object {
            if ($st -eq 'My') { Remove-Item $_.PSPath -DeleteKey -ErrorAction SilentlyContinue } else { Remove-Item $_.PSPath -ErrorAction SilentlyContinue } }
    }
    foreach ($f in $newFiles) { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }
    throw
}
Write-Host "Certificado $($cert.Thumbprint) ($Subject) criado; .cer e .pfx em $OutDir"
[pscustomobject]@{ Thumbprint = $cert.Thumbprint; Subject = $cert.Subject; Cer = $cer; Pfx = $pfx; NotAfter = $cert.NotAfter }
