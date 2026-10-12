<#
.SYNOPSIS  Instala um certificado de editor (.cer) como Editor Confiavel (e, com -TrustRoot, como raiz confiavel) do usuario atual ou da maquina.
.EXAMPLE   pwsh tools/Trust-ThinQQPublisher.ps1 -CerPath .\ThinQQ-publisher.cer -TrustRoot
.NOTES     -Scope LocalMachine exige administrador. Em empresas, prefira distribuir o .cer por GPO (Computer Configuration > Policies >
           Windows Settings > Security Settings > Public Key Policies > Trusted Publishers).
#>
param(
    [Parameter(Mandatory)][string]$CerPath,
    [switch]$TrustRoot,
    [ValidateSet('CurrentUser', 'LocalMachine')][string]$Scope = 'CurrentUser'
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $CerPath)) { throw "Certificado nao encontrado: $CerPath" }
$c = Import-Certificate -FilePath $CerPath -CertStoreLocation "Cert:\$Scope\TrustedPublisher"
if ($TrustRoot) {
    Write-Host 'Instalando em Autoridades Raiz Confiaveis: o Windows pode pedir confirmacao.' -ForegroundColor Yellow
    Import-Certificate -FilePath $CerPath -CertStoreLocation "Cert:\$Scope\Root" | Out-Null
}
Write-Host "Editor confiavel ($Scope): $($c.Subject) [$($c.Thumbprint)]"
$c.Thumbprint
