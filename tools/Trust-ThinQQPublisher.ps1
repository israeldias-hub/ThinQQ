<#
.SYNOPSIS  Instala um certificado de editor (.cer) como Editor Confiavel (e, com -TrustRoot, como raiz confiavel) do usuario atual ou da maquina.
.EXAMPLE   pwsh tools/Trust-ThinQQPublisher.ps1 -CerPath .\ThinQQ-publisher.cer -TrustRoot
.NOTES     -Scope LocalMachine exige administrador: sem ele o script sai com codigo 2 e uma mensagem clara, sem tentar gravar. Em CurrentUser, a
           importacao em Raiz (-TrustRoot) faz o Windows pedir confirmacao numa janela. Em empresas, prefira distribuir o .cer por GPO
           (Computer Configuration > Policies > Windows Settings > Security Settings > Public Key Policies > Trusted Publishers).
#>
param(
    [Parameter(Mandatory)][string]$CerPath,
    [switch]$TrustRoot,
    [ValidateSet('CurrentUser', 'LocalMachine')][string]$Scope = 'CurrentUser'
)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $CerPath)) { throw "Certificado nao encontrado: $CerPath" }
if ($Scope -eq 'LocalMachine') {
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host '-Scope LocalMachine exige administrador: abra o PowerShell como administrador (ou use -Scope CurrentUser, ou distribua o .cer por GPO). Nada foi instalado.' -ForegroundColor Red
        exit 2
    }
}
$c = Import-Certificate -FilePath $CerPath -CertStoreLocation "Cert:\$Scope\TrustedPublisher"
if ($TrustRoot) {
    if ($Scope -eq 'CurrentUser') { Write-Host 'Instalando em Autoridades Raiz Confiaveis do usuario: o Windows vai pedir confirmacao numa janela.' -ForegroundColor Yellow }
    else { Write-Host 'Instalando em Autoridades Raiz Confiaveis da maquina.' -ForegroundColor Yellow }
    Import-Certificate -FilePath $CerPath -CertStoreLocation "Cert:\$Scope\Root" | Out-Null
}
Write-Host "Editor confiavel ($Scope): $($c.Subject) [$($c.Thumbprint)]"
$c.Thumbprint
