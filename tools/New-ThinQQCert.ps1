<#
.SYNOPSIS  Cria um certificado autoassinado de assinatura de codigo para o ThinQQ (por usuario, sem administrador), exporta .cer e .pfx
           e instala o .cer como Editor Confiavel do usuario. Com -TrustRoot tambem o instala em Autoridades Raiz Confiaveis do usuario
           (o Windows pede confirmacao numa janela). Sem a raiz, o Office mostra o editor como "nao verificado".
.EXAMPLE   pwsh tools/New-ThinQQCert.ps1 -Subject "CN=Minha Empresa - ThinQQ" -TrustRoot
.PARAMETER Force  Permite sobrescrever ThinQQ-publisher.cer / ThinQQ-signing.pfx que ja existam em -OutDir. Sem ele, arquivos existentes
           fazem o script falhar ANTES de criar qualquer coisa. Cada execucao cria um certificado NOVO (mesmo assunto, thumbprint diferente):
           rodar de novo acumula certificados em Cert:\CurrentUser\My e em Editores Confiaveis; remova os antigos pelo thumbprint (certmgr.msc).
.NOTES     Falha no meio: tudo o que esta execucao gravou (certificado em My com a chave, copia em CA, entrada em Editores Confiaveis, e .cer/.pfx
           e pasta de saida que ela criou) e desfeito e o erro e repassado. A importacao em Raiz Confiavel (-TrustRoot) e o ULTIMO passo e NAO e
           desfeita por esse rollback; se voce recusar a janela do Windows, ela nao acontece e a execucao inteira e desfeita.
           A senha do PFX gerada (se -PfxPassword vazio) e impressa uma vez; uma senha passada em -PfxPassword fica no historico do shell.
#>
param(
    [string]$Subject = 'CN=ThinQQ (autoassinado)',
    [string]$OutDir = (Join-Path $PSScriptRoot '..\dist\cert'),
    [int]$Years = 3,
    [string]$PfxPassword = '',
    [switch]$TrustRoot,
    [switch]$NoTrustPublisher,
    [switch]$Force
)
$ErrorActionPreference = 'Stop'
$cer = Join-Path $OutDir 'ThinQQ-publisher.cer'
$pfx = Join-Path $OutDir 'ThinQQ-signing.pfx'
if (-not $Force) {
    $existing = @($cer, $pfx) | Where-Object { Test-Path -LiteralPath $_ }
    if ($existing) { throw "Arquivo(s) ja existe(m) em ${OutDir}: $((@($existing) | ForEach-Object { Split-Path $_ -Leaf }) -join ', '). Nada foi criado. Use -Force para sobrescrever (cada execucao cria um certificado novo; os antigos continuam nas lojas) ou escolha outro -OutDir." }
}
$outDirCreated = -not (Test-Path -LiteralPath $OutDir)
$newFiles = @($cer, $pfx) | Where-Object { -not (Test-Path -LiteralPath $_) }   # so desfaz arquivos que esta execucao cria
New-Item -ItemType Directory -Force $OutDir | Out-Null
$cert = $null
try {
    $cert = New-SelfSignedCertificate -Type CodeSigningCert -Subject $Subject -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 `
        -KeyExportPolicy Exportable -CertStoreLocation Cert:\CurrentUser\My -NotAfter (Get-Date).AddYears($Years) -FriendlyName 'ThinQQ assinatura de codigo'
    # o cmdlet tambem grava uma copia em "Autoridades de Certificacao Intermediarias" (Cert:\CurrentUser\CA); ela nao serve para nada aqui
    Get-ChildItem Cert:\CurrentUser\CA | Where-Object Thumbprint -eq $cert.Thumbprint | Remove-Item
    Export-Certificate -Cert $cert -FilePath $cer -Type CERT | Out-Null
    if (-not $PfxPassword) {
        $PfxPassword = -join ((48..57 + 65..90 + 97..122) | Get-Random -Count 16 | ForEach-Object { [char]$_ })
        Write-Host "Senha do PFX gerada (guarde-a; nao e gravada em lugar nenhum): $PfxPassword"
    }
    Export-PfxCertificate -Cert $cert -FilePath $pfx -Password (ConvertTo-SecureString $PfxPassword -AsPlainText -Force) | Out-Null
    if (-not $NoTrustPublisher) { Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null }
    if ($TrustRoot) {
        # ultimo passo (nao ha nada depois dele que possa falhar e deixar a raiz para tras)
        Write-Host 'Instalando o certificado em Autoridades Raiz Confiaveis do usuario: o Windows vai pedir confirmacao.' -ForegroundColor Yellow
        Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
    }
} catch {
    # falha: desfaz o que esta execucao fez (nao deixa certificado orfao nas lojas do usuario) e repassa o erro
    if ($cert) {
        foreach ($st in 'My', 'CA', 'TrustedPublisher') {
            Get-ChildItem "Cert:\CurrentUser\$st" -ErrorAction SilentlyContinue | Where-Object Thumbprint -eq $cert.Thumbprint | ForEach-Object {
                if ($st -eq 'My') { Remove-Item $_.PSPath -DeleteKey -ErrorAction SilentlyContinue } else { Remove-Item $_.PSPath -ErrorAction SilentlyContinue } }
        }
    }
    foreach ($f in $newFiles) { Remove-Item -LiteralPath $f -Force -ErrorAction SilentlyContinue }
    if ($outDirCreated -and -not (Get-ChildItem -LiteralPath $OutDir -Force -ErrorAction SilentlyContinue)) { Remove-Item -LiteralPath $OutDir -Force -ErrorAction SilentlyContinue }
    Write-Host 'Falha ao criar o certificado: o que esta execucao gravou (certificado, chave, .cer/.pfx novos) foi desfeito. Obs.: a importacao em Raiz Confiavel (-TrustRoot) e o ultimo passo e nao e desfeita por este rollback; se voce recusar a janela do Windows, a execucao inteira e desfeita.' -ForegroundColor Yellow
    throw
}
Write-Host "Certificado $($cert.Thumbprint) ($Subject) criado; .cer e .pfx em $OutDir"
[pscustomobject]@{ Thumbprint = $cert.Thumbprint; Subject = $cert.Subject; Cer = $cer; Pfx = $pfx; NotAfter = $cert.NotAfter }
