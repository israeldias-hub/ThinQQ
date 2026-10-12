<#
.SYNOPSIS  Assina os projetos VBA dos .xlam por linha de comando (SignTool do Windows SDK + Office Subject Interface Packages) ou so diagnostica (-WhatIf).
.EXAMPLE   pwsh tools/Sign-ThinQQ.ps1 -Thumbprint <thumbprint do certificado em Cert:\CurrentUser\My> -TimestampUrl http://timestamp.digicert.com
.EXAMPLE   pwsh tools/Sign-ThinQQ.ps1 -PfxPath .\dist\cert\ThinQQ-signing.pfx -PfxPassword ... -SkipVerify
.NOTES     Pre-requisitos (uma vez, com administrador): Windows SDK (signtool.exe) e o pacote "Office Subject Interface Packages for Digitally
           Signing VBA Projects" (Microsoft), com as DLLs registradas por regsvr32 (leia o readme.txt do pacote). Sem eles este script sai com
           codigo 2 e nada e assinado. O VBE (Ferramentas > Assinatura Digital) continua sendo o caminho manual.
           ESTADO DE VERIFICACAO: so o diagnostico e o caminho de saida com codigo 2 foram exercitados (maquina sem o Windows SDK nem as SIPs).
           O ramo que chama signtool (sign e verify) NAO foi executado ponta a ponta numa maquina com SDK e SIPs: confira o resultado abrindo o
           .xlam no Excel (a assinatura de projeto VBA e a do Office) antes de distribuir.
           Prefira -Thumbprint (certificado em Cert:\CurrentUser\My, sem senha na linha de comando) a -PfxPath/-PfxPassword: uma senha em
           texto puro fica no historico do shell e e visivel a outros processos do usuario enquanto o signtool roda.
           VERIFICACAO (`signtool verify /pa`): ela monta a cadeia ate uma raiz confiavel. Um certificado autoassinado confiado so como Editor
           Confiavel (padrao de New-ThinQQCert) termina em CERT_E_UNTRUSTEDROOT: por padrao isso vira AVISO (assinou, codigo 0); qualquer outra
           falha da verificacao falha o script. Uma verificacao /pa completa exige New-ThinQQCert -TrustRoot (raiz confiavel do usuario) ou um
           certificado de uma CA corporativa. -SkipVerify pula a verificacao.
#>
param(
    [string[]]$Files = @((Join-Path $PSScriptRoot '..\dist\ThinQQ.xlam'), (Join-Path $PSScriptRoot '..\dist\ThinQQ.core.xlam')),
    [string]$PfxPath = '',
    [string]$PfxPassword = '',
    [string]$Thumbprint = '',
    [string]$TimestampUrl = '',
    [switch]$SkipVerify,
    [switch]$WhatIf
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\SignVerify.ps1')
function Find-SignTool {
    $c = Get-Command signtool.exe -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    foreach ($arch in 'x86', 'x64') {
        $cands = Get-ChildItem -Path (Join-Path ${env:ProgramFiles(x86)} 'Windows Kits\10\bin') -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue | Where-Object { $_.FullName -like "*\$arch\signtool.exe" } | Sort-Object FullName -Descending
        if ($cands) { return $cands[0].FullName }
    }
    return $null
}
# A SIP do Office fica na visao de 64 bits ou, registrada pelo regsvr32 de 32 bits (a do SignTool x86, que este script prefere), na
# WOW6432Node: procura nas duas. PSObject.Properties evita o erro de modo estrito numa subchave sem o valor Dll.
function Test-OfficeSip {
    foreach ($root in 'HKLM:\SOFTWARE\Microsoft\Cryptography\OID\EncodingType 0\CryptSIPDllGetSignedDataMsg', 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Cryptography\OID\EncodingType 0\CryptSIPDllGetSignedDataMsg') {
        if (-not (Test-Path $root)) { continue }
        foreach ($k in (Get-ChildItem $root -ErrorAction SilentlyContinue)) {
            $p = Get-ItemProperty $k.PSPath -ErrorAction SilentlyContinue
            if ($p -and $p.PSObject.Properties['Dll'] -and $p.Dll -match 'msosip') { return $true }
        }
    }
    return $false
}
$signtool = Find-SignTool
$sip = Test-OfficeSip
Write-Host ("signtool.exe: " + $(if ($signtool) { $signtool } else { 'NAO ENCONTRADO (instale o Windows SDK)' }))
Write-Host ("SIP do Office (msosip/msosipx registradas): " + $(if ($sip) { 'presente' } else { 'AUSENTE (instale o pacote Office SIPs e registre as DLLs com regsvr32, como administrador)' }))
if (-not $signtool -or -not $sip) {
    Write-Host 'Pre-requisitos ausentes: nada foi assinado. Caminho manual: VBE > Ferramentas > Assinatura Digital.' -ForegroundColor Yellow
    exit 2
}
if ($WhatIf) { Write-Host 'WhatIf: pre-requisitos presentes; a assinatura rodaria agora.'; exit 0 }
if (-not $PfxPath -and -not $Thumbprint) { throw 'Informe -PfxPath (e -PfxPassword) ou -Thumbprint.' }
foreach ($f in $Files) {
    if (-not (Test-Path -LiteralPath $f)) { throw "Arquivo nao encontrado: $f" }
    $stArgs = @('sign', '/fd', 'SHA256')                 # nunca $args: e a variavel automatica dos argumentos do script
    if ($PfxPath) { $stArgs += @('/f', $PfxPath); if ($PfxPassword) { $stArgs += @('/p', $PfxPassword) } } else { $stArgs += @('/sha1', $Thumbprint) }
    if ($TimestampUrl) { $stArgs += @('/tr', $TimestampUrl, '/td', 'SHA256') }
    $stArgs += $f
    $st = Invoke-SignTool $signtool $stArgs
    if ($st.Code -ne 0) { throw "signtool sign falhou ($($st.Code)) em $f" }
    if ($SkipVerify) { Write-Host "Assinado (verificacao /pa pulada por -SkipVerify): $f"; continue }
    $vr = Invoke-SignTool $signtool @('verify', '/pa', $f)
    switch (Get-SignVerifyOutcome $vr.Code $vr.Output) {
        'ok' { Write-Host "Assinado e verificado: $f" }
        'untrusted-root' {
            Write-Host "Assinado: $f" -ForegroundColor Green
            Write-Host 'AVISO: a verificacao /pa terminou numa raiz nao confiavel (CERT_E_UNTRUSTEDROOT). E esperado com certificado autoassinado confiado so como Editor Confiavel; a assinatura foi gravada. Uma verificacao /pa completa exige New-ThinQQCert -TrustRoot (raiz confiavel do usuario) ou um certificado de CA corporativa.' -ForegroundColor Yellow
        }
        default { throw "signtool verify falhou ($($vr.Code)) em $f" }
    }
}
exit 0
