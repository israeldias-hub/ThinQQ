<#
.SYNOPSIS  Assina os projetos VBA dos .xlam por linha de comando (SignTool do Windows SDK + Office Subject Interface Packages) ou so diagnostica (-WhatIf).
.EXAMPLE   pwsh tools/Sign-ThinQQ.ps1 -PfxPath .\dist\cert\ThinQQ-signing.pfx -PfxPassword ... -TimestampUrl http://timestamp.digicert.com
.NOTES     Pre-requisitos (uma vez, com administrador): Windows SDK (signtool.exe) e o pacote "Office Subject Interface Packages for Digitally
           Signing VBA Projects" (Microsoft), com as DLLs registradas por regsvr32 (leia o readme.txt do pacote). Sem eles este script sai com
           codigo 2 e nada e assinado. O VBE (Ferramentas > Assinatura Digital) continua sendo o caminho manual.
#>
param(
    [string[]]$Files = @((Join-Path $PSScriptRoot '..\dist\ThinQQ.xlam'), (Join-Path $PSScriptRoot '..\dist\ThinQQ.core.xlam')),
    [string]$PfxPath = '',
    [string]$PfxPassword = '',
    [string]$Thumbprint = '',
    [string]$TimestampUrl = '',
    [switch]$WhatIf
)
$ErrorActionPreference = 'Stop'
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
    Write-Host 'Pre-requisitos ausentes: nada foi assinado. Caminho manual: VBE > Ferramentas > Assinatura Digital (gera assinatura V3).' -ForegroundColor Yellow
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
    & $signtool @stArgs
    if ($LASTEXITCODE -ne 0) { throw "signtool sign falhou ($LASTEXITCODE) em $f" }
    & $signtool verify /pa $f
    if ($LASTEXITCODE -ne 0) { throw "signtool verify falhou ($LASTEXITCODE) em $f" }
    Write-Host "Assinado e verificado: $f"
}
exit 0
