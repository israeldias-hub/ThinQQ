# Executa install.ps1 / uninstall.ps1 contra pastas e chave de registro temporarias.
# Roda em pwsh 7 e em Windows PowerShell 5.1.
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$tmp  = Join-Path $env:TEMP ("tqq_inst_" + [guid]::NewGuid().ToString('N'))
$pkg  = Join-Path $tmp 'pkg'; $root = Join-Path $tmp 'root'
$reg  = 'HKCU:\Software\ThinQQ_Test_' + [guid]::NewGuid().ToString('N').Substring(0, 8)
$install   = Join-Path $repo 'installer\install.ps1'
$uninstall = Join-Path $repo 'installer\uninstall.ps1'
New-Item -ItemType Directory -Force $pkg | Out-Null
Set-Content (Join-Path $pkg 'ThinQQ.xlam') 'loader'
Set-Content (Join-Path $pkg 'ThinQQ.core.xlam') 'core'
Set-Content (Join-Path $pkg 'tutorial.html') '<html></html>'
function Write-Manifest($coreHash, $loaderHash) {
    @{ coreVersion = '1.2.3'; loaderVersion = '1.0.0'; manifestUrl = 'https://exemplo/manifest.json'
       coreSha256 = $coreHash; loaderSha256 = $loaderHash } | ConvertTo-Json | Set-Content (Join-Path $pkg 'manifest.json')
}
function Sha($name) { (Get-FileHash (Join-Path $pkg $name) -Algorithm SHA256).Hash.ToLower() }
Write-Manifest (Sha 'ThinQQ.core.xlam') (Sha 'ThinQQ.xlam')
function Assert($cond, $msg) { if (-not $cond) { throw "FALHA: $msg" } }
function Expect-Failure($label, [scriptblock]$block) {
    $failed = $false
    try { & $block } catch { $failed = $true }
    Assert $failed $label
}
$optKey = Join-Path $reg 'Microsoft\Office\16.0\Excel\Options'
$upKey  = Join-Path $reg 'VB and VBA Program Settings\ThinQQ\Updater'
try {
    # Instalacao recusada: hash divergente, arquivo ausente (nada instalado)
    Write-Manifest 'deadbeef' (Sha 'ThinQQ.xlam')
    Expect-Failure 'hash do core divergente deve recusar' { & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck }
    Write-Manifest (Sha 'ThinQQ.core.xlam') 'deadbeef'
    Expect-Failure 'hash do loader divergente deve recusar' { & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck }
    Assert (-not (Test-Path $root)) 'nada copiado apos hash divergente'
    Assert (-not (Test-Path $reg)) 'registro intacto apos hash divergente'
    Write-Manifest (Sha 'ThinQQ.core.xlam') (Sha 'ThinQQ.xlam')
    Move-Item (Join-Path $pkg 'ThinQQ.core.xlam') (Join-Path $tmp 'core.bak')
    Expect-Failure 'arquivo ausente deve recusar' { & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck }
    Assert (-not (Test-Path $root)) 'nada copiado com arquivo ausente'
    Move-Item (Join-Path $tmp 'core.bak') (Join-Path $pkg 'ThinQQ.core.xlam')

    # Instalacao normal
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck
    Assert (Test-Path (Join-Path $root 'ThinQQ.xlam')) 'loader copiado'
    Assert (Test-Path (Join-Path $root 'versions\1.2.3\ThinQQ.core.xlam')) 'core copiado na pasta da versao'
    Assert (Test-Path (Join-Path $root 'versions\1.2.3\tutorial.html')) 'tutorial copiado'
    $opt = Get-ItemProperty $optKey
    Assert ($opt.OPEN -eq ('/R "' + (Join-Path $root 'ThinQQ.xlam') + '"')) "valor OPEN: $($opt.OPEN)"
    $tl = Get-ItemProperty (Join-Path $reg 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ')
    Assert ($tl.AllowSubfolders -eq 1) 'local confiavel com subpastas'
    Assert ((Get-ItemProperty $upKey).UpdateSource -eq 'https://exemplo/manifest.json') 'UpdateSource do manifesto'
    # Segunda instalacao nao duplica OPEN
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck
    Assert (-not ((Get-ItemProperty $optKey).PSObject.Properties.Name -contains 'OPEN1')) 'OPEN nao duplicado'
    # UpdateSource existente e preservado
    Set-ItemProperty $upKey -Name UpdateSource -Value 'D:\minha\fonte\manifest.json'
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck
    Assert ((Get-ItemProperty $upKey).UpdateSource -eq 'D:\minha\fonte\manifest.json') 'UpdateSource existente preservado'
    # UpdateSource explicito prevalece
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck -UpdateSource '\srv\tqq\manifest.json'
    Assert ((Get-ItemProperty $upKey).UpdateSource -eq '\srv\tqq\manifest.json') 'UpdateSource explicito'

    # Desinstalar preserva outros suplementos e renumera
    Set-ItemProperty $optKey -Name OPEN1 -Value '/R "C:\outro\A.xlam"'
    Set-ItemProperty $optKey -Name OPEN2 -Value '/R "C:\outro\B.xlam"'
    & $uninstall -Root $root -RegistryRoot $reg -SkipExcelCheck
    $names = (Get-ItemProperty $optKey).PSObject.Properties | Where-Object { $_.Name -match '^OPEN\d*$' }
    $o = Get-ItemProperty $optKey
    Assert ($o.OPEN -eq '/R "C:\outro\A.xlam"') "OPEN renumerado: $($o.OPEN)"
    Assert ($o.OPEN1 -eq '/R "C:\outro\B.xlam"') "OPEN1 renumerado: $($o.OPEN1)"
    Assert (-not ($o.PSObject.Properties.Name -contains 'OPEN2')) 'OPEN2 removido'
    Assert (-not (Test-Path (Join-Path $reg 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ'))) 'local confiavel removido'
    Assert (-not (Test-Path (Join-Path $reg 'VB and VBA Program Settings\ThinQQ'))) 'configuracoes removidas'
    Assert (-not (Test-Path $root)) 'arquivos removidos'
    Write-Host "Test-Installer: OK (PowerShell $($PSVersionTable.PSVersion))" -ForegroundColor Green
} finally {
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $reg -Recurse -Force -ErrorAction SilentlyContinue
}
