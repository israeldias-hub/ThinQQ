# Executa install.ps1 / uninstall.ps1 contra pastas e chaves de registro temporarias (HKCU:\Software\ThinQQ_Test_*).
# Roda em pwsh 7 e em Windows PowerShell 5.1. Nunca toca %LOCALAPPDATA%\ThinQQ nem as chaves reais do Excel.
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
$install   = Join-Path $repo 'installer\install.ps1'
$uninstall = Join-Path $repo 'installer\uninstall.ps1'
$tmp  = Join-Path $env:TEMP ("tqq_inst_" + [guid]::NewGuid().ToString('N'))
$regs = New-Object System.Collections.ArrayList
New-Item -ItemType Directory -Force $tmp | Out-Null
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Assert($cond, $msg) { if (-not $cond) { throw "FALHA: $msg" } }
function New-Reg { $r = 'HKCU:\Software\ThinQQ_Test_' + [guid]::NewGuid().ToString('N').Substring(0, 8); [void]$regs.Add($r); return $r }
function Expect-Failure($label, [scriptblock]$block, $pattern = '') {
    $msg = $null
    try { & $block } catch { $msg = $_.Exception.Message }
    Assert ($null -ne $msg) $label
    if ($pattern) { Assert ($msg -match $pattern) "$label (mensagem: $msg)" }
}
function Sha($dir, $name) { (Get-FileHash (Join-Path $dir $name) -Algorithm SHA256).Hash.ToLower() }
# Cria um pacote; $loaderBytesText vira o conteudo do "loader"; $loaderZipVersion != '' cria um zip com a versao no docProps
function New-Pkg([string]$name, [string]$loaderText = 'loader', [string]$loaderVer = '1.0.0', [string]$loaderZipVersion = '') {
    $p = Join-Path $tmp $name
    New-Item -ItemType Directory -Force $p | Out-Null
    if ($loaderZipVersion) {
        $zp = Join-Path $p 'ThinQQ.xlam'
        if (Test-Path $zp) { Remove-Item $zp -Force }
        $z = [IO.Compression.ZipFile]::Open($zp, 'Create')
        $e = $z.CreateEntry('docProps/core.xml')
        $w = New-Object IO.StreamWriter($e.Open())
        $w.Write("<cp:coreProperties><dc:description>ThinQQ loader $loaderZipVersion (Release)</dc:description></cp:coreProperties>")
        $w.Dispose(); $z.Dispose()
    } else { Set-Content (Join-Path $p 'ThinQQ.xlam') $loaderText }
    Set-Content (Join-Path $p 'ThinQQ.core.xlam') 'core'
    Set-Content (Join-Path $p 'tutorial.html') '<html></html>'
    @{ coreVersion = '1.2.3'; loaderVersion = $loaderVer; manifestUrl = 'https://exemplo/manifest.json'
       coreSha256 = (Sha $p 'ThinQQ.core.xlam'); loaderSha256 = (Sha $p 'ThinQQ.xlam') } | ConvertTo-Json | Set-Content (Join-Path $p 'manifest.json')
    return $p
}
function Set-Manifest($pkg, $coreHash, $loaderHash) {
    @{ coreVersion = '1.2.3'; loaderVersion = '1.0.0'; manifestUrl = 'https://exemplo/manifest.json'
       coreSha256 = $coreHash; loaderSha256 = $loaderHash } | ConvertTo-Json | Set-Content (Join-Path $pkg 'manifest.json')
}
function Snapshot($key) {
    $k = Get-Item -LiteralPath $key
    ($k.GetValueNames() | Sort-Object | ForEach-Object { "$_|$($k.GetValueKind($_))|$($k.GetValue($_, $null, 'DoNotExpandEnvironmentNames'))" }) -join "`n"
}
$optSub = 'Microsoft\Office\16.0\Excel\Options'
$tlSub  = 'Microsoft\Office\16.0\Excel\Security\Trusted Locations\LocationThinQQ'
$upSub  = 'VB and VBA Program Settings\ThinQQ\Updater'

try {
    $pkg = New-Pkg 'pkg'
    $root = Join-Path $tmp 'root\ThinQQ'
    $reg = New-Reg
    $optKey = Join-Path $reg $optSub; $upKey = Join-Path $reg $upSub

    # ---- Protecoes de caminho e de registro ----
    $bad = Join-Path $tmp 'outraPasta'
    New-Item -ItemType Directory -Force $bad | Out-Null
    Set-Content (Join-Path $bad 'importante.txt') 'nao apagar'
    Expect-Failure 'install: Root fora do padrao ThinQQ recusado' { & $install -Root $bad -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck } 'recusada'
    Expect-Failure 'uninstall: Root fora do padrao ThinQQ recusado' { & $uninstall -Root $bad -RegistryRoot $reg -SkipExcelCheck } 'recusada'
    Assert (Test-Path (Join-Path $bad 'importante.txt')) 'pasta recusada continua intacta'
    Expect-Failure 'install: raiz de disco recusada' { & $install -Root 'C:\' -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck } 'recusada'
    Expect-Failure 'uninstall: raiz de disco recusada' { & $uninstall -Root 'C:\' -RegistryRoot $reg -SkipExcelCheck } 'recusada'
    Expect-Failure 'uninstall: pasta do perfil recusada' { & $uninstall -Root $env:USERPROFILE -RegistryRoot $reg -SkipExcelCheck } 'recusada'
    $fake = Join-Path $tmp 'fake\ThinQQ'
    New-Item -ItemType Directory -Force $fake | Out-Null
    Set-Content (Join-Path $fake 'documento.txt') 'nao e do ThinQQ'
    Expect-Failure 'uninstall: pasta ThinQQ sem ThinQQ.xlam/versions recusada' { & $uninstall -Root $fake -RegistryRoot $reg -SkipExcelCheck } 'não parece'
    Assert (Test-Path (Join-Path $fake 'documento.txt')) 'pasta sem marcadores intacta'
    Expect-Failure 'install: HKCU:\Software\Microsoft recusado' { & $install -Root $root -RegistryRoot 'HKCU:\Software\Microsoft' -PackageDir $pkg -SkipExcelCheck } 'registro recusada'
    Expect-Failure 'install: HKLM recusado' { & $install -Root $root -RegistryRoot 'HKLM:\Software' -PackageDir $pkg -SkipExcelCheck } 'registro recusada'
    Expect-Failure 'uninstall: HKLM recusado' { & $uninstall -Root $root -RegistryRoot 'HKLM:\Software' -SkipExcelCheck } 'registro recusada'
    Assert (-not (Test-Path $root)) 'recusas nao criaram nada'

    # ---- Instalacao recusada: hash divergente, arquivo ausente (nada instalado) ----
    Set-Manifest $pkg 'deadbeef' (Sha $pkg 'ThinQQ.xlam')
    Expect-Failure 'hash do core divergente deve recusar' { & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck } 'SHA|válido'
    Set-Manifest $pkg (Sha $pkg 'ThinQQ.core.xlam') ('0' * 64)
    Expect-Failure 'hash do loader divergente deve recusar' { & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck } 'SHA-256'
    Assert (-not (Test-Path $root)) 'nada copiado apos hash divergente'
    Assert (-not (Test-Path $reg)) 'registro intacto apos hash divergente'
    Set-Manifest $pkg (Sha $pkg 'ThinQQ.core.xlam') (Sha $pkg 'ThinQQ.xlam')
    Move-Item (Join-Path $pkg 'ThinQQ.core.xlam') (Join-Path $tmp 'core.bak')
    Expect-Failure 'arquivo ausente deve recusar' { & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck } 'não encontrado'
    Assert (-not (Test-Path $root)) 'nada copiado com arquivo ausente'
    Move-Item (Join-Path $tmp 'core.bak') (Join-Path $pkg 'ThinQQ.core.xlam')

    # ---- Instalacao normal ----
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck
    Assert (Test-Path (Join-Path $root 'ThinQQ.xlam')) 'loader copiado'
    Assert (Test-Path (Join-Path $root 'versions\1.2.3\ThinQQ.core.xlam')) 'core copiado na pasta da versao'
    Assert (Test-Path (Join-Path $root 'versions\1.2.3\tutorial.html')) 'tutorial copiado'
    Assert (-not (Test-Path (Join-Path $root '_staging'))) 'staging removido'
    Assert (-not (Test-Path (Join-Path $root 'ThinQQ.xlam.bak'))) 'sem .bak na primeira instalacao'
    $opt = Get-ItemProperty $optKey
    Assert ($opt.OPEN -eq ('/R "' + (Join-Path $root 'ThinQQ.xlam') + '"')) "valor OPEN: $($opt.OPEN)"
    Assert ((Get-ItemProperty (Join-Path $reg $tlSub)).AllowSubfolders -eq 1) 'local confiavel com subpastas'
    Assert ((Get-ItemProperty $upKey).UpdateSource -eq 'https://exemplo/manifest.json') 'UpdateSource do manifesto'
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck
    Assert (-not ((Get-ItemProperty $optKey).PSObject.Properties.Name -contains 'OPEN1')) 'OPEN nao duplicado'
    Set-ItemProperty $upKey -Name UpdateSource -Value 'D:\minha\fonte\manifest.json'
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck
    Assert ((Get-ItemProperty $upKey).UpdateSource -eq 'D:\minha\fonte\manifest.json') 'UpdateSource existente preservado'
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck -UpdateSource '\\srv\tqq\manifest.json'
    Assert ((Get-ItemProperty $upKey).UpdateSource -eq '\\srv\tqq\manifest.json') 'UpdateSource explicito'

    # ---- Backup do loader e recusa de loader mais antigo ----
    $pkgB = New-Pkg 'pkgB' 'loader-novo'
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkgB -SkipExcelCheck
    Assert ((Get-Content (Join-Path $root 'ThinQQ.xlam.bak') -Raw).Trim() -eq 'loader') '.bak guarda o loader anterior'
    Assert ((Get-Content (Join-Path $root 'ThinQQ.xlam') -Raw).Trim() -eq 'loader-novo') 'loader novo no lugar'
    $pkgV2 = New-Pkg 'pkgV2' '' '2.0.0' '2.0.0'
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkgV2 -SkipExcelCheck
    $pkgV1 = New-Pkg 'pkgV1' '' '1.0.0' '1.0.0'
    Expect-Failure 'loader mais antigo recusado' { & $install -Root $root -RegistryRoot $reg -PackageDir $pkgV1 -SkipExcelCheck } 'mais novo'
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkgV1 -SkipExcelCheck -Force
    Assert ((Get-Item (Join-Path $root 'ThinQQ.xlam')).Length -eq (Get-Item (Join-Path $pkgV1 'ThinQQ.xlam')).Length) '-Force sobrescreve'
    & $uninstall -Root $root -RegistryRoot $reg -SkipExcelCheck
    Assert (-not (Test-Path $root)) 'arquivos removidos'
    Assert (-not (Test-Path (Join-Path $reg $tlSub))) 'local confiavel removido'
    Assert (-not (Test-Path (Join-Path $reg 'VB and VBA Program Settings\ThinQQ'))) 'configuracoes removidas'

    # ---- Outros valores de Options sobrevivem; entrada nova vira OPEN2 ----
    $reg2 = New-Reg; $opt2 = Join-Path $reg2 $optSub
    New-Item -Path $opt2 -Force | Out-Null
    Set-ItemProperty $opt2 -Name OPEN -Value '/R "C:\outro\A.xlam"'
    Set-ItemProperty $opt2 -Name OPEN1 -Value '/R "C:\outro\B.xlam"'
    Set-ItemProperty $opt2 -Name Maximized -Value 3 -Type DWord
    $root2 = Join-Path $tmp 'r2\ThinQQ'
    & $install -Root $root2 -RegistryRoot $reg2 -PackageDir $pkg -SkipExcelCheck
    $o = Get-ItemProperty $opt2
    Assert ($o.OPEN -eq '/R "C:\outro\A.xlam"' -and $o.OPEN1 -eq '/R "C:\outro\B.xlam"') 'OPEN/OPEN1 de outros suplementos preservados'
    Assert ($o.Maximized -eq 3) 'valor nao relacionado preservado'
    Assert ($o.OPEN2 -eq ('/R "' + (Join-Path $root2 'ThinQQ.xlam') + '"')) 'ThinQQ adicionado como OPEN2'
    # uninstall com ThinQQ no fim: restantes intactos
    & $uninstall -Root $root2 -RegistryRoot $reg2 -SkipExcelCheck
    $o = Get-ItemProperty $opt2
    Assert ($o.OPEN -eq '/R "C:\outro\A.xlam"' -and $o.OPEN1 -eq '/R "C:\outro\B.xlam"' -and $o.Maximized -eq 3) 'restantes intactos apos uninstall'
    Assert (-not ($o.PSObject.Properties.Name -contains 'OPEN2')) 'OPEN2 removido'

    # ---- Entrada antiga atualizada no lugar; MyThinQQ.xlam intocado ----
    $reg3 = New-Reg; $opt3 = Join-Path $reg3 $optSub
    New-Item -Path $opt3 -Force | Out-Null
    Set-ItemProperty $opt3 -Name OPEN -Value '/R "C:\x\MyThinQQ.xlam"'
    Set-ItemProperty $opt3 -Name OPEN1 -Value '/R "D:\antigo\thinqq.xlam"'
    $root3 = Join-Path $tmp 'r3\ThinQQ'
    & $install -Root $root3 -RegistryRoot $reg3 -PackageDir $pkg -SkipExcelCheck
    $o = Get-ItemProperty $opt3
    Assert ($o.OPEN -eq '/R "C:\x\MyThinQQ.xlam"') 'MyThinQQ.xlam intocado'
    Assert ($o.OPEN1 -eq ('/R "' + (Join-Path $root3 'ThinQQ.xlam') + '"')) "entrada antiga atualizada: $($o.OPEN1)"
    Assert (-not ($o.PSObject.Properties.Name -contains 'OPEN2')) 'sem duplicata'
    # uninstall: MyThinQQ continua (nome de arquivo diferente)
    & $uninstall -Root $root3 -RegistryRoot $reg3 -SkipExcelCheck
    $o = Get-ItemProperty $opt3
    Assert ($o.OPEN -eq '/R "C:\x\MyThinQQ.xlam"') 'MyThinQQ.xlam preservado no uninstall'
    Assert (-not ($o.PSObject.Properties.Name -contains 'OPEN1')) 'entrada do ThinQQ removida'

    # ---- Uninstall com ThinQQ no meio: ordem preservada e sem lacunas, tipo preservado ----
    $reg4 = New-Reg; $opt4 = Join-Path $reg4 $optSub
    New-Item -Path $opt4 -Force | Out-Null
    $root4 = Join-Path $tmp 'r4\ThinQQ'
    New-Item -ItemType Directory -Force $root4 | Out-Null
    Set-Content (Join-Path $root4 'ThinQQ.xlam') 'x'
    Set-ItemProperty $opt4 -Name OPEN -Value '/R "%TEMP%\A.xlam"' -Type ExpandString
    Set-ItemProperty $opt4 -Name OPEN1 -Value ('/R "' + (Join-Path $root4 'ThinQQ.xlam') + '"')
    Set-ItemProperty $opt4 -Name OPEN2 -Value '/R "C:\outro\B.xlam"'
    & $uninstall -Root $root4 -RegistryRoot $reg4 -SkipExcelCheck
    $k = Get-Item $opt4
    Assert ($k.GetValue('OPEN', $null, 'DoNotExpandEnvironmentNames') -eq '/R "%TEMP%\A.xlam"') 'OPEN preservado'
    Assert ($k.GetValueKind('OPEN').ToString() -eq 'ExpandString') 'tipo do valor preservado'
    Assert ($k.GetValue('OPEN1') -eq '/R "C:\outro\B.xlam"') 'OPEN2 renumerado para OPEN1'
    Assert ($k.GetValueNames() -notcontains 'OPEN2') 'sem lacuna: OPEN2 removido'
    # sem entrada do ThinQQ: Options byte a byte igual
    $before = Snapshot $opt4
    & $uninstall -Root $root4 -RegistryRoot $reg4 -SkipExcelCheck
    Assert ((Snapshot $opt4) -eq $before) 'Options inalterado sem entrada do ThinQQ'

    # ---- Falha injetada apos a copia: nada fica ----
    $reg5 = New-Reg; $root5 = Join-Path $tmp 'r5\ThinQQ'
    Expect-Failure 'falha injetada propaga' { & $install -Root $root5 -RegistryRoot $reg5 -PackageDir $pkg -SkipExcelCheck -SimulateFailureForTest } 'simulada'
    Assert (-not (Test-Path $root5)) 'sem loader/pasta apos falha em instalacao nova'
    Assert (-not (Test-Path (Join-Path $reg5 $tlSub))) 'sem local confiavel apos falha'
    # falha com instalacao previa: loader anterior restaurado
    & $install -Root $root5 -RegistryRoot $reg5 -PackageDir $pkg -SkipExcelCheck
    Expect-Failure 'falha injetada (reinstalacao)' { & $install -Root $root5 -RegistryRoot $reg5 -PackageDir $pkgB -SkipExcelCheck -SimulateFailureForTest } 'simulada'
    Assert ((Get-Content (Join-Path $root5 'ThinQQ.xlam') -Raw).Trim() -eq 'loader') 'loader anterior restaurado apos falha'
    Assert (-not (Test-Path (Join-Path $root5 '_staging'))) 'staging limpo apos falha'
    & $uninstall -Root $root5 -RegistryRoot $reg5 -SkipExcelCheck

    # ---- Caminho com acentos (instalar e desinstalar) ----
    $reg6 = New-Reg; $root6 = Join-Path $tmp 'Aplicações\ThinQQ'
    & $install -Root $root6 -RegistryRoot $reg6 -PackageDir $pkg -SkipExcelCheck
    Assert (Test-Path (Join-Path $root6 'versions\1.2.3\ThinQQ.core.xlam')) 'instalou em caminho com acento'
    Assert ((Get-ItemProperty (Join-Path $reg6 $optSub)).OPEN -eq ('/R "' + (Join-Path $root6 'ThinQQ.xlam') + '"')) 'OPEN com acento'
    & $uninstall -Root $root6 -RegistryRoot $reg6 -SkipExcelCheck
    Assert (-not (Test-Path $root6)) 'desinstalou caminho com acento'

    Write-Host "Test-Installer: OK (PowerShell $($PSVersionTable.PSVersion))" -ForegroundColor Green
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    foreach ($r in $regs) { Remove-Item -LiteralPath $r -Recurse -Force -ErrorAction SilentlyContinue }
}
