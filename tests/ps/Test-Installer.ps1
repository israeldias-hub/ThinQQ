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
function New-Pkg([string]$name, [string]$loaderText = 'loader', [string]$loaderVer = '1.0.0', [string]$loaderZipVersion = '', [switch]$WithPpam) {
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
    $m = @{ coreVersion = '1.2.3'; loaderVersion = $loaderVer; manifestUrl = 'https://exemplo/manifest.json'
       coreSha256 = (Sha $p 'ThinQQ.core.xlam'); loaderSha256 = (Sha $p 'ThinQQ.xlam') }
    if ($WithPpam) {
        Set-Content (Join-Path $p 'ThinQQ.ppam') 'ppam'
        $m.pptAddinVersion = '1.2.3'; $m.pptAddinSha256 = (Sha $p 'ThinQQ.ppam')
    }
    $m | ConvertTo-Json | Set-Content (Join-Path $p 'manifest.json')
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
$pptAddSub = 'Microsoft\Office\16.0\PowerPoint\AddIns\ThinQQ'
$pptTlSub = 'Microsoft\Office\16.0\PowerPoint\Security\Trusted Locations\LocationThinQQ'

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
    # manifesto de um release publicado no GitHub (New-ReleaseManifest com BaseUrl http): quem instala pelo ZIP passa a consultar o endereco
    # ESTAVEL (release/manifest.json da main), e nao o asset versionado da release
    . (Join-Path $repo 'build\lib\Release-Lib.ps1')
    $gm = New-ReleaseManifest -CoreVersion '1.2.3' -LoaderVersion '1.0.0' -BaseUrl 'https://github.com/x/ThinQQ/releases/download/v1.2.3' -CoreSha (Sha $pkg 'ThinQQ.core.xlam') -LoaderSha (Sha $pkg 'ThinQQ.xlam') -Notes '' -PublishedAt '2026-10-10'
    $gm | ConvertTo-Json | Set-Content (Join-Path $pkg 'manifest.json')
    Remove-ItemProperty $upKey -Name UpdateSource
    & $install -Root $root -RegistryRoot $reg -PackageDir $pkg -SkipExcelCheck
    Assert ((Get-ItemProperty $upKey).UpdateSource -eq (Get-DefaultUpdateSource)) "UpdateSource do ZIP do GitHub = endereco estavel: $((Get-ItemProperty $upKey).UpdateSource)"
    Assert ((Get-ItemProperty $upKey).UpdateSource -notmatch '/releases/download/') 'nao fica preso ao asset versionado'
    Set-Manifest $pkg (Sha $pkg 'ThinQQ.core.xlam') (Sha $pkg 'ThinQQ.xlam')

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

    # ---- Suplemento do PowerPoint (1.4.0): copiado, registrado (AddIns\ThinQQ: Path e AutoLoad) e local confiavel; pacote sem .ppam continua valido
    $regP = New-Reg; $rootP = Join-Path $tmp 'rp\ThinQQ'
    $pkgP = New-Pkg 'pkgP' -WithPpam
    & $install -Root $rootP -RegistryRoot $regP -PackageDir $pkgP -SkipExcelCheck
    Assert (Test-Path (Join-Path $rootP 'ThinQQ.ppam')) 'ppam copiado para a raiz'
    $ak = Get-ItemProperty (Join-Path $regP $pptAddSub)
    Assert ($ak.Path -eq (Join-Path $rootP 'ThinQQ.ppam')) "AddIns\ThinQQ\Path: $($ak.Path)"
    Assert ($ak.AutoLoad -eq 1) 'AutoLoad = 1'
    Assert ((Get-ItemProperty (Join-Path $regP $pptTlSub)).Path -eq ($rootP + '\')) 'local confiavel do PowerPoint'
    Assert ((Get-ItemProperty (Join-Path $regP $upSub)).PpamVersion -eq '1.2.3') 'PpamVersion gravada (o atualizador do Excel compara com pptAddinVersion)'
    Set-ItemProperty (Join-Path $regP $upSub) -Name PendingPpam -Value '1'
    Set-Content (Join-Path $rootP 'ThinQQ.ppam.new') 'velho'
    Set-Content (Join-Path $rootP 'ThinQQ.ppam.old') 'sobra de troca interrompida'
    & $install -Root $rootP -RegistryRoot $regP -PackageDir $pkgP -SkipExcelCheck
    Assert ((Get-ItemProperty (Join-Path $regP $pptAddSub)).AutoLoad -eq 1) 'reinstalacao idempotente'
    Assert (-not (Test-Path (Join-Path $rootP 'ThinQQ.ppam.new'))) 'reinstalar descarta a pendencia antiga do atualizador'
    Assert (-not (Test-Path (Join-Path $rootP 'ThinQQ.ppam.old'))) 'reinstalar limpa a sobra ThinQQ.ppam.old'
    Assert ($null -eq (Get-ItemProperty (Join-Path $regP $upSub)).PendingPpam) 'PendingPpam limpo'
    & $uninstall -Root $rootP -RegistryRoot $regP -SkipExcelCheck
    Assert (-not (Test-Path (Join-Path $regP $pptAddSub))) 'AddIns\ThinQQ removido'
    Assert (-not (Test-Path (Join-Path $regP $pptTlSub))) 'local confiavel do PowerPoint removido'
    # pacote sem .ppam (ZIP antigo): instala o Excel e nao cria chaves do PowerPoint
    $regQ = New-Reg; $rootQ = Join-Path $tmp 'rq\ThinQQ'
    & $install -Root $rootQ -RegistryRoot $regQ -PackageDir $pkg -SkipExcelCheck
    Assert (-not (Test-Path (Join-Path $rootQ 'ThinQQ.ppam'))) 'sem ppam no pacote: nada'
    Assert (-not (Test-Path (Join-Path $regQ $pptAddSub))) 'sem ppam: sem AddIns'
    & $uninstall -Root $rootQ -RegistryRoot $regQ -SkipExcelCheck
    # SHA divergente do ppam recusa e nada fica; manifesto com o campo e sem o arquivo recusa
    $pkgX = New-Pkg 'pkgX' -WithPpam
    Set-Content (Join-Path $pkgX 'ThinQQ.ppam') 'adulterado'
    $regX = New-Reg; $rootX = Join-Path $tmp 'rx\ThinQQ'
    Expect-Failure 'ppam com SHA divergente recusado' { & $install -Root $rootX -RegistryRoot $regX -PackageDir $pkgX -SkipExcelCheck } 'SHA-256 de ThinQQ.ppam'
    Assert (-not (Test-Path $rootX)) 'nada copiado com ppam adulterado'
    Remove-Item (Join-Path $pkgX 'ThinQQ.ppam')
    Expect-Failure 'manifesto cita o ppam ausente' { & $install -Root $rootX -RegistryRoot $regX -PackageDir $pkgX -SkipExcelCheck } 'ThinQQ.ppam'
    # falha injetada depois da copia: sem ppam nem chaves do PowerPoint
    $regY = New-Reg; $rootY = Join-Path $tmp 'ry\ThinQQ'
    Expect-Failure 'falha injetada (ppam)' { & $install -Root $rootY -RegistryRoot $regY -PackageDir $pkgP -SkipExcelCheck -SimulateFailureForTest } 'simulada'
    Assert (-not (Test-Path $rootY)) 'falha: nada fica'
    Assert (-not (Test-Path (Join-Path $regY $pptAddSub))) 'falha: sem AddIns'
    Assert (-not (Test-Path (Join-Path $regY $pptTlSub))) 'falha (instalacao nova): sem local confiavel do PowerPoint'
    $upY = Join-Path $regY $upSub
    Assert ((-not (Test-Path $upY)) -or ($null -eq (Get-ItemProperty $upY).PpamVersion)) 'falha (instalacao nova): sem PpamVersion'
    # falha injetada numa REINSTALACAO: o ppam anterior volta e a PpamVersion antiga e restaurada
    $regZ = New-Reg; $rootZ = Join-Path $tmp 'rz\ThinQQ'
    & $install -Root $rootZ -RegistryRoot $regZ -PackageDir $pkgP -SkipExcelCheck
    $pkgP2 = New-Pkg 'pkgP2' -WithPpam
    Set-Content (Join-Path $pkgP2 'ThinQQ.ppam') 'ppam-novo'
    @{ coreVersion = '1.2.3'; loaderVersion = '1.0.0'; manifestUrl = 'https://exemplo/manifest.json'; pptAddinVersion = '1.9.9'
       coreSha256 = (Sha $pkgP2 'ThinQQ.core.xlam'); loaderSha256 = (Sha $pkgP2 'ThinQQ.xlam'); pptAddinSha256 = (Sha $pkgP2 'ThinQQ.ppam') } | ConvertTo-Json | Set-Content (Join-Path $pkgP2 'manifest.json')
    Expect-Failure 'falha injetada (reinstalacao com ppam)' { & $install -Root $rootZ -RegistryRoot $regZ -PackageDir $pkgP2 -SkipExcelCheck -SimulateFailureForTest } 'simulada'
    Assert ((Get-Content (Join-Path $rootZ 'ThinQQ.ppam') -Raw).Trim() -eq 'ppam') 'ppam anterior restaurado apos falha'
    Assert ((Get-ItemProperty (Join-Path $regZ $upSub)).PpamVersion -eq '1.2.3') 'PpamVersion anterior restaurada apos falha'
    Assert ((Get-ItemProperty (Join-Path $regZ $pptAddSub)).AutoLoad -eq 1) 'AddIns\ThinQQ anterior mantido apos falha'
    & $install -Root $rootZ -RegistryRoot $regZ -PackageDir $pkgP2 -SkipExcelCheck
    Assert ((Get-Content (Join-Path $rootZ 'ThinQQ.ppam') -Raw).Trim() -eq 'ppam-novo') 'reinstalacao troca o ppam'
    Assert ((Get-ItemProperty (Join-Path $regZ $upSub)).PpamVersion -eq '1.9.9') 'PpamVersion nova'
    # .ppam instalado EM USO (PowerPoint aberto): a reinstalacao falha com orientacao clara e deixa tudo como estava.
    # FileShare.None: nem a copia de seguranca consegue ler; FileShare.Read: le, mas a troca falha DEPOIS de loader e core terem sido movidos (exercita o desfazer).
    $pkgL = New-Pkg 'pkgL' 'loader-l' -WithPpam
    Set-Content (Join-Path $pkgL 'ThinQQ.core.xlam') 'core-l'
    Set-Content (Join-Path $pkgL 'ThinQQ.ppam') 'ppam-l'
    @{ coreVersion = '1.2.3'; loaderVersion = '1.0.0'; manifestUrl = 'https://exemplo/manifest.json'; pptAddinVersion = '1.7.7'
       coreSha256 = (Sha $pkgL 'ThinQQ.core.xlam'); loaderSha256 = (Sha $pkgL 'ThinQQ.xlam'); pptAddinSha256 = (Sha $pkgL 'ThinQQ.ppam') } | ConvertTo-Json | Set-Content (Join-Path $pkgL 'manifest.json')
    $ppamZ = Join-Path $rootZ 'ThinQQ.ppam'; $loaderZ = Join-Path $rootZ 'ThinQQ.xlam'; $coreZ = Join-Path $rootZ 'versions\1.2.3\ThinQQ.core.xlam'
    # a copia de seguranca do loader (ThinQQ.xlam.bak, so a ultima) de uma instalacao anterior: com o ppam ilegivel (FileShare.None) o
    # instalador aborta ANTES de tocar em qualquer arquivo, inclusive o .bak
    Set-Content -LiteralPath ($loaderZ + '.bak') 'bak-antigo'
    foreach ($share in @('None', 'Read')) {
        $shaBefore = @{ ppam = (Sha $rootZ 'ThinQQ.ppam'); loader = (Sha $rootZ 'ThinQQ.xlam'); core = (Sha $rootZ 'versions\1.2.3\ThinQQ.core.xlam') }
        $lock = [IO.File]::Open($ppamZ, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]$share)
        try {
            Expect-Failure "reinstalar com o ppam em uso (FileShare.$share)" { & $install -Root $rootZ -RegistryRoot $regZ -PackageDir $pkgL -SkipExcelCheck } 'Feche o PowerPoint[^\r\n]*Detalhe: \S'
        } finally { $lock.Dispose() }
        if ($share -eq 'None') { Assert ((Get-Content -LiteralPath ($loaderZ + '.bak') -Raw).Trim() -eq 'bak-antigo') 'ppam ilegivel: o ThinQQ.xlam.bak anterior continua intacto (aborta antes de copiar o loader)' }
        Assert ((Sha $rootZ 'ThinQQ.ppam') -eq $shaBefore.ppam) "ppam em uso ($share): conteudo nao mudou"
        Assert ((Sha $rootZ 'ThinQQ.xlam') -eq $shaBefore.loader) "ppam em uso ($share): loader voltou ao estado anterior"
        Assert ((Sha $rootZ 'versions\1.2.3\ThinQQ.core.xlam') -eq $shaBefore.core) "ppam em uso ($share): core voltou ao estado anterior"
        Assert ((Get-ItemProperty (Join-Path $regZ $upSub)).PpamVersion -eq '1.9.9') "ppam em uso ($share): PpamVersion intacta"
        Assert (-not (Test-Path (Join-Path $rootZ '_staging'))) "ppam em uso ($share): staging limpo"
    }
    & $install -Root $rootZ -RegistryRoot $regZ -PackageDir $pkgL -SkipExcelCheck
    Assert ((Get-Content $ppamZ -Raw).Trim() -eq 'ppam-l') 'liberado o arquivo, a reinstalacao troca o ppam'
    # o desinstalador remove SO as chaves do ThinQQ: outros suplementos e locais confiaveis do PowerPoint sobrevivem
    $otherAdd = Join-Path $regZ 'Microsoft\Office\16.0\PowerPoint\AddIns\Outro'
    $otherTl = Join-Path $regZ 'Microsoft\Office\16.0\PowerPoint\Security\Trusted Locations\LocationOutro'
    New-Item -Path $otherAdd -Force | Out-Null; Set-ItemProperty $otherAdd -Name Path -Value 'C:\outro\Outro.ppam'
    New-Item -Path $otherTl -Force | Out-Null; Set-ItemProperty $otherTl -Name Path -Value 'C:\outro\'
    & $uninstall -Root $rootZ -RegistryRoot $regZ -SkipExcelCheck
    Assert ((Get-ItemProperty $otherAdd).Path -eq 'C:\outro\Outro.ppam') 'AddIns de outro suplemento do PowerPoint preservado'
    Assert ((Get-ItemProperty $otherTl).Path -eq 'C:\outro\') 'local confiavel de outro do PowerPoint preservado'
    Assert (-not (Test-Path (Join-Path $regZ $pptAddSub))) 'AddIns\ThinQQ removido (reinstalado)'
    Assert (-not (Test-Path (Join-Path $regZ 'VB and VBA Program Settings\ThinQQ'))) 'PpamVersion removida com as configuracoes'

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
