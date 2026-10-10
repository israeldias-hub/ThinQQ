<#
.SYNOPSIS  Monta ThinQQ.xlam (loader) e ThinQQ.core.xlam (core) a partir de src/.
.EXAMPLE   pwsh build/build.ps1 -Target all -Configuration Release -Install
#>
param(
    [ValidateSet('loader','core','all')] [string]$Target = 'all',
    [ValidateSet('Release','Test')]      [string]$Configuration = 'Release',
    [switch]$Install
)
. "$PSScriptRoot\lib\Build-Common.ps1"
Assert-VbomTrusted
$versions = Get-Content (Join-Path $RepoRoot 'version.json') -Raw | ConvertFrom-Json
New-Item -ItemType Directory -Force $DistDir | Out-Null
$targets = if ($Target -eq 'all') { @('loader','core') } else { @($Target) }

foreach ($t in $targets) {
    $srcDir  = Join-Path $RepoRoot "src\$t"
    $version = $versions.$t
    $suffix  = if ($Configuration -eq 'Test') { '.test' } else { '' }
    $outName = if ($t -eq 'loader') { "ThinQQ$suffix.xlam" } else { "ThinQQ.core$suffix.xlam" }
    $outPath = Join-Path $DistDir $outName
    if (Test-Path $outPath) { Remove-Item $outPath -Force }

    $buildInfo = Join-Path $env:TEMP "modBuildInfo_$t.bas"
    Write-BuildInfoModule -Version $version -Configuration $Configuration -OutPath $buildInfo

    $xl = New-ExcelApp
    try {
        $wb = $xl.Workbooks.Add()
        while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
        $wb.Worksheets.Item(1).Name = 'ThinQQ'
        $vbproj = $wb.VBProject
        $vbproj.Name = "ThinQQ_$t"

        Import-VbaFile $vbproj $buildInfo
        Get-ChildItem (Join-Path $srcDir 'modules') -Filter *.bas | ForEach-Object { Import-VbaFile $vbproj $_.FullName }
        $clsDir = Join-Path $srcDir 'classes'
        if (Test-Path $clsDir) { Get-ChildItem $clsDir -Filter *.cls | ForEach-Object { Import-VbaFile $vbproj $_.FullName } }
        $formsDir = Join-Path $srcDir 'forms'
        if (Test-Path $formsDir) {
            Get-ChildItem $formsDir -Filter *.json | ForEach-Object {
                Add-FormFromJson $vbproj $_.FullName (Join-Path $formsDir ($_.BaseName + '.vba'))
            }
        }
        $twb = Join-Path $srcDir 'ThisWorkbook.vba'
        # CodeName: 'ThisWorkbook' no Excel en, 'EstaPasta_de_trabalho' no pt-BR
        if (Test-Path $twb) { Add-CodeToComponent $vbproj $wb.CodeName $twb }

        if ($Configuration -eq 'Test') {
            Get-ChildItem (Join-Path $RepoRoot 'tests\vba\common') -Filter *.bas | ForEach-Object { Import-VbaFile $vbproj $_.FullName }
            $tDir = Join-Path $RepoRoot "tests\vba\$t"
            if (Test-Path $tDir) { Get-ChildItem $tDir -Filter *.bas | ForEach-Object { Import-VbaFile $vbproj $_.FullName } }
        }

        Set-DocProperty $wb 'Title' 'ThinQQ'
        Set-DocProperty $wb 'Comments' "ThinQQ $t $version ($Configuration)"
        $wb.IsAddin = $true
        $wb.SaveAs($outPath, 55)   # 55 = xlOpenXMLAddIn
        $wb.Close($false)
    } finally { Close-ExcelApp $xl }

    if ($t -eq 'loader') { Add-RibbonToPackage -PackagePath $outPath -RibbonXmlPath (Join-Path $srcDir 'ribbon\customUI14.xml') }
    Write-Host "OK  $outPath"
}

if ($Install -and $Configuration -eq 'Release') {
    $root = Join-Path $env:LOCALAPPDATA 'ThinQQ'
    $verDir = Join-Path $root "versions\$($versions.core)"
    New-Item -ItemType Directory -Force $verDir, (Join-Path $root 'tmp'), (Join-Path $root 'logs') | Out-Null
    Copy-Item (Join-Path $DistDir 'ThinQQ.xlam') $root -Force
    Copy-Item (Join-Path $DistDir 'ThinQQ.core.xlam') $verDir -Force
    Write-Host "Instalado em $root (abra o Excel; na primeira vez registre o suplemento em Arquivo > Opções > Suplementos > Ir... > Procurar)"
}
