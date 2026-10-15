<#
.SYNOPSIS  Gera um release: version.json, build Release (xlam e ppam do PowerPoint), manifest.json, zip (com o Excel de demonstração em exemplos/) e (opcional) publicação em pasta.
.EXAMPLE   pwsh build/release.ps1 -Version 1.0.0 -BaseUrl https://github.com/israeldias-hub/ThinQQ/releases/download/v1.0.0
.EXAMPLE   pwsh build/release.ps1 -Version 1.0.0 -BaseUrl '\\servidor\ThinQQ' -PublishTo '\\servidor\ThinQQ'
.PARAMETER ManifestUrl  Fonte de atualizacao gravada no manifesto (o que um usuario que instala pelo ZIP consulta). Padrao: com BaseUrl
           http(s), o endereco estavel DEFAULT_UPDATE_SOURCE do loader (release/manifest.json da main); com pasta de rede/local,
           BaseUrl\manifest.json.
.NOTES     NAO grava release/manifest.json: copie dist/manifest.json para release/manifest.json apenas
           depois de publicar com URLs reais (veja release/README.md).
#>
param(
    [Parameter(Mandatory)][string]$Version,
    [string]$LoaderVersion,
    [Parameter(Mandatory)][string]$BaseUrl,
    [string]$NotesFile = (Join-Path $PSScriptRoot '..\release\notes.txt'),
    [string]$PublishTo,
    [string]$ManifestUrl,
    [string]$SignPfx = '',
    [string]$SignPassword = '',
    [string]$SignTimestampUrl = ''
)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\lib\Build-Common.ps1"
. "$PSScriptRoot\lib\Release-Lib.ps1"
if ($Version -notmatch '^\d+\.\d+\.\d+$') { throw 'Versão deve ser x.y.z' }
if ($LoaderVersion -and $LoaderVersion -notmatch '^\d+\.\d+\.\d+$') { throw 'LoaderVersion deve ser x.y.z' }

# Pre-requisitos (antes de alterar qualquer coisa)
$tutorial = Join-Path $RepoRoot 'docs\tutorial\tutorial.html'
if (-not (Test-Path $tutorial)) { throw "Tutorial não encontrado: $tutorial. Crie docs/tutorial/tutorial.html antes de gerar o release." }
if (Test-TutorialIsStub $tutorial) { throw 'docs/tutorial/tutorial.html ainda é o esqueleto (pequeno demais ou com o marcador THINQQ-TUTORIAL-STUB). Escreva o tutorial real antes do release.' }
$installerFiles = @(Get-ReleaseInstallerFiles)
foreach ($f in $installerFiles) {
    if (-not (Test-Path (Join-Path $RepoRoot "installer\$f"))) { throw "Arquivo do instalador ausente: installer\$f" }
}
# O Excel de demonstracao viaja no ZIP (exemplos/), ao lado de Instalar-ThinQQ.cmd; o instalador nao o instala nem o abre.
$demoSource = Join-Path $RepoRoot "exemplos\$(Get-ReleaseDemoName)"
if (-not (Test-Path -LiteralPath $demoSource)) { throw "Arquivo de demonstracao ausente: exemplos\$(Get-ReleaseDemoName). Gere-o com tools/New-DemoWorkbook.ps1 antes do release." }
if (Test-XlsxHasMacros $demoSource) { throw "exemplos\$(Get-ReleaseDemoName) contem macros (vbaProject): a demonstracao deve ser .xlsx sem macros." }
if (-not (Test-Path -LiteralPath (Join-Path $RepoRoot 'installer\LEIA-ME-exemplos.txt'))) { throw 'Arquivo ausente: installer\LEIA-ME-exemplos.txt (vai no ZIP como exemplos\LEIA-ME.txt).' }
# PowerPoint fechado e VBOM do PowerPoint habilitado (o build do .ppam precisa dos dois): falha aqui, antes de gravar o version.json
Assert-PptBuildReady

# version.json: gravado antes do build (o build le dele); restaurado se qualquer etapa falhar
$versionPath = Join-Path $RepoRoot 'version.json'
$originalVersionJson = [IO.File]::ReadAllText($versionPath)
$versions = $originalVersionJson | ConvertFrom-Json
$changed = $false
if ($versions.core -ne $Version) { $versions.core = $Version; $changed = $true }
if ($LoaderVersion -and $versions.loader -ne $LoaderVersion) { $versions.loader = $LoaderVersion; $changed = $true }
$succeeded = $false
try {
    if ($changed) { [IO.File]::WriteAllText($versionPath, (($versions | ConvertTo-Json) + "`n"), (New-Object Text.UTF8Encoding($false))) }

    $global:LASTEXITCODE = 0
    & "$PSScriptRoot\build.ps1" -Target all -Configuration Release
    if ($global:LASTEXITCODE -ne 0) { throw "Build Release falhou (código $global:LASTEXITCODE)." }
    # Suplemento do PowerPoint (ThinQQ.ppam): o build usa um PowerPoint proprio (VBOM habilitado e PowerPoint fechado). Falhou = nada de release parcial.
    Remove-Item -LiteralPath (Join-Path $DistDir 'ThinQQ.ppam') -Force -ErrorAction SilentlyContinue   # um .ppam de uma execucao anterior nunca pode passar por novo
    $global:LASTEXITCODE = 0
    try { & "$PSScriptRoot\build.ps1" -Target ppt -Configuration Release }
    catch { throw "Build do suplemento do PowerPoint falhou ($($_.Exception.Message)). Habilite o VBOM do PowerPoint e feche o PowerPoint." }
    if ($global:LASTEXITCODE -ne 0) { throw "Build do suplemento do PowerPoint falhou (código $global:LASTEXITCODE). Habilite o VBOM do PowerPoint e feche o PowerPoint." }
    foreach ($f in @('ThinQQ.xlam', 'ThinQQ.core.xlam', 'ThinQQ.ppam')) {
        $p = Join-Path $DistDir $f
        if (-not (Test-Path $p)) { throw "Build não gerou dist\$f" }
        $bad = @(Get-TestMarkersInXlam $p)
        if ($bad.Count -gt 0) { throw "O artefato Release dist\$f contém módulos de teste ($($bad -join ', ')). Release cancelado." }
    }

    # assinatura opcional dos projetos VBA dos .xlam e do .ppam (signtool + Office SIPs); ANTES dos SHA-256 do manifesto, que cobrem os arquivos ja assinados
    if ($SignPfx) {
        $global:LASTEXITCODE = 0
        & "$PSScriptRoot\..\tools\Sign-ThinQQ.ps1" -Files @((Join-Path $DistDir 'ThinQQ.xlam'), (Join-Path $DistDir 'ThinQQ.core.xlam'), (Join-Path $DistDir 'ThinQQ.ppam')) -PfxPath $SignPfx -PfxPassword $SignPassword -TimestampUrl $SignTimestampUrl
        if ($global:LASTEXITCODE -ne 0) { throw "Assinatura dos .xlam e do .ppam falhou (código $global:LASTEXITCODE): o release não foi gerado. Veja tools/Sign-ThinQQ.ps1." }
    }

    $notes = ''
    if (Test-Path $NotesFile) { $notes = (Get-Content $NotesFile -Raw -Encoding UTF8).Trim() }
    $manifest = New-ReleaseManifest -CoreVersion $versions.core -LoaderVersion $versions.loader -BaseUrl $BaseUrl `
        -CoreSha (Get-FileHash (Join-Path $DistDir 'ThinQQ.core.xlam') -Algorithm SHA256).Hash.ToLower() `
        -LoaderSha (Get-FileHash (Join-Path $DistDir 'ThinQQ.xlam') -Algorithm SHA256).Hash.ToLower() `
        -PptSha (Get-FileHash (Join-Path $DistDir 'ThinQQ.ppam') -Algorithm SHA256).Hash.ToLower() `
        -Notes $notes -PublishedAt (Get-Date -Format 'yyyy-MM-dd') -ManifestUrl $ManifestUrl
    [IO.File]::WriteAllText((Join-Path $DistDir 'manifest.json'), (($manifest | ConvertTo-Json) + "`n"), (New-Object Text.UTF8Encoding($false)))
    Copy-Item $tutorial $DistDir -Force
    foreach ($f in $installerFiles) { Copy-Item (Join-Path $RepoRoot "installer\$f") $DistDir -Force }

    $names = @(Get-ReleaseTopFiles)
    $files = $names | ForEach-Object { Join-Path $DistDir $_ }
    $zip = Join-Path $DistDir "ThinQQ-$($versions.core).zip"
    New-ReleaseZip $zip (Get-ReleaseZipLayout $RepoRoot $DistDir)
    $badZip = @(Get-TestMarkersInZip $zip)
    if ($badZip.Count -gt 0) { Remove-Item $zip -Force; throw "O ZIP do release contém módulos de teste ($($badZip -join ', ')). Release cancelado." }
    $zipProblems = @(Test-ReleaseZipContent $zip $demoSource)
    if ($zipProblems.Count -gt 0) { Remove-Item $zip -Force; throw "O conteúdo do ZIP do release está errado ($($zipProblems -join '; ')). Release cancelado." }
    if ($PublishTo) {
        New-Item -ItemType Directory -Force $PublishTo | Out-Null
        Copy-Item $files $PublishTo -Force
        Copy-Item $zip $PublishTo -Force
        Write-Host "Publicado em $PublishTo"
    }
    $succeeded = $true
} finally {
    if (-not $succeeded -and $changed) {
        [IO.File]::WriteAllText($versionPath, $originalVersionJson, (New-Object Text.UTF8Encoding($false)))
        Write-Host 'version.json restaurado (release falhou).' -ForegroundColor Yellow
    }
}
Write-Host "Release pronto: $zip"
Write-Host "Publique ZIP e arquivos de dist/ em $BaseUrl; depois copie dist/manifest.json para release/manifest.json e faça commit dele e de version.json."
