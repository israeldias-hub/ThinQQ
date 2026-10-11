<#
.SYNOPSIS  Gera um release: version.json, build Release, manifest.json, zip e (opcional) publicação em pasta.
.EXAMPLE   pwsh build/release.ps1 -Version 1.0.0 -BaseUrl https://github.com/israeldias-hub/ThinQQ/releases/download/v1.0.0
.EXAMPLE   pwsh build/release.ps1 -Version 1.0.0 -BaseUrl '\servidor\ThinQQ' -PublishTo '\servidor\ThinQQ'
.NOTES     NAO grava release/manifest.json: copie dist/manifest.json para release/manifest.json apenas
           depois de publicar com URLs reais (veja release/README.md).
#>
param(
    [Parameter(Mandatory)][string]$Version,
    [string]$LoaderVersion,
    [Parameter(Mandatory)][string]$BaseUrl,
    [string]$NotesFile = (Join-Path $PSScriptRoot '..\release\notes.txt'),
    [string]$PublishTo
)
$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\lib\Build-Common.ps1"
if ($Version -notmatch '^\d+\.\d+\.\d+$') { throw 'Versão deve ser x.y.z' }
if ($LoaderVersion -and $LoaderVersion -notmatch '^\d+\.\d+\.\d+$') { throw 'LoaderVersion deve ser x.y.z' }

# Pre-requisitos (antes de alterar qualquer coisa)
$tutorial = Join-Path $RepoRoot 'docs\tutorial\tutorial.html'
if (-not (Test-Path $tutorial)) { throw "Tutorial não encontrado: $tutorial. Crie docs/tutorial/tutorial.html antes de gerar o release." }
foreach ($f in @('install.ps1', 'uninstall.ps1', 'Instalar-ThinQQ.cmd', 'LEIA-ME.txt')) {
    if (-not (Test-Path (Join-Path $RepoRoot "installer\$f"))) { throw "Arquivo do instalador ausente: installer\$f" }
}

# version.json (so regrava se mudou)
$versionPath = Join-Path $RepoRoot 'version.json'
$versions = Get-Content $versionPath -Raw | ConvertFrom-Json
$changed = $false
if ($versions.core -ne $Version) { $versions.core = $Version; $changed = $true }
if ($LoaderVersion -and $versions.loader -ne $LoaderVersion) { $versions.loader = $LoaderVersion; $changed = $true }
if ($changed) { [IO.File]::WriteAllText($versionPath, (($versions | ConvertTo-Json) + "`n"), (New-Object Text.UTF8Encoding($false))) }

$global:LASTEXITCODE = 0
& "$PSScriptRoot\build.ps1" -Target all -Configuration Release
if ($global:LASTEXITCODE -ne 0) { throw "Build Release falhou (código $global:LASTEXITCODE)." }
foreach ($f in @('ThinQQ.xlam', 'ThinQQ.core.xlam')) {
    if (-not (Test-Path (Join-Path $DistDir $f))) { throw "Build não gerou dist\$f" }
}

# URLs: separador conforme a base (http usa '/', pasta/UNC usa '\')
$sep = '\'
if ($BaseUrl -match '^https?://') { $sep = '/' }
$base = $BaseUrl.TrimEnd('/', '\') + $sep
$notes = ''
if (Test-Path $NotesFile) { $notes = (Get-Content $NotesFile -Raw -Encoding UTF8).Trim() }
$manifest = [ordered]@{
    schemaVersion = 1
    coreVersion   = $versions.core
    coreUrl       = $base + 'ThinQQ.core.xlam'
    coreSha256    = (Get-FileHash (Join-Path $DistDir 'ThinQQ.core.xlam') -Algorithm SHA256).Hash.ToLower()
    loaderVersion = $versions.loader
    loaderUrl     = $base + 'ThinQQ.xlam'
    loaderSha256  = (Get-FileHash (Join-Path $DistDir 'ThinQQ.xlam') -Algorithm SHA256).Hash.ToLower()
    tutorialUrl   = $base + 'tutorial.html'
    manifestUrl   = $base + 'manifest.json'
    publishedAt   = (Get-Date -Format 'yyyy-MM-dd')
    notes         = $notes
}
[IO.File]::WriteAllText((Join-Path $DistDir 'manifest.json'), (($manifest | ConvertTo-Json) + "`n"), (New-Object Text.UTF8Encoding($false)))
Copy-Item $tutorial $DistDir -Force
Copy-Item (Join-Path $RepoRoot 'installer\*') $DistDir -Force

$names = @('ThinQQ.xlam', 'ThinQQ.core.xlam', 'manifest.json', 'tutorial.html', 'install.ps1', 'uninstall.ps1', 'Instalar-ThinQQ.cmd', 'LEIA-ME.txt')
$files = $names | ForEach-Object { Join-Path $DistDir $_ }
$zip = Join-Path $DistDir "ThinQQ-$($versions.core).zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path $files -DestinationPath $zip
if ($PublishTo) {
    New-Item -ItemType Directory -Force $PublishTo | Out-Null
    Copy-Item $files $PublishTo -Force
    Copy-Item $zip $PublishTo -Force
    Write-Host "Publicado em $PublishTo"
}
Write-Host "Release pronto: $zip"
Write-Host "Publique ZIP e arquivos de dist/ em $BaseUrl; depois copie dist/manifest.json para release/manifest.json e faça commit dele e de version.json."
