<#
.SYNOPSIS  Publica uma versao nova do ThinQQ em duas etapas (Windows PowerShell 5.1 ou pwsh 7). Normalmente chamado por publicar-release.bat.
.DESCRIPTION
  Etapa 1, Prepare (padrao): confere as pre-condicoes (arvore limpa, branch main igual a origin/main, versao maior que a do version.json,
  release/notes.txt alterado desde a ultima tag v*, CHANGELOG com entradas em [Nao lancado], Test-Docs e Test-Privacy), cria a branch
  release/v<versao>, atualiza version.json, CHANGELOG, rodape do tutorial e exemplo do LEIA-ME, constroi numa copia em C:\temp com
  build\release.ps1, guarda os 6 arquivos e o state.json em dist\release-<versao>\, faz o commit "release: ThinQQ <versao>" (com
  release/manifest.json), envia a BRANCH e abre a pagina do pull request.
  Etapa 2, Publish (depois de mesclar o pull request com "Create a merge commit"): confere os arquivos contra o state.json, que a
  origin/main contem a preparacao e tem o mesmo release/manifest.json, que a tag e a release nao existem; cria a release em rascunho,
  envia os 6 assets, confere tamanho e SHA-256 de cada um e publica (sem -Draft).
  O token do GitHub vem do login do Git (git credential fill), fica so na memoria e nunca e impresso nem gravado. -DryRun nao le o token.
.EXAMPLE   publicar-release.bat 1.2.1             (etapa 1)
.EXAMPLE   publicar-release.bat 1.2.1 publicar    (etapa 2)
.EXAMPLE   powershell -File tools\Publish-Release.ps1 -Version 1.2.1 -Phase Publish -Draft -NoBrowser
.NOTES     Codigos de saida: 0 ok; 1 falha (a mensagem diz o que ficou feito e como retomar); 2 pre-condicao recusada ou uso errado (nada mudou).
           -RepoRoot e -StateDir existem para os testes e para ensaios; o padrao e este repositorio e dist\release-<versao>.
#>
param(
    [Parameter(Position = 0)][string]$Version = '',
    [Parameter(Position = 1)][ValidateSet('Prepare', 'Publish', 'preparar', 'publicar')][string]$Phase = 'Prepare',
    [string]$Repo = 'israeldias-hub/ThinQQ',
    [string]$NotesFile = 'release/notes.txt',
    [switch]$Draft,
    [switch]$DryRun,
    [switch]$NoBrowser,
    [switch]$RunTests,
    [switch]$AllowMainChanges,
    [string]$StateDir = '',
    [string]$RepoRoot = ''
)
$ErrorActionPreference = 'Stop'
# O script aceita -Debug (e um parametro avancado), e o Invoke-WebRequest do pwsh 7.4+ escreve os cabecalhos (com o token) no fluxo de
# depuracao; Set-PSDebug -Trace mostraria valores. Nenhum dos dois serve aqui: desligados para todo o script.
$DebugPreference = 'SilentlyContinue'
Set-PSDebug -Off
$exitCode = 1
try {
    . (Join-Path $PSScriptRoot 'lib\Release-Publish.ps1')
    if (-not $RepoRoot) { $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path }
    if (-not $Version) {
        Write-Host 'Uso: publicar-release.bat x.y.z            (etapa 1: prepara, constroi e abre o pull request)'
        Write-Host '     publicar-release.bat x.y.z publicar   (etapa 2: depois de mesclar o pull request, publica a release no GitHub)'
        Write-Host '     Opcoes: -DryRun (so mostra o que faria), -Draft (etapa 2 termina em rascunho), -NoBrowser, -RunTests (suite completa antes do build), -AllowMainChanges (etapa 2 aceita uma origin/main que mudou depois da preparacao)'
        $exitCode = 2
    } else {
        $isPublish = ($Phase -eq 'Publish' -or $Phase -eq 'publicar')
        if ($isPublish) {
            if ($RunTests) { Write-Host 'AVISO: -RunTests so vale na etapa Preparar (ignorado).' -ForegroundColor Yellow }
            [void](Invoke-ReleasePublish -RepoRoot $RepoRoot -Version $Version -Repo $Repo -NotesFile $NotesFile -StateDir $StateDir -Draft:$Draft -DryRun:$DryRun -NoBrowser:$NoBrowser -AllowMainChanges:$AllowMainChanges)
        } else {
            if ($Draft) { Write-Host 'AVISO: -Draft so vale na etapa Publicar (ignorado).' -ForegroundColor Yellow }
            [void](Invoke-ReleasePrepare -RepoRoot $RepoRoot -Version $Version -Repo $Repo -NotesFile $NotesFile -DryRun:$DryRun -NoBrowser:$NoBrowser -RunTests:$RunTests)
        }
        $exitCode = 0
    }
} catch {
    $kind = 'falha'
    if ($_.Exception.Data.Contains('RpKind')) { $kind = [string]$_.Exception.Data['RpKind'] }
    $msg = $_.Exception.Message
    if (Get-Command Protect-RpSecret -ErrorAction SilentlyContinue) { $msg = Protect-RpSecret $msg }
    Write-Host ''
    if ($kind -eq 'precondicao') { Write-Host "NAO FOI POSSIVEL CONTINUAR: $msg" -ForegroundColor Red; $exitCode = 2 }
    else { Write-Host "ERRO: $msg" -ForegroundColor Red; $exitCode = 1 }
}
exit $exitCode
