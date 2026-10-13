@echo off
rem ============================================================================
rem  publicar-release.bat
rem  Publica uma versao nova do ThinQQ no GitHub, em duas etapas.
rem
rem  Uso:  publicar-release.bat 1.2.1           (etapa 1: prepara a versao, constroi
rem                                              e abre a pagina do pull request)
rem        publicar-release.bat 1.2.1 publicar  (etapa 2: depois de mesclar o pull
rem                                              request, cria a release e publica)
rem
rem  - Opcoes extras sao repassadas ao PowerShell: -Draft (a etapa 2 termina em
rem    rascunho), -NoBrowser, -RunTests (suite completa antes do build), -DryRun,
rem    -AllowMainChanges (a etapa 2 publica mesmo com a main alterada depois da
rem    preparacao; a tag fica na main atual).
rem  - Para ensaiar sem gravar nem enviar nada:  set DRYRUN=1  e rode o script.
rem  - Para nao pausar no fim (automacao):       set NOPAUSE=1
rem  - O login e o do Git desta maquina (Git Credential Manager); o token nunca
rem    e mostrado nem gravado. Detalhes: README, secao "Publicar uma versao".
rem ============================================================================
setlocal EnableExtensions
cd /d "%~dp0"

rem Um PSModulePath herdado do PowerShell 7 (bat aberto num terminal pwsh) esconde
rem cmdlets do Windows PowerShell 5.1: sem ele, o 5.1 usa o padrao da maquina.
set "PSModulePath="
set "EXTRA="
if defined DRYRUN set "EXTRA=-DryRun"

powershell -NoProfile -ExecutionPolicy Bypass -File tools\Publish-Release.ps1 %* %EXTRA%
set "RC=%ERRORLEVEL%"

echo.
if not defined NOPAUSE pause
endlocal & exit /b %RC%
