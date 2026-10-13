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
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

rem Um PSModulePath herdado do PowerShell 7 (bat aberto num terminal pwsh) esconde
rem cmdlets do Windows PowerShell 5.1: sem ele, o 5.1 usa o padrao da maquina.
set "PSModulePath="
set "EXTRA="
if defined DRYRUN set "EXTRA=-DryRun"
set "ARGS=%*"

rem Aberto com dois cliques (sem argumentos): pergunta a versao e a etapa em vez de so
rem mostrar o uso. Com NOPAUSE (automacao e testes) mantem o comportamento antigo.
if not "%~1"=="" goto :executar
if defined NOPAUSE goto :executar
echo.
echo Qual versao voce quer publicar? Escreva no formato x.y.z, por exemplo 1.2.1.
echo Deixe em branco e tecle Enter para sair.
set "VERSAO="
set /p "VERSAO=Versao: "
if not defined VERSAO goto :fim
echo "!VERSAO!"| findstr /r /c:"^\"[0-9][0-9]*\.[0-9][0-9]*\.[0-9][0-9]*\"$" >nul
if errorlevel 1 (
  echo.
  echo [ERRO] A versao precisa ter o formato x.y.z, so numeros e pontos. Exemplo: 1.2.1
  set "RC=2"
  goto :fim
)
echo.
echo   Etapa 1 = preparar: constroi a release e abre o pull request.
echo   Etapa 2 = publicar: SO depois de mesclar o pull request da etapa 1.
choice /c 12 /m "Qual etapa"
set "ETAPA="
if errorlevel 2 set "ETAPA=publicar"
set "ARGS=!VERSAO! !ETAPA!"

:executar
powershell -NoProfile -ExecutionPolicy Bypass -File tools\Publish-Release.ps1 %ARGS% %EXTRA%
set "RC=%ERRORLEVEL%"

:fim
echo.
if not defined NOPAUSE pause
endlocal & exit /b %RC%
