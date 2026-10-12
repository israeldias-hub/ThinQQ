@echo off
rem ============================================================================
rem  criar-pull-request.bat
rem  Envia a branch atual para o GitHub e abre a pagina do pull request.
rem
rem  Uso:  criar-pull-request.bat            (base = main)
rem        criar-pull-request.bat outra-base (base = outra branch)
rem
rem  - Um pull request NAO fica velho: cada push na branch o atualiza sozinho.
rem    Se a branch ja tem um pull request aberto, a pagina mostra "View pull
rem    request"; se nao tem, mostra o formulario para criar.
rem  - Alteracoes nao commitadas nao entram no pull request (o script avisa).
rem  - Sem o GitHub CLI (gh): o script usa o Git e o navegador, como o login do
rem    Git Credential Manager ja guardado nesta maquina.
rem  - Para ensaiar sem enviar nada:  set DRYRUN=1  e depois rode o script.
rem ============================================================================
setlocal EnableExtensions
cd /d "%~dp0"

set "REPO_URL=https://github.com/israeldias-hub/ThinQQ"
set "BASE=%~1"
if "%BASE%"=="" set "BASE=main"

where git >nul 2>nul
if errorlevel 1 (
  echo [ERRO] O Git nao foi encontrado. Instale o Git for Windows e tente de novo.
  goto :fim
)

set "BRANCH="
for /f "delims=" %%B in ('git rev-parse --abbrev-ref HEAD 2^>nul') do set "BRANCH=%%B"
if "%BRANCH%"=="" (
  echo [ERRO] Esta pasta nao e um repositorio Git.
  goto :fim
)
if /i "%BRANCH%"=="HEAD" (
  echo [ERRO] O repositorio esta sem branch, em HEAD solto. Va para uma branch antes.
  goto :fim
)
if /i "%BRANCH%"=="%BASE%" (
  echo [ERRO] Voce esta na propria branch %BASE%. Va para a branch do trabalho, por exemplo: git checkout feat/fase3
  goto :fim
)

echo.
echo Branch do trabalho: %BRANCH%
echo Branch de destino : %BASE%
echo.

set "SUJO="
for /f "delims=" %%S in ('git status --porcelain') do set "SUJO=1"
if defined SUJO (
  echo [AVISO] Ha alteracoes que ainda nao foram commitadas. Elas NAO entram no pull request:
  git status --short
  echo.
  choice /c SN /m "Continuar mesmo assim"
  if errorlevel 2 goto :fim
)

echo Enviando a branch para o GitHub...
if defined DRYRUN (
  echo [TESTE] git push -u origin "%BRANCH%"
) else (
  git push -u origin "%BRANCH%"
  if errorlevel 1 (
    echo.
    echo [ERRO] O envio falhou. Confira a rede e o login do Git. Se o Git pedir uma conta, use israeldias-hub.
    goto :fim
  )
)

set "URL=%REPO_URL%/compare/%BASE%...%BRANCH%?expand=1"
echo.
echo Abrindo a pagina do pull request no navegador:
echo   %URL%
if defined DRYRUN (
  echo [TESTE] start "" "%URL%"
) else (
  start "" "%URL%"
)
echo.
echo Pronto. Se a branch ja tinha um pull request, ele foi atualizado pelo envio.
echo Para criar um novo, clique em "Create pull request" na pagina que abriu.

:fim
echo.
if not defined NOPAUSE pause
endlocal
