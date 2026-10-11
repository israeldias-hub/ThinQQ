@echo off
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
echo Instalando o ThinQQ...
echo.
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
if errorlevel 1 (
  echo.
  echo ===== FALHA NA INSTALACAO =====
  echo Leia a mensagem de erro acima. Feche o Excel e tente de novo, ou siga a
  echo instalacao manual descrita no LEIA-ME.txt.
  echo.
  pause
  exit /b 1
)
echo.
echo ===== INSTALACAO CONCLUIDA =====
echo Pronto! Abra o Excel e procure a aba ThinQQ.
echo.
pause
