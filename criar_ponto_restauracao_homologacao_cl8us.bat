@echo off
setlocal EnableExtensions EnableDelayedExpansion
set "ORIG=C:\_DesktopReal\08.clausula"

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd_HHmmss"') do set "DATA=%%I"
if "%DATA%"=="" set "DATA=backup_%RANDOM%"

set "DEST=C:\_DesktopReal\08.clausula_restore_%DATA%"

echo Criando ponto de restauracao da homologacao...
echo Origem : %ORIG%
echo Destino: %DEST%

if not exist "%ORIG%" (
  echo ERRO: pasta de homologacao nao encontrada.
  pause
  exit /b 1
)

robocopy "%ORIG%" "%DEST%" /E /XD .venv __pycache__ .git /XF *.pyc >nul
set "RC=%ERRORLEVEL%"
if %RC% GEQ 8 (
  echo ERRO ao criar backup. Codigo robocopy: %RC%
  pause
  exit /b 1
)

echo Backup criado com sucesso em:
echo %DEST%

if exist "%ORIG%\.git" (
  cd /d "%ORIG%"
  git status
  git add .
  git commit -m "RESTORE POINT: homologacao timeline com data-base inicial" 2>nul
  git tag -a v-homologacao-timeline-data-base-%DATA% -m "Ponto de restauracao da homologacao com timeline e data-base inicial" 2>nul
  echo Ponto git local criado/tentado. Nenhum push foi realizado.
) else (
  echo Pasta sem .git. Ponto de restauracao criado como backup de pasta.
)

echo.
echo CONCLUIDO. Este procedimento NAO faz push para producao/nuvem.
pause
endlocal
