@echo off
setlocal
set SRC=%~dp0
set DEST=C:\_DesktopReal\08.clausula

echo Atualizando modulo 03_Valor_Global.py na homologacao...
if not exist "%DEST%" (
  echo ERRO: Pasta de homologacao nao encontrada: %DEST%
  pause
  exit /b 1
)
if not exist "%DEST%\pages" mkdir "%DEST%\pages"
copy /Y "%SRC%pages\03_Valor_Global.py" "%DEST%\pages\03_Valor_Global.py" >nul
if errorlevel 1 (
  echo ERRO ao copiar o arquivo 03_Valor_Global.py.
  pause
  exit /b 1
)
echo Arquivo 03_Valor_Global.py atualizado com sucesso.
echo Rode: streamlit run app.py
pause
endlocal
