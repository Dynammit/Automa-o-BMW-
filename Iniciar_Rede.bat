@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Rede de Concessionarias
set PY=python
where py >nul 2>nul && set PY=py -3
%PY% -c "import pandas, numpy, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo Instalando componentes necessarios ^(so na primeira vez, precisa de internet^)...
  %PY% -m pip install --user pandas numpy openpyxl
)
%PY% -c "import truststore" >nul 2>nul
if errorlevel 1 %PY% -m pip install --user truststore >nul 2>nul
echo.
echo Abrindo o painel no navegador. Nao feche esta janela enquanto estiver usando.
%PY% app_rede.py
echo.
echo O painel foi encerrado.
pause