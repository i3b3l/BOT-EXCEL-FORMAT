@echo off
setlocal
title BotPag - Instalador de dependencias

echo ==================================================
echo  BOT PAG - INSTALADOR DE DEPENDENCIAS
echo ==================================================
echo.
echo IMPORTANTE:
echo O BotPag-Formatador.exe ja inclui Python e as bibliotecas
echo necessarias. Este instalador NAO e necessario para executar
echo o EXE. Ele prepara o Python para desenvolvimento/recompilacao.
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py"
    goto install
)
where python >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=python"
    goto install
)

echo ERRO: Python nao foi encontrado neste computador.
echo Instale o Python 3 pelo site oficial https://www.python.org/downloads/
echo e marque a opcao para adicionar Python ao PATH.
echo Depois execute este arquivo novamente.
echo.
pause
exit /b 1

:install
echo Instalando/atualizando pandas, openpyxl e PyInstaller...
%PYTHON_CMD% -m pip install --upgrade pip
if errorlevel 1 goto error
%PYTHON_CMD% -m pip install --upgrade pandas openpyxl pyinstaller
if errorlevel 1 goto error

echo.
echo Dependencias instaladas com sucesso.
echo Elas sao necessarias para executar o codigo-fonte ou recompilar,
echo nao para abrir o BotPag-Formatador.exe ja gerado.
echo.
pause
exit /b 0

:error
echo.
echo Ocorreu um erro durante a instalacao. Confira a mensagem acima
echo e verifique sua conexao com a internet.
echo.
pause
exit /b 1
