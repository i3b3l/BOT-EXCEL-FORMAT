@echo off
setlocal EnableExtensions
title BOTPAG - FORMATADOR
mode con: cols=90 lines=36
cd /d "%~dp0"
for /F %%a in ('echo prompt $E^| cmd') do set "ESC=%%a"
cls
echo.
echo %ESC%[96m         BBBBB  OOOOO  TTTTT  PPPPP   AAA   GGGGG
echo %ESC%[96m         BB  BB OO  OO   T    PP  PP A   A GG
echo %ESC%[94m         BBBBB  OO  OO   T    PPPPP  AAAAA GG GGG
echo %ESC%[94m         BB  BB OO  OO   T    PP     A   A GG  GG
echo %ESC%[96m         BBBBB  OOOOO    T    PP     A   A  GGGGG
echo.
echo %ESC%[97m                    FORMATADOR DE PLANILHAS
echo %ESC%[95m                 Feito com ^<3 por Izabel Galvao :3
echo %ESC%[95m                 Atualizacoes no meu GITHUB: https://github.com/i3b3l
echo.
echo %ESC%[94m   ^>_ BOTPAG - FORMATADOR
echo.
echo %ESC%[96m   +---------------------------------------------------------------+
echo %ESC%[96m   ^| %ESC%[97mPREPARANDO A EXECUCAO%ESC%[96m                                      ^|
echo %ESC%[96m   +---------------------------------------------------------------+
echo %ESC%[97m   O robo padroniza os dados e organiza a aba ENTRADA.
echo %ESC%[94m   [OK]%ESC%[97m Limpa abas de processamento, mantendo os cabecalhos.
echo.
echo %ESC%[95m   SAIDA: %ESC%[97mPagamentos_Tratado.xlsx
echo %ESC%[97m   O arquivo original sera preservado.
echo.
echo %ESC%[93m   ! IMPORTANTE
echo %ESC%[97m   Feche o arquivo de saida no Excel antes de executar.
echo %ESC%[96m   +---------------------------------------------------------------+
echo.
echo %ESC%[92m   [TECLADO] Pressione qualquer tecla para continuar...
echo %ESC%[0m
pause >nul
echo.
if not exist "%~dp0BotPag-Formatador.exe" goto Falha
echo %ESC%[92m   [INICIANDO] BotPag-Formatador.exe%ESC%[0m
echo.
"%~dp0BotPag-Formatador.exe"
set "CODIGO_SAIDA=%ERRORLEVEL%"
echo.
if "%CODIGO_SAIDA%"=="0" (echo %ESC%[92m   [ENCERRADO] Formatador finalizado.%ESC%[0m) else (echo %ESC%[93m   [ATENCAO] Codigo de saida: %CODIGO_SAIDA%%ESC%[0m)
echo.
echo   Pressione qualquer tecla para sair.
pause >nul
exit /b %CODIGO_SAIDA%
:Falha
echo %ESC%[91m   [ERRO] BotPag-Formatador.exe nao encontrado.%ESC%[0m
echo   Pressione qualquer tecla para sair.
pause >nul
exit /b 1
