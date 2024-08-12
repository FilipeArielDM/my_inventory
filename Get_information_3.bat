@echo off
setlocal
for /f "tokens=*" %%i in ('hostname') do set "computer=%%i"

:: Criação do arquivo de log
set "logfile=%~dp0inventario/powertrain/%computer%.log"

:: Cabeçalho
(
    echo ============================================
    echo.
    echo Configuração de IP:
    echo.
    ipconfig
    echo.
    echo ============================================
    echo.
    echo Informações do Sistema:
    echo.
    systeminfo | findstr /B /C:"OS Name" /C:"Host Name" /C:"System Model" /C:"System Manufacturer"
    echo.
    echo ============================================
    echo.
    echo Softwares:
    echo.
    wmic product get name | findstr /R "TXOne CrowdStrike Symantec"
    
) > "%logfile%"

:: Fim do script
goto :eof

:: Função de pausa alternativa ao timeout, utilizando ping
:customTimeout
set /a "delay=%1+1"
ping 127.0.0.1 -n %delay% > nul
goto :eof