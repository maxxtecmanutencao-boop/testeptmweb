@echo off
title PTM JSL - Sistema de Instalacao
color 0A

echo ========================================
echo    PTM JSL - SISTEMA DE INSTALACAO
echo ========================================
echo.

echo Verificando componentes necessarios...
echo.

REM Verifica se Python esta instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo [!] Python NAO encontrado!
    echo.
    echo Deseja baixar e instalar o Python agora?
    echo [S] Sim - Abrir pagina de download
    echo [N] Nao - Cancelar instalacao
    echo.
    choice /C SN /M "Sua escolha"
    
    if errorlevel 2 goto :fim
    if errorlevel 1 (
        start https://www.python.org/downloads/
        echo.
        echo Apos instalar o Python, execute este instalador novamente.
        pause
        exit /b 1
    )
)

echo [OK] Python encontrado!
echo.

REM Verifica se o Chrome esta instalado
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe" >nul 2>&1
if errorlevel 1 (
    echo [!] Google Chrome NAO encontrado!
    echo.
    echo O sistema precisa do Google Chrome para funcionar.
    echo Deseja baixar e instalar o Chrome agora?
    echo [S] Sim - Abrir pagina de download
    echo [N] Nao - Continuar sem Chrome
    echo.
    choice /C SN /M "Sua escolha"
    
    if errorlevel 1 (
        start https://www.google.com/chrome/
        echo.
        echo Apos instalar o Chrome, execute este instalador novamente.
        pause
    )
) else (
    echo [OK] Google Chrome encontrado!
    echo.
)

echo Instalando dependencias Python...
echo Por favor aguarde, isso pode levar alguns minutos...
echo.

pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo [!] Erro ao instalar dependencias!
    pause
    exit /b 1
)

echo.
echo ========================================
echo   INSTALACAO CONCLUIDA COM SUCESSO!
echo ========================================
echo.
echo Para iniciar o sistema, execute:
echo    INICIAR_SISTEMA.bat
echo.
pause