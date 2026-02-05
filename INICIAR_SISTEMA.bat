@echo off
title PTM JSL - Sistema de Consultas 2.0
color 0B

echo ========================================
echo    PTM JSL - SISTEMA DE CONSULTAS 2.0
echo ========================================
echo.
echo Iniciando o sistema...
echo Por favor aguarde...
echo.

REM Inicia o Streamlit
streamlit run app_melhorado_atualizado_2.py --server.port 8501 --server.headless true

if errorlevel 1 (
    echo.
    echo [!] Erro ao iniciar o sistema!
    echo.
    echo Possivel solucao:
    echo 1. Execute o INSTALADOR.bat novamente
    echo 2. Verifique se o Python esta instalado
    echo 3. Verifique se todas as dependencias estao instaladas
    echo.
    pause
)