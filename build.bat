@echo off
echo ========================================
echo  PTM JSL - EMPACOTAMENTO DO SISTEMA
echo ========================================
echo.

echo [1/5] Verificando Python...
python --version
if errorlevel 1 (
    echo ERRO: Python nao encontrado!
    pause
    exit /b 1
)

echo.
echo [2/5] Instalando dependencias...
pip install -r requirements.txt

echo.
echo [3/5] Criando executavel com PyInstaller...
pyinstaller --name="PTM_JSL_Sistema" ^
    --onefile ^
    --windowed ^
    --add-data "Petrobras.png;." ^
    --add-data "logo jsl.png;." ^
    --add-data "BD.xlsm;." ^
    --hidden-import=streamlit.web.cli ^
    --hidden-import=streamlit.web.bootstrap ^
    --hidden-import=selenium ^
    --hidden-import=bs4 ^
    --hidden-import=openpyxl ^
    --hidden-import=plotly ^
    --collect-all streamlit ^
    app_melhorado.py

echo.
echo [4/5] Copiando arquivos necessarios...
if not exist "dist\PTM_JSL_Sistema" mkdir "dist\PTM_JSL_Sistema"
copy "Petrobras.png" "dist\PTM_JSL_Sistema\"
copy "logo jsl.png" "dist\PTM_JSL_Sistema\"
copy "BD.xlsm" "dist\PTM_JSL_Sistema\"
copy "requirements.txt" "dist\PTM_JSL_Sistema\"

echo.
echo [5/5] Criando pasta de distribuicao...
if not exist "DISTRIBUICAO" mkdir "DISTRIBUICAO"
xcopy /E /I /Y "dist\PTM_JSL_Sistema" "DISTRIBUICAO\PTM_JSL_Sistema"

echo.
echo ========================================
echo  EMPACOTAMENTO CONCLUIDO COM SUCESSO!
echo ========================================
echo.
echo Pasta de distribuicao: DISTRIBUICAO\PTM_JSL_Sistema
echo.
pause