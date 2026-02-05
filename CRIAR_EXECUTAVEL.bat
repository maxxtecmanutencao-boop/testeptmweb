@echo off
echo ========================================
echo  CRIANDO EXECUTAVEL STANDALONE
echo ========================================
echo.

echo Instalando PyInstaller...
pip install pyinstaller

echo.
echo Criando executavel...
pyinstaller --name="PTM_JSL_Sistema_v2.0" ^
    --onefile ^
    --noconsole ^
    --icon=Petrobras.ico ^
    --add-data "Petrobras.png;." ^
    --add-data "logo jsl.png;." ^
    --add-data "BD.xlsm;." ^
    --hidden-import=streamlit.web.cli ^
    --hidden-import=streamlit.runtime.scriptrunner.magic_funcs ^
    --hidden-import=streamlit.web.bootstrap ^
    --hidden-import=selenium.webdriver.chrome.service ^
    --hidden-import=selenium.webdriver.common.service ^
    --collect-all streamlit ^
    --collect-all selenium ^
    --collect-all plotly ^
    app_melhorado.py

echo.
echo Copiando ChromeDriver...
if exist "chromedriver.exe" (
    copy "chromedriver.exe" "dist\"
)

echo.
echo ========================================
echo  EXECUTAVEL CRIADO COM SUCESSO!
echo ========================================
echo.
echo Arquivo: dist\PTM_JSL_Sistema_v2.0.exe
echo.
pause