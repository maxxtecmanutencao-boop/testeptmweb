import sys
from cx_Freeze import setup, Executable

# Dependências necessárias
build_exe_options = {
    "packages": [
        "streamlit",
        "pandas",
        "openpyxl",
        "plotly",
        "PIL",
        "selenium",
        "bs4",
        "lxml",
        "pathlib",
        "datetime",
        "io",
        "time",
        "shutil",
        "os"
    ],
    "includes": [
        "streamlit.web.cli",
        "streamlit.web.bootstrap"
    ],
    "include_files": [
        ("Petrobras.png", "Petrobras.png"),
        ("logo jsl.png", "logo jsl.png"),
        ("BD.xlsm", "BD.xlsm")
    ],
    "excludes": [],
}

base = None
if sys.platform == "win32":
    base = "Win32GUI"

setup(
    name="PTM_JSL_Sistema",
    version="2.0",
    description="Sistema PTM JSL - Consultas e Monitoramento",
    options={"build_exe": build_exe_options},
    executables=[Executable("app_melhorado.py", base=base, icon=None)]
)