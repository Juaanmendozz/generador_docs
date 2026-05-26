@echo off
echo ================================================
echo   Instalando Generador de Documentos Word
echo ================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python no esta instalado.
    echo Por favor descarga Python desde https://python.org e instala con la opcion "Add to PATH"
    pause
    exit /b 1
)

echo Instalando dependencias...
pip install python-docx pillow --quiet

echo.
echo Iniciando la aplicacion...
python app.py
pause
