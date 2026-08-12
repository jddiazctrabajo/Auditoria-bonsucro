@echo off
REM Script de instalación para Windows
REM Crea un entorno virtual e instala dependencias

echo.
echo =========================================
echo Instalador - Sistema de Fertilizacion
echo =========================================
echo.

REM Verificar si Python está instalado
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no está instalado o no está en el PATH
    echo Por favor instala Python desde https://www.python.org
    pause
    exit /b 1
)

echo [1/4] Creando entorno virtual...
python -m venv venv
if errorlevel 1 (
    echo [ERROR] No se pudo crear el entorno virtual
    pause
    exit /b 1
)

echo [2/4] Activando entorno virtual...
call venv\Scripts\activate.bat

echo [3/4] Actualizando pip...
python -m pip install --upgrade pip

echo [4/4] Instalando dependencias...
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] No se pudieron instalar las dependencias
    pause
    exit /b 1
)

echo.
echo =========================================
echo [SUCCESS] Instalación completada!
echo =========================================
echo.
echo Para ejecutar la aplicación:
echo   - Abre una terminal en esta carpeta
echo   - Ejecuta: venv\Scripts\activate
echo   - Luego: streamlit run app.py
echo.
pause
