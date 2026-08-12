@echo off
REM Script para ejecutar la aplicación Streamlit

echo.
echo =========================================
echo Sistema de Trazabilidad de Fertilizacion
echo =========================================
echo.

REM Verificar si el entorno virtual existe
if not exist "venv\Scripts\activate.bat" (
    echo [ERROR] Entorno virtual no encontrado.
    echo Por favor ejecuta 'install.bat' primero.
    pause
    exit /b 1
)

echo Activando entorno virtual...
call venv\Scripts\activate.bat

echo.
echo Iniciando aplicacion Streamlit...
echo La aplicacion se abrira en: http://localhost:8501
echo.
echo [Presiona CTRL+C para detener]
echo.

streamlit run app.py

pause
