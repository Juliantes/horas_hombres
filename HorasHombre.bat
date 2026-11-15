@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "ROOT=%~dp0"
set "APP_URL=http://127.0.0.1:8501"
set "FIREFOX="
set "PYTHON="
set "APP="

if exist "%ROOT%venv\Scripts\python.exe" (
    set "PYTHON=%ROOT%venv\Scripts\python.exe"
    set "APP=%ROOT%horas_hombre.py"
)
if not defined PYTHON if exist "%ROOT%python\python.exe" (
    set "PYTHON=%ROOT%python\python.exe"
    if exist "%ROOT%app\horas_hombre.py" (
        set "APP=%ROOT%app\horas_hombre.py"
    ) else (
        set "APP=%ROOT%horas_hombre.py"
    )
)

if not defined PYTHON (
    echo [ERROR] No se encontro Python en:
    echo   %ROOT%venv\Scripts\python.exe
    echo   ni %ROOT%python\python.exe
    echo.
    echo En una PC con Internet ejecute build.ps1 una vez.
    echo O use el zip completo generado en Mac (carpeta dist\HorasHombre).
    pause
    exit /b 1
)

if not exist "%ROOT%data" mkdir "%ROOT%data"
set "HH_DATA_DIR=%ROOT%data"
set "MPLCONFIGDIR=%ROOT%data\.matplotlib"

where firefox >nul 2>&1
if %ERRORLEVEL%==0 set "FIREFOX=firefox"
if not defined FIREFOX if exist "%ProgramFiles%\Mozilla Firefox\firefox.exe" (
    set "FIREFOX=%ProgramFiles%\Mozilla Firefox\firefox.exe"
)
if not defined FIREFOX if exist "%ProgramFiles(x86)%\Mozilla Firefox\firefox.exe" (
    set "FIREFOX=%ProgramFiles(x86)%\Mozilla Firefox\firefox.exe"
)
if not defined FIREFOX if exist "%LOCALAPPDATA%\Mozilla Firefox\firefox.exe" (
    set "FIREFOX=%LOCALAPPDATA%\Mozilla Firefox\firefox.exe"
)

echo ============================================
echo   Horas hombre vs OT
echo ============================================
echo.
echo Datos: %HH_DATA_DIR%
echo Iniciando servidor local...
echo.

if defined FIREFOX (
    start "" /b cmd /c "ping -n 5 127.0.0.1 >nul && start "" "%FIREFOX%" "%APP_URL%""
    echo Se abrira Firefox en unos segundos en %APP_URL%
) else (
    echo [AVISO] No se encontro Firefox instalado.
    echo Cuando vea "You can now view your Streamlit app" abajo,
    echo abra el navegador manualmente y vaya a:
    echo   %APP_URL%
    echo.
)

echo.
echo Para cerrar: Ctrl+C o cierre esta ventana.
echo.

"%PYTHON%" -m streamlit run "%APP%" ^
    --server.headless true ^
    --server.address 127.0.0.1 ^
    --server.port 8501 ^
    --browser.gatherUsageStats false ^
    --server.fileWatcherType none

pause