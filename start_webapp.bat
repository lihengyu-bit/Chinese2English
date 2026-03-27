@echo off
setlocal

cd /d "%~dp0"

set "VENV_DIR=%~dp0.venv"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    echo [1/4] Creating virtual environment...
    python -m venv "%VENV_DIR%"
    if errorlevel 1 goto :error
)

echo [2/4] Upgrading pip...
"%PYTHON_EXE%" -m pip install --upgrade pip
if errorlevel 1 goto :error

echo [3/4] Installing requirements...
"%PYTHON_EXE%" -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo [4/4] Starting web app...
echo.
echo Open in browser:
echo   http://127.0.0.1:8501
echo   http://SERVER-IP:8501
echo.

"%PYTHON_EXE%" -m streamlit run app.py --server.address 0.0.0.0 --server.port 8501
goto :eof

:error
echo.
echo Startup failed. Please check network access and Python installation.
pause
exit /b 1
