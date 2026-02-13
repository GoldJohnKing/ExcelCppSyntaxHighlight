@echo off
setlocal EnableDelayedExpansion

:: Windows EXE Build Script for cpp_highlight.py
:: Run this directly in Windows or via WSL interop

set "PROJECT_DIR=%~dp0"
cd /d "%PROJECT_DIR%"

echo ============================================
echo  C++ Excel Syntax Highlighter - Windows Build
echo ============================================
echo.

:: Check for uv
where uv >nul 2>nul
if errorlevel 1 (
    echo [ERROR] uv not found. Please install uv first:
    echo   https://github.com/astral-sh/uv
    echo.
    echo Trying to use python/pip directly...
    goto USE_PIP
)

echo [Step 1/5] uv version:
uv --version
echo.

:: Create Windows venv
if not exist ".venv_windows" (
    echo [Step 2/5] Creating Windows virtual environment...
    uv venv .venv_windows
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment
        exit /b 1
    )
) else (
    echo [Step 2/5] Windows virtual environment exists
)
echo.

:: Install dependencies
echo [Step 3/5] Installing dependencies...
call .venv_windows\Scripts\activate.bat
uv pip install pyinstaller openpyxl pygments
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies
    exit /b 1
)
echo.

:: Build
echo [Step 4/5] Building EXE...
echo Using spec file: cpp_highlight.spec
uv run pyinstaller --clean cpp_highlight.spec

goto BUILD_DONE

:USE_PIP
echo [Step 1/3] Using pip directly...
echo.

:: Create venv with python
echo [Step 2/3] Creating virtual environment...
python -m venv .venv_windows
if errorlevel 1 (
    echo [ERROR] Failed to create virtual environment
    exit /b 1
)
echo.

:: Install dependencies
echo [Step 3/3] Installing dependencies and building...
call .venv_windows\Scripts\activate.bat
pip install pyinstaller openpyxl pygments
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies
    exit /b 1
)
echo.

echo Building EXE using spec file...
pyinstaller --clean cpp_highlight.spec

:BUILD_DONE
if errorlevel 1 (
    echo [ERROR] Build failed
    exit /b 1
)
echo.

echo [Step 5/5] Build completed!
echo.
echo Output: %PROJECT_DIR%dist\cpp_highlight.exe
echo.

:: Test
echo Testing: cpp_highlight.exe --help
"%PROJECT_DIR%dist\cpp_highlight.exe" --help
echo.

pause
