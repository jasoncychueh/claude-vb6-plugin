@echo off
setlocal

:: VB6 Command-Line Compiler Wrapper
:: Usage: compile.bat <project.vbp>
:: Bundled with vb6-dev skill for use across projects.

if "%~1"=="" (
    echo Usage: compile.bat ^<project.vbp^>
    exit /b 1
)

if not exist "%~1" (
    echo [ERROR] File not found: %~1
    exit /b 1
)

:: Find VB6.EXE (set VB6_PATH env var to override)
if defined VB6_PATH (
    set VB6="%VB6_PATH%"
) else (
    set VB6="C:\Program Files (x86)\Microsoft Visual Studio\VB6.EXE"
    if not exist %VB6% set VB6="C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE"
)
if not exist %VB6% (
    echo [ERROR] VB6.EXE not found. Set VB6_PATH environment variable.
    exit /b 1
)

set VBP=%~f1
set LOG=%~dp1compile.log

echo Compiling %VBP% ...
start "" /min /wait %VB6% /make "%VBP%" /out "%LOG%"
set RESULT=%ERRORLEVEL%

if exist "%LOG%" (
    echo.
    echo === Compile Log ===
    type "%LOG%"
    echo.
)

if %RESULT% EQU 0 (
    echo [OK] Compile succeeded.
) else (
    echo [FAIL] Compile failed. Check compile.log for details.
)

exit /b %RESULT%
