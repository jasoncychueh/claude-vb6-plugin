@echo off
setlocal

:: VB6 Resource Compiler Wrapper
:: Usage: compile_rc.bat <file.rc>
:: Compiles .rc to .RES using Windows SDK rc.exe (x86)

if "%~1"=="" (
    echo Usage: compile_rc.bat ^<file.rc^>
    exit /b 1
)

if not exist "%~1" (
    echo [ERROR] File not found: %~1
    exit /b 1
)

:: Find rc.exe (x86 version for VB6 compatibility)
set RC_EXE=
for /d %%D in ("C:\Program Files (x86)\Windows Kits\10\bin\10.*") do (
    if exist "%%D\x86\rc.exe" set RC_EXE="%%D\x86\rc.exe"
)

if not defined RC_EXE (
    echo [ERROR] rc.exe not found. Install Windows SDK.
    exit /b 1
)

set RC_FILE=%~f1
set RC_DIR=%~dp1

echo Compiling %RC_FILE% ...
echo Using %RC_EXE%

pushd "%RC_DIR%"
%RC_EXE% "%RC_FILE%"
set RESULT=%ERRORLEVEL%
popd

if %RESULT% EQU 0 (
    echo [OK] Resource compilation succeeded.
) else (
    echo [FAIL] Resource compilation failed.
)

exit /b %RESULT%
