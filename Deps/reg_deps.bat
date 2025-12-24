@echo off
setlocal EnableDelayedExpansion

REM === Config ===
set VBP_FILE=%~1

if "%VBP_FILE%"=="" (
    echo Gebruik: check_vb6_dependencies.bat project.vbp
    exit /b 1
)

if not exist "%VBP_FILE%" (
    echo FOUT: VBP bestand niet gevonden: %VBP_FILE%
    exit /b 1
)

echo ======================================
echo VB6 Dependency Check
echo Project: %VBP_FILE%
echo ======================================
echo.

set ERROR_FOUND=0

REM === Lees alle OCX en DLL regels ===
for /f "usebackq delims=" %%L in ("%VBP_FILE%") do (
    echo %%L | findstr /i ".ocx .dll" >nul
	if not errorlevel 1 (
        call :PROCESS_LINE "%%L"
    )
)

echo.
echo ======================================
if %ERROR_FOUND%==0 (
    echo Alle dependencies succesvol verwerkt.
    exit /b 0
) else (
    echo Er zijn fouten opgetreden.
    exit /b 1
)

REM ================================
REM Subroutine
REM ================================
:PROCESS_LINE
set LINE=%~1

REM Haal alles na de laatste ; of |
for %%A in ("%LINE:;=" "%") do set FILE=%%~A
for %%A in ("%FILE:|=" "%") do set FILE=%%~A

REM Trim quotes
set FILE=%FILE:"=%
for /f "tokens=* delims= " %%F in ("%FILE%") do set FILE=%%F

echo Dependency: "%FILE%"

if exist "%FILE%" (
    echo Status: Registering...
    %systemroot%\SysWoW64\regsvr32.exe /s "%FILE%"
    if errorlevel 1 (
        echo [ERROR]
        set ERROR_FOUND=1
    ) else (
        echo [OK]
    )
) else (
    echo Status: Not Found
    set ERROR_FOUND=1
)