@echo off
setlocal enabledelayedexpansion

:: ==== CONFIG - Edit here ====
set HOST=localhost
set PORT=3306
set USER=root
set PASS=
set DBNAME=test
set SQLDIR=%~dp0split_output
set MYSQL=mysql.exe
:: ============================

echo ========================================
echo  MySQL Batch Import
echo ========================================
echo  Host   : %HOST%:%PORT%
echo  DB     : %DBNAME%
echo  Dir    : %SQLDIR%
echo ========================================
echo.

if not exist "%SQLDIR%" (
    echo [ERROR] Directory not found: %SQLDIR%
    pause
    exit /b 1
)

if not exist "%MYSQL%" (
    echo [ERROR] mysql.exe not found: %MYSQL%
    pause
    exit /b 1
)

set COUNT=0
for %%f in ("%SQLDIR%\part_*.sql") do set /a COUNT+=1

if %COUNT%==0 (
    echo [ERROR] No part_*.sql files found in: %SQLDIR%
    pause
    exit /b 1
)

echo Found %COUNT% files. Starting import...
echo.

set OK=0
set FAIL=0
set CUR=0

for %%f in ("%SQLDIR%\part_*.sql") do (
    set /a CUR+=1
    echo [!CUR!/%COUNT%] %%~nxf
    if "%PASS%"=="" (
        cmd /c ""%MYSQL%" -h%HOST% -P%PORT% -u%USER% --password= --default-character-set=utf8mb4 %DBNAME% < "%%f""
    ) else (
        cmd /c ""%MYSQL%" -h%HOST% -P%PORT% -u%USER% --password=%PASS% --default-character-set=utf8mb4 %DBNAME% < "%%f""
    )
    if !errorlevel!==0 (
        echo        [OK]
        set /a OK+=1
    ) else (
        echo        [FAILED]
        set /a FAIL+=1
        echo [!date! !time!] FAILED: %%f >> "%~dp0import_errors.log"
    )
)

echo.
echo ========================================
echo  Done: %OK% OK, %FAIL% FAILED
if %FAIL% gtr 0 echo  See: %~dp0import_errors.log
echo ========================================
pause
