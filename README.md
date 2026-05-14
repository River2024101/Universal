@echo off
REM ============================================================
REM SAP CREDITORS AGEING - AUTOMATED RUNNER
REM ============================================================

setlocal enabledelayedexpansion

echo.
echo ============================================================
echo  SAP CREDITORS AGEING REPORT - AUTOMATED PROCESSOR
echo ============================================================
echo.

REM Run Sap_Creditors.py
echo [1/2] Processing SAP OS data...
python Sap_Creditors.py
if errorlevel 1 (
    echo ERROR: Sap_Creditors.py failed!
    pause
    exit /b 1
)

REM Run Sap_Creditors_Ageing.py
echo.
echo [2/2] Generating ageing report...
python Sap_Creditors_Ageing.py
if errorlevel 1 (
    echo ERROR: Sap_Creditors_Ageing.py failed!
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  PROCESS COMPLETED SUCCESSFULLY
echo ============================================================
echo.
pause
