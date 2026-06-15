@echo off
setlocal

set PYTHON=C:\Users\austi\AppData\Local\Programs\Python\Python312\python.exe
set SCRIPTS=C:\Users\austi\Claude\Projects\Marketing Dashboard\scripts
set LOG=C:\Users\austi\Claude\Projects\Marketing Dashboard\update_log.txt

echo [%date% %time%] Starting dashboard update >> "%LOG%"

echo Step 1: Downloading reports from LeadPerfection...
"%PYTHON%" "%SCRIPTS%\download_reports.py" >> "%LOG%" 2>&1
if %ERRORLEVEL% neq 0 (
    echo ERROR: Download failed. Check update_log.txt
    exit /b 1
)

echo Step 2: Building dashboard HTML...
"%PYTHON%" "%SCRIPTS%\build_dashboard.py" >> "%LOG%" 2>&1
if %ERRORLEVEL% neq 0 (
    echo ERROR: Build failed. Check update_log.txt
    exit /b 1
)

echo Step 3: Uploading to Google Drive...
"%PYTHON%" "%SCRIPTS%\upload_to_drive.py" >> "%LOG%" 2>&1
if %ERRORLEVEL% neq 0 (
    echo ERROR: Upload failed. Check update_log.txt
    exit /b 1
)

echo [%date% %time%] Dashboard update complete! >> "%LOG%"
echo Done! Dashboard updated and uploaded to Google Drive.
