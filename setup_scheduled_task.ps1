# Run this script as Administrator (right-click PowerShell -> "Run as administrator")
# Creates a Windows Scheduled Task that runs the marketing dashboard pipeline
# every day at 8:00 AM, independent of whether Claude Code is open.

$TaskName   = "SuperFast Marketing Dashboard"
$PythonExe  = "C:\Users\austi\AppData\Local\Programs\Python\Python312\python.exe"
$ScriptPath = "C:\Users\austi\Claude\Projects\Marketing Dashboard\scripts\run_pipeline.py"
$WorkingDir = "C:\Users\austi\Claude\Projects\Marketing Dashboard"

$Action = New-ScheduledTaskAction `
    -Execute $PythonExe `
    -Argument "`"$ScriptPath`"" `
    -WorkingDirectory $WorkingDir

$Trigger = New-ScheduledTaskTrigger -Daily -At "8:00AM"

$Settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -WakeToRun $false

$Principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME `
    -LogonType InteractiveToken `
    -RunLevel Highest

# Remove existing task if present
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask `
    -TaskName   $TaskName `
    -Action     $Action `
    -Trigger    $Trigger `
    -Settings   $Settings `
    -Principal  $Principal `
    -Description "Downloads LP reports, builds dashboard, uploads to Drive, emails team"

Write-Host ""
Write-Host "SUCCESS: Task '$TaskName' registered." -ForegroundColor Green
Write-Host "Runs every day at 8:00 AM — no need for Claude Code to be open."
Write-Host ""
Write-Host "To test it right now, run:"
Write-Host "  Start-ScheduledTask -TaskName '$TaskName'"
