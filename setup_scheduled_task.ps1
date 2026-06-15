# Creates a Windows Scheduled Task to run the dashboard update
# every Tuesday at 8:00 AM Mountain Time (UTC-6 / UTC-7 DST).
# Run this once from an elevated PowerShell prompt.

$TaskName    = "MarketingDashboardUpdate"
$BatFile     = "C:\Users\austi\Claude\Projects\Marketing Dashboard\run_dashboard_update.bat"
$Description = "Downloads LP reports, rebuilds dashboard HTML, uploads to Google Drive"

# Tuesday = 3 in PowerShell's DayOfWeek (Sun=0, Mon=1, Tue=2, Wed=3 ... actually Tue=2)
# 8am Mountain = 8am local (Windows will use the machine's local timezone)
$Trigger = New-ScheduledTaskTrigger `
    -Weekly `
    -DaysOfWeek Tuesday `
    -At "08:00AM"

$Action = New-ScheduledTaskAction `
    -Execute "cmd.exe" `
    -Argument "/c `"$BatFile`""

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -RunOnlyIfNetworkAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)

$Principal = New-ScheduledTaskPrincipal `
    -UserId $env:USERNAME `
    -LogonType InteractiveToken `
    -RunLevel Highest

# Remove existing task if present
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue

Register-ScheduledTask `
    -TaskName    $TaskName `
    -Trigger     $Trigger `
    -Action      $Action `
    -Settings    $Settings `
    -Principal   $Principal `
    -Description $Description

Write-Host ""
Write-Host "Scheduled task '$TaskName' created." -ForegroundColor Green
Write-Host "Runs every Tuesday at 8:00 AM."
Write-Host ""
Write-Host "To test it now: Start-ScheduledTask -TaskName '$TaskName'"
Write-Host "To view status: Get-ScheduledTask -TaskName '$TaskName'"
