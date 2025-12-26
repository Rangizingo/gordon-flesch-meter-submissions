# Install-ScheduledTask.ps1 - Set up Windows Scheduled Task for meter submissions

param(
    [switch]$Uninstall,
    [string]$Time = "08:00",
    [switch]$RunNow
)

$taskName = "GordonFlesch-MeterSubmission"
$scriptPath = Join-Path $PSScriptRoot "src\Main.ps1"

function Install-MeterSubmissionTask {
    param(
        [string]$TriggerTime
    )

    Write-Host "Installing scheduled task: $taskName" -ForegroundColor Cyan

    # Check if task exists
    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if ($existingTask) {
        Write-Host "Task already exists, removing old version..." -ForegroundColor Yellow
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    }

    # Parse time
    $timeParts = $TriggerTime.Split(':')
    $hour = [int]$timeParts[0]
    $minute = [int]$timeParts[1]

    # Create trigger (daily at specified time)
    $trigger = New-ScheduledTaskTrigger -Daily -At "${hour}:${minute}"

    # Create action
    $action = New-ScheduledTaskAction `
        -Execute "pwsh.exe" `
        -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" `
        -WorkingDirectory $PSScriptRoot

    # Create settings
    $settings = New-ScheduledTaskSettingsSet `
        -AllowStartIfOnBatteries `
        -DontStopIfGoingOnBatteries `
        -StartWhenAvailable `
        -RunOnlyIfNetworkAvailable `
        -WakeToRun

    # Create principal (run as current user)
    $principal = New-ScheduledTaskPrincipal `
        -UserId $env:USERNAME `
        -LogonType Interactive `
        -RunLevel Highest

    # Register task
    Register-ScheduledTask `
        -TaskName $taskName `
        -Trigger $trigger `
        -Action $action `
        -Settings $settings `
        -Principal $principal `
        -Description "Automatically submits printer meter readings to Gordon Flesch Company"

    Write-Host ""
    Write-Host "Task installed successfully!" -ForegroundColor Green
    Write-Host "  Name: $taskName" -ForegroundColor White
    Write-Host "  Runs: Daily at $TriggerTime" -ForegroundColor White
    Write-Host "  Script: $scriptPath" -ForegroundColor White
    Write-Host ""
    Write-Host "To test, run: .\src\Main.ps1 -Test" -ForegroundColor Yellow
}

function Uninstall-MeterSubmissionTask {
    Write-Host "Uninstalling scheduled task: $taskName" -ForegroundColor Cyan

    $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if ($existingTask) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "Task uninstalled successfully!" -ForegroundColor Green
    } else {
        Write-Host "Task not found." -ForegroundColor Yellow
    }
}

function Start-TaskNow {
    Write-Host "Starting task immediately..." -ForegroundColor Cyan

    $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue

    if ($task) {
        Start-ScheduledTask -TaskName $taskName
        Write-Host "Task started. Check logs for results." -ForegroundColor Green
    } else {
        Write-Host "Task not found. Install it first." -ForegroundColor Red
    }
}

# Main
if ($Uninstall) {
    Uninstall-MeterSubmissionTask
} elseif ($RunNow) {
    Start-TaskNow
} else {
    Install-MeterSubmissionTask -TriggerTime $Time
}
