# Logger.ps1 - Logging utilities for Gordon Flesch Meter Submission

$script:LogDirectory = Join-Path $PSScriptRoot "..\logs"

function Get-LogPath {
    param(
        [string]$Date = (Get-Date -Format "yyyy-MM-dd")
    )

    if (-not (Test-Path $script:LogDirectory)) {
        New-Item -ItemType Directory -Path $script:LogDirectory -Force | Out-Null
    }

    return Join-Path $script:LogDirectory "submissions_$Date.log"
}

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",

        [switch]$NoConsole
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logLine = "[$timestamp] [$Level] $Message"

    $logPath = Get-LogPath
    Add-Content -Path $logPath -Value $logLine -Encoding UTF8

    if (-not $NoConsole) {
        $color = switch ($Level) {
            "INFO"    { "White" }
            "WARN"    { "Yellow" }
            "ERROR"   { "Red" }
            "SUCCESS" { "Green" }
        }
        Write-Host $logLine -ForegroundColor $color
    }
}

function Clear-OldLogs {
    param(
        [int]$RetentionDays = 30
    )

    $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
    $logFiles = Get-ChildItem -Path $script:LogDirectory -Filter "submissions_*.log" -ErrorAction SilentlyContinue

    foreach ($file in $logFiles) {
        if ($file.LastWriteTime -lt $cutoffDate) {
            Remove-Item $file.FullName -Force
            Write-Log "Removed old log file: $($file.Name)" -Level "INFO"
        }
    }
}

function Get-LogSummary {
    param(
        [string]$Date = (Get-Date -Format "yyyy-MM-dd")
    )

    $logPath = Get-LogPath -Date $Date

    if (-not (Test-Path $logPath)) {
        return @{
            Exists = $false
            Entries = 0
            Errors = 0
            Successes = 0
        }
    }

    $content = Get-Content $logPath

    return @{
        Exists = $true
        Entries = $content.Count
        Errors = ($content | Where-Object { $_ -match "\[ERROR\]" }).Count
        Successes = ($content | Where-Object { $_ -match "\[SUCCESS\]" }).Count
    }
}

Export-ModuleMember -Function Get-LogPath, Write-Log, Clear-OldLogs, Get-LogSummary
