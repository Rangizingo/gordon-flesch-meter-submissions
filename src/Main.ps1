# Main.ps1 - Gordon Flesch Meter Submission Automation
# Entry point that orchestrates all modules

param(
    [switch]$Test,      # Run in test mode (no actual submission)
    [switch]$Force,     # Process all emails regardless of processed status
    [switch]$Verbose
)

$ErrorActionPreference = "Stop"
$scriptRoot = $PSScriptRoot

# Import modules
. "$scriptRoot\Logger.ps1"
. "$scriptRoot\SnmpReader.ps1"
. "$scriptRoot\EmailMonitor.ps1"
. "$scriptRoot\GordonFlesch.ps1"
. "$scriptRoot\Notifications.ps1"

function Get-Configuration {
    $configPath = Join-Path $scriptRoot "..\config"

    $settings = Get-Content (Join-Path $configPath "settings.json") | ConvertFrom-Json
    $printers = Get-Content (Join-Path $configPath "printers.json") | ConvertFrom-Json

    return @{
        Settings = $settings
        Printers = $printers.printers | Where-Object { $_.enabled -eq $true }
    }
}

function Find-PrinterConfig {
    param(
        [string]$Serial,
        [string]$EquipmentId,
        [array]$PrinterConfigs
    )

    # Try to match by serial first
    $match = $PrinterConfigs | Where-Object { $_.serial -eq $Serial }

    if (-not $match) {
        # Fall back to equipment ID
        $match = $PrinterConfigs | Where-Object { $_.equipmentId -eq $EquipmentId }
    }

    return $match
}

function Invoke-MeterSubmission {
    param(
        [switch]$TestMode,
        [switch]$ForceProcess
    )

    Write-Log "========== Gordon Flesch Meter Submission Started ==========" -Level "INFO"
    Write-Log "Mode: $(if ($TestMode) { 'TEST (no submission)' } else { 'LIVE' })" -Level "INFO"

    # Load configuration
    try {
        $config = Get-Configuration
        Write-Log "Loaded $($config.Printers.Count) printer configuration(s)" -Level "INFO"
    } catch {
        Write-Log "Failed to load configuration: $($_.Exception.Message)" -Level "ERROR"
        return
    }

    # Connect to O365
    $connected = Connect-GraphAPI
    if (-not $connected) {
        Write-Log "Cannot proceed without Graph API connection" -Level "ERROR"
        return
    }

    # Clear old logs
    Clear-OldLogs -RetentionDays $config.Settings.logRetentionDays

    # Get pending meter requests
    $emails = Get-PendingMeterRequests `
        -SenderFilter $config.Settings.senderFilter `
        -SubjectFilter $config.Settings.subjectFilter

    if ($emails.Count -eq 0) {
        Write-Log "No pending meter request emails found" -Level "INFO"
        Disconnect-GraphAPI
        return
    }

    $results = @()

    foreach ($email in $emails) {
        # Check if already processed
        if (-not $ForceProcess -and (Test-EmailProcessed -EmailId $email.Id)) {
            Write-Log "Skipping already processed email: $($email.Subject)" -Level "INFO"
            continue
        }

        Write-Log "Processing email: $($email.Subject)" -Level "INFO"

        # Parse email
        $parsed = Parse-MeterRequestEmail -Email $email

        if (-not $parsed.SubmissionUrl) {
            Write-Log "No submission URL found in email" -Level "WARN"
            continue
        }

        # Process each printer in the email
        foreach ($printer in $parsed.Printers) {
            Write-Log "Processing printer: $($printer.EquipmentId) ($($printer.Serial))" -Level "INFO"

            # Find matching config
            $printerConfig = Find-PrinterConfig -Serial $printer.Serial -EquipmentId $printer.EquipmentId -PrinterConfigs $config.Printers

            if (-not $printerConfig) {
                Write-Log "No configuration found for printer $($printer.EquipmentId)" -Level "WARN"
                $results += @{
                    Success = $false
                    EquipmentId = $printer.EquipmentId
                    Error = "No configuration found"
                }
                continue
            }

            # Test printer connectivity
            if (-not (Test-PrinterConnection -IP $printerConfig.ip)) {
                Write-Log "Printer $($printerConfig.ip) is not reachable" -Level "ERROR"
                $results += @{
                    Success = $false
                    EquipmentId = $printer.EquipmentId
                    Error = "Printer not reachable"
                }
                continue
            }

            # Get SNMP reading
            $snmpResult = Get-PrinterMeterReading `
                -IP $printerConfig.ip `
                -Community $printerConfig.snmpCommunity `
                -OID $printerConfig.meterOid `
                -Retries $config.Settings.snmp.retries `
                -RetryDelay $config.Settings.snmp.retryDelay `
                -Timeout $config.Settings.snmp.timeout

            if (-not $snmpResult.Success) {
                Write-Log "Failed to get SNMP reading for $($printer.EquipmentId)" -Level "ERROR"
                $results += @{
                    Success = $false
                    EquipmentId = $printer.EquipmentId
                    Error = "SNMP failed: $($snmpResult.Error)"
                }
                continue
            }

            Write-Log "SNMP reading: $($snmpResult.Reading)" -Level "INFO"

            if ($TestMode) {
                Write-Log "[TEST MODE] Would submit reading $($snmpResult.Reading) for $($printer.EquipmentId)" -Level "INFO"
                $results += @{
                    Success = $true
                    EquipmentId = $printer.EquipmentId
                    Reading = $snmpResult.Reading
                    TestMode = $true
                }
                continue
            }

            # Get GF session data
            $gfSession = Get-GFSessionData -SubmissionUrl $parsed.SubmissionUrl

            if (-not $gfSession.Success) {
                Write-Log "Failed to get GF session: $($gfSession.Error)" -Level "ERROR"
                $results += @{
                    Success = $false
                    EquipmentId = $printer.EquipmentId
                    Error = "GF session failed"
                }
                continue
            }

            # Find matching equipment in GF data
            $gfEquipment = $gfSession.Equipment | Select-Object -First 1

            if (-not $gfEquipment) {
                Write-Log "No equipment found on GF submission page" -Level "ERROR"
                $results += @{
                    Success = $false
                    EquipmentId = $printer.EquipmentId
                    Error = "No equipment on GF page"
                }
                continue
            }

            # Submit meter reading
            $submitResult = Submit-MeterReading `
                -SubmissionUrl $parsed.SubmissionUrl `
                -InternalEquipmentId $gfEquipment.InternalId `
                -MeterId $gfEquipment.MeterId `
                -Reading $snmpResult.Reading `
                -Session $gfSession.Session `
                -Retries $config.Settings.submission.retries `
                -RetryDelay $config.Settings.submission.retryDelay

            $results += @{
                Success = $submitResult.Success
                EquipmentId = $printer.EquipmentId
                Reading = $snmpResult.Reading
                Error = $submitResult.Error
            }
        }

        # Mark email as processed
        if (-not $TestMode) {
            Add-ProcessedEmailId -EmailId $email.Id
        }
    }

    # Send summary notification
    if ($results.Count -gt 0) {
        Send-SubmissionSummary -Results $results -Settings $config.Settings
    }

    # Log summary
    $successCount = ($results | Where-Object { $_.Success }).Count
    $failCount = ($results | Where-Object { -not $_.Success }).Count

    Write-Log "========== Submission Complete ==========" -Level "INFO"
    Write-Log "Success: $successCount | Failed: $failCount | Total: $($results.Count)" -Level $(if ($failCount -eq 0) { "SUCCESS" } else { "WARN" })

    Disconnect-GraphAPI
}

# Run main function
try {
    Invoke-MeterSubmission -TestMode:$Test -ForceProcess:$Force
} catch {
    Write-Log "Fatal error: $($_.Exception.Message)" -Level "ERROR"
    Write-Log $_.ScriptStackTrace -Level "ERROR"

    # Try to send failure notification
    try {
        Send-ToastNotification -Title "Meter Submission Failed" -Message $_.Exception.Message -Type "Failure"
    } catch { }

    exit 1
}

exit 0
