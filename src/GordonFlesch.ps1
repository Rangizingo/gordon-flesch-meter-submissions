# GordonFlesch.ps1 - Gordon Flesch meter submission API

. "$PSScriptRoot\Logger.ps1"

function Get-GFSessionData {
    param(
        [Parameter(Mandatory)]
        [string]$SubmissionUrl,

        [int]$Timeout = 30000
    )

    Write-Log "Fetching GF submission page..." -Level "INFO"

    try {
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $response = Invoke-WebRequest -Uri $SubmissionUrl -SessionVariable session -UseBasicParsing -TimeoutSec ($Timeout / 1000)

        if ($response.StatusCode -ne 200) {
            return @{ Success = $false; Error = "HTTP $($response.StatusCode)" }
        }

        $html = $response.Content

        # Extract equipment panel IDs (internal IDs)
        $equipPanelPattern = 'divEquipmentPanel_(\d+)'
        $equipPanelMatches = [regex]::Matches($html, $equipPanelPattern)

        # Extract meter reading IDs
        $meterIdPattern = 'newReading_(\d+)'
        $meterIdMatches = [regex]::Matches($html, $meterIdPattern)

        # Extract equipment number (display ID like MA7502)
        $equipNumPattern = '<span[^>]*class="[^"]*equipNum[^"]*"[^>]*>\s*\n?\s*([A-Z]{2}\d+)'
        $equipNumMatch = [regex]::Match($html, $equipNumPattern, 'IgnoreCase,Singleline')

        # Fallback pattern for equipment number
        if (-not $equipNumMatch.Success) {
            $equipNumPattern2 = '>([A-Z]{2}\d{4})\s*</span>'
            $equipNumMatch = [regex]::Match($html, $equipNumPattern2, 'IgnoreCase')
        }

        # Extract last reading value
        $lastReadingPattern = 'txtReadingValue[^>]*>([0-9,]+)</td>'
        $lastReadingMatch = [regex]::Match($html, $lastReadingPattern, 'IgnoreCase')

        $result = @{
            Success = $true
            Session = $session
            Cookies = $session.Cookies.GetCookies($SubmissionUrl)
            Equipment = @()
        }

        # Build equipment data
        for ($i = 0; $i -lt $equipPanelMatches.Count; $i++) {
            $equip = @{
                InternalId = $equipPanelMatches[$i].Groups[1].Value
                MeterId = if ($meterIdMatches.Count -gt $i) { $meterIdMatches[$i].Groups[1].Value } else { $null }
                EquipmentNumber = if ($equipNumMatch.Success) { $equipNumMatch.Groups[1].Value } else { $null }
                LastReading = if ($lastReadingMatch.Success) { $lastReadingMatch.Groups[1].Value -replace ',', '' } else { $null }
            }
            $result.Equipment += $equip
        }

        Write-Log "Found $($result.Equipment.Count) equipment(s) on submission page" -Level "INFO"

        return $result

    } catch {
        Write-Log "Failed to fetch GF page: $($_.Exception.Message)" -Level "ERROR"
        return @{ Success = $false; Error = $_.Exception.Message }
    }
}

function Submit-MeterReading {
    param(
        [Parameter(Mandatory)]
        [string]$SubmissionUrl,

        [Parameter(Mandatory)]
        [string]$InternalEquipmentId,

        [Parameter(Mandatory)]
        [string]$MeterId,

        [Parameter(Mandatory)]
        [int]$Reading,

        [Parameter(Mandatory)]
        $Session,

        [int]$Retries = 3,

        [int]$RetryDelay = 5000
    )

    $baseUrl = "https://einfo.gflesch.com"
    $validateUrl = "$baseUrl/einfo/Service/ValidateMeter"
    $saveUrl = "$baseUrl/einfo/Service/SaveAllMeters"

    $readingDate = (Get-Date).ToString("MM/dd/yyyy")

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        Write-Log "Submission attempt $attempt/$Retries..." -Level "INFO"

        try {
            # Step 1: Validate the meter reading
            $validatePayload = @{
                EquipmentID = $InternalEquipmentId
                MeterID = $MeterId
                ReadingDate = $readingDate
                DisplayReading = $Reading.ToString()
                IsMissing = $false
            } | ConvertTo-Json

            $validateResponse = Invoke-RestMethod -Uri $validateUrl `
                -Method POST `
                -Body $validatePayload `
                -ContentType "application/json;charset=utf-8" `
                -WebSession $Session

            Write-Log "Validation response received" -Level "INFO"

            # Step 2: Save all meters
            $savePayload = @{
                allMeters = @(
                    @{
                        EquipmentID = $InternalEquipmentId
                        ReadingDate = $readingDate
                        MeterReads = @(
                            @{
                                MeterID = $MeterId
                                DisplayReading = $Reading.ToString()
                                ActualReading = $Reading.ToString()
                                IsRollover = "false"
                                SeverityType = ""
                                Severity = ""
                            }
                        )
                    }
                )
            } | ConvertTo-Json -Depth 5

            $saveResponse = Invoke-RestMethod -Uri $saveUrl `
                -Method POST `
                -Body $savePayload `
                -ContentType "application/json;charset=utf-8" `
                -WebSession $Session

            Write-Log "Meter reading submitted successfully" -Level "SUCCESS"

            return @{
                Success = $true
                Reading = $Reading
                EquipmentId = $InternalEquipmentId
                MeterId = $MeterId
                Timestamp = Get-Date
                Attempts = $attempt
            }

        } catch {
            Write-Log "Submission attempt $attempt failed: $($_.Exception.Message)" -Level "WARN"

            if ($attempt -lt $Retries) {
                Start-Sleep -Milliseconds $RetryDelay
            }
        }
    }

    Write-Log "Submission failed after $Retries attempts" -Level "ERROR"

    return @{
        Success = $false
        Error = "Failed after $Retries attempts"
        EquipmentId = $InternalEquipmentId
        Timestamp = Get-Date
        Attempts = $Retries
    }
}

function Test-SubmissionUrl {
    param(
        [Parameter(Mandatory)]
        [string]$Url
    )

    try {
        $response = Invoke-WebRequest -Uri $Url -Method HEAD -UseBasicParsing -TimeoutSec 10
        return $response.StatusCode -eq 200
    } catch {
        return $false
    }
}

