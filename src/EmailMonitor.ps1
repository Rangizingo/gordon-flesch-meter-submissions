# EmailMonitor.ps1 - Office 365 email monitoring via Microsoft Graph

. "$PSScriptRoot\Logger.ps1"

$script:TokenCachePath = Join-Path $PSScriptRoot "..\config\.graph_token.json"

function Initialize-GraphAuth {
    # Ensure Microsoft.Graph module is available
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Mail)) {
        Write-Log "Installing Microsoft.Graph module..." -Level "INFO"
        Install-Module Microsoft.Graph -Scope CurrentUser -Force
    }

    Import-Module Microsoft.Graph.Mail -ErrorAction Stop
}

function Connect-GraphAPI {
    param(
        [switch]$Force
    )

    Initialize-GraphAuth

    $context = Get-MgContext -ErrorAction SilentlyContinue

    if ($context -and -not $Force) {
        Write-Log "Already connected to Graph API as $($context.Account)" -Level "INFO"
        return $true
    }

    try {
        Write-Log "Connecting to Microsoft Graph (browser auth)..." -Level "INFO"
        Connect-MgGraph -Scopes "Mail.Read", "Mail.ReadBasic", "Mail.Send" -NoWelcome
        $context = Get-MgContext
        Write-Log "Connected as $($context.Account)" -Level "SUCCESS"
        return $true
    } catch {
        Write-Log "Failed to connect to Graph API: $($_.Exception.Message)" -Level "ERROR"
        return $false
    }
}

function Disconnect-GraphAPI {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

function Get-PendingMeterRequests {
    param(
        [string]$SenderFilter = "gfc.contracts-d@gflesch.com",
        [string]$SubjectFilter = "meter reading request",
        [int]$DaysBack = 7
    )

    $user = (Get-MgContext).Account

    if (-not $user) {
        Write-Log "Not connected to Graph API" -Level "ERROR"
        return @()
    }

    Write-Log "Checking emails for $user..." -Level "INFO"

    try {
        # Search for meter reading request emails
        $messages = Get-MgUserMessage -UserId $user `
            -Filter "contains(subject, '$SubjectFilter')" `
            -Top 50 `
            -Property Id, Subject, Body, ReceivedDateTime, From, IsRead `
            -OrderBy "receivedDateTime desc"

        # Filter by sender and date
        $cutoffDate = (Get-Date).AddDays(-$DaysBack)

        $filtered = $messages | Where-Object {
            $_.From.EmailAddress.Address -eq $SenderFilter -and
            $_.ReceivedDateTime -gt $cutoffDate
        }

        Write-Log "Found $($filtered.Count) meter request emails" -Level "INFO"
        return $filtered

    } catch {
        Write-Log "Error fetching emails: $($_.Exception.Message)" -Level "ERROR"
        return @()
    }
}

function Parse-MeterRequestEmail {
    param(
        [Parameter(Mandatory)]
        $Email
    )

    $body = $Email.Body.Content

    $result = @{
        EmailId = $Email.Id
        Subject = $Email.Subject
        ReceivedDate = $Email.ReceivedDateTime
        IsRead = $Email.IsRead
        SubmissionUrl = $null
        Printers = @()
    }

    # Extract submission URL (ac= token)
    $urlPattern = 'href="(https://einfo\.gflesch\.com/einfo//aem\.aspx\?ac=[^"]+)"'
    $urlMatch = [regex]::Match($body, $urlPattern, 'IgnoreCase')

    if ($urlMatch.Success) {
        $result.SubmissionUrl = $urlMatch.Groups[1].Value
    }

    # Extract printer info from table
    # Pattern: Equipment ID and Serial in same cell
    $equipPattern = '<td[^>]*>([A-Z]{2}\d+)<br[^/]*/?>\s*\n?\s*([A-Z0-9]+)</td>'
    $equipMatches = [regex]::Matches($body, $equipPattern, 'IgnoreCase,Singleline')

    # Pattern: Make and Model
    $makePattern = '<td[^>]*>(Ricoh|HP|Canon|Lexmark|Xerox|Brother|Kyocera|Sharp|Konica|Toshiba)<br[^/]*/?>\s*\n?\s*([A-Z0-9]+)</td>'
    $makeMatches = [regex]::Matches($body, $makePattern, 'IgnoreCase,Singleline')

    # Location notes
    $locPattern = 'Location Notes:\s*</b>\s*([^<]+)</td>'
    $locMatches = [regex]::Matches($body, $locPattern, 'IgnoreCase')

    for ($i = 0; $i -lt $equipMatches.Count; $i++) {
        $printer = @{
            EquipmentId = $equipMatches[$i].Groups[1].Value.Trim()
            Serial = $equipMatches[$i].Groups[2].Value.Trim()
            Make = if ($makeMatches.Count -gt $i) { $makeMatches[$i].Groups[1].Value.Trim() } else { "Unknown" }
            Model = if ($makeMatches.Count -gt $i) { $makeMatches[$i].Groups[2].Value.Trim() } else { "Unknown" }
            Location = if ($locMatches.Count -gt $i) { $locMatches[$i].Groups[1].Value.Trim() } else { "Unknown" }
        }
        $result.Printers += $printer
    }

    Write-Log "Parsed email: $($result.Printers.Count) printer(s), URL: $(if ($result.SubmissionUrl) { 'Found' } else { 'Missing' })" -Level "INFO"

    return $result
}

function Get-ProcessedEmailIds {
    $processedPath = Join-Path $PSScriptRoot "..\config\.processed_emails.json"

    if (Test-Path $processedPath) {
        return (Get-Content $processedPath | ConvertFrom-Json)
    }
    return @()
}

function Add-ProcessedEmailId {
    param(
        [Parameter(Mandatory)]
        [string]$EmailId
    )

    $processedPath = Join-Path $PSScriptRoot "..\config\.processed_emails.json"
    $processed = Get-ProcessedEmailIds

    if ($EmailId -notin $processed) {
        $processed += $EmailId
        $processed | ConvertTo-Json | Set-Content $processedPath -Encoding UTF8
    }
}

function Test-EmailProcessed {
    param(
        [Parameter(Mandatory)]
        [string]$EmailId
    )

    $processed = Get-ProcessedEmailIds
    return $EmailId -in $processed
}

Export-ModuleMember -Function Connect-GraphAPI, Disconnect-GraphAPI, Get-PendingMeterRequests, Parse-MeterRequestEmail, Test-EmailProcessed, Add-ProcessedEmailId
