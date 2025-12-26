# Find all printers from Gordon Flesch emails in O365 - DEBUG VERSION
# Uses Microsoft Graph API with interactive login

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Mail

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Mail.Read", "Mail.ReadBasic" -NoWelcome

$user = (Get-MgContext).Account
Write-Host "Connected as: $user" -ForegroundColor Green

# Search broader - use subject contains instead of from filter
Write-Host "`nSearching for Gordon Flesch / meter reading emails..." -ForegroundColor Cyan

# Method 1: Search by subject
$searchResults = Get-MgUserMessage -UserId $user -Search '"meter reading request"' -Top 200 -Property Subject,Body,ReceivedDateTime,From

Write-Host "Found $($searchResults.Count) emails via subject search" -ForegroundColor Yellow

# Method 2: Also try filter by from address variations
$fromFilter1 = Get-MgUserMessage -UserId $user -Filter "contains(from/emailAddress/address, 'gflesch')" -Top 200 -Property Subject,Body,ReceivedDateTime,From -ErrorAction SilentlyContinue
$fromFilter2 = Get-MgUserMessage -UserId $user -Filter "contains(subject, 'meter')" -Top 200 -Property Subject,Body,ReceivedDateTime,From -ErrorAction SilentlyContinue

Write-Host "Found $($fromFilter1.Count) emails via from filter" -ForegroundColor Yellow
Write-Host "Found $($fromFilter2.Count) emails via subject filter" -ForegroundColor Yellow

# Combine all results
$allMessages = @{}
$combined = @($searchResults) + @($fromFilter1) + @($fromFilter2) | Where-Object { $_ -ne $null }

foreach ($msg in $combined) {
    if ($msg.Id -and -not $allMessages.ContainsKey($msg.Id)) {
        $allMessages[$msg.Id] = $msg
    }
}

$messages = $allMessages.Values
Write-Host "`nTotal unique emails to parse: $($messages.Count)" -ForegroundColor Green

# Show sample subjects
Write-Host "`nSample email subjects:" -ForegroundColor Cyan
$messages | Select-Object -First 10 | ForEach-Object {
    Write-Host "  - $($_.Subject) (from: $($_.From.EmailAddress.Address))"
}

# Parse printer info
$printers = @{}

foreach ($msg in $messages) {
    # Skip if no body
    if (-not $msg.Body -or -not $msg.Body.Content) { continue }

    $body = $msg.Body.Content

    # More flexible pattern - look for table rows with equipment data
    # Pattern 1: Equipment ID and Serial in same cell with <br>
    $pattern1 = '<td[^>]*class="Black"[^>]*>([A-Z]{2}\d+)<br[^/]*/?>\s*([A-Z0-9]+)</td>'

    # Pattern 2: Make and Model in same cell with <br>
    $pattern2 = '<td[^>]*class="Black"[^>]*>(Ricoh|HP|Canon|Lexmark|Xerox|Brother|Kyocera|Sharp|Konica|Toshiba)<br[^/]*/?>\s*([A-Z0-9]+)</td>'

    $equipMatches = [regex]::Matches($body, $pattern1, 'IgnoreCase,Singleline')
    $makeMatches = [regex]::Matches($body, $pattern2, 'IgnoreCase,Singleline')

    for ($i = 0; $i -lt $equipMatches.Count; $i++) {
        $equipId = $equipMatches[$i].Groups[1].Value.Trim()
        $serial = $equipMatches[$i].Groups[2].Value.Trim()

        $make = if ($makeMatches.Count -gt $i) { $makeMatches[$i].Groups[1].Value.Trim() } else { "Unknown" }
        $model = if ($makeMatches.Count -gt $i) { $makeMatches[$i].Groups[2].Value.Trim() } else { "Unknown" }

        # Location
        $locPattern = 'Location Notes:\s*</b>\s*([^<]+)</td>'
        $locMatches = [regex]::Matches($body, $locPattern, 'IgnoreCase')
        $location = if ($locMatches.Count -gt $i) { $locMatches[$i].Groups[1].Value.Trim() } else { "Unknown" }

        $key = $serial
        if ($serial -and -not $printers.ContainsKey($key)) {
            $printers[$key] = @{
                EquipmentID = $equipId
                SerialNumber = $serial
                Make = $make
                Model = $model
                Location = $location
                LastSeen = $msg.ReceivedDateTime
                Subject = $msg.Subject
            }
            Write-Host "  Found: $equipId / $serial / $make $model" -ForegroundColor DarkGray
        }
    }

    # Fallback: try to find ANY serial-looking pattern near equipment IDs
    if ($equipMatches.Count -eq 0) {
        # Look for pattern like "MA7502" followed somewhere by serial
        $fallbackPattern = '([A-Z]{2}\d{4})\s*<br[^>]*>\s*\n?\s*([A-Z]\d{3}[A-Z]\d{6})'
        $fallbackMatches = [regex]::Matches($body, $fallbackPattern, 'IgnoreCase,Singleline')

        foreach ($match in $fallbackMatches) {
            $equipId = $match.Groups[1].Value.Trim()
            $serial = $match.Groups[2].Value.Trim()

            $key = $serial
            if ($serial -and -not $printers.ContainsKey($key)) {
                $printers[$key] = @{
                    EquipmentID = $equipId
                    SerialNumber = $serial
                    Make = "Unknown"
                    Model = "Unknown"
                    Location = "Unknown"
                    LastSeen = $msg.ReceivedDateTime
                    Subject = $msg.Subject
                }
                Write-Host "  Found (fallback): $equipId / $serial" -ForegroundColor DarkGray
            }
        }
    }
}

# Display results
Write-Host "`n========== PRINTERS FOUND ==========" -ForegroundColor Yellow

$printerList = $printers.Values | Sort-Object EquipmentID

foreach ($p in $printerList) {
    Write-Host "$($p.EquipmentID) | $($p.SerialNumber) | $($p.Make) $($p.Model) | $($p.Location)" -ForegroundColor White
}

Write-Host "`n========== SUMMARY ==========" -ForegroundColor Yellow
Write-Host "Total unique printers: $($printers.Count)" -ForegroundColor Green

# Export to CSV
$csvPath = "$PSScriptRoot\gf_printers.csv"
$printerList | ForEach-Object { [PSCustomObject]$_ } | Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "Exported to: $csvPath" -ForegroundColor Cyan

# Debug: Save one email body for inspection
if ($messages.Count -gt 0) {
    $debugPath = "$PSScriptRoot\debug_email.html"
    $messages[0].Body.Content | Out-File -FilePath $debugPath -Encoding UTF8
    Write-Host "Saved sample email HTML to: $debugPath" -ForegroundColor DarkGray
}

Disconnect-MgGraph | Out-Null
Write-Host "`nDone." -ForegroundColor Green
