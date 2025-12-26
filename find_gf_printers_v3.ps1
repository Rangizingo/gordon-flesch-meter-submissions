# Find all printers from Gordon Flesch emails in O365 - FIXED VERSION
# Handles plain text email format

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Mail

Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Mail.Read", "Mail.ReadBasic" -NoWelcome

$user = (Get-MgContext).Account
Write-Host "Connected as: $user" -ForegroundColor Green

Write-Host "`nSearching for Gordon Flesch meter emails..." -ForegroundColor Cyan

# Search multiple ways
$allMessages = @{}

# Search by subject
$search1 = Get-MgUserMessage -UserId $user -Search '"meter reading request"' -Top 500 -Property Id,Subject,Body,ReceivedDateTime,From -ErrorAction SilentlyContinue
foreach ($m in $search1) { if ($m.Id) { $allMessages[$m.Id] = $m } }
Write-Host "  Subject search: $($search1.Count) emails"

# Search by from address
$search2 = Get-MgUserMessage -UserId $user -Search '"gflesch.com"' -Top 500 -Property Id,Subject,Body,ReceivedDateTime,From -ErrorAction SilentlyContinue
foreach ($m in $search2) { if ($m.Id) { $allMessages[$m.Id] = $m } }
Write-Host "  From search: $($search2.Count) emails"

$messages = $allMessages.Values
Write-Host "`nTotal unique emails: $($messages.Count)" -ForegroundColor Green

# Parse printer info - PLAIN TEXT FORMAT
$printers = @{}

foreach ($msg in $messages) {
    if (-not $msg.Body -or -not $msg.Body.Content) { continue }

    $body = $msg.Body.Content

    # Plain text patterns:
    # Equipment        MA7502
    # Serial Number    W433L400252
    # Make             Ricoh
    # Model            MP3352SP

    # Pattern for Equipment ID (2 letters + 4 digits)
    $equipPattern = 'Equipment\s+([A-Z]{2}\d{4})'
    $serialPattern = 'Serial Number\s+([A-Z0-9]+)'
    $makePattern = 'Make\s+(\w+)'
    $modelPattern = 'Model\s+([A-Z0-9\-]+)'
    $locationPattern = 'Location\s+(.+?)(?:\r?\n|$)'

    $equipMatches = [regex]::Matches($body, $equipPattern, 'IgnoreCase')
    $serialMatches = [regex]::Matches($body, $serialPattern, 'IgnoreCase')
    $makeMatches = [regex]::Matches($body, $makePattern, 'IgnoreCase')
    $modelMatches = [regex]::Matches($body, $modelPattern, 'IgnoreCase')
    $locationMatches = [regex]::Matches($body, $locationPattern, 'IgnoreCase')

    for ($i = 0; $i -lt $equipMatches.Count; $i++) {
        $equipId = $equipMatches[$i].Groups[1].Value.Trim()
        $serial = if ($serialMatches.Count -gt $i) { $serialMatches[$i].Groups[1].Value.Trim() } else { "" }
        $make = if ($makeMatches.Count -gt $i) { $makeMatches[$i].Groups[1].Value.Trim() } else { "Unknown" }
        $model = if ($modelMatches.Count -gt $i) { $modelMatches[$i].Groups[1].Value.Trim() } else { "Unknown" }
        $location = if ($locationMatches.Count -gt $i) { $locationMatches[$i].Groups[1].Value.Trim() } else { "Unknown" }

        # Use serial as unique key, fallback to equipId
        $key = if ($serial) { $serial } else { $equipId }

        if ($key -and -not $printers.ContainsKey($key)) {
            $printers[$key] = @{
                EquipmentID = $equipId
                SerialNumber = $serial
                Make = $make
                Model = $model
                Location = $location
                LastSeen = $msg.ReceivedDateTime
            }
        }
    }

    # Also try HTML table format (original .msg style)
    if ($equipMatches.Count -eq 0) {
        $htmlPattern = '<td[^>]*>([A-Z]{2}\d{4})<br[^>]*>\s*([A-Z0-9]+)</td>\s*<td[^>]*>([^<]+)<br[^>]*>\s*([^<]+)</td>'
        $htmlMatches = [regex]::Matches($body, $htmlPattern, 'IgnoreCase,Singleline')

        foreach ($match in $htmlMatches) {
            $equipId = $match.Groups[1].Value.Trim()
            $serial = $match.Groups[2].Value.Trim()
            $make = $match.Groups[3].Value.Trim()
            $model = $match.Groups[4].Value.Trim()

            $key = if ($serial) { $serial } else { $equipId }

            if ($key -and -not $printers.ContainsKey($key)) {
                $printers[$key] = @{
                    EquipmentID = $equipId
                    SerialNumber = $serial
                    Make = $make
                    Model = $model
                    Location = "Unknown"
                    LastSeen = $msg.ReceivedDateTime
                }
            }
        }
    }
}

# Display results
Write-Host "`n========== PRINTERS FOUND ==========" -ForegroundColor Yellow
Write-Host ""

$printerList = $printers.Values | Sort-Object EquipmentID

$printerList | ForEach-Object {
    Write-Host "$($_.EquipmentID) | $($_.SerialNumber) | $($_.Make) $($_.Model) | $($_.Location)" -ForegroundColor White
}

Write-Host "`n========== SUMMARY ==========" -ForegroundColor Yellow
Write-Host "Total unique printers: $($printers.Count)" -ForegroundColor Green

# Export
$csvPath = "$PSScriptRoot\gf_printers.csv"
$printerList | ForEach-Object { [PSCustomObject]$_ } | Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "Exported to: $csvPath" -ForegroundColor Cyan

Disconnect-MgGraph | Out-Null
Write-Host "`nDone." -ForegroundColor Green
