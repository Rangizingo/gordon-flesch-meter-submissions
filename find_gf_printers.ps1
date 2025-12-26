# Find all printers from Gordon Flesch emails in O365
# Uses Microsoft Graph API with interactive login

# Install module if needed
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}

Import-Module Microsoft.Graph.Mail

# Connect to Graph with mail read permissions
Write-Host "Connecting to Microsoft Graph (will open browser for login)..." -ForegroundColor Cyan
Connect-MgGraph -Scopes "Mail.Read" -NoWelcome

# Get current user
$user = (Get-MgContext).Account
Write-Host "Connected as: $user" -ForegroundColor Green

# Search for Gordon Flesch meter emails
Write-Host "`nSearching for Gordon Flesch meter emails..." -ForegroundColor Cyan

$messages = Get-MgUserMessage -UserId $user -Filter "from/emailAddress/address eq 'gfc.contracts-d@gflesch.com'" -Top 100 -Property Subject,Body,ReceivedDateTime

Write-Host "Found $($messages.Count) emails from Gordon Flesch" -ForegroundColor Green

# Parse printer info from emails
$printers = @{}

foreach ($msg in $messages) {
    $body = $msg.Body.Content

    # Extract equipment/serial rows using regex
    # Pattern matches: Equipment ID, Serial, Make, Model from the HTML table
    $pattern = '<td[^>]*>([A-Z0-9]+)<br[^>]*>\s*([A-Z0-9]+)</td>\s*<td[^>]*>([^<]+)<br[^>]*>\s*([^<]+)</td>'

    $matches = [regex]::Matches($body, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

    foreach ($match in $matches) {
        $equipId = $match.Groups[1].Value.Trim()
        $serial = $match.Groups[2].Value.Trim()
        $make = $match.Groups[3].Value.Trim()
        $model = $match.Groups[4].Value.Trim()

        # Extract location from "Location Notes:" line
        $locPattern = "Location Notes:\s*</b>([^<]+)</td>"
        $locMatch = [regex]::Match($body, $locPattern)
        $location = if ($locMatch.Success) { $locMatch.Groups[1].Value.Trim() } else { "Unknown" }

        $key = "$equipId|$serial"
        if (-not $printers.ContainsKey($key)) {
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
}

# Display results
Write-Host "`n========== PRINTERS REQUIRING METER SUBMISSIONS ==========" -ForegroundColor Yellow
Write-Host ""

$printerList = $printers.Values | Sort-Object Make, Model

foreach ($p in $printerList) {
    Write-Host "Equipment ID: $($p.EquipmentID)" -ForegroundColor White
    Write-Host "  Serial:     $($p.SerialNumber)"
    Write-Host "  Make/Model: $($p.Make) $($p.Model)"
    Write-Host "  Location:   $($p.Location)"
    Write-Host "  Last Email: $($p.LastSeen)"
    Write-Host ""
}

Write-Host "========== SUMMARY ==========" -ForegroundColor Yellow
Write-Host "Total unique printers: $($printers.Count)" -ForegroundColor Green

# Export to CSV
$csvPath = "$PSScriptRoot\gf_printers.csv"
$printerList | ForEach-Object { [PSCustomObject]$_ } | Export-Csv -Path $csvPath -NoTypeInformation
Write-Host "Exported to: $csvPath" -ForegroundColor Cyan

# Disconnect
Disconnect-MgGraph | Out-Null
Write-Host "`nDone." -ForegroundColor Green
