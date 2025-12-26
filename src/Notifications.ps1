# Notifications.ps1 - Windows Toast and Email notifications

. "$PSScriptRoot\Logger.ps1"

function Initialize-ToastNotifications {
    # Ensure BurntToast module is available
    if (-not (Get-Module -ListAvailable -Name BurntToast)) {
        Write-Log "Installing BurntToast module..." -Level "INFO"
        Install-Module BurntToast -Scope CurrentUser -Force
    }

    Import-Module BurntToast -ErrorAction Stop
}

function Send-ToastNotification {
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("Success", "Failure", "Info")]
        [string]$Type = "Info"
    )

    try {
        Initialize-ToastNotifications

        $icon = switch ($Type) {
            "Success" { "✅" }
            "Failure" { "❌" }
            "Info"    { "ℹ️" }
        }

        $toastTitle = "$icon $Title"

        New-BurntToastNotification -Text $toastTitle, $Message -AppLogo $null

        Write-Log "Toast notification sent: $Title" -Level "INFO"

    } catch {
        Write-Log "Failed to send toast notification: $($_.Exception.Message)" -Level "WARN"
    }
}

function Send-EmailNotification {
    param(
        [Parameter(Mandatory)]
        [string]$To,

        [Parameter(Mandatory)]
        [string]$Subject,

        [Parameter(Mandatory)]
        [string]$Body,

        [ValidateSet("Success", "Failure", "Info")]
        [string]$Type = "Info"
    )

    try {
        # Ensure connected to Graph
        $context = Get-MgContext -ErrorAction SilentlyContinue

        if (-not $context) {
            Write-Log "Not connected to Graph API, cannot send email" -Level "WARN"
            return $false
        }

        $icon = switch ($Type) {
            "Success" { "✅" }
            "Failure" { "❌" }
            "Info"    { "ℹ️" }
        }

        $htmlBody = @"
<html>
<body style="font-family: Arial, sans-serif;">
<h2>$icon $Subject</h2>
<pre style="background-color: #f4f4f4; padding: 15px; border-radius: 5px;">
$Body
</pre>
<hr>
<p style="color: #666; font-size: 12px;">
Automated message from Gordon Flesch Meter Submission Tool<br>
Generated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
</p>
</body>
</html>
"@

        $message = @{
            Subject = "$icon $Subject"
            Body = @{
                ContentType = "HTML"
                Content = $htmlBody
            }
            ToRecipients = @(
                @{
                    EmailAddress = @{
                        Address = $To
                    }
                }
            )
        }

        Send-MgUserMail -UserId $context.Account -Message $message

        Write-Log "Email notification sent to $To" -Level "INFO"
        return $true

    } catch {
        Write-Log "Failed to send email notification: $($_.Exception.Message)" -Level "WARN"
        return $false
    }
}

function Send-SubmissionSummary {
    param(
        [Parameter(Mandatory)]
        [array]$Results,

        [Parameter(Mandatory)]
        [hashtable]$Settings
    )

    $successCount = ($Results | Where-Object { $_.Success }).Count
    $failCount = ($Results | Where-Object { -not $_.Success }).Count
    $total = $Results.Count

    $status = if ($failCount -eq 0) { "Success" } elseif ($successCount -eq 0) { "Failure" } else { "Info" }

    $title = "Meter Submission: $successCount/$total Succeeded"

    $details = $Results | ForEach-Object {
        $icon = if ($_.Success) { "✓" } else { "✗" }
        "$icon $($_.EquipmentId): $(if ($_.Success) { $_.Reading } else { $_.Error })"
    }

    $message = $details -join "`n"

    # Toast notification
    if ($Settings.notifications.toast) {
        if ($status -eq "Success" -and $Settings.notifications.notifyOnSuccess) {
            Send-ToastNotification -Title $title -Message $message -Type $status
        } elseif ($status -ne "Success" -and $Settings.notifications.notifyOnFailure) {
            Send-ToastNotification -Title $title -Message $message -Type $status
        }
    }

    # Email notification
    if ($Settings.notifications.email -and $Settings.notifications.notifyEmail) {
        if ($status -eq "Success" -and $Settings.notifications.notifyOnSuccess) {
            Send-EmailNotification -To $Settings.notifications.notifyEmail -Subject $title -Body $message -Type $status
        } elseif ($status -ne "Success" -and $Settings.notifications.notifyOnFailure) {
            Send-EmailNotification -To $Settings.notifications.notifyEmail -Subject $title -Body $message -Type $status
        }
    }
}

Export-ModuleMember -Function Send-ToastNotification, Send-EmailNotification, Send-SubmissionSummary
