# Gordon Flesch Meter Submission Automation

## Project Scope & Development Checklist

---

## Overview

### Project Goals
Automate the monthly submission of printer meter readings to Gordon Flesch Company in response to their automated email requests.

### Requirements
- Monitor Office 365 email for Gordon Flesch meter reading requests
- Automatically retrieve current meter readings from printers via SNMP
- Submit readings to Gordon Flesch web portal via HTTP API
- Notify user of success/failure via Windows Toast and email
- Run daily as a Windows Scheduled Task
- Modular printer configuration (add/remove printers easily)

### Out of Scope
- Browser automation (direct HTTP confirmed working)
- Azure/cloud hosting (local workstation solution)
- Multi-user support (single user)

---

## Architecture

### Technology Stack
| Component | Technology |
|-----------|------------|
| Language | PowerShell 7+ |
| Email Access | Microsoft Graph API |
| Printer Metrics | SNMP (UDP port 161) |
| Form Submission | HTTP REST (JSON) |
| Scheduling | Windows Task Scheduler |
| Notifications | BurntToast (Windows Toast) + Send-MailMessage |
| Configuration | JSON file |
| Logging | Local text file |

### Data Flow
```
O365 Email → Parse Request → Lookup Printer Config → SNMP Query → HTTP Submit → Log & Notify
```

### File Structure
```
printerusage/
├── src/
│   ├── Main.ps1                    # Entry point
│   ├── EmailMonitor.ps1            # O365 email checking
│   ├── SnmpReader.ps1              # Printer SNMP queries
│   ├── GordonFlesch.ps1            # GF API submission
│   ├── Notifications.ps1           # Toast + email alerts
│   └── Logger.ps1                  # Logging utilities
├── config/
│   ├── printers.json               # Printer configuration
│   └── settings.json               # App settings (email, etc.)
├── logs/
│   └── submissions.log             # Activity log
├── Install-ScheduledTask.ps1       # Task scheduler setup
├── PROJECT_SCOPE.md
└── README.md
```

---

## Development Phases

### Phase 1: Foundation (Data & Config)
Core configuration and data structures.

- [ ] **1.1** Create project folder structure
- [ ] **1.2** Create `printers.json` schema with sample data
  - Fields: equipmentId, serial, ip, snmpCommunity, meterOid, location
- [ ] **1.3** Create `settings.json` schema
  - Fields: emailAddress, checkIntervalHours, logRetentionDays, notifyEmail
- [ ] **1.4** Create `Logger.ps1` with functions:
  - `Write-Log` (timestamped entries)
  - `Get-LogPath` (daily log rotation)
- [ ] **1.5** Create log file rotation (keep last 30 days)

**Acceptance Criteria**: Config files load without error, logs write to correct location.

---

### Phase 2: SNMP Reader
Retrieve meter readings from printers.

- [ ] **2.1** Create `SnmpReader.ps1` module
- [ ] **2.2** Implement `Get-PrinterMeterReading` function
  - Parameters: IP, Community, OID
  - Returns: Integer meter value or error
- [ ] **2.3** Add retry logic (3 attempts, 5s delay)
- [ ] **2.4** Add timeout handling (10s per attempt)
- [ ] **2.5** Add error handling for unreachable printers
- [ ] **2.6** Test against real printer (192.168.11.222)

**Acceptance Criteria**: Successfully retrieves meter reading from test printer, handles failures gracefully.

---

### Phase 3: Email Monitor
Check O365 for Gordon Flesch meter requests.

- [ ] **3.1** Create `EmailMonitor.ps1` module
- [ ] **3.2** Implement Microsoft Graph authentication
  - Use device code flow (interactive first-time)
  - Cache token for subsequent runs
- [ ] **3.3** Implement `Get-PendingMeterRequests` function
  - Filter: from = gfc.contracts-d@gflesch.com
  - Filter: subject contains "meter reading request"
  - Filter: unread or received within last 7 days
- [ ] **3.4** Implement `Parse-MeterRequestEmail` function
  - Extract: submission URL (ac= token)
  - Extract: equipment ID
  - Extract: serial number
  - Extract: due date
- [ ] **3.5** Implement duplicate detection (track processed emails)
- [ ] **3.6** Test with real O365 mailbox

**Acceptance Criteria**: Finds and parses Gordon Flesch emails correctly, no duplicate processing.

---

### Phase 4: Gordon Flesch API Submission
Submit meter readings to the GF web portal.

- [ ] **4.1** Create `GordonFlesch.ps1` module
- [ ] **4.2** Implement `Get-GFSessionData` function
  - GET submission URL with ac= token
  - Extract: session cookies
  - Extract: internal equipment ID
  - Extract: meter reading ID
- [ ] **4.3** Implement `Submit-MeterReading` function
  - POST to /einfo/Service/ValidateMeter
  - POST to /einfo/Service/SaveAllMeters
  - Return: success/failure status
- [ ] **4.4** Add retry logic (3 attempts)
- [ ] **4.5** Add response validation (confirm submission accepted)
- [ ] **4.6** Test with real submission URL (dry run first)

**Acceptance Criteria**: Successfully submits reading to GF portal, handles API errors.

---

### Phase 5: Notifications
Alert user to results.

- [ ] **5.1** Create `Notifications.ps1` module
- [ ] **5.2** Install BurntToast module (if not present)
- [ ] **5.3** Implement `Send-ToastNotification` function
  - Success: green checkmark, printer name, reading value
  - Failure: red X, error message
- [ ] **5.4** Implement `Send-EmailNotification` function
  - Use Microsoft Graph (same auth as email monitor)
  - Include: timestamp, printer, reading, status
- [ ] **5.5** Test both notification methods

**Acceptance Criteria**: Toast appears on screen, email received in inbox.

---

### Phase 6: Main Orchestration
Tie all components together.

- [ ] **6.1** Create `Main.ps1` entry point
- [ ] **6.2** Implement main workflow:
  1. Load config files
  2. Check for pending email requests
  3. For each request:
     - Match to printer config
     - Get SNMP reading
     - Submit to Gordon Flesch
     - Log result
  4. Send summary notification
- [ ] **6.3** Add global error handling (try/catch wrapper)
- [ ] **6.4** Add execution summary (X succeeded, Y failed)
- [ ] **6.5** Test full end-to-end flow

**Acceptance Criteria**: Complete workflow executes, all components integrate correctly.

---

### Phase 7: Scheduling & Deployment

- [ ] **7.1** Create `Install-ScheduledTask.ps1`
  - Task name: GordonFlesch-MeterSubmission
  - Trigger: Daily at 8:00 AM
  - Action: Run Main.ps1
  - Settings: Run whether user logged in or not
- [ ] **7.2** Add uninstall option
- [ ] **7.3** Test scheduled execution
- [ ] **7.4** Create README.md with:
  - Setup instructions
  - Configuration guide
  - Troubleshooting tips
- [ ] **7.5** Initial git commit and push

**Acceptance Criteria**: Task runs automatically, documentation complete.

---

### Phase 8: Hardening & Polish

- [ ] **8.1** Add input validation for all config values
- [ ] **8.2** Add credential encryption for stored tokens
- [ ] **8.3** Add health check mode (`-Test` parameter)
- [ ] **8.4** Add verbose mode (`-Verbose` parameter)
- [ ] **8.5** Add manual trigger mode (`-Force` parameter)
- [ ] **8.6** Test failure scenarios:
  - Printer offline
  - Invalid email credentials
  - GF portal unreachable
  - Malformed email
- [ ] **8.7** Final code review and cleanup

**Acceptance Criteria**: Handles all edge cases gracefully, no unhandled exceptions.

---

## Configuration Reference

### printers.json
```json
{
  "printers": [
    {
      "equipmentId": "MA7502",
      "serial": "W433L400252",
      "ip": "192.168.11.222",
      "snmpCommunity": "public",
      "meterOid": "1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.1",
      "location": "3rd FL Accounting"
    }
  ]
}
```

### settings.json
```json
{
  "emailAddress": "pblanco@equippers.com",
  "senderFilter": "gfc.contracts-d@gflesch.com",
  "checkIntervalDays": 1,
  "logRetentionDays": 30,
  "notifyOnSuccess": true,
  "notifyOnFailure": true,
  "notifyEmail": "pblanco@equippers.com"
}
```

---

## Testing Checklist

| Test | Expected Result |
|------|-----------------|
| SNMP query to valid printer | Returns integer reading |
| SNMP query to invalid IP | Returns error after 3 retries |
| Parse sample GF email | Extracts URL, equipment ID, serial |
| GET GF submission page | Returns 200, extracts IDs |
| POST meter reading | Returns success confirmation |
| Toast notification | Popup appears on screen |
| Email notification | Email received |
| Scheduled task trigger | Script executes at scheduled time |
| No pending emails | Script exits cleanly, logs "no requests" |

---

## Rollback Plan

If issues occur in production:
1. Disable scheduled task: `Disable-ScheduledTask -TaskName "GordonFlesch-MeterSubmission"`
2. Submit readings manually via GF portal
3. Review logs at `./logs/submissions.log`
4. Fix and re-enable

---

## Changelog

| Version | Date | Changes |
|---------|------|---------|
| 0.1 | TBD | Initial development |
| 1.0 | TBD | Production release |

---

*Last updated: 2024-12-26*
