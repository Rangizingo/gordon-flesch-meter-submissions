# Dry run submission walkthrough for both printers

cd $PSScriptRoot
. .\src\Logger.ps1
. .\src\SnmpReader.ps1
. .\src\GordonFlesch.ps1

Write-Host '========== FULL DRY RUN: BOTH PRINTERS ==========' -ForegroundColor Yellow
Write-Host ''

# Load config
$config = Get-Content .\config\printers.json | ConvertFrom-Json

foreach ($printer in $config.printers) {
    Write-Host "===== $($printer.equipmentId) ($($printer.make) $($printer.model)) =====" -ForegroundColor Cyan
    Write-Host "IP: $($printer.ip)"
    Write-Host "Serial: $($printer.serial)"
    Write-Host "Location: $($printer.location)"
    Write-Host ''

    # Step 1: SNMP Reading
    Write-Host '[STEP 1] Getting SNMP meter reading...' -ForegroundColor Yellow
    $reading = Get-PrinterMeterReading -IP $printer.ip -OID $printer.meterOid -Retries 1
    if ($reading.Success) {
        Write-Host "  Current Reading: $($reading.Reading)" -ForegroundColor Green
    } else {
        Write-Host "  FAILED: $($reading.Error)" -ForegroundColor Red
        continue
    }
    Write-Host ''

    # Step 2: Simulate GF submission URL (we'd get this from email)
    Write-Host '[STEP 2] Would parse submission URL from email...' -ForegroundColor Yellow
    Write-Host '  URL format: https://einfo.gflesch.com/einfo//aem.aspx?ac=<token>' -ForegroundColor Gray
    Write-Host ''

    # Step 3: Show what submission payload would look like
    Write-Host '[STEP 3] Submission payload would be:' -ForegroundColor Yellow
    $payload = @{
        allMeters = @(
            @{
                EquipmentID = '<internal-id-from-page>'
                ReadingDate = (Get-Date).ToString('MM/dd/yyyy')
                MeterReads = @(
                    @{
                        MeterID = '<meter-id-from-page>'
                        DisplayReading = $reading.Reading.ToString()
                        ActualReading = $reading.Reading.ToString()
                        IsRollover = 'false'
                        SeverityType = ''
                        Severity = ''
                    }
                )
            }
        )
    }
    Write-Host ($payload | ConvertTo-Json -Depth 5) -ForegroundColor DarkGray
    Write-Host ''

    # Step 4: Summary
    Write-Host '[STEP 4] Submission summary:' -ForegroundColor Yellow
    Write-Host "  Equipment: $($printer.equipmentId)" -ForegroundColor White
    Write-Host "  Reading: $($reading.Reading)" -ForegroundColor White
    Write-Host "  Date: $((Get-Date).ToString('MM/dd/yyyy'))" -ForegroundColor White
    Write-Host "  Status: READY TO SUBMIT" -ForegroundColor Green
    Write-Host ''
    Write-Host '-------------------------------------------'
    Write-Host ''
}

Write-Host '========== DRY RUN COMPLETE ==========' -ForegroundColor Green
Write-Host 'Both printers ready. Run .\src\Main.ps1 without -Test to submit for real.' -ForegroundColor Yellow
