# PrinterDiscovery.ps1 - Auto-detect printer info via SNMP

. "$PSScriptRoot\Logger.ps1"

function Send-SnmpGetRaw {
    param(
        [Parameter(Mandatory)]
        [string]$Target,
        [string]$Community = "public",
        [Parameter(Mandatory)]
        [string]$OidString,
        [int]$Timeout = 5000
    )

    $udp = New-Object System.Net.Sockets.UdpClient
    $udp.Client.ReceiveTimeout = $Timeout

    # Parse OID
    $oidParts = $OidString.Split('.') | ForEach-Object { [int]$_ }
    $oidBytes = [System.Collections.ArrayList]@()
    [void]$oidBytes.Add([byte](40 * $oidParts[0] + $oidParts[1]))

    for ($i = 2; $i -lt $oidParts.Count; $i++) {
        $val = $oidParts[$i]
        if ($val -lt 128) {
            [void]$oidBytes.Add([byte]$val)
        } else {
            $bytes = [System.Collections.ArrayList]@()
            while ($val -gt 0) {
                [void]$bytes.Insert(0, [byte]($val -band 0x7F))
                $val = $val -shr 7
            }
            for ($j = 0; $j -lt $bytes.Count - 1; $j++) {
                $bytes[$j] = [byte]($bytes[$j] -bor 0x80)
            }
            foreach ($b in $bytes) { [void]$oidBytes.Add($b) }
        }
    }

    $commBytes = [System.Text.Encoding]::ASCII.GetBytes($Community)

    # Build SNMP v1 GET packet
    $oidSeq = @([byte]0x06, [byte]$oidBytes.Count) + $oidBytes.ToArray()
    $nullVal = @([byte]0x05, [byte]0x00)
    $varbind = @([byte]0x30, [byte]($oidSeq.Count + $nullVal.Count)) + $oidSeq + $nullVal
    $varbindList = @([byte]0x30, [byte]$varbind.Count) + $varbind

    $reqId = @([byte]0x02, [byte]0x01, [byte]0x01)
    $errStatus = @([byte]0x02, [byte]0x01, [byte]0x00)
    $errIndex = @([byte]0x02, [byte]0x01, [byte]0x00)
    $pduContent = $reqId + $errStatus + $errIndex + $varbindList
    $pdu = @([byte]0xA0, [byte]$pduContent.Count) + $pduContent

    $version = @([byte]0x02, [byte]0x01, [byte]0x00)
    $commSeq = @([byte]0x04, [byte]$commBytes.Count) + $commBytes
    $msgContent = $version + $commSeq + $pdu
    $packet = @([byte]0x30, [byte]$msgContent.Count) + $msgContent

    try {
        $udp.Connect($Target, 161)
        [void]$udp.Send([byte[]]$packet, $packet.Count)

        $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
        $response = $udp.Receive([ref]$ep)
        $udp.Close()

        # Parse response - find value after OID
        for ($i = 20; $i -lt $response.Count - 2; $i++) {
            # Look for type byte after OID sequence
            if ($response[$i] -eq 0x04) {  # OCTET STRING
                $len = $response[$i + 1]
                if ($len -gt 0 -and ($i + 2 + $len) -le $response.Count) {
                    $str = [System.Text.Encoding]::ASCII.GetString($response, $i + 2, $len)
                    return @{ Success = $true; Value = $str.Trim(); Type = "String" }
                }
            }
            if ($response[$i] -eq 0x02 -and $i -gt 30) {  # INTEGER (skip early integers)
                $len = $response[$i + 1]
                if ($len -gt 0 -and $len -le 4) {
                    $val = 0
                    for ($j = 0; $j -lt $len; $j++) {
                        $val = ($val -shl 8) + $response[$i + 2 + $j]
                    }
                    return @{ Success = $true; Value = $val; Type = "Integer" }
                }
            }
            if ($response[$i] -eq 0x41) {  # Counter32
                $len = $response[$i + 1]
                $val = 0
                for ($j = 0; $j -lt $len; $j++) {
                    $val = ($val -shl 8) + $response[$i + 2 + $j]
                }
                return @{ Success = $true; Value = $val; Type = "Counter" }
            }
        }
        return @{ Success = $false; Error = "Parse failed" }
    } catch {
        return @{ Success = $false; Error = $_.Exception.Message }
    } finally {
        if ($udp) { $udp.Close() }
    }
}

function Get-PrinterInfo {
    param(
        [Parameter(Mandatory)]
        [string]$IP,
        [string]$Community = "public"
    )

    Write-Host "Discovering printer at $IP..." -ForegroundColor Cyan

    $info = @{
        IP = $IP
        Description = $null
        SerialNumber = $null
        PageCount = $null
        PageCountOID = $null
    }

    # Standard OIDs
    $descResult = Send-SnmpGetRaw -Target $IP -Community $Community -OidString "1.3.6.1.2.1.1.1.0"
    if ($descResult.Success) { $info.Description = $descResult.Value }

    $serialResult = Send-SnmpGetRaw -Target $IP -Community $Community -OidString "1.3.6.1.2.1.43.5.1.1.17.1"
    if ($serialResult.Success) { $info.SerialNumber = $serialResult.Value }

    # Try multiple page count OIDs
    $pageOids = @(
        @{ Name = "Standard"; OID = "1.3.6.1.2.1.43.10.2.1.4.1.1" }
        @{ Name = "Ricoh"; OID = "1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.1" }
        @{ Name = "HP"; OID = "1.3.6.1.4.1.11.2.3.9.4.2.1.4.1.2.5.0" }
        @{ Name = "Canon"; OID = "1.3.6.1.4.1.1602.1.11.1.3.1.4.101" }
        @{ Name = "Xerox"; OID = "1.3.6.1.4.1.253.8.53.13.2.1.6.1.20.1" }
        @{ Name = "Brother"; OID = "1.3.6.1.4.1.2435.2.3.9.4.2.1.5.5.10.0" }
        @{ Name = "Kyocera"; OID = "1.3.6.1.4.1.1347.43.10.1.1.12.1.1" }
    )

    foreach ($oid in $pageOids) {
        $result = Send-SnmpGetRaw -Target $IP -Community $Community -OidString $oid.OID
        if ($result.Success -and $result.Value -gt 100) {  # Reasonable page count
            $info.PageCount = $result.Value
            $info.PageCountOID = $oid.OID
            $info.PageCountType = $oid.Name
            break
        }
    }

    # Display results
    Write-Host ""
    Write-Host "========== PRINTER INFO ==========" -ForegroundColor Yellow
    Write-Host "IP:          $($info.IP)"
    Write-Host "Description: $($info.Description)"
    Write-Host "Serial:      $($info.SerialNumber)"
    Write-Host "Page Count:  $($info.PageCount) (via $($info.PageCountType) OID)"
    Write-Host "OID:         $($info.PageCountOID)"
    Write-Host ""

    return $info
}

function Add-PrinterToConfig {
    param(
        [Parameter(Mandatory)]
        [hashtable]$PrinterInfo,
        [string]$EquipmentId,
        [string]$Location = "Unknown"
    )

    $configPath = Join-Path $PSScriptRoot "..\config\printers.json"
    $config = Get-Content $configPath | ConvertFrom-Json

    # Parse make/model from description
    $make = "Unknown"
    $model = "Unknown"
    if ($PrinterInfo.Description) {
        if ($PrinterInfo.Description -match "(Ricoh|HP|Canon|Xerox|Brother|Kyocera|Lexmark|Sharp|Konica|Toshiba)") {
            $make = $matches[1]
        }
        $model = $PrinterInfo.Description
    }

    $newPrinter = @{
        equipmentId = $EquipmentId
        serial = $PrinterInfo.SerialNumber
        ip = $PrinterInfo.IP
        snmpCommunity = "public"
        meterOid = $PrinterInfo.PageCountOID
        location = $Location
        make = $make
        model = $model
        enabled = $true
    }

    $config.printers += $newPrinter
    $config | ConvertTo-Json -Depth 5 | Set-Content $configPath -Encoding UTF8

    Write-Host "Printer added to config!" -ForegroundColor Green
    return $newPrinter
}
