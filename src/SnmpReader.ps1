# SnmpReader.ps1 - SNMP meter reading retrieval for printers

. "$PSScriptRoot\Logger.ps1"

function Send-SnmpGet {
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

    # First two parts encoded as 40*first + second
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

    $reqId = @([byte]0x02, [byte]0x04, [byte]0x00, [byte]0x00, [byte]0x00, [byte]0x01)  # 4-byte request-id
    $errStatus = @([byte]0x02, [byte]0x01, [byte]0x00)
    $errIndex = @([byte]0x02, [byte]0x01, [byte]0x00)
    $pduContent = $reqId + $errStatus + $errIndex + $varbindList
    $pdu = @([byte]0xA0, [byte]$pduContent.Count) + $pduContent

    $version = @([byte]0x02, [byte]0x01, [byte]0x01)  # v2c (0x01) instead of v1 (0x00)
    $commSeq = @([byte]0x04, [byte]$commBytes.Count) + $commBytes
    $msgContent = $version + $commSeq + $pdu
    $packet = @([byte]0x30, [byte]$msgContent.Count) + $msgContent

    try {
        $udp.Connect($Target, 161)
        [void]$udp.Send([byte[]]$packet, $packet.Count)

        $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
        $response = $udp.Receive([ref]$ep)
        $udp.Close()

        # Parse response - find Counter32 (0x41), Gauge32 (0x42), or Integer (0x02)
        # Scan from end backwards to find the value
        for ($i = $response.Count - 3; $i -ge 20; $i--) {
            $type = $response[$i]
            # Counter32=0x41, Gauge32=0x42, Integer=0x02
            if ($type -eq 0x41 -or $type -eq 0x42 -or ($type -eq 0x02 -and $i -gt 30)) {
                $len = $response[$i + 1]
                if ($len -gt 0 -and $len -le 5 -and ($i + 2 + $len) -le $response.Count) {
                    $val = [uint64]0
                    for ($j = 0; $j -lt $len; $j++) {
                        $val = ($val -shl 8) + $response[$i + 2 + $j]
                    }
                    # Only accept reasonable page counts (> 100)
                    if ($val -gt 100) {
                        return @{ Success = $true; Value = $val; Error = $null }
                    }
                }
            }
        }
        return @{ Success = $false; Value = $null; Error = "Could not parse SNMP response" }
    } catch {
        return @{ Success = $false; Value = $null; Error = $_.Exception.Message }
    } finally {
        if ($udp) { $udp.Close() }
    }
}

function Get-PrinterMeterReading {
    param(
        [Parameter(Mandatory)]
        [string]$IP,

        [string]$Community = "public",

        [string]$OID = "1.3.6.1.4.1.367.3.2.1.2.19.5.1.9.1",

        [int]$Retries = 3,

        [int]$RetryDelay = 5000,

        [int]$Timeout = 10000
    )

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        Write-Log "SNMP query attempt $attempt/$Retries to $IP" -Level "INFO"

        $result = Send-SnmpGet -Target $IP -Community $Community -OidString $OID -Timeout $Timeout

        if ($result.Success) {
            Write-Log "SNMP success: $IP returned $($result.Value)" -Level "SUCCESS"
            return @{
                Success = $true
                Reading = $result.Value
                IP = $IP
                Timestamp = Get-Date
                Attempts = $attempt
            }
        }

        Write-Log "SNMP attempt $attempt failed: $($result.Error)" -Level "WARN"

        if ($attempt -lt $Retries) {
            Start-Sleep -Milliseconds $RetryDelay
        }
    }

    Write-Log "SNMP failed after $Retries attempts for $IP" -Level "ERROR"
    return @{
        Success = $false
        Reading = $null
        IP = $IP
        Error = "Failed after $Retries attempts"
        Timestamp = Get-Date
        Attempts = $Retries
    }
}

function Test-PrinterConnection {
    param(
        [Parameter(Mandatory)]
        [string]$IP,

        [int]$Timeout = 1000
    )

    $ping = Test-Connection -ComputerName $IP -Count 1 -TimeoutSeconds ($Timeout / 1000) -Quiet
    return $ping
}

