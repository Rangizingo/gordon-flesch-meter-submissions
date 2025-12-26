# Debug Lexmark SNMP using the SnmpReader Send-SnmpGet function

cd $PSScriptRoot
. .\src\Logger.ps1

$ip = '10.3.1.40'
$community = 'public'
$oid = '1.3.6.1.2.1.43.10.2.1.4.1.1'

Write-Host "Testing Send-SnmpGet to $ip for OID $oid" -ForegroundColor Cyan

$udp = New-Object System.Net.Sockets.UdpClient
$udp.Client.ReceiveTimeout = 5000

# Build exactly like Send-SnmpGet does now
$oidParts = $oid.Split('.') | ForEach-Object { [int]$_ }
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

$commBytes = [System.Text.Encoding]::ASCII.GetBytes($community)

$oidSeq = @([byte]0x06, [byte]$oidBytes.Count) + $oidBytes.ToArray()
$nullVal = @([byte]0x05, [byte]0x00)
$varbind = @([byte]0x30, [byte]($oidSeq.Count + $nullVal.Count)) + $oidSeq + $nullVal
$varbindList = @([byte]0x30, [byte]$varbind.Count) + $varbind

$reqId = @([byte]0x02, [byte]0x04, [byte]0x00, [byte]0x00, [byte]0x00, [byte]0x01)
$errStatus = @([byte]0x02, [byte]0x01, [byte]0x00)
$errIndex = @([byte]0x02, [byte]0x01, [byte]0x00)
$pduContent = $reqId + $errStatus + $errIndex + $varbindList
$pdu = @([byte]0xA0, [byte]$pduContent.Count) + $pduContent

$version = @([byte]0x02, [byte]0x01, [byte]0x01)
$commSeq = @([byte]0x04, [byte]$commBytes.Count) + $commBytes
$msgContent = $version + $commSeq + $pdu
$packet = @([byte]0x30, [byte]$msgContent.Count) + $msgContent

Write-Host "Sending packet ($(${packet}.Length) bytes):" -ForegroundColor Yellow
Write-Host (($packet | ForEach-Object { $_.ToString('X2') }) -join ' ')

try {
    $udp.Connect($ip, 161)
    [void]$udp.Send([byte[]]$packet, $packet.Length)

    $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
    $response = $udp.Receive([ref]$ep)
    $udp.Close()

    Write-Host ""
    Write-Host "Response ($($response.Length) bytes):" -ForegroundColor Green
    Write-Host (($response | ForEach-Object { $_.ToString('X2') }) -join ' ')

    Write-Host ""
    Write-Host "Parsing for Counter32 (0x41)..." -ForegroundColor Yellow

    for ($i = $response.Length - 3; $i -ge 20; $i--) {
        $type = $response[$i]
        if ($type -eq 0x41 -or $type -eq 0x42 -or ($type -eq 0x02 -and $i -gt 30)) {
            $len = $response[$i + 1]
            if ($len -gt 0 -and $len -le 5 -and ($i + 2 + $len) -le $response.Length) {
                $val = [uint64]0
                for ($j = 0; $j -lt $len; $j++) {
                    $val = ($val -shl 8) + $response[$i + 2 + $j]
                }
                $typeName = switch ($type) { 0x41 { "Counter32" }; 0x42 { "Gauge32" }; 0x02 { "Integer" } }
                Write-Host "  Found $typeName at pos $i, len=$len, value=$val" -ForegroundColor $(if ($val -gt 100) { 'Green' } else { 'Gray' })
            }
        }
    }

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
