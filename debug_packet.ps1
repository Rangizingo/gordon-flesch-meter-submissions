# Debug SNMP packet building

cd $PSScriptRoot
. .\src\SnmpReader.ps1

$ip = '10.3.1.40'
$community = 'public'
$oid = '1.3.6.1.2.1.43.10.2.1.4.1.1'

Write-Host "Building packet for OID: $oid" -ForegroundColor Cyan

# Parse OID same as Send-SnmpGet
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

Write-Host "OID bytes: $(($oidBytes | ForEach-Object { $_.ToString('X2') }) -join ' ')" -ForegroundColor Yellow
Write-Host "Expected:  2B 06 01 02 01 2B 0A 02 01 04 01 01" -ForegroundColor Gray

# The working debug script packet OID section was:
# 06 0C 2B 06 01 02 01 2B 0A 02 01 04 01 01
# Where 06=OID type, 0C=length 12, then the 12 OID bytes

$commBytes = [System.Text.Encoding]::ASCII.GetBytes($community)

# Build the packet
$oidSeq = @([byte]0x06, [byte]$oidBytes.Count) + $oidBytes.ToArray()
$nullVal = @([byte]0x05, [byte]0x00)
$varbind = @([byte]0x30, [byte]($oidSeq.Count + $nullVal.Count)) + $oidSeq + $nullVal
$varbindList = @([byte]0x30, [byte]$varbind.Count) + $varbind

$reqId = @([byte]0x02, [byte]0x01, [byte]0x01)
$errStatus = @([byte]0x02, [byte]0x01, [byte]0x00)
$errIndex = @([byte]0x02, [byte]0x01, [byte]0x00)
$pduContent = $reqId + $errStatus + $errIndex + $varbindList
$pdu = @([byte]0xA0, [byte]$pduContent.Count) + $pduContent

$version = @([byte]0x02, [byte]0x01, [byte]0x01)
$commSeq = @([byte]0x04, [byte]$commBytes.Count) + $commBytes
$msgContent = $version + $commSeq + $pdu
$packet = @([byte]0x30, [byte]$msgContent.Count) + $msgContent

Write-Host ""
Write-Host "Built packet ($(${packet}.Length) bytes):" -ForegroundColor Cyan
Write-Host (($packet | ForEach-Object { $_.ToString('X2') }) -join ' ')
Write-Host ""
Write-Host "Working packet (45 bytes):" -ForegroundColor Gray
Write-Host "30 2D 02 01 01 04 06 70 75 62 6C 69 63 A0 20 02 04 00 00 00 01 02 01 00 02 01 00 30 12 30 10 06 0C 2B 06 01 02 01 2B 0A 02 01 04 01 01 05 00"
