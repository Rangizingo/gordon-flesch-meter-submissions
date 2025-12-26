# Debug Lexmark SNMP response

$ip = '10.3.1.40'
$port = 161
$community = 'public'

$udp = New-Object System.Net.Sockets.UdpClient
$udp.Client.ReceiveTimeout = 5000

# Build SNMPv2c GET for prtMarkerLifeCount (1.3.6.1.2.1.43.10.2.1.4.1.1)
$packet = [byte[]]@(
    0x30, 0x2d,  # SEQUENCE, length 45
    0x02, 0x01, 0x01,  # version: v2c (1)
    0x04, 0x06, 0x70, 0x75, 0x62, 0x6c, 0x69, 0x63,  # community: "public"
    0xa0, 0x20,  # GET-REQUEST, length 32
    0x02, 0x04, 0x00, 0x00, 0x00, 0x01,  # request-id: 1
    0x02, 0x01, 0x00,  # error-status: 0
    0x02, 0x01, 0x00,  # error-index: 0
    0x30, 0x12,  # varbind list, length 18
    0x30, 0x10,  # varbind, length 16
    0x06, 0x0c, 0x2b, 0x06, 0x01, 0x02, 0x01, 0x2b, 0x0a, 0x02, 0x01, 0x04, 0x01, 0x01,  # OID: 1.3.6.1.2.1.43.10.2.1.4.1.1
    0x05, 0x00   # NULL value
)

try {
    $udp.Connect($ip, $port)
    [void]$udp.Send($packet, $packet.Length)

    $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
    $response = $udp.Receive([ref]$ep)
    $udp.Close()

    Write-Host "Response: $($response.Length) bytes" -ForegroundColor Cyan
    Write-Host ""

    # Full hex dump
    $hex = ($response | ForEach-Object { $_.ToString("X2") }) -join " "
    Write-Host "Full response:" -ForegroundColor Yellow
    Write-Host $hex
    Write-Host ""

    # Parse ASN.1 to find the value
    # Look for Counter32 (0x41) or Gauge32 (0x42) or Integer (0x02) near the end
    Write-Host "Scanning for values..." -ForegroundColor Yellow

    for ($i = 0; $i -lt $response.Length - 2; $i++) {
        $type = $response[$i]
        $len = $response[$i + 1]

        if (($type -eq 0x41 -or $type -eq 0x42 -or $type -eq 0x02) -and $len -gt 0 -and $len -le 5 -and ($i + 2 + $len) -le $response.Length) {
            $val = [uint64]0
            for ($j = 0; $j -lt $len; $j++) {
                $val = ($val -shl 8) + $response[$i + 2 + $j]
            }
            $typeName = switch ($type) { 0x41 { "Counter32" }; 0x42 { "Gauge32" }; 0x02 { "Integer" } }
            Write-Host "  Position $i : $typeName (len=$len) = $val" -ForegroundColor Green
        }
    }

} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
}
