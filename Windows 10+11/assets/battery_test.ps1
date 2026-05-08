param(
    [int]$durationMinutes = 15,
    [scriptblock]$GetBatteryInfo
)

function Get-Battery {
    return & $GetBatteryInfo
}

function Get-BatteryPercent {
    $b = Get-Battery
    if (-not $b) { return $null }

    # already pre-aggregated from main script
    return $b.EstimatedChargeRemaining
}

Write-Host "Preparing laptop for battery test..."

# brightness
try {
    $brightnessObj = Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness
    $originalBrightness = $brightnessObj.CurrentBrightness
} catch {
    $originalBrightness = $null
}

powercfg -setactive SCHEME_BALANCED

try {
    (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods)
        .WmiSetBrightness(1,50)
} catch {}

$battery = Get-Battery
$startPercent = $battery.EstimatedChargeRemaining
$startTime = Get-Date

Write-Host "Starting battery: $startPercent%"

$cpuInfo = Get-CimInstance Win32_Processor | Select-Object -First 1
$iterations = [math]::Max(10000, [math]::Round(35000 * ($cpuInfo.MaxClockSpeed / 2500)))

$loadJob = Start-Job {
    param($iterations)

    while ($true) {
        # CPU Load
        1..$iterations | ForEach-Object { [math]::Sqrt($_) * [math]::Pow($_,0.7) } | Out-Null
        # Small disk activity 
        $tmp = "$env:TEMP\battery_test.tmp" 
        Get-Random -Minimum 1000 -Maximum 5000 | Out-File $tmp 
        Remove-Item $tmp -ErrorAction SilentlyContinue 
        # Occasional network activity 
        try { 
            Invoke-WebRequest -Uri "https://www.wikipedia.org" -UseBasicParsing -TimeoutSec 3 | Out-Null 
        } catch {}
        Start-Sleep -Milliseconds 500
    }
} -ArgumentList $iterations

try {
    $testExtended = $false

    do {
        Start-Sleep -Seconds ($durationMinutes * 60)

        $battery = Get-Battery
        $current = $battery.EstimatedChargeRemaining
        $drop = $startPercent - $current

        if ($drop -lt 2 -and -not $testExtended) {
            Write-Host "Extending test..."
            $testExtended = $true
        } else {
            break
        }

    } while ($true)

}
finally {
    Stop-Job $loadJob -Force -ErrorAction SilentlyContinue
    Remove-Job $loadJob -Force -ErrorAction SilentlyContinue
}

$battery = Get-Battery
$endPercent = $battery.EstimatedChargeRemaining
$endTime = Get-Date

$drop = $startPercent - $endPercent
$elapsed = ($endTime - $startTime).TotalMinutes

if ($drop -gt 0) {
    $rate = $drop / $elapsed
    $runtime = [int](100 / $rate)
} else {
    $runtime = 0
}

$output = [int]$runtime

if ($originalBrightness) {
    try {
        (Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods)
            .WmiSetBrightness(1,$originalBrightness)
    } catch {}
}

Write-Host $output
return $output