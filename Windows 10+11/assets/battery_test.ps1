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
    return $b.Percent
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
$startPercent = $battery.Percent
$startTime = Get-Date

Write-Host "Starting battery: $startPercent%"

$cpuInfo = Get-CimInstance Win32_Processor | Select-Object -First 1
$iterations = [math]::Max(10000, [math]::Round(35000 * ($cpuInfo.MaxClockSpeed / 2500)))

$workerCount = [math]::Max(2, [math]::Min(3, [Environment]::ProcessorCount))

$jobs = 1..$workerCount | ForEach-Object {

    Start-Job -ArgumentList $iterations -ScriptBlock {

        param($iterations)

        while ($true) {

            # --- dynamic workload scaling (realism) ---
            $localIterations = [math]::Round(
                $iterations * (Get-Random -Minimum 0.7 -Maximum 1.1)
            )

            # --- CPU burst phase (bursty, not constant) ---
            if ((Get-Random -Minimum 1 -Maximum 10) -le 6) {

                for ($i = 1; $i -le $localIterations; $i++) {

                    # lightweight mixed compute (closer to real app work than pure math loops)
                    $x = $i + 1
                    $val = [math]::Sqrt($x) * [math]::Log($x + 1)

                    # occasionally mix branch behavior (simulates app logic)
                    if (($i % 100) -eq 0) {
                        $val = $val * 0.99
                    }
                }
            }

            # --- light disk activity (occasional, not constant) ---
            if ((Get-Random -Minimum 1 -Maximum 10) -eq 1) {
                try {
                    $tmp = "$env:TEMP\battery_sim.tmp"
                    Set-Content -Path $tmp -Value (Get-Random)
                    Remove-Item $tmp -ErrorAction SilentlyContinue
                } catch {}
            }

            # --- rare network activity (very occasional like real apps) ---
            if ((Get-Random -Minimum 1 -Maximum 30) -eq 1) {
                try {
                    Invoke-WebRequest `
                        -Uri "https://www.wikipedia.org" `
                        -UseBasicParsing `
                        -TimeoutSec 3 | Out-Null
                } catch {}
            }

            # --- irregular idle time (important for realism) ---
            Start-Sleep -Milliseconds (Get-Random -Minimum 80 -Maximum 400)
        }
    }
}

try {
    $testExtended = $false

    do {
        Start-Sleep -Seconds ($durationMinutes * 60)

        $battery = Get-Battery
        $currentPercent = $battery.Percent
        $drop = $startPercent - $currentPercent

        if ($drop -lt 2 -and -not $testExtended) {
            Write-Host "Extending test..."
            $testExtended = $true
        } else {
            break
        }

    } while ($true)

}
finally {
    $jobs | Stop-Job
    Start-Sleep -Seconds 1
    $jobs | Remove-Job
}

$battery = Get-Battery
$endPercent = $battery.Percent
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