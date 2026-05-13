param(
    [int]$durationMinutes = 15,
    [string]$batteryCorePath
)
. $batteryCorePath
function Get-BatteryPercent {

    $battery = Get-OverallBatteryStatus

    if (-not $battery) {
        return $null
    }

    return $battery.Percent
}

Write-Host "Preparing battery test..."

try {

    $brightnessObj = Get-WmiObject `
        -Namespace root/WMI `
        -Class WmiMonitorBrightness

    $originalBrightness = $brightnessObj.CurrentBrightness
}
catch {
    $originalBrightness = $null
}

powercfg -setactive SCHEME_BALANCED

try {

    (
        Get-WmiObject `
            -Namespace root/WMI `
            -Class WmiMonitorBrightnessMethods
    ).WmiSetBrightness(1, 50)

}
catch {}

$startPercent = Get-BatteryPercent
$startTime = Get-Date

Write-Host "Starting battery: $startPercent%"

$cpuInfo = Get-CimInstance Win32_Processor |
    Select-Object -First 1

$iterations = [math]::Max(
    10000,
    [math]::Round(
        35000 * ($cpuInfo.MaxClockSpeed / 2500)
    )
)

$workerCount = [math]::Max(
    2,
    [math]::Min(
        3,
        [Environment]::ProcessorCount
    )
)

$workers = 1..$workerCount | ForEach-Object {

    Start-Job -ArgumentList $iterations -ScriptBlock {

        param($iterations)

        while ($true) {

            $localIterations = [math]::Round(
                $iterations * (
                    Get-Random -Minimum 0.7 -Maximum 1.1
                )
            )

            if ((Get-Random -Minimum 1 -Maximum 10) -le 6) {

                for ($i = 1; $i -le $localIterations; $i++) {

                    $x = $i + 1

                    $null = (
                        [math]::Sqrt($x) *
                        [math]::Log($x + 1)
                    )
                }
            }

            if ((Get-Random -Minimum 1 -Maximum 10) -eq 1) {

                try {

                    $tmp = "$env:TEMP\battery_sim.tmp"

                    Set-Content $tmp (Get-Random)

                    Remove-Item $tmp -ErrorAction SilentlyContinue

                }
                catch {}
            }

            Start-Sleep -Milliseconds (
                Get-Random -Minimum 80 -Maximum 400
            )
        }
    }
}

try {

    $testExtended = $false

    do {

        Start-Sleep -Seconds ($durationMinutes * 60)

        $currentPercent = Get-BatteryPercent
        $drop = $startPercent - $currentPercent

        if ($drop -lt 2 -and -not $testExtended) {

            Write-Host "Extending test..."
            $testExtended = $true
        }
        else {
            break
        }

    } while ($true)
}
finally {
    $workers | Stop-Job 
    $workers | Remove-Job -Force
}

$endPercent = Get-BatteryPercent
$endTime = Get-Date

$drop = $startPercent - $endPercent
$elapsed = ($endTime - $startTime).TotalMinutes

if ($drop -gt 0) {

    $rate = $drop / $elapsed
    $runtime = [int](100 / $rate)
}
else {
    $runtime = 0
}

if ($originalBrightness) {

    try {

        (
            Get-WmiObject `
                -Namespace root/WMI `
                -Class WmiMonitorBrightnessMethods
        ).WmiSetBrightness(1, $originalBrightness)

    }
    catch {}
}

Write-Host $runtime
return [int]$runtime