function Test-IsCharging {
    param($batteryInfo)

    if (-not $batteryInfo -or -not $batteryInfo.Batteries) {
        return $false
    }

    foreach ($b in @($batteryInfo.Batteries)) {
        if ($b.BatteryStatus -eq 2) {
            return $true
        }
    }

    return $false
}
function Wait-ForUnplug {

    do {
        Start-Sleep 2
        $batteryInfo = Get-OverallBatteryStatus
    }
    while (Test-IsCharging $batteryInfo)
}
function Update-BatteryProgress {

    param(
        [datetime]$StartTime,
        [int]$DurationSeconds,
        [object]$Progress,
        [string]$Label
    )

    $elapsed = ((Get-Date) - $StartTime).TotalSeconds
    $remaining = [math]::Max(0, $DurationSeconds - $elapsed)

    $remaining = [int][math]::Ceiling($remaining)

    $minutes = [int][math]::Floor($remaining / 60)
    $seconds = [int]$remaining % 60

    $percent = [math]::Min(
        100,
        [int](($elapsed / $DurationSeconds) * 100)
    )

    $Progress.Update.Invoke(
        $Label,
        ("Time remaining: {0:D2}m {1:D2}s" -f $minutes, $seconds),
        $percent
    )
}
function Invoke-BatteryTest {

    param(
        [int]$DurationMinutes = 15,
        [string]$BatteryScriptPath,
        [string]$MainScriptPath,
        [object]$Progress
    )
    $batteryCorePath = Join-Path $MainScriptPath "modules\battery_core.ps1"
    . $batteryCorePath
    
    $batteryInfo = Get-OverallBatteryStatus

    if (-not $batteryInfo.Exists) {

        return @{
            BatteryResult  = "No battery detected"
            BatteryHours   = 0
            BatteryMinutes = 0
        }
    }

    if (Test-IsCharging $batteryInfo) {

        Write-Host ""
        Write-Host "Laptop is plugged into power."

        do {

            $choice = Read-Host "Skip battery test? (Y/N)"

            if ($choice -match '^[Yy]$') {

                return @{
                    BatteryResult  = "Skipped"
                    BatteryHours   = 0
                    BatteryMinutes = 0
                }
            }

            if ($choice -match '^[Nn]$') {

                Write-Host "Please unplug charger..."
                Wait-ForUnplug
                Write-Host "Running on battery."
                break
            }

        } while ($true)
    }

    $job = Start-Job -ScriptBlock {

        param(
            $duration,
            $testPath,
            $corePath
        )

        & $testPath -durationMinutes $duration -batteryCorePath $corePath

    } -ArgumentList @(
        $DurationMinutes,
        $BatteryScriptPath,
        $batteryCorePath
    )

    $startTime = Get-Date
    $durationSeconds = $DurationMinutes * 60
    $cycle = 1
    $elapsed = 0

    while ($job.State -in @('Running', 'NotStarted')) {

        $elapsed = ((Get-Date) - $startTime).TotalSeconds

        $label = if ($cycle -eq 1) {
            "Battery Test Running"
        }
        else {
            "Extended Battery Test (Cycle $cycle)"
        }

        Update-BatteryProgress `
            -StartTime $startTime `
            -DurationSeconds $durationSeconds `
            -Progress $Progress `
            -Label $label

        if ($elapsed -ge $durationSeconds) {
            $cycle++
            $startTime = Get-Date
        }

        Start-Sleep 1
        $job = Get-Job -Id $job.Id
    }

    Wait-Job $job

    $batteryMinutes = Receive-Job $job |
        Select-Object -Last 1

    Remove-Job $job

    $Progress.Complete.Invoke()

    $batteryHours = [int]($batteryMinutes / 60)

    return @{
        BatteryResult  = "$batteryHours timer"
        BatteryHours   = $batteryHours
        BatteryMinutes = $batteryMinutes
    }
}