function Get-BatteryInfo {
    Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
}

function Get-Batteries {
    # Always returns an array (safe for multi-battery systems)
    $b = Get-BatteryInfo
    if (-not $b) { return @() }
    return @($b)
}

function Get-OverallBatteryStatus {

    $batteries = Get-Batteries

    if ($batteries.Count -eq 0) {
        return [pscustomobject]@{
            Exists = $false
            Percent = $null
            Batteries = @()
        }
    }

    $valid = $batteries | Where-Object {
        $null -ne $_.EstimatedChargeRemaining
    }

    if ($valid.Count -eq 0) {
        return [pscustomobject]@{
            Exists = $true
            Percent = $null
            Batteries = $batteries
        }
    }

    $percent = ($valid |
        Measure-Object EstimatedChargeRemaining -Average
    ).Average

    return [pscustomobject]@{
        Exists = $true
        Percent = [math]::Round($percent, 2)
        Batteries = $batteries
    }
}

function Test-IsCharging {
    param($batteryInfo)

    if (-not $batteryInfo -or -not $batteryInfo.Batteries) {
        return $false
    }

    # Multiple batteries: if ANY is charging, treat as charging
    foreach ($b in @($batteryInfo.Batteries)) {
        if ($b.BatteryStatus -eq 2) { return $true }
    }

    return $false
}

function Wait-ForUnplug {
    do {
        Start-Sleep -Seconds 2
        $batteryInfo = Get-OverallBatteryStatus
    } while (Test-IsCharging -batteryInfo $batteryInfo)
}

function Show-BatteryProgress {
    param(
        [datetime]$StartTime,
        [int]$DurationSeconds,
        [object]$Progress,
        [string]$Label = "Battery Test Running"
    )

    $endTime = (Get-Date).AddSeconds($DurationSeconds)
    while ($true) {

        $remaining = ($endTime - (Get-Date)).TotalSeconds
        if ($remaining -le 0) { break }
    
        $remaining = [math]::Ceiling($remaining)
    
        $minutes = [int]($remaining / 60)
        $seconds = $remaining % 60
    
        $Progress.Update.Invoke(
            $Label,
            "Time remaining: ${minutes}m ${seconds}s",
            [int](100 - (($remaining / $DurationSeconds) * 100))
        )
        Start-Sleep 1
    }

    $Progress.Complete.Invoke()
}

function Invoke-BatteryTest {
    param(
        [int]$DurationMinutes = 15,
        [string]$BatteryScriptPath,
        [pscustomobject]$batteryInfo,
        [object]$Progress
    )

    if (-not (Test-Path $BatteryScriptPath)) {
        throw "Battery script not found: $BatteryScriptPath"
    }

    Write-Host "Running battery test..."

    if ($batteryInfo.Exists) {

        $charging = Test-IsCharging -batteryInfo $batteryInfo

        if ($charging) {
            Write-Host ""
            Write-Host "Laptop is currently plugged into power."

            do {
                $choice = Read-Host "Do you want to skip the battery test? (Y/N)"

                if ($choice -match "^[Yy]$") {
                    return @{
                        BatteryResult = "Skipped"
                        BatteryMinutes = 0
                    }
                }
                elseif ($choice -match "^[Nn]$") {
                    Write-Host "Please unplug the charger..."
                    Wait-ForUnplug
                    Write-Host "Now running on battery."
                    break
                }

            } while ($true)
        }

        $batteryFunctionString = ${function:Get-OverallBatteryStatus}.ToString()

        $job = Start-Job -ScriptBlock {
            param($path, $duration, $batteryFunctionString)
            & $path -durationMinutes $duration -GetBatteryInfo $batteryFunctionString
        } -ArgumentList $BatteryScriptPath, $DurationMinutes, $batteryFunctionString

        $cycle = 1
        $durationSeconds = $DurationMinutes * 60

        while ($job.State -in @('Running', 'NotStarted')) {

            $label = if ($cycle -eq 1) {
                "Battery Test Running"
            } else {
                "Extended Battery Test (Cycle $cycle)"
            }

            Show-BatteryProgress `
                -DurationSeconds $durationSeconds `
                -Progress $Progress `
                -Label $label

            $cycle++

            if ($job.State -notin @('Running', 'NotStarted')) { break }
        }

        Wait-Job $job
        $batteryMinutes = Receive-Job $job | Select-Object -Last 1
        Remove-Job $job

        $Progress.Complete.Invoke()

    } else {
        return @{
            BatteryResult = "No battery detected"
            BatteryHours = 0
            BatteryMinutes = 0
        }
    }

    $batteryHours = [int]($batteryMinutes / 60)

    return @{
        BatteryResult = "$batteryHours timer"
        BatteryHours = $batteryHours
        BatteryMinutes = $batteryMinutes
    }
}