function Get-BatteryInfo {
    Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
}

function Test-IsCharging {
    param($batteryInfo)
    $battery = $batteryInfo.Battery
    if (-not $battery) { return $false }
    return $battery.BatteryStatus -eq 2  # 2 = charging
}

function Wait-ForUnplug {
    do {
        Start-Sleep -Seconds 2
        $batteryInfo = Get-Battery
    } while (Test-IsCharging $batteryInfo)
}

function Show-BatteryProgress {
    param(
        [datetime]$StartTime,
        [int]$DurationSeconds,
        [object]$Progress,
        [string]$Label = "Battery Test Running"
    )

    while ((Get-Date) -lt $StartTime.AddSeconds($DurationSeconds)) {

        $elapsed = [int]((Get-Date) - $StartTime).TotalSeconds

        if ($elapsed -gt $DurationSeconds) {
            $elapsed = $DurationSeconds
        }

        $remaining = $DurationSeconds - $elapsed

        $minutes = [int]($remaining / 60)
        $seconds = $remaining % 60

        $percent = [int](($elapsed / $DurationSeconds) * 100)

        $Progress.Update.Invoke(
            $Label,
            "Time remaining: ${minutes}m ${seconds}s",
            $percent
        )

        Start-Sleep 1
    }

    $Progress.Complete.Invoke()
}
function Get-Battery{

    $battery = Get-BatteryInfo
    $batteryExists = $false

    if ($null -ne $battery) { $batteryExists = $true }

    return [PSCustomObject]@{
        BatteryExists = $batteryExists
        Battery = $battery
    }
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

    if ($batteryInfo.BatteryExists) {
        $charging = Test-IsCharging -batteryInfo $batteryInfo

        if ($charging) {
            Write-Host ""
            Write-Host "Laptop is currently plugged into power."

            do {
                $choice = Read-Host "Do you want to skip the battery test? (Y/N)"

                if ($choice -match "^[Yy]$") {
                    Write-Host "Skipping battery test as requested."
                    return @{
                        BatteryResult  = "Skipped"
                        BatteryMinutes = 0
                    }
                }
                elseif ($choice -match "^[Nn]$") {
                    Write-Host "Please unplug the charger to continue..."
                    Wait-ForUnplug
                    Write-Host "Laptop is now on battery. Starting test..."
                    break
                }

            } while ($true)
        }

        $batteryFunction = ${function:Get-BatteryInfo}
        # Run external battery script as a job
        $job = Start-Job -ScriptBlock {
            param($path, $duration, $batteryFunction)
        
            & $path -durationMinutes $duration -GetBatteryInfo $batteryFunction
        } -ArgumentList $BatteryScriptPath, $DurationMinutes, $batteryFunction

        $cycle = 1
    $durationSeconds = $DurationMinutes * 60

    while ($job.State -in @('Running', 'NotStarted')) {

        $cycleLabel = if ($cycle -eq 1) {
            "Battery Test Running"
        }
        else {
            "Extended Battery Test (Cycle $cycle)"
        }

        $cycleStart = Get-Date

        Show-BatteryProgress `
            -StartTime $cycleStart `
            -DurationSeconds $durationSeconds `
            -Progress $Progress `
            -Label $cycleLabel

        $cycle++

        if ($job.State -notin @('Running', 'NotStarted')) {
            break
        }
    }

        Wait-Job $job
        $batteryMinutes = Receive-Job $job | Select-Object -Last 1
        Remove-Job $job

        # This completes progress bar
        $Progress.Complete.Invoke()
    } else {
        Write-Host "No battery detected. Skipping battery test."
        return @{
            BatteryResult  = "No battery detected"
            BatteryHours = 0
            BatteryMinutes = 0
        }
    }

    $batteryHours = [int]($batteryMinutes / 60)

    return @{
        BatteryResult  = "$batteryHours timer"
        BatteryHours = $batteryHours
        BatteryMinutes = $batteryMinutes
    }
}