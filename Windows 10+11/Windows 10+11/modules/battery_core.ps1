function Get-BatteryInfo {
    Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
}

function Get-Batteries {
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

    $percent = (
        $valid |
        Measure-Object EstimatedChargeRemaining -Average
    ).Average

    [pscustomobject]@{
        Exists = $true
        Percent = [math]::Round($percent, 2)
        Batteries = $batteries
    }
}