# Default path for saved settings
$Global:PowerSettingsFile = Join-Path `
	(Join-Path (Split-Path $PSScriptRoot -Parent) "config") `
	"powerSettings.json"

function Get-ActiveScheme {
    $output = powercfg /getactivescheme
    if ($output -match 'GUID:\s+([a-f0-9-]+)') {
        return $matches[1]
    }
    throw "Could not determine active power scheme."
}

function Set-ActiveScheme {
    param (
        [Parameter(Mandatory)]
        [string]$Guid
    )

    Write-Verbose "Switching to power scheme $Guid"
    powercfg /setactive $Guid | Out-Null
}

function Get-PowerValue {
    param (
        [string]$SubGroup,
        [string]$Setting
    )

    $output = powercfg /query SCHEME_CURRENT $SubGroup $Setting

    $acLine = $output | Select-String "Current AC Power Setting Index"
    $dcLine = $output | Select-String "Current DC Power Setting Index"

    if (-not $acLine -or -not $dcLine) {
        throw "Failed to parse powercfg output for $SubGroup / $Setting"
    }

    $acHex = ($acLine -split '\s+')[-1]
    $dcHex = ($dcLine -split '\s+')[-1]

    return @{
        AC = [int]$acHex
        DC = [int]$dcHex
    }
}

function Save-PowerSettings {
    [CmdletBinding()]
    param (
        [string]$Path = $Global:PowerSettingsFile
    )

    Write-Host "Saving power settings to $Path"

    $settings = @{
        SchemeGuid = Get-ActiveScheme

        MonitorTimeout = Get-PowerValue "SUB_VIDEO" "VIDEOIDLE"
        SleepTimeout   = Get-PowerValue "SUB_SLEEP" "STANDBYIDLE"
        HibernateTimeout = Get-PowerValue "SUB_SLEEP" "HIBERNATEIDLE"
    }

    $settings | ConvertTo-Json -Depth 3 | Set-Content -Path $Path -Encoding UTF8
}

function Set-NoTimeouts {
    [CmdletBinding()]
    param (
        [string]$SchemeGuid
    )

    if ($SchemeGuid) {
        Set-ActiveScheme -Guid $SchemeGuid
    }

    Write-Verbose "Disabling monitor timeout"
    powercfg /change monitor-timeout-ac 0
    powercfg /change monitor-timeout-dc 0

    Write-Verbose "Disabling sleep timeout"
    powercfg /change standby-timeout-ac 0
    powercfg /change standby-timeout-dc 0

    Write-Verbose "Disabling hibernate timeout"
    powercfg /change hibernate-timeout-ac 0
    powercfg /change hibernate-timeout-dc 0
}

function Restore-PowerSettings {
    [CmdletBinding()]
    param (
        [string]$Path = $Global:PowerSettingsFile,
        [switch]$RestoreScheme
    )

    if (-not (Test-Path $Path)) {
        throw "Power settings file not found: $Path"
    }

    Write-Host "Restoring power settings from $Path"

    $settings = Get-Content $Path | ConvertFrom-Json

    if ($RestoreScheme -and $settings.SchemeGuid) {
        Set-ActiveScheme -Guid $settings.SchemeGuid
    }

    # Restore monitor timeout
    powercfg /change monitor-timeout-ac $settings.MonitorTimeout.AC
    powercfg /change monitor-timeout-dc $settings.MonitorTimeout.DC

    # Restore sleep timeout
    powercfg /change standby-timeout-ac $settings.SleepTimeout.AC
    powercfg /change standby-timeout-dc $settings.SleepTimeout.DC

    # Restore hibernate timeout
    powercfg /change hibernate-timeout-ac $settings.HibernateTimeout.AC
    powercfg /change hibernate-timeout-dc $settings.HibernateTimeout.DC
}