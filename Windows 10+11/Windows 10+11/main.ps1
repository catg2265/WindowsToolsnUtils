$scriptroot = $PSScriptRoot

# Logging setup
$LogPath = Join-Path $scriptroot "transcript.txt"
$LogArchivePath = Join-Path $scriptroot "logs"

if (Test-Path $LogPath) {

    if (!(Test-Path $LogArchivePath)) {
        New-Item -ItemType Directory -Path $LogArchivePath | Out-Null
    }

    $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $newName = "transcript_$timestamp.txt"

    Move-Item $LogPath (Join-Path $LogArchivePath $newName) -Force

    # Keep only the 10 newest transcript files
    Get-ChildItem -Path $LogArchivePath -Filter "transcript_*.txt" |
        Sort-Object LastWriteTime -Descending |
        Select-Object -Skip 10 |
        Remove-Item -Force
}

Start-Transcript -Path $LogPath -Append -ErrorAction SilentlyContinue

try{
    
    # Ensure script runs as Administrator
    $principal = [Security.Principal.WindowsPrincipal]::new(
        [Security.Principal.WindowsIdentity]::GetCurrent()
    )
    
    if (-not $PSCommandPath) {
        throw "Cannot self-elevate: script path is not available."
    }

    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {

        Write-Host "Restarting script with Administrator privileges..."

        $exe = if ($PSVersionTable.PSEdition -eq "Core") {
            "pwsh.exe"
        } else {
            "powershell.exe"
        }

        Start-Process $exe -Verb RunAs -ArgumentList (
            "-ExecutionPolicy Bypass -File `"$PSCommandPath`""
        ) -WorkingDirectory $scriptroot

        exit
    }

    # Window maximize (safe check)
    Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Win {
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
    }
"@

    $proc = Get-Process -Id $PID -ErrorAction SilentlyContinue

    if ($proc -and $proc.MainWindowHandle -ne 0) {
        [Win]::ShowWindow($proc.MainWindowHandle, 3)
        }

    #########################################
    #########   Dot Source Modules  #########
    #########################################

    $modules = @(
        "progress_bar_controller.ps1",
        "power_settings.ps1",
        "packages_install.ps1",
        "battery_core.ps1"
        "run_battery_test.ps1",
        "system_info.ps1",
        "reports.ps1"
    )

    foreach ($module in $modules) {
        $path = Join-Path "$scriptroot\modules" $module

        if (Test-Path $path) {
            try {
                . $path
            } catch {
                Write-Warning "Failed to load $module : $_"
            }
        } else {
            Write-Warning "Missing module: $module"
        }
    }

    #########################################
    #########         Setup         #########
    #########################################

    $ErrorActionPreference = "Stop"

    Save-PowerSettings -Verbose
    Set-NoTimeouts -Verbose

    Add-Process "Installing Packages"
    Add-Process "Battery Test"
    Add-Process "Gather System Info"
    Add-Process "Generate Reports"

    Initialize-ProgressBar -Activity "Windows Automation"

    #########################################
    #########    Package Install    #########
    #########################################

    $installProgress = New-ProgressContext -ParentId 1 -Id 2
    Initialize-PackageEnvironment
    Install-PackageManagers
    Install-Packages `
        -JsonPath (Join-Path $scriptroot\config "packages.json") `
        -Progress  $installProgress
    Set-DefaultApps -XmlPath (Join-Path $scriptroot\config "defaultapps.xml")

    Move-Process

    #########################################
    ##########    Battery Test    ###########
    #########################################

    $batteryProgress = New-ProgressContext -ParentId 1 -Id 3
    $batteryInfo = Get-OverallBatteryStatus
    $batteryTest = Invoke-BatteryTest `
        -DurationMinutes 15 `
        -BatteryScriptPath (Join-Path $scriptroot "assets\battery_test.ps1") `
        -MainScriptPath $scriptroot `
        -Progress $batteryProgress

    Move-Process

    #########################################
    ##########     System Info    ###########
    #########################################

    $systemInfo = Get-SystemInfo

    $estimatedPrice = Get-EstimatedPrice `
        -cpuName $systemInfo.CPU `
        -ram $systemInfo.RAM_GB `
        -diskInfo $systemInfo.Disks `
        -gpuName $systemInfo.GPU `
        -screenSize $systemInfo.ScreenSize `
        -batteryHours $batteryTest.BatteryHours

    Move-Process
    
    #########################################
    ##########  Generate Reports  ###########
    #########################################

    $resultFolder = New-ResultsFolder -ScriptRoot $scriptroot -Model $systemInfo.Model

    $specFile = New-SpecSheetFile `
        -DestinationFolder $resultFolder `
        -cpuName $systemInfo.CPU `
        -ram $systemInfo.RAM_GB `
        -diskRows $systemInfo.DiskHtml `
        -gpuName $systemInfo.GPU `
        -screenSize $systemInfo.ScreenSize `
        -os $systemInfo.OS `
        -batteryExists $batteryInfo.Exists `
        -batteryResult $batteryTest.BatteryResult `
        -estimatedPrice $estimatedPrice
    Write-Host "Created: $specFile in $resultFolder"

    $reportFile = New-InternalReportFile `
        -DestinationFolder $resultFolder `
        -manufacturer $systemInfo.Manufacturer `
        -model $systemInfo.Model `
        -serialNumber $systemInfo.SerialNumber `
        -cpuName $systemInfo.CPU `
        -ram $systemInfo.RAM_GB `
        -physicalDisks $systemInfo.Disks `
        -gpuName $systemInfo.GPU `
        -screenSize $systemInfo.ScreenSize `
        -os $systemInfo.OS `
        -batteryResult $batteryTest.BatteryResult `
        -estimatedPrice $estimatedPrice
    Write-Host "Created: $reportFile in $resultFolder"

    #########################################
    ##########      Finalize      ###########
    #########################################

    Restore-PowerSettings -Verbose
    Complete-Progress

    Read-Host "Press any key to exit..." -ForegroundColor Yellow
} catch {
    Stop-Transcript -ErrorAction SilentlyContinue
    $_ | Out-String | Out-File $LogPath -Append
    throw
}
finally{
    Stop-Transcript -ErrorAction SilentlyContinue
}