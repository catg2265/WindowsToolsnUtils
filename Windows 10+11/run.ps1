# Ensure script runs as Administrator
if (-not ([Security.Principal.WindowsPrincipal] `
[Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole]::Administrator)) {

    Write-Host "Restarting script with Administrator privileges..."
    
    Start-Process powershell `
	"-NoExit -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
	-WorkingDirectory $PSScriptRoot `
	-Verb RunAs

    exit
}
$ErrorLog = Join-Path $PSScriptRoot "error_log.txt"

$ErrorActionPreference = "Stop"

Register-EngineEvent PowerShell.Exiting -Action {
    $global:LASTERROR | Out-String | Out-File $using:ErrorLog -Append
}
# Ensure window maximises
Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win {
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@

$hwnd = (Get-Process -Id $pid).MainWindowHandle
[Win]::ShowWindow($hwnd, 3)  # 3 = maximize

# Save current power plan settings
$acTimeoutOriginal = (powercfg /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE).Split()[7]
$dcTimeoutOriginal = (powercfg /query SCHEME_CURRENT SUB_VIDEO VIDEOIDLE).Split()[7]
$acSleepOriginal = (powercfg /query SCHEME_CURRENT SUB_SLEEP STANDBYIDLE).Split()[7]
$dcSleepOriginal = (powercfg /query SCHEME_CURRENT SUB_SLEEP STANDBYIDLE).Split()[7]

# Disable display off
powercfg /change monitor-timeout-ac 0
powercfg /change monitor-timeout-dc 0

# Disable sleep
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0

# Detect USB location
$scriptRoot = $PSScriptRoot

# Paths on USB
$batteryScript = Join-Path $scriptRoot "battery_test.ps1"
$wingetInstaller = Join-Path $scriptRoot "windows-package-manager-winget-2025-1213-120-0.msixbundle"
$xmlFile = Join-Path $scriptRoot "defaultapps.xml"
$JSONPath = Join-Path $scriptRoot "packages.json"

# ------------------------
# Ensure Windows Package Manager (winget) is available
# ------------------------

function Install-Winget {
    Write-Host "Attempting to install Winget..."

    if (Test-Path $wingetInstaller) {
        try {
            Write-Host "Using local App Installer bundle..."
            Add-AppxPackage -Path $wingetInstaller -ErrorAction Stop
            Write-Host "Winget installed successfully from local installer."
            return $true
        } catch {
            Write-Warning "Failed local install: $($_.Exception.Message)"
        }
    }

    # Try remote download
    $downloadUrl = "https://aka.ms/Microsoft.WindowsAppInstaller.msixbundle"
    $tmpPath = Join-Path $env:TEMP "AppInstaller.msixbundle"

    try {
        Write-Host "Downloading App Installer bundle..."
        Invoke-WebRequest -Uri $downloadUrl -OutFile $tmpPath -UseBasicParsing -ErrorAction Stop
        Add-AppxPackage -Path $tmpPath -ErrorAction Stop
        Write-Host "Winget installed successfully from online installer."
        return $true
    } catch {
        Write-Warning "Failed to install Winget automatically."
        return $false
    }
}
function Repair-WingetSources {

    Write-Host "Checking Winget sources..." -ForegroundColor Cyan

    try {
        # 1. Force reset if anything is corrupted
        winget source reset --force 2>&1 | Out-Null

        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Winget source reset failed:"
            Write-Warning $resetResult
        } else {
            Write-Host "Winget source reset successful"
        }

        # 2. Always update sources after reset
        winget source update | Out-Null

        Write-Host "Winget sources updated." -ForegroundColor Green
    }
    catch {
        Write-Warning "Winget source repair failed: $($_.Exception.Message)"
    }

    # 3. Verify sources are usable
    $sources = winget source list 2>&1

    if ($sources -match "No sources" -or $sources -match "error") {
        Write-Host "Re-adding default sources..." -ForegroundColor Yellow

        winget source reset --force | Out-Null
        winget source update | Out-Null
    }

    Write-Host "Winget source check complete." -ForegroundColor Green
}
function Repair-AppInstaller {

    Write-Host "Checking Winget / App Installer health..." -ForegroundColor Cyan

    # --- 1. Check if winget command exists ---
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        Write-Warning "Winget not found. Attempting reinstall..."

        $bundle = Get-ChildItem "$PSScriptRoot\*.msixbundle" -ErrorAction SilentlyContinue | Select-Object -First 1

        if ($bundle) {
            Add-AppxPackage -Path $bundle.FullName -ErrorAction Stop
        } else {
            Invoke-WebRequest "https://aka.ms/Microsoft.WindowsAppInstaller.msixbundle" -OutFile "$env:TEMP\winget.msixbundle"
            Add-AppxPackage -Path "$env:TEMP\winget.msixbundle"
        }
    }

    # --- 2. Reset sources ---
    winget source reset --force | Out-Null
    winget source update | Out-Null

    # --- 3. Check if winget source is enabled ---
    $sources = winget source list 2>$null

    if ($sources -notmatch "winget.*true") {

        Write-Warning "Winget source is missing or disabled. Attempting repair..."

        # Clear cache (only once here)
        Remove-Item "$env:LOCALAPPDATA\Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState" `
            -Recurse -Force -ErrorAction SilentlyContinue

        # Kill App Installer processes
        Get-Process -Name AppInstaller -ErrorAction SilentlyContinue | Stop-Process -Force

        # Re-add source manually
        winget source add `
            -n winget `
            -t Microsoft.PreIndexed.Package `
            -a https://cdn.winget.microsoft.com/cache

        winget source update | Out-Null

        $sources = winget source list 2>$null
    }

    # --- 4. If still broken → re-register App Installer ---
    if ($sources -notmatch "winget.*true") {

        Write-Warning "Winget still broken. Re-registering App Installer..."

        try {
            Add-AppxPackage -RegisterByFamilyName -MainPackage Microsoft.DesktopAppInstaller_8wekyb3d8bbwe -ErrorAction Stop
        }
        catch {
            Write-Warning "Re-register failed. Attempting full reinstall..."

            $bundle = Get-ChildItem "$PSScriptRoot\*.msixbundle" -ErrorAction SilentlyContinue | Select-Object -First 1

            if ($bundle) {
                Add-AppxPackage -Path $bundle.FullName -ErrorAction Stop
            } else {
                Invoke-WebRequest "https://aka.ms/Microsoft.WindowsAppInstaller.msixbundle" -OutFile "$env:TEMP\winget.msixbundle"
                Add-AppxPackage -Path "$env:TEMP\winget.msixbundle"
            }
        }

        winget source reset --force | Out-Null
        winget source update | Out-Null
    }

    # --- 5. Final validation ---
    $sources = winget source list 2>$null

    if ($sources -notmatch "winget.*true") {
        throw "Winget is still not functional after repair attempts."
    }

    # --- 6. Disable msstore for automation ---
    winget source disable msstore | Out-Null

    Write-Host "Winget is healthy and ready." -ForegroundColor Green
}

Repair-AppInstaller

if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
    Write-Host "Winget not detected."
    $installed = Install-Winget

    if (-not $installed) {
        Write-Host "Please install Windows Package Manager (Winget) manually and re-run this script."
        exit
    }
} else {
    Write-Host "Winget is already installed."
}

# -----------------------------
# Make Errors stop script instead of silently failing
# -----------------------------

$ErrorActionPreference = "Stop"

# -----------------------------
# Disable msstore 
# -----------------------------
try {
    winget source disable msstore | Out-Null
} catch {}

Write-Host "Resetting Winget sources..." -ForegroundColor Cyan

Repair-WingetSources

Remove-Item "$env:LOCALAPPDATA\Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\*.json" -Force -ErrorAction SilentlyContinue

winget source reset --force | Out-Null
winget source update | Out-Null

winget settings --enable LocalManifestFiles

# -----------------------------
# Load JSON config
# -----------------------------
if (-not (Test-Path $JSONPath)) {
    throw "Package config not found: $JSONPath"
}

$jsonRaw = Get-Content $JSONPath -Raw -Encoding UTF8

try {
    $config = $jsonRaw | ConvertFrom-Json -ErrorAction Stop
}
catch {
    Write-Host "JSON FAILED TO PARSE" -ForegroundColor Red
    Write-Host $_.Exception.Message
    Write-Host "`n--- RAW CONTENT ---"
    Write-Host $jsonRaw
    Read-Host "Press Enter to exit"
    exit
}

$packages = $config.packages

# -----------------------------
# INSTALL
# -----------------------------
Write-Host "Starting package installation..."

#$wingetPath = (Get-Command winget).Source

$results = @()

foreach ($pkg in $packages) {

    $id = $pkg.id

    try {
        Write-Host "Installing $id ..."

        #if ($pkg.override) {
        #    $arguments += $pkg.override
        #}

        $arguments = @()
        $maxRetries = 2
        $attempt = 0
        $success = $false

        while (-not $success -and $attempt -lt $maxRetries) {

            $attempt++

            Write-Host "Installing $id (attempt $attempt)..." -ForegroundColor Cyan

            $arguments = @(
                "upgrade",
                "--id", $id
                "-e"
                "--silent"
                "--accept-package-agreements"
                "--accept-source-agreements"
            )

            $process = Start-Process winget `
                -ArgumentList $arguments `
                -Wait -PassThru -NoNewWindow

            if ($process.ExitCode -eq 0 -or $process.ExitCode -eq 3010) {
                $success = $true
                break
            }

            Write-Warning "Install failed for $id (exit $($process.ExitCode)). Repairing sources..."

            Repair-WingetSources
        }

if (-not $success) {
    throw "Final failure installing $id after retries"
}

        if ($process.ExitCode -eq 0) {
            Write-Host "[OK] $id installed"
            $results += [PSCustomObject]@{ Id = $id; Status = "Success" }
            continue
        }
        else {
			throw "Winget failed with exit code $($process.ExitCode)"
		}
    }
    catch {
        Write-Warning ("Winget failed for {0}: {1}" -f $id, $_.Exception.Message)

        # Fallback (Chrome etc.)
        if ($pkg.fallback) {

            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) "$id.exe"

            Write-Host "Downloading fallback for $id..."
            Invoke-WebRequest -Uri $pkg.fallback.url -OutFile $tmp

            if ($pkg.fallback.sha256) {
                $hash = (Get-FileHash $tmp -Algorithm SHA256).Hash
                if ($hash -ne $pkg.fallback.sha256) {
                    throw "Hash mismatch for $id"
                }
            }

            Start-Process $tmp -ArgumentList $pkg.fallback.silentArgs -Wait

            Write-Host "[OK] $id installed via fallback"
            $results += [PSCustomObject]@{ Id = $id; Status = "FallbackSuccess" }
            continue
        }

        $results += [PSCustomObject]@{ Id = $id; Status = "Failed" }
    }
}

# -----------------------------
# SUMMARY
# -----------------------------
Write-Host "`n--- INSTALL SUMMARY ---"

$results | ForEach-Object {
    Write-Host "$($_.Id): $($_.Status)"
}

Write-Host "Applying default apps..."
Write-Progress -Activity "Installing Packages" -Status "Applying default apps..." -PercentComplete 100

if (!(Test-Path $xmlFile)) {
    Write-Warning "Default apps XML not found at $xmlFile. Skipping default apps."
} else {
    try {
        dism /online /Import-DefaultAppAssociations:"$xmlFile"
        Write-Host "Default apps imported successfully for new users."
    } catch {
        Write-Warning "Failed to import default apps XML: $($_.Exception.Message)"
    }
}

# Complete progress bar
Write-Progress -Activity "Installing Packages" -Completed -Status "All packages installed and defaults applied."
Write-Host "Installation complete."

# ----------------------
# 2️ Run Battery Test with Progress
# ----------------------
$durationMinutes = 15  # matches battery test duration
Write-Host "Running battery test..."

# Check if system has a battery
$batteryCheck = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
$batteryExists = $null -ne $batteryCheck

if ($batteryExists) {

    # Check if currently charging
    $charging = $batteryCheck.BatteryStatus -eq 2  # 2 = charging

    if ($charging) {
        Write-Host ""
        Write-Host "Laptop is currently plugged into power."

        do {
            $choice = Read-Host "Do you want to skip the battery test? (Y/N)"

            if ($choice -match "^[Yy]$") {
                Write-Host "Skipping battery test as requested."
                $batteryResult = "Skipped"
                $batteryMinutes = 0
                break
            } elseif ($choice -match "^[Nn]$") {
                Write-Host "Please unplug the charger to continue..."
                
                # Wait until unplugged
                do {
                    Start-Sleep -Seconds 2
                    $batteryCheck = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
                    $charging = $batteryCheck.BatteryStatus -eq 2

                } while ($charging)

                Write-Host "Laptop is now on battery. Starting test..."
                break
            }

        } while ($true)
    }

    # Only run test if not skipped
    if ($batteryResult -ne "Skipped") {

        # Run battery test as a job so we can track progress
        $job = Start-Job -ScriptBlock {
            param($batteryScript)
            & $batteryScript
        } -ArgumentList $batteryScript

        $durationSeconds = $durationMinutes * 60
        for ($i=0; $i -lt $durationSeconds; $i++) {
            $percent = [int](($i / $durationSeconds) * 100)

            $remainingSeconds = $durationSeconds - $i
            $remainingMinutes = [int](($remainingSeconds % 3600) / 60)
            $remainingSec = $remainingSeconds % 60
            $timeLeft = "$remainingMinutes min $remainingSec sec"

            Write-Progress `
                -Activity "Battery Test Running" `
                -Status "Time remaining: $timeLeft" `
                -PercentComplete $percent

            Start-Sleep -Seconds 1
        }

        Wait-Job $job
        $batteryMinutes = Receive-Job $job | Select-Object -Last 1
        Remove-Job $job

        Write-Progress -Activity "Battery Test Running" -Completed -Status "Battery test complete"
    }
}
else {
    Write-Host "No battery detected. Skipping battery test."
    $batteryResult = "No battery detected"
}

# Convert minutes → hours
if ($batteryResult -eq "Skipped" -or $batteryResult -eq "No battery detected") {
    # keep existing value
}
else {
    $batteryHours = [int]($batteryMinutes / 60)
    $batteryResult = "$batteryHours timer"
}
Write-Host "Battery result = " + $batteryResult

# ----------------------
# 3️ Gather System Info
# ----------------------
# Disk detection
function GetDisks {
    $disks = Get-PhysicalDisk | Where-Object { $_.BusType -ne 'USB' -and $_.FriendlyName -ne 'Verbatim STORE N GO'}
    $diskString = ""
    $diskObjects = @()   # New array to store structured disk info

    $validSizes = @(
        64, 120, 128, 248, 256, 320, 480, 
        500, 512, 960, 1000, 1024, 2000, 
        3000, 4000, 6000, 8000, 10000, 
        12000, 16000, 20000, 22000
    )
    
    foreach ($disk in $disks) {
        $newLine = ""
        $diskSizeGB = [math]::Round($disk.Size / 1GB)
        $nearestSizeUp = $validSizes | Where-Object { $_ -ge $diskSizeGB } | Sort-Object | Select-Object -First 1

        # Build HTML row
        if ($disk.MediaType -eq 'SSD' -or $disk.SpindleSpeed -eq 0) {
            $newLine = "<tr><td>SSD</td><td>$nearestSizeUp GB</td></tr>"
            $type = "SSD"
			$diskModel = $($disk.FriendlyName)
        }
        elseif ($disk.MediaType -eq 'HDD' -or $disk.SpindleSpeed -gt 0) {
            $newLine = "<tr><td>Harddisk</td><td>$nearestSizeUp GB</td></tr>"
            $type = "HDD"
			$diskModel = $($disk.FriendlyName)
        }
        else {
            Write-Host "Unknown: $($disk.FriendlyName) $diskSizeGB GB not included in specsheet"
            continue
        }

        # Append to HTML string
        if ($newLine) {
            if ($diskString) {
                $diskString += "`n$newLine"
            } else {
                $diskString = $newLine
            }
        }

        # Add to structured disk array for pricing
        $diskObjects += [PSCustomObject]@{
            SizeGB = $nearestSizeUp
            MediaType = $type
			Model = $diskModel
        }
    }

    # Return both: HTML string and disk objects
    return @{
        Html = $diskString
        Disks = $diskObjects
    }
}

$diskInfo = GetDisks
$physicalDisks = $diskInfo.Disks

$computerSystem = Get-CimInstance Win32_ComputerSystem
$cpu = Get-CimInstance Win32_Processor
$gpu = Get-CimInstance Win32_VideoController | Select-Object -First 1
$bios = Get-CimInstance Win32_BIOS
$os = Get-CimInstance Win32_OperatingSystem

$model = $computerSystem.Model
$manufacturer = $computerSystem.Manufacturer
$cpuName = $cpu.Name
$ram = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB)
$diskRows = $diskInfo.Html
$gpuName = $gpu.Name



# Screen size detection
function Get-ScreenSize {
    $monitors = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams
    foreach ($m in $monitors) {
        $width = $m.MaxHorizontalImageSize
        $height = $m.MaxVerticalImageSize
        if ($width -gt 0 -and $height -gt 0) {
            $diagonal = [math]::Sqrt(($width*$width)+($height*$height))
            $inches = [math]::Round($diagonal / 2.54,1)
            return "$inches`""
        }
    }
    return "Unknown"
}

$screenSize = Get-ScreenSize

# ----------------------
# Price Estimation Logic
# ----------------------
function Get-EstimatedPrice {

    param(
        $cpuName,
        $ram,
        $diskInfo,   # Pass $diskInfo object containing Disks info
        $gpuName,
        $screenSize,
        $batteryHours
    )

    $price = 0

    # ----- CPU tier contribution (DKK) -----
    if ($cpuName -match "i9|Ryzen 9") { $price += 4500 }    # ~€600
    elseif ($cpuName -match "i7|Ryzen 7") { $price += 3400 } # ~€450
    elseif ($cpuName -match "i5|Ryzen 5") { $price += 2400 } # ~€320
    elseif ($cpuName -match "i3|Ryzen 3") { $price += 1650 } # ~€220
    else { $price += 1125 }                                  # ~€150

    # ----- RAM contribution -----
    if ($ram -ge 32) { $price += 1350 }   # ~€180
    elseif ($ram -ge 16) { $price += 900 } # ~€120
    elseif ($ram -ge 8) { $price += 450 }  # ~€60

    # ----- Storage contribution (based on all disks) -----
    foreach ($disk in $diskInfo.Disks) {
        $diskType = $disk.MediaType
        $diskSize = $disk.SizeGB

        # SSD bonus
        if ($diskType -eq "SSD") {
            if ($diskSize -ge 1000) { $price += 750 }   # ~€100
            elseif ($diskSize -ge 512) { $price += 525 } # ~€70
            elseif ($diskSize -ge 256) { $price += 300 } # ~€40
        }
        # HDD bonus
        elseif ($diskType -eq "HDD") {
            if ($diskSize -ge 1000) { $price += 750 }    # ~€100
            elseif ($diskSize -ge 512) { $price += 525 }  # ~€70
            elseif ($diskSize -ge 256) { $price += 300 }  # ~€40
        }
    }

    # ----- GPU bonus -----
    if ($gpuName -match "RTX|GTX|RX") { $price += 900 } # ~€120

    # ----- Screen size adjustment -----
    if ($batteryHours -gt 0) { # Only if there’s a battery
        if ($screenSize -match "17") { $price += 300 } # ~€40
    }

    # ----- Battery quality adjustment -----
    if ($batteryHours -ge 8) { $price += 600 }    # ~€80
    elseif ($batteryHours -ge 5) { $price += 300 } # ~€40
    elseif ($batteryHours -le 2) { $price -= 300 } # ~€-40

    # Minimum safeguard
    if ($price -lt 600) { $price = 600 } # ~€80

    return [int]($price)
}

$estimatedPrice = Get-EstimatedPrice `
    -cpuName $cpuName `
    -ram $ram `
    -diskInfo $physicalDisks `
    -gpuName $gpuName `
    -screenSize $screenSize `
    -batteryHours $batteryHours

# ----------------------
# 4️ Create Results Folder
# ----------------------
$resultsFolder = Join-Path $scriptRoot "Results"
if (!(Test-Path $resultsFolder)) { New-Item -ItemType Directory -Path $resultsFolder }

$safeModel = $model -replace '[\\/:*?"<>|]', ''
$destinationFolder = Join-Path $resultsFolder $safeModel
$counter = 1
while (Test-Path $destinationFolder) {
    $destinationFolder = Join-Path $resultsFolder "$safeModel-$counter"
    $counter++
}

# Create the folder
New-Item -ItemType Directory -Path $destinationFolder

$specSheet = Join-Path $destinationFolder "spec_sheet.html"
$internalFile = Join-Path $destinationFolder "internal_report.txt"

# ----------------------
# 5️ Create Professional A4 Spec Sheet
# ----------------------
if ($batteryExists) {
    $batteryRow = "<tr><td>Batteritid op til</td><td class='highlight'>$batteryResult</td></tr>"
	$displayRow = @"
<tr><td>Sk<span>&#230;</span>rm</td><td>$screenSize</td></tr>
"@
}
else {
    $batteryRow = ""
	$displayRow = ""
}

$html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset='UTF-8'>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Spec Sheet</title>
<style>
body { font-family: Arial, sans-serif; width:210mm; margin:auto; background-color:#fff; color:#000; }
.container { border:2px solid #000; padding:20px 30px; box-sizing:border-box; }

/* Flex header for model*/
.header {
    display: flex;
    justify-content: space-between; /* model left, price right */
    align-items: center;
    margin-bottom: 25px;
}
.header .model {
    font-size: 28px;
    font-weight: bold;
    margin-left: 5px;
}

/* Full-width divider with margins */
.divider {
    border-bottom: 2px solid #000;
    margin: 15px 0 25px 0; /* spacing above/below */
}

/* Price styling */
.price-section {
    display: flex;
    justify-content: flex-start; /* left-align price */
    font-size: 32px;
    font-weight: bold;
    margin-top: 10px;
}
.price-label {
    margin-left: 5px; /* space between label and value */
}

.price-value {
    color: #1a73e8; /* highlight color */
    margin-left: 180px;
}

/* Table styling */
table { width:100%; border-collapse:collapse; font-size:18px; }
td { padding:12px 10px; border-bottom:1px solid #ccc; }
td:first-child { font-weight:bold; width:35%; }
.highlight { font-size:22px; font-weight:bold; color:#1a73e8; }

@media print { 
    body { margin:0; } 
    .container { border:none; padding:0; } 
}
</style>
</head>
<body>
<div class='container'>

<!-- Header with Model and Price and Divider -->
<div class="header">
    <div class="model">Model: __________________</div>
</div>
<div class="divider"></div>

<table>
<tr><td>Processor</td><td>$cpuName</td></tr>
<tr><td>Hukommelse</td><td>$ram GB RAM</td></tr>
$diskRows
<tr><td>Grafik</td><td>$gpuName</td></tr>
$displayRow
<tr><td>Styresystem</td><td>$($os.Caption)</td></tr>
$batteryRow
</table>
<div class="divider"></div>

<div class="price-section">
    <span class="price-label">Pris: <span class="price-value">$estimatedPrice,-</span></span>
</div>
</div>
</body>
</html>
"@

[System.IO.File]::WriteAllText($specSheet, $html, [System.Text.UTF8Encoding]::new($false))

# ----------------------
# 6️ Create Internal Technician Report
# ----------------------
# Convert each disk to a readable string
$diskDetails = $physicalDisks | ForEach-Object { 
    "$($_.SizeGB) GB $($_.MediaType) - $($_.Model)" 
} | Out-String
$internalReport = @"
INTERNAL TECH REPORT
====================

Manufacturer: $manufacturer
Model: $model
Serial Number: $($bios.SerialNumber)

CPU: $cpuName
RAM: $ram GB

Disks: 
$diskDetails

GPU: $gpuName

Screen Size: $screenSize

Operating System:
$($os.Caption)

Battery Runtime:
$batteryResult

Estimated Refurbished Value:
€$estimatedPrice
"@

$internalReport | Out-File $internalFile

# Restore original power settings
if ($null -ne $acTimeoutOriginal) { powercfg /change monitor-timeout-ac $acTimeoutOriginal }
if ($null -ne $dcTimeoutOriginal) { powercfg /change monitor-timeout-dc $dcTimeoutOriginal }
if ($null -ne $acSleepOriginal) { powercfg /change standby-timeout-ac $acSleepOriginal }
if ($null -ne $dcSleepOriginal) { powercfg /change standby-timeout-dc $dcSleepOriginal }

Write-Host "Original power settings restored."

Write-Host ""
Write-Host "Professional spec sheet and internal report created in:"
Write-Host $destinationFolder
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Yellow
Pause