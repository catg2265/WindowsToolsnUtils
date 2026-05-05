# Ensure script runs as Administrator
if (-not ([Security.Principal.WindowsPrincipal] `
[Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
[Security.Principal.WindowsBuiltInRole]::Administrator)) {

    Write-Host "Restarting script with Administrator privileges..."
    
    Start-Process powershell `
    "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
    -Verb RunAs

    exit
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
$xmlFile = Join-Path $PSScriptRoot "defaultapps.xml"

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

# ----------------------
# 1️ Run Winget Loop
# ----------------------
$packages = @(
    # Dev / Runtime
    "Oracle.JavaRuntimeEnvironment",
	"AdoptOpenJDK.OpenJDK.17",           
    "Microsoft.DotNet.Runtime.7",
    "Microsoft.VC++2015-2022.Redist.x64",

    # Browsers
    "Google.Chrome",

    # Productivity / Utilities
    "VideoLAN.VLC",
	"TheDocumentFoundation.LibreOffice"
)

# Updates Winget Sources
Write-Host "Updating Winget sources..."
winget source update


# Install All Packages
$total = $packages.Count
for ($i = 0; $i -lt $total; $i++) {
    $pkg = $packages[$i]
    Write-Host "Installing $pkg..."

    # Calculate progress safely
    $percentComplete = if ($total -gt 0) { [int](($i / $total) * 100) } else { 0 }
    Write-Progress -Activity "Installing Packages" -Status "Installing $pkg ($($i+1)/$total)" -PercentComplete $percentComplete

	if ($pkg -eq "Google.Chrome") {
    Write-Host "Installing Chrome..."

    # Option 1: Use Winget if available
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Host "Winget detected. Installing Chrome via Winget..."
        winget install --id Google.Chrome -e --silent --accept-package-agreements --accept-source-agreements --disable-interactivity
    }
    else {
        # Option 2: Streamed download + offline installer
        $tmpPath = Join-Path $env:TEMP "chrome_installer.exe"
        $chromeUrl = "https://dl.google.com/chrome/install/standalonesetup64.exe"

        # Remove previous installer if exists
        if (Test-Path $tmpPath) { Remove-Item $tmpPath -Force }

        Write-Host "Downloading Chrome Standalone Installer (streamed, 1 MB buffer)..."

        $client = [System.Net.Http.HttpClient]::new()
        $response = $client.GetAsync($chromeUrl, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).Result
        $stream = $response.Content.ReadAsStreamAsync().Result
        $fileStream = [System.IO.File]::Create($tmpPath)

        $buffer = New-Object byte[] 1048576  # 1 MB buffer
        while (($read = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $fileStream.Write($buffer, 0, $read)
        }

        $fileStream.Close()
        $stream.Close()
        $client.Dispose()

        Write-Host "Download complete. Running Chrome installer..."
        Start-Process -FilePath $tmpPath -ArgumentList "/silent /norestart" -Wait
        Write-Host "Chrome installation complete."
    }
		# Skip regular Winget loop for Chrome
		continue
	}
	
    # Automatic silent install with overrides for .NET and other MSIs
    if ($pkg -eq "Microsoft.DotNet.Runtime.7") {
		winget install --id $pkg `
			--silent `
			--accept-package-agreements `
			--accept-source-agreements `
			--override "/quiet /norestart"
	}

    $maxAttempts = 3
    $attempt = 0
    $installed = $false
    while (-not $installed -and $attempt -lt $maxAttempts) {
        $attempt++
        try {
            winget install --id $pkg --silent --accept-package-agreements --accept-source-agreements $overrideArgs
            $installed = $true
        } catch {
            Write-Warning "Attempt $attempt failed for $pkg. Retrying..."
            Start-Sleep -Seconds 5
        }
    }
    if (-not $installed) { Write-Warning "$pkg could not be installed automatically." }
}

Write-Host "Applying default apps..."
Write-Progress -Activity "Installing Packages" -Status "Applying default apps..." -PercentComplete 100

if (!(Test-Path $xmlFile)) {
    Write-Warning "Default apps XML not found at $xmlFile. Skipping default apps."
} else {
    try {
        dism /online /Import-DefaultAppAssociations:$xmlFile
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
$batteryExists = $batteryCheck -ne $null

if ($batteryExists) {
    # Loop until laptop is on battery
    do {
        # Refresh battery status
        $batteryCheck = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
        $charging = $batteryCheck.BatteryStatus -eq 2  # 2 = charging

        if ($charging) {
            Write-Host ""
            Write-Host "====================================================="
            Write-Host "ERROR: Laptop is plugged into power. Battery test"
            Write-Host "cannot run while charging. Please unplug AC power."
            Write-Host "====================================================="
            Write-Host "Press any key to continue once the laptop is unplugged..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }

    } while ($charging)  # repeat until unplugged

    Write-Host "Laptop is on battery. Running battery test..."
	
	# Run battery test as a job so we can track progress
	$job = Start-Job -ScriptBlock {
		param($batteryScript)
		& $batteryScript
	} -ArgumentList $batteryScript

	$durationSeconds = $durationMinutes * 60
	for ($i=0; $i -lt $durationSeconds; $i++) {
		$percent = [int](($i / $durationSeconds) * 100)
    
		# Calculate remaining time
		$remainingSeconds = $durationSeconds - $i
		$remainingHours = [int]($remainingSeconds / 3600)
		$remainingMinutes = [int](($remainingSeconds % 3600) / 60)
		$remainingSec = $remainingSeconds % 60
    
		# Format string
		if ($remainingHours -gt 0) {
			$timeLeft = "$remainingHours h $remainingMinutes min"
		} else {
			$timeLeft = "$remainingMinutes min $remainingSec sec"
		}

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
else {
    Write-Host "No battery detected. Skipping battery test."
    $batteryResult = "No battery detected"
}

# Convert minutes → hours
$batteryHours = [int]($batteryMinutes / 60)
$batteryResult = "$batteryHours timer"
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
        $diskType = $disk.Type
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
    -disks $physicalDisks `
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
if ($acTimeoutOriginal -ne $null) { powercfg /change monitor-timeout-ac $acTimeoutOriginal }
if ($dcTimeoutOriginal -ne $null) { powercfg /change monitor-timeout-dc $dcTimeoutOriginal }
if ($acSleepOriginal -ne $null) { powercfg /change standby-timeout-ac $acSleepOriginal }
if ($dcSleepOriginal -ne $null) { powercfg /change standby-timeout-dc $dcSleepOriginal }

Write-Host "Original power settings restored."

Write-Host ""
Write-Host "Professional spec sheet and internal report created in:"
Write-Host $destinationFolder
Write-Host ""
Write-Host "Press any key to exit..." -ForegroundColor Yellow
Pause