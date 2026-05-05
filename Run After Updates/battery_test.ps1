# Battery Quick Test – Safe Version 

# Requirements: NirCmd.exe in same folder 
$nircmd = Join-Path $PSScriptRoot "nircmd.exe"

$durationMinutes = 15 
$logFile = "$env:USERPROFILE\Desktop\battery_test_result.txt" 
Write-Host "Preparing laptop for battery test..." 

# Store original brightness 
try { 
	$brightnessObj = Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightness 
	$originalBrightness = $brightnessObj.CurrentBrightness 
} 
	catch { 
		Write-Host "Cannot read brightness, skipping restore later." 
		$originalBrightness = $null 
	} 

# Set power plan to Balanced 
powercfg -setactive SCHEME_BALANCED 

# Set brightness to 50% for test 
try { 
	(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,50) 
} 
	catch { 
		Write-Host "Cannot set brightness." 
	} 

# Store original audio state # 0 = unmuted, 1 = muted 
$audioStatus = & $nircmd mutesysvolume 2 # toggle to read current state 
# NirCmd doesn't provide direct read, so we assume user knows original state 
# We'll restore at the end 

# Force mute 
Start-Process $nircmd -ArgumentList "mutesysvolume 1" 
Write-Host "Audio muted for test." 

# Get starting battery level 
$battery = Get-WmiObject Win32_Battery 
$startPercent = $battery.EstimatedChargeRemaining 
$startTime = Get-Date 

Write-Host "Starting battery: $startPercent%" 

# -----------------------------
# Adaptive Moderate CPU Load Job
# -----------------------------
$cpuInfo = Get-CimInstance Win32_Processor | Select-Object -First 1
$cpuFreqMHz = $cpuInfo.MaxClockSpeed        # e.g., 2500 for 2.5 GHz
$cpuCores = $cpuInfo.NumberOfLogicalProcessors

# Determine loop size based on CPU frequency
# Base: 35000 iterations for ~2.5 GHz (moderate load)
$baseFreqMHz = 2500
$baseIterations = 35000
$scaledIterations = [math]::Max(10000, [math]::Round($baseIterations * ($cpuFreqMHz / $baseFreqMHz)))

Write-Host "CPU: $($cpuInfo.Name)"
Write-Host "Cores: $cpuCores, MaxFreq: $cpuFreqMHz MHz, Loop iterations: $scaledIterations"

$loadJob = Start-Job {

    param($iterations)

    while ($true) {
        # Moderate CPU computation, scaled to CPU
        1..$iterations | ForEach-Object { [math]::Sqrt($_) * [math]::Pow($_,0.7) } | Out-Null

        # Small disk activity
        $tmp = "$env:TEMP\battery_test.tmp"
        Get-Random -Minimum 1000 -Maximum 5000 | Out-File $tmp
        Remove-Item $tmp -ErrorAction SilentlyContinue

        # Occasional network activity
        try {
            Invoke-WebRequest -Uri "https://www.wikipedia.org" -UseBasicParsing -TimeoutSec 3 | Out-Null
        } catch {}

        # Short idle like a user reading
        Start-Sleep -Milliseconds 500
    }

} -ArgumentList $scaledIterations

# Wait for test duration 
$testExtended = $false

do {

    Start-Sleep -Seconds ($durationMinutes * 60)

    # Check battery after this cycle
    $battery = Get-WmiObject Win32_Battery
    $currentPercent = $battery.EstimatedChargeRemaining
    $drop = $startPercent - $currentPercent
	
    if ($drop -lt 2 -and -not $testExtended) {

        Write-Host "Battery drop too small ($drop%). Extending test another $durationMinutes minutes..."
        $testExtended = $true

    } else {
        break
    }

} while ($true)

# Get ending battery level 
$battery = Get-WmiObject Win32_Battery 
$endPercent = $battery.EstimatedChargeRemaining 
$endTime = Get-Date 

Write-Host "Ending battery: $endPercent%" 

# Stop CPU load 
Stop-Job $loadJob 
Remove-Job $loadJob 

# Calculate battery usage
$drop = $startPercent - $endPercent
$elapsedMinutes = ($endTime - $startTime).TotalMinutes

if ($drop -le 0) {
    Write-Host "Battery did not drop enough to estimate runtime."
    $runtimeMinutes = 0
}
else {

    # percent used per minute
    $drainRate = $drop / $elapsedMinutes

    # estimated full runtime
    $runtimeMinutes = [int](100 / $drainRate)

    Write-Host "Battery drop: $drop %"
    Write-Host "Drain rate: $([math]::Round($drainRate,3)) % per minute"
    Write-Host "Estimated runtime: $runtimeMinutes minutes"
}

# Prepare output for Spec-Sheet script
$output = [int][math]::Floor($runtimeMinutes)

# Restore brightness 
if ($originalBrightness -ne $null) { 
	try { 
		(Get-WmiObject -Namespace root/WMI -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1,$originalBrightness) 
		Write-Host "Brightness restored." 
	} 
		catch { 
			Write-Host "Could not restore brightness." 
		} 
} 

# Unmute audio 
Start-Process $nircmd -ArgumentList "mutesysvolume 0" 
Write-Host "Audio restored."

Write-Host $output
return $output