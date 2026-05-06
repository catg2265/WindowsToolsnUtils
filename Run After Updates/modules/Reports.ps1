# ----------------------
# Create Results Folder Structure
# ----------------------
function New-ResultsFolder {
    param (
        [string]$ScriptRoot,
        [string]$Model
    )

    $resultsFolder = Join-Path $ScriptRoot "Results"
    if (!(Test-Path $resultsFolder)) {
        New-Item -ItemType Directory -Path $resultsFolder | Out-Null
    }

    $safeModel = $Model -replace '[\\/:*?"<>|]', ''
    $destinationFolder = Join-Path $resultsFolder $safeModel

    $counter = 1
    while (Test-Path $destinationFolder) {
        $destinationFolder = Join-Path $resultsFolder "$safeModel-$counter"
        $counter++
    }

    New-Item -ItemType Directory -Path $destinationFolder | Out-Null
    return $destinationFolder
}

# ----------------------
# Create Spec Sheet File
# ----------------------
function New-SpecSheetFile {
    param (
        [string]$DestinationFolder,
        $cpuName,
        $ram,
        $diskRows,
        $gpuName,
        $screenSize,
        $os,
        $batteryExists,
        $batteryResult,
        $estimatedPrice
    )

    $specSheetPath = Join-Path $DestinationFolder "spec_sheet.html"

    if ($batteryExists) {
        $batteryRow = "<tr><td>Batteritid op til</td><td class='highlight'>$batteryResult</td></tr>"
        $displayRow = "<tr><td>Sk<span>&#230;</span>rm</td><td>$screenSize</td></tr>"
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
<title>Spec Sheet</title>
<style>
body { font-family: Arial; width:210mm; margin:auto; }
.container { border:2px solid #000; padding:20px 30px; }

.header { display:flex; justify-content:space-between; margin-bottom:25px; }
.model { font-size:28px; font-weight:bold; }

.divider { border-bottom:2px solid #000; margin:15px 0 25px 0; }

.price-section { font-size:32px; font-weight:bold; }
.price-value { color:#1a73e8; margin-left:180px; }

table { width:100%; border-collapse:collapse; font-size:18px; }
td { padding:12px 10px; border-bottom:1px solid #ccc; }
td:first-child { font-weight:bold; width:35%; }
.highlight { font-size:22px; font-weight:bold; color:#1a73e8; }
</style>
</head>
<body>
<div class='container'>

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
Pris: <span class="price-value">$estimatedPrice,-</span>
</div>

</div>
</body>
</html>
"@

    [System.IO.File]::WriteAllText(
        $specSheetPath,
        $html,
        [System.Text.UTF8Encoding]::new($false)
    )

    return $specSheetPath
}

# ----------------------
# Create Internal Report File
# ----------------------
function New-InternalReportFile {
    param (
        [string]$DestinationFolder,
        $manufacturer,
        $model,
        $serialNumber,
        $cpuName,
        $ram,
        $physicalDisks,
        $gpuName,
        $screenSize,
        $os,
        $batteryResult,
        $estimatedPrice
    )

    $internalFilePath = Join-Path $DestinationFolder "internal_report.txt"

    $diskDetails = $physicalDisks | ForEach-Object {
        "$($_.SizeGB) GB $($_.MediaType) - $($_.Model)"
    } | Out-String

$report = @"
INTERNAL TECH REPORT
====================

Manufacturer: $manufacturer
Model: $model
Serial Number: $serialNumber

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

    $report | Out-File $internalFilePath

    return $internalFilePath
}