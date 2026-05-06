. .\Progress_Bar_Controller.ps1

Add-Process "Create folder"
Add-Process "Run battery test"
Add-Process "Generate report"

Initialize-ProgressBar -Activity "Windows Automation"

# Step 1
New-ResultsFolder ...
Move-Process

# Step 2 (battery test with sub-progress)
$batteryProgress = New-ProgressContext -ParentId 1 -Id 2
Start-BatteryTest -Progress $batteryProgress
Move-Process

# Step 3
New-InternalReportFile ...
Move-Process

Complete-Progress


. .\SystemInfo.ps1   # dot-source the file

$system = Get-SystemInfo

$system.Model
$system.CPU
$system.RAM_GB
$system.GPU
$system.BIOSVersion
$system.OS
$system.ScreenSize
$system.DiskHtml
$system.Disks


. .\ResultsModule.ps1    # dot-source the file

$folder = New-ResultsFolder -ScriptRoot $scriptRoot -Model $model

$specFile = New-SpecSheetFile `
    -DestinationFolder $folder `
    -cpuName $cpuName `
    -ram $ram `
    -diskRows $diskRows `
    -gpuName $gpuName `
    -screenSize $screenSize `
    -os $os `
    -batteryExists $batteryExists `
    -batteryResult $batteryResult `
    -estimatedPrice $estimatedPrice
Write-Host "Created: $specFile"

$reportFile = New-InternalReportFile `
    -DestinationFolder $folder `
    -manufacturer $manufacturer `
    -model $model `
    -serialNumber $bios.SerialNumber `
    -cpuName $cpuName `
    -ram $ram `
    -physicalDisks $physicalDisks `
    -gpuName $gpuName `
    -screenSize $screenSize `
    -os $os `
    -batteryResult $batteryResult `
    -estimatedPrice $estimatedPrice
Write-Host "Created: $reportFile"