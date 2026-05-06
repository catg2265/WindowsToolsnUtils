# ----------------------
# Disk Detection
# ----------------------
function Get-Disks {
    $disks = Get-PhysicalDisk | Where-Object { $_.BusType -ne 'USB' -and $_.FriendlyName -ne 'Verbatim STORE N GO' }
    $diskString = ""
    $diskRows = @()
    $diskObjects = @()

    $validSizes = @(
        64, 120, 128, 248, 256, 320, 480,
        500, 512, 960, 1000, 1024, 2000,
        3000, 4000, 6000, 8000, 10000,
        12000, 16000, 20000, 22000
    )

    foreach ($disk in $disks) {
        $diskSizeGB = [math]::Round($disk.Size / 1GB)
        $nearestSizeUp = $validSizes | Where-Object { $_ -ge $diskSizeGB } | Sort-Object | Select-Object -First 1

        $type = $null
        $newLine = ""

        if ($disk.MediaType -eq 'SSD' -or $disk.SpindleSpeed -eq 0) {
            $type = "SSD"
            $newLine = "<tr><td>SSD</td><td>$nearestSizeUp GB</td></tr>"
        }
        elseif ($disk.MediaType -eq 'HDD' -or $disk.SpindleSpeed -gt 0) {
            $type = "HDD"
            $newLine = "<tr><td>Harddisk</td><td>$nearestSizeUp GB</td></tr>"
        }
        else {
            Write-Verbose "Unknown disk type: $($disk.FriendlyName) $diskSizeGB GB"
            continue
        }

        # Build HTML string
        if ($newLine) {
            $diskRows += $newLine
        }

        # Add structured object
        $diskObjects += [PSCustomObject]@{
            SizeGB    = $nearestSizeUp
            MediaType = $type
            Model     = $disk.FriendlyName
        }
    }
    $diskString = $diskRows -join "`n"

    return @{
        Html  = $diskString
        Disks = $diskObjects
    }
}

# ----------------------
# Screen Size Detection
# ----------------------
function Get-ScreenSize {
    $monitors = Get-CimInstance -Namespace root\wmi -ClassName WmiMonitorBasicDisplayParams

    if (-not $monitors) {
        return "Unknown"
    }

    # Take first monitor only
    $m = $monitors | Select-Object -First 1

    $width  = $m.MaxHorizontalImageSize
    $height = $m.MaxVerticalImageSize

    if ($width -le 0 -or $height -le 0) {
        return "Unknown"
    }

    # Calculate diagonal in inches
    $diagonal = [math]::Sqrt(($width * $width) + ($height * $height))
    $inchesRaw = $diagonal / 2.54

    # Common marketed screen sizes
    $validSizes = @(
        11.6, 12, 12.3, 13, 13.3, 14, 15, 15.6, 
        16, 16.1, 17, 17.3, 18
    )

    # Find nearest size
    $nearest = $validSizes | Sort-Object {
        [math]::Abs($_ - $inchesRaw)
    } | Select-Object -First 1

    # Returns the correct screen size
    return "$nearest`""
}

# ----------------------
# Core System Info
# ----------------------
function Get-SystemCoreInfo {
    $computerSystem = Get-CimInstance Win32_ComputerSystem
    $cpu  = Get-CimInstance Win32_Processor
    $gpu  = Get-CimInstance Win32_VideoController | Select-Object -First 1
    $bios = Get-CimInstance Win32_BIOS
    $os   = Get-CimInstance Win32_OperatingSystem

    return [PSCustomObject]@{
        Manufacturer = $computerSystem.Manufacturer
        Model        = $computerSystem.Model
        CPU          = $cpu.Name
        RAM_GB       = [math]::Round($computerSystem.TotalPhysicalMemory / 1GB)
        GPU          = $gpu.Name
        BIOSVersion  = $bios.SMBIOSBIOSVersion
        OS           = $os.Caption
    }
}

# ----------------------
# Full Aggregated Info
# ----------------------
function Get-SystemInfo {
    $diskInfo = Get-Disks
    $coreInfo = Get-SystemCoreInfo
    $screen   = Get-ScreenSize

    return [PSCustomObject]@{
        Manufacturer = $coreInfo.Manufacturer
        Model        = $coreInfo.Model
        CPU          = $coreInfo.CPU
        RAM_GB       = $coreInfo.RAM_GB
        GPU          = $coreInfo.GPU
        BIOSVersion  = $coreInfo.BIOSVersion
        OS           = $coreInfo.OS
        ScreenSize   = $screen
        DiskHtml     = $diskInfo.Html
        Disks        = $diskInfo.Disks
    }
}