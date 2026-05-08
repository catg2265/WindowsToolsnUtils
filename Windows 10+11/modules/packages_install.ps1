
function Initialize-PackageEnvironment {

    try {
        [Net.ServicePointManager]::SecurityProtocol = `
            [Net.SecurityProtocolType]::Tls12
    }
    catch {
        Write-Warning "Failed to set TLS 1.2 (may already be enforced by OS)"
    }
}

function Install-Winget {

    Write-Host "Checking WinGet availability..."

    $winget = Get-Command winget -ErrorAction SilentlyContinue

    if ($winget) {
        Write-Host "WinGet already installed. Attempting repair..."

        try {
            if (-not (Get-Module -ListAvailable Microsoft.WinGet.Client)) {
		        # Trust PSGallery to avoid prompts
	            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
                
		        # Ensure NuGet Provider is installed
            	if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
                    Install-PackageProvider -Name NuGet -Force -ForceBootstrap -Scope AllUsers
	    	    }

		        Install-Module Microsoft.WinGet.Client -Force -Scope AllUsers
            }

            Import-Module Microsoft.WinGet.Client

            if (Get-Command Repair-WinGetPackageManager -ErrorAction SilentlyContinue) {
                Repair-WinGetPackageManager
            }

            Write-Host "WinGet repair completed."
            return
        }
        catch {
            try {
                Repair-WinGetPackageManager -ErrorAction Stop
                Write-Host "WinGet repair completed."
            }
            catch {
                Write-Warning "Repair failed. Checking if WinGet is still usable..."
            }
            
            # Validate winget installation
            $wingetWorks = $false
            try {
                $null = winget --version
                $wingetWorks = $true
            }
            catch {
                $wingetWorks = $false
            }
            
            if ($wingetWorks) {
                Write-Host "WinGet is functional. Skipping reinstall."
                return
            }
            
            Write-Warning "WinGet is not functional. Proceeding with reinstall."
        }
    }

    Write-Host "WinGet not found or repair failed. Installing App Installer..."

    # --- Install winget MSIX ---
    $bundleUrl  = "https://aka.ms/getwinget"
    $bundlePath = "$env:TEMP\Microsoft.DesktopAppInstaller.msixbundle"

    Invoke-WebRequest -Uri $bundleUrl -OutFile $bundlePath -ErrorAction Stop

    if (!(Test-Path $bundlePath)) {
        throw "Failed to download App Installer package"
    }

    # Install winget
    Add-AppxPackage -Path $bundlePath -ForceApplicationShutdown -ErrorAction Stop

    Start-Sleep -Seconds 3

    # Validate installation
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) {
        throw "WinGet installation failed"
    }

    Write-Host "WinGet installed successfully."
}

function Install-Chocolatey {
    if (Get-Command choco -ErrorAction SilentlyContinue) {
        Write-Host "Chocolatey already installed."
        return
    }

    Write-Host "Installing Chocolatey..."

    Set-ExecutionPolicy Bypass -Scope Process -Force

    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))

    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path","User")
}
function Test-ChocoPackageInstalled {
    param([string]$Name)

    $result = choco list --local-only --exact $Name 2>$null

    return ($result -match $Name)
}
function Test-PackageInstalled {
    param(
        [string]$ChocoId
    )

    if ($ChocoId) {
        if (Test-ChocoPackageInstalled -Name $ChocoId) {
            return $true
        }
    }

    return $false
}

function Install-WithWinget {
    param($id)

    Write-Host "Trying winget install: $id"

    $arguments = @(
        "install",
        "--id", $id,
        "--silent",
        "--accept-package-agreements",
        "--accept-source-agreements",
        "--exact",
        "--scope", "machine"
    )

    $p = Start-Process winget -ArgumentList $arguments -NoNewWindow -Wait -PassThru
    $code = $p.ExitCode

    if ($code -ne 0) {
        Write-Host "Install failed, trying upgrade..."

        $arguments = @(
            "upgrade",
            "--id", $id,
            "--silent",
            "--accept-package-agreements",
            "--accept-source-agreements",
            "--exact",
            "--scope", "machine"
        )

        $p2 = Start-Process winget -ArgumentList $arguments -NoNewWindow -Wait -PassThru
        $code = $p2.ExitCode
    }

    return $code
}

function Install-WithChoco {
    param($name)

    Write-Host "Falling back to Chocolatey: $name"

    choco install $name -y
}

function Install-PackageManagers{
    Install-Winget
    Install-Chocolatey
}
function Install-Packages{
    param(
        [string]$JsonPath,
        [Object]$Progress
    )

    # Test if JSON file exists
    if (!(Test-Path $JsonPath)) {
        Write-Error "JSON file not found: $JsonPath"
        exit 1
    }

    # Import JSON
    $data = Get-Content $JsonPath | ConvertFrom-Json

    $total = $data.packages.Count
    $current = 0

    # Install loop
    foreach ($pkg in $data.packages) {
        $current++

        $percent = [int](($current / $total) * 100)

        $Progress.Update.Invoke(
            "Installing packages",
            "$($pkg.name) ($current of $total)",
            $percent
        ) 

        Write-Host "`nInstalling $($pkg.name)..."

        $wingetReturnCode = Install-WithWinget -id $pkg.winget
        Write-Debug "`nwinget install of $($pkg.name) gave exit code: $($wingetReturnCode)"

        $installed = winget list --id $pkg.winget --exact 2>$null

        if (-not $installed) {
            if (Test-PackageInstalled -ChocoId $pkg.choco) {
                Write-Host "Skipping $($pkg.name) (already installed)"
                continue
            }
            Install-WithChoco -name $pkg.choco
        } else {
            Write-Host "$($pkg.name) installed via winget."
        }
    }
    # Finish progress bar
    $Progress.Complete.Invoke()
}
function Set-DefaultApps {
    param([string]$XmlPath)

    if (!(Test-Path $XmlPath)) {
        Write-Warning "Default app XML not found: $XmlPath"
        return
    }

    Write-Host "Applying default app associations..."

    Start-Process dism -ArgumentList "/Online /Import-DefaultAppAssociations:`"$XmlPath`"" -Wait -NoNewWindow

    Write-Host "***FOR FUTURE USERS ONLY*** Default apps import completed. ***FOR FUTURE USERS ONLY***"
}


