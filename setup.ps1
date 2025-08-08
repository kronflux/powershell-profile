# Ensure the script can run with elevated privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Please run this script as an Administrator!"
    break
}

# Function to test internet connectivity
function Test-InternetConnection {
    try {
        Test-Connection -ComputerName www.google.com -Count 1 -ErrorAction Stop | Out-Null
        return $true
    }
    catch {
        Write-Warning "Internet connection is required but not available. Please check your connection."
        return $false
    }
}

# Function to install Nerd Fonts
function Install-NerdFonts {
    param (
        [string]$FontName = "Mononoki",
        [string]$FontDisplayName = "Mononoki NF",
        [string]$Version = "3.2.1"
    )

    try {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
        $fontFamilies = (New-Object System.Drawing.Text.InstalledFontCollection).Families.Name
        
        if ($fontFamilies -notcontains "${FontDisplayName}") {
            Write-Host "Installing ${FontDisplayName} font..." -ForegroundColor Cyan
            $fontZipUrl = "https://github.com/ryanoasis/nerd-fonts/releases/download/v${Version}/${FontName}.zip"
            $zipFilePath = "$env:TEMP\${FontName}.zip"
            $extractPath = "$env:TEMP\${FontName}"

            $webClient = New-Object System.Net.WebClient
            $webClient.DownloadFileAsync((New-Object System.Uri($fontZipUrl)), $zipFilePath)

            while ($webClient.IsBusy) {
                Start-Sleep -Seconds 2
            }

            Expand-Archive -Path $zipFilePath -DestinationPath $extractPath -Force
            $destination = (New-Object -ComObject Shell.Application).Namespace(0x14)
            Get-ChildItem -Path $extractPath -Recurse -Filter "*.ttf" | ForEach-Object {
                If (-not(Test-Path "C:\Windows\Fonts\$($_.Name)")) {
                    $destination.CopyHere($_.FullName, 0x10)
                }
            }

            Remove-Item -Path $extractPath -Recurse -Force
            Remove-Item -Path $zipFilePath -Force
            Write-Host "Font ${FontDisplayName} installed successfully" -ForegroundColor Green
        } else {
            Write-Host "Font ${FontDisplayName} already installed" -ForegroundColor DarkGray
        }
        
        return $fontFamilies -contains "${FontDisplayName}"
    }
    catch {
        Write-Error "Failed to download or install ${FontDisplayName} font. Error: $_"
        return $false
    }
}

# Function to ensure Starship is installed and configured
function Ensure-StarshipConfigured {
    param(
        [switch]$ForcePreset
    )
    
    try {
        # Attempt to locate starship.exe even if PATH hasn't refreshed
        $starshipExe = Join-Path $env:ProgramFiles 'starship\bin\starship.exe'
        if (-not (Test-Path $starshipExe)) {
            $cmd = Get-Command starship -ErrorAction SilentlyContinue
            if ($cmd) { 
                $starshipExe = $cmd.Source 
            } else { 
                $starshipExe = $null 
            }
        }

        if (-not $starshipExe) {
            Write-Host "[setup] Installing Starship..." -ForegroundColor Cyan
            $winget = Get-Command winget -ErrorAction SilentlyContinue
            if (-not $winget) { 
                throw "winget not found. Install winget or install Starship manually." 
            }
            
            # Install Starship using winget
            winget install -e --accept-source-agreements --accept-package-agreements Starship.Starship

            # Try to resolve again after install
            $starshipExe = Join-Path $env:ProgramFiles 'starship\bin\starship.exe'
            if (-not (Test-Path $starshipExe)) {
                $cmd = Get-Command starship -ErrorAction SilentlyContinue
                if ($cmd) { 
                    $starshipExe = $cmd.Source 
                } else { 
                    $starshipExe = $null 
                }
            }
        }

        if (-not $starshipExe) {
            Write-Warning "[setup] Starship installed but not yet on PATH. Close and reopen PowerShell, then re-run this script."
            return $false
        }

        # Ensure config directory exists
        $cfgDir = Join-Path $HOME '.config'
        if (-not (Test-Path $cfgDir)) { 
            New-Item -ItemType Directory -Path $cfgDir -Force | Out-Null 
        }
        $cfg = Join-Path $cfgDir 'starship.toml'

        # Only write preset if file doesn't exist, or if forced
        if ($ForcePreset -or -not (Test-Path $cfg)) {
            & $starshipExe preset bracketed-segments -o $cfg | Out-Null
            Write-Host "[setup] Wrote Starship preset to $cfg" -ForegroundColor Green
        } else {
            Write-Host "[setup] Starship config already exists at $cfg; leaving as-is." -ForegroundColor DarkGray
        }

        # Set the STARSHIP_CONFIG environment variable
        $env:STARSHIP_CONFIG = $cfg
        [System.Environment]::SetEnvironmentVariable("STARSHIP_CONFIG", $cfg, [System.EnvironmentVariableTarget]::User)
        Write-Host "[setup] Environment variable STARSHIP_CONFIG set to: $cfg" -ForegroundColor Green

        # Ensure profile initialization (only append once)
        $profilePath = $PROFILE
        if ($profilePath -and (Test-Path $profilePath)) {
            $profileText = Get-Content -Raw -Path $profilePath
            $initLine = 'Invoke-Expression (&starship init powershell)'
            if ($profileText -notmatch [regex]::Escape($initLine)) {
                Add-Content -Path $profilePath -Value "`n# Initialize Starship prompt`n$initLine`n"
                Write-Host "[setup] Appended Starship init to $profilePath" -ForegroundColor Green
            } else {
                Write-Host "[setup] Starship init already present in $profilePath" -ForegroundColor DarkGray
            }
        }
        
        return $true
    } 
    catch {
        Write-Error "[setup] Failed to configure Starship. Error: $_"
        return $false
    }
}

# Main script execution starts here

# Check for internet connectivity before proceeding
if (-not (Test-InternetConnection)) {
    Write-Error "Internet connection required. Exiting."
    break
}

# Profile creation or update
Write-Host "`n=== Setting up PowerShell Profile ===" -ForegroundColor Cyan

if (!(Test-Path -Path $PROFILE -PathType Leaf)) {
    try {
        # Detect Version of PowerShell & Create Profile directories if they do not exist.
        $profilePath = ""
        if ($PSVersionTable.PSEdition -eq "Core") {
            $profilePath = "$env:userprofile\Documents\PowerShell"
        }
        elseif ($PSVersionTable.PSEdition -eq "Desktop") {
            $profilePath = "$env:userprofile\Documents\WindowsPowerShell"
        }

        if (!(Test-Path -Path $profilePath)) {
            New-Item -Path $profilePath -ItemType "directory" | Out-Null
        }

        Invoke-RestMethod https://github.com/kronflux/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        Write-Host "✅ The profile @ [$PROFILE] has been created." -ForegroundColor Green
        Write-Host "ℹ️  If you want to make any personal changes or customizations, please do so at [$profilePath\Profile.ps1]" -ForegroundColor Yellow
        Write-Host "    as there is an updater in the installed profile which uses the hash to update the profile" -ForegroundColor Yellow
        Write-Host "    and will lead to loss of changes if modified directly." -ForegroundColor Yellow
    }
    catch {
        Write-Error "Failed to create or update the profile. Error: $_"
    }
}
else {
    try {
        $backupPath = Join-Path (Split-Path $PROFILE) "oldprofile.ps1"
        Move-Item -Path $PROFILE -Destination $backupPath -Force
        Invoke-RestMethod https://github.com/kronflux/powershell-profile/raw/main/Microsoft.PowerShell_profile.ps1 -OutFile $PROFILE
        Write-Host "✅ PowerShell profile at [$PROFILE] has been updated." -ForegroundColor Green
        Write-Host "📦 Your old profile has been backed up to [$backupPath]" -ForegroundColor Yellow
        Write-Host "⚠️  NOTE: Please back up any persistent components of your old profile to" -ForegroundColor Yellow
        Write-Host "    [$HOME\Documents\PowerShell\Profile.ps1]" -ForegroundColor Yellow
        Write-Host "    as there is an updater in the installed profile which uses the hash to update" -ForegroundColor Yellow
        Write-Host "    the profile and will lead to loss of changes." -ForegroundColor Yellow
    }
    catch {
        Write-Error "❌ Failed to backup and update the profile. Error: $_"
    }
}

# Install and configure Starship
Write-Host "`n=== Installing and Configuring Starship ===" -ForegroundColor Cyan
$starshipSuccess = Ensure-StarshipConfigured

# Install Nerd Font
Write-Host "`n=== Installing Nerd Fonts ===" -ForegroundColor Cyan
$fontSuccess = Install-NerdFonts -FontName "Mononoki" -FontDisplayName "Mononoki NF"

# Install Chocolatey
Write-Host "`n=== Installing Chocolatey Package Manager ===" -ForegroundColor Cyan
try {
    $chocoInstalled = Get-Command choco -ErrorAction SilentlyContinue
    if (-not $chocoInstalled) {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
        Write-Host "✅ Chocolatey installed successfully" -ForegroundColor Green
    } else {
        Write-Host "Chocolatey already installed" -ForegroundColor DarkGray
    }
}
catch {
    Write-Error "Failed to install Chocolatey. Error: $_"
}

# Install Terminal Icons Module
Write-Host "`n=== Installing Terminal Icons Module ===" -ForegroundColor Cyan
try {
    if (-not (Get-Module -ListAvailable -Name Terminal-Icons)) {
        Install-Module -Name Terminal-Icons -Repository PSGallery -Force -Scope CurrentUser
        Write-Host "✅ Terminal Icons module installed successfully" -ForegroundColor Green
    } else {
        Write-Host "Terminal Icons module already installed" -ForegroundColor DarkGray
    }
}
catch {
    Write-Error "Failed to install Terminal Icons module. Error: $_"
}

# Final status check
Write-Host "`n=== Setup Summary ===" -ForegroundColor Cyan

$profileExists = Test-Path -Path $PROFILE
$starshipInstalled = Get-Command starship -ErrorAction SilentlyContinue
$chocoInstalled = Get-Command choco -ErrorAction SilentlyContinue
$terminalIconsInstalled = Get-Module -ListAvailable -Name Terminal-Icons

if ($profileExists) {
    Write-Host "✅ PowerShell Profile: Configured" -ForegroundColor Green
} else {
    Write-Host "❌ PowerShell Profile: Not configured" -ForegroundColor Red
}

if ($starshipInstalled) {
    Write-Host "✅ Starship: Installed" -ForegroundColor Green
} else {
    Write-Host "⚠️  Starship: Installed but PATH not updated (restart required)" -ForegroundColor Yellow
}

if ($fontSuccess) {
    Write-Host "✅ Nerd Font (Mononoki NF): Installed" -ForegroundColor Green
} else {
    Write-Host "⚠️  Nerd Font: Installation may have failed" -ForegroundColor Yellow
}

if ($chocoInstalled) {
    Write-Host "✅ Chocolatey: Installed" -ForegroundColor Green
} else {
    Write-Host "❌ Chocolatey: Not installed" -ForegroundColor Red
}

if ($terminalIconsInstalled) {
    Write-Host "✅ Terminal Icons: Installed" -ForegroundColor Green
} else {
    Write-Host "❌ Terminal Icons: Not installed" -ForegroundColor Red
}

# Final message
Write-Host "`n" -NoNewline
if ($profileExists -and ($starshipInstalled -or $starshipSuccess) -and $fontSuccess) {
    Write-Host "✅ Setup completed successfully!" -ForegroundColor Green
    Write-Host "Please restart your PowerShell session to apply all changes." -ForegroundColor Cyan
    Write-Host "Don't forget to set your terminal to use 'Mononoki NF' font for proper icon display." -ForegroundColor Yellow
} else {
    Write-Warning "Setup completed with some issues. Please check the messages above and restart PowerShell."
    Write-Host "You may need to run this script again after restarting if some components didn't install properly." -ForegroundColor Yellow
}