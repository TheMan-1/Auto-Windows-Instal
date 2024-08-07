# Get the path to the directory where the script is located
$scriptDirectory = $PSScriptRoot
$stuffFolder = Join-Path -Path $scriptDirectory -ChildPath "Stuff"

$stepResults = @{}

function Invoke-Step {
    param (
        [int]$StepNumber,
        [string]$Description,
        [scriptblock]$Action
    )

    Write-Host "Starting Step $StepNumber/10: $Description"
    try {
        & $Action
        Write-Host "Step $StepNumber/10 completed successfully: $Description"
        $stepResults[$StepNumber] = "Completed"
    }
    catch {
        Write-Host "Step $StepNumber/10 failed: $Description"
        Write-Host "Error: $_"
        $stepResults[$StepNumber] = "Failed"
    }
    Write-Host ""
}

# Step 1: Set the current date and time to US Central Time
Invoke-Step -StepNumber 1 -Description "Set current date and time to US Central Time" -Action {
    $centralTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Central Standard Time")
    $currentDateTime = [System.TimeZoneInfo]::ConvertTime([DateTime]::Now, $centralTimeZone)
    Set-Date -Date $currentDateTime
}
# Note: Finds Central Time Zone, converts current time, sets system time

# Step 2: Change power plan to never sleep
Invoke-Step -StepNumber 2 -Description "Set power plan to never sleep" -Action {
    powercfg /change standby-timeout-ac 0
    powercfg /change standby-timeout-dc 0
}
# Note: Disables sleep mode for both AC and battery power

# Step 3: Move a picture to C:\Windows\Web\Wallpaper
Invoke-Step -StepNumber 3 -Description "Move Background1.jpg to C:\Windows\Web\Wallpaper" -Action {
    $sourceImagePath = Join-Path -Path $stuffFolder -ChildPath "MSU\Background1.jpg"
    $destinationFolder = "C:\Windows\Web\Wallpaper"
    if (Test-Path $sourceImagePath) {
        Move-Item -Path $sourceImagePath -Destination $destinationFolder -Force
    }
    else {
        throw "Source image not found."
    }
}
# Note: Moves "Background1.jpg" from script folder to Windows wallpaper directory

# Step 4: Change the desktop background to the moved picture for all users
Invoke-Step -StepNumber 4 -Description "Change desktop background for all users" -Action {
    $backgroundImagePath = "C:\Windows\Web\Wallpaper\Background1.jpg"
    if (Test-Path $backgroundImagePath) {
        Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP' -Name "DesktopImagePath" -Value $backgroundImagePath
        RUNDLL32.EXE user32.dll, UpdatePerUserSystemParameters
    }
    else {
        throw "Background image not found."
    }
}
# Note: Sets the moved image as desktop background for all users

# Step 5: Disable power management for the network adapter
Invoke-Step -StepNumber 5 -Description "Disable power management for network adapters" -Action {
    $networkAdapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
    foreach ($adapter in $networkAdapters) {
        Disable-NetAdapterPowerManagement -Name $adapter.Name
    }
}
# Note: Finds all active network adapters and disables their power management

# Step 6: Enable Remote Desktop and configure settings
Invoke-Step -StepNumber 6 -Description "Enable Remote Desktop and configure settings" -Action {
    # Enable Remote Desktop
    Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name "fDenyTSConnections" -Value 0
    
    # Disable NLA (Network Level Authentication)
    Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name "UserAuthentication" -Value 0
    
    # Enable Remote Desktop firewall rule
    Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
    
    Write-Host "Remote Desktop has been enabled and NLA has been disabled."
}
# Note: Enables Remote Desktop and disables Network Level Authentication

# Step 7: Change CD drive letter to T:
Invoke-Step -StepNumber 7 -Description "Change CD drive letter to T:" -Action {
    $cdDrive = Get-WmiObject -Class Win32_Volume -Filter "DriveType=5"
    if ($cdDrive) {
        $newDriveLetter = "T:"
        $cdDrive.DriveLetter = $newDriveLetter
        $cdDrive.Put()
    }
    else {
        throw "No CD drive found."
    }
}
# Note: Identifies CD drive and changes its drive letter to T:

# Step 8: Change the product key to an educational one from Key.txt
Invoke-Step -StepNumber 8 -Description "Change product key to educational key" -Action {
    $keyFilePath = Join-Path -Path $stuffFolder -ChildPath "Key.txt"
    if (Test-Path $keyFilePath) {
        $educationalProductKey = Get-Content -Path $keyFilePath -Raw
        $result = cscript //nologo C:\Windows\System32\slmgr.vbs /ipk $educationalProductKey
        if ($result -match "successfully") {
            Write-Host $result
        }
        else {
            throw "Failed to change product key: $result"
        }
    }
    else {
        throw "Key.txt file not found."
    }
}
# Note: Reads key from "Key.txt" and applies it using Windows activation script

# Step 9: Run specific applications from the Stuff folder
Invoke-Step -StepNumber 9 -Description "Run specific applications" -Action {
    $applications = @(
        "Office 2021\install.cmd",
        "AcroRdr20202000130002_MUI.exe",
        "LanDeskAgent_MSUTexas_with_status.exe",
        "Ninite 7Zip Chrome FileZilla Firefox GIMP PuTTY Installer.exe",
        "SupportAssistLauncher.exe"
    )

    foreach ($app in $applications) {
        $appPath = Join-Path -Path $stuffFolder -ChildPath $app
        if (Test-Path $appPath) {
            $process = Start-Process -FilePath $appPath -Verb RunAs -PassThru
            $process.WaitForExit()
            if ($process.ExitCode -ne 0) {
                Write-Host "Warning: Application $app exited with code $($process.ExitCode)"
            }
        }
        else {
            Write-Host "Warning: Application $app not found in $stuffFolder."
        }
    }
}
# Note: Executes a list of installers and setup files

# Step 10: Install Sophos
Invoke-Step -StepNumber 10 -Description "Install Sophos" -Action {
    $sophosPath = Join-Path -Path $stuffFolder -ChildPath "01-22-24-SophosSetup - Shortcut.lnk"
    if (Test-Path $sophosPath) {
        $process = Start-Process -FilePath $sophosPath -Verb RunAs -PassThru
        $process.WaitForExit()
        if ($process.ExitCode -ne 0) {
            Write-Host "Warning: Sophos installation exited with code $($process.ExitCode)"
        }
    }
    else {
        Write-Host "Warning: Sophos installer not found at $sophosPath."
    }
}

# Summary of results
Write-Host "Summary of Steps:"
foreach ($step in 1..10) {
    $status = $stepResults[$step]
    $description = switch ($step) {
        1 { "Set current date and time to US Central Time" }
        2 { "Set power plan to never sleep" }
        3 { "Move Background1.jpg to C:\Windows\Web\Wallpaper" }
        4 { "Change desktop background for all users" }
        5 { "Disable power management for network adapters" }
        6 { "Enable Remote Desktop and configure settings" }
        7 { "Change CD drive letter to T:" }
        8 { "Change product key to educational key" }
        9 { "Run specific applications" }
        10 { "Install Sophos" }
    }
    Write-Host "Step $step`: $status - $description"
}

# Final Instructions
Write-Host ""
Write-Host "All tasks have been attempted."
Write-Host "Please follow these steps:"
Write-Host "1. Manually update Adobe."
Write-Host "2. Manually update Excel."
Write-Host "3. Manually check for updates with SupportAssist."
Write-Host "4. Manually check for Windows updates."
Write-Host "5. After all updates are complete, please delete SupportAssist."