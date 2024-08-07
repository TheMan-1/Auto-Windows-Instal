# Updates this function to look for the specific path
function Find-USBDrive {
    $expectedPath = "D:\Run This After\Stuff"
    if (Test-Path $expectedPath) {
        return "D:"
    }
    $drives = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 2 }
    foreach ($drive in $drives) {
        $testPath = Join-Path -Path $drive.DeviceID -ChildPath "Run This After\Stuff"
        if (Test-Path $testPath) {
            return $drive.DeviceID
        }
    }
    throw "USB drive with 'Run This After\Stuff' folder not found."
}

# Updates the USB drive path
try {
    $usbDrive = Find-USBDrive
    $stuffFolder = Join-Path -Path $usbDrive -ChildPath "Run This After\Stuff"
    
    Write-Host "USB drive found at $usbDrive, Stuff folder at $stuffFolder"
    Write-Host ""  # Added line for spacing
    Write-Host ""  # Added line for spacing
}
catch {
    Write-Host "Error: $_"
    exit 1
}

# Initializes step results hashtable
$stepResults = @{}

function Invoke-Step {
    param (
        [int]$StepNumber,
        [string]$Description,
        [scriptblock]$Action
    )

    # Checks if the step has already been completed
    if ($stepResults.ContainsKey($StepNumber) -and $stepResults[$StepNumber] -eq "Completed") {
        Write-Host "Step $StepNumber/8 already completed: $Description"
        return
    }

    Write-Host "Starting Step $StepNumber/8: $Description"
    try {
        $result = & $Action
        if ($result -match "successfully") {
            Write-Host "Step $StepNumber/8 completed successfully: $Description"
            $stepResults[$StepNumber] = "Completed"
        }
        else {
            Write-Host "Step $StepNumber/8 failed: $Description"
            $stepResults[$StepNumber] = "Failed"
        }
    }
    catch {
        Write-Host "Step $StepNumber/8 failed: $Description"
        Write-Host "Error: $_"
        $stepResults[$StepNumber] = "Failed"
    }
    Write-Host ""
}

# Step 1: Set the time zone to US Central Time
Invoke-Step -StepNumber 1 -Description "Set time zone to US Central Time" -Action {
    Set-TimeZone -Id "Central Standard Time"
    Write-Host "Time zone set to US Central Time."
    return "successfully"
}
# Note: Sets the system time zone to Central Time without altering the current time

# Step 2: Change power plan to never sleep
Invoke-Step -StepNumber 2 -Description "Set power plan to never sleep" -Action {
    powercfg /change standby-timeout-ac 0
    powercfg /change standby-timeout-dc 0
    return "successfully"
}
# Note: Disables sleep mode for both AC and battery power

# Step 3: Copy MSU folder to C:\Windows\Web\Wallpaper
Invoke-Step -StepNumber 3 -Description "Copy MSU folder to C:\Windows\Web\Wallpaper" -Action {
    $sourceFolderPath = Join-Path -Path $stuffFolder -ChildPath "MSU"
    $destinationFolder = "C:\Windows\Web\Wallpaper"

    if (Test-Path $sourceFolderPath) {
        Copy-Item -Path $sourceFolderPath -Destination $destinationFolder -Recurse -Force
        Write-Host "Folder copied from $sourceFolderPath to $destinationFolder"
        return "successfully"
    }
    else {
        throw "Source folder not found: $sourceFolderPath"
    }
}
# Note: Copies the entire MSU folder from the USB drive to C:\Windows\Web\Wallpaper\MSU

# Step 4: Change desktop background for all users
Invoke-Step -StepNumber 4 -Description "Change desktop background for all users" -Action {
    $backgroundImagePath = "C:\Windows\Web\Wallpaper\MSU\Background1.jpg"
    
    if (Test-Path $backgroundImagePath) {
        Add-Type -TypeDefinition @"
        using System;
        using System.Runtime.InteropServices;
        public class Wallpaper {
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
        }
"@
        $SPI_SETDESKWALLPAPER = 20
        $SPIF_UPDATEINIFILE = 0x01
        $SPIF_SENDCHANGE = 0x02
        
        [Wallpaper]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $backgroundImagePath, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
        Write-Host "Desktop background changed to $backgroundImagePath"
        return "successfully"
    }
    else {
        throw "Background image not found at $backgroundImagePath."
    }
}
# Note: Sets the desktop background using the Background1.jpg from the copied MSU folder

# Step 5: Disable power management for the network adapter
Invoke-Step -StepNumber 5 -Description "Disable power management for network adapters" -Action {
    $networkAdapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
    foreach ($adapter in $networkAdapters) {
        Disable-NetAdapterPowerManagement -Name $adapter.Name
    }
    return "successfully"
}
# Note: Finds all active network adapters and disables their power management

# Step 6: Enable Remote Desktop and configure settings
Invoke-Step -StepNumber 6 -Description "Enable Remote Desktop and configure settings" -Action {
    Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name "fDenyTSConnections" -Value 0
    Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' -Name "UserAuthentication" -Value 0
    Enable-NetFirewallRule -DisplayGroup "Remote Desktop"
    
    Write-Host "Remote Desktop has been enabled and NLA has been disabled."
    return "successfully"
}
# Note: Enables Remote Desktop and disables Network Level Authentication

# Step 7: Change CD drive letter to T:
Invoke-Step -StepNumber 7 -Description "Change CD drive letter to T:" -Action {
    $cdDrive = Get-WmiObject -Class Win32_Volume -Filter "DriveType=5"
    if ($cdDrive) {
        $newDriveLetter = "T:"
        $cdDrive.DriveLetter = $newDriveLetter
        $cdDrive.Put()
        return "successfully"
    }
    else {
        throw "No CD drive found."
    }
}
# Note: Identifies CD drive and changes its drive letter to T:

# Step 8: Run other applications
Invoke-Step -StepNumber 8 -Description "Run other applications" -Action {
    $applicationsToRun = @(
        "D:\Run This After\Stuff\AcroRdr20202000130002_MUI.exe", 
        "D:\Run This After\Stuff\Ninite 7Zip Chrome FileZilla Firefox GIMP PuTTY Installer.exe",
        "D:\Run This After\Stuff\LanDeskAgent_MSUTexas_with_status.exe", 
        "D:\Run This After\Stuff\01-22-24-SophosSetup - Shortcut.lnk")
    foreach ($app in $applicationsToRun) {
        if (Test-Path $app) {
            # Check if the file is a valid executable or a .cmd file
            if ((Get-Item $app).Extension -in @('.exe', '.cmd', '.bat', '.lnk')) {
                if ($app.EndsWith('.lnk')) {
                    # Resolve the shortcut target
                    $shell = New-Object -ComObject WScript.Shell
                    $shortcut = $shell.CreateShortcut($app)
                    $targetPath = $shortcut.TargetPath
                    if (Test-Path $targetPath) {
                        Start-Process -FilePath $targetPath -NoNewWindow
                    }
                    else {
                        Write-Host "Warning: Target of shortcut $app not found."
                    }
                }
                else {
                    # Run the application directly
                    Start-Process -FilePath $app -NoNewWindow
                }
            }
            else {
                Write-Host "Warning: $app is not a valid executable."
            }
        }
        else {
            throw "Application not found at $app."
        }
    }

    # Run the install.cmd from Office 2021 in a separate command prompt
    $installCmdPath = "D:\Run This After\Stuff\Office 2021\install.cmd"
    if (Test-Path $installCmdPath) {
        $tempBatchFile = "$env:TEMP\RunInstall.cmd"
        Set-Content -Path $tempBatchFile -Value "@echo off`ncd `"$($installCmdPath | Split-Path -Parent)`ncall `"$installCmdPath`""
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c `"$tempBatchFile`"" -NoNewWindow
    }
    else {
        Write-Host "Warning: install.cmd not found at $installCmdPath."
    }

    return "successfully"
}

# Summary of steps
Write-Host ""
Write-Host "Summary of Steps:"
foreach ($step in 1..8) {
    Write-Host "Step ${step}: $($stepResults[$step])"
}

# Final Instructions
Write-Host ""
Write-Host "All tasks have been attempted."
Write-Host "Please follow these steps:"
Write-Host "1. Manually install SupportAssist from the appropriate source."
Write-Host "2. Manually update Adobe and Excel."
Write-Host "3. Manually check for updates with SupportAssist."
Write-Host "4. Manually check for Windows updates."
Write-Host "5. After all updates are complete, please delete SupportAssist."