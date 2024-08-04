# Get the path to the directory where the script is located
$scriptDirectory = $PSScriptRoot
$stuffFolder = Join-Path -Path $scriptDirectory -ChildPath "Stuff"

# Step 1: Set the current date and time to US Central Time
$centralTimeZone = [System.TimeZoneInfo]::FindSystemTimeZoneById("Central Standard Time")
$currentDateTime = [System.TimeZoneInfo]::ConvertTime([DateTime]::Now, $centralTimeZone)
Set-Date -Date $currentDateTime
Write-Host "Step 1/10: Current date and time set to US Central Time."

# Step 2: Change power plan to never sleep
powercfg -change -standby-timeout-ac 0  # Set standby timeout to 0 minutes when plugged in (AC)
powercfg -change -standby-timeout-dc 0  # Set standby timeout to 0 minutes when on battery (DC)
Write-Host "Step 2/10: Power plan has been set to never sleep."

# Step 3: Move a picture to C:\Windows\Web\Wallpaper
$sourceImagePath = Join-Path -Path $stuffFolder -ChildPath "MSU\Background1.jpg"  # Updated path to the image
$destinationFolder = "C:\Windows\Web\Wallpaper"  # Destination folder
if (Test-Path $sourceImagePath) {
    Move-Item -Path $sourceImagePath -Destination $destinationFolder -Force
    Write-Host "Step 3/10: Moved Background1.jpg to C:\Windows\Web\Wallpaper."
} else {
    Write-Host "Step 3/10: Source image not found."
}

# Step 4: Change the desktop background to the moved picture for all users
$backgroundImagePath = "C:\Windows\Web\Wallpaper\Background1.jpg"  # Path to the moved image
if (Test-Path $backgroundImagePath) {
    Set-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\PersonalizationCSP' -Name "DesktopImagePath" -Value $backgroundImagePath
    RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters
    Write-Host "Step 4/10: Desktop background has been changed to Background1.jpg for all users."
} else {
    Write-Host "Step 4/10: Background image not found."
}

# Step 5: Disable power management for the network adapter
$networkAdapters = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object { $_.NetEnabled -eq $true }
foreach ($adapter in $networkAdapters) {
    $adapter.SetPowerManagement(0)  # Disable power management
    Write-Host "Step 5/10: Disabled power management for network adapter: $($adapter.Name)."
}

# Step 6: Disable remote connections
$remoteRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server"
Set-ItemProperty -Path $remoteRegPath -Name "fDenyTSConnections" -Value 1  # Disable remote connections
Set-ItemProperty -Path $remoteRegPath -Name "UserAuthentication" -Value 0  # Disable network level authentication
Write-Host "Step 6/10: Remote connections and network level authentication have been disabled."

# Step 7: Change CD drive letter to T:
$cdDrive = Get-WmiObject -Class Win32_CDROMDrive | Where-Object { $_.Drive -ne $null }
if ($cdDrive) {
    $newDriveLetter = "T:"
    $currentDriveLetter = $cdDrive.Drive
    $result = Invoke-WmiMethod -Class Win32_LogicalDisk -Name ChangeDriveLetter -ArgumentList $currentDriveLetter, $newDriveLetter
    if ($result.ReturnValue -eq 0) {
        Write-Host "Step 7/10: CD drive letter changed from $currentDriveLetter to $newDriveLetter."
    } else {
        Write-Host "Step 7/10: Failed to change CD drive letter."
    }
} else {
    Write-Host "Step 7/10: No CD drive found."
}

# Step 8: Change the product key to an educational one from Key.txt
$keyFilePath = Join-Path -Path $stuffFolder -ChildPath "Key.txt"  # Path to the key file
if (Test-Path $keyFilePath) {
    $educationalProductKey = Get-Content -Path $keyFilePath -Raw  # Read the key from the file
    slmgr.vbs /ipk $educationalProductKey
    Write-Host "Step 8/10: Product key has been changed to the educational key from Key.txt."
} else {
    Write-Host "Step 8/10: Key.txt file not found."
}

# Step 9: Run specific applications from the Stuff folder
$applications = @(
    "Office 2021\install.cmd",  # Path to the install.cmd file in the Office 2021 folder
    "AcroRdr20202000130002_MUI.exe",
    "LanDeskAgent_MSUTexas_with_status.exe",
    "Ninite 7Zip Chrome FileZilla Firefox GIMP PuTTY Installer.exe",
    "SupportAssistLauncher.exe",
    "01-22-24-SophosSetup - Shortcut.lnk"  # Ensure the shortcut has the correct extension
)

$failedApplications = @()  # Array to keep track of failed applications

foreach ($app in $applications) {
    $appPath = Join-Path -Path $stuffFolder -ChildPath $app
    if (Test-Path $appPath) {
        # Attempt to start the application as administrator
        $process = Start-Process -FilePath $appPath -Verb RunAs -PassThru  # Run as administrator and get the process object
        $process.WaitForExit()  # Wait for the process to exit

        # Check the exit code
        if ($process.ExitCode -ne 0) {
            Write-Host "Step 9/10: Application $app failed to start. Exit code: $($process.ExitCode)"
            $failedApplications += $app  # Add to failed applications list
        } else {
            Write-Host "Step 9/10: Started application $app as administrator."
        }
    } else {
        Write-Host "Step 9/10: Application $app not found in $stuffFolder."
    }
}

# Retry failed applications
if ($failedApplications.Count -gt 0) {
    Write-Host "Retrying failed applications..."
    foreach ($failedApp in $failedApplications) {
        $appPath = Join-Path -Path $stuffFolder -ChildPath $failedApp
        if (Test-Path $appPath) {
            # Attempt to start the application again
            $process = Start-Process -FilePath $appPath -Verb RunAs -PassThru  # Run as administrator and get the process object
            $process.WaitForExit()  # Wait for the process to exit

            # Check the exit code again
            if ($process.ExitCode -ne 0) {
                Write-Host "Application $failedApp failed to start again. Exit code: $($process.ExitCode)"
            } else {
                Write-Host "Application $failedApp started successfully on retry."
            }
        } else {
            Write-Host "Application $failedApp not found in $stuffFolder."
        }
    }
}

# Step 10: Check for Windows updates
Write-Host "Step 10/10: Checking for Windows updates..."
Invoke-Expression "usoclient StartScan"  # Start the Windows Update scan
Write-Host "Windows Update scan initiated."

# Final Instructions
Write-Host "All tasks have been completed."
Write-Host "Please follow these steps:"
Write-Host "1. Manually update Adobe."
Write-Host "2. Manually update Excel."
Write-Host "3. Manually check for updates with SupportAssist."
Write-Host "4. Manually check for Windows updates."
Write-Host "5. After all updates are complete, please delete SupportAssist."