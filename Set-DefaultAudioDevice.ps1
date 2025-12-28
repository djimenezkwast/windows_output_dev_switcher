<#
.SYNOPSIS
    Easy audio output device switcher for Windows.
#>

# Set window title
$Host.UI.RawUI.WindowTitle = "Audio Output Switcher"

# Constants
$Script:Timeouts = @{
    AlreadySelected = 2000
    Success = 3000
    Goodbye = 500
}

function Get-DeviceState {
    $devices = @(Get-AudioDevice -List | Where-Object { $_.Type -eq "Playback" })
    $currentDefault = Get-AudioDevice -Playback
    $currentComms = Get-AudioDevice -PlaybackCommunication

    if ($null -eq $currentDefault -and $devices.Count -gt 0) {
        $currentDefault = $devices[0]
    }
    if ($null -eq $currentComms) {
        $currentComms = $currentDefault
    }

    return @{
        Devices = $devices
        Default = $currentDefault
        Comms = $currentComms
    }
}

function Install-AudioModule {
    if (-not (Get-Module -ListAvailable -Name AudioDeviceCmdlets)) {
        Write-Host ""
        Write-Host "  First-time setup: Installing required component..." -ForegroundColor Yellow
        Write-Host ""
        try {
            # Required for PSGallery on older Windows versions
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

            Install-Module -Name AudioDeviceCmdlets -Repository PSGallery -Force -Scope CurrentUser
            Write-Host "  Setup complete!" -ForegroundColor Green
            Write-Host ""
        }
        catch {
            Write-Host ""
            Write-Host "  ERROR: Could not install required component." -ForegroundColor Red
            Write-Host ""
            Write-Host "  Please try running PowerShell as Administrator and run this script again." -ForegroundColor Yellow
            Write-Host ""
            Write-Host "  Press any key to exit..." -ForegroundColor Gray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            exit 1
        }
    }

    try {
        Import-Module AudioDeviceCmdlets -ErrorAction Stop
    }
    catch {
        Write-Host ""
        Write-Host "  ERROR: Could not load audio module." -ForegroundColor Red
        Write-Host "  Try restarting the script or reinstalling." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "  Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }
}

function Show-Menu {
    param(
        [array]$Devices,
        [object]$CurrentDefault,
        [object]$CurrentComms,
        [int]$SelectedIndex
    )

    Clear-Host

    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host "         AUDIO OUTPUT DEVICE SWITCHER         " -ForegroundColor Cyan
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Use " -NoNewline -ForegroundColor White
    Write-Host "[^]" -NoNewline -ForegroundColor Yellow
    Write-Host " " -NoNewline
    Write-Host "[v]" -NoNewline -ForegroundColor Yellow
    Write-Host " arrows to select, " -NoNewline -ForegroundColor White
    Write-Host "[Enter]" -NoNewline -ForegroundColor Yellow
    Write-Host " to confirm" -ForegroundColor White
    Write-Host ""

    for ($i = 0; $i -lt $Devices.Count; $i++) {
        $device = $Devices[$i]
        $isSelected = ($i -eq $SelectedIndex)
        $isDefault = ($null -ne $CurrentDefault) -and ($device.ID -eq $CurrentDefault.ID)
        $isComms = ($null -ne $CurrentComms) -and ($device.ID -eq $CurrentComms.ID)

        # Build prefix
        if ($isSelected) {
            $prefix = "  > "
        }
        else {
            $prefix = "    "
        }

        # Build suffix based on which defaults this device is
        if ($isDefault -and $isComms) {
            $suffix = "  (default + comms)"
        }
        elseif ($isDefault) {
            $suffix = "  (default)"
        }
        elseif ($isComms) {
            $suffix = "  (comms)"
        }
        else {
            $suffix = ""
        }

        $isCurrent = $isDefault -and $isComms

        # Determine colors and write
        if ($isSelected -and $isCurrent) {
            Write-Host $prefix -NoNewline -ForegroundColor Cyan
            Write-Host $device.Name -NoNewline -ForegroundColor Green
            Write-Host $suffix -ForegroundColor Green
        }
        elseif ($isSelected -and ($isDefault -or $isComms)) {
            Write-Host $prefix -NoNewline -ForegroundColor Cyan
            Write-Host $device.Name -NoNewline -ForegroundColor Cyan
            Write-Host $suffix -ForegroundColor Yellow
        }
        elseif ($isSelected) {
            Write-Host $prefix -NoNewline -ForegroundColor Cyan
            Write-Host "$($device.Name)$suffix" -ForegroundColor Cyan
        }
        elseif ($isCurrent) {
            Write-Host $prefix -NoNewline
            Write-Host $device.Name -NoNewline -ForegroundColor Green
            Write-Host $suffix -ForegroundColor Green
        }
        elseif ($isDefault -or $isComms) {
            Write-Host $prefix -NoNewline
            Write-Host $device.Name -NoNewline -ForegroundColor White
            Write-Host $suffix -ForegroundColor Yellow
        }
        else {
            Write-Host "$prefix$($device.Name)$suffix" -ForegroundColor White
        }
    }

    Write-Host ""
    Write-Host "  Press " -NoNewline -ForegroundColor Gray
    Write-Host "[F5]" -NoNewline -ForegroundColor Yellow
    Write-Host " to refresh, " -NoNewline -ForegroundColor Gray
    Write-Host "[Escape]" -NoNewline -ForegroundColor Yellow
    Write-Host " to exit" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
}

function Select-AudioDevice {
    # Install module if needed
    Install-AudioModule

    $state = Get-DeviceState
    $devices = $state.Devices
    $currentDefault = $state.Default
    $currentComms = $state.Comms

    if ($devices.Count -eq 0) {
        Clear-Host
        Write-Host ""
        Write-Host "  No audio devices found!" -ForegroundColor Red
        Write-Host ""
        Write-Host "  Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }

    # Start with current device selected
    $selectedIndex = 0
    for ($i = 0; $i -lt $devices.Count; $i++) {
        if ($devices[$i].ID -eq $currentDefault.ID) {
            $selectedIndex = $i
            break
        }
    }

    while ($true) {
        Show-Menu -Devices $devices -CurrentDefault $currentDefault -CurrentComms $currentComms -SelectedIndex $selectedIndex

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

        switch ($key.VirtualKeyCode) {
            38 {
                # Up arrow
                $selectedIndex--
                if ($selectedIndex -lt 0) {
                    $selectedIndex = $devices.Count - 1
                }
            }
            40 {
                # Down arrow
                $selectedIndex++
                if ($selectedIndex -ge $devices.Count) {
                    $selectedIndex = 0
                }
            }
            13 {
                # Enter
                $selectedDevice = $devices[$selectedIndex]

                # Check if already set as both default and communication device
                $isAlreadyDefault = ($selectedDevice.ID -eq $currentDefault.ID)
                $isAlreadyComms = ($selectedDevice.ID -eq $currentComms.ID)

                if ($isAlreadyDefault -and $isAlreadyComms) {
                    Write-Host ""
                    Write-Host "  Already using: $($selectedDevice.Name)" -ForegroundColor Yellow
                    Write-Host "  (both default and communication device)" -ForegroundColor Yellow
                    Write-Host ""
                    Write-Host "  Press any key to exit, or wait 2 seconds..." -ForegroundColor Gray

                    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    while ($stopwatch.ElapsedMilliseconds -lt $Script:Timeouts.AlreadySelected) {
                        if ($Host.UI.RawUI.KeyAvailable) {
                            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                            break
                        }
                        Start-Sleep -Milliseconds 100
                    }
                    exit 0
                }

                try {
                    # Set as both default playback and communication device
                    Set-AudioDevice -ID $selectedDevice.ID | Out-Null

                    Write-Host ""
                    Write-Host "  SUCCESS! Now using: $($selectedDevice.Name)" -ForegroundColor Green
                    Write-Host "  (set as default and communication device)" -ForegroundColor Green
                    Write-Host ""
                    Write-Host "  Press any key to exit, or wait 3 seconds..." -ForegroundColor Gray

                    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    while ($stopwatch.ElapsedMilliseconds -lt $Script:Timeouts.Success) {
                        if ($Host.UI.RawUI.KeyAvailable) {
                            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                            break
                        }
                        Start-Sleep -Milliseconds 100
                    }
                    exit 0
                }
                catch {
                    Write-Host ""
                    Write-Host "  ERROR: Could not switch audio device." -ForegroundColor Red
                    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
                    Write-Host ""
                    Write-Host "  Press any key to try again..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

                    # Refresh device list in case something changed
                    $state = Get-DeviceState
                    $devices = $state.Devices
                    $currentDefault = $state.Default
                    $currentComms = $state.Comms
                    # Reset selection after error
                    $selectedIndex = 0
                }
            }
            116 {
                # F5 - refresh device list
                $state = Get-DeviceState
                $devices = $state.Devices
                $currentDefault = $state.Default
                $currentComms = $state.Comms
                # Keep selection in bounds after refresh
                if ($selectedIndex -ge $devices.Count) {
                    $selectedIndex = [Math]::Max(0, $devices.Count - 1)
                }
            }
            27 {
                # Escape
                Write-Host ""
                Write-Host "  Goodbye!" -ForegroundColor Cyan
                Write-Host ""
                Start-Sleep -Milliseconds $Script:Timeouts.Goodbye
                exit 0
            }
        }
    }
}

# Run the script
Select-AudioDevice
