<#
.SYNOPSIS
    Easy audio output device switcher for Windows.
#>

# Set window title
$Host.UI.RawUI.WindowTitle = "Audio Output Switcher"

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
    Clear-Host

    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host "         AUDIO OUTPUT DEVICE SWITCHER         " -ForegroundColor Cyan
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""

    $devices = @(Get-AudioDevice -List | Where-Object { $_.Type -eq "Playback" })
    $currentDefault = Get-AudioDevice -Playback

    if ($devices.Count -eq 0) {
        Write-Host "  No audio devices found!" -ForegroundColor Red
        Write-Host ""
        Write-Host "  Press any key to exit..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        exit 1
    }

    Write-Host "  Select your audio output device:" -ForegroundColor White
    Write-Host ""

    for ($i = 0; $i -lt $devices.Count; $i++) {
        $device = $devices[$i]
        $num = $i + 1

        if ($device.ID -eq $currentDefault.ID) {
            Write-Host "    [$num] $($device.Name)" -ForegroundColor Green -NoNewline
            Write-Host "  << CURRENT" -ForegroundColor Green
        }
        else {
            Write-Host "    [$num] $($device.Name)" -ForegroundColor White
        }
    }

    Write-Host ""
    Write-Host "    [0] Exit" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  ============================================" -ForegroundColor Cyan
    Write-Host ""

    return $devices
}

function Select-AudioDevice {
    # Install module if needed
    Install-AudioModule

    while ($true) {
        $devices = Show-Menu

        Write-Host "  Type a number and press Enter: " -ForegroundColor Yellow -NoNewline
        $selection = Read-Host

        if ($selection -match '^\d{1,2}$') {
            $index = [int]$selection

            if ($index -eq 0) {
                Write-Host ""
                Write-Host "  Goodbye!" -ForegroundColor Cyan
                Write-Host ""
                Start-Sleep -Milliseconds 500
                exit 0
            }

            if ($index -ge 1 -and $index -le $devices.Count) {
                $selectedDevice = $devices[$index - 1]

                try {
                    $selectedDevice | Set-AudioDevice | Out-Null

                    Write-Host ""
                    Write-Host "  SUCCESS! Now using: $($selectedDevice.Name)" -ForegroundColor Green
                    Write-Host ""
                    Write-Host "  Press any key to exit, or wait 3 seconds..." -ForegroundColor Gray

                    # Wait for keypress or timeout
                    $timeout = 3000
                    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
                    while ($stopwatch.ElapsedMilliseconds -lt $timeout) {
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
                }
            }
            else {
                Write-Host ""
                Write-Host "  Please enter a number between 0 and $($devices.Count)" -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
        else {
            Write-Host ""
            Write-Host "  Please enter a number" -ForegroundColor Red
            Start-Sleep -Seconds 1
        }
    }
}

# Run the script
Select-AudioDevice
