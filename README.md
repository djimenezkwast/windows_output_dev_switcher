# Windows Audio Output Switcher

Simple utility for switching the default Windows audio output device via a menu interface.

## Purpose

Allow easy switching between audio devices (e.g., bluetooth transmitter on 3.5mm jack vs soundbar) without navigating Windows sound settings.

## Files

- `Switch-Audio-Output.bat` - Double-click launcher that runs the PowerShell script
- `Set-DefaultAudioDevice.ps1` - Main script with interactive menu

## How It Works

1. User double-clicks the `.bat` file
2. PowerShell script displays list of playback devices with arrow-key navigation
3. Current default device shown in green with "(current)" suffix
4. Selected item shown with `>` prefix in cyan
5. Arrow keys move selection, Enter confirms, Escape exits
6. Uses `AudioDeviceCmdlets` module from PSGallery (auto-installs on first run)

## Key Dependencies

- PowerShell with execution policy bypass (handled by .bat launcher)
- `AudioDeviceCmdlets` module - installed automatically to CurrentUser scope
