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
3. Device status shown as suffix:
   - `(default + comms)` in green - device is both default and communication device
   - `(default)` in yellow - device is only the default device
   - `(comms)` in yellow - device is only the communication device
4. Selected item shown with `>` prefix in cyan
5. Arrow keys move selection, Enter confirms, Escape exits
6. F5 refreshes the device list (useful if devices are plugged/unplugged)
7. Selecting a device sets it as both default and communication device
8. Uses `AudioDeviceCmdlets` module from PSGallery (auto-installs on first run)

## Optional: Global Hotkey

You can assign a keyboard shortcut to launch the switcher from anywhere:

1. Right-click `Switch-Audio-Output.bat` → **Create shortcut**
2. Move the shortcut to your Desktop or Start Menu folder
3. Right-click the shortcut → **Properties**
4. Click the **Shortcut key** field and press your desired key combo (e.g., `Ctrl+Alt+A`)
5. Click **OK**

Now pressing your hotkey from any application will open the audio switcher.

**Note:** Windows only allows `Ctrl+Alt+<key>` combinations for shortcut keys.

## Key Dependencies

- PowerShell with execution policy bypass (handled by .bat launcher)
- `AudioDeviceCmdlets` module - installed automatically to CurrentUser scope
