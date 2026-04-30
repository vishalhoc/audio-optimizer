# Audio Optimizer

Native Win32 Windows audio control utility.

## Build

```powershell
.\build.ps1
```

The portable executable is created at `dist\AudioOptimizer.exe`.

## Device UI

- Highlight a row to inspect that device and sync the format/volume controls.
- Check one or more rows to choose the devices targeted by Apply/Reset/Volume/Effects/Visibility actions.
- Format controls include Apply checked, Reset checked, Outputs, Inputs, and All so input devices can be updated directly.
- Format apply logs summarize updated, skipped, and failed devices for outputs/inputs.
- The config panel shows device format, mix format, current WASAPI device-period latency, volume/mute, state, system effects/APO state, and exclusive-mode state.
- For virtual drivers such as Voicemeeter, some endpoint properties may reject direct changes. The app tries a backed registry fallback and marks it as pending restart when needed.
- Industrial native UI pass:
  - Segoe UI typography.
  - Light operations-console palette.
  - Color-coded buttons for apply, repair, restore, danger, and panel actions.
  - Live status strip for the latest action.
  - Persistent timestamped log file.

## Advanced Tweaks

- MMCSS/SystemProfile tuning with backup restore.
- Ducking disable/restore.
- UI menu-delay tweak/restore.
- Exclusive mode enable/disable for selected devices or all devices.
- Low-latency power options, USB selective suspend, PCIe link-state, audio service repair, and power report.
- Crackling repair actions:
  - Safe repair checked: stable 48 kHz / 24-bit, disable endpoint effects, disable ducking.
  - Safe repair all: applies safe repair to all endpoints and optimizes audio services.
  - Aggressive repair: safe all-device repair, disables exclusive mode for shared stability, applies low-latency power, and restarts audio services.
- Diagnostics:
  - Open Windows audio troubleshooter/settings.
  - Open Device Manager for audio/chipset/Bluetooth drivers.
  - Open Bluetooth settings for hands-free profile checks.
  - Open/Clear persistent log.
  - Export full debug report.
- Registry backups are stored under `HKCU\Software\AudioOptimizer\Backup`.
- Runtime logs are written to `dist\AudioOptimizer.log`.

Driver-specific options such as Realtek jack detection, Bluetooth AAC/HFP, Dolby/DTS, Loudness Equalization, and USB hub power-management checkboxes must be exposed by the installed driver or configured in Windows Device Manager/Sound settings.
