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
- The config panel shows device format, mix format, current WASAPI device-period latency, volume/mute, state, system effects/APO state, and exclusive-mode state.
- For virtual drivers such as Voicemeeter, some endpoint properties may reject direct changes. The app tries a backed registry fallback and marks it as pending restart when needed.

## Advanced Tweaks

- MMCSS/SystemProfile tuning with backup restore.
- Ducking disable/restore.
- UI menu-delay tweak/restore.
- Exclusive mode enable/disable for selected devices or all devices.
- Low-latency power options, USB selective suspend, PCIe link-state, audio service repair, and power report.
- Registry backups are stored under `HKCU\Software\AudioOptimizer\Backup`.

Driver-specific options such as Realtek jack detection, Bluetooth AAC/HFP, Dolby/DTS, Loudness Equalization, and USB hub power-management checkboxes must be exposed by the installed driver or configured in Windows Device Manager/Sound settings.
