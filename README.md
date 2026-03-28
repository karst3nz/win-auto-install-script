# Windows Auto-Install Script

Automated Windows setup script with Chocolatey package installation.

## Requirements

- **Windows** 10/11
- **PowerShell** 5.1 or higher
- **Administrator privileges** (script requires running as administrator)
- **Internet connection**

## Quick Start

```powershell
# Run as administrator
.\setup.ps1
```

## Command Line Parameters

| Parameter | Description |
|-----------|-------------|
| `-NoRestart` | Don't prompt for restart after completion |
| `-DryRun` | Simulation mode without actual changes |
| `-ConfigPath <path>` | Path to JSON configuration file (default: `.\packages.config.json`) |
| `-Uninstall` | Uninstall packages from configuration file |
| `-ExportConfig <path>` | Export installed packages to JSON |
| `-SkipActivation` | Skip Windows activation |
| `-DriversPath <path>` | Path to folder with drivers/applications (.exe and .msi files) |

### Usage Examples

```powershell
# Standard installation
.\setup.ps1

# Simulation mode (no actual changes)
.\setup.ps1 -DryRun

# Export installed packages
.\setup.ps1 -ExportConfig "my-packages.json"

# Uninstall packages from config
.\setup.ps1 -Uninstall

# Installation without restart
.\setup.ps1 -NoRestart

# Skip Windows activation
.\setup.ps1 -SkipActivation

# Use alternative config
.\setup.ps1 -ConfigPath "C:\path\to\custom-config.json"

# Install drivers/applications from custom folder
.\setup.ps1 -DriversPath "D:\drivers"
```

## Configuration File

The `packages.config.json` file contains the list of packages to install:

```json
{
  "packages": [
    {
      "id": "7zip",
      "name": "7-Zip"
    },
    {
      "id": "googlechrome",
      "name": "Google Chrome"
    },
    {
      "id": "vlc",
      "name": "VLC Media Player"
    }
  ]
}
```

### Required Fields

| Field | Type | Description |
|-------|------|-------------|
| `id` | string | Package identifier in Chocolatey repository |
| `name` | string | Display name of the package |

### Finding Packages

Find packages at [community.chocolatey.org/packages](https://community.chocolatey.org/packages)

## What the Script Does

1. **Checks internet connection**
2. **Installs/updates Chocolatey**
3. **Installs packages** from `packages.config.json`
4. **Activates Windows** (via MAS script)
5. **Enables dark theme**
6. **Adds context menu items:**
   - "Copy as Path" — copy file/folder path
   - "Open in PowerShell" — open PowerShell in folder
7. **Installs drivers/applications** from `drivers` folder (.exe and .msi files)
8. **Cleans up old logs** (older than 30 days)

## Project Structure

```
D:\win-auto-install-script\
├── setup.ps1              # Main script
├── packages.config.json   # Package configuration
├── README.md              # Documentation
├── setup_*.log            # Run logs
├── start.bat              # Batch launcher
└── drivers\               # Folder for .exe and .msi installers
    ├── app1.exe
    └── app2.msi
```

## Logs

Each run creates a log file with timestamp:
```
setup_20260327_103240.log
```

Logs are stored for 30 days, then automatically deleted.

## Security

### Script Hash Verification

The script downloads MAS (Microsoft Activation Script) with verification:
- Download with retries (up to 3 times)
- SHA256 hash calculation of downloaded file
- Optional verification against known hash

To set expected hash, edit `$Configuration.MAS.ExpectedHash` at the beginning of the script.

### Rollback Support

Script supports rollback:
- Uninstall packages on error
- Restore registry theme settings

### Script Signing

For production use, signing the script is recommended:

```powershell
# Create self-signed certificate
$cert = New-SelfSignedCertificate -DnsName "WinAutoInstall" -CertStoreLocation "Cert:\CurrentUser\My"

# Sign the script
Set-AuthenticodeSignature -FilePath ".\setup.ps1" -Certificate $cert
```

## Troubleshooting

### Error: "Script execution is disabled"

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Package Installation Error

1. Check package name at community.chocolatey.org
2. Try manual installation: `choco install <package-id>`
3. Check log file `setup_*.log`

### Windows Activation Error

- Ensure internet connection is available
- Try with `-SkipActivation` and activate manually
- Check [MAS GitHub](https://github.com/massgrave/MAS)

### Drivers/Applications Installation

- Place `.exe` or `.msi` installers in the `drivers` folder
- For `.exe` files, script tries common silent install arguments
- For `.msi` files, script uses `msiexec /quiet /norestart`

## Contributing

### Adding Packages

Edit `packages.config.json`:

```json
{
  "packages": [
    {
      "id": "your-package-id",
      "name": "Your Package Name"
    }
  ]
}
```

### Reporting Issues

Please include:
1. PowerShell version: `$PSVersionTable.PSVersion`
2. Log file `setup_*.log`
3. Problem description

## License

MIT License — free use and modification.

## Links

- [Chocolatey](https://chocolatey.org/)
- [MAS (Microsoft Activation Script)](https://github.com/massgrave/MAS)
- [PowerShell Documentation](https://docs.microsoft.com/powershell/)
