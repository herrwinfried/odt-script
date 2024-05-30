<div align="center">
  <img src="https://upload.wikimedia.org/wikipedia/commons/thumb/0/0e/Microsoft_365_%282022%29.svg/1862px-Microsoft_365_%282022%29.svg.png" width="150" height="150">
  <img src="https://raw.githubusercontent.com/PowerShell/PowerShell/master/assets/ps_black_64.svg?sanitize=true" width="150" height="150">
  <p><b>Powershell Script to Simplify Office Deployment Tool Usage</b></p>
</div>

<img src="./assets/image/odt-script.png">

# Office Deployment Tool PowerShell Script

This PowerShell script is designed to make the Office Deployment Tool easier to use by providing a simple and interactive interface.

## Usage

### Option 1: Run with Installer

Double-click the `installer.bat` file.

#### OR

Run the following PowerShell command:

```bat
start powershell.exe -ExecutionPolicy Bypass -File .\installer.ps1

```
### Option 2: Run in Command Line Interface (CLI)
```powershell
    Usage: 
        .\installer.ps1 [-h] [-c <ConfigFile>] [-i] [-d]

    Parameters:
        -h, -?, -Help
            Displays this help message.

        -c, -ConfigFile <ConfigFile>
            Specifies the configuration file to be used.

        -i, -Install
            Triggers the installation process.

        -d, -Download
            Starts the download process.
```
#### Example Command
```powershell
powershell.exe -ExecutionPolicy Bypass -File .\installer.ps1 -i -d
```

## Useful Links

- [config.office.com](https://config.office.com/) ([Documentation](https://learn.microsoft.com/en-us/deployoffice/office-deployment-tool-configuration-options)) 
  - An easy way to create configuration files for Office deployment.
- [SetupProd_OffScrub.exe](https://aka.ms/SaRA-OfficeUninstallFromPC) ([Documentation](https://support.microsoft.com/en-us/office/uninstall-office-automatically-9ad57b43-fa12-859a-9cf0-b694637b3b05))
  - This application helps when the Office installation is corrupted. It can be used to uninstall and reinstall Office, and is particularly useful for uninstallation purposes.