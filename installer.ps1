$GetScriptDir = $PSScriptRoot
[string]$GetScriptName = [System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Name)

# Language support
if (-Not ($Lang)) { $Lang = $PSUICulture }
$GetLanguageFile = "$PSScriptRoot/lang/$Lang/$GetScriptName.psd1"

if (-Not (Test-Path "$PSScriptRoot/lang/en-US/$GetScriptName.psd1")) {
        Write-Error "There is no language file. Script exited."
        Exit 1
} elseif (Test-Path $GetLanguageFile) {
    $Language = Import-LocalizedData -BaseDirectory "$PSScriptRoot/lang" -FileName $($GetScriptName).psd1 -UICulture $Lang
} elseif (-Not (Test-Path $GetLanguageFile)) {
    $Language = Import-LocalizedData -BaseDirectory "$PSScriptRoot/lang" -FileName $($GetScriptName).psd1 -UICulture "en-US"
}
# language support complete

$Host.UI.RawUI.WindowTitle = "HerrWinfried - $($Language.ScriptTitle)"

[string]$Global:ConfigPath = "$GetScriptDir\Configuration.xml"

function Invoke-ODTExtraction {
    $odtPath = "$GetScriptDir\odt.exe"
    $odtDownloadUrl = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17531-20046.exe"

    if (-Not (Test-Path $odtPath)) {
        Write-Host "$($Language.NotFoundOdt) $($Language.DownloadInternet)"
        Write-Information "$odtDownloadUrl"
        
        try {
            Invoke-WebRequest -Uri $odtDownloadUrl -OutFile $odtPath
            Write-Host -ForegroundColor Green "$($Language.OdtDownload)"
        } catch {
            Write-Host -ForegroundColor Red "$($Language.OdtDownloadNot)"
            return $false
        }
    }

    Write-Host -ForegroundColor Cyan "$($Language.OdtSetupExeExt)"
    & $odtPath /extract:$GetScriptDir /quiet
    if ($LASTEXITCODE -eq 0) {
        Write-Host -ForegroundColor Green "$($Language.OdtSetupExe)"
        return $true
    } else {
        Write-Host -ForegroundColor Red "$($Language.OdtSetupExeNot)"
        return $false
    }
}

function Test-SetupExe {
    $SetupPath = "$GetScriptDir\setup.exe"
    if (-Not (Test-Path $SetupPath)) {
        Write-Host "$($Language.TestSetupExe)"
        if (-Not (Invoke-ODTExtraction)) {
            return $false
        }
    }
    return $true
}

function Test-ConfigFile {
    if (-Not (Test-Path $Global:ConfigPath)) {
        Write-Host -ForegroundColor Red "$($Language.TestConfigFileInfo -f $Global:ConfigPath)"
        Write-Host -ForegroundColor DarkBlue "https://config.office.com/deploymentsettings"
        Write-Host -ForegroundColor DarkBlue "$($Language.TestConfigFileInfoHelp)"
        Show-Menu
        return $false
    }
    return $true
}

function Invoke-Office {
    if (-Not (Test-SetupExe)) {
        Write-Host -ForegroundColor Red "$($Language.InvokeOfficeNotFoundSetupExe)"
        return
    }

    if (Test-Path "$GetScriptDir\office.old") {
        Write-Host -ForegroundColor Yellow "$($Language.InvokeOfficeFoundOffice_old -f "$GetScriptDir\office.old" )"
        Remove-Item -Path "$GetScriptDir\office.old" -Force -Recurse
    }

    if (Test-Path "$GetScriptDir\office") {
        Write-Host -ForegroundColor Yellow ($Language.InvokeOfficeFoundOffice -f "$GetScriptDir\office", "$GetScriptDir\office.old")
        Move-Item -Path "$GetScriptDir\office" -Destination "$GetScriptDir\office.old" -Force
    }

    if (-Not (Test-ConfigFile)) {
        return
    }

    $SetupPath = "$GetScriptDir\setup.exe"
    Write-Host -ForegroundColor Green "$($Language.Starting -f $SetupPath) [Download]"
    Write-Host -ForegroundColor Yellow "$($Language.InvokeOfficeInfo)"
    Write-Host -ForegroundColor Yellow "$($Language.InvokeOfficeInfo1 -f "$GetScriptDir\office")"
    & $SetupPath /download $Global:ConfigPath
    if ($LASTEXITCODE -eq 0) {
        Write-Host -ForegroundColor Green "$($Language.InvokeOfficeDownload)"
    } else {
        Write-Host -ForegroundColor Red "$($Language.InvokeOfficeDownloadNot)"
    }
}

function Install-Office {
    if (-Not (Test-SetupExe)) {
        Write-Host -ForegroundColor Red "$($Language.TestSetupExe)"
        return
    }

    if (-Not (Test-ConfigFile)) {
        return
    }

    if (-Not (Test-Path "$GetScriptDir\office")) {
        Write-Host -ForegroundColor Yellow "$($Language.InstallOfficeFoundOfficeNot)"
        Invoke-Office
    }

    $SetupPath = "$GetScriptDir\setup.exe"
    Write-Host -ForegroundColor Green "$($Language.Starting -f $SetupPath) [Install]"
    & $SetupPath /configure $Global:ConfigPath
    if ($LASTEXITCODE -eq 0) {
        Write-Host -ForegroundColor Green "$($Language.InstallOfficeDownload)"
    } else {
        Write-Host -ForegroundColor Red "$($Language.InstallOfficeDownloadNot)"
    }
}

function Set-ConfigPath {
    Add-Type -AssemblyName System.Windows.Forms
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.Filter = "XML Files|*.xml"
    $OpenFileDialog.Title = "$($Language.OpenFileDialogTitle)"
    if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $newConfigPath = $OpenFileDialog.FileName
        if (Test-Path $newConfigPath) {
            $Global:ConfigPath = $newConfigPath
            Write-Host -ForegroundColor Green "$($Language.SetConfigSuccessPath -f $Global:ConfigPath)"
            Write-Warning "$($Language.SetConfigSuccessWarn)"
            Write-host
        } else {
            Write-Host -ForegroundColor Red "$($Language.SetConfigFailedPath -f $newConfigPath)"
            Write-host
        }
    }
}

function Show-Menu {
    Write-Host -ForegroundColor Green "$($Language.ShowMenuConfigPath) $Global:ConfigPath"
    Write-Host -ForegroundColor Cyan "$($Language.ShowMenuOfficePath) $GetScriptDir\office"
    Write-Host ""
    Write-Host -ForegroundColor DarkYellow "$($Language.ShowMenuSelection)`n"
    Write-Host -ForegroundColor Red "[0] $($Language.ShowMenuZero)"
    Write-Host -ForegroundColor DarkGray "[1] $($Language.ShowMenuOne)"
    Write-Host -ForegroundColor DarkBlue "[2] $($Language.ShowMenuTwo)"
    Write-Host -ForegroundColor DarkRed "[3] $($Language.ShowMenuFour)"
    $choice = Read-Host "$($Language.ShowMenuInput) (1, 2 or 3)"
    
    switch ($choice) {
        0 {
            Exit 0
        }
        1 {
            Invoke-Office
            Show-Menu
        }
        2 {
            Install-Office
            Show-Menu
        }
        3 {
            Set-ConfigPath
            Show-Menu
        }
        default {
            Write-Host -ForegroundColor Red "$($Language.ShowMenuInvalid)"
            Start-Sleep -Seconds 2
            Clear-Host
            Show-Menu
        }
    }
}

Show-Menu
