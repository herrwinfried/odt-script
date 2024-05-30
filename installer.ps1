param(
    [Alias("h", "?")]
    [switch]$Help,
    [Alias("c")]
    [string]$ConfigFile,
    [Alias("i")]
    [switch]$Install,
    [Alias("d")]
    [switch]$Download
)
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

$Host.UI.RawUI.WindowTitle = "$($Language.ScriptTitle -f 'herrwinfried')"

[string]$Global:ConfigPath = "$GetScriptDir\Configuration.xml"
[string]$Global:ODTPath = "$GetScriptDir\odt.exe"
[string]$Global:ODTDownloadURL = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17531-20046.exe"
[string]$Global:SetupPath = "$GetScriptDir\setup.exe"
[string]$Global:ConfigGeneratorURL = "https://config.office.com/deploymentsettings"
[string]$Global:GetOfficePath = "$GetScriptDir\office"

function Test-Setup {

    if (-Not (Test-Path $SetupPath)) {
        Write-Host "$($Language.TestSetup)"
        if (-Not (Invoke-ODTExtraction)) {
            return $false
        }
    }
    return $true
}

function Test-ODT {

    if (-Not (Test-Path $ODTPath)) {
        if (-Not (Invoke-ODTExtraction)) {
            return $false
        }
    }
    return $true
}

function Test-ConfigFile {
    if (-Not (Test-Path $ConfigPath)) {
        Write-Host -ForegroundColor Red "$($Language.ConfigFile1 -f $ConfigPath)"
        Write-Host -ForegroundColor DarkBlue "$ConfigGeneratorURL"
        Write-Host -ForegroundColor DarkBlue "$($Language.ConfigFile2)"
        Show-Menu
        return $false
    }
    return $true
}

function Invoke-ODTExtraction {
 
    if (-Not (Test-Path $ODTPath)) {
        Write-Host "$($Language.NotFound -f $ODTPath) $($Language.DownloadInternet)"
        Write-Information "$ODTDownloadURL"
        
        try {
            Invoke-WebRequest -Uri $ODTDownloadURL -OutFile $ODTPath
            Write-Host -ForegroundColor Green "$($Language.DownloadSuccess -f 'odt.exe')"
        } catch {
            Write-Host -ForegroundColor Red "$($Language.DownloadFailed -f 'odt.exe')"
            return $false
        }
    }

    Write-Host -ForegroundColor Cyan "$($Language.Extracting -f 'odt.exe')"
    & $ODTPath /extract:$GetScriptDir /quiet
    if ($LASTEXITCODE -eq 0) {
        Write-Host -ForegroundColor Green "$($Language.AddedSuccess -f 'setup.exe')"
        return $true
    } else {
        Write-Host -ForegroundColor Red "$($Language.ExtractingFailed -f 'setup.exe')"
        return $false
    }
}

function Invoke-Office {

    if (-Not (Test-ODT)) {
        return
    }

    if (-Not (Test-Setup)) {
        return
    }

    if (-Not (Test-ConfigFile)) {
        return
    }
    Write-Host -ForegroundColor Green "$($Language.Starting -f $SetupPath)"
    Write-Host -ForegroundColor Yellow "$($Language.InvokeOfficeInfo)"
    Write-Host -ForegroundColor Yellow "$($Language.InvokeOfficeInfo1 -f "$GetOfficePath")"
    & $SetupPath /download $ConfigPath
    if ($LASTEXITCODE -eq 0) {
        Write-Host -ForegroundColor Green "$($Language.DownloadSuccessOffice -f 'office')"
    } else {
        Write-Host -ForegroundColor Red "$($Language.DownloadFailedOffice -f 'office')"
    }
    Write-Host "`a"
}

function Install-Office {
    if (-Not (Test-ODT)) {
        return
    }

    if (-Not (Test-Setup)) {
        return
    }

    if (-Not (Test-ConfigFile)) {
        return
    }

    if (-Not (Test-Path "$GetOfficePath")) {
    Write-Host -ForegroundColor Yellow "$($Language.NotFound -f "$GetOfficePath")"
    return 
    }

    Write-Host -ForegroundColor Green "$($Language.Starting -f $SetupPath)"
    & $SetupPath /configure $ConfigPath
    if ($LASTEXITCODE -eq 0) {
        Write-Host -ForegroundColor Green "$($Language.InstallSuccessOffice -f 'office')"
    } else {
        Write-Host -ForegroundColor Red "$($Language.InstallFailedOffice -f 'office')"
    }
    Write-Host "`a"
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
            Write-Host -ForegroundColor Green "$($Language.SetConfigSuccessPath -f $ConfigPath)"
            Write-Warning "$($Language.SetConfigSuccessWarn)"
            Write-host
        } else {
            Write-Host -ForegroundColor Red "$($Language.SetConfigFailedPath -f $newConfigPath)"
            Write-host
        }
    }
}

function Show-Welcome {
    Write-Host -ForegroundColor Magenta "Github: https://github.com/herrwinfried/odt-script"
    Write-Host -ForegroundColor Green "$($Language.ShowMenuConfigPath -f $ConfigPath)"
    Write-Host -ForegroundColor Cyan "$($Language.ShowMenuOfficePath -f $GetOfficePath)"
    Write-Host ""
}

function Show-Menu {
    Show-Welcome
    Write-Host -ForegroundColor DarkYellow "$($Language.ShowMenuSelection)`n"
    Write-Host -ForegroundColor Red "[0] $($Language.ShowMenuZero)"
    Write-Host -ForegroundColor DarkCyan "[1] $($Language.ShowMenuOne)"
    Write-Host -ForegroundColor DarkBlue "[2] $($Language.ShowMenuTwo)"
    Write-Host -ForegroundColor DarkBlue "[3] $($Language.ShowMenuThree)"
    Write-Host -ForegroundColor DarkMagenta "[4] $($Language.ShowMenuFour)"
    $choice = Read-Host "$($Language.ShowMenuInput)"
    
    switch ($choice) {
        0 {
            Exit 0
        }
        1 {
            Write-Host -ForegroundColor DarkGreen "`n$($Language.Starting -f $Language.ShowMenuTwo)`n"
            Invoke-Office
            Write-Host -ForegroundColor DarkGreen "`n$($Language.Starting -f $Language.ShowMenuThree)`n"
            Install-Office
            Show-Menu
        }
        2 {
            Write-Host -ForegroundColor DarkGreen "`n$($Language.Starting -f $Language.ShowMenuTwo)`n"
            Invoke-Office
            Show-Menu
        }
        3 {
            Write-Host -ForegroundColor DarkGreen "`n$($Language.Starting -f $Language.ShowMenuThree)`n"
            Install-Office
            Show-Menu
        }
        4 {
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

if ($ConfigFile) {
    if (Test-Path $ConfigFile) {
        [string]$Global:ConfigPath = $ConfigFile
    Write-Host -ForegroundColor Green "$($Language.SetConfigSuccessPath -f $ConfigFile)"
    Write-Warning "$($Language.SetConfigSuccessWarn)"
    Write-host
} else {
    Write-Host -ForegroundColor Red "$($Language.SetConfigFailedPath -f $ConfigFile)"
    Write-host
}
}
if ($Help) {
    $helpText = @"
    $($Language.HelpUsage) 
        .\$GetScriptName.ps1 [-h] [-c <ConfigFile>] [-i] [-d]

    $($Language.HelpParam)
        -h, -?, -Help
            $($Language.Helphelp)

        -c, -ConfigFile <ConfigFile>
            $($Language.HelpConfig)

        -i, -Install
            $($Language.HelpInstall)

        -d, -Download
            $($Language.HelpDownload)
"@
    Write-Output $helpText
    exit 1
}

if ((-Not ($Install)) -and (-Not ($Download))) {
Show-Menu
}
elseif (($Install) -or ($Download)) {
    Show-Welcome

    if ($Download) {
        Write-Host -ForegroundColor DarkGreen "`n$($Language.Starting -f $Language.ShowMenuTwo)`n"
        Invoke-Office
        Read-Host
    }
    if ($Install) {
        Write-Host -ForegroundColor DarkGreen "`n$($Language.Starting -f $Language.ShowMenuThree)`n"
        Install-Office
        Read-Host
    }
}
