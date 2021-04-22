#  Uncomment this mess for deployment

<#

@@:: This prolog allows a PowerShell script to be embedded in a .CMD file.
@@:: Any non-PowerShell content must be preceeded by "@@"
@@setlocal
@@set POWERSHELL_BAT_ARGS=%*
@@if defined POWERSHELL_BAT_ARGS set POWERSHELL_BAT_ARGS=%POWERSHELL_BAT_ARGS:"=\"%
@@PowerShell -ExecutionPolicy Bypass -Command Invoke-Expression $('$args=@(^&{$args} %POWERSHELL_BAT_ARGS%);'+[String]::Join(';',$((Get-Content '%~f0') -notmatch '^^@@'))) & goto :EOF

#>


# VARIABLES
$ConfirmPreference = 'None'

$RoboticsDownload = $False
$RoboticsHash = '9DFAA79A6651A4710711AE334E52B21BF7E2AF479C6EF84EF918B483F35F3090'

$FundamentalsDownload = $False
$FundamentalsHash = '32852B1331DF3F45D61B93980D62D42CB9A333F90F73D9EE90890832B26E4AD2'


# Self-elevate the script if required
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
  Exit
 }
}


# Disable exec check 
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy Bypass -Force


# Install NuGet to get packages
if ((Get-PackageProvider -Name NuGet -ForceBootstrap).version -lt 2.8.5.201 | Out-Null ) {
    Write-Host "Adding NuGet as a Package Provider..." -ForegroundColor Green
    try {
        Install-PackageProvider -Name "NuGet" -MinimumVersion 2.8.5.201 -Scope CurrentUser -Confirm:$false -ForceBootstrap
    }
    catch {
        Write-Host "Couldn't add NuGet as a Package Provider" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        exit
    }
}
else {
    Write-Host "Version of NuGet installed = " (Get-PackageProvider -Name NuGet).version
}


# Install module to deal with Google Chrome Profiles
# It has a dependency on "TabExpansionPlusPlus" which has a conflict "-AllowClobber" supresses the warning
if (Get-Module -ListAvailable -Name PSChromeProfile) {
    Write-Host "PSChromeProfile already installed"
}
else {
    Write-Host "Installing 'PSChromeProfile' module..." -ForegroundColor Green

    # Trust the "PsGallery" NuGet repository
    Write-Host "Trusting 'PSGallery' repository..." -ForegroundColor Green
    try {
        Set-PSRepository -Name "PsGallery" -InstallationPolicy Trusted
    }
    catch {
        Write-Host "Couldn't trust 'PSGallery' repository" -ForegroundColor Red
        Write-Host $_
        exit
    }

    # Install PSChromeProfile module
    try {
        Install-Module -Name PSChromeProfile -Scope CurrentUser -AllowClobber
    }
    catch {
        Write-Host "Couldn't install 'PSChromeProfile' module" -ForegroundColor Red
        Write-Host $_
        exit
    }

    # Set "PsGallery" NuGet repository to untrusted
    Write-Host "Set 'PSGallery' repository to untrusted..." -ForegroundColor Yellow
    try {
        Set-PSRepository -Name "PsGallery" -InstallationPolicy Trusted
    }
    catch {
        Write-Host "Couldn't set 'PSGallery' repository to untrusted" -ForegroundColor Red
        Write-Host $_
        exit
    }
}

# Check if Fundamentals Assets exists and compare hash
if (Test-Path -Path $HOME\Documents\FUNdamentalsAssets.zip -PathType Leaf) {
    $FundamentalsActualHash = Get-FileHash $HOME\Documents\FUNdamentalsAssets.zip
    if ($FundamentalsHash -eq $FundamentalsActualHash.Hash) {
        Write-Host "Fundamentals Assets exist"
        Expand-Archive -Path $HOME\Documents\FUNdamentalsAssets.zip -DestinationPath $HOME\Documents\FUNdamentalsAssets -Force
    }
    else {
        $FundamentalsDownload = $True
    }
}
else {
    $FundamentalsDownload = $True
}

# Download and extract "FundamentalsAssets.zip"
if ($FundamentalsDownload) {
    Write-Host "Downloading and extracting Fundamentals Assets..." -ForegroundColor Green
    try {
    $source = 'https://archive.org/download/fundamentals-assets/FUNdamentalsAssets.zip'
    iwr -Uri $source -OutFile $HOME\Documents\FUNdamentalsAssets.zip
    Expand-Archive -Path $HOME\Documents\FUNdamentalsAssets.zip -DestinationPath $HOME\Documents\FUNdamentalsAssets -Force
}
    catch {
    Write-Host "Couldn't extract Fundamentals Assets" -ForegroundColor Red
    Write-Host $_
    exit
    }
}


# Check if Robotics PDF exists and compare hash
if (Test-Path -Path $HOME\Documents\robotics\robotics-beginner.pdf -PathType Leaf) {
    $RoboticsActualHash = Get-FileHash $HOME\Documents\robotics\robotics-beginner.pdf
    if ($RoboticsHash -eq $RoboticsActualHash.Hash) {
        Write-Host "Robotics curriculum exists"
    }
    else {
        $RoboticsDownload = $True
    }
}
else {
    $RoboticsDownload = $True
}
   

# Download Robotics PDF
if ($RoboticsDownload) {
    Write-Host "Downloading and Robotics curriculum..." -ForegroundColor Green
    try {
    $GoogleFileId = "12gIq7mU76MxUqH80i0Vkv1NGQ_r8Da5i"
    New-Item -ItemType Directory -Force -Path $HOME\Documents\robotics | Out-Null

    # set protocol to tls version 1.2
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Download the Virus Warning into _tmp.txt
    Invoke-WebRequest -Uri "https://drive.google.com/uc?export=download&id=$GoogleFileId" -OutFile "_tmp.txt" -SessionVariable googleDriveSession

    # Get confirmation code from _tmp.txt
    $searchString = Select-String -Path "_tmp.txt" -Pattern "confirm="
    $searchString -match "confirm=(?<content>.*)&amp;id="
    $confirmCode = $matches['content']

    # Delete _tmp.txt
    Remove-Item "_tmp.txt"

    # Download the real file
    Invoke-WebRequest -Uri "https://drive.google.com/uc?export=download&confirm=${confirmCode}&id=$GoogleFileId" -OutFile $HOME\Documents\robotics\robotics-beginner.pdf -WebSession $googleDriveSession
}
    catch {
        Write-Host "Couldn't download Robotics curriculum" -ForegroundColor Red
        Write-Host $_
        exit
    }
}

<#

# Download and execute Windows Debloater
Write-Host "Downloading and executing Windows Debloater..." -ForegroundColor Green
try {
    iwr -useb https://raw.githubusercontent.com/Sycnex/Windows10Debloater/master/Windows10SysPrepDebloater.ps1 | iex
}
catch {
    Write-Host "Couldn't run Windows Debloater" -ForegroundColor Red
    Write-Host $_
    exit
}

#>


# Move everything from $HOME\Downloads to $HOME\Documents\garbage
Write-Host "Clearing Downloads folder..." -ForegroundColor Green
try {
    New-Item -ItemType Directory -Force -Path $HOME\Documents\garbage | Out-Null
    Move-Item -Path $HOME\Downloads\* -Destination $HOME\Documents\garbage | Out-Null
}
catch {
    Write-Host "Couldn't clear Downloads folder" -ForegroundColor Red
    Write-Host $_
    exit
}


# Organize Desktop
Write-Host "Organizing Desktop..." -ForegroundColor Green
try {
    New-Item -ItemType Directory -Force -Path $HOME\Documents\desktop | Out-Null
    Move-Item -Path $HOME\Desktop\* -Destination $HOME\Documents\desktop -Force | Out-Null
}
catch {
    Write-Host "Couldn't organize Desktop" -ForegroundColor Red
    Write-Host $_
    exit
}

# Add shortcuts to Desktop
Write-Host "Adding shortcuts to Desktop..." -ForegroundColor Green
try {
    # Chrome shortcut
    $ChromeShortcut = "$HOME\Desktop\Chrome.lnk"
    $ChromeLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Chrome = $WScriptShell.CreateShortcut($ChromeShortcut)
    $Chrome.TargetPath = $ChromeLocation
    $Chrome.WorkingDirectory = "C:\Program Files (x86)\Google\Chrome\Application"
    $Chrome.Save()
    

    # Roblox Studio shortcut
    $RobloxShortcut = "$HOME\Desktop\Roblox Studio.lnk"
    $RobloxLocation = "$HOME\AppData\Local\Roblox\Versions\RobloxStudioLauncherBeta.exe"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Roblox = $WScriptShell.CreateShortcut($RobloxShortcut)
    $Roblox.TargetPath = $RobloxLocation
    $Roblox.Arguments = "-ide"
    $Roblox.Save()

    # White-Green assets


    # Studiuno software



}
catch {
    Write-Host "Couldn't add shortcuts to Desktop" -ForegroundColor Red
    Write-Host $_
    exit
}