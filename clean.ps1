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

$username=( ( Get-WMIObject -class Win32_ComputerSystem | Select-Object -ExpandProperty username ) -split '\\' )[1]
$Directory = "C:\Users\$username"
$StartMenu = "$Directory\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"

$ConfirmPreference = 'None'

$RoboticsDownload = $False
$RoboticsHash = '9DFAA79A6651A4710711AE334E52B21BF7E2AF479C6EF84EF918B483F35F3090'

$FundamentalsDownload = $False
$FundamentalsHash = '32852B1331DF3F45D61B93980D62D42CB9A333F90F73D9EE90890832B26E4AD2'

$StuduinoDowload = $False
$StuduinoHash = '8722B42DA261473A7B30BC036F05FFB478833D606E343D8B93AB450576488DD6'

$ArchiveDownload = $False
$ArchiveHash = '0C3A5517097E72DEA227E746388D61F65A7C2C991139B1E7D697E4D03B837BD9'

$RobloxDownload = $True

$ChromeDownload = $True
$ChromeUrl = 'https://dl.google.com/tag/s/appguid%3D%7B8A69D345-D564-463C-AFF1-A69D9E530F96%7D%26iid%3D%7BE8E3A411-32BE-217B-5101-1E794CAFE2CE%7D%26lang%3Den%26browser%3D4%26usagestats%3D0%26appname%3DGoogle%2520Chrome%26needsadmin%3Dprefers%26ap%3Dx64-stable-statsdef_1%26installdataindex%3Dempty/update2/installers/ChromeSetup.exe'

$MinecraftDownload = $True

$OBSDownload = $True

$MCreatorDownload = $True


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
    Install-PackageProvider -Name "NuGet" -MinimumVersion 2.8.5.201 -Scope CurrentUser -Confirm:$false -ForceBootstrap
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
    Set-PSRepository -Name "PsGallery" -InstallationPolicy Trusted


    # Install PSChromeProfile module
    Install-Module -Name PSChromeProfile -Scope CurrentUser -AllowClobber

    # Set "PsGallery" NuGet repository to untrusted
    Write-Host "Set 'PSGallery' repository to untrusted..." -ForegroundColor Yellow
    Set-PSRepository -Name "PsGallery" -InstallationPolicy Trusted
}


# Check if Fundamentals Assets exists and compare hash
if (Test-Path -Path $Directory\Documents\fundamentalsassets.zip -PathType Leaf) {
    $FundamentalsActualHash = Get-FileHash $Directory\Documents\fundamentalsassets.zip
    if ($FundamentalsHash -eq $FundamentalsActualHash.Hash) {
        Write-Host "Fundamentals Assets exist"
        Expand-Archive -Path $Directory\Documents\fundamentalsassets.zip -DestinationPath $Directory\Documents\fundamentalsassets -Force
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
    $source = 'https://archive.org/download/fundamentals-assets/FUNdamentalsAssets.zip'
    iwr -Uri $source -OutFile $Directory\Documents\fundamentalsassets.zip
    Expand-Archive -Path $Directory\Documents\fundamentalsassets.zip -DestinationPath $Directory\Documents\fundamentalsassets -Force
}



# Check if Image Archive exists and compare hash
if (Test-Path -Path $Directory\Documents\imagearchive.zip -PathType Leaf) {
    $ArchiveActualHash = Get-FileHash $Directory\Documents\imagearchive.zip
    if ($ArchiveHash -eq $ArchiveActualHash.Hash) {
        Write-Host "Image Archive exists"
        Expand-Archive -Path $Directory\Documents\imagearchive.zip -DestinationPath $Directory\Documents\imagearchive -Force
    }
    else {
        $ArchiveDownload = $True
    }
}
else {
    $ArchiveDownload = $True
}

# Download and extract "ImageArchive.zip"
if ($ArchiveDownload) {
    Write-Host "Downloading and extracting Image Archive..." -ForegroundColor Green
    $source = 'https://archive.org/download/image-archive/ImageArchive.zip'
    iwr -Uri $source -OutFile $Directory\Documents\imagearchive.zip
    Expand-Archive -Path $Directory\Documents\imagearchive.zip -DestinationPath $Directory\Documents\imagearchive -Force
}


# Check if Robotics PDF exists and compare hash
if (Test-Path -Path $Directory\Documents\robotics\robotics-beginner.pdf -PathType Leaf) {
    $RoboticsActualHash = Get-FileHash $Directory\Documents\robotics\robotics-beginner.pdf
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

    $GoogleFileId = "12gIq7mU76MxUqH80i0Vkv1NGQ_r8Da5i"
    New-Item -ItemType Directory -Force -Path $Directory\Documents\robotics | Out-Null

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
    Invoke-WebRequest -Uri "https://drive.google.com/uc?export=download&confirm=${confirmCode}&id=$GoogleFileId" -OutFile $Directory\Documents\robotics\robotics-beginner.pdf -WebSession $googleDriveSession
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





# Check if Studuino Software exists and compare hash
if (Test-Path -Path $Directory\Documents\studuino.zip -PathType Leaf) {
    $StuduinoActualHash = Get-FileHash $Directory\Documents\studuino.zip
    if ($StuduinoHash -eq $StuduinoActualHash.Hash) {
        Write-Host "Studuino Software exists"
    }
    else {
        $StuduinoDownload = $True
    }
}
else {
    $StuduinoDowload = $True
}   


# Install Artec Studuino Software
if ($StuduinoDowload) {
    Write-Host "Installing Artec Studuino Software..." -ForegroundColor Green
    iwr -useb 'https://www.artec-kk.co.jp/studuino/data/software/cdn/en/Studuino.zip' -OutFile $Directory\Documents\Studuino.zip | Out-Null
    New-Item -ItemType Directory -Force -Path $Directory\Documents\studuino | Out-Null
    Expand-Archive -Path $Directory\Documents\studuino.zip -DestinationPath $Directory\Documents -Force
}

# Install Roblox Studio Software
if ($RobloxDownload) {
    Write-Host "Installing Roblox Studio Software..." -ForegroundColor Green
    iwr -useb 'https://setup.rbxcdn.com/RobloxStudioLauncherBeta.exe' -OutFile $Directory\Documents\robloxstudiolauncherbeta.exe | Out-Null
}


# Install Google Chrome Software
if ($ChromeDownload) {
    Write-Host "Installing Google Chrome..." -ForegroundColor Green
    iwr -useb $ChromeUrl -OutFile $Directory\Documents\chromeinstaller.exe | Out-Null
}

# Install Minecraft Software
if ($MinecraftDownload) {
    Write-Host "Installing Minecraft..." -ForegroundColor Green
    iwr -useb 'https://launcher.mojang.com/download/MinecraftInstaller.msi' -OutFile $Directory\Documents\minecraftinstaller.msi | Out-Null
}


# Install OBS Software
if ($OBSDownload) {
    Write-Host "Installing OBS..." -ForegroundColor Green
    iwr -useb 'https://cdn-fastly.obsproject.com/downloads/OBS-Studio-26.1.1-Full-Installer-x64.exe' -OutFile $Directory\Documents\obsinstaller.exe

}

# Install MCreator Software
if ($MCreatorDownload) {
    Write-Host "Installing MCreator..." -ForegroundColor Green
    iwr -useb 'https://mcreator.net/repository/2021-1/MCreator%202021.1%20Windows%2064bit.zip' -OutFile $Directory\Documents\mcreator.zip
    Remove-Item -Path $Directory\Documents\mcreator -Recurse -Force -Confirm:$False | Out-Null
    Expand-Archive -Path $Directory\Documents\mcreator.zip -DestinationPath $Directory\Documents -Force
    Move-Item -Path $Directory\Documents\MCreator20211 -Destination $Directory\Documents\mcreator -Force
}
   
<#

# Run and then kill Chrome Installer after 30 seconds

$ChromeProcess = Start-Process -FilePath $Directory\Documents\chromeinstaller.exe
$ChromeProcess

Start-Sleep -Seconds 30
Stop-Process -Name "chrome" -Force -Confirm:$false


# Run and then kill Roblox Studio after 15 seconds

$StudioProcess = Start-Process -FilePath $Directory\Documents\robloxstudiolauncherbeta.exe
$StudioProcess

Start-Sleep -Seconds 15
Stop-Process -Name "RobloxStudioBeta" -Force -Confirm:$false


# Run Minecraft silent MSI installation

$MinecraftProcess = Start-Process -FilePath $Directory\Documents\minecraftinstaller.msi -ArgumentList "/quiet /norestart"
$MinecraftProcess

# Run OBS silent installation

$OBSProcess = Start-Process -FilePath $Directory\Documents\obsinstaller.exe -ArgumentList "/S"
$OBSProcess

#>







# Move everything from $Directory\Downloads to $Directory\Documents\garbage
Write-Host "Clearing Downloads folder..." -ForegroundColor Green

New-Item -ItemType Directory -Force -Path $Directory\Documents\garbage | Out-Null
Move-Item -Path $Directory\Downloads\* -Destination $Directory\Documents\garbage | Out-Null

# Configure Chrome

$LocalAppData = "$Directory\AppData\Local"
$ChromeProfile = Join-Path $LocalAppData "Google\Chrome\User Data\"
$ChromeDefault = Join-Path $ChromeProfile "Default\"


Write-Host "Reset Google Chrome profile..." -ForegroundColor Green
iwr -useb 'https://archive.org/download/default_202105/Default.zip' -OutFile $Directory\Documents\Default.zip | Out-Null
New-Item -ItemType Directory -Force -Path $Directory\Documents\chrome | Out-Null
Move-Item -Path $ChromeDefault -Destination $Directory\Documents\chrome
Expand-Archive -Path $Directory\Documents\Default.zip -DestinationPath $ChromeProfile -Force | Out-Null



# Organize Desktop
Write-Host "Organizing Desktop..." -ForegroundColor Green

Remove-Item -Path $Directory\Documents\desktop\FundamentalsAssets -Recurse -Force -Confirm:$false | Out-Null
Remove-Item -Path $Directory\Documents\desktop\ImageArchive -Recurse -Force -Confirm:$false | Out-Null
Remove-Item -Path $Directory\Documents\desktop\Robotics -Recurse -Force -Confirm:$false | Out-Null

New-Item -ItemType Directory -Force -Path $Directory\Documents\desktop | Out-Null
Move-Item -Path $Directory\Desktop\* -Force -Destination $Directory\Documents\desktop | Out-Null

# Copy FUNdamentals Assets to Desktop
Write-Host "Adding Fundamentals Assets to Desktop..." -ForegroundColor Green
New-Item -ItemType Directory -Force -Path $Directory\Desktop\FUNdamentalsAssets | Out-Null
Copy-Item -Path $Directory\Documents\fundamentalsassets\* -Destination $Directory\Desktop\FUNdamentalsAssets | Out-Null


# Copy Image Archive to Desktop
Write-Host "Adding Image Archive to Desktop..." -ForegroundColor Green
New-Item -ItemType Directory -Force -Path $Directory\Desktop\ImageArchive | Out-Null
Copy-Item -Path $Directory\Documents\imagearchive\* -Destination $Directory\Desktop\ImageArchive | Out-Null


# Copy Robotics to Desktop
Write-Host "Adding Robotics to Desktop..." -ForegroundColor Green
New-Item -ItemType Directory -Force -Path $Directory\Desktop\Robotics| Out-Null
Copy-Item -Path $Directory\Documents\robotics\* -Destination $Directory\Desktop\Robotics| Out-Null


# Add shortcuts to Desktop
Write-Host "Adding shortcuts to Desktop..." -ForegroundColor Green

    # Chrome shortcut
    $ChromeShortcut = "$Directory\Desktop\Chrome.lnk"
    $ChromeLocation = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Chrome = $WScriptShell.CreateShortcut($ChromeShortcut)
    $Chrome.TargetPath = $ChromeLocation
    $Chrome.WorkingDirectory = "C:\Program Files (x86)\Google\Chrome\Application"
    $Chrome.Save()
    
    

    # Roblox Studio shortcut
    $RobloxShortcut = "$Directory\Desktop\Roblox Studio.lnk"
    if (Test-Path -Path $Directory\AppData\Local\Roblox\Versions\*\RobloxStudioLauncherBeta.exe -PathType Leaf) {
        $RobloxLocation = Resolve-Path -Path $Directory\AppData\Local\Roblox\Versions\*\RobloxStudioLauncherBeta.exe | Convert-Path
    }
    else {
        $RobloxLocation = Resolve-Path -Path $Directory\AppData\Local\Roblox\Versions\RobloxStudioLauncherBeta.exe | Convert-Path
    }
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Roblox = $WScriptShell.CreateShortcut($RobloxShortcut)
    $Roblox.TargetPath = $RobloxLocation
    $Roblox.Arguments = "-ide"
    $Roblox.Save()

    

    # Studiuno software
    $StuduinoShortcut = "$Directory\Desktop\Studuino.lnk"
    $StuduinoLocation = "$Directory\Documents\studuino\Application Files\ArtecRobotStartUp_1_5_7_1\ArtecRobotStartUp.exe"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Studuino = $WScriptShell.CreateShortcut($StuduinoShortcut)
    $Studuino.TargetPath = $StuduinoLocation
    $Studuino.WorkingDirectory = "$Directory\Documents\studuino\Application Files\ArtecRobotStartUp_1_5_7_1"
    $Studuino.Save()



    # Mcreator software
    $MCreatorShortcut = "$StartMenu\MCreator.lnk"
    $McreatorLocation = "$Directory\Documents\mcreator\MCreator.exe"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $MCreator = $WScriptShell.CreateShortcut($MCreatorShortcut)
    $MCreator.TargetPath = $MCreatorLocation
    $MCreator.Save()


# Add shortcuts to start menu
Write-Host "Adding shortcuts to Start Menu..." -ForegroundColor Green
   
   Copy-Item -Path $ChromeShortcut -Destination $StartMenu\. -Force
   Copy-Item -Path $RobloxShortcut -Destination $StartMenu\. -Force
   Copy-Item -Path $StuduinoShortcut -Destination $StartMenu\. -Force


# Remove Public user desktop shortcuts because reasons

Remove-Item -Path C:\Users\Public\Desktop\* -Recurse -Force -Confirm:$False

