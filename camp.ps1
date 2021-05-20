$username=( ( Get-WMIObject -class Win32_ComputerSystem | Select-Object -ExpandProperty username ) -split '\\' )[1]
$Directory = "C:\Users\$username"
$StartMenu = "$Directory\AppData\Roaming\Microsoft\Windows\Start Menu\Programs"

$RobloxDownload = $True
$OBSDownload = $True
$MinecraftDownload = $True
$MCreatorDownload = $True

# Install Roblox Studio Software
if ($RobloxDownload) {
    Write-Host "Installing Roblox Studio Software..." -ForegroundColor Green
    iwr -useb 'https://setup.rbxcdn.com/RobloxStudioLauncherBeta.exe' -OutFile $Directory\Documents\robloxstudiolauncherbeta.exe | Out-Null
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
   



# Run and then kill Roblox Studio after 15 seconds
$StudioProcess = Start-Process -FilePath $Directory\Documents\robloxstudiolauncherbeta.exe
$StudioProcess

Start-Sleep -Seconds 20
Stop-Process -Name "RobloxStudioBeta" -Force -Confirm:$false

# Run Minecraft silent MSI installation
$MinecraftProcess = Start-Process -FilePath $Directory\Documents\minecraftinstaller.msi -ArgumentList "/quiet /norestart"
$MinecraftProcess

# Run OBS silent installation
$OBSProcess = Start-Process -FilePath $Directory\Documents\obsinstaller.exe -ArgumentList "/S"
$OBSProcess





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
Copy-Item -Path $RobloxShortcut -Destination $StartMenu\. -Force
Copy-Item -Path $StuduinoShortcut -Destination $StartMenu\. -Force