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

$ArtecCert = 'MIIFRTCCBC2gAwIBAgIRAIWIG2z4gs9NOmhMhWDoQUUwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCR0IxGzAZBgNVBAgTEkdyZWF0ZXIgTWFuY2hlc3RlcjEQMA4GA1UEBxMHU2FsZm9yZDEYMBYGA1UEChMPU2VjdGlnbyBMaW1pdGVkMSQwIgYDVQQDExtTZWN0aWdvIFJTQSBDb2RlIFNpZ25pbmcgQ0EwHhcNMTkxMTE0MDAwMDAwWhcNMjAxMTEzMjM1OTU5WjCBjzELMAkGA1UEBhMCSlAxETAPBgNVBBEMCDU4MS0wMDY2MQ4wDAYDVQQIDAVPc2FrYTEMMAoGA1UEBwwDWWFvMR0wGwYDVQQJDBQzLTItMjEsIEtpdGFrYW1laWNobzEXMBUGA1UECgwOQVJURUMgQ08uLExURC4xFzAVBgNVBAMMDkFSVEVDIENPLixMVEQuMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAuqjiYA532wfiBN6trH/UXnSFeqVJJ7E/CbxuMcII6F/XOqy6y2L6Xs8AyLvdp+8WHJ3Y4ncc9R3q5HKjJ2QWUmk2U03bs3P/yXkdJuSdsFLq2sVlyEyyOB7R4Iv8rc8rDItidThkzlMARNZn5Nj6BMBRUYFaQGKLFOU9N9B0kSp7k+ooJw4bm75/R0tH7SQVScM/ApH2af/piCp3lRU6mgRnbWf3H6pnP5MllX+8p6OrJdANgZqFeMvLIOOOoUW5rOdQe/0t3gOlD6S5Wwj+VFYV/ddXNw064RXyJoGNsfxCw7gSK9VKicjSKumENPsPOIFiJaRuYbhjFNdRy6yzFQIDAQABo4IBrDCCAagwHwYDVR0jBBgwFoAUDuE6qFM6MdWKvsG7rWcaA4WtNA4wHQYDVR0OBBYEFO7mn0g6QhJfcIcA1SmpwXcu4uNuMA4GA1UdDwEB/wQEAwIHgDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMBEGCWCGSAGG+EIBAQQEAwIEEDBABgNVHSAEOTA3MDUGDCsGAQQBsjEBAgEDAjAlMCMGCCsGAQUFBwIBFhdodHRwczovL3NlY3RpZ28uY29tL0NQUzBDBgNVHR8EPDA6MDigNqA0hjJodHRwOi8vY3JsLnNlY3RpZ28uY29tL1NlY3RpZ29SU0FDb2RlU2lnbmluZ0NBLmNybDBzBggrBgEFBQcBAQRnMGUwPgYIKwYBBQUHMAKGMmh0dHA6Ly9jcnQuc2VjdGlnby5jb20vU2VjdGlnb1JTQUNvZGVTaWduaW5nQ0EuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAkBgNVHREEHTAbgRlzLWthZ2F5YW1hQGFydGVjLWtrLmNvLmpwMA0GCSqGSIb3DQEBCwUAA4IBAQA3dteIwsEl1h8l74JyyrPO/6o/gbS9gNu3r01hWvXDuAz8b3VvhoGe8NQ3gnv2d2idi98VSaN9u+nM98pO1qKg3TYIXLU4dCmR4XesFtVO9lNHbaIIGAR1uf2MCUPuLv2jKNj5QdErfkJxHopzemuPhfCeS8woRr3qjFjQER6+Y9vs5+GAdLLF8fvCC0Ku+dncYknjpXryCd3tu+3ixUiZ3hJyfSooTzdr1haSFSZH2Owed6s1x0eMPRvYZ2tXc2MiBGSwnA2q1okKSMH0fpe2WZvFe8yH1JRAwhCAxs524jSXmfdxiycWGNJVEHherOzcvihap56Ho21hyrKoyhyU'
$SectigoCert = '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'

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

# Import certificates for ClickOnce
Write-Host "Importing Artec cert..." -ForegroundColor Green
try {
    Set-Content -Path $HOME\Documents\artec.cer -Value $ArtecCert
    Import-Certificate $HOME\Documents\artec.cer -CertStoreLocation 'Cert:\LocalMachine\Root' | Out-Null
}
catch {
    Write-Host "Couldn't import Artec cert" -ForegroundColor Red
    Write-Host $_
    exit
}



# Import certificates for ClickOnce
Write-Host "Importing Sectigo cert..." -ForegroundColor Green
try {
    Set-Content -Path $HOME\Documents\sectigo.cer -Value $SectigoCert
    Import-Certificate $HOME\Documents\sectigo.cer -CertStoreLocation 'Cert:\LocalMachine\Root' | Out-Null
}
catch {
    Write-Host "Couldn't import Sectigo cert" -ForegroundColor Red
    Write-Host $_
    exit
}




# Install Artec Studuino Software
Write-Host "Installing Artec Studuino Software..." -ForegroundColor Green
try {
    # iwr -useb 'https://www.artec-kk.co.jp/studuino/data/software/cdn/en/Studuino.zip' -OutFile $HOME\Documents\Studuino.zip | Out-Null
    New-Item -ItemType Directory -Force -Path $HOME\Documents\studuino | Out-Null
    # Expand-Archive -Path $HOME\Documents\studuino.zip -DestinationPath $HOME\Documents -Force
}
catch {
    Write-Host "Couldn't Install Artec Studuino Software" -ForegroundColor Red
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
    $StuduinoShortcut = "$HOME\Desktop\Studuino.lnk"
    $StuduinoLocation = "$HOME\Documents\studuino\Application Files\ArtecRobotStartUp_1_5_7_1\ArtecRobotStartUp.exe"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Studuino = $WScriptShell.CreateShortcut($StuduinoShortcut)
    $Studuino.TargetPath = $StuduinoLocation
    $Studuino.WorkingDirectory = "$HOME\Documents\studuino\Application Files\ArtecRobotStartUp_1_5_7_1"
    $Studuino.Save()

}
catch {
    Write-Host "Couldn't add shortcuts to Desktop" -ForegroundColor Red
    Write-Host $_
    exit
}
