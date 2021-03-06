# Install ssh client and server on Windows 10

# Elevate privledges

if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
  Exit
 }
}

Add-WindowsCapability -Online -Name OpenSSH.Client~~~~0.0.1.0
Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0
Start-Service sshd
Set-Service -Name sshd -StartupType 'Automatic'

# Generate keys

mkdir $HOME/.ssh

New-Item $HOME/.ssh/id_rsa.pub
New-Item $HOME/.ssh/authorized_keys
New-Item C:\ProgramData\ssh\administrators_authorized_keys

Write-Output 'n' | ssh-keygen -f $HOME/.ssh/id_rsa -N '""'

$Local = Get-Content -Path $HOME/.ssh/id_rsa.pub
$Remote =  Get-Content -Path \\SENSEI00\share\authorized_keys
if (!($Local -in $Remote)) {
  Add-Content -Path \\SENSEI00\share\authorized_keys -Value $Local
}

Copy-Item -Path \\SENSEI00\share\authorized_keys -Dest $HOME/.ssh/authorized_keys 
Copy-Item -Path \\SENSEI00\share\authorized_keys -Dest C:\ProgramData\ssh\administrators_authorized_keys

$acl = Get-Acl C:\ProgramData\ssh\administrators_authorized_keys
$acl.SetAccessRuleProtection($true, $false)
$administratorsRule = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","Allow")
$systemRule = New-Object system.security.accesscontrol.filesystemaccessrule("SYSTEM","FullControl","Allow")
$acl.SetAccessRule($administratorsRule)
$acl.SetAccessRule($systemRule)
$acl | Set-Acl
