# Install ssh client and server on Windows 10

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

$Source = \\SENSEI01\share\authorized_keys
$Local = Get-Content -Path $HOME/.ssh/id_rsa.pub
$Remote =  Get-Content -Path \\SENSEI01\share\authorized_keys
if (!( $Remote -match $Local )) {
  Add-Content -Path \\SENSEI01\share\authorized_keys -Value $From
}

Copy-Item $Source $HOME/.ssh/authorized_keys 
Copy-Item $Source C:\ProgramData\ssh\administrators_authorized_keys

$acl = Get-Acl C:\ProgramData\ssh\administrators_authorized_keys
$acl.SetAccessRuleProtection($true, $false)
$administratorsRule = New-Object system.security.accesscontrol.filesystemaccessrule("Administrators","FullControl","Allow")
$systemRule = New-Object system.security.accesscontrol.filesystemaccessrule("SYSTEM","FullControl","Allow")
$acl.SetAccessRule($administratorsRule)
$acl.SetAccessRule($systemRule)
$acl | Set-Acl
