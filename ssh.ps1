# Install ssh client and server on Windows 10

Add-WindowsCapability -Online -Name OpenSSH.Client~~~~0.0.1.0
Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0
Start-Service sshd
Set-Service -Name sshd -StartupType 'Automatic'

# Generate keys

mkdir ~/.ssh
ssh-keygen -f ~/.ssh/id_rsa -N '""'

$From = Get-Content -Path ~/.ssh/id_rsa.pub
Add-Content -Path \\SENSEI01\share\authorized_keys -Value $From
