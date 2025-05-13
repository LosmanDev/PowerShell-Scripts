[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 
New-Item -Type Directory -Path "C:\HWID" 
Set-Location -Path "C:\HWID" $env:Path += ";C:\Program Files\WindowsPowerShell\Scripts" 
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned 
Install-Script -Name Get-WindowsAutopilotInfo 
Get-WindowsAutopilotInfo -OutputFile AutopilotHWID.csv

# Ensure TLS1.2 for script downloads
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Get the machine’s BIOS serial number
$serial = (Get-CimInstance -ClassName Win32_BIOS).SerialNumber.Trim()


# Change location into the new folder
Set-Location -Path "C:\$serial HWID"

# Add Scripts folder to your PATH for this session
$env:Path += ";C:\Program Files\WindowsPowerShell\Scripts"

# Allow running downloaded scripts in this session
Set-ExecutionPolicy -Scope Process -ExecutionPolicy RemoteSigned -Force

# Install and run the Autopilot HWID extractor, outputting inside your serial-named folder
Install-Script -Name Get-WindowsAutopilotInfo -Force
Get-WindowsAutopilotInfo -OutputFile "AutopilotHWID.csv"
