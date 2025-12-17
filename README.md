# System Information and Configuration

```bash
# Displays detailed configuration information about the computer, including OS version, memory, and network adapters.
systeminfo

# Quickly checks the version of Windows you are running.
winver

# bios get serialnumber Retrieves the serial number of the computer's BIOS.
wmic bios get serialnumber

(Get-CimInstance Win32_ComputerSystemProduct).IdentifyingNumber

# Get system bios
wmic bios get smbiosbiosversion

# ###################################################################################################################
```

# System File and Image Repair

```bash
# ###################################################################################################################
  # System File Checker
 sfc /scannow

  # Scans all system files to detect and repair corrupted or missing files.
 DISM /Online /Cleanup-Image /CheckHealth

  # Quickly checks if there is any corruption in the system image.
 DISM /Online /Cleanup-Image /ScanHealth

  # Performs a detailed scan of the system image for corruption.
 DISM /Online /Cleanup-Image /RestoreHealth

 # 62.3% frozen
 Run:
  sc query wuauserv
  sc query bits
  sc query trustedinstaller

  # Repairs detected corruption in the system image by downloading and replacing damaged files.
 DISM /Online /Cleanup-Image /StartComponentCleanup

  # Removes obsolete system files and outdated components from the Windows Component Store (WinSxS folder) to reclaim disk space, but it keeps backup components for uninstalling updates.
 Dism.exe /Online /Cleanup-Image /AnalyzeComponentStore

  # To relaunch advanced options, then pick Reset this PC.
 shutdown /r /o

 # Commands bundled

Start-Process powershell -ArgumentList '-NoExit', '-Command', @'
Write-Host "Starting: sfc /scannow"
sfc /scannow
Write-Host "Finished: sfc /scannow"

Write-Host "Starting: RestoreHealth"
dism /online /cleanup-image /RestoreHealth
Write-Host "Finished: RestoreHealth"

Write-Host "Starting: StartComponentCleanup"
dism /online /cleanup-image /StartComponentCleanup
Write-Host "Finished: StartComponentCleanup"
'@


# tail recent DISM entries
Get-Content -Path C:\Windows\Logs\DISM\dism.log -Tail 200

# tail recent Component-Based Servicing log
Get-Content -Path C:\Windows\Logs\CBS\CBS.log -Tail 200

# ###################################################################################################################
```

# SYSTEM PROCESSES / UP-TIME / GET TEMPERATURE AND OVERALL CPU USAGE

```powershell

# Uptime
$os = Get-CimInstance Win32_OperatingSystem (New-TimeSpan $os.LastBootUpTime (Get-Date)) | Select-Object Days, Hours, Minutes

# Get system Temperature
Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace root/wmi |
Where-Object CurrentTemperature -gt 0 |
Sort-Object CurrentTemperature -Descending |
Select-Object -First 1 |
Select-Object InstanceName, @{Name = 'Fahrenheit'; Expression = { [math]::Round((($_.CurrentTemperature / 10 - 273.15) * 9 / 5) + 32, 1) } }

# CPU / memory per process
Get-Process | Sort-Object CPU -Descending | Select-Object -First 10 Name, CPU, Id

# Overall CPU usage
Get-Counter '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 5


ms-teams OUTLOOK OneDrive

# ###################################################################################################################
```

# Disk and File System

```bash
# Checks the file system and disk for errors.
 chkdsk
   Use /f # to fix errors.
   Use /r # to locate bad sectors and recover readable data.

  # Clean system files
 cleanmgr

 # Search for a text string in files (Basic)
 find 'error' C:\logs\apps\.log

 # Search for strings in files (More powerful, supports regex)
 findstr /i /s /c:"password" C:\Users\*.txt # Case sensitive, search subdirs, literal string

  # To open an item "C:\".
 Invoke-Item

 # ###################################################################################################################
```

# Clear Windows Update

```bash
 net stop wuauserv
 net stop bits
 rd /s /q %windir%\SoftwareDistribution
 net start wuauserv
 net start bits

 # ###################################################################################################################
```

# Network Information

```bash
 ipconfig /all # Shows detailed network configuration, including IP address, DNS, and MAC addresses.
 netstat -an # # Displays active connections and listening ports.
 ipconfig # Displays current network settings.
 netsh winsock reset # Resets the Winsock catalog to a clean state (fixes network stack issues).
 netsh int ip reset # Resets TCP/IP settings to default (useful for network troubleshooting).
 ipconfig /release # Releases the current IP address assigned to the device’s network adapter.
 ipconfig /renew # Renews the IP address from the DHCP server.
 ipconfig /flushdns # Flushes the DNS cache to resolve DNS-related issues.

 # ###################################################################################################################
```

# Group Policy and Updates

```bash
 gpupdate /force # Forces a refresh of Group Policy settings.
 wmic qfe list # Lists all installed Windows updates (useful for checking patch status).
 gpresult /h # List all the policies applied and security groups in HTML.
 Start-DeviceSync # Force Intune Sync.
 dsregcmd /status # Confirm the Device is Enrolled in Intune.

 # ###################################################################################################################
```

# Power Management

```bash
 powercfg /h on/off # Enables or disables hibernation mode.
 powercfg /batteryreport # Generates a detailed battery health report.
 powercfg /energy # Generates an energy efficiency report.
 powercfg.cpl

 # ###################################################################################################################
```

# Drive Encryption

```bash
 manage-bde -status # Displays the BitLocker encryption status of drives.
 manage-bde C: -off # Decrypts the system drive (turns off BitLocker encryption).
 manage-bde -on C: -used
 manage-bde C: -protectors -add -rp -tpm
 manage-bde -protectors -enable C:
 manage-bde -protectors -get C: > "%UserProfile%\Desktop\BitLocker-Recovery-Key.txt"

 # ###################################################################################################################
```

# Advanced WiFi Settings

- Set wireless mode to `802.11n`
- MIMO Power Save Mode set to `No SMPS`
- Roaming Aggressiveness set to `Highest`

# WMI Errors Check

- Open the Event Viewer `eventvwr.msc`
- Navigate to Applications and Services Logs > Microsoft > Windows > WMI-Activity › Operational.

```bash

dsa.msc # AD Run
ncpa.cpl # Network Run
gpmc.msc # Group policy
mmsys.cpl # Audio
compmgmt.msc #computer management
sysdm.cpl # System props (Add more RAM)
appwiz.cpl # control panel applications

start ms-cxh:localonly # Create a local windows account
start ms-availablenetworks: # Access Network from CMD
start ms-settings:windowsupdate # Access updates
shutdown /r /o /f /t 0 # Windows Recovery Environment (WinRE),

# ###################################################################################################################
```

# Force client actions manually

- In Control Panel > Configuration Manager > Actions, run
- Machine Policy Retrieval & Evaluation Cycle
- Software Updates Deployment Evaluation Cycle
- Software Updates Scan Cycle

# Intune Device State Reset and Forced Compliance Evaluation Steps

```bash

# ###################################################################################################################

# IME reinitializes, retrieves fresh app assignments, re-runs all detection circuitry without delay, and reports compliance faster than the standard Intune polling cycle.

Get-CimInstance Win32_Service -Filter "Name='IntuneManagementExtension'" | Select-Object Name,ProcessId

taskkill /PID 1234 /F

taskkill /IM AgentExecutor.exe /F

sc stop IntuneManagementExtension

Remove-Item "C:\ProgramData\Microsoft\IntuneManagementExtension" -Recurse -Force

restart-computer

start ms-settings:workplace

dsregcmd /refreshprt

$Shell = New-Object -ComObject Shell.Application
$Shell.open("intunemanagementextension://syncapp")


Get-Service | Where-Object { $_.Name -match "csc_umbrellaagent|stAgentSvc|CSFalconService" } | Format-Table Name, Status

# UMBRELLA - 4aade9c2-d76b-4a2e-9caf-58201c341f4d
# NETSKOPE - 2e4c26b7-12f1-4a56-9c22-6ae0d66736ea
# CROWDSTRIKE - f5c225e3-9064-4caf-9c52-0f3a8f375770


# Chrome Bookmarks

%LOCALAPPDATA%\Google\Chrome\User Data\Default

# Edge Bookmarks

%LOCALAPPDATA%\Microsoft\Edge\User Data\Default

# Signatures

%appdata%\Microsoft\Signatures

# ###################################################################################################################
```
