System Information and Configuration

```bash
# Displays detailed configuration information about the computer, including OS version, memory, and network adapters.
systeminfo

# Quickly checks the version of Windows you are running.
winver

# bios get serialnumber Retrieves the serial number of the computer's BIOS.
wmic
```

System File and Image Repair

```bash
  # Scans all system files to detect and repair corrupted or missing files.
 DISM /Online /Cleanup-Image /CheckHealth

  # Quickly checks if there is any corruption in the system image.
 DISM /Online /Cleanup-Image /ScanHealth

  # Performs a detailed scan of the system image for corruption.
 DISM /Online /Cleanup-Image /RestoreHealth

  # Repairs detected corruption in the system image by downloading and replacing damaged files.
 DISM /Online /Cleanup-Image /StartComponentCleanup

  # Removes obsolete system files and outdated components from the Windows Component Store (WinSxS folder) to reclaim disk space, but it keeps backup components for uninstalling updates.
 Dism.exe /Online /Cleanup-Image /AnalyzeComponentStore

  # To open an item.
 Invoke-Item

  # To relaunch advanced options, then pick Reset this PC.
 shutdown /r /o
```

Disk and File System

```bash
# Checks the file system and disk for errors.
 chkdsk
   Use /f # to fix errors.
   Use /r # to locate bad sectors and recover readable data.

  # Clean system files
 cleanmgr
```

Clear Windows Update

```bash
 net stop wuauserv
 net stop bits
 rd /s /q %windir%\SoftwareDistribution
 net start wuauserv
 net start bits
```

Network Information

```bash
 ipconfig /all # Shows detailed network configuration, including IP address, DNS, and MAC addresses.
 netstat -an # # Displays active connections and listening ports.
 ipconfig # Displays current network settings.
 netsh winsock reset # Resets the Winsock catalog to a clean state (fixes network stack issues).
 netsh int ip reset # Resets TCP/IP settings to default (useful for network troubleshooting).
 ipconfig /release # Releases the current IP address assigned to the device’s network adapter.
 ipconfig /renew # Renews the IP address from the DHCP server.
 ipconfig /flushdns # Flushes the DNS cache to resolve DNS-related issues.
```

Group Policy and Updates

```bash
 gpupdate /force # Forces a refresh of Group Policy settings.
 wmic qfe list # Lists all installed Windows updates (useful for checking patch status).
 gpresult /h # List all the policies applied and security groups in HTML.
 Start-DeviceSync # Force Intune Sync.
 dsregcmd /status # Confirm the Device is Enrolled in Intune.
```

Power Management

```bash
 powercfg /h on/off # Enables or disables hibernation mode.
 powercfg /batteryreport # Generates a detailed battery health report.
 powercfg /energy # Generates an energy efficiency report.
```

Drive Encryption

```bash
 manage-bde -status # Displays the BitLocker encryption status of drives.
 manage-bde C: -off # Decrypts the system drive (turns off BitLocker encryption).
 manage-bde -on C: -used
 manage-bde C: -protectors -add -rp -tpm
 manage-bde -protectors -enable C:
 manage-bde -protectors -get C: > "%UserProfile%\Desktop\BitLocker-Recovery-Key.txt"
```

Advanced WiFi Settings

- Set wireless mode to 802.11n.
- MIMO Power Save Mode set to No SMPS.
- Roaming Aggressiveness set to Highest.

WMI Errors Check

- Open the Event Viewer (eventvwr.msc).
- Navigate to Applications and Services Logs > Microsoft > Windows > WMI-Activity › Operational.

Other Commands

- AD Run dsa.msc to run.
- Network Run ncpa.cpl.
- Group policy gpmc.msc.
- Audio: mmsys.cpl.

- Autopilot

```bash
  - start ms-cxh:localonly
  - start ms-availablenetworks:
  - OOBE\BYPASSNRO
```

Services Commands

```bash
 Stop-Service -Name wuauserv -Force
 Stop-Service -Name bits -Force
 Stop-Service -Name dosvc -Force
 Stop-Service -Name cryptsvc -Force
 Remove-Item -Path "$env:windir\SoftwareDistribution" -Recurse -Force
 Remove-Item -Path "$env:windir\System32\catroot2" -Recurse -Force
 Start-Service -Name cryptsvc
 Start-Service -Name dosvc
 Start-Service -Name bits
 Start-Service -Name wuauserv
```

Force the SCCM Client to Check for Updates

```bash
# Trigger software updates scan cycle
Invoke-WmiMethod -Namespace root\ccm -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000113}"
Invoke-WmiMethod -Namespace root\ccm -Class SMS_Client -Name TriggerSchedule -ArgumentList "{00000000-0000-0000-0000-000000000108}"

C:\Windows\CCM\CCMExec.exe /detectnow
```

Force client actions manually

- In Control Panel > Configuration Manager > Actions, run:\*\*
- Machine Policy Retrieval & Evaluation Cycle\*\*
- Software Updates Scan Cycle\*\*
- Software Updates Deployment Evaluation Cycle\*\*
