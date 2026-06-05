### System Information and Configuration

```bash
# Displays detailed configuration information about the computer, including OS version, memory, and network adapters.
systeminfo

# Quickly checks the version of Windows you are running.
winver

################ bios get serialnumber Retrieves the serial number of the computer's BIOS. ###############
wmic bios get serialnumber

(Get-CimInstance Win32_ComputerSystemProduct).IdentifyingNumber

(Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion').DisplayVersion

################ Get system bios ###############
wmic bios get smbiosbiosversion


# ###################################################################################################################
```

### System File and Image Repair

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

  # Repairs detected corruption in the system image by downloading and replacing damaged files.
 DISM /Online /Cleanup-Image /StartComponentCleanup

  # Removes obsolete system files and outdated components from the Windows Component Store (WinSxS folder) to reclaim disk space, but it keeps backup components for uninstalling updates.
 Dism.exe /Online /Cleanup-Image /AnalyzeComponentStore

 ################ Commands bundled ###############

Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile', '-Command', "sfc /scannow; DISM /Online /Cleanup-Image /StartComponentCleanup; DISM /Online /Cleanup-Image /RestoreHealth; shutdown /r /t 60 /c 'Restart Initiated.'"

Start-Process powershell -Verb RunAs -ArgumentList '-NoProfile', '-Command', 'sfc /scannow; DISM /Online /Cleanup-Image /StartComponentCleanup; DISM /Online /Cleanup-Image /RestoreHealth; Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing; $f=New-Object Windows.Forms.Form; $f.Width=350; $f.Height=150; $f.StartPosition=''CenterScreen''; $f.TopMost=$true; $f.Text=''Please save your work''; $f.ControlBox=$false; $l=New-Object Windows.Forms.Label; $l.AutoSize=$true; $l.Font=New-Object Drawing.Font(''Segoe UI'', 14); $l.Top=40; $l.Left=50; $f.Controls.Add($l); $f.Show(); for($i=120; $i -gt 0; $i--){$l.Text=''Restarting in ''+$i+'' seconds...''; $f.Refresh(); Start-Sleep 1}; Restart-Computer -Force'

################ tail recent DISM entries###############
Get-Content -Path C:\Windows\Logs\DISM\dism.log -Tail 200

################ tail recent Component-Based Servicing log ###############
Get-Content -Path C:\Windows\Logs\CBS\CBS.log -Tail 200

shutdown /r /t 60 /c "Restart Initiated."
shutdown /r /t 3600 /c "System maintenance in progress. This device will restart automatically in 60 minutes."


# ###################################################################################################################
```

### SYSTEM PROCESSES / UP-TIME / GET TEMPERATURE AND OVERALL CPU USAGE

```powershell

################ Uptime ###############

$os = Get-CimInstance Win32_OperatingSystem; New-TimeSpan -Start $os.LastBootUpTime -End (Get-Date) | Select Days,Hours,Minutes

################ Get system Temperature ###############

Get-WmiObject MSAcpi_ThermalZoneTemperature -Namespace root/wmi |
Where-Object CurrentTemperature -gt 0 |
Sort-Object CurrentTemperature -Descending |
Select-Object -First 1 |
Select-Object InstanceName, @{Name = 'Fahrenheit'; Expression = { [math]::Round((($_.CurrentTemperature / 10 - 273.15) * 9 / 5) + 32, 1) } }

################ CPU / memory per process ###############

Get-Process | Sort-Object CPU -Descending | Select-Object -First 10 Name, CPU, Id

################ Kill MS sessions ###############

"POWERPNT","EXCEL","WINWORD","OneDrive","OUTLOOK","ms-teams","Teams","msedge","chrome" | ForEach-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue | ForEach-Object { try { Stop-Process -Id $_.Id -Force -ErrorAction Stop; Write-Host "Terminated: $($_.Name) (PID $($_.Id))" } catch { Write-Host "Failed to terminate: $($_.Name) (PID $($_.Id))" } } }

"POWERPNT","EXCEL","WINWORD","OneDrive","OUTLOOK","ms-teams","Teams" | ForEach-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue | ForEach-Object { try { Stop-Process -Id $_.Id -Force -ErrorAction Stop; Write-Host "Terminated: $($_.Name) (PID $($_.Id))" } catch { Write-Host "Failed to terminate: $($_.Name) (PID $($_.Id))" } } }

"OUTLOOK" | ForEach-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue | ForEach-Object { try { Stop-Process -Id $_.Id -Force -ErrorAction Stop; Write-Host "Terminated: $($_.Name) (PID $($_.Id))" } catch { Write-Host "Failed to terminate: $($_.Name) (PID $($_.Id))" } } }


################ Reliability Monitor ###############
perfmon /rel

# This diagnostic script checks the "health" of the PC to find hidden installation blockers: it verifies if the system is unstable (low reliability score), hasn't been rebooted in over a week, is waiting for a reboot (registry locks), has the Windows Installer service stuck, or has failed recent Windows Updates.

Write-Host "DIAGNOSTICS & BLOCKERS" -f Cyan; $s = (Get-CimInstance Win32_ReliabilityStabilityMetrics | select -f 1).SystemStabilityIndex; Write-Host "Stability (1-10): " -NoNewline; if ($s -lt 5) { Write-Host $s -f Red }else { Write-Host $s -f Green }; $d = ((Get-Date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime).Days; Write-Host "Uptime:             $d Days" -f $(if ($d -gt 7) { 'Yellow' }else { 'White' }); $p = @(); if (gp 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' -ea 0) { $p += 'CBS' }; if (gp 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' -ea 0) { $p += 'WU' }; if ((gp 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ea 0).PendingFileRenameOperations) { $p += 'Rename' }; Write-Host "Reboot Pending:     " -NoNewline; if ($p) { Write-Host "YES ($($p -join ','))" -f Red }else { Write-Host "NO" -f Green }; $m = (gps msiexec -ea 0); Write-Host "MSI Exec Busy:      " -NoNewline; if ($m) { Write-Host "YES" -f Yellow }else { Write-Host "NO" -f Green }; Write-Host "`nLast 5 Updates:" -f Cyan; (New-Object -Com Microsoft.Update.Searcher).QueryHistory(0, 5) | % { Write-Host ("[{0}] {1}" -f $_.Date.ToString('MM-dd'), $_.Title.SubString(0, [math]::Min(45, $_.Title.Length))) -f $(if ($_.ResultCode -eq 2) { 'Green' }else { 'Red' }) }

# ###################################################################################################################
```

### Disk and File System

```bash

cd "$env:USERPROFILE\downloads"

# Downloads size check

Get-ChildItem "$env:USERPROFILE\Downloads" -Recurse -File -Force -ErrorAction SilentlyContinue | Measure-Object Length -Sum | Select-Object @{N='Folder';E={"$env:USERPROFILE\Downloads"}}, @{N='SizeMB';E={[math]::Round($_.Sum / 1MB, 2)}}, @{N='SizeGB';E={[math]::Round($_.Sum / 1GB, 2)}}


# Checks the file system and disk for errors.
 chkdsk /f; chkdsk /r
   Use /f # to fix errors.
   Use /r # to locate bad sectors and recover readable data.

  # Clean system files
 cleanmgr

# SysMain pre-loads apps into RAM Disable SysMain
Stop-Service -Name "SysMain" -Force
Set-Service -Name "SysMain" -StartupType Disabled

# Optimize System Drive
Optimize-Volume -DriveLetter C -ReTrim -Verbose

# Empty the Recycle Bin silently
Clear-RecycleBin -Force

# Clear temporary files
Remove-Item -Path "$env:TEMP\*" -Recurse -Force -ErrorAction SilentlyContinue

 # Search for a text string in files
 Select-String -Path  "C:\logs\apps\.log" -Pattern'error'

 # Search for strings in files (More powerful, supports regex)
 findstr /i /s /c:"password" C:\Users\*.txt # Case sensitive, search subdirs, literal string

  # To open an item "C:\".
 Invoke-Item

 # ###################################################################################################################
```

### Clear Windows Update

```bash
 net stop wuauserv
 net stop bits
 rd /s /q %windir%\SoftwareDistribution
 net start wuauserv
 net start bits

 # Windows will immediately create a fresh, empty SoftwareDistribution folder. It forgets all downloaded update files and local history, forcing a full scan against the update server (Intune/WSUS) on the next cycle.

 $s = "wuauserv","bits","cryptsvc","dosvc"; Stop-Service $s -Force; Rename-Item -Path "$env:windir\SoftwareDistribution" -NewName "SoftwareDistribution.bak" -Force; Start-Service $s

 # ###################################################################################################################
```

### Network Information

```bash
 ipconfig /all # Shows detailed network configuration, including IP address, DNS, and MAC addresses.
 netstat -an # # Displays active connections and listening ports.
 ipconfig # Displays current network settings.

 ipconfig /release # Releases the current IP address assigned to the device’s network adapter.
 ipconfig /flushdns # Flushes the DNS cache to resolve DNS-related issues.
 ipconfig /renew # Renews the IP address from the DHCP server.

 #Restart Required
 netsh winsock reset # Resets the Winsock catalog to a clean state (fixes network stack issues).
 netsh int ip reset # Resets TCP/IP settings to default (useful for network troubleshooting).

 Start-Process powershell -Verb RunAs -ArgumentList '-NoExit', '-Command', 'ipconfig /release; ipconfig /flushdns; ipconfig /renew; netsh winsock reset' 

 
 # ###################################################################################################################
```

### Group Policy and Intune Policies

```bash
 gpupdate /force # Forces a refresh of Group Policy settings.
 wmic qfe list # Lists all installed Windows updates (useful for checking patch status).
 gpresult /h # List all the policies applied and security groups in HTML.
 dsregcmd /status # Confirm the Device is Enrolled in Intune.
 dsregcmd /refreshprt #Forces the device to immediately refresh its Primary Refresh Token (PRT) re-establishing authentication state

 # Retrieve the 20 most recent AAD Operational events:
 Get-WinEvent -LogName "Microsoft-Windows-AAD/Operational" -MaxEvents 20 | Select-Object TimeCreated, Id, LevelDisplayName, Message | Format-List

 # Filter specifically for Warning and Error events:
 Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-AAD/Operational'; Level=2,3} -MaxEvents 10 | Select-Object TimeCreated, Id, Message | Format-List

# Force Windows device to immediately check in with Microsoft Intune and sync win32 apps and compliance
 Get-ScheduledTask | ? {$_.TaskName -eq 'PushLaunch'} | % { $_ | Start-ScheduledTask; sleep 2; $_ | Get-ScheduledTaskInfo | select TaskName, Last* }
 $Shell = New-Object -ComObject Shell.Application; $Shell.open("intunemanagementextension://syncapp")
 $Shell = New-Object -ComObject Shell.Application; $Shell.open("intunemanagementextension://synccompliance")

 Get-ScheduledTask | ? {$_.TaskName -eq 'PushLaunch'} | % { $_ | Start-ScheduledTask; sleep 2; $_ | Get-ScheduledTaskInfo | select TaskName, Last* }; $Shell = New-Object -ComObject Shell.Application; $Shell.open("intunemanagementextension://syncapp"); $Shell.open("intunemanagementextension://synccompliance")

 
start ms-cxh:localonly # Create a local windows account
start ms-availablenetworks: # Access Network from CMD
start ms-settings:windowsupdate # Access updates
start ms-settings:workplace # Intune Sync
shutdown /r /o /f /t 0 # Windows Recovery Environment (WinRE),
shutdown /r /fw /t 0 # Motherboard Firmware Interface (UEFI / BIOS)


 # ###################################################################################################################
```

### Power Management

```bash
 powercfg /h on/off # Enables or disables hibernation mode.
 powercfg /batteryreport # Generates a detailed battery health report.
 powercfg /energy # Generates an energy efficiency report.
 powercfg.cpl

# Quick extraction of battery report
 powercfg /batteryreport /output "$env:TEMP\br.html" > $null; $h = Get-Content "$env:TEMP\br.html" -Raw; $o = [regex]::Match($h, '(?s)Since OS install.*?<td class="hms">([^<]+)</td>.*?<div[^>]*>([^<]+)</div>.*?<td class="hms">([^<]+)</td>.*?<div[^>]*>([^<]+)</div>'); [PSCustomObject]@{ 'Design Capacity' = [regex]::Match($h, 'DESIGN CAPACITY</span></td><td>(.*?)\s*mWh').Groups[1].Value; 'Full Charge Capacity' = [regex]::Match($h, 'FULL CHARGE CAPACITY</span></td><td>(.*?)\s*mWh').Groups[1].Value; 'Cycle Count' = [regex]::Match($h, 'CYCLE COUNT</span></td><td>([^<]+)</td>').Groups[1].Value; 'Active (Full Charge)' = $o.Groups[1].Value; 'Standby (Full Charge)' = $o.Groups[2].Value; 'Active (Design Capacity)' = $o.Groups[3].Value; 'Standby (Design Capacity)' = $o.Groups[4].Value }

# Event Viewer Battery reports [524 Critical]
 Get-WinEvent -FilterHashtable @{LogName='System'; ProviderName='Microsoft-Windows-Kernel-Power'; ID=@(524)} | Select-Object TimeCreated, Id, @{Name='Context'; Expression={switch($_.Id){524{'Critical Battery Depletion'}}}}, Message | Format-Table -AutoSize -Wrap

 # ###################################################################################################################
```

### Drive Encryption

```bash
 manage-bde -status # Displays the BitLocker encryption status of drives.
 manage-bde C: -off # Decrypts the system drive (turns off BitLocker encryption).
 manage-bde -on C: -RecoveryPassword

 # ###################################################################################################################
```

### Advanced WiFi Settings

- Set wireless mode to `802.11n`
- MIMO Power Save Mode set to `No SMPS`
- Roaming Aggressiveness set to `Highest`

### WMI Errors Check

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


# ###################################################################################################################
```

### Intune troubleshooting

````powershell

# ###################################################################################################################

# ############### Software Versions installed on system ###############

Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName } | Select-Object DisplayName, DisplayVersion | Sort-Object DisplayName

# ############### Event Viwer softwares installed by recent date###############

Get-WinEvent -FilterHashtable @{LogName='Application';ProviderName='MsiInstaller';ID=1033,11724;StartTime=(Get-Date).AddDays(-5)} | % { $m=$_.Message; [PSCustomObject]@{Time=$_.TimeCreated.ToString('MM/dd/yyyy HH:mm:ss'); Action=if($_.Id -eq 11724){"Removed"}elseif($m -match "status: 0\."){"Installed"}else{"Failed"}; Name=if($m -match "Product(?: Name)?: (.*?)(?:\. Product Version:| --)"){$Matches[1]}else{"Unknown"}; Version=if($m -match "Product Version: ([0-9.]+)"){$Matches[1]}else{"N/A"}} } | ft -AutoSize


################ Scans the Application, System, and Intune MDM logs for "Critical" or "Error" level events from the last 24 hours, printing the most recent 20 failures ###############

$scanevnt=24; $s=(Get-Date).AddHours(-$scanevnt); @{N='Application';L='APP FAILURES'},@{N='System';L='SYSTEM FAILURES'},@{N='Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin';L='INTUNE FAILURES'} | % { Write-Host "`nLOG: $($_.L)" -f Cyan; Write-Host ('='*60); try { Get-WinEvent -FilterHashtable @{LogName=$_.N; Level=1,2; StartTime=$s} -MaxEvents 20 -EA Stop | Sort TimeCreated | % { Write-Host ("[{0}] {1:MM-dd HH:mm} ID={2} Src={3}" -f $_.LogName,$_.TimeCreated,$_.Id,$_.ProviderName) -f Magenta; ($_.Message -split "`r?`n" | ?{$_} | select -f 5) | % { Write-Host "    $_" -f White }; Write-Host ('-'*60) -f DarkGray } } catch { Write-Host "  FAIL REASON: $($_.Exception.Message)" -f Red } }


# ############### System Policies pushed from Intune ###############

gci 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device' -Rec | % { $p=$_.Name; $c=($p -replace '.*\\device\\?','').Split('\')[0]; $_|gp|% { $_.PSObject.Properties | ? Name -notmatch '^PS' | % { [pscustomobject]@{Source='PolicyManager'; Category=$c; Name=$_.Name; Value=$_.Value; Key=$p} } } } | sort Category,Name | ogv -Title 'Policy Manager View'

# ############### Check if 3 agents are running ###############

Get-Service | Where-Object { $_.Name -match "csc_umbrellaagent|stAgentSvc|CSFalconService|IntuneManagementExtension" } | Format-Table Name, Status

# ############### Remove Cisco ###############

"csc_ui.exe", "csc_ui_toast.exe", "vpnui.exe", "vpnagent.exe", "ciscod.exe" | ForEach-Object { taskkill /F /IM $_ /T 2>$null }; @("C:\Program Files (x86)\Cisco\Cisco Secure Client", "C:\ProgramData\Cisco\Cisco Secure Client") | Where-Object { Test-Path $_ } | ForEach-Object { takeown /F $_ /R /D Y | Out-Null; icacls $_ /grant Administrators:F /T /C /Q | Out-Null; Remove-Item -Path $_ -Recurse -Force }

@("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall") | Get-ChildItem | Get-ItemProperty | Where-Object { $_.DisplayName -match "Cisco Secure Client" } | ForEach-Object { Remove-Item -Path $_.PSPath -Recurse -Force }; @("HKLM:\SOFTWARE\Cisco\Cisco Secure Client", "HKLM:\SOFTWARE\WOW6432Node\Cisco\Cisco Secure Client") | Where-Object { Test-Path $_ } | Remove-Item -Recurse -Force

Stop-Service -Name "csc_umbrellaagent" -Force -EA SilentlyContinue; sc.exe delete "csc_umbrellaagent" | Out-Null; Remove-Item -Path "HKLM:\SYSTEM\CurrentControlSet\Services\csc_umbrellaagent" -Recurse -Force -EA SilentlyContinue

################ Scan Intune Extension Logs for specific ID's with error messages. ###############
'4aade9c2-d76b-4a2e-9caf-58201c341f4d' = 'Umbrella'; 
'2e4c26b7-12f1-4a56-9c22-6ae0d66736ea' = 'Netskope';
'f5c225e3-9064-4caf-9c52-0f3a8f375770' = 'CsFalcon'; 
'9df64576-1eff-47b6-886f-00ce74f51b27' = 'Company Portal'

'f74971b0-13e6-42c8-a52d-1f1336e78647','5e811505-aa71-4046-815d-68d931bfbe92' | % { $i=$_; sls $i 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\*.log' -Context 0,20 | % { [pscustomobject]@{ID=$i; File=$_.Filename; Match=$_.Line.Trim(); Context=($_.Context.PostContext | ? {$_ -match 'ExitCode|Error|Fail'} | Out-String).Trim()} } | select -last 5 } | fl


# ############### Checks status for the AppID ###############

$scanApp=@{'f74971b0-13e6-42c8-a52d-1f1336e78647'='Win 24H2 Installer';'5e811505-aa71-4046-815d-68d931bfbe92'='Win 24H2 Feature Update'}; $s=@{1000='Success';2000='Pending';3000='In Progress';4000='Failed'}; gci 'HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps' -Rec | ? {$scanApp.ContainsKey($_.PSChildName)} | % { $r=($_|gp -Name EnforcementStateMessage -ea 0).EnforcementStateMessage; $j=if($r){$r|ConvertFrom-Json}; $c=$j.EnforcementState; $e=$j.ErrorCode; if(!$c){$c=$_.GetValue('EnforcementState');$e=$_.GetValue('LastErrorCode')}; [pscustomobject]@{App=$scanApp.$_.PSChildName; Status=$s[[int]$c]; Err=$e; Time=$_.GetValue('LastUpdatedTimeUtc'); ID=$_.PSChildName} } | ft -a

# ############### Reset Intune Service to re-install AppID ###############

$resetAppInstall=@('f74971b0-13e6-42c8-a52d-1f1336e78647','5e811505-aa71-4046-815d-68d931bfbe92'); $r='HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps'; $resetAppInstall | % { $d=$_; write-host "Scanning $d" -f Cyan; $t=gci $r -Rec -ea 0 | ? {$_.PSChildName -eq $d}; if($t){ $t | % { write-host "Deleting $($_.Name)" -f Yellow; ri $_.PSPath -Rec -Force } } else { write-host "No keys found" -f Gray } }; write-host "Restarting Service..." -f Green; Restart-Service "IntuneManagementExtension" -Force


# ############### Stops service, kills history, hunts down hidden GRS keys for both apps, and restarts service ###############

Stop-Service "IntuneManagementExtension" -Force -ea 0; $t=@('f74971b0-13e6-42c8-a52d-1f1336e78647','5e811505-aa71-4046-815d-68d931bfbe92'); $r="HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps"; $t | % { $id=$_; Write-Host "Cleaning $id" -f Cyan; gci $r -Rec -ea 0 | ? {$_.PSChildName -eq $id} | ri -Rec -Force; gci $r | % { gci "$($_.PSPath)\GRS" -ea 0 | ? {$_.PSChildName -eq $id} | ri -Rec -Force } }; Start-Service "IntuneManagementExtension"

# ############### App ID Log Error Tracker ###############

sls 'f74971b0-13e6-42c8-a52d-1f1336e78647|5e811505-aa71-4046-815d-68d931bfbe92' 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\*.log' -Context 0,15 | % { $c=$_.Context.PostContext | ? {$_ -match 'ExitCode|Error|Fail|GRS'} | Out-String; if($c){ [pscustomobject]@{Log=$_.Filename; Match=$_.Line.Trim().Substring(0, [math]::Min(80,$_.Line.Length)); Context=$c.Trim()} } } | fl

# ############### Force Run (Installer Only) ###############

$e="C:\Program Files (x86)\Microsoft Intune Management Extension\AgentExecutor.exe"; if(Test-Path $e){ & $e -configFile "C:\Program Files (x86)\Microsoft Intune Management Extension\AgentExecutorConfig.xml" -appId 'f74971b0-13e6-42c8-a52d-1f1336e78647' -operation 1 } else { echo "AgentExecutor not found" }


# ###################################################################################################################

# IME reinitializes, retrieves fresh app assignments, re-runs all detection circuitry without delay, and reports compliance faster than the standard Intune polling cycle.

Function Reset-Intune { Write-Host ">>> RESETTING INTUNE AGENT <<<"; Stop-Service "IntuneManagementExtension" -Force -ErrorAction SilentlyContinue; "AgentExecutor", "Microsoft.Management.Services.IntuneWindowsAgent" | ForEach-Object { Get-Process $_ -ErrorAction SilentlyContinue | Stop-Process -Force }; Remove-Item "C:\ProgramData\Microsoft\IntuneManagementExtension" -Recurse -Force -ErrorAction SilentlyContinue; dsregcmd /refreshprt; Start-Service "IntuneManagementExtension"; Get-ScheduledTask | Where-Object { $_.TaskName -eq 'PushLaunch' } | Start-ScheduledTask; Write-Host ">>> DONE. Sync Triggered. <<<" }; Reset-Intune

# ###################################################################################################################

# ############### Chrome Bookmarks ###############

%LOCALAPPDATA%\Google\Chrome\User Data\Default

# ############### Edge Bookmarks ###############

%LOCALAPPDATA%\Microsoft\Edge\User Data\Default

# ############### Signatures ###############

%appdata%\Microsoft\Signatures

```bash

################ Local Outlook signatures→ OneDrive backup ###############
$src="$env:APPDATA\Microsoft\Signatures";$dst="$env:USERPROFILE\OneDrive - BeiGene\Desktop\Signatures";if(!(Test-Path $dst)){New-Item $dst -ItemType Directory|Out-Null};Copy-Item "$src\*" $dst -Recurse -Force

################ Reverse: OneDrive → local Outlook signatures ###############
$src="$env:USERPROFILE\OneDrive - BeiGene\Desktop\Signatures";$dst="$env:APPDATA\Microsoft\Signatures";if(!(Test-Path $dst)){New-Item $dst -ItemType Directory|Out-Null};Move-Item "$src\*" $dst -Recurse -Force


################ Notepad ###############
Add-AppxPackage -Path "C:\Users\liban.osman\OneDrive - BeiGene\Documents\Softwares\Notepad\Microsoft.VCLibs.140.00_14.0.33519.0_x64__8wekyb3d8bbwe.Appx"
Add-AppxPackage -Path "C:\Users\liban.osman\OneDrive - BeiGene\Documents\Softwares\Notepad\Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64__8wekyb3d8bbwe.Appx"
Add-AppxPackage -Path "C:\Users\liban.osman\OneDrive - BeiGene\Documents\Softwares\Notepad\Microsoft.UI.Xaml.2.7_7.2409.9001.0_x64__8wekyb3d8bbwe.Appx"
Add-AppxPackage -Path "C:\Users\liban.osman\OneDrive - BeiGene\Documents\Softwares\Notepad\Microsoft.WindowsNotepad_11.2604.5.0_neutral_~_8wekyb3d8bbwe.Msixbundle"

$UWPPath = (Get-AppxPackage Microsoft.WindowsNotepad).InstallLocation + "\Notepad\Notepad.exe"
& $UWPPath

# 1. Purge the broken user-context installation and custom registry routing
Get-AppxPackage *WindowsNotepad* | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\notepad.exe" -Force -ErrorAction SilentlyContinue

# 2. Define absolute payload paths
$Dir = "C:\Users\liban.osman\OneDrive - BeiGene\Documents\Softwares\Notepad"
$Bundle = "$Dir\Microsoft.WindowsNotepad_11.2604.5.0_neutral_~_8wekyb3d8bbwe.Msixbundle"
$Dep1 = "$Dir\Microsoft.VCLibs.140.00_14.0.33519.0_x64__8wekyb3d8bbwe.Appx"
$Dep2 = "$Dir\Microsoft.VCLibs.140.00.UWPDesktop_14.0.33728.0_x64__8wekyb3d8bbwe.Appx"
$Dep3 = "$Dir\Microsoft.UI.Xaml.2.7_7.2409.9001.0_x64__8wekyb3d8bbwe.Appx"

Add-AppxProvisionedPackage -Online -PackagePath $Bundle -DependencyPackagePath $Dep1, $Dep2, $Dep3 -SkipLicense

# 4. Force OS registration for the current active session
$Manifest = (Get-ChildItem -Path "C:\Program Files\WindowsApps" -Filter "Microsoft.WindowsNotepad_*_x64__8wekyb3d8bbwe" -Directory | Select-Object -First 1).FullName + "\AppxManifest.xml"
Add-AppxPackage -Register $Manifest -DisableDevelopmentMode


```

# ############### Kyocera Logs ###############

%APPDATA%\Kyocera Cloud Print and Scan - Print status\logs\errors

# ############### Outlook Monthly Channel ###############

Set-Location "C:\Program Files\Common Files\Microsoft Shared\ClickToRun"
.\OfficeC2RClient.exe /changesetting Channel=MonthlyEnterprise
.\OfficeC2RClient.exe /update user


# ############### Outlook Legacy Room Finder ###############

$s=(Get-CimInstance Win32_UserProfile | ? LocalPath -match "mike.kao").SID; $p1="Registry::HKEY_USERS\$s\SOFTWARE\Policies\Microsoft\Office\16.0\Outlook\Options\Calendar"; $p2="Registry::HKEY_USERS\$s\SOFTWARE\Microsoft\Office\16.0\Outlook\Preferences"; if(!(Test-Path $p1)){New-Item $p1 -Force | Out-Null}; New-ItemProperty -Path $p1 -Name "ShowLegacyRoomFinder" -Value 1 -PropertyType DWord -Force | Out-Null; if(!(Test-Path $p2)){New-Item $p2 -Force | Out-Null}; New-ItemProperty -Path $p2 -Name "RoomFinderForceWebView" -Value 0 -PropertyType DWord -Force | Out-Null


# ############### Intune Log collection ###############

md C:\temp\odc
cd c:\temp\odc
wget https://aka.ms/intunePS1 -outfile IntuneODCStandAlone.ps1
wget https://aka.ms/intuneXML -outfile Intune.xml
Set-ExecutionPolicy Bypass
.\IntuneODCStandAlone.ps1

# ############### Edge Fix ###############

Remove-Item -Path "C:\Program Files (x86)\Microsoft\Edge\Application\149.*" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "C:\Program Files (x86)\Microsoft\EdgeUpdate\Download\*" -Recurse -Force -ErrorAction SilentlyContinue

$ClientKeys = @(
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}",
    "HKLM:\SOFTWARE\Microsoft\EdgeUpdate\Clients\{56EB18F8-B008-4CBD-B6D2-8C97FE7E9062}"
)
foreach ($Key in $ClientKeys) { Remove-Item -Path $Key -Recurse -Force -ErrorAction SilentlyContinue }

Start-Process msiexec.exe -ArgumentList '/i "C:\Users\ray.nunez\Downloads\MicrosoftEdgeEnterpriseX64.msi" /qn /norestart /L*V "C:\edge_install.log"' -Wait -NoNewWindow

https://www.microsoft.com/en-us/edge/business/download?form=MA13FJ

````

### AUTOMATED SOFTWARE INSTALLS

```powershell

# Incognito Edge
Start-Process msedge -ArgumentList "-inprivate"

# ############### Lenovo System Update ###############

Add-Type -A System.Windows.Forms,System.Drawing; function n($m){$b=New-Object System.Windows.Forms.NotifyIcon;$b.Icon=[System.Drawing.SystemIcons]::Information;$b.Visible=$true;$b.ShowBalloonTip(5000,'Software Install',$m,[System.Windows.Forms.ToolTipIcon]::Info);sleep -m 600;$b.Dispose()}; $u='https://download.lenovo.com/pccbbs/thinkvantage_en/system_update_5.08.03.59.exe'; $p="$env:TEMP\lenovo_update.exe"; n 'Downloading Lenovo System Update...'; (New-Object System.Net.WebClient).DownloadFile($u, $p); n 'Installing Lenovo System Update...'; start $p -Arg '/VERYSILENT /NORESTART' -Wait; ri $p -Force; n 'Lenovo System Update Installed Successfully'; sleep 2

################ Surface Laptop 5 ############### https://www.microsoft.com/en-us/download/details.aspx?id=104679

Add-Type -A System.Windows.Forms,System.Drawing; function n($m){$b=New-Object System.Windows.Forms.NotifyIcon;$b.Icon=[System.Drawing.SystemIcons]::Information;$b.Visible=$true;$b.ShowBalloonTip(5000,'Software Install',$m,[System.Windows.Forms.ToolTipIcon]::Info);sleep -m 600;$b.Dispose()}; $u='https://download.microsoft.com/download/68992368-8d70-4231-a9e4-23dfaede832b/SurfaceLaptop5_Win11_22631_26.043.30647.0.msi'; $p="$env:TEMP\surface5_update.msi"; n 'Downloading Surface Laptop 5 Drivers...'; (New-Object System.Net.WebClient).DownloadFile($u, $p); n 'Installing Surface Laptop 5 Drivers...'; start msiexec -Arg "/i `"$p`" /qn /norestart" -Wait; ri $p -Force; n 'Surface Laptop 5 Drivers Installed Successfully'; sleep 2


################ Surface Laptop 6 ############### https://www.microsoft.com/en-us/download/details.aspx?id=105946

Add-Type -A System.Windows.Forms,System.Drawing; function n($m){$b=New-Object System.Windows.Forms.NotifyIcon;$b.Icon=[System.Drawing.SystemIcons]::Information;$b.Visible=$true;$b.ShowBalloonTip(5000,'Software Install',$m,[System.Windows.Forms.ToolTipIcon]::Info);sleep -m 600;$b.Dispose()}; $u='https://download.microsoft.com/download/a53facb0-c939-4302-a0d3-53aa18217230/SurfaceLaptop6forBusiness_Win11_22631_26.051.6840.0.msi'; $p="$env:TEMP\surface6_update.msi"; n 'Downloading Surface Laptop 6 Drivers...'; (New-Object System.Net.WebClient).DownloadFile($u, $p); n 'Installing Surface Laptop 6 Drivers...'; start msiexec -Arg "/i `"$p`" /qn /norestart" -Wait; ri $p -Force; n 'Surface Laptop 6 Drivers Installed Successfully'; sleep 2

################ Surface Laptop 7 ############### https://www.microsoft.com/en-us/download/details.aspx?id=108014

Add-Type -A System.Windows.Forms,System.Drawing; function n($m){$b=New-Object System.Windows.Forms.NotifyIcon;$b.Icon=[System.Drawing.SystemIcons]::Information;$b.Visible=$true;$b.ShowBalloonTip(5000,'Software Install',$m,[System.Windows.Forms.ToolTipIcon]::Info);sleep -m 600;$b.Dispose()}; $u='https://download.microsoft.com/download/1543bd80-9cae-498d-8b0f-9841e4d7b2a8/SurfaceLaptop7withIntel_Win11_22631_26.043.33704.0.msi'; $p="$env:TEMP\surface7_update.msi"; n 'Downloading Surface Laptop 7 Drivers...'; (New-Object System.Net.WebClient).DownloadFile($u, $p); n 'Installing Surface Laptop 7 Drivers...'; start msiexec -Arg "/i `"$p`" /qn /norestart" -Wait; ri $p -Force; n 'Surface Laptop 7 Drivers Installed Successfully'; sleep 2

################ Chrome ###############

Add-Type -A System.Windows.Forms,System.Drawing; function n($m){$b=New-Object System.Windows.Forms.NotifyIcon;$b.Icon=[System.Drawing.SystemIcons]::Information;$b.Visible=$true;$b.ShowBalloonTip(5000,'Software Install',$m,[System.Windows.Forms.ToolTipIcon]::Info);sleep -m 600;$b.Dispose()}; $u='https://dl.google.com/chrome/install/latest/chrome_installer.exe'; $p="$env:TEMP\chrome_installer.exe"; n 'Downloading Google Chrome...'; (New-Object System.Net.WebClient).DownloadFile($u, $p); n 'Installing Google Chrome...'; start $p -Arg '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Wait; ri $p -Force; n 'Google Chrome Installed Successfully'; sleep 2

```

# Zsh/Bash Commands

```bash

# ###################################################################################################################
/System/Volumes/Data/Library/SystemExtensions/> 
                     
102G    /Library/SystemExtensions/.staging
systemextensionsctl list | Select-String "com.crowdstrike.falcon"

X9E956P446 com.crowdstrike.falcon.Agent (7.36/208.07) Falcon Sensor [activated enabled]
X9E956P446 com.crowdstrike.falcon.Agent (7.35/207.04)  Falcon Sensor [terminated waiting to uninstall on reboot]

df -h /                                                        
Filesystem        Size    Used   Avail Capacity iused ifree %iused  Mounted on
/dev/disk3s1s1   228Gi    12Gi    34Gi    26%    458k  359M    0%   /


###################################################################################################################

# print the top 20 largest ghost files
sudo lsof +L1 | awk '{print $1, $2, $3, $7, $10}' | sort -nk 4 | tail -n 20

sudo du -sh /System/Volumes/Data/Library/* 2>/dev/null | sort -rh | head -n 10

sudo bash -c "du -sh /Library/SystemExtensions/.* 2>/dev/null | sort -rh | head -n 5"

sudo find /Library/SystemExtensions/.staging -name "*.systemextension" | head -n 20

# Find specific bloated temp folder
sudo du -sh /private/var/folders/*/* | sort -rh | head -n 10

# Size of virtual memory
ls -lh /private/var/vm

# See exactly how much space snapshots are taking by opening Terminal and running:
tmutil listlocalsnapshots /

# Identify large directories within the hidden Library folder:
du -sh ~/Library/* | sort -rh | head -n 15

# To reset Spotlight
sudo mdutil -E /

# Check System Temp Bloat
sudo du -sh /private/var/folders/* | sort -rh


 ############################## JAMF / MDM ############################## 

# Jamf Log File Diagnostics
cat /var/log/jamf.log

# Retrieves native Apple MDM framework logs generated within the last 60 minutes.
log show --predicate 'subsystem == "com.apple.MDM"' --last 1h 

# Streams real-time execution logs for the Jamf binary.
log stream --predicate 'process == "jamf"' 

# Forces execution of pending policies scoped to the device.
sudo jamf policy 

# Initiates an inventory collection and submits data to the Jamf Pro server.
sudo jamf recon 

# Re-applies the management framework and MDM profile.
sudo jamf manage 

# Prompts for user-level MDM profile installation if missing.
sudo jamf mdm -userLevelMdm 

############################### System & Hardware Auditing ############################## 

# Outputs macOS product version, build version, and product name.
sw_vers: 

# Outputs hardware UUID, serial number, processor architecture, and memory configuration.
system_profiler SPHardwareDataType 

# Displays system uptime and load averages.
uptime 

# Outputs current FileVault 2 encryption state.
fdesetup status 

# Verifies if a specific user possesses a SecureToken (required for FileVault decryption).
sysadminctl -secureTokenStatus <username> 

# Displays the status of System Integrity Protection (SIP).
csrutil status 

############################## Network Diagnostics ############################## 

# Lists all registered network interfaces.
networksetup -listallnetworkservices 

# Retrieves the MAC address for the specified network service.
networksetup -getmacaddress "Wi-Fi" 

# Executes standard ICMP echo requests with a defined packet count.
ping -c 4 <hostname> 

Directory & Account Operations

# Lists all local user accounts.
dscl . -list /Users 

# Outputs all directory attributes (UID, GID, home directory path) for a specified account.
dscl . -read /Users/<username> 

# Grants local administrator privileges to a standard user.
sudo dseditgroup -o edit -a <username> -t user admin 

```
