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

Start-Process powershell -Verb RunAs -ArgumentList '-NoExit', '-Command', 'Write-Host "Starting: sfc /scannow"; sfc /scannow; Write-Host "Finished: sfc /scannow"; Write-Host "Starting: RestoreHealth"; dism /online /cleanup-image /RestoreHealth; Write-Host "Finished: RestoreHealth"; Write-Host "Starting: StartComponentCleanup"; dism /online /cleanup-image /StartComponentCleanup; Write-Host "Finished: StartComponentCleanup"'

################ tail recent DISM entries###############
Get-Content -Path C:\Windows\Logs\DISM\dism.log -Tail 200

################ tail recent Component-Based Servicing log ###############
Get-Content -Path C:\Windows\Logs\CBS\CBS.log -Tail 200

shutdown /r /t 60 /c "Restart Initiated."
shutdown /r /t 3600 /c "System maintenance in progress. This device will restart automatically in 60 minutes.

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

################ Overall CPU usage ###############

Get-Counter '\Processor(_Total)\% Processor Time' -SampleInterval 1 -MaxSamples 5

################ Kill MS sessions ###############

"POWERPNT","EXCEL","WINWORD","OneDrive","OUTLOOK","ms-teams","Teams","msedge","chrome" | ForEach-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue | ForEach-Object { try { Stop-Process -Id $_.Id -Force -ErrorAction Stop; Write-Host "Terminated: $($_.Name) (PID $($_.Id))" } catch { Write-Host "Failed to terminate: $($_.Name) (PID $($_.Id))" } } }


################ Reliability Monitor ###############
perfmon /rel

# This diagnostic script checks the "health" of the PC to find hidden installation blockers: it verifies if the system is unstable (low reliability score), hasn't been rebooted in over a week, is waiting for a reboot (registry locks), has the Windows Installer service stuck, or has failed recent Windows Updates.

Write-Host "DIAGNOSTICS & BLOCKERS" -f Cyan; $s = (Get-CimInstance Win32_ReliabilityStabilityMetrics | select -f 1).SystemStabilityIndex; Write-Host "Stability (1-10): " -NoNewline; if ($s -lt 5) { Write-Host $s -f Red }else { Write-Host $s -f Green }; $d = ((Get-Date) - (Get-CimInstance Win32_OperatingSystem).LastBootUpTime).Days; Write-Host "Uptime:             $d Days" -f $(if ($d -gt 7) { 'Yellow' }else { 'White' }); $p = @(); if (gp 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending' -ea 0) { $p += 'CBS' }; if (gp 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired' -ea 0) { $p += 'WU' }; if ((gp 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -ea 0).PendingFileRenameOperations) { $p += 'Rename' }; Write-Host "Reboot Pending:     " -NoNewline; if ($p) { Write-Host "YES ($($p -join ','))" -f Red }else { Write-Host "NO" -f Green }; $m = (gps msiexec -ea 0); Write-Host "MSI Exec Busy:      " -NoNewline; if ($m) { Write-Host "YES" -f Yellow }else { Write-Host "NO" -f Green }; Write-Host "`nLast 5 Updates:" -f Cyan; (New-Object -Com Microsoft.Update.Searcher).QueryHistory(0, 5) | % { Write-Host ("[{0}] {1}" -f $_.Date.ToString('MM-dd'), $_.Title.SubString(0, [math]::Min(45, $_.Title.Length))) -f $(if ($_.ResultCode -eq 2) { 'Green' }else { 'Red' }) }


# Scans the Application, System, and Intune MDM logs for "Critical" or "Error" level events from the last 24 hours, printing the most recent 20 failures ###############

$h=24; $s=(Get-Date).AddHours(-$h); @{N='Application';L='APP FAILURES'},@{N='System';L='SYSTEM FAILURES'},@{N='Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin';L='INTUNE FAILURES'} | % { Write-Host "`nLOG: $($_.L)" -f Cyan; Write-Host ('='*60); try { Get-WinEvent -FilterHashtable @{LogName=$_.N; Level=1,2; StartTime=$s} -MaxEvents 20 -EA Stop | Sort TimeCreated | % { Write-Host "[{0}] {1:MM-dd HH:mm} ID={2} Src={3}" -f $_.LogName,$_.TimeCreated,$_.Id,$_.ProviderName -f Magenta; ($_.Message -split "`r?`n" | ?{$_} | select -f 5) | % { Write-Host "    $_" -f White }; Write-Host ('-'*60) -f DarkGray } } catch { Write-Host "  No errors found or log unavailable." -f Green } }

# ###################################################################################################################
```

### Disk and File System

```bash
# Checks the file system and disk for errors.
 chkdsk /f; chkdsk /r
   Use /f # to fix errors.
   Use /r # to locate bad sectors and recover readable data.

  # Clean system files
 cleanmgr

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


 # ###################################################################################################################
```

### Group Policy and Intune Policies

```bash
 gpupdate /force # Forces a refresh of Group Policy settings.
 wmic qfe list # Lists all installed Windows updates (useful for checking patch status).
 gpresult /h # List all the policies applied and security groups in HTML.
 dsregcmd /status # Confirm the Device is Enrolled in Intune.
 dsregcmd /refreshprt #Forces the device to immediately refresh its Primary Refresh Token (PRT) re-establishing authentication state

# Force Windows device to immediately check in with Microsoft Intune and sync win32 apps and compliance
 Get-ScheduledTask | ? {$_.TaskName -eq 'PushLaunch'} | % { $_ | Start-ScheduledTask; sleep 2; $_ | Get-ScheduledTaskInfo | select TaskName, Last* }
 $Shell = New-Object -ComObject Shell.Application; $Shell.open("intunemanagementextension://syncapp")
 $Shell = New-Object -ComObject Shell.Application; $Shell.open("intunemanagementextension://synccompliance")
 start "intunemanagementextension://syncapp"
 start "intunemanagementextension://synccompliance"

$t=(Get-Date).AddMinutes(-30); $events=@(); $events+=Get-WinEvent -LogName "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Operational" -EA 0|?{$_.TimeCreated -ge $t -and $_.Id -in 72,73}|Select @{N='Time';E={$_.TimeCreated}},@{N='Source';E={'MDM'}},@{N='Detail';E={"Session $(if($_.Id -eq 72){'Start'}else{'End'})"}}; $events+=Get-WinEvent -LogName "Security" -EA 0|?{$_.TimeCreated -ge $t -and $_.Id -eq 4688 -and $_.Properties[5].Value -match 'AgentExecutor.exe'}|Select @{N='Time';E={$_.TimeCreated}},@{N='Source';E={'PROC'}},@{N='Detail';E={$_.Properties[8].Value}}; $events+=Get-Content "$env:ProgramData\Microsoft\IntuneManagementExtension\Logs\IntuneManagementExtension.log" -Tail 600 -EA 0|%{if($_ -match 'date="([^"]+)".*time="([^"]+)"'){$d=[datetime]"$($Matches[1]) $($Matches[2])"; if($d -ge $t){[pscustomobject]@{Time=$d;Source='IME';Detail=$_.Replace('<![LOG[','').Split(']')[0]}}}}; $events|Sort Time|FT -AutoSize
 # ###################################################################################################################
```

### Power Management

```bash
 powercfg /h on/off # Enables or disables hibernation mode.
 powercfg /batteryreport # Generates a detailed battery health report.
 powercfg /energy # Generates an energy efficiency report.
 powercfg.cpl

 # ###################################################################################################################
```

### Drive Encryption

```bash
 manage-bde -status # Displays the BitLocker encryption status of drives.
 manage-bde C: -off # Decrypts the system drive (turns off BitLocker encryption).
 manage-bde -on C: -used
 manage-bde C: -protectors -add -rp -tpm
 manage-bde -protectors -enable C:
 manage-bde -protectors -get C: > "%UserProfile%\Desktop\BitLocker-Recovery-Key.txt"

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

start ms-cxh:localonly # Create a local windows account
start ms-availablenetworks: # Access Network from CMD
start ms-settings:windowsupdate # Access updates
start ms-settings:workplace #Intune Sync
shutdown /r /o /f /t 0 # Windows Recovery Environment (WinRE),


# ###################################################################################################################
```

### Intune troubleshooting

````powershell

# ###################################################################################################################

# ############### System Policies pushed from Intune ###############

gci 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device' -Rec | % { $p=$_.Name; $c=($p -replace '.*\\device\\?','').Split('\')[0]; $_|gp|% { $_.PSObject.Properties | ? Name -notmatch '^PS' | % { [pscustomobject]@{Source='PolicyManager'; Category=$c; Name=$_.Name; Value=$_.Value; Key=$p} } } } | sort Category,Name | ogv -Title 'Policy Manager View'

Get-Service | Where-Object { $_.Name -match "csc_umbrellaagent|stAgentSvc|CSFalconService|IntuneManagementExtension" } | Format-Table Name, Status

Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Applets\Regedit" -Name "LastKey" -Value "Computer\HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps"; Start-Process regedit

'4aade9c2-d76b-4a2e-9caf-58201c341f4d' = 'Umbrella'; '2e4c26b7-12f1-4a56-9c22-6ae0d66736ea' = 'Netskope'; 'f5c225e3-9064-4caf-9c52-0f3a8f375770' = 'CsFalcon'; '9df64576-1eff-47b6-886f-00ce74f51b27' = 'Company Portal'

################ Scan Intune Extension Logs for specific ID's with error messages. ###############

'f74971b0-13e6-42c8-a52d-1f1336e78647','5e811505-aa71-4046-815d-68d931bfbe92' | % { $i=$_; sls $i 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\*.log' -Context 0,20 | % { [pscustomobject]@{ID=$i; File=$_.Filename; Match=$_.Line.Trim(); Context=($_.Context.PostContext | ? {$_ -match 'ExitCode|Error|Fail'} | Out-String).Trim()} } | select -last 5 } | fl


# ############### Checks status for the AppID ###############

$m=@{'f74971b0-13e6-42c8-a52d-1f1336e78647'='Win 24H2 Installer';'5e811505-aa71-4046-815d-68d931bfbe92'='Win 24H2 Feature Update'}; $s=@{1000='Success';2000='Pending';3000='In Progress';4000='Failed'}; gci 'HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps' -Rec | ? {$m.ContainsKey($_.PSChildName)} | % { $r=($_|gp -Name EnforcementStateMessage -ea 0).EnforcementStateMessage; $j=if($r){$r|ConvertFrom-Json}; $c=$j.EnforcementState; $e=$j.ErrorCode; if(!$c){$c=$_.GetValue('EnforcementState');$e=$_.GetValue('LastErrorCode')}; [pscustomobject]@{App=$m.$_.PSChildName; Status=$s[[int]$c]; Err=$e; Time=$_.GetValue('LastUpdatedTimeUtc'); ID=$_.PSChildName} } | ft -a

$i=@('f74971b0-13e6-42c8-a52d-1f1336e78647','5e811505-aa71-4046-815d-68d931bfbe92'); $r='HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps'; $i | % { $d=$_; write-host "Scanning $d" -f Cyan; $t=gci $r -Rec -ea 0 | ? {$_.PSChildName -eq $d}; if($t){ $t | % { write-host "Deleting $($_.Name)" -f Yellow; ri $_.PSPath -Rec -Force } } else { write-host "No keys found" -f Gray } }; write-host "Restarting Service..." -f Green; Restart-Service "IntuneManagementExtension" -Force


# ############### Stops service, kills history, hunts down hidden GRS keys for both apps, and restarts service ###############

Stop-Service "IntuneManagementExtension" -Force -ea 0; $t=@('f74971b0-13e6-42c8-a52d-1f1336e78647','5e811505-aa71-4046-815d-68d931bfbe92'); $r="HKLM:\SOFTWARE\Microsoft\IntuneManagementExtension\Win32Apps"; $t | % { $id=$_; Write-Host "Cleaning $id" -f Cyan; gci $r -Rec -ea 0 | ? {$_.PSChildName -eq $id} | ri -Rec -Force; gci $r | % { gci "$($_.PSPath)\GRS" -ea 0 | ? {$_.PSChildName -eq $id} | ri -Rec -Force } }; Start-Service "IntuneManagementExtension"

# ############### Log Error Tracker ###############

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

```

# ############### Kyocera Logs ###############

%APPDATA%\Kyocera Cloud Print and Scan - Print status\logs\errors

# ############### Outlook Monthly Channel ###############

Set-Location "C:\Program Files\Common Files\Microsoft Shared\ClickToRun"
.\OfficeC2RClient.exe /changesetting Channel=MonthlyEnterprise
.\OfficeC2RClient.exe /update user

````

### AUTOMATED SOFTWARE INSTALLS

```bash

# ############### Lenovo System Update ###############

Add-Type -A System.Windows.Forms,System.Drawing; function n($m){$b=New-Object System.Windows.Forms.NotifyIcon;$b.Icon=[System.Drawing.SystemIcons]::Information;$b.Visible=$true;$b.ShowBalloonTip(5000,'Software Install',$m,[System.Windows.Forms.ToolTipIcon]::Info);sleep -m 600;$b.Dispose()}; $u='https://download.lenovo.com/pccbbs/thinkvantage_en/system_update_5.08.03.59.exe'; $p="$env:TEMP\lenovo_update.exe"; n 'Downloading Lenovo System Update...'; (New-Object System.Net.WebClient).DownloadFile($u, $p); n 'Installing Lenovo System Update...'; start $p -Arg '/VERYSILENT /NORESTART' -Wait; ri $p -Force; n 'Lenovo System Update Installed Successfully'; sleep 2

################ Surface Laptop 5 ###############

Add-Type -A System.Windows.Forms,System.Drawing; function n($m){$b=New-Object System.Windows.Forms.NotifyIcon;$b.Icon=[System.Drawing.SystemIcons]::Information;$b.Visible=$true;$b.ShowBalloonTip(5000,'Software Install',$m,[System.Windows.Forms.ToolTipIcon]::Info);sleep -m 600;$b.Dispose()}; $u='https://download.microsoft.com/download/68992368-8d70-4231-a9e4-23dfaede832b/SurfaceLaptop5_Win11_22631_25.120.4884.0.msi'; $p="$env:TEMP\surface5_update.msi"; n 'Downloading Surface Laptop 5 Drivers...'; (New-Object System.Net.WebClient).DownloadFile($u, $p); n 'Installing Surface Laptop 5 Drivers...'; start msiexec -Arg "/i `"$p`" /qn /norestart" -Wait; ri $p -Force; n 'Surface Laptop 5 Drivers Installed Successfully'; sleep 2

################ Surface Laptop 6 ###############

Add-Type -A System.Windows.Forms,System.Drawing; function n($m){$b=New-Object System.Windows.Forms.NotifyIcon;$b.Icon=[System.Drawing.SystemIcons]::Information;$b.Visible=$true;$b.ShowBalloonTip(5000,'Software Install',$m,[System.Windows.Forms.ToolTipIcon]::Info);sleep -m 600;$b.Dispose()}; $u='https://download.microsoft.com/download/a53facb0-c939-4302-a0d3-53aa18217230/SurfaceLaptop6forBusiness_Win11_22631_25.120.480.0.msi'; $p="$env:TEMP\surface6_update.msi"; n 'Downloading Surface Laptop 6 Drivers...'; (New-Object System.Net.WebClient).DownloadFile($u, $p); n 'Installing Surface Laptop 6 Drivers...'; start msiexec -Arg "/i `"$p`" /qn /norestart" -Wait; ri $p -Force; n 'Surface Laptop 6 Drivers Installed Successfully'; sleep 2

################ Surface Laptop 7 ###############

Add-Type -A System.Windows.Forms,System.Drawing; function n($m){$b=New-Object System.Windows.Forms.NotifyIcon;$b.Icon=[System.Drawing.SystemIcons]::Information;$b.Visible=$true;$b.ShowBalloonTip(5000,'Software Install',$m,[System.Windows.Forms.ToolTipIcon]::Info);sleep -m 600;$b.Dispose()}; $u='https://download.microsoft.com/download/1543bd80-9cae-498d-8b0f-9841e4d7b2a8/SurfaceLaptop7withIntel_Win11_22631_25.122.21761.0.msi'; $p="$env:TEMP\surface7_update.msi"; n 'Downloading Surface Laptop 7 Drivers...'; (New-Object System.Net.WebClient).DownloadFile($u, $p); n 'Installing Surface Laptop 7 Drivers...'; start msiexec -Arg "/i `"$p`" /qn /norestart" -Wait; ri $p -Force; n 'Surface Laptop 7 Drivers Installed Successfully'; sleep 2

################ Chrome ###############

Add-Type -A System.Windows.Forms,System.Drawing; function n($m){$b=New-Object System.Windows.Forms.NotifyIcon;$b.Icon=[System.Drawing.SystemIcons]::Information;$b.Visible=$true;$b.ShowBalloonTip(5000,'Software Install',$m,[System.Windows.Forms.ToolTipIcon]::Info);sleep -m 600;$b.Dispose()}; $u='https://dl.google.com/chrome/install/latest/chrome_installer.exe'; $p="$env:TEMP\chrome_installer.exe"; n 'Downloading Google Chrome...'; (New-Object System.Net.WebClient).DownloadFile($u, $p); n 'Installing Google Chrome...'; start $p -Arg '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART' -Wait; ri $p -Force; n 'Google Chrome Installed Successfully'; sleep 2

```
