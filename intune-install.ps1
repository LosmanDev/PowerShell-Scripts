$ErrorActionPreference = "SilentlyContinue"

$IMEPath = "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs"
$PMPCPath = "C:\ProgramData\PatchMyPCIntuneLogs"

function Convert-ToDateTime {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    try { return [datetime]::Parse($Value) } catch { return $null }
}

function Get-IMEWin32Inventory {
    param([string]$LogFolder)

    $file = Join-Path $LogFolder "Win32AppInventory.log"
    if (-not (Test-Path $file)) { return @() }

    $pattern = "Id:\s*([0-9a-f\-]+)\s*Name:\s*(.*?)\s*Version:\s*([\w\.\-]+)"
    $results = @()

    Select-String -Path $file -Pattern "Id:" | ForEach-Object {
        if ($_.Line -match $pattern) {
            $appId = $matches[1]
            $appName = $matches[2]
            $version = $matches[3]

            $date = $null
            $time = $null

            if ($_.Line -match 'date="([^"]+)"') { $date = $matches[1] }
            if ($_.Line -match 'time="([^"]+)"') { $time = $matches[1] }

            if ($date -and $time) {
                $tsString = "$date $time"
            }
            elseif ($time) {
                $tsString = $time
            }
            else {
                $idx = $_.Line.IndexOf("<time=")
                if ($idx -ge 0) { $tsString = $_.Line.Substring($idx + 6, 12) }
            }

            $timestamp = Convert-ToDateTime $tsString

            $results += [PSCustomObject]@{
                Timestamp = $timestamp
                AppName   = $appName
                Version   = $version
                Status    = "Inventory"
                Detail    = "Win32 inventory: Installed (AppId=$appId)"
                ExitCode  = $null
                Source    = "Intune-Inventory"
                LogFile   = $file
            }
        }
    }

    return $results
}

function Get-PMPCDetectionEvents {
    param([string]$LogFolder)

    if (-not (Test-Path $LogFolder)) { return @() }

    $files = Get-ChildItem $LogFolder -Filter "PatchMyPC-Software*DetectionScript*.log" -File
    if (-not $files) { return @() }

    $events = @()

    foreach ($file in $files) {
        Get-Content $file | ForEach-Object {
            $line = $_
            if ([string]::IsNullOrWhiteSpace($line)) { return }

            $parts = $line -split "~"
            if ($parts.Count -lt 2) { return }

            $rawTs = $parts[0].Trim()
            $rawName = $parts[1].Trim("[] ")
            $found = ($parts | Where-Object { $_ -like "*Found:*" } | Select-Object -First 1)
            $purpose = ($parts | Where-Object { $_ -like "*Purpose:*" } | Select-Object -First 1)

            $timestamp = Convert-ToDateTime $rawTs

            $foundValue = $null
            if ($found -match "Found:(True|False)") { $foundValue = $matches[1] }

            $purposeValue = $null
            if ($purpose -match "Purpose:([^]]+)") { $purposeValue = $matches[1] }

            $appName = $rawName
            $version = $null

            $tokens = $rawName -split "\s+"
            if ($tokens.Count -gt 1 -and $tokens[-1] -match "^[0-9A-Za-z\.\-]+$") {
                $version = $tokens[-1]
                $appName = ($tokens[0..($tokens.Count - 2)] -join " ")
            }

            $status =
            if ($purposeValue -eq "Detection" -and $foundValue -eq "True") { "Detected" }
            elseif ($purposeValue -eq "Detection" -and $foundValue -eq "False") { "NotDetected" }
            elseif ($purposeValue -eq "Requirement" -and $foundValue -eq "False") { "RequirementFailed" }
            elseif ($purposeValue -eq "Requirement" -and $foundValue -eq "True") { "RequirementMet" }
            else { "Unknown" }

            $detail = "Patch My PC ${purposeValue}: Found=$foundValue"

            $events += [PSCustomObject]@{
                Timestamp = $timestamp
                AppName   = $appName
                Version   = $version
                Status    = $status
                Detail    = $detail
                ExitCode  = $null
                Source    = "PMPC-Detection"
                LogFile   = $file.FullName
            }
        }
    }

    return $events
}

function Get-IMEAppActionEvents {
    param([string]$LogFolder)

    $files = Get-ChildItem $LogFolder -Filter "AppActionProcessor*.log" -File
    if (-not $files) { return @() }

    $events = @()

    foreach ($file in $files) {
        Get-Content $file | ForEach-Object {
            $line = $_
            if (-not $line) { return }

            if ($line -match "app\s+'(?<AppName>[^']+)'\s*\(Id[:\s]+(?<AppId>[0-9a-f\-]+)\).*assignment type[:\s]+(?<AssignType>\w+)") {

                $timestamp = $null
                if ($line -match "^(?<ts>\d{2}[-/]\d{2}[-/]\d{4}\s+\d{2}:\d{2}:\d{2})") {
                    $timestamp = Convert-ToDateTime $matches.ts
                }

                $events += [PSCustomObject]@{
                    Timestamp = $timestamp
                    AppName   = $matches.AppName
                    Version   = $null
                    Status    = "Evaluating-$($matches.AssignType)"
                    Detail    = "AppActionProcessor evaluating assignment"
                    ExitCode  = $null
                    Source    = "IME-AppActionProcessor"
                    LogFile   = $file.FullName
                }
            }

            elseif ($line -match "Detection for app\s+'(?<AppName>[^']+)'\s*\(Id[:\s]+(?<AppId>[0-9a-f\-]+)\).*result[:\s]+(?<Result>\w+)") {

                $timestamp = $null
                if ($line -match "^(?<ts>\d{2}[-/]\d{2}[-/]\d{4}\s+\d{2}:\d{2}:\d{2})") {
                    $timestamp = Convert-ToDateTime $matches.ts
                }

                $result = $matches.Result
                $status =
                if ($result -eq "Installed") { "Detection-Installed" }
                elseif ($result -eq "NotInstalled") { "Detection-NotInstalled" }
                else { "Detection-$result" }

                $events += [PSCustomObject]@{
                    Timestamp = $timestamp
                    AppName   = $matches.AppName
                    Version   = $null
                    Status    = $status
                    Detail    = "AppActionProcessor detection result = $result"
                    ExitCode  = $null
                    Source    = "IME-AppActionProcessor"
                    LogFile   = $file.FullName
                }
            }
        }
    }

    return $events
}

function Get-IMEWorkloadEvents {
    param([string]$LogFolder)

    $files = Get-ChildItem $LogFolder -Filter "AppWorkload*.log" -File
    if (-not $files) { return @() }

    $events = @()

    foreach ($file in $files) {
        Get-Content $file | ForEach-Object {
            $line = $_
            if (-not $line) { return }

            if ($line -match "Queu\w+\s+app\s+'(?<AppName>[^']+)'\s*\(Id[:\s]+(?<AppId>[0-9a-f\-]+)\).*for\s+(?<Action>\w+)") {

                $timestamp = $null
                if ($line -match "^(?<ts>\d{2}[-/]\d{2}[-/]\d{4}\s+\d{2}:\d{2}:\d{2})") {
                    $timestamp = Convert-ToDateTime $matches.ts
                }

                $events += [PSCustomObject]@{
                    Timestamp = $timestamp
                    AppName   = $matches.AppName
                    Version   = $null
                    Status    = "Queued-$($matches.Action)"
                    Detail    = "AppWorkload queued action"
                    ExitCode  = $null
                    Source    = "IME-AppWorkload"
                    LogFile   = $file.FullName
                }
            }

            elseif ($line -match "execution state\s*=\s*(?<State>\w+).+app\s+'(?<AppName>[^']+)'") {

                $timestamp = $null
                if ($line -match "^(?<ts>\d{2}[-/]\d{2}[-/]\d{4}\s+\d{2}:\d{2}:\d{2})") {
                    $timestamp = Convert-ToDateTime $matches.ts
                }

                $events += [PSCustomObject]@{
                    Timestamp = $timestamp
                    AppName   = $matches.AppName
                    Version   = $null
                    Status    = "Execution-$($matches.State)"
                    Detail    = "AppWorkload execution state"
                    ExitCode  = $null
                    Source    = "IME-AppWorkload"
                    LogFile   = $file.FullName
                }
            }
        }
    }

    return $events
}

function Get-IMEAgentExecutorEvents {
    param([string]$LogFolder)

    $files = Get-ChildItem $LogFolder -Filter "AgentExecutor*.log" -File
    if (-not $files) { return @() }

    $events = @()

    foreach ($file in $files) {
        Get-Content $file | ForEach-Object {
            $line = $_
            if (-not $line) { return }

            if ($line -match "Starting .*Win32 app\s+'(?<AppName>[^']+)'\s*\(Id[:\s]+(?<AppId>[0-9a-f\-]+)\)") {

                $timestamp = $null
                if ($line -match "^(?<ts>\d{2}[-/]\d{2}[-/]\d{4}\s+\d{2}:\d{2}:\d{2})") {
                    $timestamp = Convert-ToDateTime $matches.ts
                }

                $events += [PSCustomObject]@{
                    Timestamp = $timestamp
                    AppName   = $matches.AppName
                    Version   = $null
                    Status    = "Install-Starting"
                    Detail    = "AgentExecutor starting install"
                    ExitCode  = $null
                    Source    = "IME-AgentExecutor"
                    LogFile   = $file.FullName
                }
            }

            elseif ($line -match "Execution .*for app\s+'(?<AppName>[^']+)'\s*\(Id[:\s]+(?<AppId>[0-9a-f\-]+)\).*(exit code|ExitCode)\s*(?<Code>-?\d+)") {

                $code = [int]$matches.Code

                $timestamp = $null
                if ($line -match "^(?<ts>\d{2}[-/]\d{2}[-/]\d{4}\s+\d{2}:\d{2}:\d{2})") {
                    $timestamp = Convert-ToDateTime $matches.ts
                }

                $status = if ($code -eq 0) { "Install-Success" } else { "Install-Failed" }

                $events += [PSCustomObject]@{
                    Timestamp = $timestamp
                    AppName   = $matches.AppName
                    Version   = $null
                    Status    = $status
                    Detail    = "AgentExecutor completed with code $code"
                    ExitCode  = $code
                    Source    = "IME-AgentExecutor"
                    LogFile   = $file.FullName
                }
            }
        }
    }

    return $events
}

function Get-IMEHighLevelEvents {
    param([string]$LogFolder)

    $files = Get-ChildItem $LogFolder -Filter "IntuneManagementExtension*.log" -File
    if (-not $files) { return @() }

    $events = @()

    foreach ($file in $files) {
        Get-Content $file | ForEach-Object {
            $line = $_
            if (-not $line) { return }

            if ($line -match "Policies synchronized|Policy sync completed|Syncing policies") {

                $timestamp = $null
                if ($line -match "^(?<ts>\d{2}[-/]\d{2}[-/]\d{4}\s+\d{2}:\d{2}:\d{2})") {
                    $timestamp = Convert-ToDateTime $matches.ts
                }

                $events += [PSCustomObject]@{
                    Timestamp = $timestamp
                    AppName   = "[IME]"
                    Version   = $null
                    Status    = "PolicySync"
                    Detail    = $line
                    ExitCode  = $null
                    Source    = "IME-Core"
                    LogFile   = $file.FullName
                }
            }

            elseif ($line -match "Failed .*policy") {

                $timestamp = $null
                if ($line -match "^(?<ts>\d{2}[-/]\d{2}[-/]\d{4}\s+\d{2}:\d{2}:\d{2})") {
                    $timestamp = Convert-ToDateTime $matches.ts
                }

                $events += [PSCustomObject]@{
                    Timestamp = $timestamp
                    AppName   = "[IME]"
                    Version   = $null
                    Status    = "PolicyError"
                    Detail    = $line
                    ExitCode  = $null
                    Source    = "IME-Core"
                    LogFile   = $file.FullName
                }
            }
        }
    }

    return $events
}

Write-Host "Collecting data..." -ForegroundColor Cyan

$Dashboard = @()
$Dashboard += Get-IMEWin32Inventory -LogFolder $IMEPath
$Dashboard += Get-PMPCDetectionEvents -LogFolder $PMPCPath
$Dashboard += Get-IMEAppActionEvents -LogFolder $IMEPath
$Dashboard += Get-IMEWorkloadEvents -LogFolder $IMEPath
$Dashboard += Get-IMEAgentExecutorEvents -LogFolder $IMEPath
$Dashboard += Get-IMEHighLevelEvents -LogFolder $IMEPath

$Dashboard = $Dashboard | Where-Object { $_ } | Sort-Object Timestamp -Descending

$Dashboard | Out-GridView -Title "Intune + Patch My PC Unified App Insight"
