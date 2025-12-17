param(
    [int]$HoursBack = 24,
    [int]$MaxEventsPerLog = 200
)

$logs = @('Application', 'System')
$start = (Get-Date).AddHours(-$HoursBack)

foreach ($log in $logs) {

    Write-Host ""
    Write-Host ("=" * 80)
    Write-Host ("LOG: {0} (ERRORS since {1})" -f $log, $start) -ForegroundColor Cyan
    Write-Host ("=" * 80)

    $events = Get-WinEvent -FilterHashtable @{
        LogName   = $log
        Level     = 2          # ERROR only
        StartTime = $start
    } -MaxEvents $MaxEventsPerLog 2>$null |
    Sort-Object TimeCreated

    if (-not $events) {
        Write-Host "No ERROR events found." -ForegroundColor Green
        continue
    }

    foreach ($evt in $events) {

        $header = "[{0}] {1:yyyy-MM-dd HH:mm:ss}  ID={2}  Source={3}" -f `
            $log, $evt.TimeCreated, $evt.Id, $evt.ProviderName

        Write-Host $header -ForegroundColor Magenta

        $lines = $evt.Message -split "`r?`n"
        $max = 8
        $count = 0

        foreach ($l in $lines) {
            if ([string]::IsNullOrWhiteSpace($l)) { continue }
            $count++
            if ($count -gt $max) { break }
            Write-Host ("    {0}" -f $l) -ForegroundColor White
        }

        if ($lines.Count -gt $max) {
            Write-Host "    [Message truncated]" -ForegroundColor Yellow
        }

        Write-Host ("-" * 80) -ForegroundColor DarkGray
    }
}





