<#
.SYNOPSIS
    The scheduled run. Refreshes the caches, rebuilds the alert list and checks
    document control.

.DESCRIPTION
    Opens the Console with Excel automation, calls modUnattended.RunNightly, and
    closes down cleanly whatever happens. Every message that would normally be a
    dialog goes to the day's log instead, so the task cannot sit waiting for a
    click that nobody is there to give it.

    It deliberately does not import weighbridge files, scrape Outlook or create
    records. Those actions need somebody to see what they did.

.PARAMETER Root
    The AcornOps folder. Defaults to the parent of the folder holding this
    script, which is right when it runs from 99_Scripts.

.PARAMETER TimeoutMinutes
    Give up and close Excel after this long. Stops a stuck run from leaving an
    EXCEL.EXE holding the data workbooks open all night, which would make the
    system read-only for everyone in the morning.

.EXAMPLE
    Register-ScheduledTask -TaskName 'Acorn Ops nightly' -Trigger (
        New-ScheduledTaskTrigger -Daily -At 5am) -Action (
        New-ScheduledTaskAction -Execute 'powershell.exe' `
            -Argument '-NoProfile -ExecutionPolicy Bypass -File C:\AcornOps\99_Scripts\Invoke-AcornNightly.ps1')
#>
[CmdletBinding()]
param(
    [string] $Root,
    [int] $TimeoutMinutes = 20
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if (-not $Root) { $Root = Split-Path -Parent $PSScriptRoot }
$consolePath = Join-Path $Root '03_Console\AcornOps_Console.xlsm'
$logDir = Join-Path $Root '00_Admin\Logs'
$logPath = Join-Path $logDir ("AcornOps-Nightly-{0:yyyy-MM-dd}.log" -f (Get-Date))

if (-not (Test-Path $logDir)) { New-Item -Path $logDir -ItemType Directory -Force | Out-Null }

function Write-Log {
    param([string] $Message)
    $line = "{0:yyyy-MM-dd HH:mm:ss}`t{1}" -f (Get-Date), $Message
    Add-Content -Path $logPath -Value $line
    Write-Verbose $Message
}

Write-Log 'Nightly run starting.'

if (-not (Test-Path $consolePath)) {
    Write-Log "FAILED: cannot find $consolePath"
    exit 1
}

$excel = $null
$book = $null
$exitCode = 0

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    # AutomationSecurity 1 = msoAutomationSecurityLow: run macros without asking.
    # This applies to this automation instance only, not to Excel generally.
    $excel.AutomationSecurity = 1

    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $book = $excel.Workbooks.Open($consolePath, 0, $false)
    Write-Log 'Console opened.'

    $excel.Run('modUnattended.RunNightly') | Out-Null
    Write-Log 'RunNightly completed.'

    if ((Get-Date) -gt $deadline) {
        Write-Log "WARNING: the run took longer than the $TimeoutMinutes minute budget."
    }

    $book.Save()
    $book.Close($false)
    $book = $null
    Write-Log 'Console saved and closed.'
}
catch {
    Write-Log "FAILED: $($_.Exception.Message)"
    $exitCode = 1
}
finally {
    # Excel started by automation does not close itself, and one left running
    # holds the data workbooks open - which locks everybody out in the morning.
    if ($book) { try { $book.Close($false) } catch { } }
    if ($excel) {
        try { $excel.Quit() } catch { }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch { }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Write-Log "Nightly run finished with exit code $exitCode."
}

exit $exitCode
