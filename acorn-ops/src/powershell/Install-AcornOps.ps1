<#
.SYNOPSIS
    Installs the Acorn Recyclers operations platform.

.DESCRIPTION
    Builds the folder tree, copies the workbooks and templates into it, imports
    the VBA modules into the Console and saves it as a macro-enabled workbook,
    then sets RootPath so the system knows where it lives.

    Importing VBA needs "Trust access to the VBA project object model", which is
    off by default and for good reason. This script turns it on, does the import,
    and turns it straight back off to whatever it was before - including if the
    import fails. Nothing is left more permissive than it was found.

.PARAMETER Root
    Where the platform should live. Use a UNC path for a shared drive
    (\\SERVER\Acorn\AcornOps) rather than a mapped letter, because drive letters
    differ between PCs and the paths are stored in the workbook.

.PARAMETER Source
    The acorn-ops folder from the repository. Defaults to two levels above this
    script.

.PARAMETER SkipTrustedLocation
    Do not register the root as an Excel trusted location. Without it the macros
    are blocked behind the yellow "Enable Content" bar every single time.

.EXAMPLE
    .\Install-AcornOps.ps1 -Root 'D:\AcornOps'

.EXAMPLE
    .\Install-AcornOps.ps1 -Root '\\ACORN-SRV\Ops\AcornOps' -Verbose
#>
[CmdletBinding()]
param(
    [string] $Root = 'C:\AcornOps',
    [string] $Source,
    [switch] $SkipTrustedLocation,
    [switch] $Force
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# --------------------------------------------------------------------------- #
# Layout
# --------------------------------------------------------------------------- #

$Folders = @(
    '00_Admin', '00_Admin\Logs', '00_Admin\Setup',
    '01_Data',
    '02_Inbox', '02_Inbox\Weighbridge', '02_Inbox\Weighbridge\_imported', '02_Inbox\Email',
    '03_Console',
    '04_Documents', '04_Documents\Templates', '04_Documents\Controlled',
    '04_Documents\Controlled\01_Policies', '04_Documents\Controlled\02_Manual',
    '04_Documents\Controlled\03_Procedures', '04_Documents\Controlled\04_WorkInstructions',
    '04_Documents\Controlled\05_Forms', '04_Documents\Controlled\06_ExternalDocuments',
    '04_Documents\Jobs', '04_Documents\Customers', '04_Documents\General',
    '05_Compliance', '05_Compliance\Audits', '05_Compliance\NCR', '05_Compliance\Incidents',
    '05_Compliance\Training', '05_Compliance\Calibration', '05_Compliance\Permits',
    '05_Compliance\LegalRegister', '05_Compliance\ManagementReview',
    '05_Compliance\RiskAssessments', '05_Compliance\SupplierEvaluation',
    '06_Archive', '07_Exports', '99_Scripts'
)

$DataWorkbooks = @('AcornOps_Master.xlsx', 'AcornOps_Operations.xlsx', 'AcornOps_Compliance.xlsx')

# Import order matters only in that modUtil and modConfig are what everything
# else calls; VBA itself does not care, but a failure part-way is easier to read
# when the foundations went in first.
$VbaModules = @(
    'modUtil.bas', 'modConfig.bas', 'modSync.bas', 'modWeighbridge.bas',
    'modOutlook.bas', 'modJobs.bas', 'modDocuments.bas', 'modInvoicing.bas',
    'modAlerts.bas', 'modAdmin.bas', 'modUnattended.bas'
)

$xlOpenXMLWorkbookMacroEnabled = 52

# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

function Write-Step { param([string] $Message) Write-Host "  $Message" -ForegroundColor Cyan }
function Write-Ok   { param([string] $Message) Write-Host "  $Message" -ForegroundColor Green }
function Write-Warn { param([string] $Message) Write-Host "  $Message" -ForegroundColor Yellow }

function Get-ExcelSecurityKey {
    <#  The registry path carries the Office version number, which differs
        between installs (16.0 for 2016 through Microsoft 365). Ask the running
        Excel rather than guessing, and fall back to whatever keys exist.        #>
    param([object] $Excel)

    $version = $null
    if ($Excel) { $version = $Excel.Version }
    if ($version) {
        $key = "HKCU:\Software\Microsoft\Office\$version\Excel\Security"
        if (Test-Path $key) { return $key }
    }
    $candidates = Get-ChildItem 'HKCU:\Software\Microsoft\Office' -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
        Sort-Object { [double] $_.PSChildName } -Descending
    foreach ($c in $candidates) {
        $key = "HKCU:\Software\Microsoft\Office\$($c.PSChildName)\Excel\Security"
        if (Test-Path $key) { return $key }
    }
    return $null
}

function Set-VbomAccess {
    param([string] $SecurityKey, [int] $Value)
    New-ItemProperty -Path $SecurityKey -Name 'AccessVBOM' -Value $Value `
        -PropertyType DWord -Force | Out-Null
}

function Get-VbomAccess {
    param([string] $SecurityKey)
    $prop = Get-ItemProperty -Path $SecurityKey -Name 'AccessVBOM' -ErrorAction SilentlyContinue
    if ($null -eq $prop) { return $null }
    return [int] $prop.AccessVBOM
}

function Add-TrustedLocation {
    param([string] $SecurityKey, [string] $Path)

    $base = Join-Path $SecurityKey 'Trusted Locations'
    if (-not (Test-Path $base)) { New-Item -Path $base -Force | Out-Null }

    # Reuse our own entry if it is already there rather than adding a second one
    # every time the installer is run.
    $target = Join-Path $base 'AcornOps'
    if (-not (Test-Path $target)) { New-Item -Path $target -Force | Out-Null }

    New-ItemProperty -Path $target -Name 'Path' -Value $Path -PropertyType String -Force | Out-Null
    New-ItemProperty -Path $target -Name 'AllowSubfolders' -Value 1 -PropertyType DWord -Force | Out-Null
    New-ItemProperty -Path $target -Name 'Description' `
        -Value 'Acorn Recyclers operations platform' -PropertyType String -Force | Out-Null

    if ($Path -like '\\*') {
        # Network locations are not trusted unless this is also set - and without
        # it a UNC install shows the security bar on every open.
        New-ItemProperty -Path $base -Name 'AllowNetworkLocations' -Value 1 `
            -PropertyType DWord -Force | Out-Null
    }
}

function Set-ConfigValue {
    <#  Writes one setting into the Config table of an open workbook. Finds the
        row by name rather than by position, so inserting a setting later does
        not silently write into the wrong one.                                   #>
    param([object] $Workbook, [string] $Name, [string] $Value)

    $sheet = $Workbook.Worksheets.Item('Config')
    $lastRow = $sheet.Cells($sheet.Rows.Count, 1).End(-4162).Row   # xlUp
    for ($r = 1; $r -le $lastRow; $r++) {
        if ([string] $sheet.Cells($r, 1).Value2 -eq $Name) {
            $sheet.Cells($r, 2).Value2 = $Value
            return $true
        }
    }
    return $false
}

# --------------------------------------------------------------------------- #
# Preflight
# --------------------------------------------------------------------------- #

Write-Host ''
Write-Host 'Acorn Recyclers - operations platform installer' -ForegroundColor White
Write-Host ('=' * 55)
Write-Host ''

if (-not $Source) {
    $Source = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
}
$Source = (Resolve-Path $Source).Path
$dist = Join-Path $Source 'dist'
$vbaDir = Join-Path $Source 'src\vba'
$wordDir = Join-Path $Source 'src\word'

Write-Step "Source:      $Source"
Write-Step "Destination: $Root"
Write-Host ''

foreach ($required in @($dist, $vbaDir)) {
    if (-not (Test-Path $required)) {
        throw "Cannot find '$required'. Point -Source at the acorn-ops folder from the repository."
    }
}

$console = Join-Path $dist 'AcornOps_Console.xlsx'
if (-not (Test-Path $console)) {
    throw "Cannot find $console. Run build/build_all.py first."
}

$targetXlsm = Join-Path $Root '03_Console\AcornOps_Console.xlsm'
if ((Test-Path $targetXlsm) -and -not $Force) {
    throw ("$targetXlsm already exists. Installing again would overwrite it and lose the Config " +
           "settings and counters it holds. Move it aside first, or re-run with -Force if you " +
           "genuinely want it replaced.")
}

# --------------------------------------------------------------------------- #
# 1. Folders
# --------------------------------------------------------------------------- #

Write-Step 'Creating the folder structure...'
$made = 0
foreach ($f in $Folders) {
    $path = Join-Path $Root $f
    if (-not (Test-Path $path)) {
        New-Item -Path $path -ItemType Directory -Force | Out-Null
        $made++
    }
}
Write-Ok "$made folder(s) created, $($Folders.Count - $made) already there."

# --------------------------------------------------------------------------- #
# 2. Files
# --------------------------------------------------------------------------- #

Write-Step 'Copying the data workbooks...'
foreach ($wb in $DataWorkbooks) {
    $src = Join-Path $dist $wb
    $dst = Join-Path $Root "01_Data\$wb"
    if (Test-Path $dst) {
        # Never overwrite live data. This is the one thing in the installer that
        # could destroy a year's records, so it simply refuses.
        Write-Warn "$wb already exists in 01_Data - left alone."
    } else {
        Copy-Item $src $dst
        Write-Ok "$wb copied."
    }
}

Write-Step 'Copying the Word templates...'
if (Test-Path $wordDir) {
    $templates = Get-ChildItem (Join-Path $wordDir '*.docx') -ErrorAction SilentlyContinue
    foreach ($t in $templates) {
        $dst = Join-Path $Root "04_Documents\Templates\$($t.Name)"
        if (Test-Path $dst) {
            Write-Warn "$($t.Name) already there - left alone so your edits survive."
        } else {
            Copy-Item $t.FullName $dst
        }
    }
    Write-Ok "$($templates.Count) template(s) handled."
} else {
    Write-Warn 'No src\word folder - templates skipped.'
}

Write-Step 'Copying the source into 99_Scripts (so the system can be rebuilt)...'
Copy-Item $vbaDir (Join-Path $Root '99_Scripts\vba') -Recurse -Force
Copy-Item (Join-Path $Source 'src\powershell\*') (Join-Path $Root '99_Scripts') -Force
if (Test-Path (Join-Path $Source 'docs')) {
    Copy-Item (Join-Path $Source 'docs') (Join-Path $Root '00_Admin\Setup\docs') -Recurse -Force
}
Copy-Item $console (Join-Path $Root '00_Admin\Setup\AcornOps_Console_original.xlsx') -Force
Write-Ok 'Source and documentation copied.'

# --------------------------------------------------------------------------- #
# 3. Build the macro-enabled Console
# --------------------------------------------------------------------------- #

Write-Step 'Building AcornOps_Console.xlsm...'

$excel = $null
$book = $null
$securityKey = $null
$previousVbom = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $securityKey = Get-ExcelSecurityKey -Excel $excel
    if (-not $securityKey) {
        throw ('Could not find the Excel security registry key. Open Excel once on this PC to ' +
               'create it, then run this again.')
    }

    $previousVbom = Get-VbomAccess -SecurityKey $securityKey
    if ($previousVbom -ne 1) {
        Write-Step 'Temporarily allowing access to the VBA project (restored afterwards)...'
        Set-VbomAccess -SecurityKey $securityKey -Value 1
        # Excel reads this at start-up, so the instance already running has to go.
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
        Start-Sleep -Milliseconds 500
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $false
    }

    $book = $excel.Workbooks.Open($console)

    try {
        $project = $book.VBProject
    } catch {
        throw ('Excel would not let the installer write macros into the workbook. Turn on ' +
               'File > Options > Trust Center > Trust Center Settings > Macro Settings > ' +
               '"Trust access to the VBA project object model", then run this again.')
    }

    $imported = 0
    foreach ($m in $VbaModules) {
        $path = Join-Path $vbaDir $m
        if (-not (Test-Path $path)) { throw "Missing VBA module: $path" }
        $project.VBComponents.Import($path) | Out-Null
        $imported++
    }
    Write-Ok "$imported VBA module(s) imported."

    if (-not (Set-ConfigValue -Workbook $book -Name 'RootPath' -Value $Root)) {
        Write-Warn 'Could not find RootPath on the Config sheet - set it by hand after opening.'
    }

    $book.SaveAs($targetXlsm, $xlOpenXMLWorkbookMacroEnabled)
    Write-Ok "Saved $targetXlsm"

    # Buttons are drawn by the code that was just imported, so this has to happen
    # after the save, not before.
    try {
        $excel.Run('modAdmin.AddButtons') | Out-Null
        $book.Save()
        Write-Ok 'Buttons placed on the Start sheet.'
    } catch {
        Write-Warn ("Could not place the buttons automatically: $($_.Exception.Message)")
        Write-Warn 'Open the Console and run modAdmin.AddButtons from the macro list (Alt+F8).'
    }

    $book.Close($true)
    $book = $null
}
finally {
    # Whatever happened above, the registry goes back how it was found.
    if ($securityKey -and $previousVbom -ne 1) {
        $restore = if ($null -eq $previousVbom) { 0 } else { $previousVbom }
        Set-VbomAccess -SecurityKey $securityKey -Value $restore
        Write-Ok 'VBA project access restored to its previous setting.'
    }
    if ($book) { try { $book.Close($false) } catch { } }
    if ($excel) {
        try { $excel.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# --------------------------------------------------------------------------- #
# 4. Trusted location
# --------------------------------------------------------------------------- #

if (-not $SkipTrustedLocation) {
    Write-Step 'Registering the folder as an Excel trusted location...'
    try {
        Add-TrustedLocation -SecurityKey $securityKey -Path $Root
        Write-Ok 'Done - macros will run without the yellow security bar.'
    } catch {
        Write-Warn ("Could not register the trusted location: $($_.Exception.Message)")
        Write-Warn 'Add it by hand: File > Options > Trust Center > Trusted Locations.'
    }
    Write-Warn 'This applies to THIS Windows user only. Run the installer as each person, or push'
    Write-Warn 'the trusted location by group policy.'
}

# --------------------------------------------------------------------------- #
# Done
# --------------------------------------------------------------------------- #

Write-Host ''
Write-Host 'Installed.' -ForegroundColor Green
Write-Host ''
Write-Host 'Next, in this order:'
Write-Host "  1. Open $targetXlsm"
Write-Host '  2. Config sheet: put your initials in CurrentUser, and fill in the company block'
Write-Host '     (CarrierLicence is legally required on every transfer note).'
Write-Host '  3. Config sheet: set OutlookIntakeFolder, and create that folder in Outlook.'
Write-Host '  4. WB_Map sheet: open a real weighbridge export and copy its column headings in.'
Write-Host '  5. Admin > Clear Example Rows, to remove the worked examples.'
Write-Host '  6. Press Sync.'
Write-Host ''
Write-Host 'The data workbooks in 01_Data are the system of record. Make sure they are inside'
Write-Host 'whatever gets backed up.' -ForegroundColor Yellow
Write-Host ''
