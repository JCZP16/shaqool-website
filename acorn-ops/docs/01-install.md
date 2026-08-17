# Installing

## What you need

- **Windows, with Excel for Microsoft 365 or Excel 2016 or later, desktop version.**
  Excel for Mac and Excel on the web cannot drive Outlook or Word, and the whole point of
  this is that it drives Outlook and Word.
- **Outlook**, desktop, signed into the mailbox the enquiries arrive in.
- **Word**, desktop, for generating transfer notes and confirmations.
- A folder everyone in the office can reach - a mapped server share, a synced SharePoint
  or OneDrive folder, or a local folder if only one person is using it.

Use the **UNC path** (`\\ACORN-SRV\Ops\AcornOps`), not a mapped drive letter. The paths
get stored in the workbook, and `S:\` on your PC may well be `T:\` on someone else's.

## The quick way

From a PowerShell window on a PC that has Excel installed:

```powershell
cd acorn-ops\src\powershell
.\Install-AcornOps.ps1 -Root '\\ACORN-SRV\Ops\AcornOps'
```

If PowerShell refuses to run it:

```powershell
powershell -ExecutionPolicy Bypass -File .\Install-AcornOps.ps1 -Root 'C:\AcornOps'
```

## What the installer does, and what it touches

1. Creates the folder tree under `-Root`.
2. Copies the three data workbooks into `01_Data`. **If one is already there it is left
   alone** - the installer will not overwrite live data, ever.
3. Copies the Word templates into `04_Documents\Templates`, again leaving any that already
   exist so your edits survive a reinstall.
4. Copies the VBA source, the scripts and the documentation into `99_Scripts` and
   `00_Admin\Setup`, so the system can be rebuilt from what is on the share.
5. Imports the eleven VBA modules into the Console and saves it as
   `03_Console\AcornOps_Console.xlsm`.
6. Writes `RootPath` into the Config sheet.
7. Puts the buttons on the Start sheet.
8. Registers the root folder as an Excel **trusted location**.

Two of those touch the registry, so they are worth being clear about:

**Trust access to the VBA project object model.** Importing macros into a workbook needs
this, and it is off by default for a good reason - it is what lets one macro write
another. The installer turns it on, does the import, and puts it back exactly how it found
it, including if the import fails partway. Nothing is left more permissive than it was.

**Trusted location.** Without this, the Console shows the yellow "Enable Content" bar
every time anyone opens it, and sooner or later someone stops enabling it and quietly
starts using a Console with no macros. Pass `-SkipTrustedLocation` if your IT policy says
otherwise and you would rather add it centrally.

Both settings are **per Windows user**. Run the installer as each person who will use the
system, or push the trusted location by group policy. Running it again on a second PC is
safe: the data workbooks are already there and get left alone.

If `03_Console\AcornOps_Console.xlsm` already exists the installer stops rather than
replacing it, because that file holds your Config settings and the reference counters.
`-Force` overrides that, but move the old one aside first.

## Doing it by hand

If the script will not run - locked-down PowerShell, no permission to touch the registry -
this is all it was doing:

1. Create the folder tree. Or: open `AcornOps_Console.xlsx`, set `RootPath` on the Config
   sheet, and use **Admin > Create Folder Structure** once the macros are in.
2. Copy the three workbooks from `dist/` into `01_Data`, the eight `.docx` from `src/word/`
   into `04_Documents\Templates`.
3. In Excel: **File > Options > Trust Center > Trust Center Settings > Macro Settings**,
   tick *Trust access to the VBA project object model*.
4. Open `dist\AcornOps_Console.xlsx`, press **Alt+F11**, then **File > Import File** for
   each `.bas` in `src/vba/`. All eleven.
5. **File > Save As**, change the type to **Excel Macro-Enabled Workbook (\*.xlsm)**, save
   it as `03_Console\AcornOps_Console.xlsm`.
6. Press **Alt+F8**, run `modAdmin.AddButtons`.
7. Untick *Trust access to the VBA project object model* again.
8. **Trust Center > Trusted Locations > Add new location**, point it at the root, tick
   *Subfolders of this location are also trusted*.

## First run - in this order

Open `03_Console\AcornOps_Console.xlsm`. Work down the **Config** sheet:

| Setting | What to put |
|---|---|
| `RootPath` | Already filled in by the installer. Check it is the UNC path, not a drive letter. |
| `CurrentUser` | Your initials. Stamped on everything you create. **Set this on each PC.** |
| `CompanyName` … `CompanyEmail` | Printed on every document that goes to a customer. |
| `CarrierLicence` | Your waste carrier registration. **Legally required on every transfer note** - a note without it is not a valid note. |
| `EnvironmentalPermit` | The transfer station's permit number. |
| `CompanySICCode` | Defaults to 38320 (recovery of sorted materials). Change if yours differs. |
| `VATRate` | 0.20 unless something has changed. |
| `OutlookIntakeFolder` | Defaults to `Inbox\Acorn Intake`. **Create that folder in Outlook first.** |
| `LongHireDays` | How long a container can sit before it wants chasing. |
| `RecoveryTargetPct` | Your diversion target - drives the red/amber/green on the Dashboard. |

Then:

1. **WB_Map sheet** - open a real export from the skip software and copy its column
   headings into the SourceHeader column. See
   [04-weighbridge-mapping.md](04-weighbridge-mapping.md). Get this right before the first
   import and everything downstream follows.
2. **Admin > Clear Example Rows** - removes the amber worked-example row from every table.
   It only removes a row that is still exactly the example that shipped, so it cannot
   touch anything you have typed.
3. Fill in the master data: customers, sites, containers, vehicles, drivers, EWC codes,
   outlets, rate card. This is the tedious afternoon, and it is the one that decides
   whether the rest of it works.
4. Press **Sync**.
5. Press **Alerts**. It should be empty, or should show exactly what you already know is
   due.

## The EWC codes are the bit that matters

Of all the master data, `tblWasteStreams` in the Master workbook is the one to take
seriously. Every weighbridge ticket carries an EWC code, and the **RecoveryRoute** against
that code is what turns tonnage into a recycling rate.

A code with no row on that sheet is counted as **Unclassified**, and unclassified tonnage
never counts as diverted. That is deliberate - it makes the gap visible instead of
flattering the number - but it does mean the recovery rate reads low until the sheet is
complete.

## Backing it up

`01_Data` holds everything. Three files. If they are not in whatever gets backed up, none
of the rest of this matters. See
[07-admin-and-backup.md](07-admin-and-backup.md#backups).
