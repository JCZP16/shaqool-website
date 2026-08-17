# Administration, backup and troubleshooting

## Backups

`01_Data` holds everything. Three files. Everything else - the Console, the caches, the
generated documents - can be rebuilt.

| | |
|---|---|
| **What** | `01_Data\*.xlsx`, plus `04_Documents` and `05_Compliance` for the paperwork |
| **How often** | Daily, and keep at least a month of daily copies |
| **Where** | Somewhere that is not the same disk, and not only a sync folder |
| **Test it** | Restore one, once, and open it. An untested backup is a hope, not a backup |

A word on OneDrive and SharePoint: file sync is **not** a backup. It replicates a
corruption to every copy within seconds. Version history helps and is worth turning on,
but it is not a substitute for a copy somewhere else.

The Console is worth backing up too, not for the data but because it holds your Config
settings and the reference counters. `00_Admin\Setup` keeps a copy of the original.

## The nightly task

`99_Scripts\Invoke-AcornNightly.ps1` opens the Console, refreshes the caches, rebuilds the
alert list and runs the document control check, then closes cleanly. Every message that
would normally be a dialog goes to the log instead, so it cannot sit waiting for a click.

To schedule it:

```powershell
Register-ScheduledTask -TaskName 'Acorn Ops nightly' `
  -Trigger (New-ScheduledTaskTrigger -Daily -At 5am) `
  -Action (New-ScheduledTaskAction -Execute 'powershell.exe' `
     -Argument '-NoProfile -ExecutionPolicy Bypass -File C:\AcornOps\99_Scripts\Invoke-AcornNightly.ps1')
```

Run it on one PC, not several. It has to be a PC that is on at 5am and logged in - Excel
automation will not run in a session that does not exist.

It deliberately does **not** import weighbridge files, scrape Outlook or create records.
Anything that creates a record or moves a file should be done by somebody who can see what
it did, so that when a figure looks wrong on Monday there is a name against the decision
rather than a scheduled task nobody remembers setting up.

To test it before trusting it, run `modUnattended.RunNightlyInteractive` from Alt+F8 -
same work, dialogs left on.

## Archiving

**Archive Year** moves a closed year's job folders into `06_Archive\<year>`.

Two safeguards:

- It works out which year a job belongs to from its **completion date**, not from when
  Windows last touched the folder. A folder opened last week to reprint a note does not
  thereby escape the archive.
- It refuses to move any job whose transfer note is still inside its **retention period** -
  two years for a transfer note, three for a hazardous consignment note. A tidy-up cannot
  destroy a statutory record.

A folder with no matching completed job is skipped and logged rather than moved.

Only documents move. The records stay in the data workbooks, so last year's tonnage is
still on the Dashboard.

## Logs

`00_Admin\Logs`, one tab-separated file per day. Each line: timestamp, Windows username,
the initials from Config, an action code, and the detail.

| Code | |
|---|---|
| `SYNC` | rows cached |
| `WB-STAGE` / `WB-COMMIT` / `WB-RECONCILE` | the weighbridge import, with batch reference |
| `OUTLOOK` | messages scanned, logged, skipped |
| `JOBS` | jobs created from intake |
| `WTN` | transfer notes raised |
| `INVOICE` / `EXPORT` | invoice lines drafted and exported |
| `ALERTS` / `DOCCHECK` | the compliance sweeps |
| `ADMIN` / `ARCHIVE` | folder creation, example clean-up, archiving |
| `WARN` / `ERROR` | anything that did not go to plan |

When someone asks in March how a figure was produced in January, this is where the answer
is. Keep them.

## Troubleshooting

**"Cannot find …\01_Data\AcornOps_Master.xlsx"**
`RootPath` on the Config sheet is wrong, or the share is not mapped. If it is set to a
drive letter, change it to the UNC path.

**"… is open on another PC"**
Excel does not merge two people's edits, so the routine stops rather than silently opening
read-only and throwing your changes away. Find who has it open. Sync itself opens
everything read-only, so it never blocks anyone.

**"Column N of tblJobs is 'X' but the Data_Jobs cache expects 'Y'"**
Somebody has added, removed or renamed a column in a data workbook. Put it back. Every
figure on the Dashboard is a SUMIFS against a fixed column in the cache, so if the sync
carried on regardless, the numbers would quietly start reading the wrong field. That is
why it refuses rather than coping.

If the change was deliberate, it belongs in `build/schema.py` with the Console rebuilt
from it - see [working on the platform](../README.md#working-on-the-platform-itself).

**The buttons have gone**
Alt+F8, run `modAdmin.AddButtons`.

**Macros are disabled / the yellow bar is back**
The trusted location is per Windows user. Re-run the installer as that user, or add the
folder by hand: File > Options > Trust Center > Trusted Locations.

**The recovery rate looks far too low**
Almost always EWC codes with no row on the **WasteStreams** sheet. Those tickets are
counted as `Unclassified`, and unclassified tonnage never counts as diverted. Filter
`Data_WeighTickets` on `RecoveryRoute = Unclassified` to see which codes are missing.

That behaviour is deliberate. The alternative - assuming an unknown code was recycled -
would make the number look better and mean less.

**The mass balance variance is large**
Usually Out tickets that were never recorded, which is a permit problem rather than a
paperwork one. Occasionally it is genuine stock build-up on site. Either way it is worth
knowing which.

**Everything on the Dashboard is zero**
Press Sync. Check `I4` does not say `never`.

## When to outgrow this

This is Excel on a shared drive. It suits a yard with a handful of people; it is not a
multi-user database and should not be asked to be one. The honest signals that you have
outgrown it:

- **More than three or four people needing to write at the same time.** One writer at a
  time is the hard constraint here, and no amount of tuning changes it.
- **The Operations workbook past roughly 50,000 weighbridge tickets** - three to four
  years at a fair volume. It will still work, but opening and syncing get slow enough to
  be annoying.
- **You want it on phones in the cabs.** Drivers confirming a movement from a handset is
  the single biggest operational win available, and it needs a real database and an app.
- **A customer or a certification body asks for a full audit trail of edits**, not just of
  automated actions. This logs what the system did; it does not log every keystroke
  someone typed into a cell.

None of that is wasted work when it comes. The data model in
[03-data-model.md](03-data-model.md) is a normalised relational design that came out of
thinking about the business, not about Excel - thirty tables with proper keys. It ports to
SQL Server, Postgres or a low-code platform more or less directly, and by then you will
have a year or two of real data proving which fields you actually use and which ones
nobody ever filled in. That knowledge is worth more than the spreadsheets.
