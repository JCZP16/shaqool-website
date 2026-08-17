# Acorn Recyclers - operations platform

An Excel-based operations and compliance system for a skip hire and waste recycling
business, joined up with Outlook and Word, laid out as an ISO-style folder and register
structure, and automated as far as it is sensible to automate.

It is built to sit **alongside** the weighbridge software, not to replace it. The
weighbridge keeps doing what it is good at - weighing lorries and printing tickets. This
reads its export and does everything the skip software does not: hire tracking, duty of
care, recovery rates, compliance dates, document control and the paperwork.

---

## What it actually does

| | |
|---|---|
| **Job book** | Every container movement from enquiry to invoice, with the skip, driver, lorry and site against it. Containers on hire, and which ones have been out too long. |
| **Weighbridge import** | Reads the skip software's CSV export through a column map you control on a spreadsheet. Recalculates net weight rather than trusting the file, rejects rows that do not stand up, and matches tickets to jobs. |
| **Outlook intake** | Reads a mail folder, classifies each message against rules you can edit, works out the customer, postcode, skip size and wanted-for date, saves the attachments into a folder named after the email, and logs it. A weighbridge CSV that arrives by email lands in the import folder on its own. |
| **Word documents** | Waste transfer notes, hazardous consignment notes, job confirmations, quotes, NCR reports, toolbox talks and inductions - filled from the data and filed against the job as .docx and .pdf. |
| **Recovery rate** | Tonnes in, tonnes out, tonnes diverted, and the mass balance between them. Measured on outputs, which is the figure the regulator and your customers will ask for. |
| **ISO registers** | Document control, NCR and corrective action, audits, training and the training matrix, calibration, permits, incidents, legal register, objectives, risk register, management review and supplier evaluation. |
| **One alert list** | Every renewal, licence, calibration, training expiry and overdue action across all three workbooks in a single sheet, red for overdue and amber for due soon. If it is empty, you are straight. |
| **Invoice lines** | Turns completed jobs and weighed tonnage into charge lines - including the excess tonnage and excess hire days that get missed - and exports them as CSV for the accounts package. |

## What it deliberately does not do

- **It is not the accounting system.** It prices work and hands charge lines to Sage,
  Xero or whatever you use. VAT, credit control and the statutory accounting records stay
  where they are.
- **It does not replace the weighbridge software.** It has no connection to the bridge
  itself, and it should not have one.
- **It does not answer emails or take bookings on its own.** It reads, files and
  suggests. Creating a job is a decision a person makes, on a sheet, with the evidence in
  front of them. There is a switch to make it more automatic; leave it off until the
  classification has been right for a few weeks.
- **It is not a database server.** Excel on a shared drive means one person writing at a
  time. That is fine for a yard; it would not be fine for fifty concurrent users. See
  [When to outgrow this](docs/07-admin-and-backup.md#when-to-outgrow-this).

---

## The shape of it

```
AcornOps/
├── 00_Admin/          logs, installer, a copy of everything needed to rebuild
├── 01_Data/           the three data workbooks - THIS is the system of record
├── 02_Inbox/          weighbridge exports and Outlook attachments, waiting to be processed
├── 03_Console/        AcornOps_Console.xlsm - the only file anyone opens
├── 04_Documents/      Word templates, controlled documents, and one folder per job
├── 05_Compliance/     certificates, audit reports, NCR evidence, permits
├── 06_Archive/        closed years
├── 07_Exports/        CSVs for the accounts package
└── 99_Scripts/        installer, VBA source, template source
```

Four workbooks, in two layers:

**The data layer** - three workbooks in `01_Data`, thirty tables between them. This is
the system of record and the thing that gets backed up.

- `AcornOps_Master.xlsx` - customers, sites, containers, vehicles, drivers, staff, EWC
  codes, onward outlets, rate card, suppliers
- `AcornOps_Operations.xlsx` - jobs, movements, weighbridge tickets, transfer notes,
  invoice lines, the email log, generated documents
- `AcornOps_Compliance.xlsx` - the ISO 9001 / 14001 / 45001 registers

**The front end** - `AcornOps_Console.xlsm` in `03_Console`. It holds no data of its own.
It reads the three data workbooks into hidden cache sheets and reports over them, so it
can be closed, copied or rebuilt from scratch without putting a record at risk.

That split is the single most important design decision here. It means a corrupted or
out-of-date Console is never a data loss - you re-sync it and carry on.

---

## Getting it running

Full detail in **[docs/01-install.md](docs/01-install.md)**. The short version, on a
Windows PC with Excel and Outlook:

```powershell
cd acorn-ops\src\powershell
.\Install-AcornOps.ps1 -Root '\\ACORN-SRV\Ops\AcornOps'
```

Then open the Console and, in this order: set `CurrentUser` and the company block on the
**Config** sheet, point **WB_Map** at your real weighbridge export headings, run
**Admin > Clear Example Rows**, and press **Sync**.

You need Excel for Microsoft 365 or Excel 2016+ on Windows desktop. Excel for Mac and
Excel on the web cannot do the Outlook and Word automation this depends on.

---

## Documentation

| | |
|---|---|
| [01-install.md](docs/01-install.md) | Installing it, what the installer touches, and doing it by hand if the script will not run |
| [02-daily-use.md](docs/02-daily-use.md) | The daily, weekly and monthly rhythm - what to press and when |
| [03-data-model.md](docs/03-data-model.md) | Every table and field. *Generated from the schema* |
| [04-weighbridge-mapping.md](docs/04-weighbridge-mapping.md) | Pointing the importer at your export, whatever shape it is |
| [05-outlook-intake.md](docs/05-outlook-intake.md) | The mail rules, what gets extracted, and how to tune it |
| [06-iso-mapping.md](docs/06-iso-mapping.md) | Which register answers which clause. *Generated* |
| [07-admin-and-backup.md](docs/07-admin-and-backup.md) | Backups, the nightly task, archiving, troubleshooting, and when to outgrow this |

---

## Working on the platform itself

Everything is generated from one schema file. Do not hand-edit the workbooks in `dist/` -
change `build/schema.py` and rebuild.

```bash
pip install openpyxl python-docx formulas
cd acorn-ops/build
python3 build_all.py
```

`build_all.py` regenerates the four workbooks, the eight Word templates and the reference
documentation, then runs two sets of checks and fails loudly if anything disagrees:

- **`lint_vba.py`** - block balance, line continuations, cross-module calls, button
  wiring, duplicate procedure names, and `ReDim Preserve` on the wrong dimension. There is
  no VBA compiler here, so this stands in for one.
- **`verify.py`** - every formula parsed, its functions checked against what Excel can
  evaluate without an `_xlfn.` prefix, and the whole workbook computed with the `formulas`
  library to catch anything that would land as an error value. It also cross-checks the
  hand-written table list in `modSync` and the example-row IDs in `modAdmin` against the
  schema, and checks that every `{{Token}}` in a Word template is one the code can supply.

Those cross-checks exist because the VBA carries a few hand-written copies of things the
schema also defines. Drift there is silent - a table simply stops syncing - so it is
checked on every build rather than left to be discovered six months later.

```
acorn-ops/
├── build/          schema.py and the generators - the source of truth
├── src/vba/        11 modules, imported into the Console by the installer
├── src/powershell/ installer and the nightly scheduled task
├── src/word/       generated Word templates
├── dist/           generated workbooks
├── sample-data/    a realistic weighbridge export to test the import against
└── docs/
```
