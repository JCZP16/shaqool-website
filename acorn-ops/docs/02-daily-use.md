# Using it day to day

Everyone opens one file: `03_Console\AcornOps_Console.xlsm`. The buttons are on the Start
sheet.

Behind the buttons, the data lives in three workbooks in `01_Data`. You can open those
directly to type into them - that is how you add a customer or close an NCR - but the
Console is where the work gets driven from.

---

## The morning

**1. Sync.** Pulls the three data workbooks into the Console. Two or three seconds.
Everything else you look at is only as current as the last sync, so do this when you sit
down and again after anyone else has been typing.

**2. Today.** The job board: everything booked, allocated, on site or awaiting collection,
plus anything finished today. Sorted by scheduled date. Late jobs are red, today's are
green.

The board is a view, not a form. To change something, change it in the Operations
workbook - the Console will pick it up on the next Sync.

**3. Scrape Outlook.** Reads the intake folder, classifies each message, saves the
attachments and logs it. Safe to run as often as you like: a message that has already been
logged is recognised and skipped, even if it has been moved since.

**4. Intake.** Work down the list. Each row shows what the message was about, which
customer it looks like, the postcode, the skip size and the date asked for.

- Bookings: type **Book** in the Outcome column, then press **Convert to Jobs**.
- Anything else: type whatever you did (`quoted`, `passed to accounts`, `junk`). Anything
  other than *Book* just marks it dealt with and drops it off the list next time.

**Convert to Jobs** raises a job at status Booked for every row marked Book. If the
customer could not be identified it skips the row and says so, rather than guessing - a
job on the wrong account takes a lot longer to unpick than one that needed thirty seconds
of typing. Where the site cannot be matched it creates one from the postcode and flags it
in the job notes for somebody to complete.

---

## Through the day

**Raise Transfer Note.** Click the job on the Today board, press the button, confirm the
job reference. It raises the record, works out the tonnage from the matched weighbridge
tickets, picks a **hazardous consignment note** automatically if the EWC code is flagged
hazardous, produces the Word and PDF, and files them in the job's folder.

**Generate Document.** For everything else - confirmations, quotes, NCR reports, toolbox
talks, inductions. Click the row on the **DocGen** sheet for the template you want, press
the button, give it the reference it asks for.

To change what a document looks like, edit the `.docx` in `04_Documents\Templates` in
Word. Keep the `{{Tokens}}` and you can move them anywhere, restyle them, or drop your
logo in the header. The full token list is on the **TokenRef** sheet.

---

## When the weighbridge exports arrive

**1.** Export from the skip software into `02_Inbox\Weighbridge`.

**2. Import Weighbridge.** Reads every CSV in that folder through the column map on
**WB_Map**, and puts the result on the **WB_Staging** sheet. Nothing has been written to
the data yet.

Rows that do not stand up are marked Rejected in red with the reason: no ticket number, an
unreadable date, a direction that is neither In nor Out, an EWC code that is not six
digits, gross less than tare. Fix the map, run it again - it is entirely safe to repeat.

**3. Write Staged Tickets.** Writes the rows marked OK into the Operations workbook. A
ticket number that already exists is **updated**, not duplicated, so re-importing an
overlapping export corrects rather than doubles. The source files are moved to
`_imported` with the batch stamp on them - they are the evidence behind every tonne you
invoice, so they are kept.

**4. Reconcile Tickets.** Matches tickets that arrived without a job reference to jobs, on
vehicle registration and date. It only accepts an unambiguous single match. If the same
lorry did three jobs that day it leaves the ticket alone rather than guessing, and you set
the JobID by hand.

> Net weight is always recalculated from the gross and tare actually recorded, never taken
> from the export. If the bridge and this system ever disagree, the arithmetic on the
> recorded weights is the one you can defend.

---

## Weekly

**Alerts.** Everything overdue or falling due in the next thirty days, from all three
workbooks in one list: vehicle tax, MOT, insurance, LOLER, tacho calibration, driver
licences, CPC, medicals, weighbridge calibration, container inspections, training
expiries, document reviews, legal register reviews, outlet permits, supplier licences and
insurance, NCR actions, planned audits, objectives and risk reviews.

Plus three things that are not dates but want chasing: containers out longer than
`LongHireDays`, weighbridge tickets with no job against them, and incidents still open
after a fortnight.

Red is overdue, amber is due soon, blue is a chase. An empty sheet means you are straight.

**Build Invoice Lines.** Turns completed jobs into charge lines - the job itself, excess
hire days, and excess tonnage worked out from the weighbridge. A job that already has any
line against it is skipped, so running it twice cannot bill the same work twice.

Check the lines on the InvoiceLines sheet, set the ones you are happy with to **Approved**,
then **Export Invoice Lines** writes them to `07_Exports` as CSV and marks them Exported.
Draft lines are deliberately left behind.

---

## Monthly

**Dashboard.** Live position on the left, the reporting month in the middle, compliance on
the right, and a rolling twelve months underneath.

`C4` sets the month. It defaults to the current one; type any month start over it to look
back, or put `=EOMONTH(TODAY(),-1)+1` back to return to automatic.

Two figures worth understanding:

**Recovery rate** is tonnes despatched to a non-landfill outlet, divided by total tonnes
despatched. It is measured on **outputs** - where the material actually ended up - because
that is what the regulator and your customers will ask for. An inputs-based figure flatters
you by ignoring residue. The number is only ever as honest as the RecoveryRoute column on
the WasteStreams sheet, so review that whenever you change outlet.

**Mass balance variance** is tonnes in minus tonnes out, as a share of tonnes in. Some
variance is normal - stock on site, moisture loss. A big one usually means Out tickets are
missing, and missing Out tickets are a permit problem, not a paperwork one.

**Document Control Check.** The check an auditor makes: is there anything in
`04_Documents\Controlled` that is not on the register, is there anything on the register
whose file has gone, and is anything issued past its review date. Findings land on the
DocCheck sheet.

---

## Yearly

**Archive Year.** Moves a closed year's job folders into `06_Archive`. It will not touch a
job whose transfer note is still inside its retention period - two years for a transfer
note, three for a hazardous consignment note - so a tidy-up cannot destroy a statutory
record. It works out which year a job belongs to from its completion date, not from when
Windows last touched the folder.

The records themselves stay in the data workbooks. Only documents move.

---

## Things worth knowing

**Two people, one workbook.** Excel does not merge edits. If someone else has a data
workbook open, the Console will tell you rather than opening it read-only and quietly
losing your changes. Sync reads all three read-only, so it never blocks anyone.

**Everything is logged.** `00_Admin\Logs`, one file a day. Every import, every document,
every error, with the Windows username and the initials from Config. When someone asks in
March how a figure was produced in January, that is where the answer is.

**Nothing is half-written.** Every routine that writes either completes or leaves the data
workbooks untouched. If something fails you get a message saying so, and the detail goes
to the log.

**References are never reused.** A failed run may leave a gap in the numbering. A gap is
fine. A duplicate is not.
