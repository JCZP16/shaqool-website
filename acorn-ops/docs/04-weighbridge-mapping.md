# Pointing the importer at your weighbridge export

The importer knows nothing about your skip software's file format. Every column heading,
every coded value and every unit conversion is declared on the Console's **WB_Map** sheet.

That is the whole design. When the software is updated and the headings move, or you
change supplier entirely, the fix is a spreadsheet edit - not a change to any code.

## Setting it up

1. Export a real file from the skip software into `02_Inbox\Weighbridge`.
2. Open it and look at the heading row.
3. On **WB_Map**, put each heading in the **SourceHeader** column against the field it
   feeds. Leave SourceHeader blank for anything the export does not have.
4. Press **Import Weighbridge** and read the Message column on WB_Staging.

Nothing is written to the data until you press **Write Staged Tickets**, so this loop is
completely safe to repeat until it looks right.

## The fields

| TargetField | Needed? | Notes |
|---|---|---|
| `TicketNo` | **yes** | Must be unique. Re-importing the same number updates the row instead of duplicating it, which is what makes an overlapping export harmless. |
| `TicketDateTime` | **yes** | If the export splits date and time into two columns, put the date heading here and the time heading in SourceHeader2. |
| `Direction` | **yes** | Must end up as exactly `In` or `Out`. Use a `Map:` transform for coded values. |
| `EWCCode` | **yes** | Six digits. `DigitsOnly` strips spaces and the `*` hazardous marker. |
| `GrossKg` | **yes** | Use `TonnesToKg` if the bridge exports tonnes. |
| `TareKg` | **yes** | |
| `VehicleReg` | recommended | Needed for the reconcile step to match tickets to jobs. |
| `JobID` | if you have it | Many exports have an order or docket reference that maps here. If not, leave it blank and reconcile on vehicle and date instead. |
| `CustomerID` | optional | Only useful if the skip software uses the same account codes. |
| `OutletID` | on Out tickets | Where the material went. |
| `Weighbridge`, `Operator`, `Notes` | optional | Carried through for the audit trail. |

`NetKg` is **not** mapped. It is always recalculated from the gross and tare actually
recorded, never read from the file. If the bridge and this system ever disagree about a
weight, the arithmetic on the recorded figures is the one that can be defended.

## Transforms

| Transform | What it does |
|---|---|
| `Trim` | Strips leading and trailing spaces. The default. |
| `Upper` | Upper case. |
| `UpperNoSpace` | Upper case with spaces and hyphens removed. **Use this on the vehicle registration** - otherwise `YJ71 KLM` and `YJ71KLM` become two different lorries and the reconcile stops matching. |
| `DigitsOnly` | Keeps only the digits. For EWC codes: turns `17 09 04` and `170904*` both into `170904`. |
| `Number` | Reads a number through thousands separators, stray units and `(1,234)` negatives. |
| `TonnesToKg` | As `Number`, then multiplies by 1000. |
| `Date` / `DateTime` | Parses the date. **UK order is forced** - on a UK site `03/04` is April. |
| `Map:FROM=TO;FROM=TO` | Rewrites coded values. An unmatched value passes through unchanged so it shows up in validation rather than vanishing. |

The shipped Direction mapping is `Map:IN=In;OUT=Out;I=In;O=Out`. If your export uses
`1`/`2`, or `Delivery`/`Collection`, extend it: `Map:1=In;2=Out`.

## What the file itself can look like

The reader copes with more than a plain comma split:

- **Comma, semicolon or tab separated.** Detected from the heading line, so a bridge with
  European regional settings works without being told.
- **Quoted fields containing commas.** A note reading `"some plasterboard, quarantined"`
  stays one field. Splitting on commas breaks the first time an address or a comment
  contains one, and it breaks silently, which is worse.
- **Doubled quotes** inside a quoted field, and **newlines** inside one.
- **A UTF-8 byte order mark** on the first heading - invisible, and otherwise stops the
  first column ever matching the map.
- **Short rows** where the line ended early: the missing fields come through blank rather
  than failing the whole import.
- `.csv` and `.txt`.

## Why a row gets rejected

Rejected rows are red on WB_Staging with the reason in the Message column, and nothing
gets written until they are dealt with.

| Message | What it usually means |
|---|---|
| `Column 'X' is not in this file` | The SourceHeader does not match the export. Copy it again - trailing spaces count. |
| `No ticket number` | The TicketNo mapping is wrong, or the export has a totals row at the bottom. |
| `No usable date/time` | The date column is mapped but not parsing. Check the transform is `DateTime`, and check whether date and time are in separate columns. |
| `Direction 'X' is not In or Out` | Add the code to the `Map:` rule. |
| `EWC code is not six digits` | The column is mapped to the wrong field, or the codes genuinely are not being recorded - which is a bigger problem than the import. |
| `Gross is less than tare` | The gross and tare columns are the wrong way round. |
| `No weights` | Weights are being read as text. Try `Number`, or `TonnesToKg` if the bridge works in tonnes. |
| `Duplicate of an earlier row in this batch` | Two exports in the folder overlap. Harmless - the first one wins. |

## Testing it

`sample-data/weighbridge-export-sample.csv` is a realistic export with the headings the
shipped WB_Map already expects, and with the awkward cases in it on purpose: a quoted
comment containing a comma, an EWC code with spaces, a weight with a thousands separator,
and the same lorry written both `YJ71 KLM` and `YJ71KLM`.

Drop it in `02_Inbox\Weighbridge` and press **Import Weighbridge**. Ten rows should stage
clean. It is a good way to prove the plumbing works before pointing it at live data.

## If the export cannot be automated at all

Some older skip systems will only print. In that case type the tickets straight into the
**WeighTickets** sheet of the Operations workbook - set Source to `Manual` and use the
`WBM-` counter for the ticket number. Everything downstream (recovery rate, transfer note
tonnage, excess tonnage charges) works identically. You just lose the time the import
would have saved.
