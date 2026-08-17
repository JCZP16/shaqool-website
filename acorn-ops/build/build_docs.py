"""
Generates the reference documentation from schema.py.

Anything describing a field, a table or an ISO clause mapping is written here
rather than by hand, so the documentation cannot drift away from the workbooks
it describes. The narrative guides in docs/ are hand-written.

    python3 build_docs.py [docs_dir]
"""

import sys
from pathlib import Path

import schema
from schema import LISTS, WORKBOOKS

import build_console


KIND_NAMES = {
    "text": "text", "memo": "long text", "int": "whole number", "num": "number",
    "money": "money (GBP)", "kg": "kilograms", "tonnes": "tonnes",
    "pct": "percentage", "date": "date", "datetime": "date and time",
    "yn": "Yes / No", "list": "pick list", "formula": "calculated",
}


def data_model(path: Path):
    lines = [
        "# Data model",
        "",
        "*Generated from `build/schema.py` by `build/build_docs.py`. Do not edit by hand -",
        "change the schema and rebuild, or the documentation and the workbooks will disagree.*",
        "",
        "Three workbooks, thirty tables. Each table lives on its own sheet as a real Excel Table",
        "(a ListObject), which is what lets formulas, the sync and the macros all refer to columns",
        "by name rather than by position.",
        "",
        "**Calculated columns are not to be typed in.** They are shaded grey in the workbook and",
        "marked *calculated* below. Typing over one replaces a formula with a literal, and the next",
        "row added will silently disagree with it.",
        "",
    ]

    lines += ["## Contents", ""]
    for filename, title, tables, _purpose in WORKBOOKS:
        lines.append(f"- **{title}** (`{filename}`)")
        for t in tables:
            anchor = t.table.lower()
            lines.append(f"  - [{t.sheet}](#{anchor}) - {t.purpose.split('.')[0]}.")
    lines.append("")

    for filename, title, tables, purpose in WORKBOOKS:
        lines += [
            "---", "",
            f"# {title}",
            "",
            f"`{filename}` - {purpose}",
            "",
        ]
        for t in tables:
            lines += [
                f"## {t.table}",
                "",
                f"**Sheet:** {t.sheet}  ",
                f"**Key:** {t.key}  ",
            ]
            if t.iso:
                lines.append(f"**Reference:** {t.iso}  ")
            lines += ["", t.purpose, "", "| Field | Type | Required | Notes |",
                      "|---|---|---|---|"]
            for c in t.cols:
                kind = KIND_NAMES.get(c.kind, c.kind)
                if c.choices:
                    kind = f"pick list ({c.choices})"
                note = c.note or ""
                if c.kind == "formula":
                    note = ("Calculated. " + note).strip()
                note = note.replace("|", "\\|")
                lines.append(f"| `{c.name}` | {kind} | {'yes' if c.required else ''} | {note} |")
            lines.append("")

    lines += [
        "---", "",
        "# Pick lists",
        "",
        "Every dropdown in the workbooks is driven by one of these. They live on the `_Lists`",
        "sheet of each workbook. To add an option, insert a row *inside* an existing block so the",
        "named range grows with it - a value added below the block is not picked up.",
        "",
    ]
    for name in sorted(LISTS):
        lines.append(f"**{name}** - " + ", ".join(LISTS[name]))
        lines.append("")

    lines += [
        "---", "",
        "# Console configuration",
        "",
        "Settings on the Console's Config sheet. Nothing in the macros hard-codes a path, a rate",
        "or a threshold; it all comes from here.",
        "",
        "| Setting | Default | What it does |",
        "|---|---|---|",
    ]
    for key, value, note in build_console.CONFIG:
        if key.startswith("---"):
            lines.append(f"| **{key.strip('- ')}** | | |")
            continue
        shown = "" if value == "" else f"`{value}`"
        lines.append(f"| `{key}` | {shown} | {note.replace('|', chr(92) + '|')} |")
    lines.append("")

    path.write_text("\n".join(lines), encoding="utf-8")
    return len(lines)


def iso_mapping(path: Path):
    """Which register answers which clause. This is the page to hand an auditor."""
    clauses = [
        ("4.1 / 4.2", "Context, interested parties",
         "tblRisksOpps (Context column), tblLegalRegister",
         "Issues and interested parties are recorded as entries on the risk register."),
        ("5.2", "Policy", "tblDocRegister",
         "Signed policies live in 04_Documents\\Controlled\\01_Policies and are on the register."),
        ("6.1", "Risks and opportunities", "tblRisksOpps",
         "Scored before and after controls, with an owner and a review date."),
        ("6.1.3", "Compliance obligations", "tblLegalRegister, tblPermits",
         "What applies, how we comply, where the evidence is, and when it was last reviewed."),
        ("6.2", "Objectives", "tblObjectives",
         "Each objective carries the KPI that proves it; the recovery rate KPI comes straight "
         "off the Dashboard."),
        ("7.1.5", "Monitoring and measuring resources", "tblCalibration",
         "The weighbridge is the critical entry. Out of calibration invalidates every ticket."),
        ("7.2", "Competence", "tblTraining, tblRequiredTraining, tblStaff, tblDrivers",
         "Required training by role, actual records with calculated expiry, and licence dates."),
        ("7.3", "Awareness", "Toolbox talk records",
         "ToolboxTalkRecord.docx, filed in 05_Compliance\\Training and logged in tblTraining."),
        ("7.5", "Documented information", "tblDocRegister + Document Control Check",
         "The register is the index; the check reconciles it against what is actually in the "
         "controlled folder and flags anything overdue review."),
        ("8.1", "Operational planning and control", "tblJobs, tblMovements, tblWasteStreams, "
         "tblOutlets",
         "Waste acceptance, container movements, and the permit status of every onward outlet."),
        ("8.1 (Duty of care)", "Waste transfer", "tblTransferNotes",
         "Retention calculated automatically: two years for a transfer note, three for a "
         "hazardous consignment note. The archive routine will not move a job inside its window."),
        ("8.4", "Externally provided processes", "tblSuppliers, tblSupplierEval",
         "Approval status, carrier licence and insurance expiry, periodic scoring."),
        ("9.1", "Monitoring, measurement, analysis", "Dashboard, tblWeighTickets",
         "Tonnage in and out, recovery rate, mass balance variance, all from weighbridge tickets "
         "rather than estimates."),
        ("9.2", "Internal audit", "tblAudits",
         "Programme with planned and actual dates; an audit whose planned date has passed with "
         "no actual date appears on Alerts."),
        ("9.3", "Management review", "tblMgmtReview",
         "Inputs covered, decisions, actions raised, and the date of the next one."),
        ("10.2", "Nonconformity and corrective action", "tblNCR, tblIncidents",
         "Root cause and effectiveness verification are separate fields on purpose - an action "
         "is closed when someone has checked it worked, not when it was done."),
    ]

    lines = [
        "# ISO clause mapping",
        "",
        "*Generated by `build/build_docs.py`.*",
        "",
        "Where each clause of ISO 9001, ISO 14001 and ISO 45001 is answered. This is the page to",
        "put in front of an auditor at the opening meeting.",
        "",
        "A caveat worth stating plainly: **a register is not a management system.** These tables",
        "hold the evidence and make the gaps visible. Whether the business actually does what its",
        "procedures say is the thing being audited, and no spreadsheet settles that.",
        "",
        "| Clause | Requirement | Where it lives | How it is answered |",
        "|---|---|---|---|",
    ]
    for clause, requirement, where, how in clauses:
        lines.append(f"| {clause} | {requirement} | {where} | {how} |")

    lines += [
        "",
        "## Statutory records and how long they are kept",
        "",
        "| Record | Retention | Enforced by |",
        "|---|---|---|",
        "| Waste transfer note | 2 years | `tblTransferNotes.RetentionUntil`, calculated |",
        "| Hazardous waste consignment note | 3 years | `tblTransferNotes.RetentionUntil`, calculated |",
        "| Weighbridge tickets | Keep alongside the notes they support | Source CSVs kept in "
        "`02_Inbox\\Weighbridge\\_imported` |",
        "| Accident records (RIDDOR) | 3 years from the date of the entry | `tblIncidents` |",
        "| LOLER thorough examination reports | 2 years, or until the next report | "
        "`tblVehicles.LOLERDue` and the certificate in 05_Compliance |",
        "",
        "The Archive routine refuses to move any job folder whose transfer note is still inside",
        "its retention window, so a tidy-up cannot destroy a statutory record.",
        "",
        "## What this system does not do",
        "",
        "Worth being straight about, because an auditor will ask:",
        "",
        "- It is not the accounting system. It prices work and exports charge lines; VAT, credit",
        "  control and the statutory accounting records stay where they are.",
        "- It does not submit hazardous waste consignment notes to the regulator. It produces the",
        "  note; submission is still a separate job.",
        "- It does not do vehicle tracking, driver hours or tachograph analysis.",
        "- It holds no personal data beyond names, work contact details and training records. If",
        "  you start putting anything more sensitive in the Notes fields, that is a GDPR decision",
        "  that needs making deliberately rather than by accident.",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")
    return len(lines)


def main():
    docs = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).parent.parent / "docs")
    docs.mkdir(parents=True, exist_ok=True)
    n1 = data_model(docs / "03-data-model.md")
    n2 = iso_mapping(docs / "06-iso-mapping.md")
    tables = sum(len(t) for _f, _t, t, _p in WORKBOOKS)
    fields = sum(len(t.cols) for t in schema.all_tables())
    print(f"  wrote 03-data-model.md   {n1:>4} lines ({tables} tables, {fields} fields)")
    print(f"  wrote 06-iso-mapping.md  {n2:>4} lines")


if __name__ == "__main__":
    main()
