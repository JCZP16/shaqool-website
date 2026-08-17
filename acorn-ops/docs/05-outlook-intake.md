# Outlook intake

## What it does

**Scrape Outlook** reads a mail folder and, for each message it has not seen before:

1. Classifies it against the rules on the **Rules** sheet.
2. Works out which customer it is from - sender address, then sender domain, then a
   postcode found in the text.
3. Looks for a `JOB-nnnn` reference in the subject or body and checks it against real jobs.
4. Pulls out the postcode, phone number, container size and the date being asked for.
5. Saves the attachments into `02_Inbox\Email\<date>\<EmailID>\`.
6. Writes a row to the email log.
7. Stamps the message with a category and moves it to the processed folder.

Then it rebuilds the **Intake** sheet so somebody can work through it.

## What it does not do

It does not reply, does not delete, and does not take a booking. It reads, files and
suggests. Turning an email into a job is a decision a person makes, on the Intake sheet,
with the evidence in front of them.

There is a `AutoCreateJobFromEmail` switch on Config. Leave it off until the
classification has been reliably right for a few weeks - and probably leave it off after
that too. The cost of a wrong automatic booking is a lorry sent to the wrong address.

## Setting it up

1. In Outlook, create the folder - `Inbox\Acorn Intake` by default - and a `Logged`
   subfolder under it.
2. Add an Outlook rule that moves enquiries into it. Anything to `skips@`, or from the
   website form, or whatever suits how work actually arrives.
3. On the Console's Config sheet:

| Setting | |
|---|---|
| `OutlookAccount` | Blank for the default mailbox. For a shared mailbox, the name exactly as it appears at the very top of the Outlook folder pane. |
| `OutlookIntakeFolder` | `Inbox\Acorn Intake` |
| `OutlookProcessedFolder` | `Inbox\Acorn Intake\Logged`. Blank leaves mail where it is. |
| `OutlookLookbackDays` | 14. Older mail is ignored, which keeps the scrape quick. |
| `OutlookMarkCategory` | `Acorn: Logged` |
| `OutlookSaveAttachments` | Yes |
| `OutlookMaxBodyChars` | 1000 - how much body text to keep for searching |

Outlook has to be **running and signed in**. This drives the Outlook already on the PC; it
does not connect to the mail server itself.

## It is safe to run repeatedly

Every message is keyed twice: on Outlook's own EntryID, and on a fingerprint of the
received time, sender and subject.

The second key exists because Outlook reissues an EntryID when a message moves between
stores. Key on the ID alone and a message that has been dragged to another folder comes
back as new, and gets logged a second time. Two keys mean a message is recognised however
it has been shuffled about.

## The rules

The **Rules** sheet is applied in Priority order, lowest number first, and the first rule
that matches wins.

| | |
|---|---|
| **Priority** | Lower runs first. Leave gaps (10, 20, 30) so a rule can be slotted in later. |
| **Field** | `Subject`, `Body`, `From`, or `Any` (subject + body + sender). |
| **MatchType** | `Contains`, `StartsWith`, `EndsWith`, `Equals`. Case is always ignored. |
| **Keyword** | What to look for. |
| **Category** | One of the categories on the pick list. |
| **Action** | Free text, shown on the Intake sheet. This is what tells whoever is working the list what to do. |
| **Enabled** | Yes / No. |

The rule on priority 900 has an **empty keyword**, so it always matches. Keep it last:
it is what stops a message falling through the whole list and being logged with no
category at all.

To tune it, look at how many rows come out **Unclassified**. Read a few, find the words
they have in common, add a rule above 900. It is worth doing this properly in the first
fortnight - the classification is only ever as good as the words in that sheet.

### One rule earns its keep on its own

The `Weighbridge Ticket` category does something the others do not: if the message has a
`.csv` or `.txt` attachment, it is **copied into the weighbridge inbox** as well as being
archived. A ticket export that arrives by email is then ready for **Import Weighbridge**
without anyone dragging a file anywhere.

## What gets extracted

| Field | How |
|---|---|
| Customer | Exact sender address, then sender domain, then a postcode matching a customer record. Domain first, because people mail from their own address but the account is held against the company. |
| Job | `JOB-nnnn` in subject or body, checked against real jobs so a typo does not attach the enquiry to the wrong one. |
| Postcode | UK postcode pattern, tolerant of a missing space. |
| Phone | UK number pattern. |
| Container | `8 yard`, `8yd`, `8yrd`; `roro` / `roll on` / `hook` with a 20, 35 or 40. |
| Wanted for | An explicit date first; then `tomorrow` / `today` / `asap`; then a weekday name meaning the next one coming, with `next Tuesday` a further week out. |

Anything less certain than that is left blank rather than guessed. A wrong date on a
booking is worse than no date, because no date gets asked about and a wrong one does not.

## Attachments

Saved to `02_Inbox\Email\<yyyy-mm-dd>\<EmailID>\` under their original name, with
characters Windows will not accept replaced.

Signature images are skipped - anything named `imageNNN`, or containing `logo` or
`signature`, with an image extension. Without that filter the archive fills up with other
people's letterheads and the genuine attachments get lost in it.

## Troubleshooting

**"Outlook is not available on this PC"** - Outlook is not running, or Excel is running
elevated and Outlook is not (or the other way round). Two processes at different
elevations cannot talk to each other. Run both normally.

**"Cannot find the mailbox 'X'"** - `OutlookAccount` must match the name at the very top
of the folder pane, exactly. Leave it blank for the default mailbox.

**"Outlook has no folder 'X'"** - the path is relative to the mailbox root, so a folder
under the inbox is `Inbox\Acorn Intake`, not `Acorn Intake`.

**Nothing comes through but there is clearly mail there** - check
`OutlookLookbackDays`, and check the messages have not already been logged. A message that
was logged once is skipped forever, by design.

**Customer never matches** - the customer's `Email` field in the Master workbook is what
this matches against. If accounts@ mails you but the record holds the sales contact's
address, the domain match should still catch it; if the record is blank, nothing will.
