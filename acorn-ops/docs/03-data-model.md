# Data model

*Generated from `build/schema.py` by `build/build_docs.py`. Do not edit by hand -
change the schema and rebuild, or the documentation and the workbooks will disagree.*

Three workbooks, thirty tables. Each table lives on its own sheet as a real Excel Table
(a ListObject), which is what lets formulas, the sync and the macros all refer to columns
by name rather than by position.

**Calculated columns are not to be typed in.** They are shaded grey in the workbook and
marked *calculated* below. Typing over one replaces a formula with a literal, and the next
row added will silently disagree with it.

## Contents

- **Master Data** (`AcornOps_Master.xlsx`)
  - [Customers](#tblcustomers) - One row per trading account.
  - [Sites](#tblsites) - Where containers actually go.
  - [Assets](#tblassets) - Every skip, RoRo and bin we own.
  - [Vehicles](#tblvehicles) - Fleet compliance dates.
  - [Drivers](#tbldrivers) - Driver competence and licence expiry.
  - [Staff](#tblstaff) - Everyone who needs a training record or a document to be issued to them.
  - [WasteStreams](#tblwastestreams) - EWC code master.
  - [Outlets](#tbloutlets) - Where material goes after us.
  - [PriceList](#tblpricelist) - Rate card.
  - [Suppliers](#tblsuppliers) - Anyone who supplies us or hauls for us.
- **Operations** (`AcornOps_Operations.xlsx`)
  - [Jobs](#tbljobs) - The spine of the system.
  - [Movements](#tblmovements) - Audit trail of where each container physically went and when.
  - [WeighTickets](#tblweightickets) - Every weighbridge event, imported from the skip software's CSV export or keyed by hand.
  - [TransferNotes](#tbltransfernotes) - Statutory duty-of-care records.
  - [InvoiceLines](#tblinvoicelines) - Charge lines built from completed jobs and weighed tonnage, ready to export to the accounts package.
  - [EmailLog](#tblemaillog) - One row per email harvested from Outlook.
  - [DocsIssued](#tbldocsissued) - Every Word/PDF the platform produced, so a document can always be traced back to its data.
- **Compliance** (`AcornOps_Compliance.xlsx`)
  - [DocRegister](#tbldocregister) - Master index of the management system.
  - [NCR](#tblncr) - Nonconformities, their root cause and the action taken.
  - [Audits](#tblaudits) - Planned and completed audits across all three standards.
  - [Training](#tbltraining) - Completed training.
  - [RequiredTraining](#tblrequiredtraining) - Drives the Training Matrix sheet: which roles must hold which competences.
  - [Calibration](#tblcalibration) - Calibration control.
  - [Permits](#tblpermits) - Every licence that would stop the site trading if it lapsed.
  - [Incidents](#tblincidents) - Health, safety and environmental events.
  - [LegalRegister](#tbllegalregister) - What law applies to us, how we comply, and where the evidence is.
  - [Objectives](#tblobjectives) - Measurable objectives per standard, with the KPI that proves them.
  - [RisksOpps](#tblrisksopps) - Clause 6.
  - [MgmtReview](#tblmgmtreview) - Clause 9.
  - [SupplierEval](#tblsuppliereval) - Periodic scoring of approved suppliers.

---

# Master Data

`AcornOps_Master.xlsx` - Slow-moving reference data: who we deal with, what we own, what we charge.

## tblCustomers

**Sheet:** Customers  
**Key:** CustomerID  

One row per trading account. The single source of truth for who we invoice.

| Field | Type | Required | Notes |
|---|---|---|---|
| `CustomerID` | text | yes | ACC-0001 style. Allocated by the Console, never reused. |
| `CustomerName` | text | yes |  |
| `CustomerType` | pick list (CustomerType) | yes |  |
| `ContactName` | text |  |  |
| `Email` | text |  | Used to match inbound Outlook mail to an account. |
| `Phone` | text |  |  |
| `Address1` | text |  |  |
| `Address2` | text |  |  |
| `Town` | text |  |  |
| `Postcode` | text |  | Second matching key for Outlook intake. |
| `VATNumber` | text |  |  |
| `PaymentTerms` | pick list (PaymentTerms) |  |  |
| `CreditLimit` | money (GBP) |  |  |
| `OnStop` | pick list (YesNo) |  | Yes blocks new bookings on the Job Entry sheet. |
| `Status` | pick list (RecordStatus) |  |  |
| `Notes` | long text |  |  |
| `CreatedOn` | date |  |  |
| `CreatedBy` | text |  |  |

## tblSites

**Sheet:** Sites  
**Key:** SiteID  

Where containers actually go. One customer may have many sites.

| Field | Type | Required | Notes |
|---|---|---|---|
| `SiteID` | text | yes |  |
| `CustomerID` | text | yes | Must exist in tblCustomers. |
| `SiteName` | text | yes |  |
| `Address1` | text |  |  |
| `Town` | text |  |  |
| `Postcode` | text | yes |  |
| `What3Words` | text |  |  |
| `AccessNotes` | long text |  | Height restrictions, gate codes, keep-clear times. |
| `PermitRequired` | pick list (YesNo) |  | Highway/footway skip permit. |
| `PermitAuthority` | text |  |  |
| `SiteContact` | text |  |  |
| `SiteContactPhone` | text |  |  |
| `Status` | pick list (RecordStatus) |  |  |

## tblAssets

**Sheet:** Assets  
**Key:** AssetID  

Every skip, RoRo and bin we own. Drives the on-hire / available count.

| Field | Type | Required | Notes |
|---|---|---|---|
| `AssetID` | text | yes |  |
| `AssetRef` | text |  | Number painted on the container. |
| `ContainerType` | pick list (ContainerType) | yes |  |
| `SizeYd3` | number |  |  |
| `AcquiredOn` | date |  |  |
| `Status` | pick list (AssetStatus) | yes |  |
| `CurrentSiteID` | text |  | Blank = in the yard. |
| `CurrentJobID` | text |  |  |
| `OnHireSince` | date |  | Feeds the long-standing-hire alert. |
| `LastInspection` | date |  |  |
| `InspectionMonths` | whole number |  |  |
| `NextInspection` | calculated |  | Calculated. |
| `Notes` | long text |  |  |

## tblVehicles

**Sheet:** Vehicles  
**Key:** VehicleID  
**Reference:** ISO 45001 7.1 / O-Licence undertakings  

Fleet compliance dates. Everything dated here appears on the Alerts sheet.

| Field | Type | Required | Notes |
|---|---|---|---|
| `VehicleID` | text | yes |  |
| `Registration` | text | yes |  |
| `VehicleType` | text |  | Skip Loader / Hook Loader / RCV / Van |
| `MakeModel` | text |  |  |
| `GrossWeightKg` | kilograms |  |  |
| `UnladenWeightKg` | kilograms |  | Cross-check against weighbridge tare. |
| `TaxDue` | date |  |  |
| `MOTDue` | date |  |  |
| `InsuranceDue` | date |  |  |
| `LOLERDue` | date |  | 6-monthly on lifting equipment. |
| `TachoCalDue` | date |  |  |
| `ServiceDue` | date |  |  |
| `OLicenceRef` | text |  |  |
| `Status` | pick list (RecordStatus) |  |  |

## tblDrivers

**Sheet:** Drivers  
**Key:** DriverID  
**Reference:** ISO 45001 7.2  

Driver competence and licence expiry. Feeds Alerts and the training matrix.

| Field | Type | Required | Notes |
|---|---|---|---|
| `DriverID` | text | yes |  |
| `StaffID` | text |  | Links to tblStaff for the training matrix. |
| `FullName` | text | yes |  |
| `Mobile` | text |  |  |
| `LicenceNumber` | text |  |  |
| `LicenceExpiry` | date |  |  |
| `CPCExpiry` | date |  |  |
| `ADRExpiry` | date |  | Only if they move hazardous loads. |
| `MedicalDue` | date |  |  |
| `StartDate` | date |  |  |
| `Status` | pick list (RecordStatus) |  |  |

## tblStaff

**Sheet:** Staff  
**Key:** StaffID  

Everyone who needs a training record or a document to be issued to them.

| Field | Type | Required | Notes |
|---|---|---|---|
| `StaffID` | text | yes |  |
| `FullName` | text | yes |  |
| `Role` | text | yes | Must match the Role column of tblRequiredTraining. |
| `Department` | text |  |  |
| `Email` | text |  |  |
| `StartDate` | date |  |  |
| `Status` | pick list (RecordStatus) |  |  |

## tblWasteStreams

**Sheet:** WasteStreams  
**Key:** EWCCode  
**Reference:** ISO 14001 6.1.2 / Duty of Care  

EWC code master. The RecoveryRoute column is what turns tonnage into a recycling rate.

| Field | Type | Required | Notes |
|---|---|---|---|
| `EWCCode` | text | yes | Six digits, no spaces, e.g. 170904. |
| `Description` | text | yes |  |
| `StreamGroup` | pick list (StreamGroup) | yes |  |
| `Hazardous` | pick list (YesNo) | yes | Yes forces a Hazardous Consignment Note instead of a WTN. |
| `RecoveryRoute` | pick list (RecoveryRoute) | yes | Recycling/Reuse/Recovery count as diverted; Landfill does not. |
| `DefaultOutletID` | text |  |  |
| `LandfillTaxBand` | text |  | Standard / Lower / Exempt |
| `Notes` | long text |  |  |

## tblOutlets

**Sheet:** Outlets  
**Key:** OutletID  
**Reference:** ISO 14001 8.1  

Where material goes after us. Permit expiry here is a duty-of-care control.

| Field | Type | Required | Notes |
|---|---|---|---|
| `OutletID` | text | yes |  |
| `OutletName` | text | yes |  |
| `OutletType` | pick list (OutletType) | yes |  |
| `PermitNumber` | text | yes |  |
| `PermitExpiry` | date |  | Blank = permit does not expire; still verify annually. |
| `PermitVerifiedOn` | date |  |  |
| `Address` | text |  |  |
| `Postcode` | text |  |  |
| `AcceptedEWC` | long text |  | Comma separated, or ALL. |
| `GateFeePerTonne` | money (GBP) |  |  |
| `Contact` | text |  |  |
| `Status` | pick list (ApprovalStatus) |  |  |

## tblPriceList

**Sheet:** PriceList  
**Key:** PriceID  

Rate card. A row with a CustomerID beats a row without, so you can hold account-specific rates.

| Field | Type | Required | Notes |
|---|---|---|---|
| `PriceID` | text | yes |  |
| `EffectiveFrom` | date | yes |  |
| `EffectiveTo` | date |  | Blank = still current. |
| `CustomerID` | text |  | Blank = standard rate for everybody. |
| `ServiceType` | pick list (ServiceType) | yes |  |
| `ContainerType` | pick list (ContainerType) |  |  |
| `WasteGroup` | pick list (StreamGroup) |  |  |
| `HireDaysIncluded` | whole number |  |  |
| `Price` | money (GBP) | yes |  |
| `ExcessPerDay` | money (GBP) |  |  |
| `TonnageIncluded` | tonnes |  |  |
| `ExcessPerTonne` | money (GBP) |  |  |
| `Notes` | long text |  |  |

## tblSuppliers

**Sheet:** Suppliers  
**Key:** SupplierID  
**Reference:** ISO 9001 8.4  

Anyone who supplies us or hauls for us. Approval status is an ISO 9001 8.4 control.

| Field | Type | Required | Notes |
|---|---|---|---|
| `SupplierID` | text | yes |  |
| `SupplierName` | text | yes |  |
| `Category` | text |  | Haulage / Plant / Parts / PPE / Disposal / Professional |
| `Contact` | text |  |  |
| `Email` | text |  |  |
| `Phone` | text |  |  |
| `ApprovalStatus` | pick list (ApprovalStatus) | yes |  |
| `CarrierLicence` | text |  | Mandatory for anyone moving our waste. |
| `CarrierLicExpiry` | date |  |  |
| `InsuranceExpiry` | date |  |  |
| `LastReviewDate` | date |  |  |
| `ReviewMonths` | whole number |  |  |
| `NextReviewDue` | calculated |  | Calculated. |
| `Notes` | long text |  |  |

---

# Operations

`AcornOps_Operations.xlsx` - Day-to-day transactions: jobs, movements, weighbridge tickets, duty-of-care notes, billing lines.

## tblJobs

**Sheet:** Jobs  
**Key:** JobID  

The spine of the system. One row per container movement request.

| Field | Type | Required | Notes |
|---|---|---|---|
| `JobID` | text | yes |  |
| `CreatedOn` | date and time |  |  |
| `CustomerID` | text | yes |  |
| `SiteID` | text | yes |  |
| `ServiceType` | pick list (ServiceType) | yes |  |
| `ContainerType` | pick list (ContainerType) | yes |  |
| `AssetID` | text |  |  |
| `RequestedDate` | date | yes |  |
| `ScheduledDate` | date |  |  |
| `CompletedDate` | date |  |  |
| `Status` | pick list (JobStatus) | yes |  |
| `DriverID` | text |  |  |
| `VehicleReg` | text |  |  |
| `PermitRef` | text |  | Council skip permit number where the site needs one. |
| `PONumber` | text |  |  |
| `PriceAgreed` | money (GBP) |  |  |
| `EWCCode` | text |  |  |
| `WasteDescription` | text |  |  |
| `DaysOnHire` | calculated |  | Calculated. Live counter until the job is completed. |
| `SourceRef` | text |  | EmailID if the job came in from Outlook intake. |
| `Notes` | long text |  |  |

## tblMovements

**Sheet:** Movements  
**Key:** MovementID  

Audit trail of where each container physically went and when.

| Field | Type | Required | Notes |
|---|---|---|---|
| `MovementID` | text | yes |  |
| `JobID` | text | yes |  |
| `MovementType` | pick list (MovementType) | yes |  |
| `MovementDateTime` | date and time | yes |  |
| `AssetID` | text | yes |  |
| `FromLocation` | text |  |  |
| `ToLocation` | text |  |  |
| `DriverID` | text |  |  |
| `VehicleReg` | text |  |  |
| `Notes` | long text |  |  |

## tblWeighTickets

**Sheet:** WeighTickets  
**Key:** TicketNo  
**Reference:** ISO 14001 9.1 / Weights & Measures  

Every weighbridge event, imported from the skip software's CSV export or keyed by hand. NetKg is calculated here, not trusted from the import.

| Field | Type | Required | Notes |
|---|---|---|---|
| `TicketNo` | text | yes | Unique. Re-importing the same ticket updates rather than duplicates. |
| `TicketDateTime` | date and time | yes |  |
| `Direction` | pick list (Direction) | yes | In = material arriving on site. Out = material leaving to an outlet. |
| `VehicleReg` | text |  |  |
| `JobID` | text |  | Blank until reconciled. |
| `CustomerID` | text |  |  |
| `OutletID` | text |  | Populated on Direction = Out. |
| `EWCCode` | text | yes |  |
| `GrossKg` | kilograms |  |  |
| `TareKg` | kilograms |  |  |
| `NetKg` | calculated |  | Calculated. |
| `NetTonnes` | calculated |  | Calculated. |
| `Weighbridge` | text |  |  |
| `Operator` | text |  |  |
| `Source` | text |  | Import / Manual / Adjustment |
| `ImportBatch` | text |  |  |
| `Matched` | pick list (YesNo) |  | No = appears on the Unmatched Tickets report. |
| `Notes` | long text |  |  |

## tblTransferNotes

**Sheet:** TransferNotes  
**Key:** WTNRef  
**Reference:** Environmental Protection Act 1990 s.34 / ISO 14001 8.1  

Statutory duty-of-care records. WTNs must be kept 2 years, hazardous consignment notes 3 years - the Archive routine will not delete anything inside those windows.

| Field | Type | Required | Notes |
|---|---|---|---|
| `WTNRef` | text | yes |  |
| `DocType` | pick list (WTNDocType) | yes |  |
| `IssueDate` | date | yes |  |
| `JobID` | text |  |  |
| `CustomerID` | text | yes |  |
| `SiteID` | text | yes |  |
| `EWCCode` | text | yes |  |
| `WasteDescription` | text | yes |  |
| `SICCode` | text |  | Producer's SIC 2007 code - legally required on the note. |
| `ContainerType` | pick list (ContainerType) |  |  |
| `QuantityTonnes` | tonnes |  |  |
| `CarrierLicence` | text |  |  |
| `ProducerSignedBy` | text |  |  |
| `ProducerSignedOn` | date |  |  |
| `CarrierSignedBy` | text |  |  |
| `DocumentPath` | text |  |  |
| `RetentionUntil` | calculated |  | Calculated. |
| `Status` | pick list (DocStatus) |  |  |

## tblInvoiceLines

**Sheet:** InvoiceLines  
**Key:** LineID  

Charge lines built from completed jobs and weighed tonnage, ready to export to the accounts package. This platform prices work; it is not the accounting system of record.

| Field | Type | Required | Notes |
|---|---|---|---|
| `LineID` | text | yes |  |
| `InvoiceNo` | text |  |  |
| `InvoiceDate` | date |  |  |
| `CustomerID` | text | yes |  |
| `JobID` | text |  |  |
| `Description` | text | yes |  |
| `Qty` | number | yes |  |
| `UOM` | pick list (UOM) |  |  |
| `UnitPrice` | money (GBP) | yes |  |
| `NetAmount` | calculated |  | Calculated. |
| `VATRate` | percentage |  | Standard rate lives on the Console Config sheet. |
| `VATAmount` | calculated |  | Calculated. |
| `GrossAmount` | calculated |  | Calculated. |
| `Status` | pick list (InvoiceStatus) |  |  |
| `ExportedOn` | date |  |  |

## tblEmailLog

**Sheet:** EmailLog  
**Key:** EmailID  

One row per email harvested from Outlook. OutlookEntryID makes the scrape idempotent - running it twice never double-logs a message.

| Field | Type | Required | Notes |
|---|---|---|---|
| `EmailID` | text | yes |  |
| `ReceivedOn` | date and time | yes |  |
| `SenderName` | text |  |  |
| `SenderEmail` | text |  |  |
| `Subject` | text |  |  |
| `Category` | pick list (EmailCategory) |  | Set by the rules on the Console Rules sheet. |
| `MatchedCustomerID` | text |  | Matched on sender domain, then postcode. |
| `MatchedJobID` | text |  | Matched on a JOB-nnnn reference in the subject or body. |
| `ExtractedPostcode` | text |  |  |
| `ExtractedPhone` | text |  |  |
| `ExtractedContainer` | text |  |  |
| `ExtractedDate` | date |  |  |
| `AttachmentCount` | whole number |  |  |
| `AttachmentFolder` | text |  |  |
| `OutlookEntryID` | text |  | Outlook's own key. Do not edit. |
| `Processed` | pick list (YesNo) |  |  |
| `ProcessedOn` | date and time |  |  |
| `ActionTaken` | text |  |  |
| `BodyExtract` | long text |  | First 1,000 characters, for searching. |

## tblDocsIssued

**Sheet:** DocsIssued  
**Key:** IssueID  
**Reference:** ISO 9001 7.5.3  

Every Word/PDF the platform produced, so a document can always be traced back to its data.

| Field | Type | Required | Notes |
|---|---|---|---|
| `IssueID` | text | yes |  |
| `GeneratedOn` | date and time | yes |  |
| `TemplateName` | text | yes |  |
| `DocType` | text |  |  |
| `JobID` | text |  |  |
| `CustomerID` | text |  |  |
| `FilePath` | text |  |  |
| `GeneratedBy` | text |  |  |

---

# Compliance

`AcornOps_Compliance.xlsx` - The ISO 9001 / 14001 / 45001 management system registers.

## tblDocRegister

**Sheet:** DocRegister  
**Key:** DocID  
**Reference:** ISO 9001 / 14001 / 45001 clause 7.5  

Master index of the management system. Nothing is 'controlled' unless it is on this list.

| Field | Type | Required | Notes |
|---|---|---|---|
| `DocID` | text | yes | e.g. AR-QP-04 |
| `DocTitle` | text | yes |  |
| `DocType` | pick list (DocType) | yes |  |
| `Standard` | pick list (Standard) |  |  |
| `ISOClause` | text |  |  |
| `Owner` | text | yes |  |
| `Revision` | text | yes |  |
| `IssueDate` | date | yes |  |
| `ReviewMonths` | whole number |  |  |
| `ReviewDue` | calculated |  | Calculated. |
| `Status` | pick list (DocStatus) | yes |  |
| `Location` | text |  | Path relative to the AcornOps root. |
| `SupersededBy` | text |  |  |
| `Notes` | long text |  |  |

## tblNCR

**Sheet:** NCR  
**Key:** NCRID  
**Reference:** ISO 9001 10.2 / 14001 10.2 / 45001 10.2  

Nonconformities, their root cause and the action taken. The heart of an ISO audit.

| Field | Type | Required | Notes |
|---|---|---|---|
| `NCRID` | text | yes |  |
| `RaisedOn` | date | yes |  |
| `RaisedBy` | text | yes |  |
| `Source` | pick list (NCRSource) | yes |  |
| `Type` | pick list (NCRType) | yes |  |
| `Severity` | pick list (Severity) | yes |  |
| `Description` | long text | yes |  |
| `LinkedJobID` | text |  |  |
| `LinkedCustomerID` | text |  |  |
| `ImmediateCorrection` | long text |  |  |
| `RootCause` | long text |  | Say why it happened, not what happened. |
| `CorrectiveAction` | long text |  |  |
| `ActionOwner` | text | yes |  |
| `DueDate` | date | yes |  |
| `ClosedOn` | date |  |  |
| `EffectivenessCheck` | long text |  | Verified how, by whom, when. |
| `Status` | pick list (OpenClosed) | yes |  |
| `DaysOpen` | calculated |  | Calculated. |
| `Overdue` | calculated |  | Calculated. |

## tblAudits

**Sheet:** Audits  
**Key:** AuditID  
**Reference:** ISO 9001 9.2  

Planned and completed audits across all three standards.

| Field | Type | Required | Notes |
|---|---|---|---|
| `AuditID` | text | yes |  |
| `AuditType` | pick list (AuditType) | yes |  |
| `Standard` | pick list (Standard) | yes |  |
| `Scope` | text | yes |  |
| `PlannedDate` | date | yes |  |
| `ActualDate` | date |  |  |
| `Auditor` | text |  |  |
| `Auditee` | text |  |  |
| `FindingsMajor` | whole number |  |  |
| `FindingsMinor` | whole number |  |  |
| `Observations` | whole number |  |  |
| `ReportPath` | text |  |  |
| `Status` | pick list (OpenClosed) | yes |  |

## tblTraining

**Sheet:** Training  
**Key:** TrainingID  
**Reference:** ISO 9001 7.2 / 45001 7.2  

Completed training. Expiry is calculated, so refreshers surface on Alerts automatically.

| Field | Type | Required | Notes |
|---|---|---|---|
| `TrainingID` | text | yes |  |
| `StaffID` | text | yes |  |
| `StaffName` | text |  |  |
| `Course` | text | yes | Spell it exactly as in tblRequiredTraining. |
| `Provider` | text |  |  |
| `CompletedOn` | date | yes |  |
| `ValidMonths` | whole number |  | Blank = does not expire. |
| `ExpiresOn` | calculated |  | Calculated. |
| `CertificatePath` | text |  |  |
| `Cost` | money (GBP) |  |  |

## tblRequiredTraining

**Sheet:** RequiredTraining  
**Key:** ReqID  
**Reference:** ISO 45001 7.2  

Drives the Training Matrix sheet: which roles must hold which competences.

| Field | Type | Required | Notes |
|---|---|---|---|
| `ReqID` | text | yes |  |
| `Role` | text | yes | Must match tblStaff.Role exactly. |
| `Course` | text | yes |  |
| `Mandatory` | pick list (YesNo) | yes |  |
| `RefreshMonths` | whole number |  |  |
| `Notes` | long text |  |  |

## tblCalibration

**Sheet:** Calibration  
**Key:** EquipmentID  
**Reference:** ISO 9001 7.1.5  

Calibration control. The weighbridge belongs here - an out-of-calibration bridge invalidates every ticket and every tonne you have invoiced.

| Field | Type | Required | Notes |
|---|---|---|---|
| `EquipmentID` | text | yes |  |
| `Description` | text | yes |  |
| `Location` | text |  |  |
| `SerialNo` | text |  |  |
| `Provider` | text |  |  |
| `LastCalibration` | date | yes |  |
| `FrequencyMonths` | whole number | yes |  |
| `NextDue` | calculated |  | Calculated. |
| `CertificatePath` | text |  |  |
| `Status` | pick list (RecordStatus) |  |  |

## tblPermits

**Sheet:** Permits  
**Key:** PermitID  
**Reference:** ISO 14001 6.1.3  

Every licence that would stop the site trading if it lapsed.

| Field | Type | Required | Notes |
|---|---|---|---|
| `PermitID` | text | yes |  |
| `PermitType` | pick list (PermitType) | yes |  |
| `Reference` | text | yes |  |
| `Holder` | text |  |  |
| `IssuingBody` | text |  |  |
| `IssueDate` | date |  |  |
| `ExpiryDate` | date |  | Blank = does not expire, but still verify annually. |
| `RenewalLeadDays` | whole number |  | How far ahead the alert should fire. Default 90. |
| `AnnualCost` | money (GBP) |  |  |
| `DocumentPath` | text |  |  |
| `Status` | pick list (RecordStatus) | yes |  |
| `Notes` | long text |  |  |

## tblIncidents

**Sheet:** Incidents  
**Key:** IncidentID  
**Reference:** ISO 45001 10.2 / RIDDOR 2013  

Health, safety and environmental events. Near misses matter as much as injuries.

| Field | Type | Required | Notes |
|---|---|---|---|
| `IncidentID` | text | yes |  |
| `IncidentDateTime` | date and time | yes |  |
| `Type` | pick list (IncidentType) | yes |  |
| `Location` | text | yes |  |
| `PersonInvolved` | text |  |  |
| `Description` | long text | yes |  |
| `RIDDORReportable` | pick list (YesNo) | yes |  |
| `RIDDORRef` | text |  |  |
| `ReportedOn` | date |  |  |
| `LostTimeDays` | whole number |  |  |
| `ImmediateAction` | long text |  |  |
| `RootCause` | long text |  |  |
| `ActionsTaken` | long text |  |  |
| `LinkedNCRID` | text |  |  |
| `ClosedOn` | date |  |  |
| `Status` | pick list (OpenClosed) | yes |  |

## tblLegalRegister

**Sheet:** LegalRegister  
**Key:** RegID  
**Reference:** ISO 14001 6.1.3 / 45001 6.1.3  

What law applies to us, how we comply, and where the evidence is.

| Field | Type | Required | Notes |
|---|---|---|---|
| `RegID` | text | yes |  |
| `Legislation` | text | yes |  |
| `Applicability` | text | yes |  |
| `Requirement` | long text | yes |  |
| `HowWeComply` | long text | yes |  |
| `Evidence` | text |  | Point at a DocID, a register or a file path. |
| `Owner` | text |  |  |
| `LastReviewed` | date | yes |  |
| `ReviewMonths` | whole number |  |  |
| `NextReview` | calculated |  | Calculated. |
| `ComplianceStatus` | pick list (Compliance) | yes |  |

## tblObjectives

**Sheet:** Objectives  
**Key:** ObjID  
**Reference:** ISO 9001 6.2 / 14001 6.2 / 45001 6.2  

Measurable objectives per standard, with the KPI that proves them.

| Field | Type | Required | Notes |
|---|---|---|---|
| `ObjID` | text | yes |  |
| `Standard` | pick list (Standard) | yes |  |
| `Objective` | text | yes |  |
| `KPI` | text | yes |  |
| `Baseline` | number |  |  |
| `Target` | number | yes |  |
| `CurrentValue` | number |  |  |
| `Progress` | calculated |  | Calculated. |
| `TargetDate` | date | yes |  |
| `Owner` | text | yes |  |
| `Status` | pick list (OpenClosed) | yes |  |
| `ReviewNotes` | long text |  |  |

## tblRisksOpps

**Sheet:** RisksOpps  
**Key:** RiskID  
**Reference:** ISO 9001 6.1 / 14001 6.1 / 45001 6.1  

Clause 6.1 risk register, scored before and after controls.

| Field | Type | Required | Notes |
|---|---|---|---|
| `RiskID` | text | yes |  |
| `Type` | pick list (RiskType) | yes |  |
| `Category` | text | yes | Operational / Environmental / H&S / Commercial / Legal |
| `Description` | long text | yes |  |
| `Likelihood` | pick list (LikertScore) | yes |  |
| `Impact` | pick list (LikertScore) | yes |  |
| `GrossScore` | calculated |  | Calculated. |
| `Controls` | long text | yes |  |
| `ResLikelihood` | pick list (LikertScore) |  |  |
| `ResImpact` | pick list (LikertScore) |  |  |
| `NetScore` | calculated |  | Calculated. |
| `Owner` | text | yes |  |
| `ReviewDate` | date | yes |  |
| `Status` | pick list (OpenClosed) | yes |  |

## tblMgmtReview

**Sheet:** MgmtReview  
**Key:** MRID  
**Reference:** ISO 9001 9.3  

Clause 9.3 review meetings, their inputs and their outputs.

| Field | Type | Required | Notes |
|---|---|---|---|
| `MRID` | text | yes |  |
| `ReviewDate` | date | yes |  |
| `Chair` | text | yes |  |
| `Attendees` | long text | yes |  |
| `InputsCovered` | long text | yes | Audit results, customer feedback, KPI performance, NCR status, risks, resources. |
| `Decisions` | long text | yes |  |
| `ActionsRaised` | long text |  |  |
| `MinutesPath` | text |  |  |
| `NextReviewDue` | date |  |  |

## tblSupplierEval

**Sheet:** SupplierEval  
**Key:** EvalID  
**Reference:** ISO 9001 8.4.1  

Periodic scoring of approved suppliers. Feeds the approval status in tblSuppliers.

| Field | Type | Required | Notes |
|---|---|---|---|
| `EvalID` | text | yes |  |
| `SupplierID` | text | yes |  |
| `SupplierName` | text |  |  |
| `EvalDate` | date | yes |  |
| `ScoreQuality` | pick list (LikertScore) | yes |  |
| `ScoreDelivery` | pick list (LikertScore) | yes |  |
| `ScoreCompliance` | pick list (LikertScore) | yes | Licences, permits and insurance in date. |
| `ScorePrice` | pick list (LikertScore) | yes |  |
| `ScoreHSE` | pick list (LikertScore) | yes |  |
| `TotalScore` | calculated |  | Calculated. |
| `Outcome` | calculated |  | Calculated. 20+ approved, 14-19 provisional, below 14 goes back for review. |
| `Reviewer` | text | yes |  |
| `Comments` | long text |  |  |
| `NextEvalDue` | date |  |  |

---

# Pick lists

Every dropdown in the workbooks is driven by one of these. They live on the `_Lists`
sheet of each workbook. To add an option, insert a row *inside* an existing block so the
named range grows with it - a value added below the block is not picked up.

**ApprovalStatus** - Approved, Provisional, Under Review, Suspended, Rejected

**AssetStatus** - Available, On Hire, In Transit, Maintenance, Repair, Lost, Scrapped

**AuditType** - Internal, Certification, Surveillance, Regulator, Customer, Supplier

**Compliance** - Compliant, Partially Compliant, Non-Compliant, Not Applicable, Under Assessment

**ContainerType** - 2yd Skip, 4yd Skip, 6yd Skip, 8yd Skip, 12yd Skip, 14yd Skip, 20yd RoRo, 35yd RoRo, 40yd RoRo, Front End Loader, Wheelie Bin 1100L

**CustomerType** - Trade, Domestic, Local Authority, Main Contractor, Broker, Internal

**Direction** - In, Out

**DocStatus** - Draft, In Review, Issued, Superseded, Withdrawn

**DocType** - Policy, Manual, Procedure, Work Instruction, Form, Register, Record, Plan, External Document

**EmailCategory** - Booking, Enquiry, Quote Request, Weighbridge Ticket, Invoice Query, Complaint, Supplier, Compliance, Internal, Spam, Unclassified

**IncidentType** - Injury, Near Miss, Environmental, Spill, Fire, Road Traffic Collision, Property Damage, Dangerous Occurrence, Ill Health

**InvoiceStatus** - Draft, Approved, Exported, Invoiced, Paid, Credited, Disputed

**JobStatus** - Enquiry, Quoted, Booked, Allocated, On Site, Awaiting Collection, Completed, Invoiced, Cancelled, Aborted

**LikertScore** - 1, 2, 3, 4, 5

**MovementType** - Delivered, Exchanged Out, Exchanged In, Collected, Returned to Yard, Transferred, Off Hire

**NCRSource** - Internal Audit, External Audit, Customer Complaint, Driver Report, Weighbridge, Site Inspection, Supplier, Incident, Management Review

**NCRType** - Quality, Environmental, Health & Safety, Compliance, Commercial

**OpenClosed** - Open, In Progress, Awaiting Verification, Closed, Cancelled

**OutletType** - Materials Recovery Facility, Transfer Station, Reprocessor, EfW, Landfill, Inert Recovery, Composting, Metal Merchant, Hazardous Facility

**PaymentTerms** - Pro Forma, Card on Booking, 7 Days, 14 Days, 30 Days, 30 Days EOM, 60 Days

**PermitType** - Environmental Permit, Waste Carrier Licence, Waste Broker Registration, Operator Licence, Planning Consent, Trade Effluent Consent, Duty of Care Exemption, Insurance

**RecordStatus** - Active, Inactive, Archived

**RecoveryRoute** - Recycling, Reuse, Recovery (Energy), Treatment, Landfill

**RiskType** - Risk, Opportunity

**ServiceType** - Delivery, Exchange, Collection, Wait and Load, Tip and Return, Muckaway, Aborted Visit

**Severity** - Major, Minor, Observation

**Standard** - ISO 9001, ISO 14001, ISO 45001, Multiple, Statutory, Internal

**StreamGroup** - Mixed C&D, Inert, Soil, Wood, Metal, Plasterboard, Green Waste, Mixed General, Cardboard, Plastics, WEEE, Hazardous, Fines/Residue

**UOM** - Job, Tonne, Day, Load, Item, Hour, Mile

**WTNDocType** - Waste Transfer Note, Season Ticket (Annual WTN), Hazardous Consignment Note

**YesNo** - Yes, No

---

# Console configuration

Settings on the Console's Config sheet. Nothing in the macros hard-codes a path, a rate
or a threshold; it all comes from here.

| Setting | Default | What it does |
|---|---|---|
| **Paths** | | |
| `RootPath` | `C:\AcornOps` | The one setting you must get right. Everything else hangs off it. On a shared drive use the UNC path (\\SERVER\Acorn\AcornOps), not a mapped letter - drive letters differ between PCs. |
| `MasterWorkbook` | `01_Data\AcornOps_Master.xlsx` | Relative to RootPath. |
| `OperationsWorkbook` | `01_Data\AcornOps_Operations.xlsx` | Relative to RootPath. |
| `ComplianceWorkbook` | `01_Data\AcornOps_Compliance.xlsx` | Relative to RootPath. |
| `TemplatesFolder` | `04_Documents\Templates` | Word templates with {{Token}} placeholders. |
| `JobDocsFolder` | `04_Documents\Jobs` | One subfolder per job, created on demand. |
| `ControlledDocsFolder` | `04_Documents\Controlled` | Scanned by the document control check. |
| `WeighbridgeInbox` | `02_Inbox\Weighbridge` | Drop the weighbridge CSV exports here. |
| `EmailInbox` | `02_Inbox\Email` | Outlook attachments land here, foldered by EmailID. |
| `ArchiveFolder` | `06_Archive` | Where the archive routine moves closed years. |
| `LogFolder` | `00_Admin\Logs` | One log file per day. Every automated action is recorded. |
| **Outlook intake** | | |
| `OutlookAccount` |  | Blank = the default mailbox. For a shared mailbox put its display name exactly as it appears in the Outlook folder pane. |
| `OutlookIntakeFolder` | `Inbox\Acorn Intake` | Backslash-separated path below the mailbox root. Set an Outlook rule to drop enquiries here. |
| `OutlookProcessedFolder` | `Inbox\Acorn Intake\Logged` | Logged mail is moved here. Leave blank to leave mail where it is. |
| `OutlookLookbackDays` | `14` | Ignore anything older. Keeps the scrape quick. |
| `OutlookMarkCategory` | `Acorn: Logged` | Category stamped on a message once it is logged. |
| `OutlookSaveAttachments` | `Yes` | Yes = save attachments into EmailInbox\<EmailID>\. |
| `OutlookMaxBodyChars` | `1000` | How much body text to keep in the log, for searching. |
| **Company details (used on generated documents)** | | |
| `CompanyName` | `Acorn Recyclers Ltd` |  |
| `CompanyAddress` |  | Single line - it is written straight onto documents. |
| `CompanyPostcode` |  |  |
| `CompanyPhone` |  |  |
| `CompanyEmail` |  |  |
| `CompanyVATNumber` |  |  |
| `CarrierLicence` |  | Waste carrier registration. Legally required on every transfer note. |
| `EnvironmentalPermit` |  | Permit number for the transfer station. |
| `CompanySICCode` | `38320` | Recovery of sorted materials. Change if yours differs. |
| **Rules & thresholds** | | |
| `VATRate` | `0.2` | Applied to new invoice lines. Stored as a fraction. |
| `AlertAmberDays` | `30` | How far ahead a renewal turns amber on the Alerts sheet. |
| `LongHireDays` | `14` | A container out longer than this is flagged for chasing. |
| `RecoveryTargetPct` | `0.9` | Green/amber/red banding on the recovery rate. |
| `MassBalanceTolerancePct` | `0.05` | In-vs-out variance above this is flagged. Some variance is normal (moisture, stock on site); a large one usually means missing Out tickets. |
| `AutoCreateJobFromEmail` | `No` | Yes lets the intake routine raise a Booked job straight from a matched email. Leave No until you trust the classification rules. |
| `ExportPDF` | `Yes` | Also save a PDF alongside every generated Word document. |
| `CurrentUser` |  | Initials stamped on records you create. Set this on each PC. |
