# Dirty Ledger — Financial Fraud Case (HC-GH-2025-FF-0047)
<img width="1254" height="1254" alt="ledger old" src="https://github.com/user-attachments/assets/0a1ee95a-32f0-4460-b83d-672f5ad5587e" />

Digital forensic investigation case files for Operation "Dirty Ledger" — a Ghana Financial Intelligence Centre (FIC) referred money laundering investigation involving suspect Paul Addo and the front company Delta Ventures Ltd.

Case origin: `DIGITAL-FORENSICS-CTF-LAB / Dirty Ledger - Financial Fraud Case`

## Repository structure

```
dirty-ledger-case/
├── README.md
├── report/
│   ├── build.js                          # docx generation script (docx-js)
│   └── Dirty_Ledger_Forensic_Report.docx # generated final report (23-task write-up)
├── evidence/
│   ├── autopsy_screenshots/   # raw Autopsy screenshots/exports (Sections A–F source material)
│   ├── whatsapp/              # msgstore.db / messages table exports (Section B, Tasks 4–8)
│   ├── financial_documents/   # ledger_offshore_Q1.xlsx, invoices (Section C, Tasks 9–12)
│   ├── exports_2025_files/    # Dropbox /EXPORTS_2025/ + /OLD/ recovered files (Task 15)
│   ├── email_cloud/           # .eml files — wire confirmations, invoice emails (Tasks 11,13,14)
│   ├── protonmail/            # ProtonMail export (Task 16)
│   └── browser_artifacts/     # Chrome History, Login Data, keyword_search_terms (Tasks 17–19)
└── diagrams/                  # network/timeline diagrams (Task 23, supporting Task 20)
```

## Where to upload your files

Place extracted/exported files from Autopsy and other sources into the matching folder below. Suggested filenames are shown — exact names aren't critical, but keeping them grouped by task helps cross-referencing in the final report.

| Folder | What goes here | Related Tasks |
|---|---|---|
| `evidence/autopsy_screenshots/` | All raw Autopsy screenshots (Data Artifacts, File Metadata, Analysis Results views) — e.g. `01_image_integrity.png`, `02_filesystem.png`, `03_file_count.png` | 1, 2, 3 |
| `evidence/whatsapp/` | `msgstore.db`, exported `messages` table CSV (`whatsapp_messages.csv`) | 4–8 |
| `evidence/financial_documents/` | `ledger_offshore_Q1.xlsx`, `invoice_Jan2025.pdf`, `invoice_Feb2025.pdf`, `falsified_invoices_Q1.zip` | 9–12 |
| `evidence/exports_2025_files/` | Everything from Dropbox `/EXPORTS_2025/` and `/OLD/` — `Boss_directives_Q1.txt`, `FIC_contact_details.txt`, `Delta_Ventures_Registration.pdf`, `Accra_Micro_Batch_Jan.csv`, `account_routing_offshore.txt`, `collection_schedule_Feb.xlsx`, `Q2_planning_draft.docx`, `invoice_template_blank.docx`, `dropbox_file_journal.csv`, and the deleted `/OLD/` files (`ledger_2024_Q3.xlsx`, `contacts_dirty.txt`, `offshore_routing_2024.txt`) | 11, 12, 14, 15, 21, 22 |
| `evidence/email_cloud/` | `.eml` files — `Wire_Transfer_UAE_Confirmation.eml`, `Wire_Transfer_Cyprus_Confirmation.eml`, `RE_EXPORTS_2025_Dropbox.eml`, `RE_Invoice_GH85000_Confirmation.eml`, `Q2_Expansion_Brief.eml` | 11, 13, 14 |
| `evidence/protonmail/` | ProtonMail export (`protonmail_export.csv` or `.eml` set — PM-2025-0001 to 0005) | 16 |
| `evidence/browser_artifacts/` | `keyword_search_terms.csv`, `History` (SQLite), `Login Data` (SQLite) or exported `logins.xlsx` | 17, 18, 19 |
| `diagrams/` | Network diagram (Task 23) as image — e.g. exported from draw.io/Lucidchart — and any timeline visualisation (Task 20) | 20, 23 |

## Report

The consolidated 23-task forensic report is generated via `report/build.js` (Node.js, `docx` package). To regenerate after adding evidence:

```bash
cd report
npm install docx
node build.js
```

This produces `Dirty_Ledger_Forensic_Report.docx`.

## Chain of custody note

All evidence files placed in `evidence/` should retain their original metadata (timestamps, hashes) where possible. Do not modify file contents after upload — if corrections are needed, add a new versioned file and note the discrepancy in the report rather than overwriting.

## Status

Draft — for internal review. See "Priority Follow-Up Actions" in the report for outstanding investigative steps (TRK-2025-0118 cross-reference, Wise/TransferWise subpoena, ecowasbank-alerts.net domain verification, Cyprus/UAE FIU referrals, 2024 prior-cycle analysis).
