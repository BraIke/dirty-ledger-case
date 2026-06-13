const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, TableOfContents,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, PageBreak
} = require('docx');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}
function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
}
function p(text, opts = {}) {
  return new Paragraph({ children: [new TextRun({ text, ...opts })], spacing: { after: 120 } });
}
function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun(text)],
    spacing: { after: 60 }
  });
}
function numbered(text) {
  return new Paragraph({
    numbering: { reference: "numbers", level: 0 },
    children: [new TextRun(text)],
    spacing: { after: 60 }
  });
}

function makeTable(header, rows, widths) {
  const headerRow = new TableRow({
    tableHeader: true,
    children: header.map((text, i) => new TableCell({
      borders,
      width: { size: widths[i], type: WidthType.DXA },
      shading: { fill: "2E5C8A", type: ShadingType.CLEAR },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 20 })] })]
    }))
  });
  const dataRows = rows.map((row, idx) => new TableRow({
    children: row.map((text, i) => new TableCell({
      borders,
      width: { size: widths[i], type: WidthType.DXA },
      shading: { fill: idx % 2 === 0 ? "FFFFFF" : "EEF3F8", type: ShadingType.CLEAR },
      margins: { top: 60, bottom: 60, left: 120, right: 120 },
      children: [new Paragraph({ children: [new TextRun({ text: String(text), size: 20 })] })]
    }))
  }));
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: widths,
    rows: [headerRow, ...dataRows]
  });
}

const children = [];

// ===================== TITLE PAGE =====================
children.push(
  new Paragraph({ spacing: { after: 400 }, children: [] }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "HIVE CONSULT", bold: true, size: 28, color: "2E5C8A" })],
    spacing: { after: 60 }
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Digital Forensics Laboratory, Accra", size: 22, color: "555555" })],
    spacing: { after: 600 }
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "DIGITAL FORENSIC INVESTIGATION REPORT", bold: true, size: 36 })],
    spacing: { after: 200 }
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Operation \u201CDirty Ledger\u201D \u2014 Financial Fraud & Money Laundering Investigation", size: 26 })],
    spacing: { after: 600 }
  }),
  makeTable(
    ["Field", "Detail"],
    [
      ["Case Reference", "DIGITAL-FORENSICS-CTF-LAB / Dirty Ledger - Financial Fraud Case"],
      ["Primary Suspect", "Paul Addo"],
      ["Front Company (Search Site)", "Delta Ventures Ltd"],
      ["Date of Seizure", "15 March 2025"],
      ["Investigating Agency", "Ghana Financial Intelligence Centre (FIC)"],
      ["Forensic Laboratory", "HIVE CONSULT Digital Forensics Laboratory, Accra"],
      ["Evidence Item", "Laptop computer (E01 forensic image), Autopsy analysis"],
      ["Report Status", "Draft \u2014 for internal review"],
    ],
    [3120, 6240]
  ),
  new Paragraph({ children: [new PageBreak()] }),
);

// ===================== TOC =====================
children.push(
  h1("Table of Contents"),
  new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-3" }),
  new Paragraph({ children: [new PageBreak()] }),
);

// ===================== EXECUTIVE SUMMARY =====================
children.push(
  h1("Executive Summary"),
  p("This report documents the findings of a digital forensic examination conducted on a laptop computer seized on 15 March 2025 from the premises of Delta Ventures Ltd, pursuant to a search and seizure order issued in connection with a coordinated money laundering investigation by the Ghana Financial Intelligence Centre (FIC). The laptop, belonging to the primary suspect Paul Addo, was forensically acquired as an E01 disk image and examined using Autopsy at HIVE CONSULT Digital Forensics Laboratory, Accra."),
  p("The examination recovered extensive evidence of a structured money laundering operation spanning at least two financial cycles (2024 and Q1 2025), involving multiple front companies, a network of named associates with defined operational roles, falsified invoicing, structured cash deposits designed to evade Financial Intelligence Centre reporting thresholds, international wire transfers to Dubai (UAE) and Limassol (Cyprus), and direct corruption of a Senior Compliance Officer at the FIC to suppress a regulatory query."),
  p("Evidence was recovered across WhatsApp communications, financial spreadsheets, falsified invoices, email correspondence, ProtonMail encrypted communications, Dropbox cloud-sync artifacts, browser search history, and saved browser credentials. The evidence is highly cross-corroborated: key figures (notably the Q1 2025 target of GH\u00a2500,000 and actual processed amount of GH\u00a2487,000, and the offshore total of USD 105,000) appear consistently across at least three independent artifact categories."),
  p("Key findings include:"),
  bullet("A hierarchical criminal network with a concealed superior (\u201CBoss\u201D) directing operations via Proton Mail, and at least five named subordinates with defined roles (financial processing, offshore transfers, cash collection, and intermediary bribery delivery)."),
  bullet("Three front companies \u2014 Delta Ventures Ltd, Gyimah Construction & Supplies, and West Coast Logistics Ltd \u2014 used to issue 14 falsified invoices totalling GH\u00a2677,000 in Q1 2025 alone."),
  bullet("Structured (\u201Csmurfed\u201D) cash deposits kept below GH\u00a29,500 per transaction, a figure traced directly to a browser search for the FIC reporting threshold."),
  bullet("Two completed international wire transfers totalling USD 105,000 to Dubai and Cyprus, disguised as \u201Cconsultancy\u201D and \u201Csoftware licensing\u201D fees."),
  bullet("Direct, documented bribery of a named FIC Senior Compliance Officer (GH\u00a25,000 per quarter) resulting in the withdrawal of regulatory query TRK-2025-0118."),
  bullet("Evidence of attempted (and largely unsuccessful) anti-forensic activity, including instructions to delete WhatsApp messages, a browser search for evidence-wiping software, and an order to destroy a physical receipt linking cash collections to Delta Ventures Ltd."),
  bullet("Evidence of a prior-year (2024) iteration of the same scheme, indicating an ongoing, recurring operation rather than an isolated incident."),
  p("The findings below are organised according to the 23 investigation tasks set out in the case briefing, across six sections: Image Integrity & Filesystem (A), WhatsApp Evidence (B), Financial Documents (C), Email & Cloud Evidence (D), Browser & Anti-Forensics (E), and Timeline & Analysis (F)."),
  new Paragraph({ children: [new PageBreak()] }),
);

// ===================== SECTION A =====================
children.push(h1("Section A: Image Integrity & Filesystem"));

children.push(h2("Task 1: Verify Image Integrity"));
children.push(p("The E01 forensic image was acquired and verified prior to analysis in accordance with standard forensic procedure. Hash verification (MD5/SHA-1) should be performed on the acquired image and compared against the value recorded at the time of acquisition to confirm that no alteration occurred between seizure and examination."));
children.push(p("Note: hash values were not directly captured in the screenshots reviewed during this examination. It is recommended that the original acquisition hash log produced by the imaging tool (e.g., FTK Imager, Guymager) be appended to this report as Appendix A to formally document chain-of-custody integrity verification."));

children.push(h2("Task 2: Filesystem Identification"));
children.push(p("The disk image was processed in Autopsy and mounted successfully, with Autopsy's standard module set (Data Artifacts, Analysis Results, File Metadata, OS Account) populating correctly \u2014 consistent with an NTFS filesystem, the default for a Windows-based laptop. User profile artifacts were located under the standard Windows path structure C:\\Users\\paddo\\..., confirming a Windows operating system with a local user account associated with the username \u201Cpaddo\u201D (Paul Addo)."));

children.push(h2("Task 3: File Count"));
children.push(p("A full file count was not directly captured during this examination; this figure is typically obtained from the Autopsy case summary / ingest module results (File Metadata view, total file count). It is recommended that the total allocated and unallocated file counts be extracted from the Autopsy \u201CFile Views\u201D summary and recorded here, including a breakdown of deleted/unallocated files identified (cross-reference Task 22, which identified at least 3 deleted Dropbox-synced files in the /OLD/ directory)."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===================== SECTION B =====================
children.push(h1("Section B: WhatsApp Evidence"));

children.push(h2("Task 4: Locate WhatsApp Database"));
children.push(p("The WhatsApp message database (msgstore.db, SQLite format) was located on the suspect's device partition at the standard WhatsApp installation path. The \u201Cmessages\u201D table was extracted via Autopsy's Data Artifacts module, yielding 55 entries (identifiers MSG001\u2013MSG055) spanning the period 24 December 2024 to 10 March 2025 \u2014 a continuous communications record immediately preceding the 15 March 2025 seizure."));

children.push(h2("Task 5: Suspect Communication Network"));
children.push(p("Analysis of the key_remote_jid field identified four distinct communication channels involving the suspect:"));
children.push(makeTable(
  ["Contact / Group JID", "Identity (Inferred)", "Role"],
  [
    ["233283913821@s.whatsapp.net", "Nancy", "Microfinance-side payment processing"],
    ["233287462810@s.whatsapp.net", "Mary", "Offshore transfer handler (Dubai / Cyprus)"],
    ["233284281931@s.whatsapp.net", "Philip", "Cash courier / collector"],
    ["233283913821-1620000000@g.us", "\u201CTeam\u201D group chat", "Operational command channel, includes \u201CBoss\u201D"],
  ],
  [3120, 2640, 3600]
));
children.push(p("The presence of a group chat with the @g.us suffix confirms a hierarchical structure above Paul Addo, with a concealed superior (\u201CBoss\u201D) receiving periodic financial performance summaries."));

children.push(h2("Task 6: Role Assignments"));
children.push(p("Based on message content across the recovered communications, the following roles within the network were established:"));
children.push(makeTable(
  ["Individual", "Role"],
  [
    ["Paul Addo (suspect)", "Central coordinator; primary point of contact for all subordinates and the \u201CBoss\u201D"],
    ["Nancy", "Processes payments through microfinance institutions (e.g., Accra Microfinance) using structured deposit references"],
    ["Mary", "Manages offshore fund transfers (Dubai / Cyprus); also coordinates FIC bribery logistics with Paul"],
    ["Philip", "Physical cash collection/courier across multiple regions; sources falsified invoices; delivers FIC bribe payments"],
    ["Kofi", "Subordinate courier; collected cash in Kumasi"],
    ["\u201CKweku\u201D", "Intermediary who passes cash bribes from Philip to the FIC contact"],
    ["\u201CBoss\u201D", "Concealed superior; sets quarterly targets and strategic direction via the group chat and Proton Mail"],
  ],
  [2400, 6960]
));

children.push(h2("Task 7: Cash Collection Orders"));
children.push(p("The following explicit cash collection instructions were identified in Philip's communication thread:"));
children.push(makeTable(
  ["Approx. Date", "Location", "Amount", "Instruction"],
  [
    ["15 Jan 2025", "Tarkwa branch", "GH\u00a275,000", "Collect cash"],
    ["15 Jan 2025", "Accra Central", "\u2014", "Drop-off via address provided by Mary; instructed not to use own name"],
    ["10 Feb 2025", "Kumasi (Stanbic)", "GH\u00a255,000", "Collected by Kofi"],
    ["7 Mar 2025", "Cape Coast", "GH\u00a240,000", "Scheduled for Wednesday; instructed to use coast road to avoid checkpoints"],
  ],
  [1700, 1900, 1500, 4260]
));
children.push(p("This pattern establishes a recurring, multi-regional cash courier operation consistent with deliberate structuring of cash deposits to remain below regulatory reporting thresholds (see Task 21, Stage 1 \u2014 Placement)."));

children.push(h2("Task 8: Evidence Destruction Intent"));
children.push(p("Two explicit instances of evidence destruction instructions were identified within the WhatsApp dataset:"));
children.push(numbered("8 January 2025 (MSG005): \u201CPerfect. Delete these messages after reading.\u201D \u2014 issued immediately following discussion of split-deposit instructions under reference REF-AML-2025-0041."));
children.push(numbered("15 January 2025 (MSG031): \u201CGood. Burn the receipt.\u201D \u2014 issued in direct response to Philip's report (MSG030) that a cash drop-off receipt had been issued under the name \u201CDelta Ventures Ltd.\u201D"));
children.push(p("These two messages constitute strong evidence of consciousness of guilt and active obstruction. MSG031 in particular directly links Delta Ventures Ltd \u2014 the company at the seized premises \u2014 to a deliberately destroyed item of physical evidence, corroborating the legitimacy and relevance of the search and seizure warrant. The instruction in MSG005 was not in fact carried out in full: all 55 messages remained recoverable in the msgstore.db (see Task 22)."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===================== SECTION C =====================
children.push(h1("Section C: Financial Documents"));

children.push(h2("Task 9: Master Transaction Ledger"));
children.push(p("The file ledger_offshore_Q1.xlsx was recovered from the Dropbox-synced folder C:\\Users\\paddo\\Dropbox\\EXPORTS_2025\\. The spreadsheet is internally referenced under case identifier HC-GH-2025-FF-0047 and records three international wire transfers for Q1 2025, totalling USD 105,000 (cleared transactions) against a GHS equivalent figure of 1,365,000."));
children.push(p("This ledger directly corroborates figures referenced independently in WhatsApp communications (MSG016, MSG024, MSG026) and in the Q1 Summary email sent via ProtonMail (PM-2025-0001), providing strong cross-source verification."));

children.push(h2("Task 10: Offshore Routing"));
children.push(makeTable(
  ["Date", "TXN Reference", "Beneficiary", "Bank / Country", "Amount (USD)", "Routing Description", "Status"],
  [
    ["08 Jan 2025", "TXN-INTL-2025-00882", "Al Maktoum Investment Holdings LLC", "Emirates NBD, UAE", "45,000", "\u201CConsultancy Services Q4 2024\u201D", "CLEARED"],
    ["18 Jan 2025", "TXN-INTL-2025-00943", "Limassol Trade & Advisory Ltd", "Hellenic Bank, Cyprus", "60,000", "\u201CSoftware License Fees Q4 2024\u201D", "CLEARED"],
    ["14 Feb 2025", "TXN-INTL-2025-(pending)", "Limassol Trade & Advisory Ltd", "Hellenic Bank, Cyprus", "(not recorded)", "\u201CQ1 2025 Adjustment\u201D", "IN_PROGRESS"],
  ],
  [1100, 1700, 2100, 1700, 1100, 1660, 1100]
));
children.push(p("The two completed transfers (USD 45,000 to Dubai and USD 60,000 to Cyprus) total USD 105,000, matching the Q1 offshore figure reported to \u201CBoss\u201D (PM-2025-0001: \u201Coffshore USD 105k. Dubai and Cyprus confirmed.\u201D). A third transaction to the same Cyprus beneficiary was recorded as IN_PROGRESS at the time of seizure, with no amount populated in the ledger \u2014 however, the GHS_Equivalent total column (1,365,000) is consistent with only the two cleared transactions, suggesting the third transaction's value field may be incomplete or pending entry. This third transaction should be flagged for urgent follow-up with Cypriot authorities, as it may represent funds still in transit or recoverable."));
children.push(p("Full wire transfer confirmations (Task 13) provide the supporting banking details: Dubai account AE070331234567890123456 (Al Maktoum Investment Holdings LLC) and Cyprus IBAN CY17002001280000001200527600 (Limassol Trade & Advisory Ltd)."));

children.push(h2("Task 11: Falsified Invoices"));
children.push(p("An email dated 20 February 2025 from mary.gyasi.finance@gmail.com to paul.addo.accragh@gmail.com (subject: \u201CRE: EXPORTS_2025 \u2014 Docs Ready\u201D) confirms the preparation of a complete set of falsified invoices for Q1 2025:"));
children.push(makeTable(
  ["Front Company", "Invoice Count", "Total Value (GH\u00a2)"],
  [
    ["Delta Ventures Ltd", "6", "312,000"],
    ["Gyimah Construction & Supplies", "5", "218,000"],
    ["West Coast Logistics Ltd", "3", "147,000"],
    ["TOTAL", "14", "677,000"],
  ],
  [4680, 2340, 2340]
));
children.push(p("The complete set was archived as Dropbox/EXPORTS_2025/falsified_invoices_Q1.zip, password-protected, with the email noting the password was \u201Csame as before\u201D \u2014 indicating reuse of security practices from a prior laundering cycle. Individual invoice files invoice_Jan2025.pdf and invoice_Feb2025.pdf were also recovered from the same folder; the February invoice (GH\u00a2200,000) is directly referenced in WhatsApp message MSG010 as routing \u201Cthrough the construction company front\u201D, almost certainly Gyimah Construction & Supplies."));
children.push(p("This directly fulfils the request made in MSG019 (\u201CCan you get Gyimah to prepare falsified invoices?\u201D), establishing a clear evidential chain from instruction to delivery."));

children.push(h2("Task 12: Money Laundering Strategy"));
children.push(p("The overarching laundering strategy is documented most explicitly in Boss_directives_Q1.txt (recovered from Dropbox/EXPORTS_2025/), received via ProtonMail on 3 January 2025 (message PM-2025-0005). The directive sets out seven operational rules:"));
children.push(numbered("Target: GH\u00a2500,000 \u201Cclean\u201D by 31 March 2025."));
children.push(numbered("No single transaction to exceed GH\u00a29,500 (structuring below the GH\u00a210,000 FIC reporting threshold)."));
children.push(numbered("Use existing front companies only: Delta Ventures, Gyimah, West Coast."));
children.push(numbered("Offshore routing priority: Dubai first, Cyprus for overflow."));
children.push(numbered("FIC budget: GH\u00a25,000 per quarter, arranged via cash through Philip."));
children.push(numbered("No phone calls regarding amounts \u2014 WhatsApp or Proton Mail only."));
children.push(numbered("Weekly status updates via Proton Mail."));
children.push(p("This directive was marked \u201CDO NOT PRINT. DELETE AFTER MEMORISING\u201D but was recovered intact, representing a significant operational security failure on the part of the network and a key piece of strategic evidence. Every operational element observed elsewhere in this investigation \u2014 the GH\u00a29,500 structuring cap, the three named front companies, the Dubai/Cyprus routing order, and the FIC bribery arrangement \u2014 traces directly back to this single document, which is therefore considered the central strategic exhibit of this case."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===================== SECTION D =====================
children.push(h1("Section D: Email & Cloud Evidence"));

children.push(h2("Task 13: Wire Transfer Confirmation"));
children.push(p("Two wire transfer confirmation emails were recovered, both sent from notifications@ecowasbank-alerts.net to paul.addo.accragh@gmail.com:"));
children.push(makeTable(
  ["Field", "Wire 1 (UAE)", "Wire 2 (Cyprus)"],
  [
    ["Reference", "TXN-INTL-2025-00882", "TXN-INTL-2025-00943"],
    ["Amount", "USD 45,000.00", "USD 60,000.00"],
    ["Beneficiary", "Al Maktoum Investment Holdings LLC", "Limassol Trade & Advisory Ltd"],
    ["Destination", "Dubai, UAE \u2014 Account AE070331234567890123456", "Cyprus \u2014 IBAN CY17002001280000001200527600"],
    ["Description", "Consultancy Services Q4 2024", "Software License Fees Q4 2024"],
    ["Status", "COMPLETED", "COMPLETED"],
  ],
  [2160, 3600, 3600]
));
children.push(p("The \u201CDescription\u201D fields represent the falsified commercial pretexts used to disguise the true purpose of the funds, consistent with the use of \u201Cexisting fronts only\u201D directed in Boss_directives_Q1.txt. Both descriptions reference \u201CQ4 2024\u201D despite the transfers occurring in January 2025, a discrepancy that may indicate reuse of documentation templates from the prior year's iteration of the scheme (see Task 22, /OLD/ folder)."));
children.push(p("The sender domain \u201Cecowasbank-alerts.net\u201D does not correspond to a recognised major Ghanaian or West African banking institution (cf. Ecobank, GCB Bank, Stanbic) and should be independently verified. If this domain is found to be unaffiliated with a regulated financial institution, it may indicate the use of an unregulated or shell payment intermediary, which would itself constitute additional evidence relevant to the investigation."));
children.push(p("The account/IBAN details identified above should be forwarded to the relevant national Financial Intelligence Units (UAE FIU and Cyprus MOKAS) to support asset tracing and potential freeze requests."));

children.push(h2("Task 14: FIC Obstruction"));
children.push(p("The file FIC_contact_details.txt, recovered from Dropbox/EXPORTS_2025/, provides direct documentary evidence of the corruption of a Financial Intelligence Centre official:"));
children.push(bullet("Position: Senior Compliance Officer (name redacted in the recovered file; described as \u201Cknown to Paul only\u201D)"));
children.push(bullet("Cost: GH\u00a25,000 per quarter, cash only"));
children.push(bullet("Payment chain: Philip delivers cash to an intermediary identified only as \u201CKweku\u201D, who passes the payment to the FIC contact"));
children.push(bullet("Documented outcome: \u201CLast payment 2025-01-19 (Q1 FIC query TRK-2025-0118 withdrawn)\u201D"));
children.push(p("This is independently corroborated by ProtonMail message PM-2025-0003 (19 January 2025, paddo.secure@proton.me to mary.offshore@proton.me): \u201CFIC query TRK-2025-0118 will be withdrawn. Cost: GHS 5,000 this quarter. Arrange via Accra\u2026\u201D, and by Mary's reply (PM-2025-0004) confirming the arrangement and querying delivery method. This exchange also corresponds to WhatsApp messages MSG020\u2013MSG021, in which Paul informs the group that \u201Cthe FIC almost asked questions about the Takoradi transactions\u201D and that \u201Cthe report [has been] amended.\u201D"));
children.push(p("Reference TRK-2025-0118 is a specific, traceable internal FIC case/query identifier and should be cross-referenced against the FIC's own internal records as a priority action item. The individual identified only as \u201CKweku\u201D represents an unidentified person of interest requiring further investigation."));

children.push(h2("Task 15: Dropbox File Journal"));
children.push(p("The Dropbox file_journal table (15 entries) was recovered, comprising 12 active files within /EXPORTS_2025/ and 3 deleted files (is_deleted = 1) within a folder named /OLD/. The active folder contents are summarised below:"));
children.push(makeTable(
  ["File", "Significance"],
  [
    ["falsified_invoices_Q1.zip", "Falsified invoice archive (Task 11)"],
    ["ledger_offshore_Q1.xlsx", "Master offshore transaction ledger (Tasks 9\u201310)"],
    ["invoice_Jan2025.pdf / invoice_Feb2025.pdf", "Individual falsified invoices (Task 11)"],
    ["Delta_Ventures_Registration.pdf", "Front company incorporation document \u2014 formal paper trail for Delta Ventures Ltd"],
    ["Accra_Micro_Batch_Jan.csv", "Corresponds to microfinance split-deposit batch (MSG002/004)"],
    ["account_routing_offshore.txt", "Account/routing details for offshore transfers"],
    ["collection_schedule_Feb.xlsx", "Cash collection schedule (Task 7)"],
    ["FIC_contact_details.txt", "FIC bribery arrangement (Task 14)"],
    ["Boss_directives_Q1.txt", "Operational strategy document (Task 12)"],
    ["Q2_planning_draft.docx", "Forward-looking Q2 expansion plans (new Kumasi front, AMF channel expansion)"],
    ["invoice_template_blank.docx", "Blank invoice template \u2014 demonstrates the invoices were internally fabricated rather than genuinely issued"],
  ],
  [3000, 6360]
));
children.push(p("The deleted /OLD/ folder contained ledger_2024_Q3.xlsx, contacts_dirty.txt, and offshore_routing_2024.txt \u2014 evidence of a prior-year iteration of the same scheme (see Task 22 and Task 21 supporting analysis)."));

children.push(h2("Task 16: ProtonMail Evidence"));
children.push(p("Five encrypted ProtonMail messages (is_encrypted = 1) were recovered for the account paddo.secure@proton.me \u2014 a dedicated \u201Cclean\u201D alias used exclusively for communication with the network's concealed superior and with Mary's offshore-dedicated alias."));
children.push(makeTable(
  ["ID", "Date", "From \u2192 To", "Summary"],
  [
    ["PM-2025-0005", "03 Jan 2025", "bigdirector.anon \u2192 paddo.secure", "Boss Directive \u2014 Q1 target GH\u00a2500,000, GH\u00a29,500 cap, fronts, offshore routing, OPSEC rules (= Boss_directives_Q1.txt)"],
    ["PM-2025-0003", "19 Jan 2025", "paddo.secure \u2192 mary.offshore", "FIC query TRK-2025-0118 to be withdrawn for GH\u00a25,000 this quarter"],
    ["PM-2025-0004", "20 Jan 2025", "mary.offshore \u2192 paddo.secure", "Confirms FIC payment arrangement; asks delivery method (Philip cash vs. bank transfer)"],
    ["PM-2025-0001", "28 Feb 2025", "paddo.secure \u2192 bigdirector.anon", "Q1 Summary (Confidential): GH\u00a2487k processed; offshore USD 105k (Dubai + Cyprus); no alerts"],
    ["PM-2025-0002", "01 Mar 2025", "bigdirector.anon \u2192 paddo.secure", "Q2 Expansion: register new front in Kumasi by April; expand AMF channel"],
  ],
  [1500, 1100, 2500, 4260]
));
children.push(p("Key findings from this artifact category:"));
children.push(bullet("A previously unseen alias, mary.offshore@proton.me, is identified \u2014 corroborating Mary's expanded role in FIC bribery logistics, not solely offshore transfers."));
children.push(bullet("The Boss\u2013Paul command channel (bigdirector.anon@proton.me \u2194 paddo.secure@proton.me) is now corroborated across three independent artifact types: the WhatsApp group chat, Boss_directives_Q1.txt, and this ProtonMail export."));
children.push(bullet("The Q1 figures (GH\u00a2487,000 / USD 105,000) appear consistently in WhatsApp (MSG043), the offshore ledger (Task 9), and this email \u2014 a strong evidentiary triangulation."));
children.push(bullet("TRK-2025-0118 is now documented in two independent artifact types (text file and email) with a matching date (19 January 2025), strengthening the FIC corruption finding."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===================== SECTION E =====================
children.push(h1("Section E: Browser & Anti-Forensics"));

children.push(h2("Task 17: Incriminating Searches"));
children.push(p("Eight entries were recovered from the Chrome keyword_search_terms table. Six relate to operational planning and evasion research:"));
children.push(makeTable(
  ["Search Term", "Relevance"],
  [
    ["how to move money offshore ghana", "Pre-planning for offshore routing (Task 10)"],
    ["anti money laundering detection ghana", "Research into AML detection mechanisms"],
    ["FIC reporting threshold Ghana 2024", "Directly explains the GH\u00a29,500 structuring cap in Boss_directives_Q1.txt"],
    ["falsify invoice Ghana company", "Pre-planning for the falsified invoice scheme (Task 11)"],
    ["SWIFT wire transfer no trace ghana", "Pre-planning for untraceable international wires"],
    ["police checkpoint Accra Kumasi highway", "Matches WhatsApp MSG032/033 \u2014 searched immediately before rerouting a Kumasi cash collection"],
  ],
  [4680, 4680]
));

children.push(h2("Task 18: Anti-Forensic Evidence"));
children.push(p("The remaining two search entries demonstrate direct anti-forensic intent:"));
children.push(makeTable(
  ["Search Term", "Significance"],
  [
    ["how to delete whatsapp messages permanently", "Corresponds to MSG005 (\u201CDelete these messages after reading\u201D)"],
    ["CCleaner delete browser history permanently", "Suspect attempted to wipe his own browser history using a third-party tool"],
  ],
  [4680, 4680]
));
children.push(p("These searches, combined with the instruction to \u201Cburn the receipt\u201D (Task 8) and the GH\u00a25,000/quarter payment to suppress the FIC query (Task 14), establish a consistent pattern of consciousness of guilt and active obstruction across digital and physical evidence categories. Notably, the search query relating to CCleaner itself survived within the browser's keyword_search_terms table, despite being a tool intended to erase such records \u2014 the attempted anti-forensic action did not achieve its purpose."));

children.push(h2("Task 19: Saved Credentials"));
children.push(p("Seven credential records were recovered from Chrome's Login Data SQLite database (table: logins). Password values were stored as encrypted BLOBs and were not recoverable without the associated Windows DPAPI/OS keychain key; however, the associated metadata is highly significant."));
children.push(makeTable(
  ["Site", "Username", "Notes"],
  [
    ["mail.google.com", "paul.addo.accragh@gmail.com", "Primary day-to-day operational email"],
    ["dropbox.com/login", "paul.addo.ventures@gmail.com", "Distinct alias used specifically for Dropbox / EXPORTS_2025"],
    ["accramicrofinance.com", "p.addo", "Direct personal/admin login to the microfinance institution used for laundering deposits"],
    ["deltaventuresltd.business", "admin", "Front company website \u2014 general administrative access"],
    ["web.facebook.com", "paul.addo.gh", "Personal/business social media"],
    ["transferwise.com (Wise)", "paul.addo.ventures@gmail.com", "Previously undocumented money transfer channel \u2014 new lead"],
    ["proton.me", "paddo.secure@proton.me", "Confirms the Boss-facing ProtonMail alias (Task 16)"],
    ["deltaventuresltd.business/admin", "paddo_admin", "Separate, personal admin-level credential for the front company backend"],
  ],
  [2800, 2960, 3600]
));
children.push(p("This artifact reveals at least three distinct identity tiers used by the suspect: a public/day-to-day Gmail account (paul.addo.accragh@gmail.com), a laundering-operations-specific Gmail account (paul.addo.ventures@gmail.com, used for Dropbox and Wise), and a Boss-facing ProtonMail alias (paddo.secure@proton.me)."));
children.push(p("Two findings are of particular investigative significance:"));
children.push(bullet("Direct personal credentialed access to accramicrofinance.com indicates either an insider relationship with, or a controlled account at, that institution \u2014 beyond simply being a deposit beneficiary."));
children.push(bullet("A previously undocumented Wise (TransferWise) account registered to paul.addo.ventures@gmail.com represents an additional money transfer channel not referenced in any other recovered artifact, and should be prioritised for a data request/subpoena to Wise."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===================== SECTION F =====================
children.push(h1("Section F: Timeline & Analysis"));

children.push(h2("Task 20: Event Timeline"));
children.push(makeTable(
  ["Date", "Event", "Source"],
  [
    ["24 Sep 2024", "offshore_routing_2024.txt last modified (prior-year operation)", "Dropbox journal (/OLD/)"],
    ["24 Dec 2024", "Nancy: GH\u00a285,000 invoice ready, to be sent to Koforidua account", "WhatsApp MSG001"],
    ["24 Dec 2024", "Paul: split deposit GH\u00a250k + GH\u00a235k under REF-AML-2025-0041", "WhatsApp MSG003\u2013004"],
    ["24 Dec 2024", "Paul: \u201CDelete these messages after reading\u201D", "WhatsApp MSG005"],
    ["03 Jan 2025", "Boss_directives_Q1.txt received via ProtonMail \u2014 Q1 strategy set out", "PM-2025-0005 / Dropbox"],
    ["05\u201307 Jan 2025", "Searches: offshore movement, AML detection, FIC reporting threshold", "Browser history"],
    ["08 Jan 2025", "Wire TXN-INTL-2025-00882: USD 45,000 to Al Maktoum Holdings, Dubai (\u201CConsultancy Services\u201D)", "Email confirmation"],
    ["12 Jan 2025", "Mary confirms Dubai transfer landed, no flags; next tranche USD 60,000 planned", "WhatsApp MSG015\u2013017"],
    ["13 Jan 2025", "Philip asked to source falsified invoices via Gyimah for the large amount", "WhatsApp MSG018\u2013019"],
    ["15 Jan 2025", "Kumasi batch GH\u00a2120,000 split across 4 accounts; roles assigned (Philip / Mary)", "WhatsApp MSG006\u2013008"],
    ["15 Jan 2025", "Philip ordered to collect GH\u00a275,000 from Tarkwa, drop at Accra Central without using his name", "WhatsApp MSG027\u2013029"],
    ["15 Jan 2025", "Receipt issued under Delta Ventures Ltd name; instruction to \u201Cburn the receipt\u201D", "WhatsApp MSG030\u2013031"],
    ["17 Jan 2025", "Search: \u201Cfalsify invoice ghana company\u201D", "Browser history"],
    ["18 Jan 2025", "Wire TXN-INTL-2025-00943: USD 60,000 to Limassol Trade & Advisory Ltd, Cyprus (\u201CSoftware License Fees\u201D)", "Email confirmation"],
    ["19 Jan 2025", "FIC query TRK-2025-0118 ordered withdrawn for GH\u00a25,000 (FIC corruption)", "PM-2025-0003 / FIC_contact_details.txt"],
    ["19\u201320 Jan 2025", "Second offshore tranche confirmed; ledger_offshore_Q1.xlsx uploaded to Dropbox", "WhatsApp MSG022\u2013024"],
    ["20 Jan 2025", "Mary confirms FIC payment arrangement, queries delivery method", "PM-2025-0004"],
    ["31 Jan 2025", "Philip confirms collection of further cash, awaits drop-off location", "WhatsApp MSG028"],
    ["01 Feb 2025", "New GH\u00a2200,000 invoice routed through construction company front (invoice_Feb2025.pdf)", "WhatsApp MSG009\u2013010 / Dropbox"],
    ["03\u201304 Feb 2025", "Police checkpoint on Accra-Kumasi road; instruction to reschedule cash run", "WhatsApp MSG032\u2013033"],
    ["03 Feb 2025", "Search: \u201Cpolice checkpoint accra kumasi highway\u201D", "Browser history"],
    ["10 Feb 2025", "Kumasi Stanbic GH\u00a255,000 collected by Kofi, confirmed \u201Cback room as usual\u201D", "WhatsApp MSG034\u2013036"],
    ["20 Feb 2025", "14 falsified invoices confirmed ready (GH\u00a2677,000), archived to falsified_invoices_Q1.zip", "Email \u2014 RE_EXPORTS_2025_Dropbox.eml"],
    ["23 Feb 2025", "Search: \u201CSWIFT wire transfer no trace ghana\u201D", "Browser history"],
    ["28 Feb 2025", "Q1 Summary to Boss: GH\u00a2487,000 processed, USD 105,000 offshore, no alerts", "PM-2025-0001"],
    ["01 Mar 2025", "Boss directs Q2 expansion: new Kumasi front, expand AMF channel", "PM-2025-0002"],
    ["07 Mar 2025", "Cape Coast cash run, GH\u00a240,000, instructed to take coast road", "WhatsApp MSG037\u2013039"],
    ["10 Mar 2025", "Group chat: Q1 target confirmed (GH\u00a2500,000), actual GH\u00a2487,000; proceed to Q2", "WhatsApp MSG040\u2013046"],
    ["(undated)", "Searches relating to deleting WhatsApp messages and wiping browser history (anti-forensic attempt)", "Browser history (URL IDs 21\u201322)"],
    ["15 Mar 2025", "Search and seizure executed at Delta Ventures Ltd premises; laptop seized", "Case briefing"],
  ],
  [1300, 5800, 2260]
));

children.push(h2("Task 21: Identify the Three Stages"));
children.push(p("The operation conforms closely to the classical three-stage money laundering model (Placement, Layering, Integration), with a fourth, cross-cutting layer of regulatory corruption protecting the entire process."));

children.push(h3("Stage 1 \u2014 Placement"));
children.push(p("Cash collected from multiple regional sources (Tarkwa: GH\u00a275,000; Kumasi: GH\u00a255,000 and GH\u00a2120,000; Cape Coast: GH\u00a240,000) by Philip and Kofi was introduced into the financial system in structured amounts, each kept below the GH\u00a29,500 cap specified in Boss_directives_Q1.txt (itself derived from a browser search of the FIC reporting threshold), via Nancy through Accra Microfinance under reference REF-AML-2025-0041."));

children.push(h3("Stage 2 \u2014 Layering"));
children.push(p("Domestic funds were routed through three front companies \u2014 Delta Ventures Ltd, Gyimah Construction & Supplies, and West Coast Logistics Ltd \u2014 supported by 14 falsified invoices totalling GH\u00a2677,000. These funds were then converted into international wire transfers disguised as legitimate trade transactions: USD 45,000 to Al Maktoum Investment Holdings (Dubai) described as \u201CConsultancy Services\u201D, and USD 60,000 to Limassol Trade & Advisory Ltd (Cyprus) described as \u201CSoftware License Fees.\u201D"));

children.push(h3("Stage 3 \u2014 Integration"));
children.push(p("The offshore accounts in Dubai and Cyprus hold the laundered USD 105,000, reported to \u201CBoss\u201D as fully processed with \u201Cno alerts\u201D (PM-2025-0001). These funds now present as legitimate consultancy and software-licensing revenue in jurisdictions disconnected from the originating cash, available for reinvestment \u2014 consistent with the directed Q2 expansion (new Kumasi front, expanded AMF channel)."));

children.push(h3("Cross-Cutting Layer \u2014 Corruption & Obstruction"));
children.push(p("Running in parallel to all three stages, a quarterly GH\u00a25,000 cash payment to a Senior Compliance Officer at the FIC (delivered via Philip and the intermediary \u201CKweku\u201D) directly resulted in the withdrawal of regulatory query TRK-2025-0118, shielding the operation from detection at a critical point in the placement/layering boundary."));

children.push(h2("Task 22: Deleted Evidence Recovery"));
children.push(p("Several deleted or attempted-to-be-deleted items were recovered:"));
children.push(numbered("/OLD/ledger_2024_Q3.xlsx (deleted, Dropbox journal id 13) \u2014 a prior quarter's laundering ledger, demonstrating that the scheme is recurring rather than isolated to Q1 2025."));
children.push(numbered("/OLD/contacts_dirty.txt (deleted, Dropbox journal id 14) \u2014 an earlier version of the network's associate contact list."));
children.push(numbered("/OLD/offshore_routing_2024.txt (deleted, Dropbox journal id 15) \u2014 2024 offshore routing details, predating the Dubai/Cyprus routes used in Q1 2025, indicating route rotation as an anti-detection technique across operational cycles."));
children.push(numbered("WhatsApp messages 1\u201355 remained fully recoverable in msgstore.db despite the explicit instruction in MSG005 to delete them \u2014 either the instruction was not followed, or deletion (if attempted) removed messages from the user interface without purging the underlying SQLite records."));
children.push(numbered("The browser search \u201CCCleaner delete browser history permanently\u201D (Task 18) survived within the keyword_search_terms table despite apparently being made in an attempt to use such a tool to erase browsing history \u2014 the cleaning tool's own search query was not itself removed."));
children.push(numbered("The physical receipt referenced in MSG030\u2013031 was reportedly burned and cannot be recovered; however, its prior existence \u2014 and its link to Delta Ventures Ltd \u2014 is documented in the WhatsApp exchange and is relevant to chain-of-custody and provenance arguments concerning the seized premises."));

children.push(h2("Task 23: Suspect Network Diagram"));
children.push(p("The diagram below summarises the command structure, named individuals, front companies, and offshore destinations identified during this examination."));

children.push(h3("Command Structure"));
children.push(makeTable(
  ["Tier", "Entity", "Identifiers", "Function"],
  [
    ["1", "\u201CBoss\u201D (concealed superior)", "bigdirector.anon@proton.me", "Sets quarterly targets and strategic direction; receives weekly/periodic reports"],
    ["2", "Paul Addo (suspect)", "233283913821 (WhatsApp) / paul.addo.accragh@gmail.com / paul.addo.ventures@gmail.com / paddo.secure@proton.me", "Central coordinator; communicates with Boss and all subordinates under multiple aliases"],
    ["3", "Nancy", "233283913821 (WhatsApp)", "Microfinance payment processing (Accra Microfinance)"],
    ["3", "Mary", "233287462810 (WhatsApp) / mary.offshore@proton.me / mary.gyasi.finance@gmail.com", "Offshore transfers (Dubai/Cyprus); FIC bribery logistics"],
    ["3", "Philip", "233284281931 (WhatsApp)", "Cash courier; sources falsified invoices; delivers FIC bribe to \u201CKweku\u201D"],
    ["4", "Kofi", "(not separately identified)", "Sub-courier; Kumasi cash collection"],
    ["4", "\u201CKweku\u201D", "(unidentified \u2014 person of interest)", "Intermediary passing cash bribes to FIC contact"],
    ["5", "FIC Senior Compliance Officer", "(redacted in source document)", "Withdrew regulatory query TRK-2025-0118 in exchange for payment"],
  ],
  [800, 2400, 3360, 2800]
));

children.push(h3("Front Companies and Offshore Destinations"));
children.push(makeTable(
  ["Front Company (Layering)", "Invoices (Q1 2025)", "Linked Offshore Destination (Integration)", "Amount"],
  [
    ["Delta Ventures Ltd (seizure site)", "6 invoices, GH\u00a2312,000", "Al Maktoum Investment Holdings LLC, Dubai, UAE", "USD 45,000"],
    ["Gyimah Construction & Supplies", "5 invoices, GH\u00a2218,000", "Limassol Trade & Advisory Ltd, Cyprus", "USD 60,000"],
    ["West Coast Logistics Ltd", "3 invoices, GH\u00a2147,000", "(unconfirmed \u2014 see Task 10, third transaction)", "(pending)"],
  ],
  [2700, 2200, 3060, 1400]
));
children.push(p("Note: the network diagram above is presented in tabular form for compatibility with this document format. It is recommended that a graphical version (e.g., flowchart) be produced as a supporting exhibit for presentation purposes, illustrating the flow of funds from cash placement (Philip/Kofi) through the named front companies to the offshore beneficiary accounts, with the command and corruption channels (Boss, FIC contact, \u201CKweku\u201D) shown as parallel overlays."));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===================== RECOMMENDATIONS =====================
children.push(h1("Priority Follow-Up Actions"));
children.push(numbered("Cross-reference FIC query reference TRK-2025-0118 against the Financial Intelligence Centre's internal records; identify the Senior Compliance Officer concerned and the individual identified only as \u201CKweku.\u201D"));
children.push(numbered("Issue a data preservation request / subpoena to Wise (TransferWise) for the account registered to paul.addo.ventures@gmail.com, identified as a previously undocumented transfer channel."));
children.push(numbered("Verify the legitimacy and ownership of the domain ecowasbank-alerts.net as a banking notification source."));
children.push(numbered("Forward account/IBAN details (AE070331234567890123456, Dubai; CY17002001280000001200527600, Cyprus) to the UAE Financial Intelligence Unit and Cyprus MOKAS for asset tracing and potential freeze action, including the unresolved third (IN_PROGRESS) transaction to Limassol Trade & Advisory Ltd."));
children.push(numbered("Obtain and analyse the recovered 2024 ledger, contacts list, and offshore routing files from the /OLD/ Dropbox folder to establish the scope of a potential broader charge period predating Q1 2025."));
children.push(numbered("Confirm hash verification of the original E01 image against acquisition-time values and append the acquisition log to this report (Task 1)."));
children.push(numbered("Obtain the full Autopsy file count summary, including allocated/unallocated breakdown, for inclusion in Task 3."));

children.push(new Paragraph({ spacing: { before: 400 }, children: [] }));
children.push(p("End of Report.", { bold: true }));

// ===================== DOCUMENT ASSEMBLY =====================
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "\u2022", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "Arial", color: "1F3864" },
        paragraph: { spacing: { before: 320, after: 200 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "1F3864", space: 4 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2E5C8A" },
        paragraph: { spacing: { before: 240, after: 160 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "44546A" },
        paragraph: { spacing: { before: 160, after: 120 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children: [new TextRun({ text: "HIVE CONSULT DFL \u2014 Case HC-GH-2025-FF-0047 \u2014 \u201CDirty Ledger\u201D Investigation", size: 16, color: "888888" })]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({ text: "Page ", size: 18 }),
            new TextRun({ children: [PageNumber.CURRENT], size: 18 }),
            new TextRun({ text: " of ", size: 18 }),
            new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 18 }),
          ]
        })]
      })
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => fs.writeFileSync("/home/claude/report/Dirty_Ledger_Forensic_Report.docx", buffer));
