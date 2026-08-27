"""
Wizard 01 -- PDF classify / extract / name logic.

Pure logic module: no FastAPI, no job/storage/filesystem concerns (mirrors
how talent_extractor.py is separated from main.py's endpoints). The job
layer (wizard01_jobs.py) is responsible for everything about where bytes
live on disk; this module only ever deals with bytes in memory and returns
a DocResult describing what to do with them.

Naming conventions, prompt wording, and all string/PDF helpers below are
ported near-verbatim from the standalone Tkinter tool this replaces
(`Rename Residency and Diversity/rename_tool.py`), which already handles
these exact document types correctly -- the only genuinely new pieces here
are:
  - the type-classifier step (that tool always knew the type already,
    since a human pre-sorted into per-type folders; here we don't)
  - BatchContext (Client/Agency/ProdCo names+addresses, received_from) and
    the sender-mismatch check on Vendor documents
  - closed reason codes instead of free-text-only reasons, for a UI that
    needs to group/count/label these
"""

import os
import io
import re
import json
import base64
from dataclasses import dataclass, field

import fitz  # PyMuPDF
from PIL import Image, ImageOps

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_BUSINESS_WORDS = {
    "llc", "inc", "incorporated", "corp", "corporation", "dba", "productions",
    "production", "company", "co", "ltd", "llp", "enterprises", "group",
    "studios", "studio", "media", "films", "film", "pictures", "entertainment",
}
_UPPERCASE_EXCEPTIONS = {"llc", "inc", "caps", "nda"}
_NAME_SUFFIXES = {"jr", "jr.", "sr", "sr.", "ii", "iii", "iv", "v", "2nd", "3rd", "4th"}


# ── Reason codes ──────────────────────────────────────────────────────────────
# Closed set so a UI can group/count/label distinctly. reason_detail carries
# the free-text specifics alongside every code below (may be "" for codes
# that are self-explanatory, e.g. REASON_OK).

REASON_OK               = ""                  # success, no reason needed
REASON_UNCLASSIFIED     = "unclassified"      # readable, but not a known type
REASON_UNREADABLE       = "unreadable"        # corrupt/blank, both providers failed
REASON_MISSING_NAME     = "missing_name"
REASON_MISSING_DATE     = "missing_date"
REASON_MISSING_STATE    = "missing_state"
REASON_MISSING_COMPANY  = "missing_company"
REASON_SENDER_MISMATCH  = "sender_mismatch"
REASON_LOW_CONFIDENCE   = "low_confidence"
REASON_PROVIDER_ERROR   = "provider_error"

BUCKET_RENAMED          = "renamed"
BUCKET_UNABLE_TO_RENAME = "unable_to_rename"
BUCKET_NOT_READABLE     = "not_readable"

FOLDER_RESIDENCY  = "Residency"
FOLDER_DIVERSITY  = "Diversity Form"
FOLDER_VENDOR     = "Vendor Invoices"
FOLDER_FREELANCE  = "Freelance Crew Invoice"
FOLDER_PETTY_CASH = "Petty Cash"
FOLDER_PRODCC    = "ProdCC"
FOLDER_CREW_PAYROLL = "Crew Payroll"
FOLDER_OTHER_FORMS = "Other Forms"
FOLDER_NOT_READABLE = "Not Readable"


@dataclass
class BatchContext:
    """One per job, filled once from the Overview-style intake form and
    threaded into every per-file call. received_from is exactly one of
    "prodco" / "agency" -- a batch is never mixed-source, and Client is not
    a valid sender (a batch is never "received from" the Client -- it can
    only be received from the ProdCo or the Agency). client_name is still
    collected as a third named entity, since a vendor invoice's Bill-To can
    legitimately show the Client's name even when the batch itself came
    from the ProdCo or Agency -- it's used for sender-mismatch comparison,
    just never as the "expected" party. No address fields -- they don't
    affect a filename, so they were never actually read anywhere below.
    vendor_naming is one of "invoice_number" (default) / "po_number" -- a
    user-chosen preference for how Vendor Invoice files get named; falls
    back to "invoice_number" automatically per-file whenever a PO number
    isn't actually printed on that invoice, regardless of this setting."""
    client_name: str = ""
    agency_name: str = ""
    prodco_name: str = ""
    received_from: str = ""
    vendor_naming: str = "invoice_number"


@dataclass
class DocResult:
    bucket: str            # BUCKET_* constant
    doc_type: str          # "residency" | "diversity" | "vendor" | "unknown"
    subfolder: str         # FOLDER_* constant -- where this lands in the zip
    output_pdf_bytes: bytes  # possibly image->PDF-converted bytes to write
    new_filename: str = "" # set only when bucket == BUCKET_RENAMED
    reason_code: str = REASON_OK
    reason_detail: str = ""
    provider: str = ""     # "openai" | "anthropic" | ""


# ── Prompts ────────────────────────────────────────────────────────────────────

TYPE_CLASSIFIER_PROMPT = """You are looking at a single scanned or photographed document submitted as part of a film/TV production crew paperwork batch. Determine which broad family it belongs to:

- "residency" -- a Driver's License, a State ID card, or a USCIS Form I-9 (Employment Eligibility Verification).
- "diversity" -- an Illinois Department of Commerce & Economic Opportunity (DCEO) "Illinois Film Tax Credit Tracking Sheet."
- "vendor" -- a hotel guest folio, a purchase receipt/order confirmation, or a formal invoice (vendor or freelance) billing the production for goods, services, or labor. Never a payroll company's own invoice -- see "crew_payroll" below.
- "petty_cash" -- ANY petty cash document: a "Petty Cash Summary"/"Summary of Petty Cash Expenses" reimbursement/reconciliation form for one named custodian (a running list of small cash purchases, an "Employee Name" field, an envelope/receipt log) -- if it's this kind of form at all, it's petty_cash, full stop, regardless of which way its balance-due settlement happens to run; an internal cash-advance Purchase Order that states its purpose is Petty Cash -- however that's expressed on this particular PO template: a Vendor field literally reading "Petty Cash"/"Petty Cash Advance", a "FOR" or "Purpose" field saying "PETTY CASH", or a line-item description literally reading "Petty Cash" (e.g. "PC / Petty Cash / $3,000.00") -- handing a lump sum float to a named custodian; a multi-person petty cash log/spreadsheet summarizing several custodians' totals at once (e.g. a "Petty Cash Spreadsheet"/"PC Master Summary" export); OR just a bare stack of purchase receipts with NO cover sheet, PO, or summary form of any kind in front of them identifying whose batch this is -- an uncovered stack of receipts defaults to petty_cash too. Never "vendor" for any of these, even the PO-cover-sheet variant -- the word "Petty Cash" appearing anywhere on the PO as its stated reason is the deciding signal, regardless of which field it's in or what the rest of that PO template looks like.
- "prodcc" -- a Production Credit Card (ProdCC) reimbursement: the crew member paid with their OWN money/personal card and the production owes THEM back. This is normally identifiable ONLY by an explicit cover sheet clearly built for that purpose -- either an internal Purchase Order-style request where the Vendor field names the crew member directly followed by wording like "- CC Reimbursement"/"- CC Reimb" (never "Petty Cash"), or a human-made cover sheet (often an Excel printout) that itself is specifically about a credit card reimbursement. NEVER classify a "Petty Cash Summary"/"Summary of Petty Cash Expenses" form as prodcc, even if its balance-due line seems to run the "company owes the employee" direction -- that form always means petty_cash regardless.
- "crew_payroll" -- any payroll-company document for the production's crew as a whole (never a single crew member's own invoice): a payroll company's own consolidated invoice for one batch of crew wages (e.g. "New C.A.P.S. LLC," "TakeOne Network Corp. DBA Wrapbook") -- identified by payroll-report page types like "Invoice Fee Summary," "Fringe Recap Report," "Wage/Payroll Register," or "Payroll Check Register," and payroll-specific line items (Corporate/Taxable/Gross Wages, Union Fringes, Employer Taxes, Workers Compensation, Handling Fees); OR a whole-project aggregate export covering every crew member at once for the full project date range, not tied to any one batch/invoice -- a "Fringe Report" (a big table, one row per crew member, of wages/fringes/benefits) or a "Payroll Register" (a running ledger of every payment across the whole project, often hundreds of pages).
- "misc_form" -- any other standardized production paperwork form not covered above, belonging to either one company or one specific crew member and tied to a specific date/period -- e.g. a state tax withholding return (such as Georgia's "G-7 Quarterly Return for Film Payer," usually filed by the payroll company), a crew member's time card, or similar standardized forms.
- "unknown" -- anything else, or if you cannot confidently tell from what's visible.

Return ONLY a JSON object with exactly this key: {"document_family": "residency" | "diversity" | "vendor" | "petty_cash" | "prodcc" | "crew_payroll" | "misc_form" | "unknown"}
No explanation. No markdown. No code fences. JSON object only."""

ORIENTATION_PROMPT = """You are looking at page images from a single PDF document -- a scan or phone photo submitted as part of a film/TV production crew paperwork batch. For each page, determine how many degrees it needs to be rotated CLOCKWISE for its content (text, photos) to appear upright/right-side-up. Common causes: a phone photo taken in landscape when the document is portrait (or vice versa), or a scanner fed the page sideways or upside-down.

Return one entry per page, in the same order given (page 1 first).
- rotation: one of 0, 90, 180, 270 -- how many degrees clockwise this page needs to be rotated to become upright. 0 if it's already upright.

Return ONLY a JSON object with exactly this key: {"pages": [{"rotation": 0 | 90 | 180 | 270}, ...]}
No explanation. No markdown. No code fences. JSON object only."""

RESIDENCY_SYSTEM_PROMPT = """You are analyzing a scanned identification or employment eligibility document submitted by a film production crew member. It will be one of: a Driver's License, a State ID card, or a USCIS Form I-9 (Employment Eligibility Verification). If a second page is included, it is almost always just the BACK of the same card (blood type, restrictions, barcodes) -- treat all pages together as ONE document, never two.

Determine which type this is:
- "drivers_license" -- explicitly says DRIVER'S LICENSE
- "state_id" -- a state-issued ID card that is NOT a driver's license (says ID CARD / IDENTIFICATION CARD, no driving class/privileges)
- "i9" -- USCIS Form I-9, Employment Eligibility Verification
- "unknown" -- if you cannot confidently tell from what's visible

Then extract:
- last_name: the person's last/family name, exactly as printed, title case. For a driver's license/state ID this is field "1" (the first name line, e.g. "ALLEN"). For an I-9 this is "Last Name (Family Name)" in Section 1. If a generational suffix (Jr, Sr, II, III, 3rd, etc.) is printed as part of that same name, keep it attached to last_name (e.g. "Sanders Jr") -- never treat the suffix as first_name or as the whole last_name by itself.
- first_name: the person's FIRST/given name ONLY -- never include a middle name. For a driver's license/state ID, field "2" often has first + middle name together (e.g. "CAROLINE ROSE") -- use only the first word ("Caroline"). For an I-9 use "First Name (Given Name)" in Section 1.
- state: ONLY for drivers_license or state_id, the 2-letter issuing state abbreviation exactly as printed on the card (e.g. "IL", "CA") -- this is usually part of the header ("ILLINOIS DRIVER LICENSE") or printed near the state seal. Empty string for i9 or unknown.
- expiration_date: ONLY for drivers_license or state_id, in exactly MM/DD/YYYY as printed next to "EXP" or field "4b". Empty string for i9 or unknown.

Return ONLY a JSON object with exactly these keys: {"document_type": "...", "last_name": "...", "first_name": "...", "state": "...", "expiration_date": "..."}
No explanation. No markdown. No code fences. JSON object only."""

RESIDENCY_MULTI_SYSTEM_PROMPT = """You are analyzing a multi-page PDF that was batch-scanned by a film production crew coordinator. It contains one or more separate people's identification documents -- each is a Driver's License, a State ID card, or a USCIS Form I-9 (Employment Eligibility Verification). A single person's card is sometimes 2 consecutive pages (front then back -- the back has no photo, just barcodes/restrictions/small print) and sometimes just 1 page (front only, no back scanned). Different people's documents never share a page.

Look at every page, in the order given, and group them into documents -- one group per person. For each document, report which pages belong to it and extract its fields.

For each document:
- pages: a list of the page numbers belonging to this document (1-indexed, in the order given -- e.g. [1, 2] for a front+back pair, or [3] for a front-only page).
- document_type: "drivers_license" / "state_id" / "i9" / "unknown".
- last_name: the person's last/family name, exactly as printed, title case. If a generational suffix (Jr, Sr, II, III, 3rd, etc.) is printed as part of that same name, keep it attached to last_name -- never treat the suffix as first_name or as the whole last_name by itself.
- first_name: the person's FIRST/given name ONLY -- never include a middle name.
- state: ONLY for drivers_license or state_id, the 2-letter issuing state abbreviation exactly as printed on the card (e.g. "IL", "CA"). Empty string for i9 or unknown.
- expiration_date: ONLY for drivers_license or state_id, in exactly MM/DD/YYYY as printed next to "EXP" or field "4b". Empty string for i9 or unknown.

Every page must belong to exactly one document's "pages" list -- never omit a page and never assign one page to two documents.

Return ONLY a JSON object with exactly this key: {"documents": [{"pages": [...], "document_type": "...", "last_name": "...", "first_name": "...", "state": "...", "expiration_date": "..."}, ...]}
No explanation. No markdown. No code fences. JSON object only."""

DIVERSITY_SYSTEM_PROMPT = """You are analyzing an Illinois Department of Commerce & Economic Opportunity (DCEO) "Illinois Film Tax Credit Tracking Sheet" submitted by a film production crew member.

Extract:
- employee_name_field: the exact text written in the "Employee Name" field, verbatim -- even if it looks like a business/loan-out company name rather than an individual person.
- looks_like_person: true if employee_name_field is clearly one individual's name (e.g. "First Last"), false if it looks like a business name (contains words like LLC, Inc, Corp, dba, Productions, Company, etc.), is blank, or is illegible.
- signature_name_guess: your best-effort reading of the person's name from the handwritten cursive signature near the "Signature" line, as First and Last name. Handwriting is often imperfect -- give your best guess. Empty string only if truly illegible/blank.

Return ONLY a JSON object with exactly these keys: {"employee_name_field": "...", "looks_like_person": true or false, "signature_name_guess": "..."}
No explanation. No markdown. No code fences. JSON object only."""

VENDOR_SYSTEM_PROMPT_TEMPLATE = """You are analyzing a vendor document submitted as part of a film/TV production's Accounts Payable packet. It is one of: a hotel guest folio, a purchase receipt/order confirmation, or a formal vendor invoice.
{prodco_line}
Some of these PDFs begin with an internal Purchase Order cover sheet (a page showing "Purchase Order: PO-XXXXXXX", a Subsidiary line, and a Vendor/Accounts/Description/Amount summary) before the vendor's own actual invoice/receipt/folio appears on a later page. When this cover sheet is present, completely ignore its "Purchase Order" number and its summary line items for every field below EXCEPT po_number itself -- always read the ACTUAL vendor document (the page(s) that look like a real invoice, receipt, or folio) for every other field, never the internal PO cover sheet.

You may be looking at a multi-page packet rather than a single document -- e.g. an invoice page, then an internal PO cover sheet, then a Form W-9, then a payment-confirmation/check page, in any order. Read every page before answering. A later page in the packet is often the most reliable place to find the actual billing individual's legal name when the invoice itself only shows a business/loan-out/DBA/studio name: a W-9 always states it on line 1 ("Name of entity/individual" -- the disregarded/business name goes on line 2 instead), and a payment confirmation or check memo page will sometimes spell it out directly as "[Person Name] dba [Company Name]".

- po_number: the internal Purchase Order number/code from that cover sheet, if one is present anywhere in this PDF (e.g. "2528-06" from "Purchase Order: PO-2528-06" -- strip a leading "PO"/"PO-" prefix, keep just the number/code itself). Empty string if no PO cover sheet is present.

FIRST, determine which of these three this document is:

1. HOTEL FOLIO -- a hotel guest bill/folio: shows a hotel name, Room Number, Arrival/Departure dates, and a running list of room/tax charges for one guest's stay. Set is_hotel_folio=true.

2. RECEIPT -- a purchase receipt or order confirmation for a straightforward purchase (a restaurant/catering order, a retail purchase, a delivery confirmation email) that has NO formal invoice structure (no "Bill To" block, no invoice number). Set is_receipt=true (and is_hotel_folio=false). A production's own "Crew Invoice" / crew timesheet form (see is_freelance_crew_labor below) is NEVER a receipt even though it also lacks a Bill To block or invoice number -- set is_receipt=false for those and let is_freelance_crew_labor carry the distinction instead.

3. VENDOR INVOICE -- a formal invoice from a business or individual billing the production, with a "Bill To"/client block and (usually) its own invoice number. Set both is_hotel_folio=false and is_receipt=false.

Then extract the fields relevant to whichever type you picked:

FOR A HOTEL FOLIO ONLY:
- hotel_name: the hotel's name as printed. Required.
- folio_number: the number printed specifically next to a "Folio No." / "Folio #" label. Empty string if that field is blank or not shown -- do NOT substitute a confirmation number, reservation number, or room number instead, even if one is visible nearby.

FOR A RECEIPT OR VENDOR INVOICE, identify the BILLER -- the individual and/or company actually issuing this document and being paid, from its own header/logo/contact-info area. NEVER use the "Bill To"/client/production-company block for these fields, no matter what it's labeled, and never use the internal PO cover sheet's "Subsidiary" or "Vendor" fields if a more specific header appears on the actual document itself.
- has_person_name: true if an individual person's own name is presented as who this document is FROM -- either directly (it's headed with a person's own name, made out personally, or says "make checks payable to [person]"), OR indirectly, when the invoice itself is headed only with a business/loan-out/DBA/studio name (e.g. "Location Media LLC", "Alexa Viscius Studio") but another page in this same packet (a W-9's line 1, a "[Person] dba [Company]" payment-confirmation page, a Fed ID/SS contact line) identifies the individual behind it. false only if no individual can be identified anywhere in the packet.
- last_name / first_name: that person's name, split, ONLY if has_person_name is true. Empty strings otherwise. A generational suffix (Jr, Sr, II, III, etc.) stays attached to last_name.
- has_company_name: true if a business name is shown as the billing entity/merchant.
- company_name: that company's name, exactly as printed, ONLY if has_company_name is true. Empty string otherwise.
- bill_to_name: whoever this document is addressed TO -- the "Bill To" / client / recipient name, exactly as printed, regardless of who that turns out to be. This is the party being billed, never the biller. Empty string if no Bill To block is present (common on a receipt).

FOR A VENDOR INVOICE ONLY (skip these two for a receipt):
- invoice_number: the vendor's OWN invoice number, exactly as printed on the actual invoice itself (e.g. "1080" from "Invoice no.: 1080", "CHIC-26-000208" from "Invoice CHIC-26-000208"). This is never the internal "PO-XXXXXXX" purchase order number from a cover sheet. Empty string if the actual invoice has no invoice number printed anywhere.
- is_freelance_crew_labor: true if this document, in substance, is one individual production professional billing the production for their own personal service/time on this shoot -- not limited to below-the-line "crew" job titles; a Director's fee, a Photographer's creative fee, or a Wardrobe Stylist's day rate all count exactly the same way. Judge this by actually examining the document's content and structure, not by searching for any one specific word. Signals that this IS freelance labor (any one can be enough, and none of them require the literal word "Labor" to appear anywhere): it is titled or functions as a "Crew Invoice" or crew timesheet; it shows a Social Security Number instead of (or alongside) a business Tax ID/EIN; it names a specific person's role/title on this production (e.g. PA, Art Director, Line Producer, Stills, Wardrobe, Grip, HMU/HMUA, Production Coordinator, Set Dresser, Director, Photographer, Stylist, Location Sound Mixer, Storyboard Artist, etc.); it bills a day rate, hourly rate, or flat creative/service fee tied to this specific shoot, with prep/shoot/wrap-style time tracking, call/wrap times, or overtime; it carries a "Sub-contractor Signature" and/or production-manager sign-off line. This still counts as true even when the invoice is issued through the person's own loan-out/DBA/LLC/studio name rather than their personal name -- what matters is that the substance of the charges is one individual's personal service/time on this shoot, e.g. "Art Director, 8 days" through "MSDP Inc.", a "Director Fee" through "Thou Swell Thou Witty, Inc.", or a "Photographer Creative + Usage Fee" through "Alexa Viscius Studio" all still count as true. A single freelancer's invoice commonly bundles in a few incidental line items alongside their main fee -- their OWN kit/equipment/gear rental (e.g. a sound mixer's own recording kit, a photographer's own camera gear), reimbursed travel/meal/parking/supply expenses, film processing, or a helper/assistant they personally brought -- none of that disqualifies it; the whole invoice is still that one person's freelance labor. Set false only for a genuine third-party vendor business -- a dedicated equipment rental house, prop house, venue, or goods supplier -- invoicing the production for goods/services that are not tied to one specific billing individual's personal day-rate/creative-service work on this shoot.

Return ONLY a JSON object with exactly these keys: {"po_number": "...", "is_hotel_folio": true or false, "hotel_name": "...", "folio_number": "...", "is_receipt": true or false, "has_person_name": true or false, "last_name": "...", "first_name": "...", "has_company_name": true or false, "company_name": "...", "bill_to_name": "...", "invoice_number": "...", "is_freelance_crew_labor": true or false}
No explanation. No markdown. No code fences. JSON object only."""

PETTY_CASH_SYSTEM_PROMPT = """You are analyzing a petty cash document submitted as part of a film/TV production's Accounts Payable packet. It will be one of:

1. A PETTY CASH SUMMARY form for ONE named custodian -- a running log of small cash purchases/receipts (Date, To Whom, For What, Amount) with an "Employee Name" field identifying whose envelope this is. Extra scanned receipt images (rideshare receipts, store receipts, email confirmations) are commonly attached on later pages behind this cover sheet -- ignore them, the cover sheet's Employee Name is all you need.

2. An internal cash-advance Purchase Order handing a lump-sum float to a custodian -- identified by "Petty Cash" being stated as the PO's own reason for existing, however that happens to be expressed on this particular PO template: a Vendor field literally reading "Petty Cash"/"Petty Cash Advance", a "FOR"/Purpose field saying "PETTY CASH", or a line-item description literally reading "Petty Cash" (e.g. "PC / Petty Cash / $3,000.00"). The custodian's name is whichever person's name is actually presented as who this float is FOR -- never a company/vendor name -- read whatever field that PO template uses for it: a "TO"/"ATTN" field naming the recipient directly, a Purpose/description block where it typically appears directly below wording like "PC Float" (e.g. "PC Float\\nHolli McGinley"), or a Contact name if that's the only name given. If the same person's name appears with slightly different spelling in different fields on the same PO (a typo is common on these), prefer the spelling in the most prominent header-level field (e.g. "TO") over one in a smaller supporting field (e.g. a payment/ACH line).

3. A PETTY CASH LOG / multi-person summary spreadsheet -- an aggregate report (often a "Petty Cash Spreadsheet" or "PC Master Summary" export from production budgeting software) that totals up SEVERAL different custodians' petty cash at once, typically as columns of initials rather than one full name. This has no single custodian to name.

4. A bare stack of purchase receipts with NO cover sheet, PO, summary form, or handwritten note of any kind in front of them identifying whose batch this is -- just receipt images/scans back to back. This has no name to extract at all, by design.

Determine:
- has_cover_sheet: false ONLY for case 4 -- no summary form, PO cover sheet, or any other page identifying a custodian appears anywhere in this document, just raw receipts. true for cases 1, 2, and 3.
- is_petty_cash_log: true ONLY for case 3 above -- a multi-custodian aggregate/summary report with no one specific named custodian. false otherwise.
- has_person_name: true if a single custodian's name can be identified (cases 1 and 2). Always false when is_petty_cash_log is true or has_cover_sheet is false.
- last_name / first_name: that custodian's name, split, ONLY if has_person_name is true. Empty strings otherwise. A generational suffix (Jr, Sr, II, III, etc.) stays attached to last_name.

Return ONLY a JSON object with exactly these keys: {"has_cover_sheet": true or false, "is_petty_cash_log": true or false, "has_person_name": true or false, "last_name": "...", "first_name": "..."}
No explanation. No markdown. No code fences. JSON object only."""

PRODCC_SYSTEM_PROMPT = """You are analyzing a Production Credit Card (ProdCC) reimbursement document submitted as part of a film/TV production's Accounts Payable packet. ProdCC is how a crew member gets paid back for production expenses they already paid for with their OWN money/personal card -- the reverse of Petty Cash, where the company hands cash to the crew member up front instead. It will be one of:

1. An internal Purchase Order-style reimbursement request -- the Vendor field on these names the crew member directly, followed by wording like "- CC Reimbursement" or "- CC Reimb" (e.g. "Logan Gilmore - CC Reimbursement"), never a real company and never "Petty Cash". Pages behind the PO cover sheet are typically the crew member's own scanned credit-card receipts.

2. A human-made cover sheet (often an Excel printout, not the standard "Petty Cash Summary" form) that a crew member or coordinator put together themselves specifically to itemize a credit card reimbursement request for one named person.

Determine:
- has_person_name: true if a single named crew member can be identified as this document's owner.
- last_name / first_name: that crew member's name, split, ONLY if has_person_name is true. Empty strings otherwise. A generational suffix (Jr, Sr, II, III, etc.) stays attached to last_name.

Return ONLY a JSON object with exactly these keys: {"has_person_name": true or false, "last_name": "...", "first_name": "..."}
No explanation. No markdown. No code fences. JSON object only."""

CREW_PAYROLL_SYSTEM_PROMPT = """You are analyzing a Crew Payroll document submitted as part of a film/TV production's Accounts Payable packet -- always something a payroll company (e.g. "New C.A.P.S. LLC," "TakeOne Network Corp. DBA Wrapbook") produced for the production's crew as a whole, never any one crew member's own invoice. Ignore individual employee names anywhere in the document entirely -- they never matter for naming this file. It will be one of:

1. A per-batch INVOICE -- the payroll company billing the production for one batch of crew wages, covering a given batch/work-date period. Typically bundles several report pages together: an Invoice summary, an "Invoice Fee Summary," a "Fringe Recap Report," a "Wage/Payroll Register" (per-employee paystub-style breakdown), and/or a "Payroll Check Register" -- with line items like Corporate/Taxable/Gross Wages, Union Fringes, Employer Taxes, Workers Compensation, and Handling Fees. Its own reference number for this specific batch may be labeled "Invoice Number," "Payroll ID," or similar -- never the internal "Invoice Group," "Batch," or Project/Job number.

2. A whole-project FRINGE REPORT -- a single wide table, one row per crew member, of wages/reimbursements/fringes/benefits, covering the ENTIRE project's date range at once rather than one batch. Not tied to any invoice number.

3. A whole-project PAYROLL REGISTER -- a running ledger of every payment across the whole project's date range, often hundreds of pages, titled "Payroll Register." Not tied to any invoice number.

4/5. A whole-project PAYROLL LOG -- a "PAYROLL LOG" table (columns like Line, Payee, PO, Rate, Days, Taxable/Non-taxable, Total, etc.) covering every payroll line item for the whole project. Comes in two variants with identical titles/structure, distinguished only by row order: "by Line #" is sorted by ascending Line number, so consecutive rows are usually different payees; "by Payee" is grouped by payee instead, so the same payee's name typically repeats across several consecutive rows before moving to the next one.

Determine:
- doc_kind: "invoice" for case 1, "fringe_report" for case 2, "payroll_register" for case 3, "payroll_log_by_line" or "payroll_log_by_payee" for case 4/5 depending on that row-order tell.
- company_name: ONLY for doc_kind "invoice" -- the payroll company's own short/common name. Prefer its own prominent header/logo/address block (e.g. the big "THE TEAM COMPANIES" logo and "The TEAM Companies, 2300 Empire Avenue..." address on its own invoice) over a smaller "Employer of Record" field elsewhere on the same page, which can name a different, less-recognizable legal entity the payroll company processes this particular job's checks through (e.g. "Talent, Ent. & Media Svcs, LLC" on that same Team Companies invoice) -- the header/letterhead name is the one to use, not the Employer of Record. Normalize whichever name you use the way someone would casually refer to it in conversation: drop a leading "New "/legal suffix like "LLC"/"Inc"/"Corp", drop a "DBA" business-name wrapper down to just the brand name that follows it, and collapse a dotted initialism into a plain acronym (e.g. "New C.A.P.S. LLC" -> "CAPS", "TakeOne Network Corp. DBA Wrapbook" -> "Wrapbook"). Never the production company being billed (the "Bill To"/"To" party). Empty string for every other doc_kind.
- invoice_number: ONLY for doc_kind "invoice" -- that batch's own reference number, exactly as printed. Prefer a field literally labeled "Invoice Number"/"Invoice #" when one is present; only fall back to another field the payroll company uses instead (e.g. "Payroll ID") when no "Invoice Number" field exists at all. Never the internal "Invoice Group," "Batch," "Client #," or Project/Job number. Empty string if none is printed, or for every other doc_kind.

Return ONLY a JSON object with exactly these keys: {"doc_kind": "invoice" | "fringe_report" | "payroll_register" | "payroll_log_by_line" | "payroll_log_by_payee", "company_name": "...", "invoice_number": "..."}
No explanation. No markdown. No code fences. JSON object only."""

MISC_FORM_MULTI_SYSTEM_PROMPT = """You are analyzing a multi-page PDF that may contain one or more separate standardized production paperwork forms -- forms that don't belong to any of Residency, Diversity, Vendor, Petty Cash, ProdCC, or Crew Payroll. Examples include a state tax withholding return (e.g. Georgia's "G-7 Quarterly Return for Film Payer," usually filed by the payroll company on the production's behalf), a crew member's time card, or a Kit/Box Rental fee record (a "Kit / Box Fee" summary page -- Employee, Amount, Date -- paired with its own "Box/Kit Rental Inventory" detail page) -- more form types beyond these can appear too. Each individual form belongs to either ONE company or ONE specific named crew member, and is tied to one specific date/period.

A single form is sometimes 1 page and sometimes spans several consecutive pages (e.g. a tax return with an attached schedule, or a Kit Rental's fee-summary page plus its own inventory page) -- pages belonging to the same form never get split apart, and different forms never share a page. The SAME company or crew member can appear multiple times in a row throughout the document, each occurrence a completely separate form instance with its own date -- e.g. one crew member submitting several dated Kit Rental fee records back to back, each its own summary+inventory pair. Never merge repeated occurrences of the same person/company into one form just because the name matches; a fresh repetition of the form's own header/summary block (even with identical-looking boilerplate) marks the start of a new, separate instance.

Look at every page, in the order given, and group them into distinct forms. For each form, report which pages belong to it and extract its fields.

For each form:
- pages: a list of the page numbers belonging to this form (1-indexed, in the order given).
- owner_type: "company" if this form is filed/owned by a company/production/payroll entity as a whole (e.g. a tax withholding return filed by the payroll company on the production's behalf), or "person" if it belongs to one specific named crew member (e.g. a time card, a Kit Rental).
- company_name: ONLY if owner_type is "company" -- that company's name, exactly as shown on the form's own header/letterhead (e.g. "WRAPBOOK"). Empty string otherwise.
- last_name / first_name: ONLY if owner_type is "person" -- that crew member's name, split. Empty strings otherwise. A generational suffix (Jr, Sr, II, III, etc.) stays attached to last_name.
- form_type: a short, human-readable label for what this specific form actually is, derived from its own visible title (e.g. "G-7 QUARTERLY RETURN FOR FILM PAYER" -> "G-7 Form", a timecard's own heading -> "Time Card", a "Kit / Box Fee" summary -> "Kit Rental") -- keep it short and in Title Case, not the form's full legal title verbatim.
- relevant_date: the ONE date on this form that identifies which specific period/instance it represents (e.g. the "Period Ending" date on a tax return, the work-week/pay-period-ending date on a time card, the fee "Date" on a Kit Rental), in exactly MM/DD/YYYY. A form can show several date-like fields -- some may be blank template fields with no value entered at all; use whichever one is actually filled in and genuinely identifies this specific instance, never a blank one. Empty string only if no date is filled in anywhere on the form.

Every page must belong to exactly one form's "pages" list -- never omit a page and never assign one page to two forms.

SPECIAL CASE -- a W-9 (or a similar purely-supporting tax/identity document) is only ever its own separate form when it is the ONLY thing present in the entire PDF. If a W-9 appears attached behind (or in front of) a different, more substantial form in this same PDF -- e.g. a Deal Memo followed by that same person's W-9 -- do NOT give it its own entry: fold its page(s) into the pages list of whichever form it's supporting instead, so the combined PDF stays together as one output file named after that primary form. Only report a standalone "form_type": "W9" entry when nothing else at all shares the PDF with it.

Return ONLY a JSON object with exactly this key: {"forms": [{"pages": [...], "owner_type": "company" | "person", "company_name": "...", "last_name": "...", "first_name": "...", "form_type": "...", "relevant_date": "..."}, ...]}
No explanation. No markdown. No code fences. JSON object only."""


def build_vendor_system_prompt(batch: BatchContext) -> str:
    """Inserts an explicit "this is the client, never the biller" hint for
    every named entity in the batch context -- without it, a ProdCo/Agency/
    Client name that happens to look distinctive can get mistaken for the
    vendor itself, especially on a receipt or a folio with no clear Bill To
    block to signal "this is the client, not the biller." This is also the
    mechanism the older tool lacked, which is why it sometimes renamed an
    invoice after who was being billed instead of who was doing the billing."""
    names = [
        n for n in (batch.client_name, batch.agency_name, batch.prodco_name)
        if n and n.strip()
    ]
    if not names:
        prodco_line = ""
    else:
        joined = ", ".join(f'"{n.strip()}"' for n in names)
        prodco_line = (
            f'\nThe production company / agency / client on this job may appear on these documents as '
            f'{joined} (or an LLC/dba variant of one of these names, e.g. "{names[0]} Productions, LLC"). '
            f'These are ALWAYS potential CLIENT / "Bill To" parties -- never the vendor issuing the document, '
            f'no matter where a name appears on the page or how prominently. Never let any of them become '
            f'has_person_name/last_name/first_name, has_company_name/company_name, or hotel_name below.\n'
        )
    return VENDOR_SYSTEM_PROMPT_TEMPLATE.replace("{prodco_line}", prodco_line)


# ── PDF / image / vision helpers ──────────────────────────────────────────────

try:
    import pillow_heif
    pillow_heif.register_heif_opener()  # lets Image.open() read iPhone HEIC/HEIF photos too
except ImportError:
    pass


def _flatten_to_rgb(img):
    """PDF has no real transparency model -- a transparent PNG/GIF/etc.
    composited straight in would render with an undefined (often black)
    background in some viewers, so anything with real alpha gets flattened
    onto white first. Everything else just gets a normal mode conversion."""
    if img.mode in ("RGBA", "LA") or (img.mode == "P" and "transparency" in img.info):
        img = img.convert("RGBA")
        background = Image.new("RGB", img.size, (255, 255, 255))
        background.paste(img, mask=img.split()[-1])
        return background
    if img.mode != "RGB":
        return img.convert("RGB")
    return img


def load_source_as_pdf_bytes(raw_bytes: bytes, filename: str) -> bytes:
    """Returns valid single-document PDF bytes regardless of source format.
    A PDF (by extension or by content, in case someone renamed/mislabeled
    it) passes through untouched. Anything else is handed to Pillow as an
    image -- covers PNG/JPG same as before, plus BMP/GIF/TIFF/WEBP/HEIC-HEIF
    (iPhone photos) and effectively any other format Pillow understands,
    regardless of the file's own extension. A multi-page/multi-frame image
    (e.g. a scanner's multi-page TIFF) becomes a multi-page PDF, feeding the
    same downstream pipeline (including Residency's page-splitting) as a
    native multi-page PDF would. Raises if the file is neither a valid PDF
    nor an image Pillow can open -- the caller already treats that as an
    unreadable file, so no separate extension whitelist is needed here."""
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".pdf" or raw_bytes[:5] == b"%PDF-":
        return raw_bytes

    try:
        img = Image.open(io.BytesIO(raw_bytes))
    except Exception:
        # PIL's own exception message embeds a raw BytesIO object repr,
        # which is meaningless in a user-facing reason_detail -- replace it
        # with a clean one instead of letting that string surface.
        raise ValueError(f"not a PDF and not a recognizable image format ({ext or 'no extension'})")
    n_frames = getattr(img, "n_frames", 1)
    frames = []
    for i in range(n_frames):
        img.seek(i)
        # exif_transpose bakes in a phone photo's stored orientation flag
        # before it ever reaches the vision-based rotation check, so that
        # check only has to catch genuine scanner/photo mistakes, not every
        # normally-oriented phone photo.
        frame = ImageOps.exif_transpose(img.copy())
        frames.append(_flatten_to_rgb(frame))

    buf = io.BytesIO()
    if len(frames) == 1:
        frames[0].save(buf, format="PDF")
    else:
        frames[0].save(buf, format="PDF", save_all=True, append_images=frames[1:])
    return buf.getvalue()


def render_pdf_bytes_to_images_b64(pdf_bytes: bytes, dpi_scale: float = 2.0, max_pages: int = 3) -> list:
    """Renders up to max_pages (default 3, the widest need across the three
    starting types -- vendor documents sometimes lead with a PO cover sheet
    pushing the real invoice to page 2-3). Rendered once per file and reused
    for both the classify call and the type-specific extract call."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []
    for i, page in enumerate(doc):
        if max_pages and i >= max_pages:
            break
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi_scale, dpi_scale))
        images.append(base64.b64encode(pix.tobytes("png")).decode())
    doc.close()
    return images


MAX_PAGES_TO_INSPECT = 30


def extract_pages_as_pdf(pdf_bytes: bytes, page_numbers_1indexed: list) -> bytes:
    """Slices specific pages (1-indexed, in the given order) out of a PDF
    into a new standalone PDF -- used when one uploaded file contains
    multiple people's residency documents (e.g. a batch-scanned stack of
    licenses), so each person ends up as their own output file."""
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = fitz.open()
    for pn in page_numbers_1indexed:
        out.insert_pdf(src, from_page=pn - 1, to_page=pn - 1)
    result = out.tobytes()
    out.close()
    src.close()
    return result


def looks_like_refusal(raw: str) -> bool:
    s = raw.strip()
    if not s:
        return True
    if "{" in s[:20] or s.startswith("```"):
        return False
    refusal_openers = ("i'm sorry", "i am sorry", "i cannot", "i can't", "sorry,", "i apologize")
    return s.lower().startswith(refusal_openers) or "{" not in s


def parse_json_object(raw: str):
    try:
        return json.loads(raw)
    except Exception:
        pass
    m = re.search(r"\{.*\}", raw, re.DOTALL)
    if m:
        try:
            return json.loads(m.group())
        except Exception:
            pass
    return None


def call_gpt_json(client, images_b64, system_prompt, user_text, max_tokens=500, detail="high"):
    content = [{"type": "text", "text": user_text}]
    for img in images_b64:
        content.append({
            "type": "image_url",
            "image_url": {"url": f"data:image/png;base64,{img}", "detail": detail},
        })
    resp = client.chat.completions.create(
        model="gpt-4o",
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content},
        ],
        temperature=0,
        max_tokens=max_tokens,
    )
    return resp.choices[0].message.content.strip()


def call_claude_json(client, images_b64, system_prompt, user_text, max_tokens=500):
    content = []
    for img in images_b64:
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/png", "data": img},
        })
    content.append({"type": "text", "text": user_text})
    resp = client.messages.create(
        model="claude-sonnet-5",
        max_tokens=max_tokens,
        system=system_prompt,
        messages=[{"role": "user", "content": content}],
    )
    text_block = next((b for b in resp.content if b.type == "text"), None)
    return text_block.text.strip() if text_block else ""


async def classify_document(images_b64, user_text, system_prompt, openai_client, anthropic_client, loop, detail="high"):
    """Tries GPT-4o first; falls back to Claude if GPT refuses/errors/comes
    back empty. Returns (parsed_dict_or_None, provider_used). Runs the
    (synchronous) SDK calls in the executor so this can be awaited alongside
    other files' calls under a semaphore. detail="low" is worth passing for
    a coarse judgment call (e.g. orientation) made over many images at
    once, where GPT-4o's high-detail mode would multiply token cost for no
    real accuracy benefit."""
    import functools
    try:
        raw = await loop.run_in_executor(
            None, functools.partial(call_gpt_json, openai_client, images_b64, system_prompt, user_text, detail=detail),
        )
        if looks_like_refusal(raw):
            raise ValueError("refusal")
        parsed = parse_json_object(raw)
        if parsed is None:
            raise ValueError("unparseable")
        return parsed, "openai"
    except Exception:
        if anthropic_client is None:
            return None, "openai"
        try:
            raw = await loop.run_in_executor(
                None, functools.partial(call_claude_json, anthropic_client, images_b64, system_prompt, user_text),
            )
            parsed = parse_json_object(raw)
            return (parsed, "anthropic") if parsed is not None else (None, "anthropic")
        except Exception:
            return None, "anthropic"


def has_form_fields(pdf_bytes: bytes) -> bool:
    """Cheap, local check (no LLM call) for whether this PDF has interactive
    AcroForm fields at all -- used to decide whether flatten_form_fields is
    worth running."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        return any(next(page.widgets(), None) is not None for page in doc)
    finally:
        doc.close()


def flatten_form_fields(pdf_bytes: bytes, dpi_scale: float = 3.0) -> bytes:
    """Rebuilds a PDF by rendering each page to a flat image and embedding
    it into a brand-new PDF -- the same effect as "Print to PDF". Some of
    the standard fillable forms submitted here (confirmed real case: the
    Illinois DCEO diversity form) have been seen with badly-scrambled
    AcroForm field name/value pairs internally -- e.g. a field literally
    named "Zip" holding the city, and a field named "Date" holding the
    signature -- while each widget still renders correctly at its own
    on-page position regardless of its internal name. Flattening only
    removes the broken interactive structure; the page's actual appearance
    (what get_pixmap() already renders, and what a human sees) is
    unaffected, so nothing downstream that reads the rendered page is
    fixing anything real -- this exists purely so the OUTPUT file itself
    doesn't carry the same scrambled structure forward for some other tool
    to trip over later."""
    src = fitz.open(stream=pdf_bytes, filetype="pdf")
    out = fitz.open()
    for page in src:
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi_scale, dpi_scale))
        new_page = out.new_page(width=page.rect.width, height=page.rect.height)
        new_page.insert_image(new_page.rect, pixmap=pix)
    result = out.tobytes()
    out.close()
    src.close()
    return result


async def correct_page_orientations(pdf_bytes: bytes, openai_client, anthropic_client, loop) -> bytes:
    """Detects and fixes any page that's sideways or upside-down (a phone
    photo taken in the wrong orientation, or a scanner fed the page wrong)
    before any classification/extraction happens -- everything downstream
    just sees an upright document, no matter how it was actually scanned.

    Sets the PDF's own per-page /Rotate attribute rather than baking in a
    pixel-level transform: every standard PDF viewer honors it, and so does
    our own get_pixmap() rendering used everywhere else in this file, so
    the fix propagates for free through classification, extraction, and
    (for Residency) the page-splitting step -- none of that code needs to
    know this ever happened."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n = min(doc.page_count, MAX_PAGES_TO_INSPECT)
    if n == 0:
        doc.close()
        return pdf_bytes

    images = []
    for i in range(n):
        pix = doc[i].get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
        images.append(base64.b64encode(pix.tobytes("png")).decode())

    user_text = "Determine the needed rotation for each page, in order, as described in the system prompt."
    parsed, _provider = await classify_document(
        images, user_text, ORIENTATION_PROMPT, openai_client, anthropic_client, loop, detail="low",
    )
    pages_info = (parsed or {}).get("pages") if parsed else None
    if not isinstance(pages_info, list):
        doc.close()
        return pdf_bytes

    changed = False
    for i, info in enumerate(pages_info[:n]):
        if not isinstance(info, dict):
            continue
        try:
            rotation = int(info.get("rotation", 0))
        except (TypeError, ValueError):
            rotation = 0
        if rotation in (90, 180, 270):
            page = doc[i]
            page.set_rotation((page.rotation + rotation) % 360)
            changed = True

    result = doc.tobytes() if changed else pdf_bytes
    doc.close()
    return result


# ── Name / filename helpers ───────────────────────────────────────────────────

def sanitize_component(s: str) -> str:
    s = INVALID_FILENAME_CHARS.sub("", s or "").strip()
    return s.rstrip(". ")


def proper_case(s: str) -> str:
    """Word-by-word Proper/Title Case, with two exceptions: LLC and INC
    (however punctuated) always render fully uppercase rather than
    title-cased, and a single letter directly after an apostrophe stays
    lowercase (a possessive/contraction suffix -- "Joe's", "McDonald's" --
    rather than a second name part, unlike "O'Brien"/"D'Angelo" where the
    run after the apostrophe is longer and does get title-cased). A run
    starting with "Mc" also capitalizes the letter right after it --
    "MCGEE"/"mcgee" -> "McGee" -- since that's effectively universal for
    real surnames and this field is always a person's name, never
    arbitrary text; "Mac" is deliberately left alone since "Mack"/"Macy"/
    "Mace" are common surnames in their own right where that doesn't apply.
    Capitalizes each contiguous run of letters independently:
    "REAGAN-BARNES" -> "Reagan-Barnes", "O'BRIEN" -> "O'Brien",
    "JOE'S BAKERY" -> "Joe's Bakery"."""
    def cap_word(word):
        core = re.sub(r"[^a-zA-Z]", "", word).lower()
        if core in _UPPERCASE_EXCEPTIONS:
            return word.upper()

        def cap_run(m):
            run = m.group(0)
            preceded_by_apostrophe = m.start() > 0 and word[m.start() - 1] == "'"
            if preceded_by_apostrophe and len(run) == 1:
                return run.lower()
            capped = run[:1].upper() + run[1:].lower()
            if len(capped) > 2 and capped[:2] == "Mc":
                capped = capped[:2] + capped[2].upper() + capped[3:]
            return capped

        return re.sub(r"[A-Za-z]+", cap_run, word)

    s = (s or "").strip()
    return " ".join(cap_word(w) for w in s.split(" ")) if s else s


def tokenize(s: str) -> set:
    return set(re.findall(r"[a-z]+", (s or "").lower()))


def names_overlap(a: str, b: str) -> bool:
    ta, tb = tokenize(a), tokenize(b)
    return bool(ta and tb and (ta & tb))


def name_looks_like_person(text: str) -> bool:
    if not text or not text.strip():
        return False
    words = re.findall(r"[a-zA-Z']+", text.lower())
    if not words or any(w in _BUSINESS_WORDS for w in words):
        return False
    return 1 <= len(words) <= 4


def reformat_date_for_filename(mmddyyyy: str) -> str:
    """MM/DD/YYYY -> YYYY_MM_DD, for the Driver's License / State ID
    filename pattern: "Last, First (ST) YYYY_MM_DD.pdf"."""
    s = (mmddyyyy or "").strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", s)
    if m:
        mm, dd, yyyy = m.groups()
        return f"{yyyy}_{int(mm):02d}_{int(dd):02d}"
    return s.replace("/", "_")


def split_person_name(full_name: str):
    """'First Last' (no comma) -> (last, first); 'Last, First' -> (last, first).
    A trailing generational suffix stays attached to the last name."""
    full_name = (full_name or "").strip()
    if "," in full_name:
        last, _, first = full_name.partition(",")
        return last.strip(), first.strip()
    parts = full_name.split()
    if len(parts) >= 3 and parts[-1].strip(".").lower() in _NAME_SUFFIXES:
        return f"{parts[-2]} {parts[-1]}", " ".join(parts[:-2])
    if len(parts) >= 2:
        return parts[-1], " ".join(parts[:-1])
    return full_name, ""


def build_residency_filename(doc_type: str, last: str, first: str, state: str, exp_date: str):
    last_s, first_s = sanitize_component(proper_case(last)), sanitize_component(proper_case(first))
    if not last_s or not first_s:
        return None
    if doc_type in ("drivers_license", "state_id"):
        state_s = sanitize_component((state or "").strip().upper())
        exp_s = sanitize_component(reformat_date_for_filename(exp_date))
        if not state_s or not exp_s:
            return None
        return f"{last_s}, {first_s} ({state_s}) {exp_s}.pdf"
    if doc_type == "i9":
        return f"{last_s}, {first_s} - I9.pdf"
    return None


def build_diversity_filename(last: str, first: str):
    last_s, first_s = sanitize_component(proper_case(last)), sanitize_component(proper_case(first))
    if not last_s or not first_s:
        return None
    return f"{last_s}, {first_s} - Diversity Form.pdf"


def build_petty_cash_filename(last: str, first: str):
    last_s, first_s = sanitize_component(proper_case(last)), sanitize_component(proper_case(first))
    if not last_s or not first_s:
        return None
    return f"{last_s}, {first_s} - Petty Cash.pdf"


def build_prodcc_filename(last: str, first: str):
    last_s, first_s = sanitize_component(proper_case(last)), sanitize_component(proper_case(first))
    if not last_s or not first_s:
        return None
    return f"{last_s}, {first_s} - CC Reimb.pdf"


def build_misc_form_filename(owner_type: str, last: str, first: str, company: str, form_type: str, relevant_date: str):
    """Shared naming convention for the "random forms" catch-all: a
    company-owned form is "{Company} - {Form Type} {Date}.pdf", a
    crew-member-owned form is "{Last}, {First} - {Form Type} {Date}.pdf"
    -- e.g. "Wrapbook - G-7 Form 2024_03_31.pdf" or
    "Marivee, Cade - Time Card 2023_12_31.pdf". The date suffix is omitted
    entirely (not left as a trailing space) when none was found."""
    form_type_s = sanitize_component(proper_case((form_type or "").strip()))
    if not form_type_s:
        return None
    date_s = reformat_date_for_filename(relevant_date) if (relevant_date or "").strip() else ""
    suffix = f"{form_type_s} {date_s}" if date_s else form_type_s

    if owner_type == "company":
        company_s = sanitize_component(proper_case(company))
        if not company_s:
            return None
        return f"{company_s} - {suffix}.pdf"
    if owner_type == "person":
        last_s, first_s = sanitize_component(proper_case(last)), sanitize_component(proper_case(first))
        if not last_s or not first_s:
            return None
        return f"{last_s}, {first_s} - {suffix}.pdf"
    return None


def build_uncovered_receipts_filename(original_filename: str):
    """A bare stack of receipts with no cover sheet at all can't be named by
    content -- keep whatever the original filename was (it may be the only
    hint of whose receipts these are) and just prefix it so it sorts to the
    top of the Petty Cash folder for manual review, same "aaa" convention
    used for an unreadable/unable-to-rename file elsewhere in this pipeline."""
    base = os.path.splitext(original_filename or "")[0]
    base_s = sanitize_component(base) or "receipts"
    return f"aaa{base_s}.pdf"


def build_freelance_filename(has_person: bool, last: str, first: str, has_company: bool, company: str, invoice_number: str):
    # A freelance crew invoice is always named by the individual alone, even
    # when it's billed through their own loan-out/DBA/LLC/studio name (e.g.
    # "NPD Media (Nathan Destro)", "Location Media LLC", "Lavi Toma
    # Styling", "Alexa Viscius Studio") -- confirmed by real reference
    # examples that drop the company entirely from the filename in every
    # such case. The company name only gets used as a last resort when no
    # person name could be extracted at all.
    last_s = sanitize_component(proper_case(last)) if has_person else ""
    first_s = sanitize_component(proper_case(first)) if has_person else ""
    company_s = sanitize_component(proper_case(company)) if has_company else ""
    num_s = sanitize_component((invoice_number or "").strip())
    invoice_suffix = f"Invoice {num_s}" if num_s else "Invoice"

    if has_person and not (last_s and first_s):
        return None
    if has_company and not company_s:
        return None
    if has_person:
        return f"{last_s}, {first_s} - {invoice_suffix}.pdf"
    if has_company:
        return f"{company_s} - {invoice_suffix}.pdf"
    return None


def build_vendor_invoice_filename(company: str, invoice_number: str):
    company_s = sanitize_component(proper_case(company))
    if not company_s:
        return None
    num_s = sanitize_component((invoice_number or "").strip())
    invoice_suffix = f"Invoice {num_s}" if num_s else "Invoice"
    return f"{company_s} - {invoice_suffix}.pdf"


def build_vendor_invoice_filename_by_po(company: str, po_number: str):
    """The alternate Vendor Invoice naming convention: "PO{number} -
    {Vendor Name}.pdf". Only usable when a PO number was actually printed
    on this specific invoice's cover sheet -- returns None otherwise so
    the caller falls back to build_vendor_invoice_filename instead."""
    company_s = sanitize_component(proper_case(company))
    po_s = sanitize_component((po_number or "").strip())
    if not company_s or not po_s:
        return None
    return f"PO{po_s} - {company_s}.pdf"


def build_vendor_receipt_filename(company: str):
    company_s = sanitize_component(proper_case(company))
    return f"{company_s} - receipt.pdf" if company_s else None


def build_hotel_folio_filename(hotel_name: str, folio_number: str):
    hotel_s = sanitize_component(proper_case(hotel_name))
    if not hotel_s:
        return None
    num_s = sanitize_component((folio_number or "").strip())
    suffix = f"Folio {num_s}" if num_s else "Folio"
    return f"{hotel_s} - {suffix}.pdf"


# ── Sender-mismatch detection (Vendor only) ───────────────────────────────────

def _expected_biller_name(batch: BatchContext) -> str:
    # No "client" case -- received_from can only ever be "prodco" or
    # "agency" (validated at the API boundary), since a batch is never
    # "received from" the Client. client_name still participates in
    # check_sender_mismatch below, just only ever as an "other" candidate.
    return {
        "prodco": batch.prodco_name,
        "agency": batch.agency_name,
    }.get(batch.received_from, "")


def check_sender_mismatch(bill_to_name: str, batch: BatchContext) -> str:
    """Returns a human-readable mismatch detail if bill_to_name confidently
    matches a DIFFERENT named entity than received_from implies, else "".
    Only flags a confident mismatch against one of the batch's OTHER two
    named entities -- never merely "didn't match the expected one", since
    that alone is too weak a signal (could just be phrased differently)."""
    bill_to_name = (bill_to_name or "").strip()
    if not bill_to_name or not batch.received_from:
        return ""
    expected = _expected_biller_name(batch)
    if expected and names_overlap(bill_to_name, expected):
        return ""
    others = {"client": batch.client_name, "agency": batch.agency_name, "prodco": batch.prodco_name}
    for key, name in others.items():
        if key == batch.received_from or not name:
            continue
        if names_overlap(bill_to_name, name):
            return (
                f"Bill-To reads '{bill_to_name}', but batch says received_from={batch.received_from} "
                f"({expected or 'not provided'})"
            )
    return ""


# ── Per-family extraction ─────────────────────────────────────────────────────

async def _extract_residency_multi(pdf_bytes, openai_client, anthropic_client, loop):
    """Residency is the one family that can't assume "one uploaded file = one
    person" -- a coordinator batch-scanning a stack of physical IDs produces a
    single multi-page PDF containing several different people's cards, each
    spanning 1 page (front only) or 2 (front+back). This renders every page
    (up to MAX_PAGES_TO_INSPECT) and asks the model to group them by person in
    one call, then physically splits the PDF so each person becomes their own
    output file -- one uploaded file can produce many DocResults here, unlike
    every other family."""
    images = render_pdf_bytes_to_images_b64(pdf_bytes, max_pages=MAX_PAGES_TO_INSPECT)
    if not images:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="residency", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=pdf_bytes, reason_code=REASON_UNREADABLE,
            reason_detail="No renderable pages found",
        )]

    user_text = (
        "Identify every person's ID document across all pages of this PDF and group "
        "pages accordingly, as described in the system prompt."
    )
    parsed, provider = await classify_document(
        images, user_text, RESIDENCY_MULTI_SYSTEM_PROMPT, openai_client, anthropic_client, loop,
    )
    documents = (parsed or {}).get("documents") if parsed else None
    if not isinstance(documents, list) or not documents:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="residency", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=pdf_bytes, reason_code=REASON_UNREADABLE,
            reason_detail="Could not read document (both providers failed)", provider=provider,
        )]

    total_pages = len(images)
    results = []
    for doc in documents:
        page_nums = [p for p in (doc.get("pages") or []) if isinstance(p, int) and 1 <= p <= total_pages]
        if not page_nums:
            continue  # nothing usable to split out for this group

        try:
            sub_pdf_bytes = extract_pages_as_pdf(pdf_bytes, page_nums)
        except Exception as e:
            results.append(DocResult(
                bucket=BUCKET_NOT_READABLE, doc_type="residency", subfolder=FOLDER_NOT_READABLE,
                output_pdf_bytes=pdf_bytes, reason_code=REASON_UNREADABLE,
                reason_detail=f"Could not split page(s) {page_nums} into their own file: {e}", provider=provider,
            ))
            continue

        doc_type = str(doc.get("document_type", "unknown")).strip().lower()
        if doc_type not in ("drivers_license", "state_id", "i9"):
            results.append(DocResult(
                bucket=BUCKET_NOT_READABLE, doc_type="residency", subfolder=FOLDER_NOT_READABLE,
                output_pdf_bytes=sub_pdf_bytes, reason_code=REASON_UNCLASSIFIED,
                reason_detail=f"Unrecognized document type ({doc_type!r}) on page(s) {page_nums}", provider=provider,
            ))
            continue

        last, first = doc.get("last_name", ""), doc.get("first_name", "")
        state = doc.get("state", "")
        exp_date = doc.get("expiration_date", "")
        new_name = build_residency_filename(doc_type, last, first, state, exp_date)
        if not new_name:
            if not last or not first:
                missing = REASON_MISSING_NAME
            elif doc_type in ("drivers_license", "state_id") and not state:
                missing = REASON_MISSING_STATE
            else:
                missing = REASON_MISSING_DATE
            results.append(DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="residency", subfolder=FOLDER_RESIDENCY,
                output_pdf_bytes=sub_pdf_bytes, reason_code=missing,
                reason_detail=(f"Missing/unreadable name, state, or expiration date on page(s) {page_nums} "
                                f"(type={doc_type}, last={last!r}, first={first!r}, state={state!r}, exp={exp_date!r})"),
                provider=provider,
            ))
            continue

        results.append(DocResult(
            bucket=BUCKET_RENAMED, doc_type="residency", subfolder=FOLDER_RESIDENCY,
            output_pdf_bytes=sub_pdf_bytes, new_filename=new_name, provider=provider,
        ))

    if not results:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="residency", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=pdf_bytes, reason_code=REASON_UNCLASSIFIED,
            reason_detail="Model returned no usable page groupings", provider=provider,
        )]
    return results


async def _extract_misc_form_multi(pdf_bytes, openai_client, anthropic_client, loop):
    """The "random forms" catch-all (G-7, Time Card, and whatever else
    comes up) can't assume one uploaded file = one form either -- a
    coordinator can batch-scan several distinct forms into one PDF. Same
    render-everything/group/split pattern as Residency."""
    images = render_pdf_bytes_to_images_b64(pdf_bytes, max_pages=MAX_PAGES_TO_INSPECT)
    if not images:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="misc_form", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=pdf_bytes, reason_code=REASON_UNREADABLE,
            reason_detail="No renderable pages found",
        )]

    user_text = (
        "Identify every distinct form across all pages of this PDF and group "
        "pages accordingly, as described in the system prompt."
    )
    parsed, provider = await classify_document(
        images, user_text, MISC_FORM_MULTI_SYSTEM_PROMPT, openai_client, anthropic_client, loop,
    )
    forms = (parsed or {}).get("forms") if parsed else None
    if not isinstance(forms, list) or not forms:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="misc_form", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=pdf_bytes, reason_code=REASON_UNREADABLE,
            reason_detail="Could not read document (both providers failed)", provider=provider,
        )]

    total_pages = len(images)
    results = []
    for form in forms:
        page_nums = [p for p in (form.get("pages") or []) if isinstance(p, int) and 1 <= p <= total_pages]
        if not page_nums:
            continue  # nothing usable to split out for this group

        try:
            sub_pdf_bytes = extract_pages_as_pdf(pdf_bytes, page_nums)
        except Exception as e:
            results.append(DocResult(
                bucket=BUCKET_NOT_READABLE, doc_type="misc_form", subfolder=FOLDER_NOT_READABLE,
                output_pdf_bytes=pdf_bytes, reason_code=REASON_UNREADABLE,
                reason_detail=f"Could not split page(s) {page_nums} into their own file: {e}", provider=provider,
            ))
            continue

        owner_type = str(form.get("owner_type", "")).strip().lower()
        company = form.get("company_name", "")
        last, first = form.get("last_name", ""), form.get("first_name", "")
        form_type = form.get("form_type", "")
        relevant_date = form.get("relevant_date", "")

        # Each distinct form_type gets its own subfolder under "Other Forms"
        # (e.g. "Other Forms/G-7 Form", "Other Forms/Kit Rental") rather than
        # dumping every kind of form into one flat folder together. Falls
        # back to the flat folder only if form_type itself couldn't be read.
        form_type_s = sanitize_component(proper_case((form_type or "").strip()))
        subfolder = f"{FOLDER_OTHER_FORMS}/{form_type_s}" if form_type_s else FOLDER_OTHER_FORMS

        new_name = build_misc_form_filename(owner_type, last, first, company, form_type, relevant_date)
        if not new_name:
            if owner_type not in ("company", "person"):
                reason = REASON_UNCLASSIFIED
            elif owner_type == "company":
                reason = REASON_MISSING_COMPANY
            else:
                reason = REASON_MISSING_NAME
            results.append(DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="misc_form", subfolder=subfolder,
                output_pdf_bytes=sub_pdf_bytes, reason_code=reason,
                reason_detail=(f"Missing/unreadable owner or form details on page(s) {page_nums} "
                                f"(owner_type={owner_type!r}, company={company!r}, last={last!r}, first={first!r}, "
                                f"form_type={form_type!r}, date={relevant_date!r})"),
                provider=provider,
            ))
            continue

        results.append(DocResult(
            bucket=BUCKET_RENAMED, doc_type="misc_form", subfolder=subfolder,
            output_pdf_bytes=sub_pdf_bytes, new_filename=new_name, provider=provider,
        ))

    if not results:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="misc_form", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=pdf_bytes, reason_code=REASON_UNCLASSIFIED,
            reason_detail="Model returned no usable page groupings", provider=provider,
        )]
    return results


async def _extract_diversity(images_b64, openai_client, anthropic_client, loop, orig_filename_hint):
    user_text = "Extract the fields described in the system prompt from this DCEO tracking sheet."
    parsed, provider = await classify_document(
        images_b64, user_text, DIVERSITY_SYSTEM_PROMPT, openai_client, anthropic_client, loop,
    )
    if parsed is None:
        return DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="diversity", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=b"", reason_code=REASON_UNREADABLE,
            reason_detail="Could not read document (both providers failed)", provider=provider,
        )

    employee_name = str(parsed.get("employee_name_field", "")).strip()
    looks_like_person = bool(parsed.get("looks_like_person")) and name_looks_like_person(employee_name)
    signature_guess = str(parsed.get("signature_name_guess", "")).strip()

    if looks_like_person:
        last, first = split_person_name(employee_name)
    else:
        # The Employee Name field reads as a company (a common real case:
        # a loan-out entity named after its owner, e.g. "Nolis Anderson
        # Photography LLC" for a photographer named Nolis Anderson) -- a
        # signature reading "Nolis Anderson" confirms this directly, and
        # doesn't need the existing filename to agree too. Checking the
        # company name FIRST also covers the common case of a generic/
        # uninformative original filename ("a.pdf", "scan001.pdf") that
        # was never going to match any real name to begin with.
        filename_guess = _guess_name_from_filename(orig_filename_hint)
        confirmed = bool(signature_guess) and (
            names_overlap(signature_guess, employee_name) or names_overlap(signature_guess, filename_guess)
        )
        if confirmed:
            last, first = split_person_name(signature_guess)
        else:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="diversity", subfolder=FOLDER_DIVERSITY,
                output_pdf_bytes=b"", reason_code=REASON_LOW_CONFIDENCE,
                reason_detail=(f"Employee Name field looks like a company ({employee_name!r}) and signature "
                               f"({signature_guess!r}) didn't confirm it against either the company name "
                               f"or the existing filename"),
                provider=provider,
            )

    new_name = build_diversity_filename(last, first)
    if not new_name:
        return DocResult(
            bucket=BUCKET_UNABLE_TO_RENAME, doc_type="diversity", subfolder=FOLDER_DIVERSITY,
            output_pdf_bytes=b"", reason_code=REASON_MISSING_NAME,
            reason_detail=f"Could not derive a usable name (last={last!r}, first={first!r})", provider=provider,
        )
    return DocResult(
        bucket=BUCKET_RENAMED, doc_type="diversity", subfolder=FOLDER_DIVERSITY,
        output_pdf_bytes=b"", new_filename=new_name, provider=provider,
    )


async def _extract_petty_cash(images_b64, openai_client, anthropic_client, loop, orig_filename_hint):
    user_text = "Extract the fields described in the system prompt from this petty cash document."
    parsed, provider = await classify_document(
        images_b64, user_text, PETTY_CASH_SYSTEM_PROMPT, openai_client, anthropic_client, loop,
    )
    if parsed is None:
        return DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="petty_cash", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=b"", reason_code=REASON_UNREADABLE,
            reason_detail="Could not read document (both providers failed)", provider=provider,
        )

    if not bool(parsed.get("has_cover_sheet")):
        # A bare stack of receipts with nothing identifying whose batch this
        # is -- not an error, just nothing to confidently name from content.
        # Keep the original filename (it may be the only clue) under an
        # "aaa" prefix so it sorts to the top for manual review.
        return DocResult(
            bucket=BUCKET_RENAMED, doc_type="petty_cash", subfolder=FOLDER_PETTY_CASH,
            output_pdf_bytes=b"", new_filename=build_uncovered_receipts_filename(orig_filename_hint),
            provider=provider,
        )

    if bool(parsed.get("is_petty_cash_log")):
        return DocResult(
            bucket=BUCKET_RENAMED, doc_type="petty_cash", subfolder=FOLDER_PETTY_CASH,
            output_pdf_bytes=b"", new_filename="Petty Cash Log.pdf", provider=provider,
        )

    if not bool(parsed.get("has_person_name")):
        return DocResult(
            bucket=BUCKET_UNABLE_TO_RENAME, doc_type="petty_cash", subfolder=FOLDER_PETTY_CASH,
            output_pdf_bytes=b"", reason_code=REASON_MISSING_NAME,
            reason_detail="Could not identify a named custodian on this petty cash document", provider=provider,
        )

    last, first = parsed.get("last_name", ""), parsed.get("first_name", "")
    new_name = build_petty_cash_filename(last, first)
    if not new_name:
        return DocResult(
            bucket=BUCKET_UNABLE_TO_RENAME, doc_type="petty_cash", subfolder=FOLDER_PETTY_CASH,
            output_pdf_bytes=b"", reason_code=REASON_MISSING_NAME,
            reason_detail=f"Could not derive a usable name (last={last!r}, first={first!r})", provider=provider,
        )
    return DocResult(
        bucket=BUCKET_RENAMED, doc_type="petty_cash", subfolder=FOLDER_PETTY_CASH,
        output_pdf_bytes=b"", new_filename=new_name, provider=provider,
    )


async def _extract_prodcc(images_b64, openai_client, anthropic_client, loop):
    user_text = "Extract the fields described in the system prompt from this ProdCC reimbursement document."
    parsed, provider = await classify_document(
        images_b64, user_text, PRODCC_SYSTEM_PROMPT, openai_client, anthropic_client, loop,
    )
    if parsed is None:
        return DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="prodcc", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=b"", reason_code=REASON_UNREADABLE,
            reason_detail="Could not read document (both providers failed)", provider=provider,
        )

    if not bool(parsed.get("has_person_name")):
        return DocResult(
            bucket=BUCKET_UNABLE_TO_RENAME, doc_type="prodcc", subfolder=FOLDER_PRODCC,
            output_pdf_bytes=b"", reason_code=REASON_MISSING_NAME,
            reason_detail="Could not identify a named crew member on this ProdCC document", provider=provider,
        )

    last, first = parsed.get("last_name", ""), parsed.get("first_name", "")
    new_name = build_prodcc_filename(last, first)
    if not new_name:
        return DocResult(
            bucket=BUCKET_UNABLE_TO_RENAME, doc_type="prodcc", subfolder=FOLDER_PRODCC,
            output_pdf_bytes=b"", reason_code=REASON_MISSING_NAME,
            reason_detail=f"Could not derive a usable name (last={last!r}, first={first!r})", provider=provider,
        )
    return DocResult(
        bucket=BUCKET_RENAMED, doc_type="prodcc", subfolder=FOLDER_PRODCC,
        output_pdf_bytes=b"", new_filename=new_name, provider=provider,
    )


async def _extract_crew_payroll(images_b64, openai_client, anthropic_client, loop):
    user_text = "Extract the fields described in the system prompt from this Crew Payroll document."
    parsed, provider = await classify_document(
        images_b64, user_text, CREW_PAYROLL_SYSTEM_PROMPT, openai_client, anthropic_client, loop,
    )
    if parsed is None:
        return DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="crew_payroll", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=b"", reason_code=REASON_UNREADABLE,
            reason_detail="Could not read document (both providers failed)", provider=provider,
        )

    doc_kind = str(parsed.get("doc_kind", "")).strip().lower()
    if doc_kind == "fringe_report":
        return DocResult(
            bucket=BUCKET_RENAMED, doc_type="crew_payroll", subfolder=FOLDER_CREW_PAYROLL,
            output_pdf_bytes=b"", new_filename="Fringe Report.pdf", provider=provider,
        )
    if doc_kind == "payroll_register":
        return DocResult(
            bucket=BUCKET_RENAMED, doc_type="crew_payroll", subfolder=FOLDER_CREW_PAYROLL,
            output_pdf_bytes=b"", new_filename="Payroll Register.pdf", provider=provider,
        )
    if doc_kind == "payroll_log_by_line":
        return DocResult(
            bucket=BUCKET_RENAMED, doc_type="crew_payroll", subfolder=FOLDER_CREW_PAYROLL,
            output_pdf_bytes=b"", new_filename="Payroll Log by Line #.pdf", provider=provider,
        )
    if doc_kind == "payroll_log_by_payee":
        return DocResult(
            bucket=BUCKET_RENAMED, doc_type="crew_payroll", subfolder=FOLDER_CREW_PAYROLL,
            output_pdf_bytes=b"", new_filename="Payroll Log by Payee.pdf", provider=provider,
        )

    company = parsed.get("company_name", "")
    invoice_number = parsed.get("invoice_number", "")
    new_name = build_vendor_invoice_filename(company, invoice_number)
    if not new_name:
        return DocResult(
            bucket=BUCKET_UNABLE_TO_RENAME, doc_type="crew_payroll", subfolder=FOLDER_CREW_PAYROLL,
            output_pdf_bytes=b"", reason_code=REASON_MISSING_COMPANY,
            reason_detail=f"Could not identify the payroll company (company={company!r})", provider=provider,
        )
    return DocResult(
        bucket=BUCKET_RENAMED, doc_type="crew_payroll", subfolder=FOLDER_CREW_PAYROLL,
        output_pdf_bytes=b"", new_filename=new_name, provider=provider,
    )


async def _extract_vendor(images_b64, openai_client, anthropic_client, loop, batch: BatchContext):
    user_text = "Extract the fields described in the system prompt from this vendor document."
    system_prompt = build_vendor_system_prompt(batch)
    parsed, provider = await classify_document(
        images_b64, user_text, system_prompt, openai_client, anthropic_client, loop,
    )
    if parsed is None:
        return DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="vendor", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=b"", reason_code=REASON_UNREADABLE,
            reason_detail="Could not read document (both providers failed)", provider=provider,
        )

    is_hotel_folio = bool(parsed.get("is_hotel_folio"))
    is_receipt = bool(parsed.get("is_receipt"))
    is_freelance_labor = bool(parsed.get("is_freelance_crew_labor"))
    hotel_name = parsed.get("hotel_name", "")
    folio_number = parsed.get("folio_number", "")
    has_person = bool(parsed.get("has_person_name"))
    has_company = bool(parsed.get("has_company_name"))
    last, first = parsed.get("last_name", ""), parsed.get("first_name", "")
    company = parsed.get("company_name", "")
    bill_to_name = parsed.get("bill_to_name", "")
    invoice_number = parsed.get("invoice_number", "")
    po_number = parsed.get("po_number", "")

    mismatch_detail = check_sender_mismatch(bill_to_name, batch)
    subfolder = FOLDER_VENDOR  # overridden below for a freelance crew labor invoice

    if is_hotel_folio:
        new_name = build_hotel_folio_filename(hotel_name, folio_number)
        if not new_name:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=FOLDER_VENDOR,
                output_pdf_bytes=b"", reason_code=REASON_MISSING_COMPANY,
                reason_detail=f"Hotel folio but missing hotel name (hotel_name={hotel_name!r})", provider=provider,
            )
    elif is_receipt:
        new_name = build_vendor_receipt_filename(company)
        if not new_name:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=FOLDER_VENDOR,
                output_pdf_bytes=b"", reason_code=REASON_MISSING_COMPANY,
                reason_detail=f"Receipt but missing company name (company={company!r})", provider=provider,
            )
    elif is_freelance_labor:
        # The invoice, in substance, represents one individual crew member's
        # own personal labor/time on this shoot (LLM judgment call based on
        # the document's actual content -- see is_freelance_crew_labor in the
        # prompt -- not a literal keyword match) -- gets its own folder
        # rather than sitting alongside ordinary vendor invoices.
        subfolder = FOLDER_FREELANCE
        if not has_person and not has_company:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=subfolder,
                output_pdf_bytes=b"", reason_code=REASON_MISSING_NAME,
                reason_detail="Freelance crew labor invoice found but could not identify a billing person or company", provider=provider,
            )
        new_name = build_freelance_filename(has_person, last, first, has_company, company, invoice_number)
        if not new_name:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=subfolder,
                output_pdf_bytes=b"", reason_code=REASON_MISSING_NAME,
                reason_detail=(f"Freelance crew labor invoice found but missing usable name/company (last={last!r}, "
                               f"first={first!r}, company={company!r})"),
                provider=provider,
            )
    elif has_company:
        # PO-based naming is a user-chosen preference (Overview intake), but
        # only usable when this specific invoice actually has a PO number on
        # it -- falls back to invoice-number naming automatically otherwise,
        # exactly as if the user had picked that convention for this file.
        new_name = None
        if batch.vendor_naming == "po_number":
            new_name = build_vendor_invoice_filename_by_po(company, po_number)
        if not new_name:
            new_name = build_vendor_invoice_filename(company, invoice_number)
        if not new_name:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=FOLDER_VENDOR,
                output_pdf_bytes=b"", reason_code=REASON_MISSING_COMPANY,
                reason_detail=f"Missing company name (company={company!r})", provider=provider,
            )
    elif has_person:
        # A vendor invoice billed directly by an individual (a freelancer, a
        # consultant, a personal-services contractor) with no company name
        # at all and not judged to be freelance crew labor -- confirmed real
        # case: a "Purchase Order / Check Request" form where the Vendor
        # field is just a person's name for things like a storyboard
        # artist's fee or a location consultation, not a business. Same
        # "Last, First - Invoice Number" naming as the is_freelance_labor
        # branch above uses for a person -- the PO-number naming choice
        # doesn't apply here (it's a company-invoice-specific alternative),
        # always by invoice number.
        new_name = build_freelance_filename(has_person, last, first, has_company, company, invoice_number)
        if not new_name:
            return DocResult(
                bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=FOLDER_VENDOR,
                output_pdf_bytes=b"", reason_code=REASON_MISSING_NAME,
                reason_detail=f"Missing usable name (last={last!r}, first={first!r})", provider=provider,
            )
    else:
        return DocResult(
            bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=FOLDER_VENDOR,
            output_pdf_bytes=b"", reason_code=REASON_MISSING_COMPANY,
            reason_detail="Could not identify a billing person or company on this invoice", provider=provider,
        )

    if mismatch_detail:
        return DocResult(
            bucket=BUCKET_UNABLE_TO_RENAME, doc_type="vendor", subfolder=subfolder,
            output_pdf_bytes=b"", reason_code=REASON_SENDER_MISMATCH,
            reason_detail=mismatch_detail, provider=provider,
        )

    return DocResult(
        bucket=BUCKET_RENAMED, doc_type="vendor", subfolder=subfolder,
        output_pdf_bytes=b"", new_filename=new_name, provider=provider,
    )


def _guess_name_from_filename(filename: str) -> str:
    base = os.path.splitext(filename)[0]
    base = re.sub(r"\([^)]*\)", " ", base)
    base = re.sub(r"_.*$", "", base)
    base = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", base)
    base = base.replace(",", " ")
    return re.sub(r"\s+", " ", base).strip()


# ── Top-level dispatcher ──────────────────────────────────────────────────────

async def process_one_document(
    raw_bytes: bytes,
    filename: str,
    batch: BatchContext,
    openai_client,
    anthropic_client,
    loop,
) -> list:
    """The single entry point the job layer calls per uploaded file. Converts
    to PDF if needed (local, no LLM cost), classifies the document family
    from a quick look at the first few pages, then dispatches to that
    family's extraction. Any unexpected exception is caught here so one bad
    file never kills a job.

    Returns a LIST of DocResult, not one -- for every family except
    Residency this is always a single-element list (one uploaded file, one
    output file), but Residency can legitimately split one uploaded file
    into several (a batch-scanned stack of different people's IDs)."""
    try:
        pdf_bytes = load_source_as_pdf_bytes(raw_bytes, filename)
    except Exception as e:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="unknown", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=raw_bytes, reason_code=REASON_UNREADABLE,
            reason_detail=f"Could not open/convert file: {e}",
        )]

    try:
        # Flatten any interactive form fields into plain static content
        # ("Print to PDF") before anything else -- some fillable forms
        # submitted here have badly-scrambled AcroForm field name/value
        # pairs internally, even though the on-page rendering is unaffected.
        # This doesn't fix anything our own vision-based reading needed
        # fixed; it just keeps that broken structure out of the output file.
        if has_form_fields(pdf_bytes):
            pdf_bytes = flatten_form_fields(pdf_bytes)
    except Exception:
        pass

    try:
        # Fix orientation next, before anything else looks at this file --
        # a sideways/upside-down scan (phone photo, misfed scanner page)
        # would otherwise confuse classification/extraction too. A failure
        # here is never fatal to the rest of the pipeline; just proceed with
        # whatever orientation the file already had.
        pdf_bytes = await correct_page_orientations(pdf_bytes, openai_client, anthropic_client, loop)
    except Exception:
        pass

    try:
        # A quick, cheap look (first few pages) is enough to tell the family
        # -- Residency's own extraction re-renders up to MAX_PAGES_TO_INSPECT
        # separately once we know that's what we're dealing with, since a
        # batch-scanned stack can run well past this classification cap.
        classify_images = render_pdf_bytes_to_images_b64(pdf_bytes, max_pages=3)
        if not classify_images:
            return [DocResult(
                bucket=BUCKET_NOT_READABLE, doc_type="unknown", subfolder=FOLDER_NOT_READABLE,
                output_pdf_bytes=pdf_bytes, reason_code=REASON_UNREADABLE,
                reason_detail="No renderable pages found",
            )]

        parsed, provider = await classify_document(
            classify_images, "Identify the broad family of this document as described in the system prompt.",
            TYPE_CLASSIFIER_PROMPT, openai_client, anthropic_client, loop,
        )
        family = str((parsed or {}).get("document_family", "unknown")).strip().lower()

        if family == "residency":
            results = await _extract_residency_multi(pdf_bytes, openai_client, anthropic_client, loop)
        elif family == "diversity":
            result = await _extract_diversity(classify_images, openai_client, anthropic_client, loop, filename)
            result.output_pdf_bytes = pdf_bytes
            results = [result]
        elif family == "vendor":
            # Re-render at a higher page cap than the classify pass -- a
            # vendor submission is often a packet (invoice + an internal PO
            # cover sheet + a W-9 + a payment confirmation), and the page
            # that actually reveals the biller's identity (e.g. a W-9's
            # "Name of entity/individual" line, or "dba" wording on a
            # payment-confirmation page) can land past page 3.
            vendor_images = render_pdf_bytes_to_images_b64(pdf_bytes, max_pages=MAX_PAGES_TO_INSPECT)
            result = await _extract_vendor(vendor_images or classify_images, openai_client, anthropic_client, loop, batch)
            result.output_pdf_bytes = pdf_bytes
            results = [result]
        elif family == "petty_cash":
            result = await _extract_petty_cash(classify_images, openai_client, anthropic_client, loop, filename)
            result.output_pdf_bytes = pdf_bytes
            results = [result]
        elif family == "prodcc":
            result = await _extract_prodcc(classify_images, openai_client, anthropic_client, loop)
            result.output_pdf_bytes = pdf_bytes
            results = [result]
        elif family == "crew_payroll":
            result = await _extract_crew_payroll(classify_images, openai_client, anthropic_client, loop)
            result.output_pdf_bytes = pdf_bytes
            results = [result]
        elif family == "misc_form":
            results = await _extract_misc_form_multi(pdf_bytes, openai_client, anthropic_client, loop)
        else:
            results = [DocResult(
                bucket=BUCKET_NOT_READABLE, doc_type="unknown", subfolder=FOLDER_NOT_READABLE,
                output_pdf_bytes=pdf_bytes, reason_code=REASON_UNCLASSIFIED,
                reason_detail="Could not confidently classify this document into a known type" if parsed is not None
                              else "Could not read document (both providers failed)",
                provider=provider,
            )]
        return results
    except Exception as e:
        return [DocResult(
            bucket=BUCKET_NOT_READABLE, doc_type="unknown", subfolder=FOLDER_NOT_READABLE,
            output_pdf_bytes=pdf_bytes, reason_code=REASON_PROVIDER_ERROR,
            reason_detail=f"Unexpected error: {e}",
        )]


# ── Naming conventions, exposed as data (GET /wizard01/conventions) ──────────

NAMING_CONVENTIONS = [
    {
        "type_id": "residency",
        "label": "Residency (Driver's License / State ID / I-9)",
        "patterns": [
            {"pattern": "{Last}, {First} ({State}) {YYYY_MM_DD}.pdf", "description": "Driver's License / State ID, named by issuing state + expiration date"},
            {"pattern": "{Last}, {First} - I9.pdf", "description": "USCIS Form I-9"},
        ],
        "notes": (
            "If one uploaded PDF contains several different people's ID documents "
            "(a batch-scanned stack), each person is split out into their own "
            "renamed output file -- one upload can produce multiple results here."
        ),
    },
    {
        "type_id": "diversity",
        "label": "Diversity Form",
        "patterns": [
            {"pattern": "{Last}, {First} - Diversity Form.pdf", "description": "Illinois DCEO Film Tax Credit Tracking Sheet"},
        ],
    },
    {
        "type_id": "vendor",
        "label": "Vendor (Invoice / Receipt / Hotel Folio)",
        "patterns": [
            {"pattern": "{Company} - Invoice {Number}.pdf", "folder": "Vendor Invoices", "description": "Company invoice, named by invoice number (number omitted if not printed) -- the default, and the automatic fallback for the PO-number option below whenever a given invoice has no PO number on it"},
            {"pattern": "PO{PO Number} - {Company}.pdf", "folder": "Vendor Invoices", "description": "Company invoice, named by PO number instead -- a user-chosen alternative (set once per batch), only used for an invoice that actually has a PO number printed on it"},
            {"pattern": "{Company} - receipt.pdf", "folder": "Vendor Invoices", "description": "Purchase receipt / order confirmation -- always by invoice number convention, not affected by the PO-number choice"},
            {"pattern": "{Hotel Name} - Folio {Number}.pdf", "folder": "Vendor Invoices", "description": "Hotel guest folio (number omitted if blank) -- not affected by the PO-number choice"},
            {"pattern": "{Last}, {First} - Invoice {Number}.pdf", "folder": "Vendor Invoices", "description": "Billed directly by an individual with no company name at all and not judged to be freelance crew labor (a freelancer/consultant invoice, e.g. a location consultation fee) -- always by invoice number, not affected by the PO-number choice"},
            {"pattern": "{Last}, {First} - Invoice {Number}.pdf", "folder": "Freelance Crew Invoice", "description": "Invoice judged (by content, not a keyword match) to represent one individual's own personal labor/service on this shoot -- e.g. a crew timesheet, a day-rate invoice with prep/shoot/wrap hours, or a Director/Photographer/Stylist fee -- always named by the person alone, even when billed through their own loan-out/DBA/LLC/studio name, gets its own folder rather than sitting alongside ordinary vendor invoices"},
        ],
        "options": [
            {
                "key": "vendor_naming",
                "label": "Vendor Invoice naming",
                "choices": [
                    {"value": "invoice_number", "label": "Vendor Name - Invoice Number (default)"},
                    {"value": "po_number", "label": "PO Number - Vendor Name"},
                ],
                "default": "invoice_number",
                "note": "Set once per batch on the Overview screen. Only changes the plain company-invoice pattern above -- Receipts, Hotel Folios, and Freelance Crew Invoices are unaffected either way.",
            },
        ],
    },
    {
        "type_id": "petty_cash",
        "label": "Petty Cash",
        "patterns": [
            {"pattern": "{Last}, {First} - Petty Cash.pdf", "folder": "Petty Cash", "description": "A named custodian's Petty Cash Summary reimbursement form, or an internal cash-advance PO handing that custodian a float (Vendor field literally \"Petty Cash\")"},
            {"pattern": "Petty Cash Log.pdf", "folder": "Petty Cash", "description": "A multi-custodian aggregate summary/spreadsheet with no single named custodian"},
            {"pattern": "aaa{Original Filename}.pdf", "folder": "Petty Cash", "description": "A bare stack of receipts with no cover sheet, PO, or summary form of any kind identifying whose batch this is -- original filename kept (it may be the only clue) and prefixed with \"aaa\" so it sorts to the top for manual review/renaming"},
        ],
    },
    {
        "type_id": "prodcc",
        "label": "ProdCC (Production Credit Card Reimbursement)",
        "patterns": [
            {"pattern": "{Last}, {First} - CC Reimb.pdf", "folder": "ProdCC", "description": "A crew member reimbursed for expenses paid with their own money/personal card -- the reverse of Petty Cash. Identified only by an explicit cover sheet built for that purpose: a \"[Name] - CC Reimbursement\" PO cover sheet, or a human-made (often Excel) cover sheet specifically about a card reimbursement -- never a \"Petty Cash Summary\" form, which is always Petty Cash regardless of its balance-due direction"},
        ],
    },
    {
        "type_id": "crew_payroll",
        "label": "Crew Payroll",
        "patterns": [
            {"pattern": "{Company} - Invoice {Number}.pdf", "folder": "Crew Payroll", "description": "A payroll company's own consolidated invoice for one batch of crew wages (Invoice Fee Summary / Fringe Recap Report / Wage-Payroll Register / Payroll Check Register pages) -- named by the payroll company's own short/common name (from its letterhead, not a possibly-different \"Employer of Record\" field) and its own batch reference number, preferring a literal \"Invoice Number\" field when present and otherwise whatever field the provider uses instead (e.g. \"Payroll ID\"), e.g. \"CAPS - Invoice 1001430894.pdf\", \"Wrapbook - Invoice 00722447.pdf\", \"The Team Companies - Invoice 26218283.pdf\", \"Revolution Entertainment Services - Invoice A23720A0136.pdf\" -- regardless of how many individual crew members' pay it covers"},
            {"pattern": "Fringe Report.pdf", "folder": "Crew Payroll", "description": "A whole-project aggregate table of every crew member's wages/fringes/benefits, not tied to one batch or invoice"},
            {"pattern": "Payroll Register.pdf", "folder": "Crew Payroll", "description": "A whole-project running payroll ledger covering every payment across the full project date range, not tied to one batch or invoice"},
            {"pattern": "Payroll Log by Line #.pdf", "folder": "Crew Payroll", "description": "A whole-project \"PAYROLL LOG\" table sorted by ascending Line number, not tied to one batch or invoice"},
            {"pattern": "Payroll Log by Payee.pdf", "folder": "Crew Payroll", "description": "Same PAYROLL LOG table, grouped by payee instead -- the same payee's name repeats across consecutive rows"},
        ],
    },
    {
        "type_id": "misc_form",
        "label": "Other Forms (G-7, Time Card, etc.)",
        "patterns": [
            {"pattern": "{Company} - {Form Type} {Date}.pdf", "folder": "Other Forms/{Form Type}", "description": "A standardized form filed/owned by a company as a whole, e.g. \"Other Forms/G-7 Form/Wrapbook - G-7 Form 2024_03_31.pdf\" for a Georgia G-7 Quarterly Return filed by the payroll company"},
            {"pattern": "{Last}, {First} - {Form Type} {Date}.pdf", "folder": "Other Forms/{Form Type}", "description": "A standardized form belonging to one specific crew member, e.g. \"Other Forms/Time Card/Marivee, Cade - Time Card 2023_12_31.pdf\""},
        ],
        "notes": (
            "Catch-all for standardized production paperwork that isn't Residency, Diversity, Vendor, "
            "Petty Cash, ProdCC, or Crew Payroll. Each distinct form_type gets its own subfolder under "
            "\"Other Forms\" (e.g. \"Other Forms/G-7 Form\", \"Other Forms/Kit Rental\", \"Other Forms/Time Card\") "
            "rather than one flat folder. If one uploaded PDF contains several distinct forms "
            "batch-scanned together, each is split out into its own renamed output file -- one upload "
            "can produce multiple results here, same as Residency."
        ),
    },
]
