import csv
import datetime as dt
import io
import re
from typing import Any, Dict, List

import openpyxl
import pdfplumber

ALLOWED_STATUSES = [
    "Missing:critical",
    "Expired:critical",
    "Valid:good",
    "Flagged:warning",
    "Pending_review:info",
    "Approved:good",
    "Rejected:critical",
    "On_hold:warning",
    "Overbilled:critical",
    "Underbilled:warning",
    "SOV_drift:warning",
    "Unapproved_change_order_billed:critical",
    "Missing_retainage:critical",
    "Retainage_rate_mismatch:critical",
    "Retainage_step_down_due:warning",
    "Retainage_step_down_not_applied:warning",
    "Prior_approval_mismatch:warning",
    "Period_continuity_break:critical",
    "Math_error:critical",
    "Duplicate_billing:critical",
    "Stored_materials_unverified:warning",
    "Missing_waiver:critical",
    "Mismatched_waiver_amount:critical",
    "Wrong_statutory_waiver_form:critical",
    "Missing_signature:warning",
    "Missing_notary:warning",
    "Tier_unknown:warning",
    "Tier_discovered:info",
    "COI_expired:critical",
    "W-9_missing:critical",
    "License_expired:critical",
    "Additional_insured_missing:warning",
    "Negative_change_order_auto_approved:warning",
    "Balance_to_finish_negative:critical",
    "Current_payment_due_mismatch:critical",
    "Contract_sum_mismatch:critical",
]

FIELD_ALIASES = {
    "project_name": ["project", "project name", "project_name", "job name", "job_name"],
    "project_number": ["project number", "project_number", "project_no", "job number", "job_number"],
    "project_address": ["project address", "project_address", "project_address", "site address"],
    "owner_name": ["owner", "owner_name", "owner name"],
    "gc_name": ["gc", "gc_name", "general_contractor", "general contractor"],
    "architect_name": ["architect", "architect_name", "architect name"],
    "contract_date": ["contract_date", "contract date"],
    "billing_period_start": ["billing_period_start", "billing period start", "period start"],
    "billing_period_end": ["billing_period_end", "billing period end", "period end"],
    "application_number": ["application_number", "application number", "app number", "app_number", "pay_app_number"],
    "application_date": ["application_date", "application date"],
    "contract_sum": ["contract_sum", "contract sum", "original contract sum", "original_contract_sum"],
    "retainage_rate": ["retainage_rate", "retainage rate", "retention rate", "retention_rate"],
    "retainage_terms": ["retainage_terms", "retainage terms", "retention terms"],
    "retainage_step_down_milestones": ["retainage_step_down_milestones", "retainage step-down milestones", "retainage step down milestones"],
    "billing_cutoff": ["billing_cutoff", "billing cutoff"],
    "sov_line_item_number": ["sov_line_item_number", "sov line item number", "line item number", "sov_line_item_no"],
    "sov_description": ["sov_description", "sov description", "description", "line item description"],
    "scheduled_value": ["scheduled_value", "scheduled value"],
    "unit": ["unit"],
    "quantity": ["quantity", "qty"],
    "unit_price": ["unit_price", "unit price", "price"],
    "scheduled_quantity": ["scheduled_quantity", "scheduled quantity"],
    "approved_co_number": ["approved_co_number", "approved change order number", "approved co number"],
    "approved_co_amount": ["approved_co_amount", "approved change order amount", "approved co amount"],
    "unapproved_co_amount": ["unapproved_co_amount", "unapproved change order amount", "unapproved co amount"],
    "line_item_retainage_rate": ["line_item_retainage_rate", "line-item retainage rate", "line item retainage rate"],
    "total_completed_and_stored_to_date": ["total_completed_and_stored_to_date", "total completed and stored to date"],
    "total_retainage": ["total_retainage", "total retainage"],
    "total_earned_less_retainage": ["total_earned_less_retainage", "total earned less retainage"],
    "previous_certificates_for_payment": ["previous_certificates_for_payment", "previous certificates for payment"],
    "current_payment_due": ["current_payment_due", "current payment due"],
    "balance_to_finish_including_retainage": ["balance_to_finish_including_retainage", "balance to finish including retainage"],
    "change_order_summary": ["change_order_summary", "change order summary"],
    "contractor_signature": ["contractor_signature", "contractor signature"],
    "architect_signature": ["architect_signature", "architect signature"],
    "owner_signature": ["owner_signature", "owner signature"],
    "notary_block": ["notary_block", "notary block", "notary"],
    "previous_completed": ["previous_completed", "previous completed"],
    "this_period_completed": ["this_period_completed", "this period completed"],
    "materials_stored": ["materials_stored", "materials stored"],
    "percent_complete": ["percent_complete", "percent complete"],
    "balance_to_finish": ["balance_to_finish", "balance to finish"],
    "retainage_on_completed": ["retainage_on_completed", "retainage on completed"],
    "retainage_on_stored": ["retainage_on_stored", "retainage on stored"],
    "total_retainage_per_line": ["total_retainage_per_line", "total retainage per line"],
    "prior_application_number": ["prior_application_number", "prior application number"],
    "prior_approved_amount": ["prior_approved_amount", "prior approved amount"],
    "prior_paid_amount": ["prior_paid_amount", "prior paid amount"],
    "prior_retainage_held": ["prior_retainage_held", "prior retainage held"],
    "prior_retainage_released": ["prior_retainage_released", "prior retainage released"],
    "prior_percent_complete_by_line": ["prior_percent_complete_by_line", "prior percent complete by line"],
    "prior_approval_date": ["prior_approval_date", "prior approval date"],
    "prior_approver_name": ["prior_approver_name", "prior approver name"],
    "co_number": ["co_number", "co number", "change order number"],
    "co_status": ["co_status", "co status", "change order status"],
    "co_amount": ["co_amount", "co amount", "change order amount"],
    "co_description": ["co_description", "co description", "change order description"],
    "co_date": ["co_date", "co date", "change order date"],
    "waiver_type": ["waiver_type", "waiver type"],
    "waiver_stage": ["waiver_stage", "waiver stage"],
    "statutory_form_state": ["statutory_form_state", "statutory form state"],
    "claimant_name": ["claimant_name", "claimant name", "claimant"],
    "claimant_tier": ["claimant_tier", "claimant tier"],
    "amount_waived": ["amount_waived", "amount waived"],
    "through_date": ["through_date", "through date"],
    "conditional_or_unconditional": ["conditional_or_unconditional", "conditional or unconditional"],
    "progress_or_final": ["progress_or_final", "progress or final"],
    "e_signature_status": ["e_signature_status", "e-signature status", "esignature status"],
    "coi_carrier": ["coi_carrier", "coi carrier", "insurance carrier", "carrier"],
    "policy_type": ["policy_type", "policy type"],
    "policy_number": ["policy_number", "policy number"],
    "limits": ["limits", "policy limits"],
    "effective_date": ["effective_date", "effective date"],
    "expiration_date": ["expiration_date", "expiration date"],
    "additional_insured_wording": ["additional_insured_wording", "additional insured wording"],
    "w9_legal_name": ["w9_legal_name", "w-9 legal name", "w9 legal name", "legal name"],
    "tin": ["tin", "tax id", "taxpayer identification number"],
    "address": ["address", "business address"],
    "license_number": ["license_number", "license number", "license no"],
    "license_type": ["license_type", "license type"],
    "license_expiration": ["license_expiration", "license expiration", "license expiration date"],
    "entity_status": ["entity_status", "entity status"],
    "ach_or_payment_status": ["ach_or_payment_status", "ach or payment status", "payment status"],
    "payment_date": ["payment_date", "payment date"],
    "payment_amount": ["payment_amount", "payment amount"],
    "payee": ["payee", "payee name"],
    "payment_reference": ["payment_reference", "payment reference"],
    "waiver_exchange_status": ["waiver_exchange_status", "waiver exchange status"],
    "release_of_retainage_status": ["release_of_retainage_status", "release of retainage status"],
    "upload_timestamp": ["upload_timestamp", "upload timestamp"],
    "page_reference": ["page_reference", "page reference"],
    "extraction_confidence": ["extraction_confidence", "extraction confidence"],
    "reviewer_name": ["reviewer_name", "reviewer name"],
    "review_date": ["review_date", "review date"],
    "flag_history": ["flag_history", "flag history"],
    "resolution_note": ["resolution_note", "resolution note"],
    "approval_or_rejection_reason": ["approval_or_rejection_reason", "approval or rejection reason"],
    "supplier": ["supplier", "vendor", "vendor_name", "subcontractor", "sub", "sub_name", "contractor", "company"],
    "product": ["product", "item", "sov_description", "description"],
    "price": ["price", "unit_price", "amount", "payment_amount", "scheduled_value"],
    "due_date": ["due_date", "due date"],
}


def _normalize_key(value: str) -> str:
    value = value.strip().lower()
    value = re.sub(r"[^a-z0-9 _-]+", " ", value)
    value = re.sub(r"[\s-]+", "_", value)
    return value.strip("_")


def _canonical_key(raw_header: str) -> str:
    normalized = _normalize_key(raw_header)
    for canonical, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            if _normalize_key(alias) == normalized:
                return canonical
    return normalized


def _extract_text(file_bytes: bytes) -> str:
    # PDF first.
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            page_texts = []
            for page in pdf.pages:
                page_text = page.extract_text() or ""
                page_texts.append(page_text)
            combined = "\n".join(page_texts).strip()
            if combined:
                return combined
    except Exception:
        pass

    # Excel second.
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        lines = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                cells = ["" if cell is None else str(cell) for cell in row]
                if any(cells):
                    lines.append(",".join(cells))
        combined = "\n".join(lines).strip()
        if combined:
            return combined
    except Exception:
        pass

    # Plain text or CSV fallback.
    return file_bytes.decode("utf-8-sig", errors="ignore")


def _parse_tabular_rows(text: str) -> List[List[str]]:
    text = text.strip()
    if not text:
        return []

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return []

    candidates = [",", "\t", ";"]
    counts = {candidate: sum(line.count(candidate) for line in lines) for candidate in candidates}
    best_delim = None
    best_count = 0
    for candidate, count in counts.items():
        if count > best_count:
            best_delim = candidate
            best_count = count

    if best_delim is None or best_count == 0:
        rows = []
        for line in lines:
            if ":" in line:
                key, _, value = line.partition(":")
                rows.append([key.strip(), value.strip()])
            else:
                rows.append([line.strip()])
        return rows

    try:
        rows = list(csv.reader(io.StringIO(text), delimiter=best_delim))
    except Exception:
        rows = [[cell.strip() for cell in line.split(best_delim)] for line in lines]

    cleaned = []
    for row in rows:
        cells = [cell.strip() for cell in row]
        if any(cells):
            cleaned.append(cells)
    return cleaned


def _parse_date_str(value: str) -> str:
    if value is None:
        return ""
    value = str(value).strip()
    formats = [
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%B %d, %Y",
        "%b %d, %Y",
        "%d-%b-%Y",
        "%d-%b-%y",
    ]
    for fmt in formats:
        try:
            return dt.datetime.strptime(value, fmt).date().isoformat()
        except ValueError:
            continue

    match = re.search(r"\b(\d{4})-(\d{2})-(\d{2})\b", value)
    if match:
        try:
            return dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3))).isoformat()
        except ValueError:
            return ""
    match = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", value)
    if match:
        try:
            return dt.date(int(match.group(3)), int(match.group(1)), int(match.group(2))).isoformat()
        except ValueError:
            return ""
    return ""


def _extract_due_date(details: Dict[str, Any], full_text: str) -> str:
    preferred_keys = ["due_date", "expiration_date", "billing_period_end", "through_date", "application_date", "contract_date"]
    for key in preferred_keys:
        candidate = details.get(key)
        if candidate:
            parsed = _parse_date_str(candidate)
            if parsed:
                return parsed

    patterns = [
        r"\b\d{4}-\d{2}-\d{2}\b",
        r"\b\d{1,2}/\d{1,2}/\d{4}\b",
        r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]* \d{1,2},? \d{4}\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, full_text, re.IGNORECASE)
        if match:
            parsed = _parse_date_str(match.group(0))
            if parsed:
                return parsed
    return ""


def _extract_title(details: Dict[str, Any], row: List[str]) -> str:
    priority_keys = [
        "supplier",
        "vendor_name",
        "subcontractor",
        "contractor",
        "claimant_name",
        "claimant",
        "payee",
        "company",
        "legal_name",
        "w9_legal_name",
        "sub_name",
        "owner_name",
        "gc_name",
    ]
    for key in priority_keys:
        candidate = details.get(key)
        if candidate:
            value = str(candidate).strip()
            if value:
                return value

    for cell in row:
        value = str(cell).strip()
        if value:
            return value
    return "Unnamed Entity"


def _derive_status(details: Dict[str, Any], full_text: str) -> str:
    combined_values = " ".join(str(value) for value in details.values())
    combined = (combined_values + " " + full_text).lower()

    checks = [
        (r"\bunapproved change order\b", "Unapproved_change_order_billed:critical"),
        (r"\bcoi expired\b|\blicense expired\b", "Expired:critical"),
        (r"\bmissing waiver\b", "Missing_waiver:critical"),
        (r"\bmismatched waiver amount\b", "Mismatched_waiver_amount:critical"),
        (r"\bnegative change order auto-approved\b", "Negative_change_order_auto_approved:warning"),
        (r"\bbalance to finish negative\b", "Balance_to_finish_negative:critical"),
        (r"\bretainage rate mismatch\b", "Retainage_rate_mismatch:critical"),
        (r"\bsov drift\b", "SOV_drift:warning"),
        (r"\bmissing retainage\b|\bmissing retention\b", "Missing_retainage:critical"),
    ]
    for pattern, status in checks:
        if re.search(pattern, combined, re.IGNORECASE):
            return status
    return "Pending_review:info"


def _row_to_record(headers: List[str], row: List[str], row_index: int, full_text: str) -> Dict[str, Any]:
    details: Dict[str, Any] = {}

    if headers:
        for idx, cell in enumerate(row):
            raw_value = str(cell).strip()
            if raw_value == "":
                continue

            if idx >= len(headers):
                details[f"column_{idx}"] = raw_value
                continue

            raw_key = headers[idx].strip()
            if not raw_key:
                details[f"column_{idx}"] = raw_value
                continue

            canonical = _canonical_key(raw_key)
            normalized = _normalize_key(raw_key)
            details[canonical] = raw_value
            if normalized != canonical:
                details[normalized] = raw_value
    else:
        for idx, cell in enumerate(row):
            raw_value = str(cell).strip()
            if raw_value:
                details[f"column_{idx}"] = raw_value

    title = _extract_title(details, row)
    status = _derive_status(details, full_text)
    due_date = _extract_due_date(details, full_text)
    details.pop("due_date", None)

    return {
        "title": title,
        "status": status,
        "details": details,
        "due_date": due_date,
    }


def _text_to_record(text: str) -> Dict[str, Any]:
    details: Dict[str, Any] = {}
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        if ":" in line:
            key, _, value = line.partition(":")
            canonical = _canonical_key(key)
            value = value.strip()
            if value:
                details[canonical] = value
            continue

        if "," in line:
            parts = [part.strip() for part in line.split(",")]
            for idx, part in enumerate(parts):
                if part:
                    details[f"column_{idx}"] = part

    available_values = [value for value in details.values() if str(value).strip()]
    title = _extract_title(details, [str(value) for value in available_values])
    due_date = _extract_due_date(details, text)
    details.pop("due_date", None)

    return {
        "title": title,
        "status": _derive_status(details, text),
        "details": details,
        "due_date": due_date,
    }


def process_file(file_bytes: bytes) -> List[Dict[str, Any]]:
    if not file_bytes:
        return [
            {
                "title": "Unnamed Entity",
                "status": "Missing:critical",
                "details": {},
                "due_date": "",
            }
        ]

    text = _extract_text(file_bytes)
    if not text.strip():
        return [
            {
                "title": "Unnamed Entity",
                "status": "Missing:critical",
                "details": {},
                "due_date": "",
            }
        ]

    rows = _parse_tabular_rows(text)

    if len(rows) <= 1:
        record = _text_to_record(text)
        return [record]

    headers = rows[0]
    data_rows = rows[1:]
    records = []
    for idx, row in enumerate(data_rows):
        if any(str(cell).strip() for cell in row):
            records.append(_row_to_record(headers, row, idx, text))

    if not records:
        records.append(_text_to_record(text))

    return records
