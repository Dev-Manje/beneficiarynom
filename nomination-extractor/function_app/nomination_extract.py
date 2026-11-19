"""Read and validate a submitted Beneficiary Nomination Form (fillable PDF, BNF-*).

Pure logic, no Azure dependencies, so it can be tested locally:
    python nomination_extract.py "path/to/form.pdf"
"""
import io
import json
import re
import sys
from datetime import date

from pypdf import PdfReader
from pypdf.generic import IndirectObject

FORM_VERSION_PREFIX = "BNF-"
NOMINEE_ROWS = 15
PAYEE_ROWS = 3

MEMBER_FIELDS = [
    "fund_name", "member_name", "member_number", "branch_code", "cost_centre",
    "home_telephone", "cell_telephone", "work_telephone", "email_address",
    "declaration_full_names",
]


def _text(values, name):
    v = values.get(name)
    return str(v).strip() if v is not None else ""


def parse_date(text):
    """dd/mm/yyyy (also d/m/yyyy, dd-mm-yyyy, dd.mm.yyyy) -> 'yyyy-mm-dd', else None."""
    m = re.fullmatch(r"(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})", text or "")
    if not m:
        return None
    d, mo, y = map(int, m.groups())
    try:
        return date(y, mo, d).isoformat()
    except ValueError:
        return None


def parse_number(text):
    """'1,234.50' -> 1234.5; '' -> None; raises ValueError if not a number."""
    t = (text or "").replace(",", "").replace("%", "").strip()
    if not t:
        return None
    return float(t)


def _resolve(obj):
    return obj.get_object() if isinstance(obj, IndirectObject) else obj


def _has_image(xobject_owner, depth=0):
    """True if an appearance stream (or a form XObject inside it) draws an image."""
    if xobject_owner is None or depth > 4:
        return False
    res = _resolve(xobject_owner.get("/Resources"))
    xobjs = _resolve(res.get("/XObject")) if res else None
    if not xobjs:
        return False
    for ref in xobjs.values():
        x = _resolve(ref)
        if x.get("/Subtype") == "/Image":
            return True
        if x.get("/Subtype") == "/Form" and _has_image(x, depth + 1):
            return True
    return False


def signature_image_present(reader):
    """Look for an image in the member_signature button's appearance (Adobe 'attach signature')."""
    for page in reader.pages:
        for annot_ref in page.get("/Annots") or []:
            a = _resolve(annot_ref)
            if a.get("/T") == "member_signature" or _resolve(a.get("/Parent") or {}).get("/T") == "member_signature":
                ap = _resolve(a.get("/AP"))
                normal = _resolve(ap.get("/N")) if ap else None
                return bool(normal is not None and _has_image(normal))
    return False


def extract(pdf_bytes):
    """Return a dict: valid, errors, warnings, form_version, member, nominees, payees, totals."""
    errors, warnings = [], []
    result = {"valid": False, "errors": errors, "warnings": warnings, "form_version": None,
              "member": None, "nominees": [], "payees": [], "totals": None}

    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        if reader.is_encrypted:
            reader.decrypt("")
        values = reader.get_form_text_fields() or {}
    except Exception as e:  # not a PDF, corrupt, or password-protected
        errors.append(f"Could not read the PDF: {e}")
        return result

    version = _text(values, "form_version")
    result["form_version"] = version or None
    if not version.startswith(FORM_VERSION_PREFIX):
        errors.append("This is not the fillable Beneficiary Nomination Form, or its form fields were removed "
                      "(for example by printing to PDF or scanning).")
        return result

    # --- member -----------------------------------------------------------
    member = {k: _text(values, k) for k in MEMBER_FIELDS}
    dob_text = _text(values, "member_date_of_birth")
    sig_date_text = _text(values, "signature_date")
    member.update({
        "date_of_birth": parse_date(dob_text),
        "date_of_birth_text": dob_text,
        "signature_date": parse_date(sig_date_text),
        "signature_date_text": sig_date_text,
        "signature_attached": _text(values, "signature_attached").lower() == "yes",
        "signature_attached_at": _text(values, "signature_attached_at"),
        "signature_image_found": signature_image_present(reader),
    })
    result["member"] = member

    if not member["signature_attached"]:
        errors.append("The form has not been signed. Open it in Adobe Acrobat Reader and click the signature box.")
    elif not member["signature_image_found"]:
        warnings.append("Signature is marked as attached, but no signature image was found on the form.")
    if not member["member_name"]:
        errors.append("Member name is empty.")
    if not member["member_number"]:
        errors.append("Member number is empty.")
    if not member["declaration_full_names"]:
        warnings.append("Full names in the declaration are empty.")
    if dob_text and not member["date_of_birth"]:
        warnings.append(f"Member date of birth '{dob_text}' is not a valid dd/mm/yyyy date.")
    if sig_date_text and not member["signature_date"]:
        warnings.append(f"Signature date '{sig_date_text}' is not a valid dd/mm/yyyy date.")

    # --- nominees -----------------------------------------------------------
    for i in range(1, NOMINEE_ROWS + 1):
        p = f"nominee_{i:02d}_"
        row = {k: _text(values, p + k) for k in
               ("full_name", "date_of_birth", "relationship", "identity_number",
                "benefit_current_pct", "benefit_revised_pct")}
        if not any(row.values()):
            continue
        if not row["full_name"]:
            warnings.append(f"Nominee row {i} has details but no name; it was skipped.")
            continue
        n = {"row": i, "full_name": row["full_name"], "relationship": row["relationship"],
             "identity_number": row["identity_number"],
             "date_of_birth": parse_date(row["date_of_birth"]), "date_of_birth_text": row["date_of_birth"]}
        for key in ("benefit_current_pct", "benefit_revised_pct"):
            try:
                n[key] = parse_number(row[key])
            except ValueError:
                n[key] = None
                errors.append(f"Nominee row {i}: '{row[key]}' is not a valid percentage.")
        if row["date_of_birth"] and not n["date_of_birth"]:
            warnings.append(f"Nominee row {i}: date of birth '{row['date_of_birth']}' is not a valid dd/mm/yyyy date.")
        result["nominees"].append(n)

    # --- court order payees ---------------------------------------------------
    for i in range(1, PAYEE_ROWS + 1):
        p = f"payee_{i:02d}_"
        row = {k: _text(values, p + k) for k in
               ("full_name", "relationship", "identity_number", "benefit_pct", "monthly_maintenance")}
        if not any(row.values()):
            continue
        if not row["full_name"]:
            warnings.append(f"Court order row {i} has details but no name; it was skipped.")
            continue
        pay = {"row": i, "full_name": row["full_name"], "relationship": row["relationship"],
               "identity_number": row["identity_number"]}
        for key in ("benefit_pct", "monthly_maintenance"):
            try:
                pay[key] = parse_number(row[key])
            except ValueError:
                pay[key] = None
                errors.append(f"Court order row {i}: '{row[key]}' is not a valid number.")
        result["payees"].append(pay)

    # --- totals ---------------------------------------------------------------
    noms = result["nominees"]
    tot_cur = round(sum(n["benefit_current_pct"] or 0 for n in noms), 2)
    tot_rev = round(sum(n["benefit_revised_pct"] or 0 for n in noms), 2)
    uses_revised = any(n["benefit_revised_pct"] is not None for n in noms)
    result["totals"] = {"nominee_count": len(noms), "payee_count": len(result["payees"]),
                        "benefit_current_pct": tot_cur, "benefit_revised_pct": tot_rev,
                        "effective_column": "revised" if uses_revised else "current"}
    if not noms:
        errors.append("No dependants / nominees are listed.")
    else:
        effective = tot_rev if uses_revised else tot_cur
        if abs(effective - 100) > 0.01:
            col = "Revised" if uses_revised else "Current"
            errors.append(f"% Benefit ({col}) adds up to {effective:g}%, it must equal 100%.")

    result["valid"] = not errors
    return result


if __name__ == "__main__":
    with open(sys.argv[1], "rb") as f:
        print(json.dumps(extract(f.read()), indent=2))
