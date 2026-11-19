"""Turn the Beneficiary Nomination Form template into a fillable PDF (AcroForm),
and write a field list (JSON + CSV) for mapping in Power Automate.

usage: python make_fillable.py <template.pdf> <out_dir>
"""
import csv
import json
import os
import sys
import fitz

SRC, OUT = sys.argv[1], sys.argv[2]
BASE = "Beneficiary Nomination Form"
FORM_VERSION = "BNF-6.06-F1.0"  # template version 6.06, fillable layout 1.0

doc = fitz.open(SRC)
page = doc[0]

# Remove the pre-printed "/ /" in the Date of Birth box so typed dates don't overlap it.
page.add_redact_annot(fitz.Rect(403, 123, 421, 134), fill=False)
# Drop the "Registration No." caption under the fund name (replaced with "Scheme Name" below)
# and the "Com_Pen @ Work Version 6.06" footer.
page.add_redact_annot(fitz.Rect(268, 46, 330, 57), fill=False)
page.add_redact_annot(fitz.Rect(478, 821, 575, 831), fill=False)
page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)

# "Scheme Name", centred where "Registration No." was (same size and baseline).
CAPTION, CAPTION_SIZE = "Scheme Name", 8.04
page.insert_text((298.9 - fitz.get_text_length(CAPTION, "helv", CAPTION_SIZE) / 2, 53.88),
                 CAPTION, fontname="helv", fontsize=CAPTION_SIZE)

DATE_FMT = 'AFDate_FormatEx("dd/mm/yyyy");'
DATE_KEY = 'AFDate_KeystrokeEx("dd/mm/yyyy");'
NUM_FMT = 'AFNumber_Format(2, 0, 0, 0, "", true);'
NUM_KEY = 'AFNumber_Keystroke(2, 0, 0, 0, "", true);'
PCT_VAL = 'AFRange_Validate(true, 0, true, 100);'

FORMATS = {"text": "", "date": "dd/mm/yyyy", "pct": "number 0-100", "amt": "number, 2 decimals",
           "sig": "image button (signature image; Adobe Reader button or signer page)", "hidden": "hidden, set automatically"}
SPEC = []

# Adobe Reader: pick an image, show it as the signature, and record that it was attached.
SIG_JS = r"""
var f = event.target;
if (f.buttonImportIcon() === 0) {
  f.buttonPosition = position.iconOnly;
  f.fillColor = color.transparent;
  f.borderColor = color.transparent;
  f.display = display.visible;
  var now = new Date();
  this.getField("signature_attached").value = "yes";
  this.getField("signature_attached_at").value = util.printd("yyyy-mm-dd", now) + "T" + util.printd("HH:MM:ss", now);
  var d = this.getField("signature_date");
  if (!d.valueAsString) d.value = util.printd("dd/mm/yyyy", now);
}
"""


def field(name, rect, tip, section, row="", size=9, align=0, kind="text", value=""):
    w = fitz.Widget()
    w.field_name = name
    w.field_label = tip + (" (dd/mm/yyyy)" if kind == "date" else "")  # tooltip
    w.rect = fitz.Rect(rect)
    w.text_font = "Helv"
    w.text_fontsize = size
    w.text_color = (0, 0, 0.45)
    w.border_width = 0
    w.fill_color = None
    if kind == "sig":
        w.field_type = fitz.PDF_WIDGET_TYPE_BUTTON
        w.button_caption = "Click here to attach your signature"
        w.text_fontsize = 8
        w.text_color = (0.12, 0.37, 0.75)
        w.fill_color = (0.91, 0.95, 1)
        w.script = SIG_JS  # runs on click (mouse-up)
    else:
        w.field_type = fitz.PDF_WIDGET_TYPE_TEXT
        if kind == "date":
            w.script_format, w.script = DATE_FMT, DATE_KEY
        elif kind == "pct":
            w.script_format, w.script, w.script_change = NUM_FMT, NUM_KEY, PCT_VAL
        elif kind == "amt":
            w.script_format, w.script = NUM_FMT, NUM_KEY
        elif kind == "hidden":
            w.field_value = value
            w.field_flags = fitz.PDF_FIELD_IS_READ_ONLY
    added = page.add_widget(w)
    if kind == "sig":
        doc.xref_set_key(added.xref, "F", "0")  # shown on screen, not printed, until a signature is attached
        doc.xref_set_key(added.xref, "MK/TP", "0")  # caption only until an image is imported
        # scale image proportionally, centred horizontally, sitting on the signature line
        doc.xref_set_key(added.xref, "MK/IF", "<</SW /A /S /P /A [0.5 0] /FB true>>")
    if kind == "hidden":
        doc.xref_set_key(added.xref, "F", "2")  # annotation flag: hidden (not shown, not printed)
    if align:
        doc.xref_set_key(added.xref, "Q", str(align))  # 1 = centre, 2 = right
    SPEC.append({
        "name": name,
        "type": "signature" if kind == "sig" else "text",
        "format": FORMATS[kind],
        "section": section,
        "row": row,
        "description": tip,
        "page": 1,
        "rect_pt": [round(v, 1) for v in rect],
    })


def cell(x0, x1, y0, y1, pad=2.5):
    return (x0 + pad, y0 + 1, x1 - pad, y1 - 1)


# --- Form version (hidden) -------------------------------------------------
field("form_version", (2, 2, 60, 12), f"Form version identifier ({FORM_VERSION})", "meta", kind="hidden", value=FORM_VERSION)
# Set when a signature image is attached (Adobe Reader button or "Beneficiary Form Signer.html").
field("signature_attached", (62, 2, 120, 12), "'yes' once a signature image has been attached", "meta", kind="hidden")
field("signature_attached_at", (122, 2, 180, 12), "When the signature was attached (yyyy-mm-ddTHH:MM:ss)", "meta", kind="hidden")

# --- Header ---------------------------------------------------------------
field("fund_name", (110, 28, 486, 46),
      "Fund registration number and name, e.g. 69707 TELEKOM NETWORKS PENSION FUND", "header", size=13, align=1)

# --- Member details (label column 28-127, value columns 127-297 / 397-567) ---
L, R = (127.4, 297.5), (396.8, 566.8)
rows = [(102.0, 118.9), (119.0, 135.8), (136.0, 152.9), (153.0, 169.9), (170.0, 187.0)]
field("member_name", cell(127.4, 566.8, *rows[0]), "Member name", "member", size=10)
field("member_number", cell(*L, *rows[1]), "Member number", "member", size=10)
field("member_date_of_birth", cell(*R, *rows[1]), "Member date of birth", "member", size=10, kind="date")
field("branch_code", cell(*L, *rows[2]), "Branch code", "member", size=10)
field("cost_centre", cell(*R, *rows[2]), "Cost centre", "member", size=10)
field("home_telephone", cell(*L, *rows[3]), "Home telephone", "member", size=10)
field("cell_telephone", cell(*R, *rows[3]), "Cell telephone", "member", size=10)
field("work_telephone", cell(*L, *rows[4]), "Work telephone", "member", size=10)
field("email_address", cell(*R, *rows[4]), "Email address", "member", size=0)  # 0 = auto-shrink to fit

# --- Declaration ------------------------------------------------------------
field("declaration_full_names", (95, 200.5, 346, 212), "Member's full names (declaration)", "declaration")
field("member_signature", (113.3, 250, 269.2, 282.5), "Member's signature: click to attach an image of your signature", "declaration", kind="sig")
field("signature_date", (325.9, 268, 453.4, 282.5), "Date signed", "declaration", size=10, align=1, kind="date")

# --- Dependants / nominees: 15 rows ----------------------------------------
cols = [(57.0, 254.9), (255.0, 311.6), (311.8, 396.7), (396.8, 481.7), (481.8, 524.3), (524.4, 566.8)]
row_tops = [345.7, 364.0, 382.1, 400.2, 418.3, 436.4, 454.6, 472.8, 490.9, 509.0,
            527.1, 545.2, 563.4, 581.5, 599.7, 617.8]
for i in range(15):
    y0, y1 = row_tops[i], row_tops[i + 1] - 0.1
    n, p = i + 1, f"nominee_{i + 1:02d}"
    field(f"{p}_full_name", cell(*cols[0], y0, y1), f"Nominee {n}: full name of dependant / nominee", "nominees", n)
    field(f"{p}_date_of_birth", cell(*cols[1], y0, y1, 1.5), f"Nominee {n}: date of birth", "nominees", n, size=8, align=1, kind="date")
    field(f"{p}_relationship", cell(*cols[2], y0, y1), f"Nominee {n}: relationship", "nominees", n)
    field(f"{p}_identity_number", cell(*cols[3], y0, y1), f"Nominee {n}: identity number", "nominees", n)
    field(f"{p}_benefit_current_pct", cell(*cols[4], y0, y1), f"Nominee {n}: % benefit (current)", "nominees", n, align=1, kind="pct")
    field(f"{p}_benefit_revised_pct", cell(*cols[5], y0, y1), f"Nominee {n}: % benefit (revised)", "nominees", n, align=1, kind="pct")

# --- Divorce / maintenance court orders: 3 rows ------------------------------
dcols = [(28.3, 269.2), (269.3, 354.1), (354.2, 425.0), (425.2, 481.7), (481.8, 566.8)]
drows = [(734.2, 751.0), (751.1, 768.0), (768.1, 785.0)]
for i, (y0, y1) in enumerate(drows):
    n, p = i + 1, f"payee_{i + 1:02d}"
    field(f"{p}_full_name", cell(*dcols[0], y0, y1), f"Court order payee {n}: full name", "court_orders", n)
    field(f"{p}_relationship", cell(*dcols[1], y0, y1), f"Court order payee {n}: relationship", "court_orders", n)
    field(f"{p}_identity_number", cell(*dcols[2], y0, y1), f"Court order payee {n}: identity number", "court_orders", n)
    field(f"{p}_benefit_pct", cell(*dcols[3], y0, y1), f"Court order payee {n}: percentage of benefit", "court_orders", n, align=1, kind="pct")
    field(f"{p}_monthly_maintenance", cell(*dcols[4], y0, y1), f"Court order payee {n}: monthly maintenance amount", "court_orders", n, align=2, kind="amt")

# --- On-screen notice (all viewers; never printed, not a form field) -----------
notice = page.add_freetext_annot(
    fitz.Rect(28.3, 6, 566.8, 26),
    "To sign this form, open it in Adobe Acrobat Reader (free) on a computer and click the blue "
    "signature box. Browsers and phone apps cannot attach the signature.",
    fontsize=8, fontname="Helv", text_color=(0.6, 0.25, 0), fill_color=(1, 0.96, 0.88), align=fitz.TEXT_ALIGN_CENTER,
)
notice.set_info(title="Signing instructions")
notice.set_border(width=0)
notice.update()
# no Print flag -> on screen only; ReadOnly + Locked -> can't be edited, moved or deleted
notice.set_flags(fitz.PDF_ANNOT_IS_READ_ONLY | fitz.PDF_ANNOT_IS_LOCKED)

doc.set_metadata({**doc.metadata, "title": "Beneficiary Nomination Form (fillable)"})
pdf_path = os.path.join(OUT, f"{BASE} - Fillable.pdf")
# Embed only the glyphs the page uses (~444 KB -> ~117 KB): keeps submissions and their data URIs small.
doc.subset_fonts()
doc.save(pdf_path, garbage=4, deflate=True, use_objstms=1)

with open(os.path.join(OUT, f"{BASE} - Fields.json"), "w", encoding="utf-8") as f:
    json.dump({"form_version": FORM_VERSION, "field_count": len(SPEC), "fields": SPEC}, f, indent=2)
with open(os.path.join(OUT, f"{BASE} - Fields.csv"), "w", encoding="utf-8", newline="") as f:
    wr = csv.DictWriter(f, fieldnames=["name", "type", "format", "section", "row", "description"], extrasaction="ignore")
    wr.writeheader()
    wr.writerows(SPEC)

print("fields:", sum(1 for _ in fitz.open(pdf_path)[0].widgets()), "spec:", len(SPEC))
