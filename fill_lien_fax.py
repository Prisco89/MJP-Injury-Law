#!/usr/bin/env python3
"""
fill_lien_fax.py — Fill the NYC DSS/HRA "Updated / Final Lien Request Fax Form"
(W-588AA (E) Rev. 04/12/2021) from a JSON spec.

Usage:
    python fill_lien_fax.py <spec.json> <output.pdf> [blank_form.pdf]

Self-contained: the blank form (text, rules, checkboxes and the NYC DSS logo)
is rebuilt by this script from embedded layout data, so no PDF file is needed.
If a path to the official blank PDF is given as the optional third argument and
the file exists, the overlay is merged onto that instead. Coordinates match the
04/12/2021 revision of the form (letter size, 612 x 792).

If a Section III / IV list has more than two entries, or an entry will not fit
on its line even at the minimum font size, the line reads "See attached rider"
and a rider page listing everything in full is appended after the form.
"""
import json
import sys
from io import BytesIO

from pypdf import PdfReader, PdfWriter
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

PAGE_W, PAGE_H = letter
FONT = "Helvetica"
BASE_SIZE = 10
MIN_SIZE = 7.5

# Firm defaults — overridden by anything in spec["attorney"].
FIRM_DEFAULTS = {
    "represents": "plaintiff",
    "firm_name": "The Law Office of Michael James Prisco PLLC",
    "firm_address": "187 Veterans Blvd, Massapequa, NY 11758",
    "attorney_name": "Michael Prisco, Esq.",
    "telephone": "(718) 709-9678",
    "email": "mp@mjpesq.com",
    "fax": "(718) 744-2033",
    "completed_by": "Michael Prisco, Esq.",
}

# Entry boxes in top-origin PDF points: (x0, x1, row_top, row_bottom).
# Each row's label text sits at row_top..row_bottom; text is drawn on the
# underline just below the label baseline.
BOXES = {
    "date_top":          (433, 555, 189, 201),
    "plaintiff_name":    (163, 500, 214, 226),
    "ssn":               (114, 330, 237, 249),
    "date_of_birth":     (456, 555, 237, 249),
    "settlement_amount": (186, 330, 260, 272),
    "date_of_incident":  (456, 555, 260, 272),
    "nyc_file_no":       (244, 330, 284, 296),
    "settlement_date":   (456, 555, 284, 296),
    "index_number":      (161, 330, 307, 319),
    "case_or_cin":       (456, 555, 307, 319),
    "injury":            (411, 555, 330, 342),
    "injury_line2":      (84, 555, 350, 362),
    "firm_name":         (147, 555, 424, 436),
    "firm_address":      (157, 555, 447, 459),
    "attorney_name":     (166, 375, 470, 482),
    "telephone":         (442, 555, 470, 482),
    "email":             (123, 375, 493, 505),
    "fax":               (442, 555, 493, 505),
    "conference_date":   (173, 330, 516, 528),
    "def_1":             (100, 555, 605, 617),
    "def_2":             (100, 555, 628, 640),
    "ins_1":             (100, 555, 690, 702),
    "ins_2":             (100, 555, 713, 725),
    "completed_by":      (160, 390, 740, 752),
    "date_bottom":       (433, 555, 740, 752),
}

# Checkbox squares (top-origin): (x0, top, x1, bottom)
CHECKBOXES = {
    "type_updated":        (224.3, 375.5, 235.9, 387.1),
    "type_final":          (323.3, 375.5, 334.8, 387.1),
    "represents_plaintiff": (296.3, 400.4, 307.8, 411.9),
    "represents_defendant": (413.4, 400.4, 424.9, 411.9),
}


# ---- Embedded blank-form layout (W-588AA (E) Rev. 04/12/2021), zlib+base64 JSON:
# {"w": [[x0, bbox_bottom, font, size, text], ...], "r": [["fill"|"box", x0, top, x1, bottom], ...]}
FORM_LAYOUT_Z = "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"
# NYC DSS logo, 1-bit PNG, placed at x=55.8, top=36.8, w=346.5, h=43.5
LOGO_PNG_B64 = "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"


def make_blank_form():
    """Rebuild the blank W-588AA form as a one-page PDF in memory."""
    import base64
    import zlib
    from reportlab.lib.utils import ImageReader

    layout = json.loads(zlib.decompress(base64.b64decode(FORM_LAYOUT_Z)))
    FORM_WORDS, FORM_RECTS = layout["w"], layout["r"]

    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    logo = ImageReader(BytesIO(base64.b64decode(LOGO_PNG_B64)))
    c.drawImage(logo, 55.8, PAGE_H - 80.3, width=346.5, height=43.5, mask=None)
    c.setFillGray(0)
    c.setStrokeGray(0)
    for kind, x0, top, x1, bottom in FORM_RECTS:
        if kind == "fill":
            c.rect(x0, PAGE_H - bottom, x1 - x0, bottom - top, stroke=0, fill=1)
        else:
            c.setLineWidth(0.75)
            c.rect(x0, PAGE_H - bottom, x1 - x0, bottom - top, stroke=1, fill=0)
    for x0, bottom, font, size, text in FORM_WORDS:
        c.setFont(font, size)
        c.drawString(x0, PAGE_H - (bottom - 0.216 * size), text)
    c.save()
    buf.seek(0)
    return buf


def fit_size(text, width, base=BASE_SIZE, minimum=MIN_SIZE):
    """Largest font size (<= base) at which text fits width; None if it never fits."""
    size = base
    while size >= minimum:
        if stringWidth(text, FONT, size) <= width:
            return size
        size -= 0.5
    return None


def draw_text(c, key, text, overflow):
    """Draw text in the named box; on overflow, write the rider marker."""
    if not text:
        return
    x0, x1, top, bottom = BOXES[key]
    width = x1 - x0 - 4
    size = fit_size(text, width)
    if size is None:
        overflow.append((key, text))
        text = "See attached rider"
        size = BASE_SIZE
    c.setFont(FONT, size)
    baseline = PAGE_H - (bottom - 1.5)
    c.drawString(x0 + 2, baseline, text)


def draw_check(c, key):
    x0, top, x1, bottom = CHECKBOXES[key]
    c.setFont("Helvetica-Bold", 10)
    c.drawCentredString((x0 + x1) / 2, PAGE_H - bottom + 2.2, "X")


def join_parts(*parts):
    return ", ".join(p.strip() for p in parts if p and p.strip())


def defendant_line(d):
    """One line: Defendant — Attorney, address, phone."""
    atty = join_parts(d.get("attorney"), d.get("attorney_address"), d.get("attorney_phone"))
    name = d.get("name", "").strip()
    return f"{name} \u2014 {atty}" if atty else name


def insurer_line(i):
    parts = [i.get("insurer", "").strip(), i.get("address", "").strip()]
    if i.get("claim_no"):
        parts.append(f"Claim/File # {i['claim_no'].strip()}")
    if i.get("insured"):
        parts.append(f"(insuring {i['insured'].strip()})")
    return ", ".join(p for p in parts if p)


def build_overlay(spec):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    overflow = []
    rider_lists = {}

    date = spec.get("date", "")
    draw_text(c, "date_top", date, overflow)
    draw_text(c, "date_bottom", date, overflow)

    p = spec.get("plaintiff", {})
    draw_text(c, "plaintiff_name", p.get("name", ""), overflow)
    draw_text(c, "ssn", p.get("ssn", ""), overflow)
    draw_text(c, "date_of_birth", p.get("date_of_birth", ""), overflow)
    draw_text(c, "case_or_cin", p.get("cin", ""), overflow)

    cs = spec.get("case", {})
    draw_text(c, "date_of_incident", cs.get("date_of_incident", ""), overflow)
    draw_text(c, "index_number", cs.get("index_number", ""), overflow)
    draw_text(c, "nyc_file_no", cs.get("nyc_file_no", ""), overflow)
    draw_text(c, "settlement_amount", cs.get("settlement_amount", ""), overflow)
    draw_text(c, "settlement_date", cs.get("settlement_date", ""), overflow)
    draw_text(c, "conference_date", cs.get("conference_date", ""), overflow)

    # Injury: try the short line; if it does not fit at base size, use the
    # full-width line below instead (then rider if even that fails).
    injury = cs.get("injury", "")
    if injury:
        x0, x1, _, _ = BOXES["injury"]
        if stringWidth(injury, FONT, BASE_SIZE) <= (x1 - x0 - 4):
            draw_text(c, "injury", injury, overflow)
        else:
            draw_text(c, "injury", "See below", overflow)
            draw_text(c, "injury_line2", injury, overflow)

    lien_type = (cs.get("lien_type") or "updated").lower()
    draw_check(c, "type_final" if lien_type.startswith("f") else "type_updated")

    a = {**FIRM_DEFAULTS, **{k: v for k, v in spec.get("attorney", {}).items() if v}}
    draw_check(c, "represents_defendant" if a["represents"].lower().startswith("d")
               else "represents_plaintiff")
    for k in ("firm_name", "firm_address", "attorney_name", "telephone", "email", "fax", "completed_by"):
        draw_text(c, k, a.get(k, ""), overflow)

    # Sections III / IV: two lines each. One entry may wrap across both
    # lines; more than two entries, or an entry that still won't fit, goes
    # to a rider page.
    def fill_section(keys, lines, title):
        if not lines:
            return
        if len(lines) > 2:
            draw_text(c, keys[0], "See attached rider", overflow)
            rider_lists[title] = lines
            return
        x0, x1, _, _ = BOXES[keys[0]]
        width = x1 - x0 - 4
        if len(lines) == 1 and fit_size(lines[0], width) is None:
            wrapped = wrap(lines[0], width, BASE_SIZE)
            if len(wrapped) <= 2:
                for key, seg in zip(keys, wrapped):
                    draw_text(c, key, seg, overflow)
                return
            draw_text(c, keys[0], "See attached rider", overflow)
            rider_lists[title] = lines
            return
        for key, line in zip(keys, lines):
            before = len(overflow)
            draw_text(c, key, line, overflow)
            if len(overflow) > before:
                overflow.pop()
                rider_lists[title] = lines
                return

    fill_section(("def_1", "def_2"),
                 [defendant_line(d) for d in spec.get("defendants", [])],
                 "III. Defendants and Defense Counsel")
    fill_section(("ins_1", "ins_2"),
                 [insurer_line(i) for i in spec.get("insurers", [])],
                 "IV. Insurance Carriers")

    c.save()
    buf.seek(0)

    # Any other single-field overflow also goes on the rider.
    if overflow:
        rider_lists["Other entries (full text)"] = [f"{k}: {v}" for k, v in overflow]
    return buf, rider_lists


def wrap(text, width, size):
    words, lines, cur = text.split(), [], ""
    for w in words:
        trial = f"{cur} {w}".strip()
        if stringWidth(trial, FONT, size) <= width:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def build_rider(spec, rider_lists):
    buf = BytesIO()
    c = canvas.Canvas(buf, pagesize=letter)
    margin, y = 72, PAGE_H - 72
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin, y, "RIDER TO UPDATED / FINAL LIEN REQUEST FAX FORM (W-588AA)")
    y -= 18
    c.setFont(FONT, 10)
    p = spec.get("plaintiff", {})
    cs = spec.get("case", {})
    hdr = f"Plaintiff: {p.get('name','')}    DOB: {p.get('date_of_birth','')}    " \
          f"Index No.: {cs.get('index_number','')}    Date of Incident: {cs.get('date_of_incident','')}"
    for line in wrap(hdr, PAGE_W - 2 * margin, 10):
        c.drawString(margin, y, line)
        y -= 14
    y -= 8
    for title, items in rider_lists.items():
        c.setFont("Helvetica-Bold", 11)
        c.drawString(margin, y, title)
        y -= 16
        c.setFont(FONT, 10)
        for n, item in enumerate(items, 1):
            for j, line in enumerate(wrap(item, PAGE_W - 2 * margin - 20, 10)):
                c.drawString(margin + (0 if j == 0 else 20), y, (f"{n}. " if j == 0 else "") + line)
                y -= 14
                if y < 72:
                    c.showPage()
                    c.setFont(FONT, 10)
                    y = PAGE_H - 72
            y -= 4
        y -= 8
    c.save()
    buf.seek(0)
    return buf


def main():
    if len(sys.argv) not in (3, 4):
        print(__doc__)
        sys.exit(1)
    spec_path, out_path = sys.argv[1], sys.argv[2]
    form_path = sys.argv[3] if len(sys.argv) == 4 else None
    spec = json.load(open(spec_path))

    overlay_buf, rider_lists = build_overlay(spec)
    import os
    if form_path and os.path.exists(form_path):
        form = PdfReader(form_path)
        source = "official blank PDF"
    else:
        form = PdfReader(make_blank_form())
        source = "embedded form"
    overlay = PdfReader(overlay_buf)
    page = form.pages[0]
    page.merge_page(overlay.pages[0])

    writer = PdfWriter()
    writer.add_page(page)
    if rider_lists:
        for rp in PdfReader(build_rider(spec, rider_lists)).pages:
            writer.add_page(rp)
    with open(out_path, "wb") as f:
        writer.write(f)
    print(f"Wrote {out_path} using {source}" + (" (with rider page)" if rider_lists else ""))


if __name__ == "__main__":
    main()
