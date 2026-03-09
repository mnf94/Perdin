"""
perdin_pdf_generator.py — Byru Form Perjalanan Dinas
Generates:
  1. form_perdin(data) → PDF bytes  (Form Perjalanan Dinas)
  2. form_lampiran(attachments) → PDF bytes  (1 page = 4 foto + deskripsi)
"""

import io
import os
from PIL import Image as PILImage
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, Image, PageBreak
)
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import ImageReader

# ── Colors ──────────────────────────────────────────────────────────────────
NAVY        = colors.HexColor("#1F3864")
ORANGE      = colors.HexColor("#F47920")
HDR_BG      = colors.HexColor("#595959")
LIGHT_GRAY  = colors.HexColor("#F2F2F2")
MID_GRAY    = colors.HexColor("#EDEDED")
DARK_GRAY   = colors.HexColor("#333333")
BORDER      = colors.HexColor("#000000")
WHITE       = colors.white

PAGE_W, PAGE_H = A4
MARGIN = 10 * mm
W = PAGE_W - 2 * MARGIN


# ── Style helpers ────────────────────────────────────────────────────────────
def ps(name, size=7, bold=False, color=DARK_GRAY, align=TA_LEFT):
    return ParagraphStyle(
        name, fontSize=size,
        fontName="Helvetica-Bold" if bold else "Helvetica",
        textColor=color, alignment=align,
        spaceAfter=0, spaceBefore=0, leading=9
    )

S   = ps("n")
SB  = ps("b", bold=True)
SC  = ps("c", align=TA_CENTER)
SBC = ps("bc", bold=True, align=TA_CENTER)
SR  = ps("r", align=TA_RIGHT)
SS  = ps("sm", size=6, color=colors.HexColor("#555555"))
SHC = ps("hc", bold=True, color=WHITE, align=TA_CENTER)
SH  = ps("h", bold=True, color=WHITE)


def fmt_rp(v):
    if v == 0 or v is None:
        return "Rp -"
    return f"Rp {int(v):,}".replace(",", ".")


def cb(checked):
    return "[X]" if checked else "[  ]"


# ── Expense rows helper ──────────────────────────────────────────────────────
EXPENSE_DEFAULTS = [
    ("BBM",                           "bbm",                         True),
    ("Toll",                          "toll",                        True),
    ("Parkir",                        "parkir",                      True),
    ("Kendaraan Umum",                "kendaraan_umum",              False),
    ("Tiket Pesawat (Pergi & Pulang)", "tiket_pesawat_pergi_pulang", False),
    ("Entertain",                     "entertain",                   False),
    ("lainnya (jika ada)",            "lainnya",                     False),
    ("Uang Saku",                     "uang_saku",                   False),
    ("Kontrakan/Rumah Kost",          "kontrakan_rumah_kost",        False),
]


def _build_expense_rows(data: dict, side: str) -> list:
    """
    side: 'pengajuan' or 'realisasi'
    Uses custom labels from data['expense_labels'] if present.
    All items (including SK/ON BILL) show their input value.
    SK/ON BILL items show the budget note in Nominal column.
    """
    expenses  = data.get(f"expenses_{side}", {})
    labels    = data.get("expense_labels", [d[0] for d in EXPENSE_DEFAULTS])
    is_sk_lst = data.get("expense_is_sk",  [d[2] for d in EXPENSE_DEFAULTS])
    keys      = [d[1] for d in EXPENSE_DEFAULTS]

    rows = []
    for i, key in enumerate(keys):
        label   = labels[i] if i < len(labels) else EXPENSE_DEFAULTS[i][0]
        is_sk   = is_sk_lst[i] if i < len(is_sk_lst) else EXPENSE_DEFAULTS[i][2]
        val     = expenses.get(key, 0) or 0
        nominal = "SESUAI SK" if (is_sk and i == 0) else ("ON BILL" if is_sk else fmt_rp(0))
        rows.append([
            Paragraph(f"{label} :", S),
            Paragraph(nominal, SC),
            Paragraph(fmt_rp(val), SR),
            Paragraph(fmt_rp(val), SR),
        ])
    return rows


def expense_rows_pengajuan(data):
    return _build_expense_rows(data, "pengajuan")


def expense_rows_realisasi(data):
    return _build_expense_rows(data, "realisasi")


# ── MAIN: generate Form Perdin PDF ──────────────────────────────────────────
def generate_perdin_pdf(data: dict, logo_path: str = None) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=landscape(A4),
        rightMargin=MARGIN, leftMargin=MARGIN,
        topMargin=8*mm, bottomMargin=8*mm
    )

    LW = PAGE_H - 2 * MARGIN   # landscape width
    story = []

    # ── DOC HEADER ───────────────────────────────────────────────────────────
    # Logo
    logo_cell = Paragraph(
        "<font color='#1F3864' size='12'><b>co </b></font>"
        "<font color='#F47920' size='12'><b>byru</b></font>", S)
    if logo_path and os.path.exists(logo_path):
        try:
            logo_cell = Image(logo_path, width=55, height=22)
        except Exception:
            pass

    meta_rows = [
        [Paragraph("No.Dokumen", SS), Paragraph(": SKB/FR-LHRG-PA-10", SS)],
        [Paragraph("Revisi", SS), Paragraph(":", SS)],
        [Paragraph("Klasifikasi", SS), Paragraph(": INTERNAL", SS)],
        [Paragraph("Halaman", SS), Paragraph(": 1/1", SS)],
    ]
    meta_tbl = Table(meta_rows, colWidths=[22*mm, 35*mm])
    meta_tbl.setStyle(TableStyle([
        ("TOPPADDING", (0,0),(-1,-1), 1),
        ("BOTTOMPADDING", (0,0),(-1,-1), 1),
    ]))

    title_cell = Paragraph(
        "<b>PT SOLUSI KERAH BYRU</b>", ps("title", size=11, bold=True, align=TA_CENTER))
    subtitle_cell = Paragraph(
        "<b>FORM PERJALANAN DINAS</b>", ps("sub", size=10, bold=True, align=TA_CENTER))

    header_tbl = Table(
        [[meta_tbl,
          Table([[title_cell], [subtitle_cell]], colWidths=[LW*0.45]),
          logo_cell]],
        colWidths=[LW*0.22, LW*0.60, LW*0.18]
    )
    header_tbl.setStyle(TableStyle([
        ("VALIGN", (0,0),(-1,-1), "MIDDLE"),
        ("BOX", (0,0),(-1,-1), 0.8, BORDER),
        ("LINEBEFORE", (1,0),(1,0), 0.5, BORDER),
        ("LINEBEFORE", (2,0),(2,0), 0.5, BORDER),
        ("ALIGN", (2,0),(2,0), "RIGHT"),
        ("RIGHTPADDING", (2,0),(2,0), 6),
    ]))
    story.append(header_tbl)
    story.append(Spacer(1, 2*mm))

    # ── IDENTITY ROWS ─────────────────────────────────────────────────────────
    id_data = [
        [Paragraph("No. Urut Perjalanan Dinas", S), Paragraph(":", S),
         Paragraph(data.get("no_urut", ""), S),
         Paragraph("", S), Paragraph("", S), Paragraph("", S)],
        [Paragraph("Nama Karyawan", S), Paragraph(":", S),
         Paragraph(f"<b>{data.get('your_name','')}</b>", SB),
         Paragraph("", S), Paragraph("", S), Paragraph("", S)],
        [Paragraph("Divisi", S), Paragraph(":", S),
         Paragraph(data.get("departement", ""), S),
         Paragraph("", S), Paragraph("", S), Paragraph("", S)],
        [Paragraph("Jabatan", S), Paragraph(":", S),
         Paragraph(data.get("jabatan", ""), S),
         Paragraph("", S), Paragraph("", S), Paragraph("", S)],
    ]
    id_tbl = Table(id_data, colWidths=[30*mm, 5*mm, LW*0.3, 5*mm, 30*mm, LW*0.2])
    id_tbl.setStyle(TableStyle([
        ("TOPPADDING", (0,0),(-1,-1), 2),
        ("BOTTOMPADDING", (0,0),(-1,-1), 2),
        ("BOX", (0,0),(-1,-1), 0.5, BORDER),
    ]))
    story.append(id_tbl)
    story.append(Spacer(1, 1*mm))

    # ── ALL CONTENT: unified 8-col tables (no nested wrappers = no overlap) ──
    half = LW / 2

    jenis     = data.get("jenis_perjalanan", "dalam_kota")
    transport = data.get("jenis_transportasi", [])

    LBL = 36 * mm
    VAL = half - LBL

    def p4(ll, lv, rl="", rv=""):
        return [Paragraph(ll, S), Paragraph(lv, S),
                Paragraph(rl, S), Paragraph(rv, S)]

    hasil_text = data.get("hasil_perjalanan", "")

    # Align left (9 rows) and right (6 rows) by matching row positions:
    # Row 0: headers
    # Row 1: Jenis Perjalanan  | Lokasi/Kota Tujuan
    # Row 2: Lokasi/Kota Tujuan | Jumlah Hari
    # Row 3: Jumlah Hari        | Tgl Keberangkatan
    # Row 4: Tgl Keberangkatan  | Tgl Kembali
    # Row 5: Tgl Kembali        | Hasil Perjalanan Dinas  ← Hasil starts here (auto height)
    # Row 6: Keperluan          | (empty)                 ← Hasil may span rows 5+
    # Row 7: Jenis Transportasi | (empty)
    # Row 8: Uang Muka          | (empty)
    info_rows = [
        # Row 0: headers
        [Paragraph("<b>PENGAJUAN PERJALANAN DINAS</b>", SHC), Paragraph("", SHC),
         Paragraph("<b>REALISASI PERJALANAN DINAS</b>", SHC), Paragraph("", SHC)],
        # Row 1
        p4("Jenis Perjalanan Dinas",
           f"{cb(jenis=='dalam_kota')} Dalam Kota  {cb(jenis=='luar_kota')} Luar Kota",
           "Lokasi / Kota Tujuan", f": {data.get('kota_tujuan','')}"),
        # Row 2
        p4("Lokasi / Kota Tujuan", f": {data.get('kota_tujuan','')}",
           "Jumlah Hari Perjalanan Dinas", f": {data.get('days_no','')}"),
        # Row 3
        p4("Jumlah Hari Perjalanan Dinas", f": {data.get('days_no','')}",
           "Tanggal Keberangkatan", f": {data.get('departure_date','')}"),
        # Row 4
        p4("Tanggal Keberangkatan", f": {data.get('departure_date','')}",
           "Tanggal Kembali", f": {data.get('return_date','')}"),
        # Row 5: Tgl Kembali kiri | Hasil Perjalanan Dinas kanan (spans rows 5-8)
        # Combine label+value in col2 because SPAN only renders the anchor cell (2,5)
        [Paragraph("Tanggal Kembali", S), Paragraph(f": {data.get('return_date','')}", S),
         Paragraph(f"Hasil Perjalanan Dinas : {hasil_text}", S), Paragraph("", S)],
        # Row 6
        p4("Keperluan Perjalanan Dinas", f": {data.get('purpose_trip','')}", "", ""),
        # Row 7
        p4("Jenis Transportasi",
           f"{cb('mobil_ops' in transport)} Mobil Operasional  "
           f"{cb('mobil_pribadi' in transport)} Mobil Pribadi  "
           f"{cb('motor' in transport)} Motor  "
           f"{cb('pesawat' in transport)} Pesawat  "
           f"{cb('kereta' in transport)} Kereta  "
           f"{cb('umum' in transport)} Umum", "", ""),
        # Row 8
        p4("Uang Muka", f": {fmt_rp(data.get('uang_muka', 0))}", "", ""),
    ]

    cw4 = [LBL, VAL, LBL, VAL]
    info_tbl = Table(info_rows, colWidths=cw4)
    info_tbl.setStyle(TableStyle([
        # Header row background + span
        ("BACKGROUND",    (0,0),(1,0), HDR_BG),
        ("BACKGROUND",    (2,0),(3,0), HDR_BG),
        ("SPAN",          (0,0),(1,0)),
        ("SPAN",          (2,0),(3,0)),
        ("TOPPADDING",    (0,0),(-1,0), 4),
        ("BOTTOMPADDING", (0,0),(-1,0), 4),
        # Content padding
        ("TOPPADDING",    (0,1),(-1,-1), 2),
        ("BOTTOMPADDING", (0,1),(-1,-1), 2),
        ("LEFTPADDING",   (0,0),(-1,-1), 3),
        # Outer boxes per side
        ("BOX",       (0,0),(1,-1), 0.8, BORDER),
        ("BOX",       (2,0),(3,-1), 0.8, BORDER),
        # Row dividers — left side all rows, right side only rows 1-4 (rows 5-8 merged)
        ("LINEBELOW", (0,1),(1,-2), 0.3, colors.lightgrey),
        ("LINEBELOW", (2,1),(3,4),  0.3, colors.lightgrey),
        # Span Hasil Perjalanan Dinas (right cols 2-3) across rows 5 to 8
        ("SPAN",      (2,5),(3,-1)),
        ("VALIGN",    (0,0),(-1,-1), "TOP"),
        ("LEFTPADDING",  (2,0),(3,-1), 6),
    ]))
    story.append(info_tbl)
    story.append(Spacer(1, 1*mm))

    # ── EXPENSE TABLE: 8 cols ────────────────────────────────────────────────
    EW = [half*0.44, half*0.18, half*0.19, half*0.19]
    cw8 = EW + EW

    def ehc(txt):
        return Paragraph(f"<b>{txt}</b>", SHC)

    rows_p = expense_rows_pengajuan(data)
    rows_r = expense_rows_realisasi(data)
    total_p = sum(data.get("expenses_pengajuan", {}).values())
    total_r = sum(data.get("expenses_realisasi", {}).values())

    exp_data = [[
        ehc("Detail Pengeluaran\nPerjalanan Dinas"), ehc("Nominal\nBudget"),
        ehc("Total Pengajuan"), ehc("Jumlah"),
        ehc("Detail Pengeluaran\nPerjalanan Dinas"), ehc("Nominal\nBudget"),
        ehc("Total Realisasi"), ehc("Jumlah"),
    ]]
    for i in range(max(len(rows_p), len(rows_r))):
        lr = rows_p[i] if i < len(rows_p) else [Paragraph("",S)]*4
        rr = rows_r[i] if i < len(rows_r) else [Paragraph("",S)]*4
        exp_data.append(lr + rr)
    exp_data.append([
        Paragraph("<b>Jumlah</b>", SBC), Paragraph("", SC),
        Paragraph(fmt_rp(total_p), SR),  Paragraph(fmt_rp(total_p), SR),
        Paragraph("<b>Jumlah</b>", SBC), Paragraph("", SC),
        Paragraph(fmt_rp(total_r), SR),  Paragraph(fmt_rp(total_r), SR),
    ])

    exp_tbl = Table(exp_data, colWidths=cw8)
    exp_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(3,0), HDR_BG),
        ("BACKGROUND",    (4,0),(7,0), HDR_BG),
        ("TOPPADDING",    (0,0),(7,0), 3),
        ("BOTTOMPADDING", (0,0),(7,0), 3),
        ("TOPPADDING",    (0,1),(7,-1), 2),
        ("BOTTOMPADDING", (0,1),(7,-1), 2),
        ("GRID",     (0,0),(3,-1), 0.3, BORDER),
        ("GRID",     (4,0),(7,-1), 0.3, BORDER),
        ("BOX",      (0,0),(3,-1), 0.8, BORDER),
        ("BOX",      (4,0),(7,-1), 0.8, BORDER),
        ("BACKGROUND", (0,-1),(3,-1), MID_GRAY),
        ("BACKGROUND", (4,-1),(7,-1), MID_GRAY),
        ("FONTNAME",   (0,-1),(7,-1), "Helvetica-Bold"),
        ("VALIGN",     (0,0),(7,-1), "MIDDLE"),
        ("ALIGN",      (1,1),(3,-1), "RIGHT"),
        ("ALIGN",      (1,0),(3,0),  "CENTER"),
        ("ALIGN",      (5,1),(7,-1), "RIGHT"),
        ("ALIGN",      (5,0),(7,0),  "CENTER"),
        ("LEFTPADDING", (4,0),(7,-1), 4),
    ]))
    story.append(exp_tbl)
    story.append(Spacer(1, 2*mm))
    story.append(Spacer(1, 2*mm))

    # ── SIGNATURE TABLE ───────────────────────────────────────────────────────
    # LEFT  (Pengajuan): Jakarta,{dep} | Karyawan ybs, | Disetujui, | Disetujui,
    #                     your_name    | approvers[0]  | approvers[1]
    # RIGHT (Realisasi): Jakarta,{ret} | Disetujui,    | Mengetahui, | (kosong)
    #                     approvers[2] | approvers[3]  |
    dep_date  = data.get("departure_date", "")
    ret_date  = data.get("return_date", "")
    your_name = data.get("your_name", "")
    approvers = list(data.get("approvers", []))
    while len(approvers) < 4:
        approvers.append("")

    SBS = ps("sbs", size=7, bold=True, align=TA_CENTER)
    SSC = ps("ssc", size=6.5, color=colors.HexColor("#444444"), align=TA_CENTER)

    # ── SIGNATURE: single 6-col table (no wrapper = no double border) ──────
    while len(approvers) < 5:
        approvers.append("")

    cw6 = [half / 3] * 6   # 6 equal columns across full width

    right_col3_hdr = "Mengetahui," if approvers[4].strip() else ""

    sig_rows = [
        # Row 0: date headers — span cols 0-2 (left) and 3-5 (right)
        [Paragraph(f"Jakarta, {dep_date}", SSC), Paragraph("", S), Paragraph("", S),
         Paragraph(f"Jakarta, {ret_date}", SSC),  Paragraph("", S), Paragraph("", S)],
        # Row 1: role labels
        [Paragraph("<b>Karyawan ybs,</b>", SBS), Paragraph("<b>Disetujui,</b>", SBS),
         Paragraph("<b>Disetujui,</b>", SBS),
         Paragraph("<b>Disetujui,</b>", SBS),    Paragraph("<b>Mengetahui,</b>", SBS),
         Paragraph(f"<b>{right_col3_hdr}</b>", SBS)],
        # Rows 2-3: signature space
        [Paragraph("", S)] * 6,
        [Paragraph("", S)] * 6,
        # Row 4: names
        [Paragraph(f"<b>{your_name}</b>",    SBS), Paragraph(f"<b>{approvers[0]}</b>", SBS),
         Paragraph(f"<b>{approvers[1]}</b>", SBS),
         Paragraph(f"<b>{approvers[2]}</b>", SBS), Paragraph(f"<b>{approvers[3]}</b>", SBS),
         Paragraph(f"<b>{approvers[4]}</b>", SBS)],
    ]

    sig_tbl = Table(sig_rows, colWidths=cw6, rowHeights=[10, 13, 20, 20, 13])
    sig_tbl.setStyle(TableStyle([
        # Outer box
        ("BOX",        (0,0),(-1,-1), 0.8, BORDER),
        # Vertical dividers between all 6 cols (rows 1-4 only, not date row)
        ("LINEAFTER",  (0,1),(4,4),   0.5, BORDER),
        # Extra visible divider between left half and right half
        ("LINEAFTER",  (2,0),(2,4),   0.8, BORDER),
        # Horizontal line above name row
        ("LINEABOVE",  (0,4),(-1,4),  0.5, BORDER),
        # Date row: span left 3 cols, right 3 cols
        ("SPAN",       (0,0),(2,0)),
        ("SPAN",       (3,0),(5,0)),
        ("BACKGROUND", (0,0),(2,0),  LIGHT_GRAY),
        ("BACKGROUND", (3,0),(5,0),  LIGHT_GRAY),
        ("ALIGN",      (0,0),(-1,-1), "CENTER"),
        ("VALIGN",     (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING", (0,0),(-1,-1), 2),
        ("BOTTOMPADDING", (0,0),(-1,-1), 2),
    ]))
    story.append(sig_tbl)

    doc.build(story)
    return buf.getvalue()


# ── LAMPIRAN PDF: 4 foto per halaman ────────────────────────────────────────
def generate_lampiran_pdf(attachments: list, perdin_data: dict = None) -> bytes:
    """
    attachments: list of dicts: [{image_bytes, description, label}]
    Renders 4 per page in a 2x2 grid.
    """
    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf, pagesize=A4)
    pw, ph = A4

    margin = 12 * mm
    cell_w = (pw - 2 * margin - 6 * mm) / 2
    cell_h = (ph - 2 * margin - 30 * mm - 6 * mm) / 2

    def draw_page_header(c, page_num, total_pages):
        c.setFont("Helvetica-Bold", 9)
        c.setFillColor(NAVY)
        c.drawString(margin, ph - margin - 5*mm, "PT SOLUSI KERAH BYRU")
        c.setFont("Helvetica", 8)
        c.setFillColor(DARK_GRAY)
        if perdin_data:
            name = perdin_data.get("your_name", "")
            dept = perdin_data.get("departement", "")
            c.drawString(margin, ph - margin - 10*mm, f"Nama: {name}   |   Divisi: {dept}")
        c.setFont("Helvetica", 7)
        c.setFillColor(colors.grey)
        c.drawRightString(pw - margin, ph - margin - 5*mm, f"Halaman {page_num} dari {total_pages}")
        # Title
        c.setFont("Helvetica-Bold", 11)
        c.setFillColor(DARK_GRAY)
        c.drawCentredString(pw / 2, ph - margin - 18*mm, "LAMPIRAN PERJALANAN DINAS")
        # Line
        c.setStrokeColor(ORANGE)
        c.setLineWidth(1.5)
        c.line(margin, ph - margin - 21*mm, pw - margin, ph - margin - 21*mm)

    def draw_slot(c, img_bytes, description, label, x, y, w, h):
        # Border
        c.setStrokeColor(colors.HexColor("#BFBFBF"))
        c.setLineWidth(0.5)
        c.roundRect(x, y, w, h, 4, stroke=1, fill=0)

        # Label header
        c.setFillColor(HDR_BG)
        c.roundRect(x, y + h - 14*mm, w, 14*mm, 4, stroke=0, fill=1)
        c.setFont("Helvetica-Bold", 8)
        c.setFillColor(WHITE)
        c.drawCentredString(x + w/2, y + h - 8*mm, label)

        # Image area
        img_area_y = y + 16*mm
        img_area_h = h - 14*mm - 18*mm - 2*mm
        img_area_w = w - 6*mm

        if img_bytes:
            try:
                # Open with PIL, normalize to RGB
                pil_img = PILImage.open(io.BytesIO(img_bytes))
                if pil_img.mode not in ("RGB", "L"):
                    pil_img = pil_img.convert("RGB")

                # Save normalized image to BytesIO, wrap with ImageReader for ReportLab
                img_io = io.BytesIO()
                pil_img.save(img_io, format="JPEG", quality=85)
                img_io.seek(0)
                rl_img = ImageReader(img_io)

                # Fit image maintaining aspect ratio inside slot
                iw, ih = pil_img.size
                scale = min(img_area_w / iw, img_area_h / ih)
                disp_w = iw * scale
                disp_h = ih * scale
                ix = x + 3*mm + (img_area_w - disp_w) / 2
                iy = img_area_y + (img_area_h - disp_h) / 2
                c.drawImage(rl_img, ix, iy, width=disp_w, height=disp_h)
            except Exception as _img_err:
                # Draw error placeholder
                c.setFillColor(LIGHT_GRAY)
                c.rect(x + 3*mm, img_area_y, img_area_w, img_area_h, fill=1, stroke=0)
                c.setFont("Helvetica", 7)
                c.setFillColor(colors.grey)
                c.drawCentredString(x + w/2, img_area_y + img_area_h/2, f"[ Error: {str(_img_err)[:40]} ]")
        else:
            # Empty slot
            c.setFillColor(colors.HexColor("#F5F5F5"))
            c.rect(x + 3*mm, img_area_y, img_area_w, img_area_h, fill=1, stroke=0)
            c.setStrokeColor(colors.HexColor("#DDDDDD"))
            c.setDash(3, 3)
            c.rect(x + 3*mm, img_area_y, img_area_w, img_area_h, fill=0, stroke=1)
            c.setDash()
            c.setFont("Helvetica", 8)
            c.setFillColor(colors.HexColor("#BBBBBB"))
            c.drawCentredString(x + w/2, img_area_y + img_area_h/2, "( tidak ada foto )")

        # Description box
        desc_y = y + 2*mm
        desc_h = 14*mm
        c.setFillColor(colors.HexColor("#F8FAFD"))
        c.rect(x + 3*mm, desc_y, w - 6*mm, desc_h, fill=1, stroke=0)
        c.setStrokeColor(colors.HexColor("#E0E8F0"))
        c.setLineWidth(0.3)
        c.rect(x + 3*mm, desc_y, w - 6*mm, desc_h, fill=0, stroke=1)
        c.setFont("Helvetica", 7)
        c.setFillColor(DARK_GRAY)

        # Word-wrap description
        max_chars = int((w - 8*mm) / 3.5)
        words = (description or "").split()
        lines = []
        cur = ""
        for word in words:
            if len(cur) + len(word) + 1 <= max_chars:
                cur = (cur + " " + word).strip()
            else:
                if cur:
                    lines.append(cur)
                cur = word
        if cur:
            lines.append(cur)
        lines = lines[:2]  # max 2 lines

        line_y = desc_y + desc_h - 5*mm
        for line in lines:
            c.drawCentredString(x + w/2, line_y, line)
            line_y -= 4*mm

    # Calculate pages
    total_pages = max(1, -(-len(attachments) // 4))  # ceiling div
    top_offset = 24 * mm

    for page_idx in range(total_pages):
        draw_page_header(c, page_idx + 1, total_pages)

        slot_indices = [page_idx * 4 + i for i in range(4)]
        positions = [
            (margin,              margin + cell_h + 6*mm),   # top-left
            (margin + cell_w + 6*mm, margin + cell_h + 6*mm), # top-right
            (margin,              margin),                     # bottom-left
            (margin + cell_w + 6*mm, margin),                 # bottom-right
        ]

        for i, slot_idx in enumerate(slot_indices):
            x, y = positions[i]
            if slot_idx < len(attachments):
                att = attachments[slot_idx]
                draw_slot(c, att.get("image_bytes"), att.get("description", ""),
                          att.get("label", f"Lampiran {slot_idx+1}"), x, y, cell_w, cell_h)
            else:
                draw_slot(c, None, "", f"Lampiran {slot_idx+1}", x, y, cell_w, cell_h)

        if page_idx < total_pages - 1:
            c.showPage()

    c.save()
    return buf.getvalue()
