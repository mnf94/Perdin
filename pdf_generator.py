"""
pdf_generator.py  —  Byru Finance Request Form
Fix:
1. Signature table: hapus garis antar baris kosong TTD
2. Item rows: skip row yang amount=0 dan description kosong
3. Logo: load dari file, fallback ke teks
"""

import io
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, Image
)

BYRU_BLUE    = colors.HexColor("#0055A5")   # untuk border/aksen tetap
BYRU_ORANGE  = colors.HexColor("#F47920")
LIGHT_BLUE   = colors.HexColor("#E6F0FA")
DARK_GRAY    = colors.HexColor("#333333")
MID_GRAY     = colors.HexColor("#666666")

# Template original colors
HDR_BG       = colors.HexColor("#A6A6A6")   # header tabel — biru gelap navy
HDR_TRX_BG   = colors.HexColor("#2E4D7B")   # sub header detail desc
ROW_ALT      = colors.HexColor("#F2F2F2")   # alternating row abu muda
TOTAL_BG     = colors.HexColor("#EDEDED")   # baris total abu
GRAND_BG     = colors.HexColor("#D9D9D9")   # grand total abu lebih gelap
BORDER_CLR   = colors.HexColor("#000000")   # border abu tipis
SIG_HDR      = colors.HexColor("#595959")   # signature header sama dengan tabel

PAGE_W, PAGE_H = A4
MARGIN = 16 * mm


def fmt_rp(v):
    return "Rp -" if v == 0 else f"Rp {int(v):,}".replace(",", ".")


def cb(checked: bool) -> str:
    return "[X]" if checked else "[  ]"


def ps(name, size=8, bold=False, color=DARK_GRAY, align=TA_LEFT):
    return ParagraphStyle(
        name, fontSize=size,
        fontName="Helvetica-Bold" if bold else "Helvetica",
        textColor=color, alignment=align, spaceAfter=1, leading=11
    )


def generate_pdf(data: dict) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            rightMargin=MARGIN, leftMargin=MARGIN,
                            topMargin=MARGIN, bottomMargin=MARGIN)

    S   = ps("n")
    SB  = ps("b",  bold=True)
    SC  = ps("c",  align=TA_CENTER)
    SR  = ps("r",  align=TA_RIGHT)
    SHC = ps("hc", bold=True, color=colors.black, align=TA_CENTER)
    SS  = ps("sm", size=7, color=MID_GRAY)
    ST  = ps("tot", bold=True, color=colors.black, align=TA_RIGHT)
    SBO = ps("bo",  bold=True, color=colors.black, align=TA_CENTER)

    W = PAGE_W - 2 * MARGIN
    story = []

    # ── HEADER ────────────────────────────────────────────────
    # Logo: coba load gambar, fallback teks
    logo_paths = [
        "byru_logo.png",
        os.path.join(os.path.dirname(__file__), "byru_logo.png"),
    ]
    logo_cell = None
    for lp in logo_paths:
        if os.path.exists(lp):
            try:
                logo_cell = Image(lp, width=70, height=28)
            except Exception:
                pass
            break
    if logo_cell is None:
        logo_cell = Paragraph(
            "<font color='#0055A5' size='15'><b>co </b></font>"
            "<font color='#F47920' size='15'><b>byru</b></font>", S)

    title = Paragraph(
        "<font color='#000000' size='13'><b>FINANCE REQUEST FORM</b></font><br/>"
        "<font color='#000000' size='8'>PT Solusi Kerah Byru</font>", S)

    transfer = [
        Paragraph("<b>Please Transfer to</b>", SB),
        Paragraph(f"<b>Bank</b>          : {data['bank']}", S),
        Paragraph(f"<b>Name</b>         : {data['recipient_name']}", S),
        Paragraph(f"<b>Account No</b> : {data['account_number']}", S),
        Paragraph(f"<b>Due Date</b>    : {data['due_date_label']}", S),
    ]

    ht = Table([[logo_cell, title, transfer]],
               colWidths=[W*0.15, W*0.38, W*0.47])
    ht.setStyle(TableStyle([
        ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
        ("BOX",          (2,0),(2,0),   0.8, BORDER_CLR),
        ("BACKGROUND",   (2,0),(2,0),   colors.white),
        ("TOPPADDING",   (2,0),(2,0),   5),
        ("BOTTOMPADDING",(2,0),(2,0),   5),
        ("LEFTPADDING",  (2,0),(2,0),   7),
        ("RIGHTPADDING", (2,0),(2,0),   7),
    ]))
    story.append(ht)
    story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER_CLR, spaceAfter=5))

    # ── DATE / NAME / COST CENTER ─────────────────────────────
    it = Table([
        [Paragraph("<b>Date</b>", SB), Paragraph(data["date"], S),
         Paragraph("", S), Paragraph("", S)],
        [Paragraph("<b>Requester Name</b>", SB), Paragraph(data["your_name"], S),
         Paragraph("<b>Cost Center</b>", SB), Paragraph(data["cost_center"], S)],
    ], colWidths=[W*0.18, W*0.32, W*0.18, W*0.32])
    it.setStyle(TableStyle([
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 4),
        ("BOTTOMPADDING", (0,0),(-1,-1), 4),
        ("BOX", (1,0),(1,0), 0.5, BORDER_CLR),
        ("BOX", (1,1),(1,1), 0.5, BORDER_CLR),
        ("BOX", (3,1),(3,1), 0.5, BORDER_CLR),
    ]))
    story.append(it)
    story.append(Spacer(1, 4))

    # ── TYPE OF EXPENSES ──────────────────────────────────────
    pic = data.get("pic_name", "")
    tt = Table([[
        Paragraph("<b>Type of Expenses</b>", SB),
        Paragraph(f"{cb(data['type_unclaimable'])} Unclaimable", S),
        Paragraph(f"{cb(data['type_claimable'])} <i>Claimable to PIC</i>", S),
        Paragraph(pic, S),
        Paragraph("<b><i>PIC</i></b>", SB),
        Paragraph("", S),
    ]], colWidths=[W*0.18, W*0.16, W*0.22, W*0.14, W*0.08, W*0.22])
    tt.setStyle(TableStyle([
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 4),
        ("BOTTOMPADDING", (0,0),(-1,-1), 4),
        ("BOX", (1,0),(1,0), 0.5, BORDER_CLR),
        ("BOX", (3,0),(3,0), 0.5, BORDER_CLR),
        ("BOX", (5,0),(5,0), 0.5, BORDER_CLR),
    ]))
    story.append(tt)
    story.append(Spacer(1, 3))

    # ── PURPOSE ───────────────────────────────────────────────
    p = data["purpose"]
    margin_txt = f"{p['margin_pct']:.1f}%" if p.get("margin_pct", 0) > 0 else "-"
    pt = Table([
        [Paragraph("<b>Purpose</b>", SB),
         Paragraph(f"{cb(p['expense'])} Expense", S),
         Paragraph(f"{cb(p['reimbursement'])} Reimbursement", S),
         Paragraph(f"{cb(p['medical'])} Medical", S),
         Paragraph(f"{cb(p['perkiraan'])} <i>Perkiraan ditagihkan</i>", S),
         Paragraph("", S)],
        [Paragraph("", S),
         Paragraph(f"{cb(p['petty_cash'])} Petty Cash", S),
         Paragraph(f"{cb(p['cash_advance'])} Cash Advance", S),
         Paragraph(f"{cb(p['settlement'])} Settlement", S),
         Paragraph("<b><i>% Margin</i></b>", SB),
         Paragraph(margin_txt, SC)],
    ], colWidths=[W*0.18, W*0.13, W*0.19, W*0.13, W*0.23, W*0.14])
    pt.setStyle(TableStyle([
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 3),
        ("BOTTOMPADDING", (0,0),(-1,-1), 3),
        ("BOX", (5,0),(5,0), 0.5, BORDER_CLR),
        ("BOX", (5,1),(5,1), 0.5, BORDER_CLR),
    ]))
    story.append(pt)
    story.append(Spacer(1, 6))

    # ── TRANSACTION TABLE ─────────────────────────────────────
    # Hanya tampilkan baris yang ada isinya (skip row kosong)
    cw = [W*0.14, W*0.61, W*0.25]
    rows = [
        [Paragraph("Trx Date", SHC),
         Paragraph("Description", SHC),
         Paragraph("Amount", SHC)],
        [Paragraph("", S),
         Paragraph(f"<b>{data.get('detail_descriptions','')}</b>", SB),
         Paragraph("", S)],
    ]

    raw_items = [i for i in data.get("items", [])
                 if i.get("description","").strip() or i.get("amount", 0) > 0]

    for item in raw_items:
        rows.append([
            Paragraph(item["date"], SC),
            Paragraph(item["description"], S),
            Paragraph(fmt_rp(item["amount"]), SR),
        ])

    n_data = len(raw_items)
    total_r = 2 + n_data

    rows.append([Paragraph("",S), Paragraph("<b>Total</b>",SB),
                 Paragraph(fmt_rp(data["subtotal"]), ST)])
    rows.append([Paragraph("",S), Paragraph("<b>Tax (if applicable)</b>",SB),
                 Paragraph(fmt_rp(data["tax"]), SR)])
    grand_r = total_r + 2
    rows.append([Paragraph("",S), Paragraph("<b>Grand Total</b>",SB),
                 Paragraph(fmt_rp(data["grand_total"]), ST)])

    # Alternating row colors hanya untuk data rows
    row_bgs = []
    for i in range(n_data):
        bg = colors.white if i % 2 == 0 else ROW_ALT
        row_bgs.append(("BACKGROUND", (0, 2+i), (-1, 2+i), bg))

    trx = Table(rows, colWidths=cw, repeatRows=1)
    trx_style = [
        ("BACKGROUND",    (0,0),(-1,0),      HDR_BG),          # header navy
        ("BACKGROUND",    (0,1),(-1,1),      ROW_ALT),         # detail desc abu muda
        ("SPAN",          (1,1),(2,1)),
        ("BACKGROUND",    (0,total_r),(-1,total_r), TOTAL_BG), # total abu
        ("BACKGROUND",    (0,total_r+1),(-1,total_r+1), colors.white),
        ("BACKGROUND",    (0,grand_r),(-1,grand_r), GRAND_BG), # grand total abu gelap
        ("LINEABOVE",     (0,total_r),(-1,total_r),   0.5, BORDER_CLR),
        ("LINEABOVE",     (0,grand_r),(-1,grand_r),   1,   DARK_GRAY),
        ("GRID",          (0,0),(-1,grand_r), 0.3, BORDER_CLR),
        ("BOX",           (0,0),(-1,grand_r), 0.8, DARK_GRAY),
        ("VALIGN",        (0,0),(-1,grand_r), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,grand_r), 3),
        ("BOTTOMPADDING", (0,0),(-1,grand_r), 3),
        ("LEFTPADDING",   (0,0),(-1,grand_r), 5),
        ("RIGHTPADDING",  (0,0),(-1,grand_r), 5),
        ("ALIGN",         (2,2),(2,grand_r),  "RIGHT"),
    ] + row_bgs

    trx.setStyle(TableStyle(trx_style))
    story.append(trx)
    story.append(Spacer(1, 7))

    # ── RULES ─────────────────────────────────────────────────
    story.append(Paragraph(
        "<i>Rules:<br/>"
        "1. Original receipt or invoice must be attached<br/>"
        "2. Receipt of current month is valid if it is claimed not later than "
        "10th of the following month<br/>"
        "3. Minimum level of approval is Manager<br/>"
        "4. Above 5 million required CEO and another C-Level Approval</i>", SS))
    story.append(Spacer(1, 8))

    # ── SIGNATURE TABLE ───────────────────────────────────────
    needs_director = data.get("needs_director", False)

    if needs_director:
        headers   = ["Claimed by", "Approved By", "Acknowledge by", "Director Approval"]
        sig_names = [data["claimer_name"], data["approver1_name"],
                     data["approver2_name"], data.get("director_name","")]
        sig_pos   = [data["claimer_dept"], data["approver1_pos"],
                     data["approver2_pos"], data.get("director_pos","")]
        col_w     = [W/4]*4
    else:
        headers   = ["Claimed by", "Approved By", "Acknowledge by"]
        sig_names = [data["claimer_name"], data["approver1_name"], data["approver2_name"]]
        sig_pos   = [data["claimer_dept"], data["approver1_pos"], data["approver2_pos"]]
        col_w     = [W/3]*3

    n = len(headers)

    # Signature rows: header / TTD space (1 baris tinggi) / nama / jabatan
    sig_rows = [
        [Paragraph(f"<b>{h}</b>", SHC) for h in headers],
        [Paragraph("", S)] * n,
        [Paragraph(f"<b>{nm}</b>",
                   SBO if (needs_director and i == n-1) else
                   ps(f"sn{i}", bold=True, align=TA_CENTER))
         for i, nm in enumerate(sig_names)],
        [Paragraph(pos, ps(f"st{i}", size=7, color=colors.grey, align=TA_CENTER))
         for i, pos in enumerate(sig_pos)],
    ]
    # rowHeights: row0=header, row1=TTD space 90pt, row2=nama, row3=jabatan
    sig_row_heights = [20, 90, 18, 14]

    sig_style = [
        ("BACKGROUND",    (0,0),(-1,0),  HDR_BG),
        ("TEXTCOLOR",     (0,0),(-1,0),  colors.white),
        ("ALIGN",         (0,0),(-1,-1), "CENTER"),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 4),
        ("BOTTOMPADDING", (0,0),(-1,-1), 4),
        ("BOX",           (0,0),(-1,-1), 0.8, colors.black),
        ("LINEAFTER",     (0,0),(n-2,-1), 0.5, BORDER_CLR),
        ("LINEABOVE",     (0,2),(-1,2),  0.5, BORDER_CLR),
        ("LINEABOVE",     (0,3),(-1,3),  0.3, BORDER_CLR),
    ]

    if needs_director:
        sig_style += [
            ("BACKGROUND", (n-1,0),(n-1,0), colors.HexColor("#595959")), # director header abu gelap
            ("BACKGROUND", (n-1,2),(n-1,3), colors.HexColor("#F2F2F2")), # director TTD space + nama/jabatan abu muda
        ]

    sig = Table(sig_rows, colWidths=col_w, rowHeights=sig_row_heights)
    sig.setStyle(TableStyle(sig_style))
    story.append(sig)

    doc.build(story)
    return buf.getvalue()
