import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
from datetime import date, timedelta
import io
import json
import os
import base64

from pdf_generator import generate_pdf
from perdin_pdf_generator import generate_perdin_pdf, generate_lampiran_pdf

# ─── Page config ──────────────────────────────────────────────
st.set_page_config(
    page_title="Byru — Form Keuangan & Perjalanan Dinas",
    page_icon="🏢",
    layout="wide",
)

# ─── CSS ──────────────────────────────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #F5F7FA; font-family: 'Segoe UI', sans-serif; }
    .section-header {
        background-color: #1F3864;
        color: white;
        padding: 7px 14px;
        border-radius: 4px;
        font-weight: bold;
        font-size: 0.9em;
        margin: 18px 0 8px 0;
        letter-spacing: 0.3px;
    }
    .info-box {
        background: #EEF3FA;
        border-left: 4px solid #1F3864;
        padding: 9px 14px;
        border-radius: 4px;
        margin-bottom: 10px;
        font-size: 0.88em;
        color: #1F3864;
    }
    .rule-box {
        background: #F9F9F9;
        border-left: 4px solid #BFBFBF;
        padding: 10px 14px;
        border-radius: 4px;
        font-size: 0.84em;
        color: #444444;
        font-style: italic;
    }
    .warning-5jt {
        background: #FFF3CD;
        border-left: 4px solid #F47920;
        padding: 10px 14px;
        border-radius: 4px;
        font-size: 0.9em;
        color: #7a3800;
        font-weight: bold;
        margin-bottom: 10px;
    }
    .stButton > button[kind="primary"] {
        background-color: #1F3864 !important;
        border-color: #1F3864 !important;
        color: white !important;
    }
    .stButton > button[kind="primary"]:hover {
        background-color: #162a4a !important;
    }
    hr { border-color: #BFBFBF !important; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background-color: #E8EDF5;
        border-radius: 6px 6px 0 0;
        padding: 8px 20px;
        font-weight: 600;
        font-size: 0.9em;
        color: #1F3864;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1F3864 !important;
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# ─── Logo ─────────────────────────────────────────────────────
_script_dir = os.path.dirname(os.path.abspath(__file__))
_logo_path  = os.path.join(_script_dir, "byru_logo.png")

col_logo, col_title = st.columns([1, 5])
with col_logo:
    if os.path.exists(_logo_path):
        st.image(_logo_path, width=110)
    else:
        st.markdown(
            "<div style='font-size:2em;font-weight:900;color:#1F3864;'>"
            "co <span style='color:#F47920'>byru</span></div>",
            unsafe_allow_html=True)
with col_title:
    st.markdown(
        "<h2 style='margin-bottom:0;color:#1F3864;'>BYRU — FORM CENTER</h2>"
        "<p style='margin-top:2px;color:#666;font-size:0.9em;'>PT Solusi Kerah Byru</p>",
        unsafe_allow_html=True)
st.markdown("<hr style='border:1px solid #BFBFBF;margin-top:0;'>", unsafe_allow_html=True)


# ─── Helper ───────────────────────────────────────────────────
def next_tue_or_thu(from_date: date) -> date:
    for i in range(1, 8):
        d = from_date + timedelta(days=i)
        if d.weekday() in (1, 3):
            return d

BANKS = ["BCA", "Mandiri", "BNI", "BRI", "CIMB Niaga", "Permata", "Danamon", "BSI", "Lainnya"]

EXPENSE_DEFAULTS = [
    ("BBM",                           "bbm",                             True),
    ("Toll",                          "toll",                            True),
    ("Parkir",                        "parkir",                          True),
    ("Kendaraan Umum",                "kendaraan_umum",                  False),
    ("Tiket Pesawat (Pergi & Pulang)", "tiket_pesawat_pergi_pulang",     False),
    ("Entertain",                     "entertain",                       False),
    ("lainnya (jika ada)",            "lainnya",                         False),
    ("Uang Saku",                     "uang_saku",                       False),
    ("Kontrakan/Rumah Kost",          "kontrakan_rumah_kost",            False),
]

# ─── TABS ─────────────────────────────────────────────────────
tab_finance, tab_perdin = st.tabs(["💰 Finance Request", "✈️ Perjalanan Dinas"])


# ╔══════════════════════════════════════════════════════════════╗
# ║  TAB 1 — FINANCE REQUEST (existing)                         ║
# ╚══════════════════════════════════════════════════════════════╝
with tab_finance:

    # ── SECTION 1: Informasi ─────────────────────────────────────
    st.markdown('<div class="section-header">📋 Informasi Pengajuan</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        req_date = st.date_input("Tanggal Pengajuan", value=date.today(), key="fr_date")
    with col2:
        your_name = st.text_input("Nama Pemohon", placeholder="Nama lengkap", key="fr_name")
    with col3:
        cost_center = st.text_input("Cost Center", placeholder="Dept / Project", key="fr_cc")

    today = date.today()
    due_date = next_tue_or_thu(today)
    st.markdown(
        f'<div class="info-box">📅 <b>Due Date:</b> '
        f'<b>{due_date.strftime("%A, %d %B %Y")}</b> — '
        f'Selasa/Kamis terdekat dari hari ini ({today.strftime("%d %b %Y")})</div>',
        unsafe_allow_html=True)

    # ── SECTION 2: Transfer ──────────────────────────────────────
    st.markdown('<div class="section-header">🏦 Tujuan Transfer</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns(3)
    with col1:
        bank = st.selectbox("Bank", BANKS, key="fr_bank")
        if bank == "Lainnya":
            bank = st.text_input("Nama Bank", placeholder="Isi nama bank", key="fr_bank_lain")
    with col2:
        recipient_name = st.text_input("Nama Rekening", placeholder="Nama pemilik rekening", key="fr_rek_name")
    with col3:
        account_number = st.text_input("Nomor Rekening", placeholder="Contoh: 1234567890", key="fr_rek_no")

    # ── SECTION 3: Type & Purpose ────────────────────────────────
    st.markdown('<div class="section-header">🗂️ Jenis & Tujuan</div>', unsafe_allow_html=True)
    col_type, col_purpose = st.columns(2)
    with col_type:
        st.markdown("**Type of Expenses**")
        type_unclaimable = st.checkbox("Unclaimable", key="fr_unclaim")
        type_claimable   = st.checkbox("Claimable to PIC", key="fr_claim")
        pic_name = ""
        if type_claimable:
            pic_name = st.text_input("Nama PIC", placeholder="PIC yang akan di-charge", key="fr_pic")
    with col_purpose:
        st.markdown("**Purpose**")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            p_expense       = st.checkbox("Expense", key="fr_exp")
            p_reimbursement = st.checkbox("Reimbursement", key="fr_reimb")
            p_medical       = st.checkbox("Medical", key="fr_med")
            p_perkiraan     = st.checkbox("Perkiraan ditagihkan", key="fr_perk")
        with col_p2:
            p_pettycash   = st.checkbox("Petty Cash", key="fr_petty")
            p_cashadvance = st.checkbox("Cash Advance", key="fr_ca")
            p_settlement  = st.checkbox("Settlement", key="fr_sett")
            margin_pct    = st.number_input("% Margin", min_value=0.0, max_value=100.0, step=0.5, value=0.0, key="fr_margin")

    # ── SECTION 4: Detail Transaksi ──────────────────────────────
    st.markdown('<div class="section-header">🧾 Detail Transaksi</div>', unsafe_allow_html=True)
    detail_description = st.text_input("Detail Descriptions",
        placeholder="Contoh: Reimbursement bulan Maret 2025", key="fr_detail")

    st.markdown("**Daftar Item** *(maks 15 baris)*")
    hcols = st.columns([2, 5, 3])
    hcols[0].markdown("<span style='color:#1F3864;font-weight:bold;'>Tanggal</span>", unsafe_allow_html=True)
    hcols[1].markdown("<span style='color:#1F3864;font-weight:bold;'>Deskripsi</span>", unsafe_allow_html=True)
    hcols[2].markdown("<span style='color:#1F3864;font-weight:bold;'>Jumlah (Rp)</span>", unsafe_allow_html=True)

    items = []
    for i in range(1, 16):
        cols = st.columns([2, 5, 3])
        with cols[0]:
            item_date = st.date_input("", value=req_date, key=f"fr_idate_{i}", label_visibility="collapsed")
        with cols[1]:
            item_desc = st.text_input("", placeholder=f"Item {i}", key=f"fr_idesc_{i}", label_visibility="collapsed")
        with cols[2]:
            item_amt = st.number_input("", min_value=0, step=1000, key=f"fr_iamt_{i}", label_visibility="collapsed")
        if item_desc.strip() or item_amt > 0:
            items.append({"date": item_date.strftime("%d/%m/%Y"), "description": item_desc, "amount": item_amt})

    # ── Totals ───────────────────────────────────────────────────
    st.markdown("<hr style='border:0.5px solid #BFBFBF;'>", unsafe_allow_html=True)
    subtotal = sum(i["amount"] for i in items)
    col_t1, col_t2 = st.columns([3, 1])
    with col_t1:
        tax_amount = st.number_input("Tax / Pajak (Rp)", min_value=0, step=1000, value=0, key="fr_tax")
    with col_t2:
        grand_total = subtotal - tax_amount
        st.metric("Subtotal", f"Rp {subtotal:,.0f}".replace(",", "."))
        st.metric("Grand Total", f"Rp {grand_total:,.0f}".replace(",", "."))

    # ── SECTION 5: Approval ──────────────────────────────────────
    st.markdown('<div class="section-header">✍️ Tanda Tangan & Persetujuan</div>', unsafe_allow_html=True)
    needs_director = grand_total >= 5_000_000
    if needs_director:
        st.markdown(
            '<div class="warning-5jt">⚠️ Grand Total ≥ Rp 5.000.000 — Wajib approval Director: '
            '<b>Nathaniel Nugroho Liman</b></div>', unsafe_allow_html=True)

    n_cols = 4 if needs_director else 3
    sig_cols = st.columns(n_cols)
    with sig_cols[0]:
        st.markdown("<b style='color:#1F3864;'>Claimed By</b>", unsafe_allow_html=True)
        claimer_name = st.text_input("Nama Pemohon", value=your_name, key="fr_signer")
        claimer_dept = st.text_input("Departemen", key="fr_dept")
    with sig_cols[1]:
        st.markdown("<b style='color:#1F3864;'>Approved By</b>", unsafe_allow_html=True)
        approver1_name = st.text_input("Nama Approver 1", key="fr_app1")
        approver1_pos  = st.text_input("Jabatan Approver 1", key="fr_app1pos")
    with sig_cols[2]:
        st.markdown("<b style='color:#1F3864;'>Acknowledge By</b>", unsafe_allow_html=True)
        approver2_name = st.text_input("Nama Approver 2", key="fr_app2")
        approver2_pos  = st.text_input("Jabatan Approver 2", key="fr_app2pos")
    if needs_director:
        with sig_cols[3]:
            st.markdown("<b style='color:#F47920;'>Director Approval ⚠️</b>", unsafe_allow_html=True)
            st.text_input("Nama Director", value="Nathaniel Nugroho Liman", disabled=True, key="fr_dir")
            st.text_input("Jabatan", value="Director", disabled=True, key="fr_dirpos")

    st.markdown("""
    <div class="rule-box"><b>Rules:</b><br>
    1. Original receipt or invoice must be attached<br>
    2. Receipt of current month is valid if claimed not later than 10th of the following month<br>
    3. Minimum level of approval is Manager<br>
    4. Above 5 million required CEO and another C-Level Approval
    </div>""", unsafe_allow_html=True)

    # ── Generate Email HTML ───────────────────────────────────────
    def generate_email_html(form_data: dict, logo_path: str) -> str:
        def fmt(v):
            return f"Rp {int(v):,}".replace(",", ".") if v else "Rp -"
        logo_tag = '<span style="font-size:18px;font-weight:900;color:#1F3864;font-family:Arial;">co <span style="color:#F47920;">byru</span></span>'
        if os.path.exists(logo_path):
            try:
                with open(logo_path, "rb") as f:
                    b64 = base64.b64encode(f.read()).decode()
                ext  = logo_path.rsplit(".", 1)[-1].lower()
                mime = "image/png" if ext == "png" else "image/jpeg"
                logo_tag = f'<img src="data:{mime};base64,{b64}" width="72" height="auto" style="display:block;width:72px;max-width:72px;height:auto;border:0;" alt="byru">'
            except Exception:
                pass
        p = form_data.get("purpose", {})
        purpose_list = [k for k, v in {
            "Expense": p.get("expense"), "Reimbursement": p.get("reimbursement"),
            "Medical": p.get("medical"), "Petty Cash": p.get("petty_cash"),
            "Cash Advance": p.get("cash_advance"), "Settlement": p.get("settlement"),
            "Perkiraan ditagihkan": p.get("perkiraan"),
        }.items() if v]
        purpose_str = ", ".join(purpose_list) if purpose_list else "-"
        item_rows = ""
        for i, item in enumerate(form_data.get("items", [])):
            bg = "#FFFFFF" if i % 2 == 0 else "#F8FAFC"
            item_rows += f"""
        <tr style="background:{bg};">
          <td style="padding:11px 16px;font-size:12px;color:#64748B;border-bottom:1px solid #E8EDF2;font-family:'Courier New',monospace;">{item["date"]}</td>
          <td style="padding:11px 16px;font-size:12px;color:#334155;border-bottom:1px solid #E8EDF2;">{item["description"]}</td>
          <td style="padding:11px 16px;font-size:12px;color:#1E293B;text-align:right;border-bottom:1px solid #E8EDF2;font-family:'Courier New',monospace;font-weight:600;">{fmt(item["amount"])}</td>
        </tr>"""
        yn   = form_data.get("your_name", "")
        dept = form_data.get("claimer_dept", "")
        return f"""<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><title>Finance Request – Byru</title></head>
<body style="margin:0;padding:0;background-color:#EEF2F7;font-family:'Segoe UI',Helvetica,Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="background-color:#EEF2F7;padding:40px 0;"><tr><td align="center">
<table width="620" cellpadding="0" cellspacing="0" style="background:#FFFFFF;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(31,56,100,0.10);">
<tr><td style="height:4px;background:linear-gradient(90deg,#1F3864 0%,#2E5BA8 50%,#F47920 100%);font-size:0;line-height:0;">&nbsp;</td></tr>
<tr><td style="padding:28px 36px 22px 36px;border-bottom:1px solid #E8EDF2;">
  <table width="100%" cellpadding="0" cellspacing="0"><tr>
    <td style="vertical-align:middle;">{logo_tag}</td>
    <td align="right" style="vertical-align:middle;">
      <span style="display:inline-block;background:#FFF4EC;color:#F47920;font-size:9px;font-weight:700;padding:5px 14px;border-radius:20px;letter-spacing:2px;text-transform:uppercase;border:1px solid #FFD4B0;">FINANCE REQUEST</span><br>
      <span style="font-size:10px;color:#94A3B8;margin-top:5px;display:block;text-align:right;">PT Solusi Kerah Byru</span>
    </td>
  </tr></table>
</td></tr>
<tr><td style="padding:28px 36px 24px 36px;background:#F8FAFD;">
  <table width="100%" cellpadding="0" cellspacing="0">
    <tr>
      <td style="vertical-align:top;">
        <p style="margin:0 0 4px 0;font-size:9px;color:#F47920;font-weight:700;letter-spacing:3px;text-transform:uppercase;">Pengajuan Baru</p>
        <h1 style="margin:0 0 6px 0;font-size:20px;font-weight:700;color:#1E293B;letter-spacing:-0.3px;">{yn}</h1>
        <p style="margin:0;font-size:11px;color:#94A3B8;">{form_data.get("date","")} &nbsp;·&nbsp; {form_data.get("cost_center","")} &nbsp;·&nbsp; {purpose_str}</p>
      </td>
      <td align="right" style="vertical-align:top;">
        <p style="margin:0 0 2px 0;font-size:9px;color:#94A3B8;letter-spacing:2px;text-transform:uppercase;text-align:right;">Grand Total</p>
        <p style="margin:0;font-size:24px;font-weight:700;color:#F47920;letter-spacing:-0.5px;">{fmt(form_data.get("grand_total",0))}</p>
      </td>
    </tr>
    <tr><td colspan="2" style="padding-top:18px;">
      <table cellpadding="0" cellspacing="0"><tr>
        <td style="background:#1F3864;border-radius:6px 0 0 6px;padding:7px 14px;"><span style="font-size:9px;color:#fff;font-weight:700;letter-spacing:2px;text-transform:uppercase;">DUE DATE</span></td>
        <td style="background:#E8EDF5;border-radius:0 6px 6px 0;padding:7px 16px;"><span style="font-size:11px;color:#1F3864;font-weight:600;">{form_data.get("due_date_label","")}</span></td>
      </tr></table>
    </td></tr>
  </table>
</td></tr>
<tr><td style="padding:28px 36px 8px 36px;">
  <p style="margin:0 0 24px 0;font-size:13px;color:#64748B;line-height:1.7;">Halo Tim Finance,<br>Berikut detail pengajuan yang perlu diproses. Mohon ditindaklanjuti sesuai due date.</p>
  <p style="margin:0 0 10px 0;font-size:9px;color:#F47920;font-weight:700;letter-spacing:3px;text-transform:uppercase;">Tujuan Transfer</p>
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#F8FAFD;border-radius:10px;border:1px solid #E0E8F0;margin-bottom:26px;">
    <tr><td style="padding:20px 22px;">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td width="50%" style="vertical-align:top;padding-bottom:16px;"><p style="margin:0 0 3px 0;font-size:9px;color:#94A3B8;letter-spacing:2px;text-transform:uppercase;">Bank</p><p style="margin:0;font-size:16px;color:#1E293B;font-weight:700;">{form_data.get("bank","")}</p></td>
          <td width="50%" style="vertical-align:top;padding-bottom:16px;"><p style="margin:0 0 3px 0;font-size:9px;color:#94A3B8;letter-spacing:2px;text-transform:uppercase;">Nama Rekening</p><p style="margin:0;font-size:13px;color:#334155;">{form_data.get("recipient_name","")}</p></td>
        </tr>
        <tr><td colspan="2" style="border-top:1px solid #E0E8F0;padding-top:16px;"><p style="margin:0 0 3px 0;font-size:9px;color:#94A3B8;letter-spacing:2px;text-transform:uppercase;">Nomor Rekening</p><p style="margin:0;font-size:20px;color:#1F3864;font-weight:700;letter-spacing:4px;font-family:'Courier New',monospace;">{form_data.get("account_number","")}</p></td></tr>
        <tr><td colspan="2" style="border-top:1px solid #E0E8F0;padding-top:12px;"><p style="margin:0 0 3px 0;font-size:9px;color:#94A3B8;letter-spacing:2px;text-transform:uppercase;">Due Date</p><p style="margin:0;font-size:13px;color:#F47920;font-weight:700;">{form_data.get("due_date_label","")}</p></td></tr>
      </table>
    </td></tr>
  </table>
  <p style="margin:0 0 10px 0;font-size:9px;color:#F47920;font-weight:700;letter-spacing:3px;text-transform:uppercase;">Detail Transaksi</p>
  <table width="100%" cellpadding="0" cellspacing="0" style="border-radius:10px;overflow:hidden;border:1px solid #E0E8F0;margin-bottom:26px;">
    <tr style="background:#1F3864;">
      <td style="padding:11px 16px;font-size:9px;font-weight:700;color:#7BA4D4;letter-spacing:2px;text-transform:uppercase;width:110px;">Tanggal</td>
      <td style="padding:11px 16px;font-size:9px;font-weight:700;color:#7BA4D4;letter-spacing:2px;text-transform:uppercase;">Deskripsi</td>
      <td style="padding:11px 16px;font-size:9px;font-weight:700;color:#7BA4D4;letter-spacing:2px;text-transform:uppercase;text-align:right;width:120px;">Jumlah</td>
    </tr>
    <tr style="background:#EEF3FA;"><td colspan="3" style="padding:8px 16px;font-size:11px;color:#64748B;font-style:italic;border-bottom:1px solid #E0E8F0;">{form_data.get("detail_descriptions","")}</td></tr>
    {item_rows}
    <tr style="background:#F1F5F9;"><td colspan="2" style="padding:11px 16px;font-size:12px;color:#64748B;border-top:1px solid #E0E8F0;">Total</td><td style="padding:11px 16px;font-size:12px;color:#64748B;text-align:right;border-top:1px solid #E0E8F0;font-family:'Courier New',monospace;">{fmt(form_data.get("subtotal",0))}</td></tr>
    <tr style="background:#F8FAFC;"><td colspan="2" style="padding:9px 16px;font-size:11px;color:#94A3B8;">Tax (if applicable)</td><td style="padding:9px 16px;font-size:11px;color:#94A3B8;text-align:right;font-family:'Courier New',monospace;">{fmt(form_data.get("tax",0))}</td></tr>
    <tr style="background:linear-gradient(90deg,#1F3864,#2E5BA8);"><td colspan="2" style="padding:14px 16px;font-size:13px;font-weight:700;color:#FFFFFF;">Grand Total</td><td style="padding:14px 16px;font-size:17px;font-weight:700;color:#F47920;text-align:right;font-family:'Courier New',monospace;">{fmt(form_data.get("grand_total",0))}</td></tr>
  </table>
</td></tr>
<tr><td style="padding:0 36px 32px 36px;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#F8FAFD;border-radius:10px;border:1px solid #E0E8F0;">
    <tr><td style="padding:20px 22px;">
      <p style="margin:0 0 14px 0;font-size:13px;color:#64748B;line-height:1.7;">Dokumen Finance Request Form terlampir sebagai referensi.<br>Mohon konfirmasi setelah transfer diproses. Terima kasih.</p>
      <p style="margin:0;font-size:13px;color:#475569;line-height:1.8;">Salam,<br><strong style="color:#1E293B;font-size:14px;">{yn}</strong><br><span style="color:#94A3B8;font-size:11px;">{dept} &nbsp;·&nbsp; PT Solusi Kerah Byru</span></p>
    </td></tr>
  </table>
</td></tr>
<tr><td style="padding:20px 36px;border-top:1px solid #E8EDF2;">
  <table width="100%" cellpadding="0" cellspacing="0"><tr>
    <td style="vertical-align:middle;">{logo_tag}<span style="font-size:10px;color:#CBD5E1;display:block;margin-top:4px;">byru.id</span></td>
    <td align="right" style="vertical-align:middle;"><span style="font-size:10px;color:#CBD5E1;line-height:1.7;display:block;text-align:right;">Email otomatis dari sistem Finance Byru.<br>Hubungi tim Finance untuk pertanyaan.</span></td>
  </tr></table>
</td></tr>
<tr><td style="height:4px;background:linear-gradient(90deg,#F47920 0%,#2E5BA8 50%,#1F3864 100%);font-size:0;line-height:0;">&nbsp;</td></tr>
</table></td></tr></table></body></html>"""

    # ── Submit buttons ───────────────────────────────────────────
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-header">🚀 Submit Pengajuan</div>', unsafe_allow_html=True)

    def validate_fr():
        errors = []
        if not your_name.strip():      errors.append("Nama Pemohon harus diisi")
        if not recipient_name.strip(): errors.append("Nama Rekening harus diisi")
        if not account_number.strip(): errors.append("Nomor Rekening harus diisi")
        if not any([type_unclaimable, type_claimable]):
            errors.append("Pilih minimal satu Type of Expenses")
        if not any([p_expense, p_reimbursement, p_medical, p_pettycash,
                    p_cashadvance, p_settlement, p_perkiraan]):
            errors.append("Pilih minimal satu Purpose")
        if len(items) == 0: errors.append("Isi minimal satu item transaksi")
        if not approver1_name.strip(): errors.append("Nama Approver 1 harus diisi")
        return errors

    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        submit_clicked = st.button("📤 Submit & Generate PDF", type="primary",
                                   use_container_width=True, key="fr_submit")
    with col_s2:
        preview_clicked = st.button("👁️ Preview Summary", use_container_width=True, key="fr_preview")
    with col_s3:
        email_clicked = st.button("✉️ Lihat Template Email", use_container_width=True, key="fr_email")

    if preview_clicked:
        st.markdown("### 📋 Summary")
        for k, v in {
            "Tanggal": req_date.strftime("%d/%m/%Y"), "Pemohon": your_name,
            "Bank": bank, "Rekening": f"{recipient_name} – {account_number}",
            "Due Date": due_date.strftime("%A, %d %B %Y"),
            "Grand Total": f"Rp {grand_total:,.0f}".replace(",", "."),
        }.items():
            st.write(f"**{k}:** {v}")
        if items:
            df_p = pd.DataFrame(items)
            df_p["amount"] = df_p["amount"].apply(lambda x: f"Rp {x:,.0f}".replace(",", "."))
            df_p.columns = ["Tanggal", "Deskripsi", "Jumlah"]
            st.dataframe(df_p, use_container_width=True)

    if email_clicked:
        errs = validate_fr()
        if errs:
            for e in errs: st.error(f"❌ {e}")
        else:
            fd = {
                "date": req_date.strftime("%d/%m/%Y"), "your_name": your_name,
                "cost_center": cost_center, "bank": bank,
                "recipient_name": recipient_name, "account_number": account_number,
                "due_date_label": due_date.strftime("%A, %d %B %Y"),
                "purpose": {"expense": p_expense, "reimbursement": p_reimbursement,
                            "medical": p_medical, "petty_cash": p_pettycash,
                            "cash_advance": p_cashadvance, "settlement": p_settlement,
                            "perkiraan": p_perkiraan},
                "detail_descriptions": detail_description, "items": items,
                "subtotal": subtotal, "tax": tax_amount, "grand_total": grand_total,
                "claimer_name": claimer_name, "claimer_dept": claimer_dept,
            }
            html_email = generate_email_html(fd, _logo_path)
            st.markdown('<div class="section-header">✉️ Template Email</div>', unsafe_allow_html=True)
            st.markdown('<div class="info-box">💡 Download file HTML, buka di browser → Select All → Copy → Paste ke Gmail/Outlook.</div>', unsafe_allow_html=True)
            with st.expander("👁️ Preview Email", expanded=True):
                components.html(html_email, height=780, scrolling=True)
            st.download_button("⬇️ Download Email HTML", data=html_email.encode("utf-8"),
                file_name=f"Email_Finance_{your_name.replace(' ','_')}_{req_date.strftime('%Y%m%d')}.html",
                mime="text/html", use_container_width=True, key="fr_dl_email")

    if submit_clicked:
        errs = validate_fr()
        if errs:
            for e in errs: st.error(f"❌ {e}")
        else:
            form_data = {
                "date": req_date.strftime("%d/%m/%Y"), "your_name": your_name,
                "cost_center": cost_center, "bank": bank,
                "recipient_name": recipient_name, "account_number": account_number,
                "due_date": due_date.strftime("%d/%m/%Y"),
                "due_date_label": due_date.strftime("%A, %d %B %Y"),
                "type_unclaimable": type_unclaimable, "type_claimable": type_claimable,
                "pic_name": pic_name,
                "purpose": {"expense": p_expense, "reimbursement": p_reimbursement,
                            "medical": p_medical, "petty_cash": p_pettycash,
                            "cash_advance": p_cashadvance, "settlement": p_settlement,
                            "perkiraan": p_perkiraan, "margin_pct": margin_pct},
                "detail_descriptions": detail_description, "items": items,
                "subtotal": subtotal, "tax": tax_amount, "grand_total": grand_total,
                "claimer_name": claimer_name, "claimer_dept": claimer_dept,
                "approver1_name": approver1_name, "approver1_pos": approver1_pos,
                "approver2_name": approver2_name, "approver2_pos": approver2_pos,
                "needs_director": needs_director,
                "director_name": "Nathaniel Nugroho Liman", "director_pos": "Director",
            }
            with st.spinner("Generating PDF..."):
                try:
                    pdf_bytes = generate_pdf(form_data)
                    fname = f"Finance_Request_{your_name.replace(' ','_')}_{req_date.strftime('%Y%m%d')}.pdf"
                    st.success("✅ PDF berhasil dibuat!")
                    st.download_button("⬇️ Download PDF", data=pdf_bytes,
                        file_name=fname, mime="application/pdf",
                        use_container_width=True, key="fr_dl_pdf")
                except Exception as ex:
                    st.error(f"Gagal generate PDF: {ex}")


# ╔══════════════════════════════════════════════════════════════╗
# ║  TAB 2 — PERJALANAN DINAS                                   ║
# ╚══════════════════════════════════════════════════════════════╝
with tab_perdin:

    st.markdown('<div class="section-header">✈️ Perjalanan Dinas</div>', unsafe_allow_html=True)

    tipe_perdin = st.radio(
        "Tipe Form",
        ["Form Perdin (Pengajuan & Realisasi)", "Cash Advance Perdin", "Reimburse Perdin"],
        horizontal=True, key="pd_tipe"
    )

    # ── INFO: CA / Reimburse langsung ke Finance Request ─────────
    if tipe_perdin in ["Cash Advance Perdin", "Reimburse Perdin"]:
        st.markdown(f"""
        <div class="info-box">
        ℹ️ <b>{tipe_perdin}</b> menggunakan <b>Finance Request Form</b> (Tab 💰 di atas).<br>
        Cukup buka tab <b>💰 Finance Request</b>, pilih Purpose sesuai tipe:
        {'<b>Cash Advance</b>' if 'Cash Advance' in tipe_perdin else '<b>Reimbursement</b>'}.
        <br><br>
        Lampiran struk/bon bisa di-upload di bagian bawah halaman ini.
        </div>
        """, unsafe_allow_html=True)

    # ── FORM PERDIN ───────────────────────────────────────────────
    if tipe_perdin == "Form Perdin (Pengajuan & Realisasi)":

        # ── Identity ─────────────────────────────────────────────
        st.markdown('<div class="section-header">👤 Identitas Karyawan</div>', unsafe_allow_html=True)
        col1, col2, col3, col4 = st.columns(4)
        with col1: pd_no_urut  = st.text_input("No. Urut Perjalanan Dinas", key="pd_no")
        with col2: pd_name     = st.text_input("Nama Karyawan", key="pd_name")
        with col3: pd_dept     = st.text_input("Divisi", key="pd_dept")
        with col4: pd_jabatan  = st.text_input("Jabatan", key="pd_jabatan")

        # ── Pengajuan ─────────────────────────────────────────────
        st.markdown('<div class="section-header">📋 Pengajuan Perjalanan Dinas</div>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            pd_jenis = st.radio("Jenis Perjalanan", ["Dalam Kota", "Luar Kota"], horizontal=True, key="pd_jenis")
            pd_kota  = st.text_input("Lokasi / Kota Tujuan", key="pd_kota")
            pd_days  = st.number_input("Jumlah Hari", min_value=1, max_value=30, value=1, key="pd_days")
        with col2:
            pd_dep_date = st.date_input("Tanggal Keberangkatan", value=date.today(), key="pd_dep")
            pd_ret_date = st.date_input("Tanggal Kembali", value=date.today(), key="pd_ret")
            pd_purpose  = st.text_area("Keperluan Perjalanan Dinas", height=80, key="pd_purpose")

        st.markdown("**Jenis Transportasi**")
        tc = st.columns(6)
        t_mobil_ops    = tc[0].checkbox("Mobil Operasional", key="pd_t1")
        t_mobil_pribadi= tc[1].checkbox("Mobil Pribadi", key="pd_t2")
        t_motor        = tc[2].checkbox("Motor", key="pd_t3")
        t_pesawat      = tc[3].checkbox("Pesawat", key="pd_t4")
        t_kereta       = tc[4].checkbox("Kereta", key="pd_t5")
        t_umum         = tc[5].checkbox("Umum", key="pd_t6")

        pd_uang_muka = st.number_input("Uang Muka (Rp)", min_value=0, step=10000, value=0, key="pd_um")

        # ── Realisasi ─────────────────────────────────────────────
        st.markdown('<div class="section-header">✅ Realisasi Perjalanan Dinas</div>', unsafe_allow_html=True)
        pd_hasil = st.text_area("Hasil Perjalanan Dinas", height=70, key="pd_hasil")

        # ── Expense Tables ────────────────────────────────────────
        st.markdown('<div class="section-header">💸 Detail Pengeluaran</div>', unsafe_allow_html=True)
        st.markdown(
            "Nama pos pengeluaran **bisa diedit**. "
            "BBM/Toll/Parkir dicatat sebagai *SESUAI SK / ON BILL* di PDF."
        )

        # Init session state untuk nama-nama expense yang bisa diedit
        if "pd_expense_labels" not in st.session_state:
            st.session_state["pd_expense_labels"] = [
                lbl for lbl, _, _ in EXPENSE_DEFAULTS
            ]

        # Row header: label editor
        st.markdown("**✏️ Edit Nama Pos Pengeluaran (opsional):**")
        label_cols = st.columns(3)
        for i, default_lbl in enumerate([d[0] for d in EXPENSE_DEFAULTS]):
            col_idx = i % 3
            new_lbl = label_cols[col_idx].text_input(
                f"Pos {i+1}",
                value=st.session_state["pd_expense_labels"][i],
                key=f"pd_elabel_{i}",
                label_visibility="collapsed",
                placeholder=default_lbl,
            )
            st.session_state["pd_expense_labels"][i] = new_lbl or default_lbl

        st.markdown("---")
        col_pengajuan, col_realisasi = st.columns(2)
        expenses_p = {}
        expenses_r = {}

        with col_pengajuan:
            st.markdown("**Pengajuan (Rp)**")
            for i, (_, key, is_sk) in enumerate(EXPENSE_DEFAULTS):
                display_label = st.session_state["pd_expense_labels"][i]
                note = " *(SESUAI SK / ON BILL)*" if is_sk else ""
                expenses_p[key] = st.number_input(
                    f"{display_label}{note}",
                    min_value=0, step=5000,
                    key=f"pd_ep_{key}",
                )

        with col_realisasi:
            st.markdown("**Realisasi (Rp)**")
            for i, (_, key, is_sk) in enumerate(EXPENSE_DEFAULTS):
                display_label = st.session_state["pd_expense_labels"][i]
                note = " *(SESUAI SK / ON BILL)*" if is_sk else ""
                expenses_r[key] = st.number_input(
                    f"{display_label}{note}",
                    min_value=0, step=5000,
                    key=f"pd_er_{key}",
                )

        total_p = sum(expenses_p.values())
        total_r = sum(expenses_r.values())
        mc = st.columns(2)
        mc[0].metric("Total Pengajuan", f"Rp {total_p:,.0f}".replace(",", "."))
        mc[1].metric("Total Realisasi", f"Rp {total_r:,.0f}".replace(",", "."))

        # ── Approvers ─────────────────────────────────────────────
        st.markdown('<div class="section-header">✍️ Persetujuan</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="info-box" style="font-size:0.82em;">
        📋 <b>Sisi Kiri PDF</b> (Pengajuan): Karyawan ybs → Disetujui (1) → Disetujui (2)<br>
        📋 <b>Sisi Kanan PDF</b> (Realisasi): Disetujui (3) → Mengetahui (1) → Mengetahui (2)
        </div>""", unsafe_allow_html=True)
        ac = st.columns(3)
        pd_app1 = ac[0].text_input("Disetujui (1) — sisi kiri", key="pd_a1")
        pd_app2 = ac[1].text_input("Disetujui (2) — sisi kiri", key="pd_a2")
        pd_app3 = ac[2].text_input("Disetujui (3) — sisi kanan", key="pd_a3")
        ac2 = st.columns(3)
        pd_app4 = ac2[0].text_input("Mengetahui (1) — sisi kanan", key="pd_a4")
        pd_app5 = ac2[1].text_input("Mengetahui (2) — sisi kanan (opsional)", key="pd_a5")
        ac2[2].markdown("")  # spacer

        # ── Generate PDF Perdin ───────────────────────────────────
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("📄 Generate Form Perdin PDF", type="primary",
                     use_container_width=True, key="pd_submit"):
            transport = []
            if t_mobil_ops:     transport.append("mobil_ops")
            if t_mobil_pribadi: transport.append("mobil_pribadi")
            if t_motor:         transport.append("motor")
            if t_pesawat:       transport.append("pesawat")
            if t_kereta:        transport.append("kereta")
            if t_umum:          transport.append("umum")

            perdin_data = {
                "no_urut": pd_no_urut, "your_name": pd_name,
                "departement": pd_dept, "jabatan": pd_jabatan,
                "jenis_perjalanan": "dalam_kota" if pd_jenis == "Dalam Kota" else "luar_kota",
                "kota_tujuan": pd_kota, "days_no": pd_days,
                "departure_date": pd_dep_date.strftime("%d/%m/%Y"),
                "return_date": pd_ret_date.strftime("%d/%m/%Y"),
                "purpose_trip": pd_purpose,
                "jenis_transportasi": transport,
                "uang_muka": pd_uang_muka,
                "hasil_perjalanan": pd_hasil,
                "expenses_pengajuan": expenses_p,
                "expenses_realisasi": expenses_r,
                "expense_labels": st.session_state.get("pd_expense_labels",
                    [d[0] for d in EXPENSE_DEFAULTS]),
                "expense_is_sk": [d[2] for d in EXPENSE_DEFAULTS],
                "approvers": [pd_app1, pd_app2, pd_app3, pd_app4, pd_app5],
            }
            with st.spinner("Generating PDF Form Perdin..."):
                try:
                    pdf_bytes = generate_perdin_pdf(perdin_data, _logo_path)
                    fname = f"Perdin_{pd_name.replace(' ','_')}_{pd_dep_date.strftime('%Y%m%d')}.pdf"
                    st.success("✅ PDF Form Perdin berhasil dibuat!")
                    st.download_button("⬇️ Download PDF Perdin", data=pdf_bytes,
                        file_name=fname, mime="application/pdf",
                        use_container_width=True, key="pd_dl")
                    # Save to session state for lampiran merge
                    st.session_state["last_perdin_data"] = perdin_data
                except Exception as ex:
                    st.error(f"Gagal generate PDF: {ex}")
                    import traceback
                    st.code(traceback.format_exc())

    # ── LAMPIRAN SECTION (semua tipe) ─────────────────────────────
    st.markdown('<div class="section-header">📎 Lampiran (Struk / Foto / Bukti)</div>', unsafe_allow_html=True)
    st.markdown("""
    <div class="info-box">
    📌 Upload foto struk/bukti. Setiap halaman PDF akan memuat <b>4 lampiran</b> (2×2 grid).<br>
    Isi deskripsi untuk masing-masing foto.
    </div>""", unsafe_allow_html=True)

    n_lampiran = st.number_input("Jumlah Lampiran", min_value=1, max_value=20, value=4, step=1, key="pd_nlamp")
    attachments = []

    for i in range(int(n_lampiran)):
        with st.expander(f"📷 Lampiran {i+1}", expanded=(i < 4)):
            lc1, lc2 = st.columns([1, 2])
            with lc1:
                uploaded = st.file_uploader(f"Upload foto #{i+1}",
                    type=["jpg", "jpeg", "png", "webp"],
                    key=f"pd_lamp_file_{i}")
                if uploaded:
                    # Read bytes FIRST, then display — avoid pointer exhaustion
                    img_bytes = uploaded.read()
                    st.image(img_bytes, use_container_width=True)
                else:
                    img_bytes = None
            with lc2:
                label = st.text_input(f"Label #{i+1}", value=f"Lampiran {i+1}",
                                      key=f"pd_lamp_label_{i}")
                desc  = st.text_area(f"Deskripsi #{i+1}",
                    placeholder="Contoh: Struk BBM tanggal 09/03/2026, SPBU Cawang",
                    height=100, key=f"pd_lamp_desc_{i}")
            attachments.append({
                "image_bytes": img_bytes,
                "label": label,
                "description": desc,
            })

    if st.button("📎 Generate PDF Lampiran", type="primary",
                 use_container_width=True, key="pd_lamp_submit"):
        perdin_ctx = st.session_state.get("last_perdin_data", {})
        with st.spinner("Generating PDF Lampiran..."):
            try:
                lamp_bytes = generate_lampiran_pdf(attachments, perdin_ctx)
                pname = perdin_ctx.get("your_name", "Karyawan")
                fname = f"Lampiran_Perdin_{pname.replace(' ','_')}.pdf"
                st.success("✅ PDF Lampiran berhasil dibuat!")
                st.download_button("⬇️ Download PDF Lampiran", data=lamp_bytes,
                    file_name=fname, mime="application/pdf",
                    use_container_width=True, key="pd_lamp_dl")
            except Exception as ex:
                st.error(f"Gagal generate Lampiran: {ex}")
                import traceback
                st.code(traceback.format_exc())
