import streamlit as st
import pytz
import openpyxl
import io
import os
from datetime import datetime
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, PageBreak, HRFlowable, KeepTogether
)

MONTHS = {
    1:"January", 2:"February", 3:"March", 4:"April",
    5:"May",     6:"June",     7:"July",  8:"August",
    9:"September",10:"October",11:"November",12:"December"
}

DARK  = colors.HexColor("#1a1a2e")
GBGD  = colors.HexColor("#f5f5f5")
GBRD  = colors.HexColor("#e8e8e8")
MID   = colors.HexColor("#666670")
RED   = colors.HexColor("#a32d2d")
GREEN = colors.HexColor("#0f6e56")


# ─── HELPERS ────────────────────────────────

def _fmt(v):
    try:    return f"${float(v):,.2f}"
    except: return ""

def _fecha(r):
    try:    return f"{int(r['month']):02d}/{int(r['day']):02d}/{r['year']}"
    except: return ""

def _prop(p):
    return p.replace("Property 1: ","").replace("Property 2: ","")

def fmt_local(r):
    try:
        v = float(r["local"] or 0)
        cur = r["cur"].strip()
        return f"${v:,.2f}" if cur == "USD" else f"{v:,.0f} {cur}"
    except:
        return ""

def extract_name(filename):
    """Extract person name from filename e.g. Cash_Ledger_Leonard.xlsx -> Leonard"""
    name = os.path.splitext(filename)[0]
    parts = name.replace("-","_").split("_")
    # Return last meaningful part (capitalized)
    for part in reversed(parts):
        if len(part) > 2 and part.lower() not in ("cash","ledger","v1","copia","cuba","lcc"):
            return part.capitalize()
    return name.capitalize()


# ─── LECTURA ────────────────────────────────

def leer_excel(file):
    wb   = openpyxl.load_workbook(file, data_only=True)
    ws   = next((wb[s] for s in wb.sheetnames if "cash" in s.lower()), wb.active)
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("The file appears to be empty.")

    hdrs = [str(h or "").strip().lower() for h in rows[0]]

    def col(name):
        for i, h in enumerate(hdrs):
            if name in h:
                return i
        return None

    ci = dict(
        year=col("year"), month=col("month"), day=col("day"),
        desc=col("description"), local=col("local amount"),
        usd=col("usd amount"), bal=col("running balance"),
        cat=col("category"), cur=col("currency"), prop=col("property"),
    )

    data = []
    for r in rows[1:]:
        if ci["year"] is None or ci["month"] is None:
            break
        if not r[ci["year"]] or not r[ci["month"]]:
            continue
        try:
            mes  = int(r[ci["month"]])
            anio = int(r[ci["year"]])
        except:
            continue
        data.append({
            "year":  anio, "month": mes,
            "day":   r[ci["day"]]   if ci["day"]   is not None else "",
            "desc":  str(r[ci["desc"]]  or "").strip() if ci["desc"]  is not None else "",
            "local": r[ci["local"]] if ci["local"] is not None else None,
            "usd":   r[ci["usd"]]   if ci["usd"]   is not None else None,
            "bal":   r[ci["bal"]]   if ci["bal"]   is not None else None,
            "cat":   str(r[ci["cat"]]   or "").strip() if ci["cat"]   is not None else "",
            "cur":   str(r[ci["cur"]]   or "").strip() if ci["cur"]   is not None else "",
            "prop":  str(r[ci["prop"]]  or "").strip() if ci["prop"]  is not None else "",
        })
    return data


def agrupar(data):
    g = {}
    for r in data:
        k = (r["year"], r["month"])
        g.setdefault(k, []).append(r)
    return dict(sorted(g.items()))


# ─── TRANSACTION TABLE BUILDER ───────────────

def tx_table(rows):
    heads = ["Date", "Prop.", "Category", "Description", "Cur.", "Local Amount", "USD"]
    col_w = [0.7*inch, 0.55*inch, 1.0*inch, 2.9*inch, 0.38*inch, 0.8*inch, 0.67*inch]

    data = [heads] + [[
        _fecha(r),
        _prop(r["prop"])[:6],
        r["cat"][:16],
        r["desc"][:60],
        r["cur"],
        fmt_local(r),
        f'${float(r["usd"] or 0):,.2f}',
    ] for r in rows]

    t = Table(data, colWidths=col_w, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,0),  DARK),
        ("TEXTCOLOR",     (0,0),(-1,0),  colors.white),
        ("FONTNAME",      (0,0),(-1,0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0),(-1,0),  7.5),
        ("ALIGN",         (0,0),(-1,0),  "CENTER"),
        ("FONTNAME",      (0,1),(-1,-1), "Helvetica"),
        ("FONTSIZE",      (0,1),(-1,-1), 7.5),
        ("ROWBACKGROUNDS",(0,1),(-1,-1), [colors.white, GBGD]),
        ("TEXTCOLOR",     (5,1),(5,-1),  MID),
        ("ALIGN",         (5,1),(5,-1),  "RIGHT"),
        ("TEXTCOLOR",     (6,1),(6,-1),  RED),
        ("FONTNAME",      (6,1),(6,-1),  "Helvetica-Bold"),
        ("ALIGN",         (6,1),(6,-1),  "RIGHT"),
        ("BACKGROUND",    (6,1),(6,-1),  colors.HexColor("#fff5f5")),
        ("LINEAFTER",     (5,0),(5,-1),  0.8, colors.HexColor("#aaaaaa")),
        ("ALIGN",         (4,1),(4,-1),  "CENTER"),
        ("GRID",          (0,0),(-1,-1), 0.3, GBRD),
        ("TOPPADDING",    (0,0),(-1,-1), 3),
        ("BOTTOMPADDING", (0,0),(-1,-1), 3),
        ("LEFTPADDING",   (0,0),(-1,-1), 4),
        ("RIGHTPADDING",  (0,0),(-1,-1), 4),
    ]))
    return t


def subtotal_row(name, rows, styles):
    total = sum(float(r["usd"] or 0) for r in rows)
    sm = ParagraphStyle("st", parent=styles["Normal"], fontSize=8,
                        fontName="Helvetica-Bold", textColor=RED)
    tbl = Table([[
        Paragraph(f"Subtotal {name}:", sm),
        Paragraph(f"${total:,.2f}", sm),
    ]], colWidths=[5.5*inch, 1.5*inch])
    tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), colors.HexColor("#fff5f5")),
        ("ALIGN",         (1,0),(1,-1),  "RIGHT"),
        ("TOPPADDING",    (0,0),(-1,-1), 4),
        ("BOTTOMPADDING", (0,0),(-1,-1), 4),
        ("LEFTPADDING",   (0,0),(0,-1),  8),
        ("RIGHTPADDING",  (-1,0),(-1,-1),8),
        ("LINEABOVE",     (0,0),(-1,0),  0.5, RED),
    ]))
    return tbl


# ─── GENERADOR PDF ──────────────────────────

def generar_pdf(sources, year, month):
    """
    sources: list of (name, rows)
    """
    mes_str = MONTHS.get(month, str(month))
    periodo = f"{mes_str} {year}"
    all_rows = [r for _, rows in sources for r in rows]
    grand_total = sum(float(r["usd"] or 0) for r in all_rows)

    styles = getSampleStyleSheet()
    def S(name, parent="Normal", **kw):
        return ParagraphStyle(name, parent=styles[parent], **kw)

    t_s  = S("t",  fontSize=15, fontName="Helvetica-Bold", textColor=colors.white)
    sb_s = S("sb", fontSize=9,  textColor=colors.HexColor("#c0c0cc"))
    h_s  = S("h",  fontSize=8.5,fontName="Helvetica-Bold", textColor=MID, spaceBefore=12, spaceAfter=5)
    sec_s= S("sec",fontSize=10, fontName="Helvetica-Bold", textColor=colors.white)
    b_s  = S("b",  fontSize=10, textColor=colors.HexColor("#222222"), leading=16)
    sm_s = S("sm", fontSize=7.5,textColor=MID)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter,
        leftMargin=0.65*inch, rightMargin=0.65*inch,
        topMargin=0.6*inch,   bottomMargin=0.6*inch)
    story = []

    def banner(title, sub, w1="60%", w2="40%"):
        tbl = Table([[Paragraph(title, t_s), Paragraph(sub, sb_s)]], colWidths=[w1, w2])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), DARK),
            ("TOPPADDING",   (0,0),(-1,-1), 10),
            ("BOTTOMPADDING",(0,0),(-1,-1), 10),
            ("LEFTPADDING",  (0,0),(0,-1),  12),
            ("RIGHTPADDING", (-1,0),(-1,-1),12),
            ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
            ("ALIGN",        (1,0),(1,-1),  "RIGHT"),
        ]))
        return tbl

    def person_banner(name, n_tx, subtotal):
        tbl = Table([[
            Paragraph(name.upper(), sec_s),
            Paragraph(f"{n_tx} transactions  |  Subtotal: ${subtotal:,.2f}", sb_s),
        ]], colWidths=["35%","65%"])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), colors.HexColor("#2d2d4e")),
            ("TOPPADDING",   (0,0),(-1,-1), 7),
            ("BOTTOMPADDING",(0,0),(-1,-1), 7),
            ("LEFTPADDING",  (0,0),(0,-1),  12),
            ("RIGHTPADDING", (-1,0),(-1,-1),12),
            ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
            ("ALIGN",        (1,0),(1,-1),  "RIGHT"),
        ]))
        return tbl

    # ── PAGE 1: LEDGER ──────────────────────
    story.append(banner("LCC — Cash Ledger (Cuba)", f"Period: {periodo}"))
    story.append(Spacer(1, 8))

    # Summary strip
    names_str = " + ".join(name for name, _ in sources)
    strip_data = [
        [Paragraph("Total Expenses (USD)", sm_s),
         Paragraph("Transactions", sm_s),
         Paragraph("Sources", sm_s)],
        [Paragraph(f'<font color="#a32d2d"><b>{_fmt(grand_total)}</b></font>', styles["Normal"]),
         Paragraph(f"<b>{len(all_rows)}</b>", styles["Normal"]),
         Paragraph(f"<b>{names_str}</b>", styles["Normal"])],
    ]
    strip = Table(strip_data, colWidths=["33%","20%","47%"])
    strip.setStyle(TableStyle([
        ("BACKGROUND",   (0,0),(-1,-1), GBGD),
        ("GRID",         (0,0),(-1,-1), 0.3, GBRD),
        ("TOPPADDING",   (0,0),(-1,-1), 6),
        ("BOTTOMPADDING",(0,0),(-1,-1), 6),
        ("LEFTPADDING",  (0,0),(-1,-1), 10),
        ("FONTSIZE",     (0,1),(-1,-1), 11),
    ]))
    story.append(strip)

    # Per-person sections
    for name, rows in sources:
        subtotal = sum(float(r["usd"] or 0) for r in rows)
        story.append(Spacer(1, 12))
        story.append(person_banner(name, len(rows), subtotal))
        story.append(Spacer(1, 4))
        story.append(tx_table(rows))
        story.append(subtotal_row(name, rows, styles))

    # Grand total
    story.append(Spacer(1, 8))
    gt = Table([[
        Paragraph("GRAND TOTAL:", ParagraphStyle("gt", parent=styles["Normal"],
                  fontSize=10, fontName="Helvetica-Bold", textColor=colors.white)),
        Paragraph(f"${grand_total:,.2f}", ParagraphStyle("gtv", parent=styles["Normal"],
                  fontSize=11, fontName="Helvetica-Bold", textColor=colors.white)),
    ]], colWidths=[5.5*inch, 1.5*inch])
    gt.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), DARK),
        ("ALIGN",         (1,0),(1,-1),  "RIGHT"),
        ("TOPPADDING",    (0,0),(-1,-1), 6),
        ("BOTTOMPADDING", (0,0),(-1,-1), 6),
        ("LEFTPADDING",   (0,0),(0,-1),  8),
        ("RIGHTPADDING",  (-1,0),(-1,-1),8),
    ]))
    story.append(gt)

    # ── PAGE 2: APPROVAL ────────────────────
    story.append(PageBreak())
    story.append(banner("LCC — Cash Ledger Approval", f"Period: {periodo}", "65%", "35%"))
    story.append(Spacer(1, 20))
    story.append(Paragraph("SUMMARY", h_s))

    recap_data = [["Period", periodo], ["Sources", names_str]]
    for name, rows in sources:
        sub = sum(float(r["usd"] or 0) for r in rows)
        recap_data.append([f"  Subtotal {name}", _fmt(sub)])
    recap_data.append(["Grand Total (USD)", _fmt(grand_total)])
    recap_data.append(["No. of Transactions", str(len(all_rows))])

    recap = Table(recap_data, colWidths=[3*inch, 2.5*inch])
    recap_style = [
        ("FONTNAME",      (0,0),(-1,-1), "Helvetica"),
        ("FONTNAME",      (0,0),(0,-1),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0),(-1,-1), 10),
        ("ALIGN",         (1,0),(1,-1),  "RIGHT"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LINEABOVE",     (0, len(recap_data)-2), (-1, len(recap_data)-2), 1, DARK),
        ("FONTNAME",      (0, len(recap_data)-2), (-1,-1), "Helvetica-Bold"),
        ("TEXTCOLOR",     (1, len(recap_data)-2), (1,-1), RED),
        ("LINEBELOW",     (0,-1),(-1,-1), 1, DARK),
    ]
    recap.setStyle(TableStyle(recap_style))
    story.append(recap)
    story.append(Spacer(1, 28))

    story.append(Paragraph("APPROVAL", h_s))
    story.append(Paragraph(
        f"Reviewed and approved. I confirm the transactions recorded in this Cash Ledger "
        f"for <b>{periodo}</b> are accurate and complete.",
        b_s))
    story.append(Spacer(1, 36))

    sig = Table([
        [Paragraph("<b>Approved by:</b>", b_s), ""],
        [Spacer(1,6), ""],
        [Paragraph("Name:&nbsp;&nbsp;&nbsp;<b>Rolando</b>", b_s), ""],
        [Spacer(1,14), ""],
        [Paragraph("Signature:", b_s), ""],
        [Spacer(1,22), ""],
        [Paragraph("Date:", b_s), ""],
    ], colWidths=[3.5*inch, 3*inch])
    sig.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), GBGD),
        ("BOX",           (0,0),(-1,-1), 0.5, GBRD),
        ("LEFTPADDING",   (0,0),(-1,-1), 16),
        ("RIGHTPADDING",  (0,0),(-1,-1), 16),
        ("TOPPADDING",    (0,0),(1,0),   12),
        ("BOTTOMPADDING", (0,-1),(1,-1), 14),
        ("TOPPADDING",    (0,1),(-1,-1), 2),
        ("BOTTOMPADDING", (0,1),(-1,-1), 2),
        ("LINEBELOW",     (0,4),(0,4),   1, DARK),
        ("LINEBELOW",     (0,6),(0,6),   1, DARK),
    ]))
    story.append(sig)
    story.append(Spacer(1, 40))
    story.append(HRFlowable(width="100%", thickness=0.5, color=GBRD, spaceAfter=6))
    story.append(Paragraph(
        f"Generated: {datetime.now().strftime('%B %d, %Y')}", sm_s))

    doc.build(story)
    buf.seek(0)
    return buf


# ─── GENERADOR EXCEL ────────────────────────

def generar_excel(sources, year, month):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    mes_str = MONTHS.get(month, str(month))
    periodo = f"{mes_str} {year}"
    all_rows = [r for _, rows in sources for r in rows]
    grand_total = sum(float(r["usd"] or 0) for r in all_rows)

    wb = Workbook()

    dark_fill  = PatternFill("solid", fgColor="1A1A2E")
    sec_fill   = PatternFill("solid", fgColor="2D2D4E")
    gray_fill  = PatternFill("solid", fgColor="F5F5F5")
    gray2_fill = PatternFill("solid", fgColor="EBEBEB")
    usd_fill   = PatternFill("solid", fgColor="FFF0F0")
    sub_fill   = PatternFill("solid", fgColor="FFE8E8")
    total_fill = PatternFill("solid", fgColor="1A1A2E")
    white_fill = PatternFill("solid", fgColor="FFFFFF")

    white_bold  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    normal      = Font(name="Arial", size=9)
    mid_font    = Font(name="Arial", size=9, color="888888")
    usd_font    = Font(name="Arial", bold=True, size=9, color="A32D2D")
    sub_font    = Font(name="Arial", bold=True, size=9, color="A32D2D")
    total_font  = Font(name="Arial", bold=True, size=11, color="FFFFFF")
    sec_font    = Font(name="Arial", bold=True, size=10, color="FFFFFF")

    thin  = Side(style="thin",   color="DDDDDD")
    med   = Side(style="medium", color="AAAAAA")
    red_s = Side(style="medium", color="A32D2D")
    bdr   = Border(left=thin, right=thin, top=thin, bottom=thin)
    usd_bdr = Border(left=med, right=thin, top=thin, bottom=thin)

    center = Alignment(horizontal="center", vertical="center")
    left   = Alignment(horizontal="left",   vertical="center")
    right  = Alignment(horizontal="right",  vertical="center")

    num_usd   = "$#,##0.00"
    num_local = "###,##0.00"
    col_widths = [12, 10, 20, 48, 10, 16, 13]
    heads = ["Date", "Property", "Category", "Description", "Currency", "Local Amount", "USD"]

    # ── Sheet per person ──────────────────────
    first = True
    for name, rows in sources:
        ws = wb.active if first else wb.create_sheet()
        ws.title = f"{name}"
        first = False

        subtotal = sum(float(r["usd"] or 0) for r in rows)

        # Title
        ws.merge_cells("A1:G1")
        ws["A1"] = f"LCC — Cash Ledger (Cuba)   |   Period: {periodo}   |   {name}"
        ws["A1"].font = white_bold
        ws["A1"].fill = dark_fill
        ws["A1"].alignment = center
        ws.row_dimensions[1].height = 22

        ws.merge_cells("A2:G2")
        ws["A2"] = "Expense review — for Rolando's approval"
        ws["A2"].font = Font(name="Arial", size=9, color="888888", italic=True)
        ws["A2"].alignment = center
        ws.row_dimensions[2].height = 16

        for i, h in enumerate(heads, 1):
            c = ws.cell(row=3, column=i, value=h)
            c.font = white_bold; c.fill = dark_fill
            c.alignment = center; c.border = bdr
            ws.column_dimensions[get_column_letter(i)].width = col_widths[i-1]
        ws.row_dimensions[3].height = 18

        for idx, r in enumerate(rows):
            row_num = idx + 4
            is_alt  = idx % 2 == 1
            bg = gray_fill if is_alt else white_fill

            def cell(col, val, font=normal, fill=None, align=left, fmt=None, border=bdr):
                c = ws.cell(row=row_num, column=col, value=val)
                c.font = font; c.fill = fill or bg
                c.alignment = align; c.border = border
                if fmt: c.number_format = fmt
                return c

            try: cell(1, f"{int(r['month']):02d}/{int(r['day']):02d}/{r['year']}", align=center)
            except: cell(1, "")
            cell(2, _prop(r["prop"]))
            cell(3, r["cat"])
            cell(4, r["desc"])
            cell(5, r["cur"], align=center)
            try:
                cell(6, float(r["local"] or 0), font=mid_font, align=right, fmt=num_local)
            except: cell(6, "")
            try:
                c = ws.cell(row=row_num, column=7, value=float(r["usd"] or 0))
                c.font = usd_font; c.fill = usd_fill
                c.alignment = right; c.number_format = num_usd; c.border = usd_bdr
            except: ws.cell(row=row_num, column=7, value="")
            ws.row_dimensions[row_num].height = 15

        # Subtotal row
        sub_row = len(rows) + 4
        ws.row_dimensions[sub_row].height = 18
        for col in range(1, 7):
            c = ws.cell(row=sub_row, column=col, value="")
            c.fill = sub_fill; c.border = bdr
        ws.cell(row=sub_row, column=1, value=f"SUBTOTAL {name.upper()}").font = sub_font
        ws.cell(row=sub_row, column=1).fill = sub_fill
        ws.cell(row=sub_row, column=1).alignment = center
        c = ws.cell(row=sub_row, column=7, value=f"=SUM(G4:G{sub_row-1})")
        c.font = sub_font; c.fill = sub_fill
        c.alignment = right; c.number_format = num_usd
        c.border = Border(left=med, right=thin, top=red_s, bottom=red_s)
        ws.freeze_panes = "A4"

    # ── Consolidated sheet ────────────────────
    wc = wb.create_sheet("Consolidated")
    wc.column_dimensions["A"].width = 28
    wc.column_dimensions["B"].width = 22

    wc.merge_cells("A1:B1")
    c = wc.cell(row=1, column=1, value=f"LCC — Cash Ledger Consolidated   |   {periodo}")
    c.font = white_bold; c.fill = dark_fill; c.alignment = center
    wc.row_dimensions[1].height = 28

    row = 3
    for name, rows in sources:
        sub = sum(float(r["usd"] or 0) for r in rows)
        wc.row_dimensions[row].height = 18
        c1 = wc.cell(row=row, column=1, value=name)
        c1.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        c1.fill = sec_fill; c1.alignment = left
        c2 = wc.cell(row=row, column=2, value=sub)
        c2.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        c2.fill = sec_fill; c2.alignment = right; c2.number_format = num_usd
        row += 1

    wc.row_dimensions[row].height = 20
    c1 = wc.cell(row=row, column=1, value="GRAND TOTAL")
    c1.font = total_font; c1.fill = total_fill; c1.alignment = left
    c2 = wc.cell(row=row, column=2, value=grand_total)
    c2.font = total_font; c2.fill = total_fill
    c2.alignment = right; c2.number_format = num_usd
    row += 2

    # Approval block
    wc.merge_cells(f"A{row}:B{row}")
    c = wc.cell(row=row, column=1,
        value=f"Reviewed and approved. I confirm the transactions recorded in this Cash Ledger for {periodo} are accurate and complete.")
    c.font = Font(name="Arial", size=10, italic=True, color="333333")
    c.alignment = Alignment(wrap_text=True, vertical="center")
    wc.row_dimensions[row].height = 32
    row += 2

    for label, val in [("Approved by:", ""), ("Name:", "Rolando"), ("Signature:", ""), ("Date:", "")]:
        wc.row_dimensions[row].height = 22
        wc.cell(row=row, column=1, value=label).font = Font(name="Arial", bold=True, size=10)
        wc.cell(row=row, column=1).fill = gray2_fill
        c = wc.cell(row=row, column=2, value=val)
        c.font = Font(name="Arial", size=10); c.fill = gray2_fill
        if label in ("Signature:", "Date:"):
            c.border = Border(bottom=Side(style="medium", color="1A1A2E"))
        row += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── STREAMLIT UI ───────────────────────────

st.set_page_config(page_title="Cash Ledger LCC", page_icon="📊", layout="centered")
st.markdown("<style>.block-container { max-width: 780px; }</style>", unsafe_allow_html=True)

st.markdown("## 📊 Cash Ledger — LCC Cuba")
st.markdown("Upload one or more Excel files, select a month, and download the consolidated PDF or Excel for Rolando's review.")
st.divider()

uploaded_files = st.file_uploader(
    "Upload Cash Ledger files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    # Load all files
    all_sources = {}  # name -> data
    errors = []
    for f in uploaded_files:
        try:
            name = extract_name(f.name)
            data = leer_excel(f)
            all_sources[name] = agrupar(data)
        except Exception as e:
            errors.append(f"{f.name}: {e}")

    if errors:
        for e in errors:
            st.error(f"Error reading {e}")

    if all_sources:
        # Find common available months across all files
        all_month_keys = set()
        for grupos in all_sources.values():
            all_month_keys.update(grupos.keys())
        all_month_keys = sorted(all_month_keys)

        opciones = {f"{MONTHS[m]} {y}": (y, m) for (y, m) in all_month_keys}
        mes_sel  = st.selectbox("Select month", list(opciones.keys()))
        year, month = opciones[mes_sel]

        # Build sources for selected month
        sources = []
        for name, grupos in all_sources.items():
            if (year, month) in grupos:
                sources.append((name, grupos[(year, month)]))

        if not sources:
            st.warning("No transactions found for this month in the uploaded files.")
        else:
            all_rows    = [r for _, rows in sources for r in rows]
            grand_total = sum(float(r["usd"] or 0) for r in all_rows)

            # Metrics
            cols = st.columns(len(sources) + 2)
            cols[0].metric("Grand Total (USD)", f"${grand_total:,.2f}")
            cols[1].metric("Transactions", len(all_rows))
            for i, (name, rows) in enumerate(sources):
                sub = sum(float(r["usd"] or 0) for r in rows)
                cols[i+2].metric(f"{name}", f"${sub:,.2f}")

            st.divider()

            # Preview per person
            for name, rows in sources:
                with st.expander(f"📋 {name} — {len(rows)} transactions", expanded=True):
                    preview = [{
                        "Date":         _fecha(r),
                        "Property":     _prop(r["prop"]),
                        "Category":     r["cat"],
                        "Description":  r["desc"],
                        "Currency":     r["cur"],
                        "Local Amount": fmt_local(r),
                        "USD":          f'${float(r["usd"] or 0):,.2f}',
                    } for r in rows]
                    st.dataframe(preview, use_container_width=True, hide_index=True)

            st.divider()
            st.markdown("**Download files**")

            tz      = pytz.timezone("America/Toronto")
            ts      = datetime.now(tz).strftime("%Y%m%d_%H%M")
            mes_str = MONTHS[month]
            base    = f"SS_CashLedger_Cuba_{mes_str}{year}_{ts}"

            col_pdf, col_xlsx = st.columns(2)

            with col_pdf:
                pdf_buf = generar_pdf(sources, year, month)
                st.download_button(
                    label="📄 Download PDF",
                    data=pdf_buf,
                    file_name=f"{base}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

            with col_xlsx:
                xl_buf = generar_excel(sources, year, month)
                st.download_button(
                    label="📊 Download Excel",
                    data=xl_buf,
                    file_name=f"{base}_Review.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

else:
    st.info("Upload one or more Excel files to get started.")
