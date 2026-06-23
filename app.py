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
    TableStyle, PageBreak, HRFlowable
)

MESES = {
    1:"January", 2:"February", 3:"March", 4:"April",
    5:"May",     6:"June",     7:"July",  8:"August",
    9:"September",10:"October",11:"November",12:"December"
}

DARK  = colors.HexColor("#1a1a2e")
GBGD  = colors.HexColor("#f5f5f5")
GBRD  = colors.HexColor("#e8e8e8")
MID   = colors.HexColor("#666670")
RED   = colors.HexColor("#a32d2d")


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


# ─── GENERADOR PDF ──────────────────────────

def generar_pdf(rows, year, month):
    mes_str = MESES.get(month, str(month))
    periodo = f"{mes_str} {year}"
    total   = sum(float(r["usd"] or 0) for r in rows)

    styles = getSampleStyleSheet()
    def S(name, parent="Normal", **kw):
        return ParagraphStyle(name, parent=styles[parent], **kw)

    t_s  = S("t",  fontSize=15, fontName="Helvetica-Bold", textColor=colors.white)
    sb_s = S("sb", fontSize=9,  textColor=colors.HexColor("#c0c0cc"))
    h_s  = S("h",  fontSize=8.5,fontName="Helvetica-Bold", textColor=MID, spaceBefore=12, spaceAfter=5)
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

    # Página 1: Ledger
    story.append(banner("LCC — Cash Ledger (Cuba)", f"Period: {periodo}"))
    story.append(Spacer(1, 8))

    strip = Table([
        [Paragraph("Total Expenses (USD)", sm_s), Paragraph("Transactions", sm_s)],
        [Paragraph(f'<font color="#a32d2d"><b>{_fmt(total)}</b></font>', styles["Normal"]),
         Paragraph(f"<b>{len(rows)}</b>", styles["Normal"])],
    ], colWidths=["50%","50%"])
    strip.setStyle(TableStyle([
        ("BACKGROUND",   (0,0),(-1,-1), GBGD),
        ("GRID",         (0,0),(-1,-1), 0.3, GBRD),
        ("TOPPADDING",   (0,0),(-1,-1), 6),
        ("BOTTOMPADDING",(0,0),(-1,-1), 6),
        ("LEFTPADDING",  (0,0),(-1,-1), 10),
        ("FONTSIZE",     (0,1),(-1,-1), 11),
    ]))
    story.append(strip)
    story.append(Spacer(1, 8))
    story.append(Paragraph("TRANSACTIONS", h_s))

    heads = ["Date", "Property", "Category", "Description", "Cur.", "Local Amount", "USD"]
    col_w = [0.7*inch, 0.85*inch, 1.0*inch, 2.6*inch, 0.38*inch, 0.8*inch, 0.67*inch]
    tx = [heads] + [[
        _fecha(r), _prop(r["prop"])[:12], r["cat"][:20], r["desc"][:35],
        r["cur"], fmt_local(r), f'${float(r["usd"] or 0):,.2f}',
    ] for r in rows]

    t = Table(tx, colWidths=col_w, repeatRows=1)
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
    story.append(t)

    # Página 2: Aprobación
    story.append(PageBreak())
    story.append(banner("LCC — Cash Ledger Approval", f"Period: {periodo}", "65%", "35%"))
    story.append(Spacer(1, 20))
    story.append(Paragraph("SUMMARY", h_s))

    recap = Table([
        ["Period",              periodo],
        ["Total Expenses (USD)", _fmt(total)],
        ["No. of Transactions",  str(len(rows))],
    ], colWidths=[3*inch, 2*inch])
    recap.setStyle(TableStyle([
        ("FONTNAME",      (0,0),(-1,-1), "Helvetica"),
        ("FONTNAME",      (0,0),(0,-1),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0),(-1,-1), 10),
        ("ALIGN",         (1,0),(1,-1),  "RIGHT"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LINEBELOW",     (0,-1),(-1,-1), 1, DARK),
    ]))
    story.append(recap)
    story.append(Spacer(1, 28))
    story.append(Paragraph("APPROVAL", h_s))
    story.append(Paragraph(
        f"Reviewed and approved. I confirm the transactions recorded in this Cash Ledger for <b>{periodo}</b> are accurate and complete.",
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
    story.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y')}", sm_s))

    doc.build(story)
    buf.seek(0)
    return buf


# ─── GENERADOR EXCEL ────────────────────────

def generar_excel(rows, year, month):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    mes_str = MESES.get(month, str(month))
    periodo = f"{mes_str} {year}"
    total_usd = sum(float(r["usd"] or 0) for r in rows)

    wb = Workbook()
    ws = wb.active
    ws.title = f"Expenses {mes_str}"

    dark_fill  = PatternFill("solid", fgColor="1A1A2E")
    gray_fill  = PatternFill("solid", fgColor="F5F5F5")
    gray2_fill = PatternFill("solid", fgColor="EBEBEB")
    usd_fill   = PatternFill("solid", fgColor="FFF0F0")
    total_fill = PatternFill("solid", fgColor="FFE8E8")
    white_fill = PatternFill("solid", fgColor="FFFFFF")

    white_bold = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    normal     = Font(name="Arial", size=9)
    mid_font   = Font(name="Arial", size=9, color="888888")
    usd_font   = Font(name="Arial", bold=True, size=9, color="A32D2D")
    total_font = Font(name="Arial", bold=True, size=10, color="A32D2D")

    thin = Side(style="thin", color="DDDDDD")
    med  = Side(style="medium", color="AAAAAA")
    bdr  = Border(left=thin, right=thin, top=thin, bottom=thin)
    usd_bdr = Border(left=med, right=thin, top=thin, bottom=thin)

    center = Alignment(horizontal="center", vertical="center")
    left   = Alignment(horizontal="left",   vertical="center")
    right  = Alignment(horizontal="right",  vertical="center")

    ws.merge_cells("A1:G1")
    ws["A1"] = f"LCC — Cash Ledger (Cuba)   |   Period: {periodo}"
    ws["A1"].font = white_bold
    ws["A1"].fill = dark_fill
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 22

    ws.merge_cells("A2:G2")
    ws["A2"] = "Expense review — for Rolando's approval"
    ws["A2"].font = Font(name="Arial", size=9, color="888888", italic=True)
    ws["A2"].alignment = center
    ws.row_dimensions[2].height = 16

    heads = ["Date", "Property", "Category", "Description", "Currency", "Local Amount", "USD"]
    col_widths = [12, 14, 20, 48, 8, 15, 13]
    for i, h in enumerate(heads, 1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = white_bold
        c.fill = dark_fill
        c.alignment = center
        c.border = bdr
        ws.column_dimensions[get_column_letter(i)].width = col_widths[i-1]
    ws.row_dimensions[3].height = 18

    num_usd = "$#,##0.00"
    num_local = "###,##0.00"

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

    total_row = len(rows) + 4
    ws.row_dimensions[total_row].height = 18
    for col in range(1, 7):
        c = ws.cell(row=total_row, column=col, value="")
        c.fill = total_fill; c.border = bdr
    ws.cell(row=total_row, column=1, value="TOTAL").font = total_font
    ws.cell(row=total_row, column=1).fill = total_fill
    ws.cell(row=total_row, column=1).alignment = center
    c = ws.cell(row=total_row, column=7, value=f"=SUM(G4:G{total_row-1})")
    c.font = total_font; c.fill = total_fill
    c.alignment = right; c.number_format = num_usd
    c.border = Border(left=med, right=thin,
                      top=Side(style="medium", color="A32D2D"),
                      bottom=Side(style="medium", color="A32D2D"))
    ws.freeze_panes = "A4"

    # Hoja aprobación
    wa = wb.create_sheet("Approval")
    wa.column_dimensions["A"].width = 28
    wa.column_dimensions["B"].width = 22

    def wa_cell(row, col, val, font=None, fill=None, align=None, fmt=None, merge_to=None):
        c = wa.cell(row=row, column=col, value=val)
        if font:  c.font = font
        if fill:  c.fill = fill
        if align: c.alignment = align
        if fmt:   c.number_format = fmt
        if merge_to:
            wa.merge_cells(start_row=row, start_column=col, end_row=row, end_column=merge_to)
        return c

    wa.row_dimensions[1].height = 28
    wa_cell(1, 1, f"LCC — Cash Ledger Approval   |   {periodo}",
            font=Font(name="Arial", bold=True, color="FFFFFF", size=13),
            fill=dark_fill, align=center, merge_to=2)
    wa.row_dimensions[2].height = 10

    for row, label, val, is_usd in [
        (3, "Period",               periodo,    False),
        (4, "Total Expenses (USD)", total_usd,  True),
        (5, "No. of Transactions",  len(rows),  False),
    ]:
        wa.row_dimensions[row].height = 18
        wa_cell(row, 1, label,
                font=Font(name="Arial", bold=True, size=10, color="555555"),
                fill=gray_fill, align=left)
        c = wa_cell(row, 2, val,
                font=Font(name="Arial", bold=True, size=10,
                          color="A32D2D" if is_usd else "1A1A2E"),
                fill=gray_fill, align=right)
        if is_usd: c.number_format = num_usd

    wa.row_dimensions[6].height = 20
    wa.merge_cells("A7:B7")
    wa_cell(7, 1,
            f"Reviewed and approved. I confirm the transactions recorded in this Cash Ledger for {periodo} are accurate and complete.",
            font=Font(name="Arial", size=10, italic=True, color="333333"),
            align=Alignment(wrap_text=True, vertical="center"))
    wa.row_dimensions[7].height = 32
    wa.row_dimensions[8].height = 14

    for row, label, val in [
        (9,  "Approved by:", ""),
        (10, "Name:",          "Rolando"),
        (11, "Signature:",    ""),
        (12, "Date:",         ""),
    ]:
        wa.row_dimensions[row].height = 22
        wa_cell(row, 1, label,
                font=Font(name="Arial", bold=True, size=10, color="1A1A2E"),
                fill=gray2_fill, align=left)
        c = wa_cell(row, 2, val,
                font=Font(name="Arial", size=10, color="1A1A2E"),
                fill=gray2_fill, align=left)
        if row in (11, 12):
            c.border = Border(bottom=Side(style="medium", color="1A1A2E"))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── STREAMLIT UI ───────────────────────────

st.set_page_config(page_title="Cash Ledger LCC", page_icon="📊", layout="centered")

st.markdown("""
    <style>
    .block-container { max-width: 760px; }
    </style>
""", unsafe_allow_html=True)

st.markdown("## 📊 Cash Ledger — LCC Cuba")
st.markdown("Upload the Cash Ledger, select a month, and download the PDF or Excel for review.")
st.divider()

uploaded = st.file_uploader("Upload Cash Ledger (.xlsx)", type=["xlsx"])

if uploaded:
    try:
        data   = leer_excel(uploaded)
        grupos = agrupar(data)

        if not grupos:
            st.error("No transactions found in the file.")
        else:
            meses_disponibles = sorted(grupos.keys())
            opciones = {f"{MESES[m]} {y}": (y, m) for (y, m) in meses_disponibles}

            mes_sel = st.selectbox("Select month", list(opciones.keys()))
            year, month = opciones[mes_sel]
            rows = grupos[(year, month)]
            total = sum(float(r["usd"] or 0) for r in rows)

            col1, col2, col3 = st.columns(3)
            col1.metric("Transactions", len(rows))
            col2.metric("Total USD", f"${total:,.2f}")
            col3.metric("Month", MESES[month])

            st.divider()

            # Preview tabla
            with st.expander("View transactions", expanded=True):
                preview = []
                for r in rows:
                    preview.append({
                        "Date":         _fecha(r),
                        "Property":     _prop(r["prop"]),
                        "Category":     r["cat"],
                        "Description":  r["desc"],
                        "Currency":     r["cur"],
                        "Local Amount": fmt_local(r),
                        "USD":         f'${float(r["usd"] or 0):,.2f}',
                    })
                st.dataframe(preview, use_container_width=True, hide_index=True)

            st.divider()
            st.markdown("**Download files**")

            col_pdf, col_xlsx = st.columns(2)
            mes_str = MESES[month]

            with col_pdf:
                pdf_buf = generar_pdf(rows, year, month)
                st.download_button(
                    label="📄 Download PDF",
                    data=pdf_buf,
                    file_name=f"SS_CashLedger_Cuba_{mes_str}{year}_{datetime.now(pytz.timezone('America/Toronto')).strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf",
                    use_container_width=True,
                )

            with col_xlsx:
                xl_buf = generar_excel(rows, year, month)
                st.download_button(
                    label="📊 Download Excel",
                    data=xl_buf,
                    file_name=f"SS_CashLedger_Cuba_{mes_str}{year}_{datetime.now(pytz.timezone('America/Toronto')).strftime('%Y%m%d_%H%M')}_Review.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

    except Exception as e:
        st.error(f"Error reading the file: {e}")

else:
    st.info("Upload the Excel file to get started.")
