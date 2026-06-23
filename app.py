#!/usr/bin/env python3
"""
Cash Ledger PDF Generator — LCC Cuba
====================================
Genera PDFs del Cash Ledger por mes, con página de aprobación para Rolando.

Uso:
    python cash_ledger_pdf.py                              # modo interactivo
    python cash_ledger_pdf.py archivo.xlsx                 # muestra meses disponibles
    python cash_ledger_pdf.py archivo.xlsx --mes 4         # genera PDF de abril
    python cash_ledger_pdf.py archivo.xlsx --mes 4 5       # genera PDFs de abril y mayo
    python cash_ledger_pdf.py archivo.xlsx --todos         # genera todos los meses

Opciones:
    --salida CARPETA   Carpeta donde guardar los PDFs (por defecto: misma carpeta que el Excel)

Requiere:
    pip install openpyxl reportlab
"""

import sys, os, argparse
from datetime import datetime
import openpyxl
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

def leer_excel(path):
    wb   = openpyxl.load_workbook(path, data_only=True)
    ws   = next((wb[s] for s in wb.sheetnames if "cash" in s.lower()), wb.active)
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError("El archivo parece estar vacío.")

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
            "year":  anio,
            "month": mes,
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


# ─── PDF ────────────────────────────────────

def _fmt(v):
    try:    return f"${float(v):,.2f}"
    except: return ""

def _fecha(r):
    try:    return f"{int(r['month']):02d}/{int(r['day']):02d}/{r['year']}"
    except: return ""

def _prop(p):
    return p.replace("Property 1: ","").replace("Property 2: ","")


def generar_pdf(rows, year, month, output_path, fuente=""):
    mes_str = MESES.get(month, str(month))
    periodo = f"{mes_str} {year}"
    total   = sum(float(r["usd"] or 0) for r in rows)
    bal_fin = float(rows[-1]["bal"] or 0) if rows else 0.0

    styles = getSampleStyleSheet()
    def S(name, parent="Normal", **kw):
        return ParagraphStyle(name, parent=styles[parent], **kw)

    t_s  = S("t",  fontSize=15, fontName="Helvetica-Bold", textColor=colors.white, spaceAfter=0)
    sb_s = S("sb", fontSize=9,  textColor=colors.HexColor("#c0c0cc"), spaceAfter=0)
    h_s  = S("h",  fontSize=8.5,fontName="Helvetica-Bold", textColor=MID, spaceBefore=12, spaceAfter=5)
    b_s  = S("b",  fontSize=10, textColor=colors.HexColor("#222222"), leading=16)
    sm_s = S("sm", fontSize=7.5,textColor=MID)

    doc = SimpleDocTemplate(
        output_path, pagesize=letter,
        leftMargin=0.65*inch, rightMargin=0.65*inch,
        topMargin=0.6*inch,   bottomMargin=0.6*inch,
    )
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

    # ── Página 1: Ledger ──────────────────────
    story.append(banner("LCC — Cash Ledger (Cuba)", f"Período: {periodo}"))
    story.append(Spacer(1, 8))

    # Strip resumen
    strip = Table([
        [Paragraph("Total gastos USD", sm_s),
         Paragraph("Running balance USD", sm_s),
         Paragraph("Transacciones", sm_s)],
        [Paragraph(f'<font color="#a32d2d"><b>{_fmt(total)}</b></font>', styles["Normal"]),
         Paragraph(f'<font color="#a32d2d"><b>{_fmt(bal_fin)}</b></font>', styles["Normal"]),
         Paragraph(f"<b>{len(rows)}</b>", styles["Normal"])],
    ], colWidths=["33%","33%","34%"])
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
    story.append(Paragraph("TRANSACCIONES", h_s))

    def fmt_local(r):
        try:
            v   = float(r["local"] or 0)
            cur = r["cur"].strip()
            return f"${v:,.2f}" if cur == "USD" else f"{v:,.0f} {cur}"
        except:
            return ""

    heads = ["Fecha", "Propiedad", "Categoría", "Descripción", "Mon.", "Monto Local", "USD"]
    col_w = [0.7*inch, 0.85*inch, 1.05*inch, 2.1*inch, 0.38*inch, 0.9*inch, 0.72*inch]
    tx    = [heads] + [[
        _fecha(r),
        _prop(r["prop"])[:12],
        r["cat"][:20],
        r["desc"][:35],
        r["cur"],
        fmt_local(r),
        f'${float(r["usd"] or 0):,.2f}',
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
        # Monto Local (5): gris discreto, alineado derecha
        ("TEXTCOLOR",     (5,1),(5,-1),  MID),
        ("ALIGN",         (5,1),(5,-1),  "RIGHT"),
        # USD (6): rojo negrita, destacado
        ("TEXTCOLOR",     (6,1),(6,-1),  RED),
        ("FONTNAME",      (6,1),(6,-1),  "Helvetica-Bold"),
        ("ALIGN",         (6,1),(6,-1),  "RIGHT"),
        ("BACKGROUND",    (6,1),(6,-1),  colors.HexColor("#fff5f5")),
        # Separador visual antes de la columna USD
        ("LINEAFTER",     (5,0),(5,-1),  0.8, colors.HexColor("#aaaaaa")),
        ("ALIGN",         (4,1),(4,-1),  "CENTER"),
        ("GRID",          (0,0),(-1,-1), 0.3, GBRD),
        ("TOPPADDING",    (0,0),(-1,-1), 3),
        ("BOTTOMPADDING", (0,0),(-1,-1), 3),
        ("LEFTPADDING",   (0,0),(-1,-1), 4),
        ("RIGHTPADDING",  (0,0),(-1,-1), 4),
    ]))
    story.append(t)

    # ── Página 2: Aprobación ──────────────────
    story.append(PageBreak())
    story.append(banner("LCC — Aprobación de Cash Ledger", f"Período: {periodo}", "65%", "35%"))
    story.append(Spacer(1, 20))

    story.append(Paragraph("RESUMEN", h_s))
    recap = Table([
        ["Período",               periodo],
        ["Total gastos USD",      _fmt(total)],
        ["Running balance USD",   _fmt(bal_fin)],
        ["N° de transacciones",   str(len(rows))],
    ], colWidths=[3*inch, 2*inch])
    recap.setStyle(TableStyle([
        ("FONTNAME",      (0,0),(-1,-1), "Helvetica"),
        ("FONTNAME",      (0,0),(0,-1),  "Helvetica-Bold"),
        ("FONTSIZE",      (0,0),(-1,-1), 10),
        ("TEXTCOLOR",     (0,0),(-1,-1), colors.HexColor("#333333")),
        ("ALIGN",         (1,0),(1,-1),  "RIGHT"),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LINEBELOW",     (0,-1),(-1,-1), 1, DARK),
    ]))
    story.append(recap)
    story.append(Spacer(1, 28))

    story.append(Paragraph("APROBACIÓN", h_s))
    story.append(Paragraph(
        f"He revisado el Cash Ledger del período <b>{periodo}</b> y confirmo que "
        "las transacciones registradas son completas y han sido correctamente documentadas.",
        b_s
    ))
    story.append(Spacer(1, 36))

    sig = Table([
        [Paragraph("<b>Aprobado por:</b>", b_s), ""],
        [Spacer(1,6),  ""],
        [Paragraph("Nombre:&nbsp;&nbsp;&nbsp;<b>Rolando</b>", b_s), ""],
        [Spacer(1,14), ""],
        [Paragraph("Firma:", b_s), ""],
        [Spacer(1,22), ""],
        [Paragraph("Fecha:", b_s), ""],
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
        f"Generado: {datetime.now().strftime('%B %d, %Y')}  |  {fuente}",
        sm_s
    ))

    doc.build(story)
    print(f"  ✅  {os.path.basename(output_path)}")


# ─── CLI ────────────────────────────────────

def mostrar_meses(grupos):
    print("\n  Meses disponibles:")
    for (y, m), rows in grupos.items():
        print(f"    {m:2d}  {MESES[m]:<12} {y}   ({len(rows)} transacciones)")
    print()

def main():
    parser = argparse.ArgumentParser(
        description="Genera PDFs de aprobación del Cash Ledger de Cuba por mes.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("excel",    nargs="?",    help="Archivo Excel (.xlsx)")
    parser.add_argument("--mes",    nargs="+", type=int, metavar="N",
                        help="Número(s) de mes (ej: --mes 4 5)")
    parser.add_argument("--todos",  action="store_true",
                        help="Genera PDFs de todos los meses disponibles")
    parser.add_argument("--salida", metavar="CARPETA",
                        help="Carpeta donde guardar los PDFs")
    args = parser.parse_args()

    # Modo interactivo
    if not args.excel:
        print("\n  Cash Ledger PDF Generator — LCC Cuba")
        print("  ─────────────────────────────────────")
        args.excel = input("  Ruta del archivo Excel: ").strip().strip('"')

    if not os.path.exists(args.excel):
        print(f"\n  ❌  Archivo no encontrado: {args.excel}")
        sys.exit(1)

    print(f"\n  Leyendo {os.path.basename(args.excel)}...")
    try:
        data = leer_excel(args.excel)
    except Exception as e:
        print(f"  ❌  Error: {e}")
        sys.exit(1)

    grupos = agrupar(data)
    if not grupos:
        print("  ❌  No se encontraron transacciones.")
        sys.exit(1)

    mostrar_meses(grupos)

    # Qué meses generar
    if args.todos:
        seleccion = list(grupos.keys())
    elif args.mes:
        seleccion = []
        for m in args.mes:
            found = [k for k in grupos if k[1] == m]
            if found:
                seleccion.extend(found)
            else:
                print(f"  ⚠️   Mes {m} sin datos, se omite.")
    else:
        disponibles = sorted(set(m for (_, m) in grupos))
        print(f"  Meses con datos: {', '.join(str(m) for m in disponibles)}")
        entrada = input("  ¿Qué mes(es) generar? (ej: 4  o  4 5  o  todos): ").strip().lower()
        if entrada == "todos":
            seleccion = list(grupos.keys())
        else:
            nums = [int(x) for x in entrada.split() if x.isdigit()]
            seleccion = [k for k in grupos if k[1] in nums]

    if not seleccion:
        print("  ❌  Ningún mes seleccionado.")
        sys.exit(1)

    # Carpeta de salida
    if args.salida:
        salida_dir = args.salida
        os.makedirs(salida_dir, exist_ok=True)
    else:
        excel_dir = os.path.dirname(os.path.abspath(args.excel))
        salida_dir = excel_dir if os.access(excel_dir, os.W_OK) else os.getcwd()

    print(f"\n  Generando {len(seleccion)} PDF(s) en: {salida_dir}\n")
    fuente = os.path.basename(args.excel)
    for (year, month) in sorted(seleccion):
        out = os.path.join(salida_dir, f"CashLedger_{MESES[month]}{year}_Cuba.pdf")
        generar_pdf(grupos[(year, month)], year, month, out, fuente)

    print(f"\n  Listo.\n")




# ─────────────────────────────────────────────────────────────
# GENERADOR EXCEL
# ─────────────────────────────────────────────────────────────

def generar_excel(rows, year, month, output_path):
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, numbers
    )
    from openpyxl.utils import get_column_letter

    mes_str = MESES.get(month, str(month))
    periodo = f"{mes_str} {year}"

    wb = Workbook()

    # ── Hoja 1: Transacciones ──────────────────
    ws = wb.active
    ws.title = f"Gastos {mes_str}"

    # Colores
    dark_fill  = PatternFill("solid", fgColor="1A1A2E")
    gray_fill  = PatternFill("solid", fgColor="F5F5F5")
    gray2_fill = PatternFill("solid", fgColor="EBEBEB")
    usd_fill   = PatternFill("solid", fgColor="FFF0F0")
    total_fill = PatternFill("solid", fgColor="FFE8E8")
    white_fill = PatternFill("solid", fgColor="FFFFFF")

    white_bold = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    dark_hdr   = Font(name="Arial", bold=True, color="1A1A2E", size=10)
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
    right  = Alignment(horizontal="right",  vertical="center", wrap_text=False)

    # ── Título ──
    ws.merge_cells("A1:G1")
    ws["A1"] = f"LCC — Cash Ledger (Cuba)   |   Período: {periodo}"
    ws["A1"].font = white_bold
    ws["A1"].fill = dark_fill
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 22

    ws.merge_cells("A2:G2")
    ws["A2"] = f"Verificación de gastos — para revisión de Rolando"
    ws["A2"].font = Font(name="Arial", size=9, color="888888", italic=True)
    ws["A2"].alignment = center
    ws.row_dimensions[2].height = 16

    # ── Encabezados ──
    heads = ["Fecha", "Propiedad", "Categoría", "Descripción", "Moneda", "Monto Local", "USD"]
    col_widths = [12, 14, 20, 38, 8, 15, 13]

    for i, h in enumerate(heads, 1):
        c = ws.cell(row=3, column=i, value=h)
        c.font   = white_bold if i != 7 else Font(name="Arial", bold=True, color="FFFFFF", size=11)
        c.fill   = dark_fill
        c.alignment = center
        c.border = bdr
        ws.column_dimensions[get_column_letter(i)].width = col_widths[i-1]
    ws.row_dimensions[3].height = 18

    # ── Filas de datos ──
    num_local = "###,##0.00"
    num_usd   = "$#,##0.00"

    for idx, r in enumerate(rows):
        row_num = idx + 4
        is_alt  = idx % 2 == 1

        def cell(col, val, font=normal, fill=None, align=left, fmt=None, border=bdr):
            c = ws.cell(row=row_num, column=col, value=val)
            c.font      = font
            c.fill      = fill or (gray_fill if is_alt else white_fill)
            c.alignment = align
            c.border    = border
            if fmt: c.number_format = fmt
            return c

        try:
            day = int(r["day"])
            cell(1, f"{int(r['month']):02d}/{day:02d}/{r['year']}", align=center)
        except:
            cell(1, "")

        prop = _prop(r["prop"])
        cell(2, prop)
        cell(3, r["cat"])
        cell(4, r["desc"])
        cell(5, r["cur"], align=center)

        try:
            local_v = float(r["local"] or 0)
            cell(6, local_v, font=mid_font, align=right, fmt=num_local)
        except:
            cell(6, "")

        try:
            usd_v = float(r["usd"] or 0)
            c = ws.cell(row=row_num, column=7, value=usd_v)
            c.font           = usd_font
            c.fill           = usd_fill
            c.alignment      = right
            c.number_format  = num_usd
            c.border         = usd_bdr
        except:
            ws.cell(row=row_num, column=7, value="")

        ws.row_dimensions[row_num].height = 15

    # ── Fila de total ──
    total_row = len(rows) + 4
    ws.row_dimensions[total_row].height = 18

    for col in range(1, 7):
        c = ws.cell(row=total_row, column=col, value="")
        c.fill   = total_fill
        c.border = bdr

    ws.cell(row=total_row, column=1, value="TOTAL").font = total_font
    ws.cell(row=total_row, column=1).fill = total_fill
    ws.cell(row=total_row, column=1).alignment = center

    last_data = total_row - 1
    formula_usd = f"=SUM(G4:G{last_data})"
    c = ws.cell(row=total_row, column=7, value=formula_usd)
    c.font          = total_font
    c.fill          = total_fill
    c.alignment     = right
    c.number_format = num_usd
    c.border        = Border(left=med, right=thin, top=Side(style="medium", color="A32D2D"),
                             bottom=Side(style="medium", color="A32D2D"))

    # Freeze panes
    ws.freeze_panes = "A4"

    # ── Hoja 2: Aprobación ────────────────────
    wa = wb.create_sheet("Aprobación")

    total_usd = sum(float(r["usd"] or 0) for r in rows)

    wa.column_dimensions["A"].width = 28
    wa.column_dimensions["B"].width = 22

    def wa_cell(row, col, val, font=None, fill=None, align=None, fmt=None, merge_to=None):
        c = wa.cell(row=row, column=col, value=val)
        if font:  c.font = font
        if fill:  c.fill = fill
        if align: c.alignment = align
        if fmt:   c.number_format = fmt
        if merge_to:
            wa.merge_cells(start_row=row, start_column=col,
                           end_row=row,   end_column=merge_to)
        return c

    wa.row_dimensions[1].height = 28
    wa_cell(1, 1, f"LCC — Aprobación de Cash Ledger   |   {periodo}",
            font=Font(name="Arial", bold=True, color="FFFFFF", size=13),
            fill=dark_fill, align=center, merge_to=2)

    wa.row_dimensions[2].height = 10

    # Resumen
    for r, label, val, is_usd in [
        (3, "Período",               periodo,     False),
        (4, "Total gastos USD",       total_usd,   True),
        (5, "N° de transacciones",    len(rows),   False),
    ]:
        wa.row_dimensions[r].height = 18
        wa_cell(r, 1, label,
                font=Font(name="Arial", bold=True, size=10, color="555555"),
                fill=gray_fill,
                align=left)

        c = wa_cell(r, 2, val,
                font=Font(name="Arial", bold=True, size=10,
                          color="A32D2D" if is_usd else "1A1A2E"),
                fill=gray_fill, align=right)
        if is_usd: c.number_format = num_usd

    wa.row_dimensions[6].height = 20

    # Texto de aprobación
    wa.merge_cells("A7:B7")
    wa_cell(7, 1,
            f"He revisado el Cash Ledger del período {periodo} y confirmo que los gastos son correctos.",
            font=Font(name="Arial", size=10, italic=True, color="333333"),
            align=Alignment(wrap_text=True, vertical="center"))
    wa.row_dimensions[7].height = 32

    wa.row_dimensions[8].height = 14

    # Bloques de firma
    for r, label, val in [
        (9,  "Aprobado por:",  ""),
        (10, "Nombre:",        "Rolando"),
        (11, "Firma:",         ""),
        (12, "Fecha:",         ""),
    ]:
        wa.row_dimensions[r].height = 22
        wa_cell(r, 1, label,
                font=Font(name="Arial", bold=True, size=10, color="1A1A2E"),
                fill=gray2_fill, align=left)
        c = wa_cell(r, 2, val,
                font=Font(name="Arial", size=10, color="1A1A2E"),
                fill=gray2_fill, align=left)
        # Línea baja en firma y fecha
        if r in (11, 12):
            c.border = Border(bottom=Side(style="medium", color="1A1A2E"))

    wb.save(output_path)
    print(f"  ✅  {os.path.basename(output_path)}")



def main():
    parser = argparse.ArgumentParser(
        description="Genera PDFs y/o Excel del Cash Ledger de Cuba por mes.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("excel",     nargs="?",        help="Archivo Excel (.xlsx)")
    parser.add_argument("--mes",     nargs="+", type=int, metavar="N")
    parser.add_argument("--todos",   action="store_true")
    parser.add_argument("--salida",  metavar="CARPETA")
    parser.add_argument("--formato", choices=["pdf","xlsx","ambos"], default="ambos",
                        help="Qué generar: pdf, xlsx, o ambos (default: ambos)")
    args = parser.parse_args()

    if not args.excel:
        print("\n  Cash Ledger Generator — LCC Cuba")
        print("  ──────────────────────────────────")
        args.excel = input("  Ruta del archivo Excel: ").strip().strip('"')

    if not os.path.exists(args.excel):
        print(f"\n  ❌  Archivo no encontrado: {args.excel}")
        sys.exit(1)

    print(f"\n  Leyendo {os.path.basename(args.excel)}...")
    try:
        data = leer_excel(args.excel)
    except Exception as e:
        print(f"  ❌  Error: {e}"); sys.exit(1)

    grupos = agrupar(data)
    if not grupos:
        print("  ❌  No se encontraron transacciones."); sys.exit(1)

    mostrar_meses(grupos)

    if args.todos:
        seleccion = list(grupos.keys())
    elif args.mes:
        seleccion = []
        for m in args.mes:
            found = [k for k in grupos if k[1] == m]
            if found: seleccion.extend(found)
            else: print(f"  ⚠️   Mes {m} sin datos.")
    else:
        disponibles = sorted(set(m for (_, m) in grupos))
        print(f"  Meses con datos: {', '.join(str(m) for m in disponibles)}")
        entrada = input("  ¿Qué mes(es)? (ej: 4  o  4 5  o  todos): ").strip().lower()
        if entrada == "todos":
            seleccion = list(grupos.keys())
        else:
            nums = [int(x) for x in entrada.split() if x.isdigit()]
            seleccion = [k for k in grupos if k[1] in nums]

    if not seleccion:
        print("  ❌  Ningún mes seleccionado."); sys.exit(1)

    excel_dir = os.path.dirname(os.path.abspath(args.excel))
    salida_dir = args.salida or (excel_dir if os.access(excel_dir, os.W_OK) else os.getcwd())
    os.makedirs(salida_dir, exist_ok=True)

    fmt = args.formato
    fuente = os.path.basename(args.excel)
    total = len(seleccion) * (2 if fmt == "ambos" else 1)
    print(f"\n  Generando {total} archivo(s) en: {salida_dir}\n")

    for (year, month) in sorted(seleccion):
        mes_str = MESES[month]
        filas   = grupos[(year, month)]
        base    = f"CashLedger_{mes_str}{year}_Cuba"

        if fmt in ("pdf", "ambos"):
            generar_pdf(filas, year, month,
                        os.path.join(salida_dir, f"{base}.pdf"), fuente)
        if fmt in ("xlsx", "ambos"):
            generar_excel(filas, year, month,
                          os.path.join(salida_dir, f"{base}_Revision.xlsx"))

    print(f"\n  Listo.\n")


if __name__ == "__main__":
    main()
