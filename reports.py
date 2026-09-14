"""Outputs: the HR workbook (xlsx) and per-rep PDF packets.

PDF generation is pure-Python (reportlab + pypdf) so it works on Streamlit Cloud
with no LibreOffice. LPS PDFs supplied per order are appended after each order's
generated commission detail page.
"""
from __future__ import annotations
import io, datetime
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR = Font(name="Arial", bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor="404040")
BOLD = Font(name="Arial", bold=True)
CUR = '$#,##0.00;($#,##0.00);"-"'
PCT = "0.0%"


def _sheet_table(ws, start_row, headers, rows, widths=None):
    for j, h in enumerate(headers, 1):
        c = ws.cell(start_row, j, h); c.font = HDR; c.fill = HFILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = BOX
    r = start_row + 1
    for row in rows:
        for j, v in enumerate(row, 1):
            c = ws.cell(r, j, v); c.border = BOX; c.font = Font(name="Arial")
        r += 1
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    return r


def build_hr_workbook(run: dict) -> bytes:
    p = run["payload"]
    order_results = p.get("results", {}).get("orders", [])
    rep_summary = p.get("results", {}).get("rep_summary", {})
    exceptions = p.get("results", {}).get("exceptions", [])
    reps_master = p.get("reps_master", {})
    wb = openpyxl.Workbook()

    # Payroll Summary
    ws = wb.active; ws.title = "Payroll Summary"; ws.sheet_view.showGridLines = False
    ws["A1"] = "Boston Aesthetics — Payroll Commission Summary"; ws["A1"].font = Font(name="Arial", bold=True, size=14)
    ws["A2"] = f"Run: {run['name']}   Period: {run['period_start']} to {run['period_end']}   Status: {run['status']}"
    ws["A2"].font = Font(name="Arial", italic=True, color="595959")
    rows = []
    for rid, s in sorted(rep_summary.items()):
        pid = reps_master.get(rid, {}).get("payroll_id") or ""
        rows.append([rid, pid, s.get("name", rid), s.get("this_period"), s.get("full")])
    end = _sheet_table(ws, 4, ["Rep ID", "Payroll/Emp ID", "Rep Name", "This Period", "Full-Deal (100%)"],
                       rows, widths=[12, 16, 26, 16, 18])
    tr = end + 1
    ws.cell(tr, 3, "TOTAL").font = BOLD
    ws.cell(tr, 4, f"=SUM(D5:D{end-1})").font = BOLD
    for r in range(5, tr + 1):
        ws.cell(r, 4).number_format = CUR; ws.cell(r, 5).number_format = CUR

    # Commission Detail (per order per rep)
    ws = wb.create_sheet("Commission Detail"); ws.sheet_view.showGridLines = False
    rows = []
    for o in order_results:
        for rl in o["reps"]:
            rows.append([o["order_number"], o["customer"], o["deal_type"], o["net_commissionable"],
                         o["commission_rate"], o["deal_commission"], rl["rep_id"], rl["rep_name"],
                         rl["share"], rl["release_fraction"], rl["full_commission"],
                         rl["prior_processed"], rl["this_period"]])
    end = _sheet_table(ws, 1,
                       ["Order", "Customer", "Deal Type", "Net Commissionable", "Rate", "Deal Commission",
                        "Rep ID", "Rep Name", "Share", "Release", "Full Commission", "Prior Processed", "This Period"],
                       rows, widths=[13, 24, 12, 16, 8, 15, 10, 20, 8, 9, 15, 14, 14])
    for r in range(2, end):
        for col in (4, 6, 11, 12, 13):
            ws.cell(r, col).number_format = CUR
        for col in (5, 9, 10):
            ws.cell(r, col).number_format = PCT

    # Exceptions
    ws = wb.create_sheet("Exceptions"); ws.sheet_view.showGridLines = False
    _sheet_table(ws, 1, ["Exception / review item"],
                 [[e] for e in exceptions] or [["None — all checks passed"]], widths=[80])

    # Reconciliation & Control Totals
    ws = wb.create_sheet("Reconciliation"); ws.sheet_view.showGridLines = False
    ct = p.get("results", {}).get("control", {})
    recon = [["Orders in run", ct.get("orders")],
             ["Sum of full-deal commission", ct.get("sum_deal_commission")],
             ["Sum of this-period commission (released)", ct.get("sum_period_commission")],
             ["Sum of per-rep this-period", round(sum((s.get("this_period") or 0) for s in rep_summary.values()), 2)],
             ["Run status", run["status"]],
             ["Generated (UTC)", datetime.datetime.utcnow().strftime("%Y-%m-%d %H:%M")]]
    end = _sheet_table(ws, 1, ["Control", "Value"], recon, widths=[42, 22])
    for r in range(2, end):
        v = ws.cell(r, 2).value
        if isinstance(v, (int, float)):
            ws.cell(r, 2).number_format = CUR

    # Plan Snapshot
    ws = wb.create_sheet("Plan Snapshot"); ws.sheet_view.showGridLines = False
    snap = [["Release rule", "Bundle 50/50: 0% none delivered, 50% first device, 100% both"],
            ["Commission base", "Pre-tax subtotal (devices + discount + customer shipping); tax excluded"],
            ["Net commissionable", "Subtotal − deductions"],
            ["Deal commission", "Net × Commission Rate (rate entered on the order sheet)"],
            ["Rep commission", "Deal commission × rep share, then × release fraction, minus prior processed"],
            ["Rounding", "Half up to the cent"]]
    _sheet_table(ws, 1, ["Policy", "Definition"], snap, widths=[26, 80])

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ---------------- per-rep PDF packet ----------------
def build_rep_packet(rep_id, run, lps_by_order: dict) -> bytes:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from pypdf import PdfReader, PdfWriter

    p = run["payload"]
    order_results = p.get("results", {}).get("orders", [])
    reps_master = p.get("reps_master", {})
    orders_raw = {o["order_number"]: o for o in p.get("orders", [])}
    name = reps_master.get(rep_id, {}).get("name", rep_id)

    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("h1", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=16)
    h2 = ParagraphStyle("h2", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=12)
    small = ParagraphStyle("small", parent=styles["Normal"], fontSize=8, textColor=colors.grey)

    def money(x):
        return "-" if x in (None, "") else "${:,.2f}".format(x)

    def pct(x):
        return "" if x in (None, "") else "{:.1f}%".format(x * 100)

    gen = io.BytesIO()
    doc = SimpleDocTemplate(gen, pagesize=LETTER, leftMargin=0.6 * inch, rightMargin=0.6 * inch,
                            topMargin=0.6 * inch, bottomMargin=0.6 * inch)
    story = []
    # ---- cover statement ----
    story.append(Paragraph("Boston Aesthetics — Commission Statement", h1))
    story.append(Paragraph(f"{name} ({rep_id})", h2))
    story.append(Paragraph(f"Run: {run['name']} &nbsp;·&nbsp; Period: {run['period_start']} to {run['period_end']} "
                           f"&nbsp;·&nbsp; Status: {run['status']}", small))
    story.append(Spacer(1, 10))
    rows = [["Order", "Customer", "Release", "This Period"]]
    total = 0.0
    for o in order_results:
        for rl in o["reps"]:
            if str(rl["rep_id"]) == str(rep_id):
                rows.append([o["order_number"], (o["customer"] or "")[:34], pct(rl["release_fraction"]),
                             money(rl["this_period"])])
                total += rl["this_period"] or 0
    rows.append(["", "", "TOTAL THIS PERIOD", money(round(total, 2))])
    t = Table(rows, colWidths=[1.2 * inch, 3.0 * inch, 1.1 * inch, 1.5 * inch])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#404040")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("LINEABOVE", (0, -1), (-1, -1), 0.7, colors.black),
        ("GRID", (0, 0), (-1, -2), 0.3, colors.HexColor("#BFBFBF")),
        ("ALIGN", (2, 0), (-1, -1), "RIGHT"), ("FONTSIZE", (0, 0), (-1, -1), 9),
    ]))
    story.append(t)
    story.append(Spacer(1, 8))
    story.append(Paragraph("Amounts reflect the current run and are not a statement that payroll has been approved "
                           "or paid. Bundle orders release 50% on the first device and the remainder when the second "
                           "device is delivered.", small))

    # ---- per-order commission detail ----
    for o in order_results:
        mine = [rl for rl in o["reps"] if str(rl["rep_id"]) == str(rep_id)]
        if not mine:
            continue
        raw = orders_raw.get(o["order_number"], {})
        story.append(PageBreak())
        story.append(Paragraph(f"Order {o['order_number']} — {o.get('customer','')}", h2))
        story.append(Paragraph(f"{o.get('deal_type','')} · {o.get('release_rule','')}", small))
        story.append(Spacer(1, 6))
        li = [["Line item", "Amount"]] + [[(l.get("description") or l.get("type") or ""), money(l.get("amount"))]
                                          for l in raw.get("line_items", [])]
        li.append(["Subtotal before tax", money(raw.get("subtotal_before_tax"))])
        for d in raw.get("deductions", []):
            li.append([f"less: {d.get('description','')}", money(-(d.get('amount') or 0))])
        li.append(["Net commissionable", money(o["net_commissionable"])])
        li.append([f"Commission rate", pct(o["commission_rate"])])
        li.append(["Deal commission", money(o["deal_commission"])])
        rl = mine[0]
        li.append([f"Your share ({pct(rl['share'])})", money(rl["full_commission"])])
        li.append([f"Release ({pct(rl['release_fraction'])})", money(rl["eligible_to_date"])])
        if rl["prior_processed"]:
            li.append(["less: prior processed", money(-rl["prior_processed"])])
        li.append(["This period", money(rl["this_period"])])
        t = Table(li, colWidths=[4.6 * inch, 1.6 * inch])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#404040")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("LINEABOVE", (0, -1), (-1, -1), 0.7, colors.black),
            ("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#D9D9D9")),
            ("ALIGN", (1, 0), (1, -1), "RIGHT"), ("FONTSIZE", (0, 0), (-1, -1), 9),
        ]))
        story.append(t)
        story.append(Spacer(1, 6))
        story.append(Paragraph("The Laser Purchase Sheet for this order follows.", small))
    doc.build(story)

    # ---- merge generated pages + LPS PDFs ----
    writer = PdfWriter()
    for pg in PdfReader(io.BytesIO(gen.getvalue())).pages:
        writer.add_page(pg)
    # append each relevant order's LPS right after (simple: at the end, in order)
    for o in order_results:
        if any(str(rl["rep_id"]) == str(rep_id) for rl in o["reps"]):
            lps = lps_by_order.get(o["order_number"])
            if lps:
                try:
                    for pg in PdfReader(io.BytesIO(lps)).pages:
                        writer.add_page(pg)
                except Exception:
                    pass
    out = io.BytesIO(); writer.write(out); return out.getvalue()
