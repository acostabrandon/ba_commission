"""Outputs: the HR workbook (xlsx) and per-rep PDF packets.

PDF generation is pure-Python (reportlab + pypdf) so it works on Streamlit Cloud
with no LibreOffice. LPS PDFs supplied per order are appended after each order's
generated commission detail page.
"""
from __future__ import annotations
import io, datetime
import openpyxl
from money import r2
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
            rows.append([o["order_number"], o["customer"], o.get("finance_type"), rl.get("commission_type"),
                         o["net_commissionable"], rl["rep_id"], rl["rep_name"], rl["share"],
                         rl.get("effective_rate"), rl.get("delivery_factor"), rl.get("collection_factor"),
                         (rl.get("full_commission") if rl.get("full_commission") is not None else rl.get("earned_to_date")),
                         rl.get("prior_processed"), rl.get("this_period")])
    end = _sheet_table(ws, 1,
                       ["Order", "Customer", "Finance Type", "Commission Type", "Net Commissionable",
                        "Rep ID", "Rep Name", "Share", "Eff. Rate", "Delivery", "Collection",
                        "Gross Commission", "Prior Processed", "This Period"],
                       rows, widths=[13, 22, 16, 16, 15, 9, 18, 8, 9, 9, 10, 15, 14, 14])
    for r in range(2, end):
        for col in (5, 12, 13, 14):
            ws.cell(r, col).number_format = CUR
        for col in (8, 9, 10, 11):
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
    snap = [["Commission base", "Pre-tax subtotal (devices + discount + customer shipping); tax excluded"],
            ["Net commissionable", "Subtotal − deductions"],
            ["Standard commission", "Net × matrix rate × share, gated by delivery and collection"],
            ["Delivery release", "Bundle 50/50 (PICO backorder): 0% none, 50% first device, 100% both"],
            ["Collection", "Pays on cash actually cleared (capped at net), gated by the delivered half — "
                           "rate × min(cleared capped at net, net × delivery)"],
            ["In-House Financing", "6% of each cleared payment (down payment + monthlies); no delivery holdback"],
            ["Quarter Accelerator", "Per rep, ≥ 6 straight-purchase devices in the quarter → 16% (retroactive)"],
            ["Super Kicker", "Per rep, ≥ $1,000,000 straight-purchase volume in the quarter → 22% (retroactive)"],
            ["Rounding", "Half up to the cent"]]
    _sheet_table(ws, 1, ["Policy", "Definition"], snap, widths=[26, 80])

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


# ---------------- per-rep PDF packet ----------------
import os
GREEN = "#38A828"        # Boston Aesthetics brand green
GREEN_DK = "#2C7A20"
INK = "#222222"
GREY = "#6B6B6B"
LIGHTG = "#EAF5E7"
LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ba_logo.png")


def _to_date(x):
    if isinstance(x, datetime.datetime):
        return x.date()
    if isinstance(x, datetime.date):
        return x
    if isinstance(x, str) and x.strip():
        try:
            return datetime.date.fromisoformat(x.strip()[:10])
        except ValueError:
            return None
    return None


def _fmt_long_date(x):
    d = _to_date(x)
    return f"{d:%B} {d.day}, {d.year}" if d else (str(x) if x else "")


def _fmt_period(a, b):
    da, db = _to_date(a), _to_date(b)
    if da and db:
        if da.year == db.year:
            return f"{da:%B} {da.day} to {db:%B} {db.day}, {db.year}"
        return f"{da:%B} {da.day}, {da.year} to {db:%B} {db.day}, {db.year}"
    return f"{a} to {b}"


def _est_today():
    try:
        from zoneinfo import ZoneInfo
        now = datetime.datetime.now(ZoneInfo("America/New_York"))
    except Exception:
        now = datetime.datetime.utcnow()
    return f"{now:%B} {now.day}, {now.year}"


def render_detail_one_page(xlsx_bytes: bytes) -> bytes | None:
    """Render an order commission-detail workbook to a single-page PDF, excluding the
    STATUS / Template Version tail (and anything below it) so nothing bleeds to page 2.
    Returns None if LibreOffice isn't available (caller falls back to a generated page)."""
    import io as _io
    import os as _os
    import shutil
    import subprocess
    import tempfile
    import openpyxl
    from openpyxl.worksheet.properties import PageSetupProperties

    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return None
    try:
        wb = openpyxl.load_workbook(_io.BytesIO(xlsx_bytes))
        ws = wb["Commission Detail"] if "Commission Detail" in wb.sheetnames else wb.active
        # find where the STATUS row starts so we can cut above it
        cut = None
        for nm in ("OrderStatus", "TemplateVersion"):
            dn = wb.defined_names.get(nm)
            if dn:
                for sh, co in dn.destinations:
                    co = (co or "").replace("$", "")
                    if co:
                        from openpyxl.utils.cell import coordinate_to_tuple
                        r = coordinate_to_tuple(co)[0]
                        cut = r if cut is None else min(cut, r)
        last_row = (cut - 1) if cut else ws.max_row
        # print area stops above the STATUS/version tail, so it never reaches page 2
        ws.print_area = "A1:E%d" % last_row
        ws.page_setup.orientation = "portrait"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_margins.left = ws.page_margins.right = 0.4
        ws.page_margins.top = ws.page_margins.bottom = 0.4
        with tempfile.TemporaryDirectory() as td:
            xp = _os.path.join(td, "detail.xlsx")
            wb.save(xp)
            prof = "-env:UserInstallation=file://" + _os.path.join(td, "lo_profile")
            subprocess.run([soffice, prof, "--headless", "--calc", "--convert-to", "pdf",
                            "--outdir", td, xp], check=True, timeout=120,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            pp = _os.path.join(td, "detail.pdf")
            if _os.path.exists(pp):
                with open(pp, "rb") as fh:
                    return fh.read()
    except Exception:
        return None
    return None


def build_rep_packet(rep_id, run, lps_by_order: dict, detail_by_order: dict = None,
                     company="Boston Aesthetics") -> bytes:
    """Branded cover statement + (attached exact commission-detail PDF, or a generated
    standard version if none supplied) + each order's LPS, per rep."""
    from reportlab.lib.pagesizes import LETTER
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                    HRFlowable, Image as RLImage)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from pypdf import PdfReader, PdfWriter

    detail_by_order = detail_by_order or {}
    p = run["payload"]
    order_results = p.get("results", {}).get("orders", [])
    reps_master = p.get("reps_master", {})
    orders_raw = {str(o["order_number"]): o for o in p.get("orders", [])}
    info = reps_master.get(str(rep_id), {})
    name = info.get("name", rep_id)

    styles = getSampleStyleSheet()
    H1 = ParagraphStyle("H1", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=18,
                        leading=21, textColor=colors.HexColor(GREEN_DK))
    SUBT = ParagraphStyle("SUBT", parent=styles["Normal"], fontName="Helvetica", fontSize=9.5,
                          textColor=colors.HexColor(GREY), spaceBefore=1)
    SEC = ParagraphStyle("SEC", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10.5,
                         textColor=colors.HexColor(GREEN_DK), spaceBefore=9, spaceAfter=3)
    BODY = ParagraphStyle("BODY", parent=styles["Normal"], fontSize=8.5, leading=11)
    SMALL = ParagraphStyle("SMALL", parent=styles["Normal"], fontSize=7.5, textColor=colors.HexColor(GREY))

    def money(x):
        if x in (None, ""):
            return "—"
        return "-${:,.2f}".format(-x) if x < 0 else "${:,.2f}".format(x)

    def pct(x):
        return "—" if x in (None, "") else "{:.1f}%".format(x * 100)

    def rule():
        return HRFlowable(width="100%", thickness=1.4, color=colors.HexColor(GREEN),
                          spaceBefore=3, spaceAfter=8)

    def kv_table(rows, col_w):
        t = Table(rows, colWidths=col_w)
        t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"), ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
            ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor(GREEN_DK)),
            ("TEXTCOLOR", (2, 0), (2, -1), colors.HexColor(GREEN_DK)),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        return t

    # ---- rep aggregates ----
    lines = []
    for o in order_results:
        for rl in o["reps"]:
            if str(rl["rep_id"]) == str(rep_id):
                lines.append((o, rl))
    sf = p.get("statement_fields", {})
    rep_sf = (sf.get("reps", {}) or {}).get(str(rep_id), {})
    title = info.get("title") or rep_sf.get("title") or "________________"
    territory = info.get("territory") or rep_sf.get("territory") or "________________"
    payroll_date = sf.get("payroll_date") or "________________"
    quarterly_rev = rep_sf.get("quarterly_revenue") or "—"
    comments = rep_sf.get("comments") or ""

    total_comm_full = 0.0        # sum of Total Commission (full deal, this rep)
    commissionable_total = 0.0
    devices_sold = 0
    tiers = set()
    detail_rows = []
    wf = []                      # holdback waterfall: (order, full, deliv%, cash_cleared, dhb, chb, earned)
    dhb_total = chb_total = cleared_total = 0.0
    for o, rl in lines:
        share = rl["share"] or 0
        inhouse = o.get("finance_type") == "In-House Financed Purchase"
        if inhouse:
            commissionable = r2((o.get("cleared") or 0) * share)
            full = rl.get("earned_to_date") or 0
        else:
            commissionable = r2((o["net_commissionable"] or 0) * share)
            full = rl.get("full_commission") or 0
        earned = rl.get("earned_to_date") or 0
        dhb = rl.get("delivery_holdback") or 0
        chb = rl.get("collection_holdback") or 0
        cash_cleared = r2((o.get("cleared") or 0) * share)
        commissionable_total = r2(commissionable_total + commissionable)
        total_comm_full = r2(total_comm_full + full)
        dhb_total = r2(dhb_total + dhb); chb_total = r2(chb_total + chb)
        cleared_total = r2(cleared_total + cash_cleared)
        devices_sold += 2 if "bundle" in str(o.get("configuration") or "").lower() else 1
        tiers.add(rl.get("commission_type"))
        ctype_display = o.get("release_rule") or rl.get("commission_type") or "—"
        detail_rows.append([(o.get("customer") or "")[:44], o["order_number"], ctype_display,
                            money(commissionable), pct(rl.get("effective_rate")), money(full)])
        wf.append((o["order_number"], full, rl.get("delivery_factor"), cash_cleared, dhb, chb, earned))
    prior = r2(sum((rl["prior_processed"] or 0) for _, rl in lines))
    earned_total = r2(sum((rl.get("earned_to_date") or 0) for _, rl in lines))
    to_pay = r2(sum((rl["this_period"] or 0) for _, rl in lines))
    if "Super Kicker" in tiers:
        tier_achieved = "Super Kicker (22%)"
    elif "Quarter Accelerator" in tiers:
        tier_achieved = "Quarter Accelerator (16%)"
    elif tiers == {"In-House Financing"}:
        tier_achieved = "In-House Financing"
    else:
        tier_achieved = "Standard"
    accel_earned = "Yes" if (tiers & {"Super Kicker", "Quarter Accelerator"}) else "No"

    gen = io.BytesIO()
    doc = SimpleDocTemplate(gen, pagesize=LETTER, leftMargin=0.6 * inch, rightMargin=0.6 * inch,
                            topMargin=0.5 * inch, bottomMargin=0.5 * inch,
                            title=f"Commission Statement — {name}")
    S = []

    # ---- letterhead: logo + title ----
    title_cell = [Paragraph("Sales Commission Statement", H1),
                  Paragraph("Pre-Payroll Review", SUBT)]
    if os.path.exists(LOGO_PATH):
        logo = RLImage(LOGO_PATH, width=1.85 * inch, height=1.85 * inch * 156.0 / 512.0)
        head = Table([[logo, title_cell]], colWidths=[2.1 * inch, 5.2 * inch])
    else:
        head = Table([[Paragraph(company, H1), title_cell]], colWidths=[2.1 * inch, 5.2 * inch])
    head.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                              ("ALIGN", (1, 0), (1, 0), "RIGHT"),
                              ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    S.append(head)
    S.append(rule())

    # ---- employee information ----
    S.append(Paragraph("Employee Information", SEC))
    emp = [["Employee Name", f"{name} ({rep_id})", "Commission Period", _fmt_period(run["period_start"], run["period_end"])],
           ["Title", title, "Payroll Date", _fmt_long_date(payroll_date) if _to_date(payroll_date) else payroll_date],
           ["Territory", Paragraph(str(territory), BODY), "Statement Date", _est_today()]]
    S.append(kv_table(emp, [1.15 * inch, 2.55 * inch, 1.2 * inch, 2.4 * inch]))

    # ---- one consolidated commission summary (deal + holdbacks + earned) ----
    S.append(Paragraph("Commission Summary", SEC))

    def hb(x):
        return "(${:,.2f})".format(x) if (x and x > 0.005) else "—"

    CELL = ParagraphStyle("cell", parent=BODY, fontSize=8, leading=9.5)
    wf_by_order = {w[0]: w for w in wf}
    srows = [["Order / Customer", "Commission\nType", "Total\nCommission",
              "Delivery\nHoldback", "Collection\nHoldback", "Earned\nThis Period"]]
    for m in detail_rows:
        cust, onum, ctype = m[0], m[1], m[2]
        _, full, _, _, dhb, chb, earned = wf_by_order[onum]
        srows.append([Paragraph(f"<b>{onum}</b><br/>{cust}", CELL), ctype,
                      money(full), hb(dhb), hb(chb), money(earned)])
    nrows = len(detail_rows)
    srows.append(["TOTAL", "", money(total_comm_full), hb(dhb_total), hb(chb_total), money(earned_total)])
    srows.append(["Less: Prior Period Adjustments", "", "", "", "", money(-prior) if prior else "$0.00"])
    srows.append(["Total Commission to be Paid", "", "", "", "", money(to_pay)])
    st = Table(srows, colWidths=[2.15 * inch, 1.0 * inch, 1.02 * inch, 1.02 * inch, 1.08 * inch, 1.03 * inch])
    total_r = nrows + 1
    less_r = nrows + 2
    pay_r = nrows + 3
    st.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(GREEN)), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, nrows), [colors.white, colors.HexColor(LIGHTG)]),
        ("ALIGN", (2, 0), (-1, -1), "RIGHT"), ("ALIGN", (0, 0), (1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, total_r), 0.25, colors.HexColor("#DDDDDD")),
        ("FONTNAME", (0, total_r), (-1, total_r), "Helvetica-Bold"),
        ("LINEABOVE", (0, total_r), (-1, total_r), 0.7, colors.black),
        ("TEXTCOLOR", (3, 1), (4, total_r), colors.HexColor("#B00000")),   # holdbacks in red
        ("SPAN", (0, less_r), (4, less_r)), ("SPAN", (0, pay_r), (4, pay_r)),
        ("FONTNAME", (0, pay_r), (-1, pay_r), "Helvetica-Bold"),
        ("BACKGROUND", (0, pay_r), (-1, pay_r), colors.HexColor(LIGHTG)),
        ("TEXTCOLOR", (0, pay_r), (-1, pay_r), colors.HexColor(GREEN_DK)),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    S.append(st)
    S.append(Paragraph("<b>Total Commission</b> is the full commission on the deal (net × rate × your share). "
                       "The <b>Delivery Holdback</b> (50% while the Boston Pico is on backorder) releases when the Pico "
                       "ships; the <b>Collection Holdback</b> releases as the customer's balance clears. "
                       "Total Commission − holdbacks = <b>Earned This Period</b>.", SMALL))

    # ---- quarter performance summary ----
    S.append(Paragraph("Quarter Performance Summary", SEC))
    perf = [["Quarterly Revenue", str(quarterly_rev), "Commission Tier Achieved", tier_achieved],
            ["Devices Sold", str(devices_sold), "Accelerator Earned", accel_earned]]
    S.append(kv_table(perf, [1.55 * inch, 2.1 * inch, 1.9 * inch, 1.75 * inch]))

    # ---- comments ----
    S.append(Paragraph("Comments", SEC))
    cm = Table([[Paragraph(comments, BODY) if comments else ""]], colWidths=[7.3 * inch],
               rowHeights=None if comments else [0.3 * inch])
    cm.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor(GREEN)),
                            ("LEFTPADDING", (0, 0), (-1, -1), 6), ("TOPPADDING", (0, 0), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    S.append(cm)

    # ---- employee review ----
    S.append(Paragraph("Employee Review", SEC))
    for para in [
        "This statement is provided prior to payroll to allow you to review the commissions scheduled for payment. "
        "If you believe there is an error or have questions regarding any transaction, please notify your manager and "
        "Human Resources within one (1) business day of receiving this statement.",
        "Unless otherwise provided in your commission agreement, commissions are based on collected revenue and are "
        "subject to adjustments for returns, credits, cancellations, chargebacks, pricing corrections, or other "
        "applicable deductions.",
        "Receipt of this statement does not alter the terms of your Compensation Plan or Employment Agreement."]:
        S.append(Paragraph(para, BODY)); S.append(Spacer(1, 3))
    S.append(Spacer(1, 6))
    sig = Table([["Prepared By:", "_____________________________", "Date:", "________________"],
                 ["Reviewed By:", "_____________________________", "Date:", "________________"]],
                colWidths=[1.0 * inch, 3.0 * inch, 0.5 * inch, 2.0 * inch])
    sig.setStyle(TableStyle([("FONTSIZE", (0, 0), (-1, -1), 9), ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                             ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"),
                             ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor(GREEN_DK)),
                             ("TOPPADDING", (0, 0), (-1, -1), 9)]))
    S.append(sig)
    doc.build(S)

    # ---- fallback "standard" order detail (used only if no exact detail PDF supplied) ----
    def order_detail_pdf(o, rl):
        raw = orders_raw.get(o["order_number"], {})
        buf = io.BytesIO()
        d2 = SimpleDocTemplate(buf, pagesize=LETTER, leftMargin=0.6 * inch, rightMargin=0.6 * inch,
                               topMargin=0.5 * inch, bottomMargin=0.5 * inch)
        F = [Paragraph(f"Order {o['order_number']} — {o.get('customer','')}", H1),
             Paragraph(f"{o.get('finance_type','')} &nbsp;·&nbsp; {o.get('configuration','')} &nbsp;·&nbsp; "
                       f"{o.get('release_rule') or rl.get('commission_type','')}", SUBT), rule(),
             Paragraph("Order Commission Detail (standard)", SEC)]
        li = [["Line item", "Amount"]]
        for l in raw.get("line_items", []):
            li.append([(l.get("description") or l.get("type") or ""), money(l.get("amount"))])
        li.append(["Subtotal before tax", money(raw.get("subtotal_before_tax"))])
        li.append(["Sales tax (excluded from commission)", money(raw.get("sales_tax"))])
        li.append(["Invoice total (customer)", money(raw.get("invoice_total"))])
        for dd in raw.get("deductions", []):
            li.append([f"less: {dd.get('description', '')}", money(-(dd.get('amount') or 0))])
        li.append(["Net commissionable", money(o["net_commissionable"])])
        t1 = Table(li, colWidths=[5.0 * inch, 2.3 * inch])
        t1.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(GREEN)), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("LINEABOVE", (0, -1), (-1, -1), 0.7, colors.black),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#DDDDDD")), ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5)]))
        F.append(t1); F.append(Spacer(1, 8)); F.append(Paragraph("This-Period Commission", SEC))
        share = rl.get("share")
        applied = r2((o.get("payable_base") or 0) * (share or 0)) if o.get("finance_type") != "In-House Financed Purchase" \
            else r2((o.get("cleared") or 0) * (share or 0))
        cb = [["Commission type", o.get("release_rule") or rl.get("commission_type", "")],
              ["Commission rate", pct(rl.get("effective_rate"))],
              [f"Your share", pct(share)],
              ["Delivery release", pct(rl.get("delivery_factor"))],
              ["Cash cleared this period", money(o.get("cleared"))],
              ["Cash applied (your share, capped by delivery)", money(applied)],
              ["Earned = cash applied × rate", money(rl.get("earned_to_date"))]]
        if rl.get("prior_processed"):
            cb.append(["Less: prior processed", money(-rl["prior_processed"])])
        cb.append(["Commission Earned — This Period", money(rl.get("this_period"))])
        t2 = Table(cb, colWidths=[5.0 * inch, 2.3 * inch])
        t2.setStyle(TableStyle([
            ("FONTSIZE", (0, 0), (-1, -1), 9), ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("LINEBELOW", (0, 0), (-1, -2), 0.25, colors.HexColor("#DDDDDD")),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor(LIGHTG)), ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("LINEABOVE", (0, -1), (-1, -1), 0.8, colors.black),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
        F.append(t2); d2.build(F)
        return buf.getvalue()

    writer = PdfWriter()
    for pg in PdfReader(io.BytesIO(gen.getvalue())).pages:
        writer.add_page(pg)
    for o, rl in lines:
        onum = o["order_number"]
        det = detail_by_order.get(onum)
        try:
            if det:
                for pg in PdfReader(io.BytesIO(det)).pages:
                    writer.add_page(pg)
            else:
                for pg in PdfReader(io.BytesIO(order_detail_pdf(o, rl))).pages:
                    writer.add_page(pg)
        except Exception:
            pass
        lps = lps_by_order.get(onum)
        if lps:
            try:
                for pg in PdfReader(io.BytesIO(lps)).pages:
                    writer.add_page(pg)
            except Exception:
                pass
    out = io.BytesIO(); writer.write(out); return out.getvalue()
