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
NAVY = "#1F3A5F"
DARK = "#333333"
LIGHT = "#EFEFEF"


def build_rep_packet(rep_id, run, lps_by_order: dict, company="Boston Aesthetics") -> bytes:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                    PageBreak, HRFlowable)
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_RIGHT
    from pypdf import PdfReader, PdfWriter

    p = run["payload"]
    order_results = p.get("results", {}).get("orders", [])
    reps_master = p.get("reps_master", {})
    orders_raw = {str(o["order_number"]): o for o in p.get("orders", [])}
    info = reps_master.get(str(rep_id), {})
    name = info.get("name", rep_id)
    today = datetime.date.today().isoformat()

    styles = getSampleStyleSheet()
    H1 = ParagraphStyle("H1", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=17, leading=21, spaceAfter=3, textColor=colors.HexColor(NAVY))
    SUBT = ParagraphStyle("SUBT", parent=styles["Normal"], fontName="Helvetica", fontSize=10, textColor=colors.HexColor(DARK))
    SEC = ParagraphStyle("SEC", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=11, textColor=colors.HexColor(NAVY), spaceBefore=7, spaceAfter=2)
    BODY = ParagraphStyle("BODY", parent=styles["Normal"], fontSize=8.5, leading=11)
    SMALL = ParagraphStyle("SMALL", parent=styles["Normal"], fontSize=7.5, textColor=colors.grey)

    def money(x):
        if x in (None, ""):
            return "—"
        return "-${:,.2f}".format(-x) if x < 0 else "${:,.2f}".format(x)

    def pct(x):
        return "—" if x in (None, "") else "{:.1f}%".format(x * 100)

    def rule():
        return HRFlowable(width="100%", thickness=0.6, color=colors.HexColor(NAVY), spaceBefore=1, spaceAfter=6)

    def kv_table(rows, col_w):
        t = Table(rows, colWidths=col_w)
        t.setStyle(TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold") if len(col_w) == 4 else ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("TEXTCOLOR", (0, 0), (0, -1), colors.HexColor(DARK)),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("LINEBELOW", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9D9D9")),
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
    territory = rep_sf.get("territory") or "________________"
    manager = rep_sf.get("manager") or "________________"
    payroll_date = sf.get("payroll_date") or "________________"
    quarterly_rev = rep_sf.get("quarterly_revenue") or "—"
    comments = rep_sf.get("comments") or ""
    FIN_SHORT = {"Straight Purchase": "Straight", "Financed Purchase": "Financed",
                 "In-House Financed Purchase": "In-House"}

    linemx = []
    total_comm = gross = deliv_hold_t = coll_hold_t = collected_cleared = 0.0
    devices_sold = 0
    tiers = set()
    for o, rl in lines:
        share = rl["share"] or 0
        inhouse = o.get("finance_type") == "In-House Financed Purchase"
        if inhouse:
            commissionable = r2((o.get("cleared") or 0) * share)
            gross_i = rl.get("earned_to_date") or 0
            dh = ch = 0.0
        else:
            commissionable = r2((o["net_commissionable"] or 0) * share)
            full = rl.get("full_commission") or 0
            gross_i = full
            # prefer explicit holdback amounts from the engine (cleared-cash model)
            dh = rl.get("delivery_holdback")
            ch = rl.get("collection_holdback")
            if dh is None:
                dfac = rl["delivery_factor"] if rl.get("delivery_factor") is not None else 1.0
                dh = r2(full * (1 - dfac))
            if ch is None:
                ch = 0.0
        total_comm = r2(total_comm + commissionable); gross = r2(gross + gross_i)
        deliv_hold_t = r2(deliv_hold_t + dh); coll_hold_t = r2(coll_hold_t + ch)
        collected_cleared = r2(collected_cleared + (o.get("cleared") or 0) * share)
        devices_sold += 2 if "bundle" in str(o.get("configuration") or "").lower() else 1
        tiers.add(rl.get("commission_type"))
        linemx.append(dict(o=o, rl=rl, commissionable=commissionable, gross_i=gross_i,
                           finance_short=FIN_SHORT.get(o.get("finance_type"), o.get("finance_type") or "—"),
                           ctype=rl.get("commission_type") or "—", eff=rl.get("effective_rate")))
    prior = r2(sum((rl["prior_processed"] or 0) for _, rl in lines))
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

    def paren(x):
        return f"(${x:,.2f})" if x and x > 0.005 else "$0.00"

    gen = io.BytesIO()
    doc = SimpleDocTemplate(gen, pagesize=LETTER, leftMargin=0.6 * inch, rightMargin=0.6 * inch,
                            topMargin=0.55 * inch, bottomMargin=0.55 * inch, title=f"Commission Statement — {name}")
    S = []

    # ---- header (mirrors the Word doc) ----
    S.append(Paragraph(f"{company} &nbsp;&nbsp;·&nbsp;&nbsp; Commissions Statement", SMALL))
    S.append(Paragraph("Sales Commission Statement (Pre-Payroll Review)", H1))
    S.append(rule())

    # ---- employee information ----
    S.append(Paragraph("Employee Information", SEC))
    emp = [["Employee Name", f"{name} ({rep_id})", "Commission Period", f"{run['period_start']} – {run['period_end']}"],
           ["Territory", territory, "Payroll Date", payroll_date],
           ["Manager", manager, "Statement Date", today]]
    S.append(kv_table(emp, [1.3 * inch, 2.35 * inch, 1.3 * inch, 2.35 * inch]))

    # ---- commission detail (per order; finance type, commission type, commissionable, rate, earned) ----
    S.append(Paragraph("Commission Detail", SEC))
    rows = [["Customer", "Order", "Finance", "Comm. Type", "Commissionable", "Rate", "Earned"]]
    for m in linemx:
        rows.append([(m["o"]["customer"] or "")[:26], m["o"]["order_number"], m["finance_short"],
                     m["ctype"], money(m["commissionable"]), pct(m["eff"]), money(m["gross_i"])])
    rows.append(["", "", "", "", "", "TOTAL", money(gross)])
    t = Table(rows, colWidths=[1.55 * inch, 0.85 * inch, 0.85 * inch, 1.15 * inch, 1.05 * inch, 0.55 * inch, 1.3 * inch])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("LINEABOVE", (0, -1), (-1, -1), 0.7, colors.black),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F6F8FA")]),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9D9D9")),
        ("ALIGN", (4, 0), (-1, -1), "RIGHT"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]))
    S.append(t)

    # ---- commission summary (mirrors the Word doc rows) ----
    S.append(Paragraph("Commission Summary", SEC))
    summ = [["Total Commissionable Revenue", money(total_comm)],
            ["Gross Commission Earned (at 100% delivered + collected)", money(gross)],
            ["Less: Delivery Holdback (Bundle 50/50 — PICO backorder)", paren(deliv_hold_t)],
            ["Less: Collection Holdback (paid on cleared cash only)", paren(coll_hold_t)],
            ["Chargebacks / Returns / Credits", "(—)"],
            ["Prior Period Adjustments", paren(prior)],
            ["Total Commission to be Paid", money(to_pay)]]
    t = Table(summ, colWidths=[5.0 * inch, 2.3 * inch])
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"), ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("LINEBELOW", (0, 0), (-1, -2), 0.25, colors.HexColor("#D9D9D9")),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor(LIGHT)),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("LINEABOVE", (0, -1), (-1, -1), 0.8, colors.black),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]))
    S.append(t)

    # ---- performance summary (Word doc fields) ----
    S.append(Paragraph("Performance Summary", SEC))
    perf = [["Devices Sold", str(devices_sold), "Commission Tier Achieved", tier_achieved],
            ["Collected Revenue (cleared)", money(collected_cleared), "Accelerator Earned", accel_earned],
            ["Quarterly Revenue", quarterly_rev, "", ""]]
    S.append(kv_table(perf, [1.55 * inch, 2.1 * inch, 1.7 * inch, 1.95 * inch]))

    # ---- comments ----
    S.append(Paragraph("Comments", SEC))
    if comments:
        cm = Table([[Paragraph(comments, BODY)]], colWidths=[7.3 * inch])
    else:
        cm = Table([[""]], colWidths=[7.3 * inch], rowHeights=[0.3 * inch])
    cm.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#BFBFBF")),
                            ("LEFTPADDING", (0, 0), (-1, -1), 6), ("TOPPADDING", (0, 0), (-1, -1), 4),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 4), ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    S.append(cm)

    # ---- employee review (verbatim from the Word doc) ----
    S.append(Paragraph("Employee Review", SEC))
    for para in [
        "This statement is provided prior to payroll to allow you to review the commissions scheduled for payment. "
        "If you believe there is an error or have questions regarding any transaction, please notify your manager and "
        "Human Resources within one (1) business day of receiving this statement.",
        "Unless otherwise provided in your commission agreement, commissions are based on collected revenue and are "
        "subject to adjustments for returns, credits, cancellations, chargebacks, pricing corrections, or other "
        "applicable deductions.",
        "Receipt of this statement does not alter the terms of your Compensation Plan or Employment Agreement."]:
        S.append(Paragraph(para, BODY))
        S.append(Spacer(1, 3))
    S.append(Spacer(1, 6))
    sig = Table([["Prepared By:", "_____________________________", "Date:", "________________"],
                 ["Reviewed By:", "_____________________________", "Date:", "________________"]],
                colWidths=[1.0 * inch, 3.0 * inch, 0.5 * inch, 2.0 * inch])
    sig.setStyle(TableStyle([("FONTSIZE", (0, 0), (-1, -1), 9), ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                             ("FONTNAME", (2, 0), (2, -1), "Helvetica-Bold"), ("TOPPADDING", (0, 0), (-1, -1), 9)]))
    S.append(sig)

    doc.build(S)

    # ---- per order: a full Commission Detail page, then that order's LPS ----
    def order_detail_pdf(o, rl):
        raw = orders_raw.get(o["order_number"], {})
        buf = io.BytesIO()
        d2 = SimpleDocTemplate(buf, pagesize=LETTER, leftMargin=0.6 * inch, rightMargin=0.6 * inch,
                               topMargin=0.55 * inch, bottomMargin=0.55 * inch)
        F = [Paragraph(f"Order {o['order_number']} — {o.get('customer','')}", H1),
             Paragraph(f"{o.get('finance_type','')} &nbsp;·&nbsp; {o.get('configuration','')} &nbsp;·&nbsp; "
                       f"Commission type: {rl.get('commission_type','')}", SUBT), rule(),
             Paragraph("Order Commission Detail", SEC)]
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
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"), ("LINEABOVE", (0, -1), (-1, -1), 0.7, colors.black),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D9D9D9")), ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5)]))
        F.append(t1)
        F.append(Spacer(1, 8)); F.append(Paragraph("Commission Calculation", SEC))
        share = rl.get("share")
        full = rl.get("full_commission")
        inhouse = o.get("finance_type") == "In-House Financed Purchase"
        if inhouse:
            cb = [["Commission type", rl.get("commission_type", "")],
                  ["Commission rate", pct(rl.get("effective_rate"))],
                  ["Contract price (customer total)", money(o.get("contract_price"))],
                  ["Cleared payments this period", money(o.get("cleared"))],
                  [f"6% of cleared × your share ({pct(share)})", money(rl.get("earned_to_date"))]]
        else:
            cb = [["Commission type", rl.get("commission_type", "")],
                  ["Commission rate", pct(rl.get("effective_rate"))],
                  [f"Full commission — net × rate × share ({pct(share)})", money(full)],
                  ["Contract price (customer total)", money(o.get("contract_price"))],
                  ["Cash cleared this period", money(o.get("cleared"))],
                  ["Delivery release", pct(rl.get("delivery_factor"))],
                  ["Less: delivery holdback (undelivered device)", money(-(rl.get("delivery_holdback") or 0))],
                  ["Less: collection holdback (awaiting cleared cash)", money(-(rl.get("collection_holdback") or 0))],
                  ["Earned to date (on cleared cash)", money(rl.get("earned_to_date"))]]
        if rl.get("prior_processed"):
            cb.append(["Less: prior processed", money(-rl["prior_processed"])])
        cb.append(["Commission Earned — This Period", money(rl.get("this_period"))])
        t2 = Table(cb, colWidths=[5.0 * inch, 2.3 * inch])
        t2.setStyle(TableStyle([
            ("FONTSIZE", (0, 0), (-1, -1), 9), ("ALIGN", (1, 0), (1, -1), "RIGHT"),
            ("LINEBELOW", (0, 0), (-1, -2), 0.25, colors.HexColor("#D9D9D9")),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor(LIGHT)), ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("LINEABOVE", (0, -1), (-1, -1), 0.8, colors.black),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
        F.append(t2)
        d2.build(F)
        return buf.getvalue()

    writer = PdfWriter()
    for pg in PdfReader(io.BytesIO(gen.getvalue())).pages:
        writer.add_page(pg)
    for o, rl in lines:
        for pg in PdfReader(io.BytesIO(order_detail_pdf(o, rl))).pages:
            writer.add_page(pg)
        lps = lps_by_order.get(o["order_number"])
        if lps:
            try:
                for pg in PdfReader(io.BytesIO(lps)).pages:
                    writer.add_page(pg)
            except Exception:
                pass
    out = io.BytesIO(); writer.write(out); return out.getvalue()
