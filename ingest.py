"""Read the Boston Aesthetics workbooks by NAME (named cells / tables), never by
cell position, and recompute every total from raw inputs (never trust cached
Excel formula results). Matches BA-COMM-SCHEMA 1.2 / template BA-COMM-DETAIL v1.x."""
from __future__ import annotations
import io, datetime
import openpyxl
from money import r2
from openpyxl.utils import range_boundaries

DEVICE_TYPES = {"ZenTite", "Boston Pico"}


def _named_value(wb, name):
    dn = wb.defined_names.get(name)
    if dn is None:
        return None
    for sheet, coord in dn.destinations:
        return wb[sheet][coord.replace("$", "")].value
    return None


def _named_range_rows(wb, name):
    dn = wb.defined_names.get(name)
    if dn is None:
        return []
    for sheet, coord in dn.destinations:
        ws = wb[sheet]
        minc, minr, maxc, maxr = range_boundaries(coord.replace("$", ""))
        return [[ws.cell(r, c).value for c in range(minc, maxc + 1)] for r in range(minr, maxr + 1)]
    return []


def _table_rows(ws, name):
    t = ws.tables.get(name)
    if t is None:
        return [], []
    minc, minr, maxc, maxr = range_boundaries(t.ref)
    headers = [ws.cell(minr, c).value for c in range(minc, maxc + 1)]
    rows = [[ws.cell(r, c).value for c in range(minc, maxc + 1)] for r in range(minr + 1, maxr + 1)]
    return headers, rows


def _find_table_anywhere(wb, name):
    for ws in wb.worksheets:
        if name in ws.tables:
            return ws
    return None


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return 0.0


def parse_order_workbook(data: bytes, filename: str = "") -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    ws = _find_table_anywhere(wb, "tblLineItems") or wb.active

    order = {
        "source_file": filename,
        "order_number": _named_value(wb, "OrderNumber"),
        "customer": _named_value(wb, "Customer"),
        "deal_type": _named_value(wb, "DealType"),
        "release_rule": _named_value(wb, "ReleaseRule"),
        "prepared_by": _named_value(wb, "PreparedBy"),
        "finalized_date": str(_named_value(wb, "FinalizedDate") or ""),
        "notes": _named_value(wb, "Notes"),
        "template_version": _named_value(wb, "TemplateVersion"),
        "sales_tax": _num(_named_value(wb, "SalesTax")),
        "commission_rate": _named_value(wb, "CommissionRate"),  # may be None
    }

    # line items (Type, Part Number, Description, Amount)
    _, li = _table_rows(ws, "tblLineItems")
    lines = []
    for row in li:
        typ, part, desc, amt = (row + [None, None, None, None])[:4]
        if typ in (None, "") and amt in (None, ""):
            continue
        lines.append({"type": typ, "part_number": part, "description": desc, "amount": _num(amt)})
    order["line_items"] = lines

    # deductions (named range: Description col 0, Amount col 3)
    ded = []
    for row in _named_range_rows(wb, "DeductionsRange")[1:]:  # skip header row
        desc = row[0] if row else None
        amt = row[3] if len(row) > 3 else None
        if (desc in (None, "")) and (amt in (None, "")):
            continue
        ded.append({"description": desc, "amount": _num(amt)})
    order["deductions"] = ded

    # rep terms (Rep ID, Share %, Commission[ignore], Note)
    reps = []
    rws = _find_table_anywhere(wb, "tblRepTerms")
    if rws is not None:
        _, rr = _table_rows(rws, "tblRepTerms")
        for row in rr:
            rid, share, _comm, note = (row + [None, None, None, None])[:4]
            if rid in (None, "") and note in (None, ""):
                continue
            reps.append({"rep_id": rid, "share": (None if share in (None, "") else _num(share)), "note": note})
    order["reps"] = reps

    # recompute totals from raw inputs
    subtotal = r2(sum(l["amount"] for l in lines))
    total_ded = r2(sum(d["amount"] for d in ded))
    net = r2(subtotal - total_ded)
    rate = order["commission_rate"]
    rate = None if rate in (None, "") else _num(rate)
    order["commission_rate"] = rate
    order["subtotal_before_tax"] = subtotal
    order["total_deductions"] = total_ded
    order["net_commissionable"] = net
    order["invoice_total"] = r2(subtotal + order["sales_tax"])
    order["deal_commission"] = r2(net * rate) if rate is not None else None
    order["device_types"] = [l["type"] for l in lines if l["type"] in DEVICE_TYPES]
    return order


def parse_delivery_register(data: bytes) -> list[dict]:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = _find_table_anywhere(wb, "tblDeliveryRegister")
    if ws is None:
        ws = wb["Delivery Register"] if "Delivery Register" in wb.sheetnames else wb.active
        headers, rows = None, [[c.value for c in r] for r in ws.iter_rows(min_row=1)]
    else:
        _, rows = _table_rows(ws, "tblDeliveryRegister")
    out = []
    for row in rows:
        onum, part, product, ddate = (list(row) + [None, None, None, None])[:4]
        if onum in (None, "") and product in (None, ""):
            continue
        if isinstance(ddate, datetime.datetime):
            ddate = ddate.date().isoformat()
        elif ddate:
            ddate = str(ddate)
        out.append({"order_number": onum, "part_number": part, "product": product,
                    "delivery_date": ddate or None})
    return out


def parse_setup_reps(data: bytes) -> dict:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = _find_table_anywhere(wb, "tblReps")
    reps = {}
    if ws is not None:
        _, rows = _table_rows(ws, "tblReps")
        for row in rows:
            rid, payroll, name, email = (list(row) + [None, None, None, None])[:4]
            if rid in (None, ""):
                continue
            reps[str(rid)] = {"payroll_id": payroll, "name": name, "email": email}
    return reps


def parse_opening_history(data: bytes) -> list[dict]:
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    ws = _find_table_anywhere(wb, "tblOpeningHistory")
    out = []
    if ws is not None:
        _, rows = _table_rows(ws, "tblOpeningHistory")
        for row in rows:
            onum, rid, ref, amt, note = (list(row) + [None] * 5)[:5]
            if onum in (None, "") and rid in (None, ""):
                continue
            out.append({"order_number": onum, "rep_id": rid, "reference": ref,
                        "amount": _num(amt), "note": note})
    return out
