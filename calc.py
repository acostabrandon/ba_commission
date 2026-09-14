"""Commission calculation engine.

Recomputes from raw inputs, applies the confirmed bundle 50/50 delivery release,
subtracts amounts already finalized in prior runs (opening history), and produces
per-order and per-rep results plus a list of exceptions for review.
"""
from __future__ import annotations
from money import r2


def _delivered_products(order_number, deliveries):
    """Set of products with an actual delivery date for this order."""
    out = set()
    for d in deliveries:
        if str(d.get("order_number")) == str(order_number) and d.get("delivery_date"):
            out.add(d.get("product"))
    return out


def release_fraction(order, deliveries):
    """Cumulative eligible fraction of the whole deal commission (0, .5, 1)."""
    delivered = _delivered_products(order["order_number"], deliveries)
    rule = order.get("release_rule")
    deal = order.get("deal_type")
    if rule == "Bundle 50/50" or deal == "Bundle":
        n = sum(1 for p in ("ZenTite", "Boston Pico") if p in delivered)
        return {0: 0.0, 1: 0.5, 2: 1.0}[min(n, 2)]
    # Full delivery / single device: released once the single device is delivered
    return 1.0 if delivered else 0.0


def _prior_for(order_number, rep_id, opening_history):
    return round(sum(h["amount"] for h in opening_history
                     if str(h.get("order_number")) == str(order_number)
                     and str(h.get("rep_id")) == str(rep_id)), 2)


def calculate(orders, deliveries, reps_master, opening_history, overrides=None):
    """Return (order_results, rep_summary, exceptions).

    overrides: {order_number: {"commission_rate": float, "reps": {rep_id: share}}}
    lets the review UI adjust rate/share without touching the source workbook.
    """
    overrides = overrides or {}
    order_results, exceptions = [], []
    rep_summary = {}  # rep_id -> {"name":, "this_period":, "full":, "lines":[]}

    for o in orders:
        onum = o["order_number"]
        ov = overrides.get(str(onum), {})
        rate = ov.get("commission_rate", o.get("commission_rate"))
        rep_shares = ov.get("reps")  # optional {rep_id: share}

        rel = release_fraction(o, deliveries)
        net = o["net_commissionable"]
        deal_comm = r2(net * rate) if rate is not None else None

        # exceptions -------------------------------------------------------
        if rate is None:
            exceptions.append(f"{onum}: commission rate not set")
        if o.get("deal_type") == "Bundle":
            dts = set(o.get("device_types", []))
            if not ({"ZenTite", "Boston Pico"} <= dts):
                exceptions.append(f"{onum}: Bundle is missing a ZenTite or Boston Pico line")
        if net < 0:
            exceptions.append(f"{onum}: net commissionable is negative")

        # reps -------------------------------------------------------------
        reps = []
        if rep_shares is not None:
            reps = [{"rep_id": rid, "share": sh} for rid, sh in rep_shares.items()]
        else:
            reps = [{"rep_id": r["rep_id"], "share": r["share"]} for r in o.get("reps", []) if r.get("rep_id")]

        share_sum = round(sum((r["share"] or 0) for r in reps), 6)
        if reps and share_sum != 1.0:
            exceptions.append(f"{onum}: rep shares total {share_sum:.4f} (should be 1.0)")
        if not reps:
            exceptions.append(f"{onum}: no reps entered")

        line_results = []
        for r in reps:
            rid = str(r["rep_id"])
            share = r["share"]
            name = reps_master.get(rid, {}).get("name", rid)
            full = r2(deal_comm * share) if (deal_comm is not None and share is not None) else None
            eligible = r2(full * rel) if full is not None else None
            prior = _prior_for(onum, rid, opening_history)
            this_period = r2((eligible or 0) - prior) if eligible is not None else None
            lr = {"rep_id": rid, "rep_name": name, "share": share,
                  "full_commission": full, "release_fraction": rel,
                  "eligible_to_date": eligible, "prior_processed": prior,
                  "this_period": this_period}
            line_results.append(lr)
            if this_period is not None:
                acc = rep_summary.setdefault(rid, {"name": name, "this_period": 0.0, "full": 0.0, "lines": []})
                acc["this_period"] = r2(acc["this_period"] + this_period)
                acc["full"] = r2(acc["full"] + (full or 0))
                acc["lines"].append({"order_number": onum, "customer": o.get("customer"),
                                     "this_period": this_period, "full_commission": full,
                                     "release_fraction": rel})

        order_results.append({
            "order_number": onum, "customer": o.get("customer"),
            "deal_type": o.get("deal_type"), "release_rule": o.get("release_rule"),
            "net_commissionable": net, "commission_rate": rate,
            "deal_commission": deal_comm, "release_fraction": rel,
            "period_commission": r2((deal_comm or 0) * rel) if deal_comm is not None else None,
            "reps": line_results,
        })

    return order_results, rep_summary, exceptions


def control_totals(order_results):
    return {
        "orders": len(order_results),
        "sum_deal_commission": r2(sum(o["deal_commission"] or 0 for o in order_results)),
        "sum_period_commission": r2(sum(o["period_commission"] or 0 for o in order_results)),
    }
