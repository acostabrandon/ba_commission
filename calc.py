"""Commission calculation engine (Boston Aesthetics comp plan).

Two gates combine on Standard/Financed deals:
  * DELIVERY  — bundle 50/50, held back ONLY because a PICO is backordered. There is
                no delivery register: the sales admin sets each order's delivery
                release (0% / 50% / 100%) in the app (carried in overrides).
  * COLLECTION — commission pays on the CASH ACTUALLY CLEARED IN THIS PERIOD, at the
                rate, capped by the delivered half. Not a percent of the whole deal.
                Only payments whose Date falls within the run's period and that are
                marked cleared are counted.

Per rep, for a Standard/Financed order:
    net              = net commissionable amount
    rel              = delivery fraction (0 / 0.5 / 1.0), set per order in the app
    cleared          = cash cleared IN PERIOD (from the payments block)
    cleared_capped   = min(cleared, net)          # never pay on more than net
    delivered_base   = net * rel                  # commissionable freed by delivery
    payable_base     = min(cleared_capped, delivered_base)
    full             = net * rate * share         # at 100% delivered + collected
    earned_to_date   = rate * payable_base * share
    delivery_holdback   = full - full*rel         # held for undelivered device(s)
    collection_holdback = full*rel - earned_to_date   # delivered but not yet cleared
    this_period      = earned_to_date - prior_processed

Commission types:
  * Standard            -> price-card matrix rate (entered per order)
  * In-House Financing  -> 6% of each cleared (in-period) payment; no delivery holdback
  * Quarter Accelerator -> per rep, >= 6 straight-purchase devices in the quarter -> 16%
  * Super Kicker        -> per rep, >= $1,000,000 straight-purchase volume in the quarter -> 22%
Accelerator/Kicker apply to STRAIGHT purchases only and are retroactive within the run.
"""
from __future__ import annotations
import datetime
from collections import defaultdict
from money import r2

INHOUSE_RATE = 0.06
ACCEL_DEVICES = 6
ACCEL_RATE = 0.16
KICKER_VOLUME = 1_000_000.0
KICKER_RATE = 0.22
STRAIGHT = "Straight Purchase"
FINANCED = "Financed Purchase"
INHOUSE = "In-House Financed Purchase"
DEVICES = ("ZenTite", "Boston Pico")


def is_bundle(order):
    """A bundle = both devices present, or the deal type says so ('ZT + PICO Bundle')."""
    dt = str(order.get("deal_type") or "").lower()
    if "bundle" in dt or ("zt" in dt and "pico" in dt):
        return True
    dtypes = set(order.get("device_types", []))
    return set(DEVICES) <= dtypes


def _parse_date(x):
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


def _in_period(pdate, period):
    """period = (start_iso, end_iso) or None. No period -> always True."""
    if not period:
        return True
    start = _parse_date(period[0]); end = _parse_date(period[1])
    if pdate is None:
        return None            # unknown — caller decides / flags
    if start and pdate < start:
        return False
    if end and pdate > end:
        return False
    return True


def default_release(order):
    """Fallback delivery fraction when the app hasn't set one: bundles assume the
    first (ZenTite) device delivered and the PICO backordered = 50%; singles = 100%."""
    return 0.5 if is_bundle(order) else 1.0


def _prior_for(order_number, rep_id, opening_history):
    return r2(sum(h["amount"] for h in opening_history
                  if str(h.get("order_number")) == str(order_number)
                  and str(h.get("rep_id")) == str(rep_id)))


def _cleared(payments, period, notes=None, onum=""):
    """Sum cleared payments whose Date falls in the period. Cleared-but-undated
    payments are counted and flagged (so the admin can verify)."""
    total = 0.0
    for p in (payments or []):
        if not p.get("cleared"):
            continue
        pdate = _parse_date(p.get("date"))
        ip = _in_period(pdate, period)
        if ip is True:
            total += (p.get("amount") or 0)
        elif ip is None:
            total += (p.get("amount") or 0)
            if notes is not None:
                notes.append(f"{onum}: a cleared payment has no date — counted this period; add a date to be sure")
        # ip False -> outside the period, skip
    return r2(total)


def calculate(orders, reps_master, opening_history, overrides=None, period=None):
    overrides = overrides or {}
    exceptions = []

    # ---- pass 1: resolve each order's inputs ----
    resolved = []
    for o in orders:
        onum = str(o["order_number"])
        ov = overrides.get(onum, {})
        finance = ov.get("finance_type") or o.get("finance_type") or STRAIGHT
        base_rate = ov.get("commission_rate", o.get("commission_rate"))
        base_rate = None if base_rate in (None, "") else float(base_rate)
        contract = ov.get("contract_price") or o.get("contract_price") or o.get("invoice_total") \
            or o.get("net_commissionable") or 0.0
        payments = ov.get("payments") or o.get("payments") or []
        cleared = _cleared(payments, period, exceptions, onum)
        has_payments = len(payments) > 0
        # delivery release: app override wins; else fall back to the confirmed default
        rel = ov.get("release_fraction")
        rel = default_release(o) if rel in (None, "") else float(rel)
        if ov.get("reps") is not None:
            sheet_names = {str(r["rep_id"]): r.get("rep_name") for r in o.get("reps", []) if r.get("rep_id")}
            reps = [{"rep_id": str(rid), "share": sh, "rep_name": sheet_names.get(str(rid))}
                    for rid, sh in ov["reps"].items()]
        else:
            reps = [{"rep_id": str(r["rep_id"]), "share": r["share"], "rep_name": r.get("rep_name")}
                    for r in o.get("reps", []) if r.get("rep_id")]
        resolved.append(dict(o=o, onum=onum, finance=finance, base_rate=base_rate, contract=float(contract),
                             cleared=cleared, has_payments=has_payments, rel=rel, reps=reps,
                             net=o["net_commissionable"], subtotal=o.get("subtotal_before_tax", o["net_commissionable"]),
                             ndev=len(o.get("device_types", []))))

    # ---- per-rep quarter accelerator / super kicker (straight purchases only) ----
    dev_by_rep = defaultdict(float)
    vol_by_rep = defaultdict(float)
    for R in resolved:
        if R["finance"] == STRAIGHT:
            for rp in R["reps"]:
                dev_by_rep[rp["rep_id"]] += R["ndev"]
                vol_by_rep[rp["rep_id"]] += (R["subtotal"] or 0) * (rp["share"] or 0)
    rep_accel = {}
    for rid in set(list(dev_by_rep) + list(vol_by_rep)):
        if vol_by_rep[rid] >= KICKER_VOLUME:
            rep_accel[rid] = (KICKER_RATE, "Super Kicker")
        elif dev_by_rep[rid] >= ACCEL_DEVICES:
            rep_accel[rid] = (ACCEL_RATE, "Quarter Accelerator")

    # ---- pass 2: compute per order / per rep ----
    order_results = []
    rep_summary = {}
    for R in resolved:
        o, onum = R["o"], R["onum"]
        finance, net, rel = R["finance"], R["net"], R["rel"]

        if is_bundle(o) and not ({"ZenTite", "Boston Pico"} <= set(o.get("device_types", []))):
            exceptions.append(f"{onum}: Bundle is missing a ZenTite or Boston Pico line")
        if net is not None and net < 0:
            exceptions.append(f"{onum}: net commissionable is negative")
        if not R["reps"]:
            exceptions.append(f"{onum}: no reps entered")
        ssum = r2(sum((rp["share"] or 0) for rp in R["reps"]))
        if R["reps"] and ssum != 1.0:
            exceptions.append(f"{onum}: rep shares total {ssum:.4f} (should be 1.0)")

        # ---- collection base: cash cleared in period, capped at net ----
        if finance == INHOUSE:
            coll_disp = None
            if not R["has_payments"]:
                exceptions.append(f"{onum}: In-House deal has no payments recorded — no commission accrues yet")
            cleared_capped = None
            payable_base = None
            delivered_base = None
        else:
            if finance in (STRAIGHT, FINANCED) and R["base_rate"] is None:
                exceptions.append(f"{onum}: commission rate not set")
            if not R["has_payments"]:
                exceptions.append(f"{onum}: no payments recorded — nothing has cleared, so no commission this period")
            cleared_capped = r2(min(R["cleared"], net)) if net is not None else r2(R["cleared"])
            delivered_base = r2((net or 0) * rel)
            payable_base = r2(min(cleared_capped, delivered_base))
            coll_disp = None if not net else round(cleared_capped / net, 6)  # % of net cleared in period

        rep_lines = []
        order_deal_comm = 0.0
        for rp in R["reps"]:
            rid, share = rp["rep_id"], (rp["share"] or 0)
            name = reps_master.get(rid, {}).get("name") or rp.get("rep_name") or rid
            prior = _prior_for(onum, rid, opening_history)

            if finance == INHOUSE:
                ctype = "In-House Financing"
                eff = INHOUSE_RATE
                full = None
                earned = r2(INHOUSE_RATE * R["cleared"] * share)   # 6% of cleared (in-period) payments
                dfac, cfac = 1.0, None
                deliv_hold = coll_hold = 0.0
            else:
                if finance == STRAIGHT and rid in rep_accel:
                    eff, ctype = rep_accel[rid]
                else:
                    eff, ctype = (R["base_rate"], "Standard")
                if eff is None:
                    full = earned = None
                    dfac, cfac = rel, None
                    deliv_hold = coll_hold = None
                else:
                    full = r2((net or 0) * eff * share)
                    earned = r2(eff * payable_base * share)
                    if payable_base >= delivered_base:   # delivered portion fully collected
                        coll_hold = 0.0
                        deliv_hold = r2(full - earned)
                    else:
                        delivered_full = r2(full * rel)
                        deliv_hold = r2(full - delivered_full)
                        coll_hold = r2(delivered_full - earned)
                    dfac = rel
                    denom = r2((net or 0) * rel)
                    cfac = None if denom == 0 else round(payable_base / denom, 6)
            this_period = r2((earned or 0) - prior) if earned is not None else None
            if full is not None:
                order_deal_comm = r2(order_deal_comm + full)

            rep_lines.append(dict(rep_id=rid, rep_name=name, share=share, commission_type=ctype,
                                  effective_rate=eff, full_commission=full, delivery_factor=dfac,
                                  collection_factor=cfac, delivery_holdback=deliv_hold,
                                  collection_holdback=coll_hold, earned_to_date=earned,
                                  prior_processed=prior, this_period=this_period))
            if this_period is not None:
                acc = rep_summary.setdefault(rid, {"name": name, "this_period": 0.0, "full": 0.0, "lines": []})
                acc["this_period"] = r2(acc["this_period"] + this_period)
                acc["full"] = r2(acc["full"] + (full or earned or 0))
                acc["lines"].append(dict(order_number=onum, customer=o.get("customer"),
                                         commission_type=ctype, this_period=this_period,
                                         full_commission=(full if full is not None else earned)))

        order_results.append(dict(
            order_number=onum, customer=o.get("customer"), configuration=o.get("deal_type"),
            finance_type=finance, release_rule=o.get("release_rule"),
            net_commissionable=net, contract_price=R["contract"], cleared=R["cleared"],
            cleared_capped=cleared_capped, payable_base=payable_base,
            collection_factor=coll_disp, release_fraction=rel,
            reps=rep_lines, deal_commission=(order_deal_comm if order_deal_comm else None),
        ))

    return order_results, rep_summary, exceptions


def control_totals(order_results):
    return {
        "orders": len(order_results),
        "sum_deal_commission": r2(sum((o.get("deal_commission") or 0) for o in order_results)),
        "sum_period_commission": r2(sum((rl["this_period"] or 0) for o in order_results for rl in o["reps"])),
    }
