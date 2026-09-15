"""Boston Aesthetics — Commission Calculator (Streamlit).

Team-shared: every payroll run is saved to the database, so anyone on the team can
open the same in-progress run and pick up where others left off, until it's approved
and locked. See README.md for one-time Streamlit Cloud + Postgres setup.
"""
import io, zipfile, datetime
import streamlit as st
import db, ingest, calc, reports

st.set_page_config(page_title="BA Commission Calculator", page_icon="🧮", layout="wide")
USER = st.session_state.setdefault("user", "")

st.title("Boston Aesthetics — Commission Calculator")

# ---------------- sidebar: identity, DB, run selection ----------------
with st.sidebar:
    st.header("Session")
    st.session_state["user"] = st.text_input("Your name (for the audit log)", value=USER)
    USER = st.session_state["user"]
    try:
        runs = db.list_runs()
        st.caption(f"Database connected · {len(runs)} run(s)")
    except Exception as e:
        st.error("Database not reachable. Check DATABASE_URL secret.")
        st.exception(e); st.stop()

    st.divider()
    st.subheader("Payroll run")
    labels = {r["id"]: f"{r['name']} [{r['status']}]" for r in runs}
    options = ["➕ New run…"] + list(labels.keys())
    pick = st.selectbox("Open a run", options, format_func=lambda x: "➕ New run…" if x == "➕ New run…" else labels[x])
    if pick == "➕ New run…":
        with st.form("newrun"):
            name = st.text_input("Run name", value="Aug 16 – Sep 15, 2026")
            c1, c2 = st.columns(2)
            ps = c1.text_input("Period start", value="2026-08-16")
            pe = c2.text_input("Period end", value="2026-09-15")
            if st.form_submit_button("Create run"):
                rid = db.create_run(name, ps, pe, USER)
                st.session_state["run_id"] = rid
                st.rerun()
    else:
        st.session_state["run_id"] = pick

run_id = st.session_state.get("run_id")
if not run_id:
    st.info("Create or open a payroll run from the sidebar to begin.")
    st.stop()

run = db.load_run(run_id)
if not run:
    st.warning("That run no longer exists."); st.session_state.pop("run_id", None); st.stop()
payload = run["payload"]
locked = run["status"] == "locked"

top = st.container()
top.subheader(f"{run['name']}  ·  {run['period_start']} → {run['period_end']}")
badge = {"draft": "🟡 Draft", "in_review": "🔵 In review", "approved": "🟢 Approved", "locked": "🔒 Locked"}
top.write(f"Status: **{badge.get(run['status'], run['status'])}**  ·  last saved by "
          f"*{run['updated_by'] or '—'}* at {run['updated_at']:%Y-%m-%d %H:%M} UTC")
if locked:
    top.warning("This run is locked. Inputs are read-only so finalized commissions can't change. "
                "Duplicate a run to make corrections.")

tabs = st.tabs(["1 · Inputs", "2 · Review & Calculate", "3 · Approve & Lock", "4 · Reports"])


def save(msg=None, status=None):
    if msg:
        payload.setdefault("log", []).append(
            {"ts": datetime.datetime.utcnow().isoformat(timespec="seconds"), "user": USER, "note": msg})
    db.save_run(run_id, payload, status=status, user=USER)


# ============================ 1 · INPUTS ============================
with tabs[0]:
    st.markdown("Upload the setup workbook (rep master + opening history) and the order workbooks "
                "for **this payroll period only** (individually, or a single ZIP that also contains the "
                "LPS PDFs named `BA-2608-00X_LPS.pdf`). No delivery register needed — you set each "
                "bundle's delivery status on the Review tab.")
    disabled = locked
    setup = st.file_uploader("Setup workbook (.xlsx)", type=["xlsx"], disabled=disabled)
    if setup is not None:
        data = setup.getvalue()
        payload["reps_master"] = ingest.parse_setup_reps(data)
        payload["opening_history"] = ingest.parse_opening_history(data)
        save("Uploaded setup workbook")
        st.success(f"Reps: {len(payload['reps_master'])} · opening-history rows: {len(payload.get('opening_history', []))}")

    st.markdown("**Order workbooks / ZIP**")
    ups = st.file_uploader("Order workbooks (.xlsx) or a .zip with workbooks + LPS PDFs",
                           type=["xlsx", "zip"], accept_multiple_files=True, disabled=disabled)
    if ups:
        orders = {o["order_number"]: o for o in payload.get("orders", [])}
        n_lps = 0
        errors = []

        def _skip(name):
            base = name.split("/")[-1]
            # skip directories and macOS/Office junk that isn't a real file
            return (name.endswith("/") or "__MACOSX" in name or base.startswith("._")
                    or base.startswith("~$") or base.startswith("."))

        def _add_order(b, base):
            try:
                o = ingest.parse_order_workbook(b, base)
            except Exception as ex:
                errors.append(f"{base}: not a readable .xlsx ({ex})"); return
            if o.get("order_number"):
                orders[o["order_number"]] = o
            else:
                errors.append(f"{base}: no Order Number found (is this a commission-detail workbook?)")

        for up in ups:
            if up.name.lower().endswith(".zip"):
                try:
                    zf = zipfile.ZipFile(io.BytesIO(up.getvalue()))
                except Exception as ex:
                    errors.append(f"{up.name}: not a valid ZIP ({ex})"); continue
                for nm in zf.namelist():
                    if _skip(nm):
                        continue
                    base = nm.split("/")[-1]
                    low = base.lower()
                    if low.endswith(".xlsx"):
                        _add_order(zf.read(nm), base)
                    elif low.endswith(".pdf"):
                        db.save_attachment(run_id, "lps", base, zf.read(nm)); n_lps += 1
            elif up.name.lower().endswith(".xlsx"):
                _add_order(up.getvalue(), up.name)
            elif up.name.lower().endswith(".pdf"):
                db.save_attachment(run_id, "lps", up.name, up.getvalue()); n_lps += 1
        payload["orders"] = list(orders.values())
        save(f"Ingested {len(orders)} order(s), {n_lps} LPS PDF(s)")
        st.success(f"{len(orders)} order(s) loaded" + (f", {n_lps} LPS PDF(s) stored" if n_lps else ""))
        if errors:
            st.warning("Some files were skipped:\n\n- " + "\n- ".join(errors))

    if payload.get("orders"):
        st.dataframe([{"Order": o["order_number"], "Customer": o["customer"], "Deal": o["deal_type"],
                       "Finance": o.get("finance_type"),
                       "Net": o["net_commissionable"], "Rate": o["commission_rate"],
                       "Reps": ", ".join(f"{(r.get('rep_name') or r['rep_id'])} ({r['rep_id']})"
                                         for r in o.get("reps", []) if r.get("rep_id"))}
                      for o in payload["orders"]], use_container_width=True, hide_index=True)
    lps_files = db.list_attachments(run_id, "lps")
    if lps_files:
        st.caption("LPS PDFs stored: " + ", ".join(a["name"] for a in lps_files))

# ============================ 2 · REVIEW ============================
with tabs[1]:
    orders = payload.get("orders", [])
    if not orders:
        st.info("Upload inputs on the Inputs tab first.")
    else:
        st.markdown("Set each order's **finance type, commission rate, rep split, and payments**. "
                    "Changes here override the sheet without editing it. Then recalculate.")
        overrides = payload.setdefault("overrides", {})

        # rep-name lookup: setup workbook first, else the name carried on the sheet
        rmaster0 = payload.get("reps_master", {})
        name_by_rid = {}
        for o in orders:
            for r in o.get("reps", []):
                if r.get("rep_id"):
                    name_by_rid[str(r["rep_id"])] = r.get("rep_name") or str(r["rep_id"])
        for rid, info in rmaster0.items():
            if info.get("name"):
                name_by_rid[str(rid)] = info["name"]

        def rep_name(rid):
            return name_by_rid.get(str(rid), str(rid))

        # ---- statement details (fill in what the workbooks don't carry) ----
        sf = payload.setdefault("statement_fields", {})
        with st.expander("Statement details — payroll date, territory, manager, quarterly revenue, comments", expanded=False):
            sf["payroll_date"] = st.text_input("Payroll date (applies to all statements)",
                                               value=sf.get("payroll_date", ""), key="payroll_date", disabled=locked)
            rmaster = payload.get("reps_master", {})
            rep_fields = sf.setdefault("reps", {})
            appearing = sorted({str(r["rep_id"]) for o in orders for r in o.get("reps", []) if r.get("rep_id")})
            for rid in appearing:
                rf = rep_fields.setdefault(rid, {})
                nm = rep_name(rid)
                st.markdown(f"**{nm} ({rid})**")
                c1, c2, c3 = st.columns(3)
                rf["territory"] = c1.text_input("Territory", value=rf.get("territory", ""), key=f"terr_{rid}", disabled=locked)
                rf["manager"] = c2.text_input("Manager", value=rf.get("manager", ""), key=f"mgr_{rid}", disabled=locked)
                rf["quarterly_revenue"] = c3.text_input("Quarterly revenue", value=rf.get("quarterly_revenue", ""), key=f"qr_{rid}", disabled=locked)
                rf["comments"] = st.text_area("Comments", value=rf.get("comments", ""), key=f"cmt_{rid}", height=68, disabled=locked)
            if st.button("💾 Save statement details", disabled=locked):
                save("Updated statement details"); st.rerun()

        for o in orders:
            onum = str(o["order_number"])
            ov = overrides.setdefault(onum, {})
            with st.expander(f"{onum} — {o['customer']}  ·  net ${o['net_commissionable']:,.2f}", expanded=True):
                default_rate = ov.get("commission_rate", o.get("commission_rate")) or 0.0
                rate_pct = st.number_input(f"Commission rate % — {onum}", min_value=0.0, max_value=100.0,
                                           value=float(default_rate) * 100, step=0.5, key=f"rate_{onum}",
                                           disabled=locked)
                ov["commission_rate"] = round(rate_pct / 100.0, 6)
                reps = o.get("reps", [])
                rep_ids = [str(r["rep_id"]) for r in reps if r.get("rep_id")]
                sh_over = ov.setdefault("reps", {})
                if rep_ids:
                    cols = st.columns(len(rep_ids))
                    for i, rid in enumerate(rep_ids):
                        nm = rep_name(rid)
                        d = sh_over.get(rid)
                        if d is None:
                            src = next((r["share"] for r in reps if str(r["rep_id"]) == rid), None)
                            d = (src if src is not None else (1.0 if len(rep_ids) == 1 else 0.0))
                        val = cols[i].number_input(f"{nm} ({rid}) share %", min_value=0.0, max_value=100.0,
                                                   value=float(d) * 100, step=1.0, key=f"sh_{onum}_{rid}",
                                                   disabled=locked)
                        sh_over[rid] = round(val / 100.0, 6)
                    tot = sum(sh_over.values())
                    (st.success if abs(tot - 1.0) < 1e-6 else st.warning)(f"Shares total {tot*100:.0f}%")

                st.markdown("**Finance, delivery & payments**")
                ft_opts = ["Straight Purchase", "Financed Purchase", "In-House Financed Purchase"]
                cur_ft = ov.get("finance_type") or o.get("finance_type") or "Straight Purchase"
                if cur_ft not in ft_opts:
                    cur_ft = "Straight Purchase"
                fc1, fc2, fc3 = st.columns(3)
                ov["finance_type"] = fc1.selectbox("Finance type", ft_opts, index=ft_opts.index(cur_ft),
                                                   key=f"ft_{onum}", disabled=locked)

                # delivery release — replaces the old delivery register
                bundle = calc.is_bundle(o)
                if bundle:
                    dlv_opts = [("Neither delivered (0%)", 0.0),
                                ("First device delivered — PICO backordered (50%)", 0.5),
                                ("Both delivered (100%)", 1.0)]
                else:
                    dlv_opts = [("Not delivered (0%)", 0.0), ("Delivered (100%)", 1.0)]
                cur_rel = ov.get("release_fraction")
                if cur_rel in (None, ""):
                    cur_rel = calc.default_release(o)
                idx = next((i for i, (_, v) in enumerate(dlv_opts) if abs(v - float(cur_rel)) < 1e-6), 0)
                pick = fc2.selectbox("Delivery status", [lbl for lbl, _ in dlv_opts], index=idx,
                                     key=f"dlv_{onum}", disabled=locked)
                ov["release_fraction"] = dict(dlv_opts)[pick]

                default_contract = ov.get("contract_price") or o.get("contract_price") or o.get("invoice_total") or o.get("net_commissionable") or 0.0
                ov["contract_price"] = round(fc3.number_input("Contract price (total the customer pays)", min_value=0.0,
                                             value=float(default_contract), step=100.0, key=f"cp_{onum}", disabled=locked), 2)
                st.caption("Payments — each collected amount with its date. Only payments dated within this run's "
                           "period and marked 'Cleared' count toward commission. Standard/Financed deals pay on the "
                           "cash cleared (capped at net), gated by the delivered half; In-House pays 6% of each cleared "
                           "payment. Check 'Cleared' once funds hit the bank.")
                seed = ov.get("payments") or o.get("payments") or [{"date": "", "amount": 0.0, "kind": "Down Payment", "cleared": True}]
                edited = st.data_editor(
                    seed, num_rows="dynamic", key=f"pay_{onum}", disabled=locked, hide_index=True,
                    column_config={
                        "date": st.column_config.TextColumn("Date"),
                        "amount": st.column_config.NumberColumn("Amount", format="$%.2f"),
                        "kind": st.column_config.SelectboxColumn("Kind", options=["Down Payment", "Installment", "Payoff", "Paid in Full", "Other"]),
                        "cleared": st.column_config.CheckboxColumn("Cleared?"),
                    })
                ov["payments"] = [dict(date=r.get("date", ""), amount=float(r.get("amount") or 0),
                                       kind=r.get("kind", "Other"), cleared=bool(r.get("cleared")))
                                  for r in edited if (r.get("amount") or 0) or r.get("date")]

        cc1, cc2 = st.columns([1, 3])
        if cc1.button("💾 Save & recalculate", disabled=locked, type="primary"):
            res, summ, exc = calc.calculate(orders, payload.get("reps_master", {}),
                                            payload.get("opening_history", []), overrides=overrides,
                                            period=(run["period_start"], run["period_end"]))
            payload["results"] = {"orders": res, "rep_summary": summ, "exceptions": exc,
                                  "control": calc.control_totals(res)}
            save("Recalculated")
            st.rerun()

        results = payload.get("results")
        if results:
            st.divider()
            ct = results["control"]
            m1, m2, m3 = st.columns(3)
            m1.metric("Full-deal commission", f"${ct['sum_deal_commission']:,.2f}")
            m2.metric("Releasing this period", f"${ct['sum_period_commission']:,.2f}")
            m3.metric("Open exceptions", len(results["exceptions"]))
            if results["exceptions"]:
                st.warning("Exceptions to resolve:\n\n- " + "\n- ".join(results["exceptions"]))
            st.markdown("**Per rep — this period**")
            st.dataframe([{"Rep ID": rid, "Rep": s["name"], "This period": round(s["this_period"], 2),
                           "Full-deal": round(s["full"], 2)} for rid, s in sorted(results["rep_summary"].items())],
                         use_container_width=True, hide_index=True)
            st.markdown("**Per order**")
            st.dataframe([{"Order": o["order_number"], "Customer": o["customer"],
                           "Finance": o.get("finance_type"),
                           "Type": ", ".join(sorted({rl["commission_type"] for rl in o["reps"]})) or "—",
                           "Net": o["net_commissionable"],
                           "Collected %": (None if o.get("collection_factor") is None else round(o["collection_factor"] * 100)),
                           "Release": o["release_fraction"],
                           "Deal comm.": o.get("deal_commission"),
                           "This period": round(sum((rl["this_period"] or 0) for rl in o["reps"]), 2)}
                          for o in results["orders"]],
                         use_container_width=True, hide_index=True)

# ============================ 3 · APPROVE & LOCK ============================
with tabs[2]:
    st.markdown("Move the run through review. Locking prevents further edits so finalized commissions "
                "can't change or be double-counted.")
    res = payload.get("results")
    exc = (res or {}).get("exceptions", [])
    if exc:
        st.warning(f"{len(exc)} open exception(s). Resolve on the Review tab before approving.")
    c1, c2, c3, c4 = st.columns(4)
    if c1.button("Mark In Review", disabled=locked):
        save("Marked in review", status="in_review"); st.rerun()
    if c2.button("Approve", disabled=locked or bool(exc)):
        save("Approved", status="approved"); st.rerun()
    if c3.button("🔒 Lock run", disabled=locked or run["status"] != "approved"):
        save("Locked", status="locked"); st.rerun()
    if c4.button("Duplicate to new run"):
        rid = db.create_run(run["name"] + " (copy)", run["period_start"], run["period_end"], USER)
        cp = dict(payload); db.save_run(rid, cp, status="draft", user=USER)
        st.session_state["run_id"] = rid; st.rerun()
    st.divider()
    st.markdown("**Activity log**")
    for e in reversed(payload.get("log", [])[-30:]):
        st.caption(f"{e['ts']} · {e.get('user') or '—'} · {e['note']}")

# ============================ 4 · REPORTS ============================
with tabs[3]:
    res = payload.get("results")
    if not res:
        st.info("Recalculate on the Review tab first.")
    else:
        st.markdown("Download the HR workbook and the per-rep PDF packets (cover statement + commission "
                    "detail + each order's LPS).")
        run_full = db.load_run(run_id)
        colA, colB = st.columns(2)
        hr = reports.build_hr_workbook(run_full)
        colA.download_button("⬇️ HR workbook (.xlsx)", hr,
                             file_name=f"HR_Commission_{run['period_end']}.xlsx",
                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        # per-rep packets zipped
        lps_by_order = {}
        for o in res["orders"]:
            onum = o["order_number"]
            blob = db.get_attachment(run_id, f"{onum}_LPS.pdf", "lps")
            if blob is None:  # fall back to any stored LPS whose name contains the order number
                for a in db.list_attachments(run_id, "lps"):
                    if str(onum) in a["name"]:
                        blob = db.get_attachment(run_id, a["name"], "lps"); break
            if blob:
                lps_by_order[onum] = blob
        missing = [o["order_number"] for o in res["orders"] if o["order_number"] not in lps_by_order]
        if missing:
            st.caption("No LPS PDF stored for: " + ", ".join(map(str, missing)) + " (packets will omit those LPS pages).")
        if st.button("Build per-rep packets", type="primary"):
            zbuf = io.BytesIO()
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                for rid, s in res["rep_summary"].items():
                    pdf = reports.build_rep_packet(rid, run_full, lps_by_order)
                    safe = (s["name"] or rid).replace(" ", "_")
                    zf.writestr(f"{rid}_{safe}_Commission_Packet.pdf", pdf)
            st.session_state["packets_zip"] = zbuf.getvalue()
            st.success(f"Built {len(res['rep_summary'])} packet(s).")
        if st.session_state.get("packets_zip"):
            colB.download_button("⬇️ All rep packets (.zip)", st.session_state["packets_zip"],
                                 file_name=f"Rep_Packets_{run['period_end']}.zip", mime="application/zip")
