# Boston Aesthetics — Commission Calculator

A team-shared Streamlit app that turns the finalized order workbooks + delivery
register into per-rep commission packets and an HR workbook. Every payroll run is
saved to a shared database, so your team can open the same in-progress run from any
browser and pick up where others left off — until it's approved and locked.

## What it does

1. **Inputs** — upload the delivery register, the setup workbook (rep master), and
   the four order workbooks (individually, or one ZIP that also holds the LPS PDFs).
2. **Review & Calculate** — set each order's finance type, commission rate, rep
   split, and payments; it recomputes everything from raw inputs, applies the bundle
   **50/50** delivery release and the **cleared-cash** collection rule, subtracts any
   prior-processed amounts, and lists exceptions.
3. **Approve & Lock** — move the run Draft → In review → Approved → Locked. Locking
   freezes it so finalized commissions can't change or be double-counted.
4. **Reports** — download the **HR workbook** (Payroll Summary, Commission Detail,
   Exceptions, Reconciliation, Plan Snapshot) and a **per-rep PDF packet** (cover
   statement + commission detail + each order's LPS).

The commission math mirrors the order sheet exactly: commission base = pre-tax
subtotal (devices + discount + customer shipping); **tax is excluded**; net = subtotal
− deductions. The matrix **rate is keyed to the net commissionable amount**.

What pays out this period, per rep:

- **Standard / Financed** — pays on the cash actually cleared, at the rate, capped by
  the delivered half:
  `this period = rate × min(cleared cash capped at net, net × delivery fraction) × share − prior`.
  Two holdbacks are shown on the statement: a **delivery holdback** (the undelivered
  device's half, released when the PICO ships) and a **collection holdback** (delivered
  but not yet cleared).
- **In-House Financing** — 6% of each cleared payment (down payment + monthlies); no
  delivery holdback (these are not pre-orders).
- **Quarter Accelerator** — per rep, ≥ 6 straight-purchase devices in the quarter →
  16%, retroactive within the run.
- **Super Kicker** — per rep, ≥ $1,000,000 straight-purchase volume in the quarter →
  22%, retroactive within the run.

Rounding is half-up to the cent.

## File naming it expects

- Order workbooks: `BA-2608-001_Commission_Detail.xlsx` (any name works; the order
  number is read from inside the file).
- LPS PDFs: **`BA-2608-001_LPS.pdf`** — the app matches an LPS to its order by the
  order number in the filename.

## Run it locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Without a database secret it uses a local `commission.db` (SQLite) — fine for one
person, but not shared.

## Deploy for the team (Streamlit Community Cloud + free Postgres)

1. **Create a free Postgres** at [neon.tech](https://neon.tech) (or Supabase). Copy
   its connection string (looks like
   `postgresql://user:pass@host/db?sslmode=require`).
2. **Put this folder in a GitHub repo.**
3. **Deploy** at [share.streamlit.io](https://share.streamlit.io) → New app → point
   at the repo and `app.py`.
4. In the app's **Settings → Secrets**, paste:
   ```toml
   DATABASE_URL = "postgresql://user:pass@host/db?sslmode=require"
   ```
5. Share the app URL with your team. Everyone sees the same runs; the tables are
   created automatically on first launch.

To keep it private to your team, set the app to private in Streamlit Cloud and invite
members by email (or host it internally).

## Notes

- The app never trusts Excel's cached formula results — it recomputes from the raw
  line items, deductions, rate, and shares, so the numbers are always live.
- A locked run is the record of what was sent. To fix something afterward, use
  **Duplicate to new run** and correct the copy, preserving the original.
- Files: `app.py` (UI) · `ingest.py` (read workbooks) · `calc.py` (engine) ·
  `reports.py` (HR workbook + PDFs) · `db.py` (shared storage).
