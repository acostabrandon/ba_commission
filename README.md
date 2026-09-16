# Boston Aesthetics — Commission Calculator

A team-shared Streamlit app that turns the finalized order workbooks + LPS PDFs into
per-rep commission packets and an HR workbook. Every payroll run is saved to a shared
database, so the team can open the same in-progress run from any browser and pick up
where others left off — until it's approved and locked.

## What it does

1. **Inputs** — upload the setup workbook (rep master + opening history) and the order
   workbooks for **this payroll period only** (individually, or one ZIP that also holds
   the LPS PDFs named `BA-2608-00X_LPS.pdf`). No delivery register — delivery status is
   set per order on the Review tab.
2. **Review & Calculate** — set each order's finance type, commission rate, rep split,
   **delivery status** (0% / 50% / 100% — the Boston Pico backorder holdback), and
   payments. It recomputes from raw inputs and lists exceptions.
3. **Approve & Lock** — move the run Draft → In review → Approved → Locked. Locking
   freezes it so finalized commissions can't change or be double-counted.
4. **Reports** — download the **HR workbook** and the **per-rep PDF packets**: a branded
   cover statement, then for each order the **exact commission-detail sheet** (rendered
   to a single page, STATUS/Template Version excluded) followed by that order's LPS.

## Commission model

Commission base = pre-tax subtotal (devices + discount + customer shipping); **tax is
excluded**; net = subtotal − deductions. The matrix rate is keyed to the net
commissionable amount. What pays out this period, per rep:

- **Standard / Financed** — pays on cash actually cleared **within the period**, at the
  rate, capped by the delivered half:
  `rate × min(cleared cash capped at net, net × delivery fraction) × share − prior`.
  The statement shows a **Delivery Holdback** (undelivered Pico half, releases when it
  ships) and a **Collection Holdback** (delivered but not yet cleared).
- **In-House Financing** — 6% of each cleared payment; no delivery holdback.
- **Quarter Accelerator** — per rep, ≥ 6 straight-purchase devices in the quarter → 16%.
- **Super Kicker** — per rep, ≥ $1,000,000 straight-purchase volume in the quarter → 22%.

Only payments dated inside the run's period and marked cleared count. Rounding is
half-up to the cent. The app never trusts Excel's cached results — it recomputes from
the raw line items, deductions, rate, shares, and payments.

## File naming it expects

- Order workbooks: `BA-2608-001_Commission_Detail.xlsx` (any name; the order number is
  read from inside the file).
- LPS PDFs: `BA-2608-001_LPS.pdf` — matched to an order by the order number in the name.

## Run it locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Without a database secret it uses a local `commission.db` (SQLite). The exact one-page
detail render needs LibreOffice installed locally (`soffice` on PATH); without it, the
app falls back to a generated one-page detail.

## Deploy for the team (Streamlit Community Cloud + free Postgres)

1. **Create a free Postgres** at [neon.tech](https://neon.tech) (or Supabase). Copy the
   connection string (`postgresql://user:pass@host/db?sslmode=require`).
2. **Push this folder to GitHub.**
3. **Deploy** at [share.streamlit.io](https://share.streamlit.io) → New app → point at
   the repo and `app.py`.
4. In **Settings → Secrets**, paste:
   ```toml
   DATABASE_URL = "postgresql://user:pass@host/db?sslmode=require"
   ```
5. Share the app URL. Everyone sees the same runs; tables are created on first launch.

`packages.txt` installs LibreOffice so the app can render each commission-detail sheet
to a one-page PDF exactly as it appears in Excel.

## Files

`app.py` (UI) · `ingest.py` (read workbooks) · `calc.py` (engine) · `reports.py`
(HR workbook + PDF packets, one-page detail render) · `db.py` (shared storage) ·
`money.py` (half-up rounding) · `ba_logo.png` (statement letterhead) ·
`requirements.txt` · `packages.txt` (LibreOffice) · `.streamlit/config.toml` (theme).
