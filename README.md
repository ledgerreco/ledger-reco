# 📒 Ledger Reconciliation · लेजर मिलान

A simple web app to reconcile **our books** (vendor/customer ledger from Tally, Zoho Books, QuickBooks or Excel)
with the **party's ledger / statement of account**.

- Upload two files: **xlsx, xls, csv or text-based PDF** (scanned PDFs are not supported yet).
- Auto-detects the header row (Tally title rows are skipped) and the columns: Date, Voucher/Invoice No.,
  Particulars, Debit, Credit (or one Amount column with +/- or Dr/Cr). You can change any of it in the app.
- Handles mirror ledgers (our Debit = party's Credit); switch it off in the sidebar if both files are from the same books.
- Matching passes:
  1. reference no. + amount (references cleaned: case, spaces, prefixes like INV/BILL/SI, fiscal-year parts)
  2. same reference but different amount → **Amount Mismatch**
  3. amount + date within N days (default 7)
  4. amount only (closest date)

  Amount tolerance default ₹1. Every match is one entry to one entry.
- On-screen summary + reconciliation statement, and an **Excel report** with sheets:
  Summary, Matched, Only in Ledger A, Only in Ledger B, Amount Mismatch.
- 🔒 Files are processed in memory on the server and are not stored.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Tests
```bash
pip install -r requirements-dev.txt
python -m pytest
python run_samples.py   # end-to-end run on samples/, writes samples/output_reconciliation*.xlsx
```

## Deploy free on Streamlit Community Cloud
1. Push this repo to a public GitHub repository.
2. Sign in at https://share.streamlit.io with GitHub → **Create app** → choose the repo, branch `main`, file `app.py`.
3. Pick a sub-domain → the app is live at `https://<name>.streamlit.app`.

## Files
| File | Purpose |
|---|---|
| `app.py` | Streamlit UI (entry point) |
| `reco.py` | Core logic: reading Excel/CSV/PDF, header & column detection, matching, Excel export |
| `.streamlit/config.toml` | 50 MB upload limit, headless, light theme |
| `tests/test_reco.py` | pytest suite |
| `run_samples.py` | End-to-end run on the sample files |
| `samples/` | Tally-style ledger + party statement (xlsx, csv, table PDF, plain-text PDF, scanned PDF), `make_samples.py` |

## Limitations
- Scanned/image PDFs need OCR (not supported yet).
- One-to-one matching only: one payment against several bills (or part payments) shows as unmatched.
- Dates are read day-first (Indian format).
- Opening balance is picked up only from an "Opening Balance" / "b/f" row; closing balance is computed.
