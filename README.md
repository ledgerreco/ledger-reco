# 📒 Ledger Reconciliation

A simple web app to reconcile **our books** (vendor/customer ledger from Tally, Zoho Books, QuickBooks or Excel)
with the **party's ledger / statement of account**.

- Upload two files: **xlsx, xls, csv, PDF, or scanned PDF / photo (jpg, png)**. Scanned files are read by Grok AI.
- Auto-detects the header row (Tally title rows are skipped) and the columns: Date, Voucher/Invoice No.,
  Particulars, Debit, Credit (or one Amount column with +/- or Dr/Cr). You can change any of it in the app.
- Handles mirror ledgers (our Debit = party's Credit); switch it off in the sidebar if both files are from the same books.
- Matching passes:
  1. reference no. + amount (references cleaned: case, spaces, prefixes like INV/BILL/SI, fiscal-year parts)
  2. same reference but different amount → **Amount Mismatch**
  3. amount + date within N days (default 7)
  4. amount only (closest date)

  Amount tolerance default ₹1. Every match is one entry to one entry.
- Nothing runs until you click **✅ Reconcile karo / Submit**. Results stay in the session, so downloading the report
  doesn't run anything again.
- **🤖 Grok AI (optional, xAI API):** after the rule passes, Grok reviews only the leftover entries and proposes
  grouped matches (one payment vs several bills, an invoice split into two entries, TDS/GST/discount/rounding
  differences), flags duplicates and missing entries, and writes a short English + Hindi summary with action items.
  The app checks every AI proposal before accepting it: the ids must exist, each entry can be used only once,
  both sides must be present, confidence must be at least 0.6, and the amounts must agree within tolerance.
  A difference is allowed only for TDS/GST/discount/rounding, and at most 10 %. Rejected proposals are listed
  in *AI Notes*. If AI fails or times out, the rule results are still shown.
  Large leftovers are sent in chunks of at most 300 rows.
- On-screen summary + reconciliation statement, and an **Excel report** with sheets:
  Summary (incl. AI summary), Matched, Only in Ledger A, Only in Ledger B, Amount Mismatch, AI Matched, AI Notes.
- 🔒 Files are processed in memory on the server and are not stored. When AI is on, the unmatched entries
  (and pages of scanned files) are sent to the xAI API.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Grok AI key
Locally: `export XAI_API_KEY=...` or create `.streamlit/secrets.toml` (git-ignored) with `XAI_API_KEY = "..."`.
On Streamlit Community Cloud: App → Settings → Secrets → `XAI_API_KEY = "..."`. Without a key the AI features
are switched off and everything else works.

## Tests
```bash
pip install -r requirements-dev.txt
python -m pytest
python run_samples.py      # rule-based end-to-end run on samples/
python run_ai_samples.py   # live Grok AI run (needs XAI_API_KEY, a few US cents)
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
| `ai.py` | Grok AI: leftover matching + validation, summary, vision extraction of scanned files |
| `.streamlit/config.toml` | 50 MB upload limit, headless, light theme |
| `tests/` | pytest suite (AI tests use a fake client, no network) |
| `run_samples.py` | End-to-end run on the sample files |
| `samples/` | Tally-style ledger + party statement (xlsx, csv, table PDF, plain-text PDF), AI samples (`ai_*_v2.xlsx`: one payment vs 2 receipts, split invoice), scanned statement (`scanned_party_statement.pdf/.jpg`), `make_samples.py` |

## Limitations
- Without AI: scanned files can't be read, and only one-to-one matching is done.
- AI results can vary slightly between runs. Every accepted AI match is checked by the app, but please review
  the AI Matched sheet. AI review takes about 30–90 s.
- Dates are read day-first (Indian format).
- Opening balance is picked up only from an "Opening Balance" / "b/f" row; closing balance is computed.
