"""Live end-to-end check of the Grok AI features on the samples (needs XAI_API_KEY; costs a few cents).
Usage: python run_ai_samples.py"""
import json
from pathlib import Path

import pandas as pd

import ai
import reco

S = Path(__file__).parent / "samples"
pd.set_option("display.width", 220)
pd.set_option("display.max_colwidth", 90)
key = ai.get_api_key()
assert key, "XAI_API_KEY not set"


def load(name):
    return reco.ledger_from_file((S / name).read_bytes(), name)[0]


# ---------- 1. AI matching on v2 (payment vs 2 receipts, split invoice) ----------
client = ai.GrokClient(key)
a, b = load("ai_our_books_tally_v2.xlsx"), load("ai_party_statement_v2.xlsx")
rule = reco.reconcile(a, b, label_a="Our Books", label_b="Sharma Traders")
res = ai.run_ai(rule, client)
s = res.summary
print("== AI match on v2 ==  status:", res.ai_status)
print(f"rule matched={s['matched_count']} | AI groups={s['ai_group_count']} | only A={s['only_a_count']} "
      f"only B={s['only_b_count']} mismatch={s['mismatch_count']} | unexplained={res.statement.iloc[-1]['Amount']}")
print(res.ai_matched[["Type", "A Entries", "B Entries", "Difference", "Status", "Confidence", "Reason"]].to_string())
print("\nAI notes:\n", res.ai_notes.to_string())
print("\nAI summary:\n", json.dumps(res.ai_summary, ensure_ascii=False, indent=1))
print("usage (match+summary):", client.usage.text())
(S / "output_ai_reconciliation_v2.xlsx").write_bytes(reco.to_excel(res))

# ---------- 2. Vision extraction of the scanned statement ----------
expected = load("party_statement_sharma_traders.xlsx").txns
for fname in ("scanned_party_statement.pdf", "scanned_party_statement.jpg"):
    vc = ai.GrokClient(key)
    imgs = ai.render_pages((S / fname).read_bytes(), fname)
    grid, info = ai.ai_extract_rows(imgs, vc)
    hdr = reco.detect_header_row(grid)
    df = reco.apply_header(grid, hdr)
    led = reco.standardize(df, reco.detect_columns(df))
    t = led.txns
    exp_keys = set(zip(expected["date"], expected["debit"].round(2), expected["credit"].round(2)))
    got_keys = set(zip(t["date"], t["debit"].round(2), t["credit"].round(2)))
    refs_ok = sum(1 for r, e in zip(t.sort_values("date")["ref"], expected.sort_values("date")["ref"]) if r == e)
    print(f"\n== Vision extract {fname}: pages={info['pages']} rows={len(t)} (expected {len(expected)}) | "
          f"rows exactly right (date+debit+credit)={len(exp_keys & got_keys)}/{len(expected)} | refs equal={refs_ok} | "
          f"opening={reco.drcr(led.opening)} (expected 50,000.00 Dr)")
    print("usage:", vc.usage.text())
    r2 = ai.run_ai(reco.reconcile(load("our_books_tally_sharma_traders.xlsx"), led, label_a="Our Books",
                                  label_b="Sharma (scanned)"), ai.GrokClient(key), with_summary=False)
    s2 = r2.summary
    print(f"   reconcile vs our books: rule matched={s2['matched_count']} {s2['match_by_type']} AI groups={s2['ai_group_count']} "
          f"only A={s2['only_a_count']} only B={s2['only_b_count']} mismatch={s2['mismatch_count']} diff={s2['difference']:,.2f}")
    if fname.endswith(".pdf"):
        (S / "output_ai_scanned.xlsx").write_bytes(reco.to_excel(r2))
