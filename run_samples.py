"""End-to-end run on the sample files; writes samples/output_reconciliation*.xlsx"""
from pathlib import Path
import reco

S = Path(__file__).parent / "samples"
a_file = S / "our_books_tally_sharma_traders.xlsx"
for b_name, out in [("party_statement_sharma_traders.xlsx", "output_reconciliation.xlsx"),
                    ("party_statement_sharma_traders.csv", None),
                    ("party_statement_sharma_traders.pdf", "output_reconciliation_pdf.xlsx"),
                    ("party_statement_sharma_traders_plain.pdf", None)]:
    a, mpa, ha, *_ = reco.ledger_from_file(a_file.read_bytes(), a_file.name)
    b, mpb, hb, _, info = reco.ledger_from_file((S / b_name).read_bytes(), b_name)
    res = reco.reconcile(a, b, flip_b=True, date_tol=7, amt_tol=1.0, label_a="Our Books", label_b="Sharma Traders")
    s = res.summary
    print(f"\n=== Our Books (Tally xlsx) vs {b_name}  [source={info['source']}, header rows A={ha} B={hb}]")
    print("  mapping B:", {k: v for k, v in mpb.items() if v})
    print(f"  rows A={s['rows_a']} B={s['rows_b']} | matched={s['matched_count']} {s['match_by_type']}")
    print(f"  only A={s['only_a_count']} ({reco.inr(s['only_a_amount'])}) | only B={s['only_b_count']} ({reco.inr(s['only_b_amount'])}) | mismatch={s['mismatch_count']} ({reco.inr(s['mismatch_amount'])})")
    print(f"  closing A={reco.drcr(s['closing_a'])}  closing B={reco.drcr(s['closing_b_own'])}  diff={reco.inr(s['difference'])}")
    print("  unexplained:", res.statement.iloc[-1]["Amount"], "| warnings:", a.warnings + b.warnings)
    if out:
        (S / out).write_bytes(reco.to_excel(res))
        print("  wrote", S / out)
try:
    reco.load_file((S / "scanned_example.pdf").read_bytes(), "scanned_example.pdf")
except reco.ScannedPDFError as e:
    print("\nscanned_example.pdf ->", e)
