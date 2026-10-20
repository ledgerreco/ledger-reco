from pathlib import Path
import io
import openpyxl
import pytest
import reco

S = Path(__file__).resolve().parent.parent / "samples"
A_FILE = S / "our_books_tally_sharma_traders.xlsx"


def _load(name):
    led, mp, hdr, df, info = reco.ledger_from_file((S / name).read_bytes(), name)
    return led, mp, hdr, info


def test_tally_header_and_columns():
    led, mp, hdr, _ = _load(A_FILE.name)
    assert hdr == 5  # 6 title/blank rows on top (blank row dropped)
    assert mp["date"] == "Date" and mp["debit"] == "Debit" and mp["credit"] == "Credit"
    assert mp["ref"] == "Vch No." and mp["narration"] == "Particulars"
    assert led.opening == -50000 and led.opening_found
    assert len(led.txns) == 14
    assert round(led.closing, 2) == -117720.0 == led.stated_closing


@pytest.mark.parametrize("party", ["party_statement_sharma_traders.xlsx", "party_statement_sharma_traders.csv",
                                   "party_statement_sharma_traders.pdf", "party_statement_sharma_traders_plain.pdf"])
def test_end_to_end(party):
    a, *_ = _load(A_FILE.name)
    b, _, _, info = _load(party)
    res = reco.reconcile(a, b, flip_b=True, date_tol=7, amt_tol=1.0)
    s = res.summary
    assert s["rows_b"] == 14
    assert s["matched_count"] == 11
    assert s["match_by_type"] == {"1. Reference + Amount": 6, "2. Amount + Date (±7d)": 4, "3. Amount only": 1}
    assert s["only_a_count"] == 2 and s["only_b_count"] == 2 and s["mismatch_count"] == 1
    assert round(s["difference"], 2) == 25849.60
    assert abs(res.statement.iloc[-1]["Amount"]) < 0.01  # fully explained
    wb = openpyxl.load_workbook(io.BytesIO(reco.to_excel(res)))
    assert wb.sheetnames == ["Summary", "Matched", "Only in Ledger A", "Only in Ledger B", "Amount Mismatch"]


def test_tighter_date_tolerance_moves_to_amount_only():
    a, *_ = _load(A_FILE.name)
    b, *_ = _load("party_statement_sharma_traders.xlsx")
    s = reco.reconcile(a, b, date_tol=1).summary
    assert s["matched_count"] == 11 and s["match_by_type"]["3. Amount only"] == 3


def test_no_flip_matches_nothing_by_amount():
    a, *_ = _load(A_FILE.name)
    b, *_ = _load("party_statement_sharma_traders.xlsx")
    assert reco.reconcile(a, b, flip_b=False).summary["matched_count"] == 0


def test_scanned_pdf_rejected():
    with pytest.raises(reco.ScannedPDFError):
        reco.load_file((S / "scanned_example.pdf").read_bytes(), "scanned_example.pdf")


def test_helpers():
    assert reco.parse_amount("1,23,456.50 Dr") == 123456.5
    assert reco.parse_amount("500.00 Cr") == -500
    assert reco.parse_amount("(250)") == -250
    assert reco.parse_amount("₹ 1,000") == 1000
    assert reco.norm_ref("Inv No. 00123") == "00123"
    assert reco.ref_number("INV/2024-25/0105") == "105"
    assert reco.ref_number("SI-105") == "105"
    assert reco.inr(12345678.9) == "1,23,45,678.90"
    assert reco.inr(-999) == "-999.00"
