"""Ledger Reconciliation — Streamlit app (free-hostable on Hugging Face Spaces / Streamlit Cloud)."""
import pandas as pd
import streamlit as st

import reco

st.set_page_config(page_title="Ledger Reconciliation | लेजर मिलान", page_icon="📒", layout="wide")

st.title("📒 Ledger Reconciliation · लेजर मिलान")
st.caption("Upload two ledgers (Excel / CSV / PDF) → get matched & unmatched entries + Excel report. "
           "दो लेजर अपलोड करें, मिलान की रिपोर्ट पाएँ।")
st.info("🔒 Your files are processed in memory on the server and are **not stored**. "
        "फ़ाइलें सर्वर पर सिर्फ़ मेमोरी में प्रोसेस होती हैं, कहीं सेव नहीं होतीं।", icon="🔒")

# ----------------------------- Sidebar settings -----------------------------
with st.sidebar:
    st.header("⚙️ Settings · सेटिंग")
    label_a = st.text_input("Name of side A (साइड A का नाम)", "Our Books")
    label_b = st.text_input("Name of side B (साइड B का नाम)", "Party Ledger")
    flip = st.checkbox("Mirror ledgers: our Debit = their Credit (उल्टी एंट्री)", value=True,
                       help="Keep ON when B is the party's own ledger/statement. Turn OFF if both files "
                            "are from the same books (e.g. two exports of our own ledger).")
    date_tol = st.number_input("Date tolerance, days (तारीख का अंतर, दिन)", 0, 365, 7)
    amt_tol = st.number_input("Amount tolerance ₹ (राशि का अंतर)", 0.0, 100000.0, 1.0, step=0.5)
    st.markdown("---")
    st.caption("Matching passes: ① Ref no. + amount → ② amount + date (± days) → ③ amount only "
               "(closest date). Same ref but different amount → *Amount Mismatch*.")

FIELD_LABELS = {
    "date": "Date (तारीख)", "ref": "Voucher / Invoice No. (बिल नं.)", "narration": "Particulars (विवरण)",
    "debit": "Debit (नामे)", "credit": "Credit (जमा)", "amount": "Single Amount column (+/- or Dr/Cr)",
    "drcr": "Dr/Cr indicator column",
}


def ledger_panel(key: str, label: str):
    """Upload + header/column mapping UI for one side. Returns a reco.Ledger or None."""
    st.subheader(f"{'🅰️' if key == 'a' else '🅱️'} {label}")
    up = st.file_uploader(f"Upload {label} (xlsx / xls / csv / pdf)", type=["xlsx", "xls", "xlsm", "csv", "pdf"],
                          key=f"file_{key}")
    if not up:
        return None
    data = up.getvalue()
    sheet = 0
    try:
        sheets = reco.list_sheets(data, up.name)
        if len(sheets) > 1:
            sheet = st.selectbox("Sheet (शीट)", sheets, key=f"sheet_{key}")
        raw, info = reco.load_file(data, up.name, sheet=sheet)
    except reco.ScannedPDFError as e:
        st.error(f"🖼️ {e}\n\nस्कैन की हुई PDF अभी सपोर्ट नहीं है — कृपया Excel/CSV या टेक्स्ट PDF भेजें।")
        return None
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read this file / फ़ाइल नहीं पढ़ पाए: {e}")
        return None
    for w in info["warnings"]:
        st.warning(w)
    if info["source"] and info["source"].startswith("pdf"):
        st.caption(f"PDF read via {'table extraction' if info['source'] == 'pdf-table' else 'text-line parsing'}"
                   " — please verify the preview below. (नीचे प्रीव्यू जाँच लें)")

    auto_hdr = reco.detect_header_row(raw)
    with st.expander("Raw file preview & header row (हेडर लाइन)", expanded=False):
        st.dataframe(raw.head(30).astype(str), width="stretch", height=260)
    hdr = st.number_input("Header row number (0 = first row) · हेडर किस लाइन पर है", 0, max(len(raw) - 1, 0),
                          int(auto_hdr), key=f"hdr_{key}",
                          help="Auto-detected. Tally exports have company name / period rows on top.")
    df = reco.apply_header(raw, int(hdr))
    auto = reco.detect_columns(df)
    cols = ["— none —"] + list(df.columns)

    with st.expander("Column mapping (कॉलम चुनें) — auto-detected, change if wrong", expanded=False):
        mapping = {}
        c1, c2 = st.columns(2)
        for n, f in enumerate(reco.FIELDS):
            default = auto.get(f)
            idx = cols.index(default) if default in cols else 0
            with (c1 if n % 2 == 0 else c2):
                choice = st.selectbox(FIELD_LABELS[f], cols, index=idx, key=f"map_{key}_{f}")
            mapping[f] = None if choice == "— none —" else choice
        st.caption("Use Debit + Credit columns, OR a single Amount column (optionally with a Dr/Cr column).")

    try:
        led = reco.standardize(df, mapping)
    except Exception as e:  # noqa: BLE001
        st.error(f"Mapping problem / कॉलम मैपिंग जाँचें: {e}")
        return None
    for w in led.warnings:
        st.warning(w)
    m1, m2, m3 = st.columns(3)
    m1.metric("Entries (एंट्री)", len(led.txns))
    m2.metric("Opening (शुरुआती)", reco.drcr(led.opening) if led.opening_found else "not found")
    m3.metric("Closing (computed)", reco.drcr(led.closing))
    show = led.txns[["row", "date", "ref", "narration", "debit", "credit"]].rename(
        columns={"row": "Line", "date": "Date", "ref": "Ref", "narration": "Particulars",
                 "debit": "Debit", "credit": "Credit"})
    st.caption("Extracted entries preview (पढ़ी गई एंट्री):")
    st.dataframe(show.head(200), width="stretch", height=240, hide_index=True)
    return led


ca, cb = st.columns(2, gap="large")
with ca:
    led_a = ledger_panel("a", label_a)
with cb:
    led_b = ledger_panel("b", label_b)

st.markdown("---")
if not (led_a and led_b):
    st.write("⬆️ Upload both ledgers to start · दोनों लेजर अपलोड करें।")
    st.stop()
if len(led_a.txns) == 0 or len(led_b.txns) == 0:
    st.error("One of the ledgers has no entries — check header row / column mapping.")
    st.stop()

res = reco.reconcile(led_a, led_b, flip_b=flip, date_tol=int(date_tol), amt_tol=float(amt_tol),
                     label_a=label_a, label_b=label_b)
s = res.summary

st.header("📊 Result · नतीजा")
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("✅ Matched (मिलान)", s["matched_count"], f"₹ {reco.inr(s['matched_amount'], 0)}", delta_color="off")
k2.metric(f"Only in {label_a}", s["only_a_count"], f"₹ {reco.inr(s['only_a_amount'], 0)}", delta_color="off")
k3.metric(f"Only in {label_b}", s["only_b_count"], f"₹ {reco.inr(s['only_b_amount'], 0)}", delta_color="off")
k4.metric("⚠️ Amount mismatch", s["mismatch_count"], f"₹ {reco.inr(s['mismatch_amount'], 0)}", delta_color="off")
k5.metric("Balance difference (अंतर)", f"₹ {reco.inr(s['difference'])}")

xlsx = reco.to_excel(res)
st.download_button("⬇️ Download Excel report (एक्सेल रिपोर्ट डाउनलोड करें)", xlsx,
                   file_name=f"reconciliation_{label_a}_vs_{label_b}.xlsx".replace(" ", "_"),
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

l, r = st.columns([1, 1], gap="large")
with l:
    st.subheader("Summary (सारांश)")
    st.dataframe(reco.summary_table(res).astype(str), width="stretch", hide_index=True)
with r:
    st.subheader("Reconciliation statement (मिलान विवरण)")
    stmt = res.statement.copy()
    stmt["Amount"] = stmt["Amount"].map(lambda v: "" if v is None or pd.isna(v) else reco.inr(v))
    st.dataframe(stmt, width="stretch", hide_index=True)
    st.caption("All amounts in our view: Dr = + , Cr = − ." + (" Party balance is sign-flipped." if flip else ""))


def fmt(df):
    if df.empty:
        return df
    d = df.copy()
    for c in d.columns:
        if "Date" in c:
            d[c] = pd.to_datetime(d[c]).dt.strftime("%d-%m-%Y")
        elif any(k in c for k in ("Debit", "Credit", "Difference", "Amount", "Net")):
            d[c] = d[c].map(lambda v: "" if pd.isna(v) or v == 0 else reco.inr(v))
    return d


t1, t2, t3, t4 = st.tabs([f"✅ Matched ({len(res.matched)})", f"Only in {label_a} ({len(res.only_a)})",
                          f"Only in {label_b} ({len(res.only_b)})", f"⚠️ Amount Mismatch ({len(res.mismatch)})"])
with t1:
    st.dataframe(fmt(res.matched), width="stretch", hide_index=True)
with t2:
    st.caption(f"In {label_a} but not found in {label_b} · सिर्फ़ {label_a} में")
    st.dataframe(fmt(res.only_a), width="stretch", hide_index=True)
with t3:
    st.caption(f"In {label_b} but not found in {label_a} · सिर्फ़ {label_b} में")
    st.dataframe(fmt(res.only_b), width="stretch", hide_index=True)
with t4:
    st.caption("Same invoice/reference number but different amount · बिल नंबर एक, राशि अलग")
    st.dataframe(fmt(res.mismatch), width="stretch", hide_index=True)
