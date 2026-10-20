"""Ledger Reconciliation — Streamlit app (server version) with optional Grok AI."""
import hashlib
import os

import pandas as pd
import streamlit as st

import ai
import reco

st.set_page_config(page_title="Ledger Reconciliation | लेजर मिलान", page_icon="📒", layout="wide")
ss = st.session_state
ss.setdefault("extract_cache", {})   # file hash -> (raw grid, info) from Grok vision
ss.setdefault("result", None)        # last reconciliation (kept so downloads don't re-run anything)
ss.setdefault("run_requested", False)

try:
    API_KEY = ai.get_api_key(st.secrets)
except Exception:  # noqa: BLE001
    API_KEY = ai.get_api_key(None)

st.title("📒 Ledger Reconciliation · लेजर मिलान")
st.caption("Upload two ledgers (Excel / CSV / PDF / scanned image) → click **Reconcile karo** → get matched & "
           "unmatched entries + Excel report. दो लेजर अपलोड करें, फिर 'Reconcile karo' दबाएँ।")
st.info("🔒 Your files are processed in memory on the server and are **not stored**. "
        "फ़ाइलें सर्वर पर सिर्फ़ मेमोरी में प्रोसेस होती हैं, कहीं सेव नहीं होतीं।", icon="🔒")

# ----------------------------- Sidebar settings -----------------------------
with st.sidebar:
    st.header("⚙️ Settings · सेटिंग")
    label_a = st.text_input("Name of side A (साइड A का नाम)", "Our Books")
    label_b = st.text_input("Name of side B (साइड B का नाम)", "Party Ledger")
    flip = st.checkbox("Mirror ledgers: our Debit = their Credit (उल्टी एंट्री)", value=True,
                       help="Keep ON when B is the party's own ledger/statement. Turn OFF if both files "
                            "are from the same books.")
    date_tol = st.number_input("Date tolerance, days (तारीख का अंतर, दिन)", 0, 365, 7)
    amt_tol = st.number_input("Amount tolerance ₹ (राशि का अंतर)", 0.0, 100000.0, 1.0, step=0.5)
    st.markdown("---")
    st.subheader("🤖 Grok AI")
    if API_KEY:
        use_ai = st.toggle("Use Grok AI (AI से बेहतर मिलान)", value=True,
                           help="After the rule matching, Grok reviews the leftover entries: one payment vs several "
                                "bills, split entries, TDS/GST/discount differences, duplicates; writes a summary "
                                "and action items; reads scanned PDFs / photos.")
        if use_ai:
            st.caption("When AI is on, the unmatched entries (and scanned pages) are sent to the xAI API for "
                       "analysis. AI matches are checked by the app before they are accepted.")
    else:
        use_ai = False
        st.toggle("Use Grok AI", value=False, disabled=True)
        st.caption("AI is off: no XAI_API_KEY configured on the server (add it in Streamlit secrets).")
    st.markdown("---")
    st.caption("Rule passes: ① Ref no. + amount → ② amount + date (± days) → ③ amount only (closest date). "
               "Same ref but different amount → *Amount Mismatch*. Then (optional) 🤖 Grok AI on the leftovers.")

FIELD_LABELS = {
    "date": "Date (तारीख)", "ref": "Voucher / Invoice No. (बिल नं.)", "narration": "Particulars (विवरण)",
    "debit": "Debit (नामे)", "credit": "Credit (जमा)", "amount": "Single Amount column (+/- or Dr/Cr)",
    "drcr": "Dr/Cr indicator column",
}


def ledger_panel(key: str, label: str):
    """Upload + header/column mapping UI for one side.
    Returns None, {'status': 'pending', ...} (needs Grok vision on submit) or {'status': 'ready', 'ledger': ..., 'sig': ...}."""
    st.subheader(f"{'🅰️' if key == 'a' else '🅱️'} {label}")
    up = st.file_uploader(f"Upload {label} (xlsx / xls / csv / pdf / jpg / png)",
                          type=["xlsx", "xls", "xlsm", "csv", "pdf", "jpg", "jpeg", "png"], key=f"file_{key}")
    if not up:
        return None
    data = up.getvalue()
    h = hashlib.sha1(data).hexdigest()
    name = up.name
    sheet, raw, info, needs_vision, err = 0, None, {"warnings": []}, False, None
    try:
        sheets = reco.list_sheets(data, name)
        if len(sheets) > 1:
            sheet = st.selectbox("Sheet (शीट)", sheets, key=f"sheet_{key}")
        raw, info = reco.load_file(data, name, sheet=sheet)
    except reco.NeedsVisionError as e:
        needs_vision, err = True, str(e)
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read this file / फ़ाइल नहीं पढ़ पाए: {e}")
        return None

    want_vision = needs_vision
    if not needs_vision and name.lower().endswith(".pdf") and use_ai:
        try:
            df0 = reco.apply_header(raw, reco.detect_header_row(raw))
            poor = ai.text_parse_looks_poor(reco.standardize(df0, reco.detect_columns(df0)), df0)
        except Exception:  # noqa: BLE001
            poor = True
        want_vision = st.checkbox("🤖 Read this PDF with Grok AI instead (better for odd layouts)", value=poor,
                                  key=f"vision_{key}",
                                  help="Auto-ticked when normal PDF reading found too few rows.")
    if want_vision:
        if not use_ai:
            st.error(f"🖼️ {err}\n\nस्कैन/फ़ोटो पढ़ने के लिए Grok AI चालू करें, या Excel/CSV भेजें।")
            return None
        cache_key = f"{h}:{sheet}"
        if cache_key not in ss.extract_cache:
            st.info("🖼️ Scanned / image file — Grok AI will read the rows when you click **Reconcile karo**. "
                    "स्कैन फ़ाइल 'Reconcile karo' दबाने पर AI से पढ़ी जाएगी।")
            return {"status": "pending", "data": data, "name": name, "cache_key": cache_key}
        raw, info = ss.extract_cache[cache_key]
        st.caption(f"🤖 Read by Grok AI from {info.get('pages', 1)} page(s) — check the rows below against the "
                   "original. (AI से पढ़ी गई एंट्री, कृपया मिलान कर लें)")
        if not info.get("rows"):
            st.error("Grok AI could not find ledger rows in this file. / इस फ़ाइल में एंट्री नहीं मिलीं।")
            return None
    for w in info.get("warnings", []):
        st.warning(w)
    if str(info.get("source", "")).startswith("pdf"):
        st.caption(f"PDF read via {'table extraction' if info['source'] == 'pdf-table' else 'text-line parsing'}"
                   " — please verify the preview below. (नीचे प्रीव्यू जाँच लें)")

    auto_hdr = reco.detect_header_row(raw)
    with st.expander("Raw file preview & header row (हेडर लाइन)", expanded=False):
        st.dataframe(raw.head(30).astype(str), width="stretch", height=260)
    hdr = st.number_input("Header row number (0 = first row) · हेडर किस लाइन पर है", 0, max(len(raw) - 1, 0),
                          int(auto_hdr), key=f"hdr_{key}_{h[:8]}",
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
                choice = st.selectbox(FIELD_LABELS[f], cols, index=idx, key=f"map_{key}_{f}_{h[:8]}_{hdr}")
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
    show["Date"] = pd.to_datetime(show["Date"]).dt.strftime("%d-%m-%Y")
    st.caption("Extracted entries preview (पढ़ी गई एंट्री):")
    st.dataframe(show.head(300), width="stretch", height=240, hide_index=True)
    return {"status": "ready", "ledger": led, "ai_usage": info.get("usage", ""), "sig": (h, sheet, int(hdr), tuple(sorted(mapping.items(), key=str)),
                                                       bool(want_vision))}


ca, cb = st.columns(2, gap="large")
with ca:
    side_a = ledger_panel("a", label_a)
with cb:
    side_b = ledger_panel("b", label_b)

st.markdown("---")
both = bool(side_a and side_b)
settings_sig = (label_a, label_b, flip, int(date_tol), float(amt_tol), bool(use_ai))
cur_sig = (side_a.get("sig") if side_a else None, side_b.get("sig") if side_b else None, settings_sig)

clicked = st.button("✅ Reconcile karo / Submit", type="primary", disabled=not both, width="stretch")
if not both:
    st.caption("⬆️ Upload both ledgers, check the previews, then click **Reconcile karo / Submit**. "
               "दोनों लेजर अपलोड करें, फिर बटन दबाएँ।")
if clicked:
    ss.run_requested = True

if ss.run_requested and both:
    # Step 1: scanned/image files -> Grok vision (only now, after the click). Then rerun so previews/mapping update.
    pending = [s for s in (side_a, side_b) if s["status"] == "pending"]
    if pending:
        client = ai.GrokClient(API_KEY, timeout=240)
        for p in pending:
            with st.spinner(f"🤖 Grok AI is reading {p['name']} … (AI फ़ाइल पढ़ रहा है)"):
                try:
                    imgs = ai.render_pages(p["data"], p["name"])
                    grid, info = ai.ai_extract_rows(imgs, client)
                    info["usage"] = client.usage.text()
                    ss.extract_cache[p["cache_key"]] = (grid, info)
                except Exception as e:  # noqa: BLE001
                    ss.run_requested = False
                    st.error(f"Grok AI could not read {p['name']}: {type(e).__name__}: {str(e)[:200]}")
                    st.stop()
        st.rerun()

    # Step 2: rules (+ AI) — runs once per click; result kept in session_state
    led_a, led_b = side_a["ledger"], side_b["ledger"]
    ss.run_requested = False
    if len(led_a.txns) == 0 or len(led_b.txns) == 0:
        st.error("One of the ledgers has no entries — check header row / column mapping.")
        st.stop()
    res = reco.reconcile(led_a, led_b, flip_b=flip, date_tol=int(date_tol), amt_tol=float(amt_tol),
                         label_a=label_a, label_b=label_b)
    usage_txt = ""
    if use_ai:
        client = ai.GrokClient(API_KEY, timeout=180)
        with st.spinner("🤖 Grok AI is reviewing unmatched entries… (30–90 s) · AI बची हुई एंट्री जाँच रहा है"):
            res = ai.run_ai(res, client)
        usage_txt = "review: " + client.usage.text()
    for sd, lab in ((side_a, label_a), (side_b, label_b)):
        if sd.get("ai_usage"):
            usage_txt += f" · reading {lab}: {sd['ai_usage']}"
    ss.result = {"res": res, "xlsx": reco.to_excel(res), "sig": cur_sig, "usage": usage_txt}
elif ss.run_requested and not both:
    ss.run_requested = False

R = ss.result
if not R:
    st.stop()
if R["sig"] != cur_sig:
    st.warning("⚠️ Files or settings changed since the last run — click **Reconcile karo / Submit** again to update. "
               "बदलाव के बाद दोबारा बटन दबाएँ।")
res = R["res"]
s = res.summary
la, lb = res.label_a, res.label_b

st.header("📊 Result · नतीजा")
k = st.columns(6 if res.ai_groups else 5)
k[0].metric("✅ Matched (मिलान)", s["matched_count"], f"₹ {reco.inr(s['matched_amount'], 0)}", delta_color="off")
k[1].metric(f"Only in {la}", s["only_a_count"], f"₹ {reco.inr(s['only_a_amount'], 0)}", delta_color="off")
k[2].metric(f"Only in {lb}", s["only_b_count"], f"₹ {reco.inr(s['only_b_amount'], 0)}", delta_color="off")
k[3].metric("⚠️ Amount mismatch", s["mismatch_count"], f"₹ {reco.inr(s['mismatch_amount'], 0)}", delta_color="off")
if res.ai_groups:
    k[4].metric("🤖 AI matched groups", s["ai_group_count"], f"₹ {reco.inr(s['ai_amount'], 0)}", delta_color="off")
k[-1].metric("Balance difference (अंतर)", f"₹ {reco.inr(s['difference'])}")

if res.ai_status.startswith("error"):
    st.warning(f"🤖 Grok AI step failed — showing rule-based results only. ({res.ai_status})")
elif res.ai_status:
    st.caption(f"🤖 Grok AI: {res.ai_status}" + (f" · {R['usage']}" if R["usage"] else ""))

st.download_button("⬇️ Download Excel report (एक्सेल रिपोर्ट डाउनलोड करें)", R["xlsx"],
                   file_name=f"reconciliation_{la}_vs_{lb}.xlsx".replace(" ", "_"),
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

if res.ai_summary:
    with st.container(border=True):
        st.subheader("🤖 AI summary · AI सारांश")
        st.write(res.ai_summary.get("summary_en", ""))
        st.write(res.ai_summary.get("summary_hi", ""))
        items = res.ai_summary.get("action_items", [])
        if items:
            st.markdown("**Action items · क्या करना है:**")
            icon = {"high": "🔴", "medium": "🟠", "low": "🟢"}
            for it in items:
                st.markdown(f"- {icon.get(it.get('priority'), '•')} {it.get('action', '')}")

l, r = st.columns([1, 1], gap="large")
with l:
    st.subheader("Summary (सारांश)")
    st.dataframe(reco.summary_table(res).astype(str), width="stretch", hide_index=True)
with r:
    st.subheader("Reconciliation statement (मिलान विवरण)")
    stmt = res.statement.copy()
    stmt["Amount"] = stmt["Amount"].map(lambda v: "" if v is None or pd.isna(v) else reco.inr(v))
    st.dataframe(stmt, width="stretch", hide_index=True)
    st.caption("All amounts in our view: Dr = + , Cr = − ." + (" Party balance is sign-flipped." if res.flip_b else ""))


def fmt(df):
    if df.empty:
        return df
    d = df.copy()
    for c in d.columns:
        if "Date" in c:
            d[c] = pd.to_datetime(d[c]).dt.strftime("%d-%m-%Y")
        elif "Reason" not in c and any(x in c for x in ("Debit", "Credit", "Difference", "Amount", "Net", "Total")):
            d[c] = d[c].map(lambda v: "" if pd.isna(v) or v == 0 else reco.inr(v))
    return d


tabs = [f"✅ Matched ({len(res.matched)})", f"Only in {la} ({len(res.only_a)})", f"Only in {lb} ({len(res.only_b)})",
        f"⚠️ Amount Mismatch ({len(res.mismatch)})"]
if res.ai_status:
    tabs += [f"🤖 AI Matched ({len(res.ai_matched)})", f"🤖 AI Notes ({len(res.ai_notes)})"]
t = st.tabs(tabs)
with t[0]:
    st.dataframe(fmt(res.matched), width="stretch", hide_index=True)
with t[1]:
    st.caption(f"In {la} but not found in {lb} · सिर्फ़ {la} में")
    st.dataframe(fmt(res.only_a), width="stretch", hide_index=True)
with t[2]:
    st.caption(f"In {lb} but not found in {la} · सिर्फ़ {lb} में")
    st.dataframe(fmt(res.only_b), width="stretch", hide_index=True)
with t[3]:
    st.caption("Same invoice/reference number but different amount · बिल नंबर एक, राशि अलग")
    st.dataframe(fmt(res.mismatch), width="stretch", hide_index=True)
if res.ai_status:
    with t[4]:
        st.caption("Groups proposed by Grok AI and verified by the app (ids exist, used once, amounts add up). "
                   "AI द्वारा मिलाई गई और ऐप द्वारा जाँची गई एंट्री।")
        st.dataframe(fmt(res.ai_matched), width="stretch", hide_index=True)
    with t[5]:
        st.caption("Observations and suggested actions from Grok AI (incl. proposals the app rejected).")
        st.dataframe(res.ai_notes, width="stretch", hide_index=True)
