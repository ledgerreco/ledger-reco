"""Ledger Reconciliation — Streamlit app (server version) with AI review, AI summary and AI chat."""
import hashlib

import pandas as pd
import streamlit as st

import ai
import reco

st.set_page_config(page_title="Ledger Reconciliation | लेजर मिलान", page_icon="📒", layout="wide")
ss = st.session_state
ss.setdefault("extract_cache", {})   # file hash -> (raw grid, info) read by AI vision
ss.setdefault("result", None)        # last reconciliation (kept so downloads/chat don't re-run anything)
ss.setdefault("run_requested", False)
ss.setdefault("chat", [])            # [{"role": "user"/"assistant", "content": str}]

try:
    API_KEY = ai.get_api_key(st.secrets)
except Exception:  # noqa: BLE001
    API_KEY = ai.get_api_key(None)

FIELD_LABELS = {
    "date": "Date (तारीख)", "ref": "Voucher / Invoice No. (बिल नं.)", "narration": "Particulars (विवरण)",
    "debit": "Debit (नामे)", "credit": "Credit (जमा)", "amount": "Single Amount column (+/- or Dr/Cr)",
    "drcr": "Dr/Cr indicator column",
}

st.title("📒 Ledger Reconciliation · लेजर मिलान")
st.caption("① Upload two ledgers (Excel / CSV / PDF / scanned photo) → ② optional AI instruction → "
           "③ settings → ④ **Reconcile karo / Submit** → results + AI chat.")
st.info("🔒 Files are processed in memory on the server and are **not stored**. Unmatched entries (and scanned "
        "pages) are sent to an AI service for review. फ़ाइलें कहीं सेव नहीं होतीं; बची हुई एंट्री AI सेवा को "
        "जाँच के लिए भेजी जाती हैं।", icon="🔒")
if not API_KEY:
    st.error("⚠️ AI is not configured: the admin must add **XAI_API_KEY** in the app's Secrets. "
             "Rule-based reconciliation still works. (AI चालू करने के लिए एडमिन को key जोड़नी होगी)")


# =============================== 1. Uploads ===============================
def upload_part(key: str, title: str):
    up = st.file_uploader(title, type=["xlsx", "xls", "xlsm", "csv", "pdf", "jpg", "jpeg", "png"], key=f"file_{key}")
    if not up:
        return None
    data = up.getvalue()
    f = {"key": key, "name": up.name, "data": data, "h": hashlib.sha1(data).hexdigest(), "sheets": [],
         "raw": None, "info": {"warnings": []}, "needs_vision": False, "poor": False, "error": None}
    try:
        f["sheets"] = reco.list_sheets(data, up.name)
        sheet = ss.get(f"sheet_{key}_{f['h'][:8]}", f["sheets"][0] if f["sheets"] else 0)
        f["raw"], f["info"] = reco.load_file(data, up.name, sheet=sheet)
        if up.name.lower().endswith(".pdf"):
            df0 = reco.apply_header(f["raw"], reco.detect_header_row(f["raw"]))
            f["poor"] = ai.text_parse_looks_poor(reco.standardize(df0, reco.detect_columns(df0)), df0)
    except reco.NeedsVisionError as e:
        f["needs_vision"], f["error"] = True, str(e)
    except Exception as e:  # noqa: BLE001
        f["error"] = f"Could not read this file / फ़ाइल नहीं पढ़ पाए: {e}"
        st.error(f["error"])
        return None
    if f["needs_vision"]:
        st.caption("🖼️ Scanned / image file — AI will read the rows when you click **Reconcile karo**.")
    for w in f["info"].get("warnings", []):
        st.caption(f"ℹ️ {w}")
    return f


u1, u2 = st.columns(2, gap="large")
with u1:
    st.subheader("🅰️ Ledger A")
    file_a = upload_part("a", "Upload side A — e.g. our books (xlsx / xls / csv / pdf / jpg / png)")
with u2:
    st.subheader("🅱️ Ledger B")
    file_b = upload_part("b", "Upload side B — e.g. party ledger / statement")


# =============================== 2–4. One form: instruction + settings + submit ===============================
def side_settings(f, key, label):
    """Settings widgets for one side (inside the form). Returns dict describing how to build the ledger."""
    st.markdown(f"**{label}** — `{f['name']}`")
    h8 = f["h"][:8]
    sheet = 0
    if len(f["sheets"]) > 1:
        sheet = st.selectbox("Sheet (शीट)", f["sheets"], key=f"sheet_{key}_{h8}")
    use_vision = f["needs_vision"]
    if not f["needs_vision"] and f["name"].lower().endswith(".pdf"):
        use_vision = st.checkbox("🤖 Read this PDF with AI (better for odd layouts / poor text)", value=f["poor"],
                                 key=f"vision_{key}_{h8}", help="Auto-ticked when normal PDF reading found too few rows.")
    elif f["needs_vision"]:
        st.caption("🤖 Will be read by AI (scanned / image).")
    cache_key = f"{f['h']}:{sheet}"
    raw = None
    if use_vision and cache_key in ss.extract_cache:
        raw = ss.extract_cache[cache_key][0]
    elif not use_vision:
        raw = f["raw"]
    out = {"file": f, "label": label, "sheet": sheet, "use_vision": use_vision, "cache_key": cache_key,
           "hdr": None, "mapping": None}
    if raw is None:
        st.caption("Column mapping will appear after the AI has read the file (after Submit).")
        return out
    src = "ai" if use_vision else "file"
    auto_hdr = reco.detect_header_row(raw)
    hdr = st.number_input("Header row number (0 = first row) · हेडर किस लाइन पर है", 0, max(len(raw) - 1, 0),
                          int(auto_hdr), key=f"hdr_{key}_{h8}_{src}",
                          help="Auto-detected. Tally exports have company name / period rows on top.")
    df = reco.apply_header(raw, int(hdr))
    auto = reco.detect_columns(df)
    cols = ["— none —"] + list(df.columns)
    mapping = {}
    c1, c2 = st.columns(2)
    for n, fld in enumerate(reco.FIELDS):
        default = auto.get(fld)
        idx = cols.index(default) if default in cols else 0
        with (c1 if n % 2 == 0 else c2):
            choice = st.selectbox(FIELD_LABELS[fld], cols, index=idx, key=f"map_{key}_{fld}_{h8}_{src}_{hdr}")
        mapping[fld] = None if choice == "— none —" else choice
    out.update(hdr=int(hdr), mapping=mapping, raw=raw)
    return out


both = bool(file_a and file_b)
with st.form("reco_form", border=True):
    st.subheader("💬 AI ko instruction do (optional)")
    instruction = st.text_area(
        "AI ko instruction do (optional)", key="instruction", height=80, label_visibility="collapsed",
        placeholder="e.g. 'TDS 2% ignore karo', 'sirf March ka data dekho', 'GST difference ko alag dikhao'")
    with st.expander("⚙️ Settings · सेटिंग (labels, matching, column mapping)", expanded=False):
        g1, g2 = st.columns(2)
        label_a = g1.text_input("Name of side A (साइड A का नाम)", "Our Books", key="label_a")
        label_b = g2.text_input("Name of side B (साइड B का नाम)", "Party Ledger", key="label_b")
        g3, g4, g5 = st.columns(3)
        flip = g3.checkbox("Mirror ledgers: our Debit = their Credit (उल्टी एंट्री)", value=True, key="flip",
                           help="Keep ON when B is the party's own ledger/statement. Turn OFF if both files are "
                                "from the same books.")
        date_tol = g4.number_input("Date tolerance, days (तारीख का अंतर)", 0, 365, 7, key="date_tol")
        amt_tol = g5.number_input("Amount tolerance ₹ (राशि का अंतर)", 0.0, 100000.0, 1.0, step=0.5, key="amt_tol")
        st.caption("Rule passes: ① Ref no. + amount → ② amount + date (± days) → ③ amount only (closest date). "
                   "Same ref but different amount → Amount Mismatch. Then 🤖 AI reviews the leftovers.")
        cfg_a = cfg_b = None
        if both:
            st.markdown("---")
            s1, s2 = st.columns(2, gap="large")
            with s1:
                cfg_a = side_settings(file_a, "a", label_a)
            with s2:
                cfg_b = side_settings(file_b, "b", label_b)
        else:
            st.caption("Column mapping options appear here after both files are uploaded.")
    submitted = st.form_submit_button("✅ Reconcile karo / Submit", type="primary", width="stretch")

if submitted:
    if not both:
        st.warning("⬆️ Please upload both ledgers first. दोनों लेजर अपलोड करें।")
    else:
        ss.run_requested = True


def build_ledger(cfg):
    if cfg["use_vision"] and cfg["cache_key"] not in ss.extract_cache:
        return None, "pending"
    if cfg["mapping"] is None:
        return None, "pending"
    df = reco.apply_header(cfg["raw"], cfg["hdr"])
    led = reco.standardize(df, cfg["mapping"])
    if cfg["use_vision"]:
        led.warnings = ss.extract_cache[cfg["cache_key"]][1].get("warnings", []) + led.warnings
    return led, "ready"


if ss.run_requested and both:
    # Step 1: files that need AI reading (scanned / image / ticked PDFs) -> read now, then rerun
    pending = [c for c in (cfg_a, cfg_b) if c and c["use_vision"] and c["cache_key"] not in ss.extract_cache]
    if pending:
        if not API_KEY:
            ss.run_requested = False
            st.error("This file needs AI to be read, but XAI_API_KEY is not configured. Please upload Excel/CSV.")
            st.stop()
        client = ai.GrokClient(API_KEY, timeout=240)
        for c in pending:
            with st.spinner(f"🤖 AI is reading {c['file']['name']} … (AI फ़ाइल पढ़ रहा है)"):
                try:
                    imgs = ai.render_pages(c["file"]["data"], c["file"]["name"])
                    grid, info = ai.ai_extract_rows(imgs, client)
                    info["usage"] = client.usage.text()
                    ss.extract_cache[c["cache_key"]] = (grid, info)
                except Exception as e:  # noqa: BLE001
                    ss.run_requested = False
                    st.error(f"AI could not read {c['file']['name']}: {type(e).__name__}: {str(e)[:200]}")
                    st.stop()
        st.rerun()

    # Step 2: rules + AI — once per submit
    ss.run_requested = False
    try:
        (led_a, _), (led_b, _) = build_ledger(cfg_a), build_ledger(cfg_b)
    except Exception as e:  # noqa: BLE001
        st.error(f"Mapping problem / कॉलम मैपिंग जाँचें (Settings): {e}")
        st.stop()
    if led_a is None or led_b is None or len(led_a.txns) == 0 or len(led_b.txns) == 0:
        st.error("One of the ledgers has no entries — check header row / column mapping in ⚙️ Settings.")
        st.stop()
    res = reco.reconcile(led_a, led_b, flip_b=flip, date_tol=int(date_tol), amt_tol=float(amt_tol),
                         label_a=label_a, label_b=label_b)
    usage_txt = ""
    if API_KEY:
        client = ai.GrokClient(API_KEY, timeout=180)
        with st.spinner("🤖 AI is reviewing unmatched entries… (30–90 s) · AI बची हुई एंट्री जाँच रहा है"):
            res = ai.run_ai(res, client, instruction=instruction)
        usage_txt = "review: " + client.usage.text()
    else:
        res.ai_status = "error: AI not configured (XAI_API_KEY missing)"
    for c in (cfg_a, cfg_b):
        u = ss.extract_cache.get(c["cache_key"], (None, {}))[1].get("usage") if c["use_vision"] else None
        if u:
            usage_txt += f" · reading {c['label']}: {u}"
    ss.result = {"res": res, "xlsx": reco.to_excel(res), "usage": usage_txt, "instruction": instruction,
                 "chat_ctx": ai.chat_context(res, instruction)}
    ss.chat = []   # new reconciliation -> fresh chat


# =============================== 5. Results ===============================
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


R = ss.result
if R:
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
        st.warning(f"🤖 AI step did not run — showing rule-based results only. ({res.ai_status[7:]})")
    elif res.ai_status:
        st.caption(f"🤖 AI: {res.ai_status}" + (f" · {R['usage']}" if R["usage"] else ""))

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

    tabs = [f"✅ Matched ({len(res.matched)})", f"Only in {la} ({len(res.only_a)})", f"Only in {lb} ({len(res.only_b)})",
            f"⚠️ Amount Mismatch ({len(res.mismatch)})", f"🤖 AI Matched ({len(res.ai_matched)})",
            f"🤖 AI Notes ({len(res.ai_notes)})", "📄 Entries read"]
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
    with t[4]:
        st.caption("Groups proposed by AI and verified by the app (ids exist, used once, amounts add up). "
                   "AI द्वारा मिलाई गई और ऐप द्वारा जाँची गई एंट्री।")
        st.dataframe(fmt(res.ai_matched), width="stretch", hide_index=True)
    with t[5]:
        st.caption("Observations and suggested actions from AI (incl. proposals the app rejected).")
        st.dataframe(res.ai_notes, width="stretch", hide_index=True)
    with t[6]:
        st.caption("Rows read from each file — check them against the originals (पढ़ी गई एंट्री जाँच लें). "
                   "Wrong? Fix header row / column mapping in ⚙️ Settings and submit again.")
        e1, e2 = st.columns(2)
        for col, led, lab in ((e1, res.a, la), (e2, res.b, lb)):
            with col:
                st.markdown(f"**{lab}** — {len(led.txns)} entries, opening "
                            f"{reco.drcr(led.opening) if led.opening_found else 'not found'}")
                for w in led.warnings:
                    st.caption(f"ℹ️ {w}")
                show = led.txns[["row", "date", "ref", "narration", "debit", "credit"]].rename(
                    columns={"row": "Line", "date": "Date", "ref": "Ref", "narration": "Particulars",
                             "debit": "Debit", "credit": "Credit"})
                st.dataframe(fmt(show), width="stretch", hide_index=True, height=300)

# =============================== 6. Chat ===============================
st.markdown("---")
st.subheader("💬 AI se poocho · Ask AI")
st.caption("Ask anything about this reconciliation (Hindi / English / Hinglish) — e.g. 'SI-110 ka kya karna hai?', "
           "'party ko kya email bheju?'. General accounting questions are fine too.")
for m in ss.chat:
    with st.chat_message(m["role"]):
        st.markdown(m["content"])
question = st.chat_input("Apna sawaal likho… / Type your question…", disabled=not API_KEY)
if question:
    with st.chat_message("user"):
        st.markdown(question)
    ctx = R["chat_ctx"] if R else ai.chat_context(None)
    with st.chat_message("assistant"):
        with st.spinner("🤖 …"):
            try:
                answer = ai.chat_reply(ai.GrokClient(API_KEY, timeout=120), ctx, ss.chat, question)
            except Exception as e:  # noqa: BLE001
                answer = f"⚠️ AI could not answer right now ({type(e).__name__}). Please try again."
        st.markdown(answer)
    ss.chat += [{"role": "user", "content": question}, {"role": "assistant", "content": answer}]
    ss.chat = ss.chat[-30:]
