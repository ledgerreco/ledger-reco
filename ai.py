"""
Grok (xAI) helpers for the ledger reconciliation app.

* ai_match()       – reviews leftovers after the rule passes and proposes grouped matches
                     (many-to-one, one-to-many, split entries, TDS/GST/discount/rounding differences),
                     duplicates and notes. Every proposal is validated in code before it is accepted.
* ai_commentary()  – short English + Hindi summary with action items.
* ai_extract_rows() – reads ledger rows from scanned PDFs / images with Grok vision.

The network call is isolated in GrokClient.call_json(), so tests can pass a fake client.
"""
from __future__ import annotations

import base64
import io
import json
import os
import time
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

import reco

MODEL = "grok-4.7"
BASE_URL = "https://api.x.ai/v1"
MAX_ROWS_PER_CALL = 300
MIN_CONFIDENCE = 0.6
MAX_DIFF_SHARE = 0.10        # an explained difference may be at most 10 % of the group amount
TICKS_PER_USD = 1e10         # xAI usage.cost_in_usd_ticks
# A group may carry a non-zero difference only when it is one of these explainable kinds
EXPLAINED_KINDS = {"tds_gst_difference", "discount_difference", "rounding_difference"}
# reasoning effort per task (grok-4.7 default is "high", which is slow/expensive)
EFFORT = {"match": "medium", "summary": "low", "extract": "low"}


def get_api_key(secrets=None) -> str | None:
    try:
        if secrets is not None and "XAI_API_KEY" in secrets:
            return str(secrets["XAI_API_KEY"]).strip() or None
    except Exception:  # noqa: BLE001  (st.secrets raises if no secrets file)
        pass
    return (os.environ.get("XAI_API_KEY") or "").strip() or None


@dataclass
class Usage:
    calls: int = 0
    input_tokens: int = 0
    output_tokens: int = 0
    cost_ticks: int = 0
    seconds: float = 0.0

    def add(self, u, secs):
        self.calls += 1
        self.seconds += secs
        if u is None:
            return
        d = u.model_dump() if hasattr(u, "model_dump") else dict(u)
        self.input_tokens += int(d.get("input_tokens") or 0)
        self.output_tokens += int(d.get("output_tokens") or 0)
        self.cost_ticks += int(d.get("cost_in_usd_ticks") or 0)

    @property
    def cost_usd(self) -> float:
        return self.cost_ticks / TICKS_PER_USD

    def text(self) -> str:
        c = f", ≈ ${self.cost_usd:.4f}" if self.cost_ticks else ""
        return (f"{self.calls} call(s), {self.input_tokens:,} input + {self.output_tokens:,} output tokens"
                f"{c}, {self.seconds:.0f}s")


class GrokClient:
    def __init__(self, api_key: str, model: str = MODEL, timeout: float = 180.0):
        from openai import OpenAI

        self.client = OpenAI(api_key=api_key, base_url=BASE_URL, timeout=timeout, max_retries=1)
        self.model = model
        self.usage = Usage()

    def call_json(self, system: str, content, schema: dict, name: str, effort: str | None = None) -> dict:
        """content: str or list of Responses-API content parts. Returns parsed JSON."""
        t = time.time()
        extra = {"reasoning": {"effort": effort}} if effort else {}
        r = self.client.responses.create(
            model=self.model, **extra,
            input=[{"role": "system", "content": system}, {"role": "user", "content": content}],
            text={"format": {"type": "json_schema", "name": name, "schema": schema, "strict": True}},
            store=False,
        )
        self.usage.add(getattr(r, "usage", None), time.time() - t)
        return json.loads(r.output_text)


# ======================================================================================
# 1) AI matching of leftovers
# ======================================================================================

MATCH_SCHEMA = {
    "type": "object",
    "properties": {
        "groups": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "a_ids": {"type": "array", "items": {"type": "string"}},
                    "b_ids": {"type": "array", "items": {"type": "string"}},
                    "kind": {"type": "string", "enum": [
                        "one_to_one", "many_to_one", "one_to_many", "many_to_many", "split_entry",
                        "partial_payment", "timing_difference", "tds_gst_difference", "discount_difference",
                        "rounding_difference", "other"]},
                    "explained_difference": {"type": "number"},
                    "difference_reason": {"type": "string"},
                    "reason": {"type": "string"},
                    "confidence": {"type": "number", "minimum": 0, "maximum": 1},
                },
                "required": ["a_ids", "b_ids", "kind", "explained_difference", "difference_reason",
                             "reason", "confidence"],
                "additionalProperties": False,
            },
        },
        "notes": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "ids": {"type": "array", "items": {"type": "string"}},
                    "category": {"type": "string", "enum": [
                        "possible_duplicate", "missing_in_party_books", "missing_in_our_books", "amount_mismatch",
                        "tds", "gst", "discount", "rounding", "timing_difference", "wrong_side_or_sign", "other"]},
                    "note": {"type": "string"},
                    "suggested_action": {"type": "string"},
                },
                "required": ["ids", "category", "note", "suggested_action"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["groups", "notes"],
    "additionalProperties": False,
}

MATCH_SYSTEM = """You are an expert Indian accountant reconciling two ledgers of the same account:
side A = "{la}" and side B = "{lb}". A rule engine already matched the easy one-to-one entries.
You get ONLY the leftovers. Every amount `amt` is already in A's sign convention (Dr = +, Cr = −) and party
entries are already sign-flipped, so matching entries have the SAME sign and (nearly) the same value.

Tasks:
1. Propose `groups` linking A entries with B entries that represent the same transaction(s):
   one payment vs several bills/receipts, an invoice split into several entries on one side, partial payments
   recorded differently, timing differences (dates far apart), or entries whose amounts differ ONLY because of
   TDS / GST / discount / rounding. Sum of A amt minus sum of B amt must be 0 (± {tol}); otherwise set
   explained_difference to exactly that difference and say why in difference_reason (e.g. "TDS 1% u/s 194Q").
   A non-zero difference is allowed ONLY for kind tds_gst_difference / discount_difference / rounding_difference
   with a concrete cause visible in the data; a plain "amounts differ" is NOT a match — put it in notes instead.
   Do not group entries that merely look similar. Each id at most once. Both a_ids and b_ids must be non-empty.
2. Add `notes` for remaining items: possible duplicates (same amount/ref twice on one side), entries missing in
   the other books, amount mismatches with a likely cause, TDS/GST items. Give a concrete suggested_action
   (e.g. "Ask party to book our debit note DN/04").
Pairs flagged `rule_mismatch` share a reference but differ in amount; you may regroup them if a better
combination exists. Keep reasons short (max 20 words, English). confidence is 0–1. Never invent ids."""


def _payload(res: reco.Result, ids_a, ids_b) -> dict:
    def rows(df, ids, side):
        out = []
        for i in ids:
            r = df.loc[i]
            out.append({"id": f"{side}{i}", "date": r["date"].strftime("%Y-%m-%d") if pd.notna(r["date"]) else None,
                        "ref": str(r["ref"])[:30], "narration": str(r["narration"])[:60], "amt": round(float(r["cmp"]), 2)})
        return out
    return {"A": rows(res.A, ids_a, "A"), "B": rows(res.B, ids_b, "B"),
            "rule_mismatch": [[f"A{i}", f"B{j}"] for i, j in res.mism_pairs if i in ids_a and j in ids_b]}


def _chunks(res: reco.Result, ids_a, ids_b, max_rows=MAX_ROWS_PER_CALL):
    """Split leftovers into date-ordered windows of at most max_rows entries (both sides together)."""
    items = [(res.A.at[i, "date"], "A", i) for i in ids_a] + [(res.B.at[j, "date"], "B", j) for j in ids_b]
    items.sort(key=lambda t: (pd.Timestamp.min if pd.isna(t[0]) else t[0], t[1], t[2]))
    for k in range(0, len(items), max_rows):
        part = items[k:k + max_rows]
        yield [i for _, s, i in part if s == "A"], [j for _, s, j in part if s == "B"]


def validate_groups(res: reco.Result, raw_groups: list[dict], allowed_a: set, allowed_b: set, amt_tol: float):
    """Return (accepted, rejected). accepted groups use integer indices."""
    accepted, rejected, used_a, used_b = [], [], set(), set()

    def parse(ids, side, allowed):
        out = []
        for x in ids:
            x = str(x).strip()
            if not x.startswith(side) or not x[1:].isdigit():
                return None, f"bad id {x}"
            v = int(x[1:])
            if v not in allowed:
                return None, f"unknown id {x}"
            out.append(v)
        if len(set(out)) != len(out):
            return None, "id repeated"
        return out, ""

    for g in sorted(raw_groups, key=lambda g: -float(g.get("confidence") or 0)):
        a, why = parse(g.get("a_ids", []), "A", allowed_a)
        b, why2 = (parse(g.get("b_ids", []), "B", allowed_b) if a is not None else (None, ""))
        reason = why or why2
        if not reason and (not a or not b):
            reason = "needs entries on both sides"
        if not reason and (used_a & set(a) or used_b & set(b)):
            reason = "id already used in another group"
        conf = float(g.get("confidence") or 0)
        if not reason and conf < MIN_CONFIDENCE:
            reason = f"low confidence {conf:.2f}"
        if not reason:
            sa, sb = float(res.A.loc[a, "cmp"].sum()), float(res.B.loc[b, "cmp"].sum())
            diff = sa - sb
            if np.sign(sa) != np.sign(sb):
                reason = "opposite Dr/Cr direction"
            elif abs(diff) <= amt_tol + 1e-9:
                status = "Matched"
            elif (g.get("kind") in EXPLAINED_KINDS
                  and abs(diff - float(g.get("explained_difference") or 0)) <= amt_tol + 1e-9
                  and abs(diff) <= MAX_DIFF_SHARE * max(abs(sa), abs(sb))
                  and str(g.get("difference_reason", "")).strip()):
                status = "Matched – difference explained"
            else:
                reason = f"amounts do not agree (difference {diff:,.2f})"
        if reason:
            rejected.append({**g, "rejected_because": reason})
            continue
        used_a |= set(a)
        used_b |= set(b)
        accepted.append({"a": a, "b": b, "kind": g.get("kind", ""), "reason": g.get("reason", ""),
                         "difference_reason": g.get("difference_reason", "") if status != "Matched" else "",
                         "confidence": round(conf, 2), "status": status})
    return accepted, rejected


def _id_label(res, x):
    try:
        side, i = x[0], int(x[1:])
        df = res.A if side == "A" else res.B
        lab = res.label_a if side == "A" else res.label_b
        return f"{lab}: {reco._entry_label(df, i)}"
    except Exception:  # noqa: BLE001
        return x


def ai_match(res: reco.Result, client: GrokClient, amt_tol: float | None = None,
             max_rows: int = MAX_ROWS_PER_CALL) -> reco.Result:
    """Run the AI pass over leftovers (only-A, only-B and rule amount-mismatch pairs). Returns a new Result."""
    amt_tol = res.amt_tol if amt_tol is None else amt_tol
    ids_a = sorted(res.free_a | {i for i, _ in res.mism_pairs})
    ids_b = sorted(res.free_b | {j for _, j in res.mism_pairs})
    if not ids_a and not ids_b:
        out = reco.apply_ai_groups(res, [])
        out.ai_status = "ok – nothing left to review"
        return out
    system = MATCH_SYSTEM.format(la=res.label_a, lb=res.label_b, tol=amt_tol)
    accepted, rejected, notes = [], [], []
    for ca, cb in _chunks(res, ids_a, ids_b, max_rows):
        if not ca and not cb:
            continue
        data = client.call_json(system, json.dumps(_payload(res, ca, cb), ensure_ascii=False), MATCH_SCHEMA,
                                "ledger_ai_match", effort=EFFORT["match"])
        used_a = {i for g in accepted for i in g["a"]}
        used_b = {j for g in accepted for j in g["b"]}
        acc, rej = validate_groups(res, data.get("groups", []), set(ca) - used_a, set(cb) - used_b, amt_tol)
        accepted += acc
        rejected += rej
        notes += data.get("notes", [])

    out = reco.apply_ai_groups(res, accepted)
    note_rows = []
    for n in notes:
        note_rows.append({"Category": n.get("category", ""), "Entries": "; ".join(_id_label(res, x) for x in n.get("ids", [])),
                          "Note": n.get("note", ""), "Suggested Action": n.get("suggested_action", "")})
    for g in rejected:
        note_rows.append({"Category": "ai_proposal_rejected",
                          "Entries": "; ".join(_id_label(res, x) for x in list(g.get("a_ids", [])) + list(g.get("b_ids", []))),
                          "Note": f"{g.get('reason', '')} (confidence {g.get('confidence')})",
                          "Suggested Action": f"Not accepted: {g['rejected_because']} — check manually"})
    out.ai_notes = pd.DataFrame(note_rows, columns=["Category", "Entries", "Note", "Suggested Action"])
    out.ai_status = "ok"
    return out


# ======================================================================================
# 2) Commentary
# ======================================================================================

COMMENT_SCHEMA = {
    "type": "object",
    "properties": {
        "summary_en": {"type": "string"},
        "summary_hi": {"type": "string"},
        "action_items": {"type": "array", "items": {
            "type": "object",
            "properties": {"action": {"type": "string"},
                           "priority": {"type": "string", "enum": ["high", "medium", "low"]}},
            "required": ["action", "priority"], "additionalProperties": False}},
    },
    "required": ["summary_en", "summary_hi", "action_items"],
    "additionalProperties": False,
}

COMMENT_SYSTEM = """You are an Indian accountant writing a short note for a business owner about a ledger
reconciliation between "{la}" (our books) and "{lb}". Use the numbers given; do not invent figures.
summary_en: max 90 words, plain English. summary_hi: the same in simple Hindi (Devanagari).
action_items: 3–7 concrete next steps (who to ask for what document / which entry to pass), most important first.
Amounts in Indian format with ₹ (e.g. ₹1,47,500)."""


def ai_commentary(res: reco.Result, client: GrokClient) -> dict:
    def lst(df, cols, n=40):
        if df is None or df.empty:
            return []
        d = df[[c for c in cols if c in df.columns]].head(n).copy()
        for c in d.columns:
            if "Date" in c:
                d[c] = pd.to_datetime(d[c]).dt.strftime("%d-%m-%Y")
        return d.to_dict("records")
    s = res.summary
    payload = {
        "labels": {"A": res.label_a, "B": res.label_b},
        "closing_A": reco.drcr(s["closing_a"]), "closing_B_own_books": reco.drcr(s["closing_b_own"]),
        "difference_A_minus_B_in_A_view": round(s["difference"], 2),
        "rule_matched": s["matched_count"], "ai_matched_groups": s.get("ai_group_count", 0),
        "statement": [(p, a) for p, a in res.statement.itertuples(index=False)],
        "only_in_A": lst(res.only_a, ["Date", "Ref", "Narration", "Net (Dr+ / Cr-)"]),
        "only_in_B": lst(res.only_b, ["Date", "Ref", "Narration", "Net (Dr+ / Cr-)"]),
        "amount_mismatch": lst(res.mismatch, ["A Ref", "A Amount", "B Ref", "B Amount (A view)", "Difference"]),
        "ai_matched": lst(res.ai_matched, ["Type", "A Entries", "B Entries", "Difference", "Reason"]),
        "ai_notes": lst(res.ai_notes, ["Category", "Entries", "Note", "Suggested Action"]),
    }
    return client.call_json(COMMENT_SYSTEM.format(la=res.label_a, lb=res.label_b),
                            json.dumps(payload, ensure_ascii=False, default=str), COMMENT_SCHEMA, "ledger_ai_summary",
                            effort=EFFORT["summary"])


def run_ai(res: reco.Result, client: GrokClient, with_summary: bool = True) -> reco.Result:
    """AI match + commentary; on any failure returns the rule result with ai_status='error: ...'."""
    try:
        out = ai_match(res, client)
    except Exception as e:  # noqa: BLE001
        res.ai_status = f"error: {type(e).__name__}: {str(e)[:200]}"
        return res
    if with_summary:
        try:
            out.ai_summary = ai_commentary(out, client)
        except Exception as e:  # noqa: BLE001
            out.ai_status = f"ok (summary failed: {type(e).__name__})"
    return out


# ======================================================================================
# 3) Vision extraction (scanned PDFs / images)
# ======================================================================================

EXTRACT_SCHEMA = {
    "type": "object",
    "properties": {
        "readable": {"type": "boolean"},
        "opening_balance": {"type": ["number", "null"]},
        "opening_balance_side": {"type": "string", "enum": ["Dr", "Cr", "none"]},
        "rows": {"type": "array", "items": {
            "type": "object",
            "properties": {
                "date": {"type": ["string", "null"]},
                "ref": {"type": "string"},
                "narration": {"type": "string"},
                "debit": {"type": ["number", "null"]},
                "credit": {"type": ["number", "null"]},
            },
            "required": ["date", "ref", "narration", "debit", "credit"], "additionalProperties": False}},
    },
    "required": ["readable", "opening_balance", "opening_balance_side", "rows"],
    "additionalProperties": False,
}

EXTRACT_SYSTEM = """You read Indian ledger statements / account statements from images (Tally, Busy, Zoho,
bank-style statements, handwritten-looking scans). Extract EVERY transaction row exactly as printed:
date as YYYY-MM-DD (Indian dates are day-first), ref = voucher/invoice/cheque number ('' if none),
narration = particulars, debit and credit as plain numbers (no commas; null if blank).
If the statement has a single Amount column with Dr/Cr, put it into debit or credit accordingly.
Do NOT include opening balance, closing balance, totals, carried/brought forward or running-balance-only lines
as rows; report the opening balance separately. If the page is not a ledger or unreadable, readable=false."""


def render_pages(data: bytes, filename: str, max_pages: int = 10, scale: float = 2.0) -> list[bytes]:
    """PDF pages / image -> list of PNG/JPEG bytes suitable for the vision model."""
    from PIL import Image

    name = filename.lower()
    out = []
    if name.endswith(".pdf"):
        import pypdfium2 as pdfium

        pdf = pdfium.PdfDocument(data)
        for k in range(min(len(pdf), max_pages)):
            img = pdf[k].render(scale=scale).to_pil().convert("L")
            buf = io.BytesIO()
            img.save(buf, format="PNG", optimize=True)
            out.append(buf.getvalue())
    else:
        img = Image.open(io.BytesIO(data))
        img = img.convert("L")
        if max(img.size) > 2400:
            img.thumbnail((2400, 2400))
        buf = io.BytesIO()
        img.save(buf, format="PNG", optimize=True)
        out.append(buf.getvalue())
    return out


def ai_extract_rows(images: list[bytes], client: GrokClient, pages_per_call: int = 4) -> tuple[pd.DataFrame, dict]:
    """Return (raw grid with header row, info) to feed into reco.detect_header_row / apply_header flow."""
    rows, opening, side, readable = [], None, "none", False
    for k in range(0, len(images), pages_per_call):
        content = [{"type": "input_text", "text": f"Ledger pages {k + 1}–{k + len(images[k:k + pages_per_call])}. Extract the rows."}]
        for img in images[k:k + pages_per_call]:
            content.append({"type": "input_image", "detail": "high",
                            "image_url": "data:image/png;base64," + base64.b64encode(img).decode()})
        data = client.call_json(EXTRACT_SYSTEM, content, EXTRACT_SCHEMA, "ledger_rows", effort=EFFORT["extract"])
        readable = readable or bool(data.get("readable"))
        if opening is None and data.get("opening_balance") not in (None, 0):
            opening, side = data["opening_balance"], data.get("opening_balance_side", "none")
        rows += data.get("rows", [])
    grid = [["Date", "Ref No", "Particulars", "Debit", "Credit"]]
    if opening is not None and side in ("Dr", "Cr"):
        first = next((r.get("date") for r in rows if r.get("date")), None)
        grid.append([first, "", "Opening Balance", opening if side == "Dr" else None, opening if side == "Cr" else None])
    for r in rows:
        grid.append([r.get("date"), r.get("ref", ""), r.get("narration", ""), r.get("debit"), r.get("credit")])
    info = {"source": "ai-vision", "pages": len(images), "rows": len(rows), "readable": readable,
            "warnings": ["Rows were read by Grok AI from an image — please check the preview against the original."]}
    return pd.DataFrame(grid, dtype=object), info


def text_parse_looks_poor(led: reco.Ledger, df: pd.DataFrame) -> bool:
    """Heuristic: few rows, or many data lines without amounts / dates."""
    n = len(led.txns)
    if n < 3:
        return True
    no_date = led.txns["date"].isna().mean()
    return no_date > 0.2 or n < 0.4 * max(1, len(df.dropna(how="all")))
