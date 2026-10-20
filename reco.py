"""
Core ledger reconciliation logic (no Streamlit dependency).

Flow:
    raw, info = load_file(bytes, filename)      # Excel / CSV / PDF  -> raw grid (no header)
    hdr      = detect_header_row(raw)
    df       = apply_header(raw, hdr)
    mapping  = detect_columns(df)                # {'date','ref','narration','debit','credit','amount','drcr'}
    led      = standardize(df, mapping)          # -> Ledger(txns, opening, stated_closing, warnings)
    res      = reconcile(ledA, ledB, flip_b=True, date_tol=7, amt_tol=1.0)
    xlsx     = to_excel(res, "Our Books", "Party Ledger")

Everything happens in memory; nothing is written to disk.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

# --------------------------------------------------------------------------------------
# Loading
# --------------------------------------------------------------------------------------


class NeedsVisionError(Exception):
    """File has no machine-readable text (scanned PDF or an image) -> needs AI (vision) extraction."""


class ScannedPDFError(NeedsVisionError):
    """Raised when a PDF has no extractable text (scanned / image-only)."""


IMAGE_EXT = (".png", ".jpg", ".jpeg")


def list_sheets(data: bytes, filename: str) -> list[str]:
    name = filename.lower()
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        engine = "xlrd" if name.endswith(".xls") else "openpyxl"
        return pd.ExcelFile(io.BytesIO(data), engine=engine).sheet_names
    return []


def load_file(data: bytes, filename: str, sheet=0) -> tuple[pd.DataFrame, dict]:
    """Return a raw cell grid (header=None, all object dtype) plus info about how it was read."""
    name = filename.lower()
    info = {"source": None, "warnings": []}
    if name.endswith(IMAGE_EXT):
        raise NeedsVisionError("This is an image. Turn on Grok AI to read ledger rows from images.")
    if name.endswith(".pdf"):
        raw, src, warns = read_pdf(data)
        info["source"] = src
        info["warnings"] += warns
    elif name.endswith((".xlsx", ".xlsm")):
        raw = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=None, dtype=object, engine="openpyxl")
        info["source"] = "excel"
    elif name.endswith(".xls"):
        raw = pd.read_excel(io.BytesIO(data), sheet_name=sheet, header=None, dtype=object, engine="xlrd")
        info["source"] = "excel"
    elif name.endswith((".csv", ".txt")):
        raw = None
        for enc in ("utf-8-sig", "cp1252", "latin1"):
            try:
                raw = pd.read_csv(io.BytesIO(data), header=None, dtype=str, sep=None, engine="python",
                                  encoding=enc, skip_blank_lines=False, on_bad_lines="skip")
                break
            except Exception:  # noqa: BLE001
                continue
        if raw is None:
            raise ValueError("Could not read CSV file")
        info["source"] = "csv"
    else:
        raise ValueError(f"Unsupported file type: {filename}")
    raw = raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    raw.columns = range(raw.shape[1])
    return raw, info


# ---------------------------------- PDF ----------------------------------------------

DATE_RE = re.compile(
    r"^(\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}|\d{4}-\d{2}-\d{2}|\d{1,2}[- ]?[A-Za-z]{3,9}[-, ]+\d{2,4})"
)
AMT_RE = re.compile(r"^\(?-?₹?(\d{1,3}(,\d{2,3})+|\d+)(\.\d{1,2})?\)?(Dr|Cr|DR|CR|dr|cr)?\.?$")


def read_pdf(data: bytes) -> tuple[pd.DataFrame, str, list[str]]:
    import pdfplumber

    warns: list[str] = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        text_len = 0
        table_rows: list[list] = []
        pages_words = []
        for page in pdf.pages:
            txt = page.extract_text() or ""
            text_len += len(txt.strip())
            for tbl in page.extract_tables() or []:
                for row in tbl:
                    if row and any(c not in (None, "") for c in row):
                        table_rows.append([(c or "").replace("\n", " ").strip() for c in row])
            pages_words.append(page.extract_words(keep_blank_chars=False, use_text_flow=False))
        if text_len < 20:
            raise ScannedPDFError(
                "This PDF has no readable text (it looks scanned / image-based). "
                "Turn on Grok AI to read it, or export the ledger as Excel/CSV or a text PDF."
            )

    # Prefer tables if they look like a ledger (enough rows with a date-ish first cells and amounts)
    if table_rows:
        width = max(len(r) for r in table_rows)
        grid = pd.DataFrame([r + [""] * (width - len(r)) for r in table_rows], dtype=object)
        dated = sum(1 for r in table_rows if any(DATE_RE.match(str(c)) for c in r[:3]))
        if width >= 3 and dated >= 2:
            return grid.replace("", np.nan), "pdf-table", warns
        warns.append("PDF tables found but did not look like a ledger; used text parsing instead.")

    grid = _pdf_text_parse(pages_words)
    if grid.empty:
        raise ValueError("Could not find any dated transaction lines in the PDF.")
    warns.append("PDF had no ruled table; rows were parsed from text lines — please check the preview.")
    return grid, "pdf-text", warns


def _group_lines(words, ytol=3):
    lines: list[list[dict]] = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        if lines and abs(lines[-1][0]["top"] - w["top"]) <= ytol:
            lines[-1].append(w)
        else:
            lines.append([w])
    return [sorted(line, key=lambda w: w["x0"]) for line in lines]


def _merge_date_words(line):
    """Join words like '01' 'Apr' '2024' into a single date token."""
    out, i = [], 0
    while i < len(line):
        merged = False
        for n in (3, 2):
            if i + n <= len(line):
                cand = " ".join(w["text"] for w in line[i:i + n])
                if re.fullmatch(r"\d{1,2}[- ][A-Za-z]{3,9}[-, ]+\d{2,4}", cand):
                    out.append({**line[i], "text": cand, "x1": line[i + n - 1]["x1"]})
                    i += n
                    merged = True
                    break
        if not merged:
            out.append(line[i])
            i += 1
    return out


def _pdf_text_parse(pages_words) -> pd.DataFrame:
    """Line-based parser: date at start, amounts are numeric tokens; columns assigned by x-position
    of header words (Debit/Credit/Balance/Amount) when available."""
    header_cols: dict[str, float] = {}
    rows = []
    for words in pages_words:
        for line in _group_lines(words):
            line = _merge_date_words(line)
            texts = [w["text"] for w in line]
            low = [t.lower().strip(".:") for t in texts]
            if not header_cols or any(k in low for k in ("debit", "credit")):
                found = {}
                for w, t in zip(line, low):
                    for key, alts in (("Debit", ("debit", "dr", "withdrawal")), ("Credit", ("credit", "cr", "deposit")),
                                      ("Balance", ("balance",)), ("Amount", ("amount",))):
                        if t in alts:
                            found[key] = w["x1"]
                if ("Debit" in found and "Credit" in found) or ("Amount" in found and not header_cols):
                    header_cols = found
                    continue
            if not line or not DATE_RE.match(line[0]["text"]):
                continue
            date = line[0]["text"]
            amounts, others = [], []
            for w in line[1:]:
                if AMT_RE.match(w["text"]) and re.search(r"[.,]", w["text"]) :
                    amounts.append(w)
                else:
                    others.append(w["text"])
            if not amounts:
                continue
            row = {"Date": date, "Debit": None, "Credit": None, "Amount": None, "Balance": None}
            if header_cols:
                for w in amounts:
                    key = min(header_cols, key=lambda k: abs(header_cols[k] - w["x1"]))
                    row[key] = w["text"]
            else:
                row["Amount"] = amounts[0]["text"]
                if len(amounts) > 1:
                    row["Balance"] = amounts[-1]["text"]
            ref = ""
            for t in others:
                if re.search(r"\d", t) and re.fullmatch(r"[A-Za-z]{0,8}[-/#]?[A-Za-z0-9/-]*\d[A-Za-z0-9/-]*", t):
                    ref = t
                    others.remove(t)
                    break
            row["Ref No"] = ref
            row["Particulars"] = " ".join(others)
            rows.append(row)
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df = df[["Date", "Ref No", "Particulars", "Debit", "Credit", "Amount", "Balance"]]
    df = df.dropna(axis=1, how="all")
    # Return as a raw grid with the header as first row so it goes through the same flow
    grid = pd.concat([pd.DataFrame([df.columns.tolist()]), pd.DataFrame(df.values)], ignore_index=True)
    return grid.astype(object)


# --------------------------------------------------------------------------------------
# Header / column detection
# --------------------------------------------------------------------------------------

SYNONYMS = {
    "date": ["date", "dt", "txn date", "transaction date", "voucher date", "vch date", "invoice date",
             "inv date", "doc date", "document date", "posting date", "bill date", "value date"],
    "ref": ["vch no", "voucher no", "voucher number", "vch no.", "invoice no", "invoice number", "inv no",
            "bill no", "bill number", "ref", "ref no", "reference", "reference no", "doc no", "document no",
            "document number", "cheque no", "chq no", "transaction no", "txn no", "number", "no"],
    "narration": ["particulars", "narration", "description", "details", "remarks", "account", "ledger",
                  "transaction details", "memo", "name"],
    "debit": ["debit", "dr", "debit amount", "debit amt", "dr amount", "withdrawal", "debit inr", "debit rs"],
    "credit": ["credit", "cr", "credit amount", "credit amt", "cr amount", "deposit", "credit inr", "credit rs"],
    "amount": ["amount", "amt", "value", "net amount", "transaction amount", "amount inr", "amount rs"],
    "drcr": ["dr cr", "drcr", "d c", "dc", "dr or cr", "type dr cr"],
}
FIELDS = ["date", "ref", "narration", "debit", "credit", "amount", "drcr"]


def _norm_hdr(s) -> str:
    s = str(s).lower().replace("₹", " rs ")
    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
    s = re.sub(r"\b(rs|inr)\b", "", s).strip() if s not in ("rs", "inr") else s
    return re.sub(r"\s+", " ", s)


def _score(header: str, field: str) -> int:
    h = _norm_hdr(header)
    if not h:
        return 0
    best = 0
    for syn in SYNONYMS[field]:
        if h == syn:
            best = max(best, 10)
        elif len(syn) > 2 and (h.startswith(syn + " ") or h.endswith(" " + syn) or f" {syn} " in f" {h} "):
            best = max(best, 6)
        elif len(syn) > 3 and syn in h:
            best = max(best, 3)
    return best


def detect_header_row(raw: pd.DataFrame, max_scan: int = 40) -> int:
    best_row, best_score = 0, -1
    for i in range(min(max_scan, len(raw))):
        cells = [c for c in raw.iloc[i].tolist() if pd.notna(c) and str(c).strip()]
        if len(cells) < 2:
            continue
        fields_hit = set()
        for c in cells:
            for f in FIELDS:
                if _score(c, f) >= 6:
                    fields_hit.add(f)
        score = len(fields_hit)
        if score > best_score:
            best_row, best_score = i, score
    return best_row if best_score >= 2 else 0


def apply_header(raw: pd.DataFrame, header_row: int) -> pd.DataFrame:
    hdr = raw.iloc[header_row].tolist()
    cols, seen = [], {}
    for j, h in enumerate(hdr):
        name = str(h).strip() if pd.notna(h) and str(h).strip() else f"Column {j + 1}"
        if name in seen:
            seen[name] += 1
            name = f"{name} ({seen[name]})"
        else:
            seen[name] = 0
        cols.append(name)
    df = raw.iloc[header_row + 1:].copy()
    df.columns = cols
    return df.reset_index(drop=True)


def detect_columns(df: pd.DataFrame) -> dict:
    """Map logical fields -> column names (or None)."""
    cols = list(df.columns)
    scores = {(f, c): _score(c, f) for f in FIELDS for c in cols}
    mapping: dict = {f: None for f in FIELDS}
    used = set()
    for (f, c), s in sorted(scores.items(), key=lambda kv: -kv[1]):
        if s < 3 or mapping[f] is not None or c in used:
            continue
        mapping[f] = c
        used.add(c)
    # If only one of debit/credit found + an amount, fine. If neither debit/credit nor amount,
    # guess the numeric columns.
    if not (mapping["debit"] or mapping["credit"] or mapping["amount"]):
        numeric = [c for c in cols if c not in used and parse_amount_series(df[c]).notna().mean() > 0.5]
        if len(numeric) >= 2:
            mapping["debit"], mapping["credit"] = numeric[0], numeric[1]
        elif numeric:
            mapping["amount"] = numeric[0]
    if mapping["date"] is None:
        for c in cols:
            if c not in used and parse_dates(df[c]).notna().mean() > 0.5:
                mapping["date"] = c
                break
    # amount column only meaningful when debit/credit pair missing
    if mapping["debit"] and mapping["credit"]:
        mapping["amount"] = None
    return mapping


# --------------------------------------------------------------------------------------
# Parsing values
# --------------------------------------------------------------------------------------


def parse_amount(v):
    """'1,23,456.00 Dr' -> 123456.0 ; '(500)' -> -500 ; '500 Cr' -> -500 ; '' -> NaN."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return np.nan
    if isinstance(v, (int, float, np.integer, np.floating)) and not isinstance(v, bool):
        return float(v)
    s = str(v).strip()
    if not s or s in ("-", "--", "nil", "Nil", "NIL"):
        return np.nan
    sign = 1.0
    m = re.search(r"\b(dr|cr)\.?$", s, re.I)
    if m:
        if m.group(1).lower() == "cr":
            sign = -1.0
        s = s[: m.start()].strip()
    if s.startswith("(") and s.endswith(")"):
        sign *= -1
        s = s[1:-1]
    s = re.sub(r"(₹|rs\.?|inr)", "", s, flags=re.I).replace(",", "").replace(" ", "")
    if s.endswith("-"):
        sign *= -1
        s = s[:-1]
    try:
        return sign * float(s)
    except ValueError:
        return np.nan


def parse_amount_series(s: pd.Series) -> pd.Series:
    return s.map(parse_amount).astype(float)


def parse_dates(s: pd.Series) -> pd.Series:
    def one(v):
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return pd.NaT
        if isinstance(v, (pd.Timestamp,)) or hasattr(v, "year"):
            return pd.Timestamp(v)
        if isinstance(v, (int, float, np.integer, np.floating)):
            if 20000 < float(v) < 80000:  # Excel serial
                return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(v))
            return pd.NaT
        t = str(v).strip()
        if not t or not re.search(r"\d", t):
            return pd.NaT
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}([ T].*)?", t):
            return pd.to_datetime(t[:10], errors="coerce")
        return pd.to_datetime(t, errors="coerce", dayfirst=True)
    return pd.Series([one(v) for v in s], index=s.index, dtype="datetime64[ns]")


# --------------------------------------------------------------------------------------
# Reference normalisation
# --------------------------------------------------------------------------------------

REF_PREFIXES = r"(TAX\s*INVOICE|INVOICE|INV|BILL|VOUCHER|VCH|REF|DOC|SALES|PURCHASE|PUR|SAL|SI|PI|CN|DN|NO|NUMBER|#)"


def norm_ref(r) -> str:
    if r is None or (isinstance(r, float) and np.isnan(r)):
        return ""
    s = str(r).upper().strip()
    if isinstance(r, float) and float(r).is_integer():
        s = str(int(r))
    s = re.sub(r"\.0$", "", s)
    prev = None
    while prev != s:  # strip repeated leading prefixes like "INV NO."
        prev = s
        s = re.sub(rf"^\s*{REF_PREFIXES}\b[\s.:#/-]*", "", s)
    s = re.sub(r"[^A-Z0-9]", "", s)
    return s


def ref_number(r) -> str:
    """Core number of a reference, ignoring fiscal-year pieces: 'INV/2024-25/0105' -> '105'."""
    if r is None or (isinstance(r, float) and np.isnan(r)):
        return ""
    s = str(r).upper()
    s = re.sub(r"\.0$", "", s)
    s2 = re.sub(r"(19|20)\d{2}\s*[-/]\s*((19|20)?\d{2})", " ", s)  # 2024-25 / 2024/2025
    groups = re.findall(r"\d+", s2) or re.findall(r"\d+", s)
    if not groups:
        return ""
    g = groups[-1].lstrip("0")
    return g if len(g) >= 2 else ""


# --------------------------------------------------------------------------------------
# Standardise a ledger
# --------------------------------------------------------------------------------------

OPENING_RE = re.compile(r"opening\s*bal|op\.?\s*bal|\bb\s*/\s*f\b|brought\s*f(?:or)?w(?:ar)?d", re.I)
CLOSING_RE = re.compile(r"closing\s*bal|cl\.?\s*bal|\bc\s*/\s*f\b|carried\s*f(?:or)?w(?:ar)?d", re.I)
TOTAL_RE = re.compile(r"^\s*(grand\s*)?totals?\b|^\s*sub\s*-?\s*total", re.I)


@dataclass
class Ledger:
    txns: pd.DataFrame
    opening: float = 0.0
    opening_found: bool = False
    stated_closing: float | None = None
    warnings: list = field(default_factory=list)

    @property
    def closing(self) -> float:
        return self.opening + float(self.txns["signed"].sum()) if len(self.txns) else self.opening


def standardize(df: pd.DataFrame, mapping: dict) -> Ledger:
    """Turn a mapped DataFrame into transactions with signed = Debit - Credit (Dr positive)."""
    warns = []
    n = len(df)
    get = lambda f: df[mapping[f]] if mapping.get(f) and mapping[f] in df.columns else pd.Series([None] * n, index=df.index)  # noqa: E731
    dates = parse_dates(get("date"))
    ref = get("ref").map(lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) else
                         (str(int(v)) if isinstance(v, float) and v.is_integer() else str(v).strip()))
    narr = get("narration").map(lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip())

    if mapping.get("debit") or mapping.get("credit"):
        dr = parse_amount_series(get("debit")).abs() if mapping.get("debit") else pd.Series(np.nan, index=df.index)
        cr = parse_amount_series(get("credit")).abs() if mapping.get("credit") else pd.Series(np.nan, index=df.index)
        has = dr.notna() | cr.notna()
        dr, cr = dr.fillna(0.0), cr.fillna(0.0)
    elif mapping.get("amount"):
        amt = parse_amount_series(get("amount"))
        if mapping.get("drcr"):
            t = get("drcr").astype(str).str.strip().str.lower()
            sign = np.where(t.str.startswith("c"), -1.0, 1.0)
            amt = amt.abs() * sign
        has = amt.notna()
        amt = amt.fillna(0.0)
        dr, cr = amt.clip(lower=0), (-amt).clip(lower=0)
    else:
        raise ValueError("Please map Debit/Credit columns or an Amount column.")

    text = (narr + " " + ref).str.strip()
    # Totals rows: no date, no particulars/ref (e.g. Tally column totals at the bottom)
    blank_total = dates.isna() & (text == "")
    # Tally puts the date only on the first line of a voucher -> forward fill
    dates = dates.ffill()
    is_open = text.str.contains(OPENING_RE)
    is_close = text.str.contains(CLOSING_RE)
    is_total = narr.str.match(TOTAL_RE) | ref.str.match(TOTAL_RE) | blank_total

    opening, opening_found = 0.0, False
    if (is_open & has).any():
        r = is_open & has
        opening = float((dr[r] - cr[r]).iloc[0])
        opening_found = True
    stated_closing = None
    if (is_close & has).any():
        r = is_close & has
        stated_closing = float((dr[r] - cr[r]).iloc[0])
        computed = opening + float((dr - cr)[has & ~is_open & ~is_close & ~is_total].sum())
        # Tally shows Closing Balance on the opposite side (to balance the totals) -> flip it back
        if abs(stated_closing + computed) < 1 and abs(computed) > 1:
            stated_closing = -stated_closing
        if abs(stated_closing - computed) > 1:
            warns.append(f"Stated closing balance {stated_closing:,.2f} differs from computed {computed:,.2f}.")

    keep = has & ~is_open & ~is_close & ~is_total & ((dr != 0) | (cr != 0))
    if (keep & dates.isna()).any():
        warns.append(f"{int((keep & dates.isna()).sum())} row(s) have no readable date.")
    tx = pd.DataFrame({
        "row": df.index[keep] + 1,
        "date": dates[keep],
        "ref": ref[keep],
        "narration": narr[keep],
        "debit": dr[keep].round(2),
        "credit": cr[keep].round(2),
    }).reset_index(drop=True)
    tx["signed"] = (tx["debit"] - tx["credit"]).round(2)
    tx["ref_norm"] = tx["ref"].map(norm_ref)
    tx["ref_num"] = tx["ref"].map(ref_number)
    # If no ref column, try to pull an invoice-like number out of narration
    if not mapping.get("ref"):
        guess = tx["narration"].str.extract(r"((?:INV|BILL|INVOICE)[\s.#:/-]*[A-Z0-9/-]*\d+)", flags=re.I)[0].fillna("")
        tx["ref_norm"] = guess.map(norm_ref)
        tx["ref_num"] = guess.map(ref_number)
    return Ledger(tx, opening, opening_found, stated_closing, warns)


# --------------------------------------------------------------------------------------
# Reconciliation
# --------------------------------------------------------------------------------------


@dataclass
class Result:
    a: Ledger
    b: Ledger
    flip_b: bool
    matched: pd.DataFrame
    mismatch: pd.DataFrame
    only_a: pd.DataFrame
    only_b: pd.DataFrame
    summary: dict
    statement: pd.DataFrame
    label_a: str = "Ledger A"
    label_b: str = "Ledger B"
    # internals (kept so an AI pass can be applied on the leftovers)
    A: pd.DataFrame | None = None
    B: pd.DataFrame | None = None
    rule_matches: list = field(default_factory=list)    # (i, j, kind)
    mism_pairs: list = field(default_factory=list)      # (i, j)  after AI removals
    free_a: set = field(default_factory=set)
    free_b: set = field(default_factory=set)
    amt_tol: float = 1.0
    # AI results (empty when AI not used)
    ai_groups: list = field(default_factory=list)       # accepted, validated groups
    ai_matched: pd.DataFrame = field(default_factory=pd.DataFrame)
    ai_notes: pd.DataFrame = field(default_factory=pd.DataFrame)
    ai_summary: dict = field(default_factory=dict)
    ai_status: str = ""                                  # "", "ok", "error: ..."


def _greedy(pairs: list[tuple]) -> list[tuple[int, int]]:
    """pairs: (sort_key..., i, j) already sorted; pick one-to-one."""
    used_a, used_b, out = set(), set(), []
    for p in pairs:
        i, j = p[-2], p[-1]
        if i in used_a or j in used_b:
            continue
        used_a.add(i)
        used_b.add(j)
        out.append((i, j))
    return out


def _days(da, db):
    if pd.isna(da) or pd.isna(db):
        return 9999
    return abs((da - db).days)


def reconcile(a: Ledger, b: Ledger, flip_b: bool = True, date_tol: int = 7, amt_tol: float = 1.0,
              label_a: str = "Ledger A", label_b: str = "Ledger B") -> Result:
    A, B = a.txns.copy(), b.txns.copy()
    A["cmp"] = A["signed"]
    B["cmp"] = -B["signed"] if flip_b else B["signed"]
    free_a, free_b = set(A.index), set(B.index)
    matches: list[tuple[int, int, str]] = []
    mism: list[tuple[int, int]] = []

    def take(pairs, kind):
        for i, j in _greedy(pairs):
            matches.append((i, j, kind))
            free_a.discard(i)
            free_b.discard(j)

    # ---- Pass 1: reference + amount
    pairs = []
    for i in free_a:
        ra, na = A.at[i, "ref_norm"], A.at[i, "ref_num"]
        if not ra and not na:
            continue
        for j in free_b:
            rb, nb = B.at[j, "ref_norm"], B.at[j, "ref_num"]
            exact = bool(ra) and ra == rb
            numeq = bool(na) and na == nb
            if not (exact or numeq):
                continue
            d = abs(A.at[i, "cmp"] - B.at[j, "cmp"])
            if d <= amt_tol + 1e-9:
                pairs.append((0 if exact else 1, d, _days(A.at[i, "date"], B.at[j, "date"]), i, j))
    take(sorted(pairs), "1. Reference + Amount")

    # ---- Amount mismatch: same reference, same direction, different amount
    pairs = []
    for i in free_a:
        ra, na = A.at[i, "ref_norm"], A.at[i, "ref_num"]
        if not ra and not na:
            continue
        for j in free_b:
            rb, nb = B.at[j, "ref_norm"], B.at[j, "ref_num"]
            exact = bool(ra) and ra == rb and len(ra) >= 2
            numeq = bool(na) and na == nb and len(na) >= 3
            if not (exact or numeq):
                continue
            if np.sign(A.at[i, "cmp"]) != np.sign(B.at[j, "cmp"]):
                continue
            pairs.append((0 if exact else 1, abs(A.at[i, "cmp"] - B.at[j, "cmp"]), i, j))
    for i, j in _greedy(sorted(pairs)):
        mism.append((i, j))
        free_a.discard(i)
        free_b.discard(j)

    # ---- Pass 2 & 3: amount (+ date)
    def amount_pass(max_days, kind):
        ia, jb = sorted(free_a), sorted(free_b)
        if not ia or not jb:
            return
        va = A.loc[ia, "cmp"].to_numpy()[:, None]
        vb = B.loc[jb, "cmp"].to_numpy()[None, :]
        diff = np.abs(va - vb)
        cand = np.argwhere((diff <= amt_tol + 1e-9) & (va != 0))
        pairs = []
        for x, y in cand:
            i, j = ia[x], jb[y]
            dd = _days(A.at[i, "date"], B.at[j, "date"])
            if max_days is not None and dd > max_days:
                continue
            pairs.append((dd, diff[x, y], i, j))
        take(sorted(pairs), kind)

    amount_pass(date_tol, f"2. Amount + Date (±{date_tol}d)")
    amount_pass(None, "3. Amount only")

    return finalize(a, b, A, B, matches, mism, free_a, free_b, flip_b, label_a, label_b, amt_tol)


def apply_ai_groups(res: Result, groups: list[dict]) -> Result:
    """Re-build the result with validated AI groups. Each group: {'a': [idx], 'b': [idx], 'kind', 'reason',
    'confidence', 'difference_reason'}. Entries may come from 'only' lists or from rule amount-mismatch pairs."""
    used_a = {i for g in groups for i in g["a"]}
    used_b = {j for g in groups for j in g["b"]}
    mism = [(i, j) for i, j in res.mism_pairs if i not in used_a and j not in used_b]
    # a mismatch pair broken up by AI -> its other half goes back to the "only" lists
    free_a = (res.free_a | {i for i, j in res.mism_pairs if (i, j) not in mism}) - used_a
    free_b = (res.free_b | {j for i, j in res.mism_pairs if (i, j) not in mism}) - used_b
    out = finalize(res.a, res.b, res.A, res.B, res.rule_matches, mism, free_a, free_b, res.flip_b,
                   res.label_a, res.label_b, res.amt_tol, ai_groups=groups)
    return out


def _side(df, idx, prefix):
    r = df.loc[idx]
    return {f"{prefix} Row": r["row"], f"{prefix} Date": r["date"], f"{prefix} Ref": r["ref"],
            f"{prefix} Narration": r["narration"], f"{prefix} Debit": r["debit"], f"{prefix} Credit": r["credit"]}


def _entry_label(df, idx):
    r = df.loc[idx]
    d = r["date"].strftime("%d-%m-%Y") if pd.notna(r["date"]) else "?"
    ref = r["ref"] or (r["narration"][:25] if r["narration"] else "-")
    return f"{ref} {d} ({drcr(r['signed'])})"


def finalize(a, b, A, B, matches, mism, free_a, free_b, flip_b, label_a, label_b, amt_tol, ai_groups=None) -> Result:
    ai_groups = ai_groups or []
    rows = []
    for i, j, kind in matches:
        d = {"Match Type": kind, **_side(A, i, "A"), **_side(B, j, "B")}
        d["Difference"] = round(A.at[i, "cmp"] - B.at[j, "cmp"], 2)
        dd = _days(A.at[i, "date"], B.at[j, "date"])
        d["Days Apart"] = dd if dd != 9999 else None
        rows.append(d)
    matched = pd.DataFrame(rows)
    if len(matched):
        matched = matched.sort_values(["Match Type", "A Date"]).reset_index(drop=True)

    rows = []
    for i, j in mism:
        d = {**_side(A, i, "A"), **_side(B, j, "B")}
        d["A Amount"] = A.at[i, "cmp"]
        d["B Amount (A view)"] = B.at[j, "cmp"]
        d["Difference"] = round(A.at[i, "cmp"] - B.at[j, "cmp"], 2)
        rows.append(d)
    mismatch = pd.DataFrame(rows)

    def only(df, free):
        o = df.loc[sorted(free), ["row", "date", "ref", "narration", "debit", "credit", "signed"]].copy()
        o.columns = ["Row", "Date", "Ref", "Narration", "Debit", "Credit", "Net (Dr+ / Cr-)"]
        return o.sort_values("Date").reset_index(drop=True)

    only_a, only_b = only(A, free_a), only(B, free_b)

    rows = []
    for n, g in enumerate(ai_groups, start=1):
        sa, sb = float(A.loc[g["a"], "cmp"].sum()), float(B.loc[g["b"], "cmp"].sum())
        rows.append({
            "Group": n, "Type": g.get("kind", ""),
            "A Entries": "; ".join(_entry_label(A, i) for i in g["a"]), "A Total": round(sa, 2),
            "B Entries": "; ".join(_entry_label(B, j) for j in g["b"]), "B Total (A view)": round(sb, 2),
            "Difference": round(sa - sb, 2), "Status": g.get("status", "Matched"),
            "Reason": g.get("reason", ""), "Difference Reason": g.get("difference_reason", ""),
            "Confidence": g.get("confidence"),
            "A Rows": ", ".join(str(int(A.at[i, "row"])) for i in g["a"]),
            "B Rows": ", ".join(str(int(B.at[j, "row"])) for j in g["b"]),
        })
    ai_matched = pd.DataFrame(rows)

    # ---- Summary & reconciliation statement (everything in A's perspective, Dr positive)
    sgn = -1 if flip_b else 1
    open_a, open_b = a.opening, sgn * b.opening
    close_a, close_b = a.closing, sgn * b.closing
    m_idx_a = [i for i, _, _ in matches]
    m_idx_b = [j for _, j, _ in matches]
    sum_only_a = float(A.loc[sorted(free_a), "cmp"].sum())
    sum_only_b = float(B.loc[sorted(free_b), "cmp"].sum())
    mm_diff = float(mismatch["Difference"].sum()) if len(mismatch) else 0.0
    round_diff = float(A.loc[m_idx_a, "cmp"].sum() - B.loc[m_idx_b, "cmp"].sum())
    ai_diff = float(ai_matched["Difference"].sum()) if len(ai_matched) else 0.0
    ai_a = [i for g in ai_groups for i in g["a"]]
    diff = close_a - close_b

    summary = {
        "rows_a": len(A), "rows_b": len(B),
        "debit_a": float(A["debit"].sum()), "credit_a": float(A["credit"].sum()),
        "debit_b": float(B["debit"].sum()), "credit_b": float(B["credit"].sum()),
        "opening_a": a.opening, "opening_b": b.opening,
        "opening_a_found": a.opening_found, "opening_b_found": b.opening_found,
        "closing_a": close_a, "closing_b_own": b.closing, "closing_b": close_b,
        "stated_closing_a": a.stated_closing, "stated_closing_b": b.stated_closing,
        "matched_count": len(matches), "matched_amount": float(A.loc[m_idx_a, "cmp"].abs().sum()),
        "match_by_type": pd.Series([k for _, _, k in matches]).value_counts().sort_index().to_dict() if matches else {},
        "only_a_count": len(free_a), "only_a_amount": sum_only_a,
        "only_b_count": len(free_b), "only_b_amount": sum_only_b,
        "mismatch_count": len(mism), "mismatch_amount": mm_diff,
        "ai_group_count": len(ai_groups), "ai_entry_count_a": len(ai_a),
        "ai_entry_count_b": sum(len(g["b"]) for g in ai_groups),
        "ai_amount": float(A.loc[ai_a, "cmp"].abs().sum()) if ai_a else 0.0, "ai_difference": ai_diff,
        "difference": diff,
    }
    lines = [
        (f"Closing balance as per {label_a}", close_a),
        (f"Closing balance as per {label_b}" + (" (sign flipped to our view)" if flip_b else ""), close_b),
        ("Difference to explain (A − B)", diff),
        ("Explained by:", None),
        ("  Opening balance difference", open_a - open_b),
        (f"  Entries only in {label_a} ({len(free_a)})", sum_only_a),
        (f"  Entries only in {label_b} ({len(free_b)}) — reversed", -sum_only_b),
        (f"  Amount mismatches, same reference ({len(mism)})", mm_diff),
        ("  Small differences within tolerance on matched entries", round_diff),
    ]
    if ai_groups:
        lines.append((f"  Differences on AI-matched groups ({len(ai_groups)})", ai_diff))
    explained = (open_a - open_b) + sum_only_a - sum_only_b + mm_diff + round_diff + ai_diff
    lines += [("Total explained", explained), ("Unexplained", diff - explained)]
    statement = pd.DataFrame(lines, columns=["Particulars", "Amount"])
    statement["Amount"] = statement["Amount"].map(lambda v: None if v is None else round(v, 2))
    return Result(a, b, flip_b, matched, mismatch, only_a, only_b, summary, statement, label_a, label_b,
                  A=A, B=B, rule_matches=list(matches), mism_pairs=list(mism), free_a=set(free_a),
                  free_b=set(free_b), amt_tol=amt_tol, ai_groups=list(ai_groups), ai_matched=ai_matched)


# --------------------------------------------------------------------------------------
# Formatting & Excel
# --------------------------------------------------------------------------------------


def inr(x, decimals: int = 2) -> str:
    """Indian grouping: 1234567.5 -> '12,34,567.50'."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    neg = x < 0
    s = f"{abs(x):.{decimals}f}"
    whole, _, frac = s.partition(".")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        head = re.sub(r"(\d)(?=(\d{2})+$)", r"\1,", head)
        whole = f"{head},{tail}"
    out = whole + (f".{frac}" if frac else "")
    return f"-{out}" if neg else out


def drcr(x) -> str:
    if x is None or abs(x) < 0.005:
        return "0.00"
    return f"{inr(abs(x))} {'Dr' if x > 0 else 'Cr'}"


INR_FMT = '[>=10000000]##\\,##\\,##\\,##0.00;[>=100000]##\\,##\\,##0.00;##,##0.00'


def summary_table(res: Result) -> pd.DataFrame:
    s, la, lb = res.summary, res.label_a, res.label_b
    rows = [
        ("Transactions (rows)", s["rows_a"], s["rows_b"]),
        ("Opening balance" , drcr(s["opening_a"]) + ("" if s["opening_a_found"] else " (not found)"),
         drcr(s["opening_b"]) + ("" if s["opening_b_found"] else " (not found)")),
        ("Total Debit", inr(s["debit_a"]), inr(s["debit_b"])),
        ("Total Credit", inr(s["credit_a"]), inr(s["credit_b"])),
        ("Closing balance (computed)", drcr(s["closing_a"]), drcr(s["closing_b_own"])),
        ("Matched entries", s["matched_count"], s["matched_count"]),
        ("Unmatched entries (only in this side)", s["only_a_count"], s["only_b_count"]),
        ("Amount mismatch (same ref)", s["mismatch_count"], s["mismatch_count"]),
    ]
    if s.get("ai_group_count"):
        rows.append((f"AI matched groups ({s['ai_group_count']}) — entries", s["ai_entry_count_a"], s["ai_entry_count_b"]))
    return pd.DataFrame(rows, columns=["Item", la, lb])


def to_excel(res: Result) -> bytes:
    la, lb = res.label_a, res.label_b
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter", datetime_format="dd-mmm-yyyy", date_format="dd-mmm-yyyy") as xw:
        wb = xw.book
        money = wb.add_format({"num_format": INR_FMT})
        bold = wb.add_format({"bold": True})
        title = wb.add_format({"bold": True, "font_size": 14})
        boldmoney = wb.add_format({"bold": True, "num_format": INR_FMT})

        # Summary sheet
        ws = wb.add_worksheet("Summary")
        xw.sheets["Summary"] = ws
        ws.write(0, 0, f"Ledger Reconciliation: {la} vs {lb}", title)
        ws.write(1, 0, f"Sign flip for {lb}: {'Yes (their Credit = our Debit)' if res.flip_b else 'No'}")
        st = summary_table(res)
        r0 = 3
        for c, h in enumerate(st.columns):
            ws.write(r0, c, h, bold)
        for r, row in enumerate(st.itertuples(index=False), start=r0 + 1):
            for c, v in enumerate(row):
                ws.write(r, c, v)
        r = r0 + len(st) + 3
        ws.write(r, 0, "Reconciliation Statement (amounts in our view, Dr positive)", title)
        r += 1
        for row in res.statement.itertuples(index=False):
            is_key = row.Particulars.startswith(("Difference", "Total explained", "Unexplained", "Closing"))
            ws.write(r, 0, row.Particulars, bold if is_key else None)
            if row.Amount is not None and not pd.isna(row.Amount):
                ws.write_number(r, 1, row.Amount, boldmoney if is_key else money)
            r += 1
        r += 1
        if res.summary["match_by_type"]:
            ws.write(r, 0, "Matches by pass", bold)
            r += 1
            for k, v in res.summary["match_by_type"].items():
                ws.write(r, 0, k)
                ws.write(r, 1, v)
                r += 1
        ws.write(r + 1, 0, "A = " + la + "   |   B = " + lb)
        r += 3
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        if res.ai_summary:
            ws.write(r, 0, "Grok AI summary", title)
            r += 1
            for key, lab in (("summary_en", "Summary"), ("summary_hi", "सारांश")):
                if res.ai_summary.get(key):
                    ws.write(r, 0, f"{lab}: {res.ai_summary[key]}", wrap)
                    ws.set_row(r, 15 * max(2, len(res.ai_summary[key]) // 70 + 1))
                    r += 1
            for k, item in enumerate(res.ai_summary.get("action_items", []), start=1):
                txt = item.get("action", "") if isinstance(item, dict) else str(item)
                pr = item.get("priority", "") if isinstance(item, dict) else ""
                ws.write(r, 0, f"Action {k}{' [' + pr + ']' if pr else ''}: {txt}", wrap)
                r += 1
        elif res.ai_status:
            ws.write(r, 0, f"Grok AI: {res.ai_status}")
        ws.set_column(0, 0, 62)
        ws.set_column(1, 2, 24)

        def dump(df, name):
            rename = {c: c.replace("A ", f"{la}: ", 1) if c.startswith("A ") else
                      (c.replace("B ", f"{lb}: ", 1) if c.startswith("B ") else c) for c in df.columns}
            d = df.rename(columns=rename) if len(df) else pd.DataFrame({"Info": ["No entries"]})
            d.to_excel(xw, sheet_name=name, index=False)
            w = xw.sheets[name]
            for c, col in enumerate(d.columns):
                is_money = "Reason" not in col and any(k in col for k in ("Debit", "Credit", "Difference", "Amount", "Net", "Total"))
                width = min(max(12, len(str(col)) + 2, int(d[col].astype(str).str.len().max() or 0) + 2), 45)
                w.set_column(c, c, width, money if is_money else None)
            w.freeze_panes(1, 0)
            w.autofilter(0, 0, len(d), len(d.columns) - 1)

        dump(res.matched, "Matched")
        dump(res.only_a, "Only in Ledger A")
        dump(res.only_b, "Only in Ledger B")
        dump(res.mismatch, "Amount Mismatch")
        if res.ai_status:
            dump(res.ai_matched, "AI Matched")
            dump(res.ai_notes, "AI Notes")
    return out.getvalue()


# Convenience one-shot ------------------------------------------------------------------

def ledger_from_file(data: bytes, filename: str, mapping: dict | None = None, header_row: int | None = None):
    raw, info = load_file(data, filename)
    hdr = detect_header_row(raw) if header_row is None else header_row
    df = apply_header(raw, hdr)
    mp = mapping or detect_columns(df)
    led = standardize(df, mp)
    led.warnings = info["warnings"] + led.warnings
    return led, mp, hdr, df, info
