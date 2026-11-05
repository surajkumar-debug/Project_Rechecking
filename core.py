import io, re, difflib
import numpy as np, pandas as pd
from openpyxl.styles import PatternFill, Font

SOURCES = ["Center Data", "PM", "Calling", "CSR", "Audit", "Audit Final"]
DEFAULT_TABS = {"Center Data": ["Center Data"], "PM": ["PM Compile", "PM"], "Calling": ["Calling Sheet", "Calling"],
                "CSR": ["CSR Audit", "CSR"], "Audit": ["Audit Home", "Audit"], "Audit Final": ["Audit Final"]}
MAP_COLS = ["Field", "Source", "Center Data header", "Source header", "Kind"]

# Field, Source, Center Data header (aliases), Source header (aliases), Kind(num/pct/text)
# Aliases: , ya | se alag karo. TC / Day rows special hain.
_ROWS = [
    ["TC", "Center Data", "", "TC Code", ""], ["TC", "PM", "", "TC Code", ""], ["TC", "Calling", "", "TC Code", ""],
    ["TC", "CSR", "", "TC Code", ""], ["TC", "Audit", "", "Venue Code | TC Code", ""], ["TC", "Audit Final", "", "TC_CODE | TC Code", ""],
    ["Day", "Center Data", "", "Exam Day Number", ""],
    ["Actual Cam", "PM", "Actual Camera", "Final Cam Count", "num"],
    ["QC Cam", "PM", "QC Camera", "QC Cam Count", "num"],
    ["Monitor", "PM", "Monitor Count", "Final Monitor Count", "num"],
    ["UPS", "PM", "UPS Count", "Final UPS Count", "num"],
    ["HMS %", "PM", "HMS%", "HMS Avg Hit Per%", "pct"],
    ["DVR Installed", "PM", "Total DVR Count", "HMS DVR Installed Count", "num"],
    ["DVR Live", "PM", "Online DVR Count", "HMS No of DVR Hit", "num"],
    ["VOIP Status", "PM", "VOIP Status", "VOIP Final Installation Status", "text"],
    ["Common Center", "PM", "Common Center", "Common Centers", "text"],
    ["Manpower", "PM", "Manpower Attendance", "Manpower Attendence Days", "num"],
    ["CCTV Installed", "Calling", "CCTV Installation Status", "CCTV Setup Installed", "text"],
    ["ISP", "Calling", "Internet Installation Status", "Internet Installation", "text"],
    ["Monitor Status", "Calling", "Monitor Installation Status", "Monitor Installed", "text"],
    ["UPS Status", "Calling", "UPS.2", "UPS Installed", "text"],
    ["Calling Status", "Calling", "Calling Status", "calling status", "text"],
    ["Remarks", "Calling", "Remarks", "remarks", "text"],
    ["Max Cam (Audit)", "Audit", "MAX Cam RCV as per Data", "As per Audit Max | Unnamed: 38", "num"],
    ["Data %", "Audit", "Data %", "Data %", "pct"],
    ["Audit %", "Audit", "Audit %", "Audit %", "pct"],
    ["Audit Status", "Audit", "Audit Status", "Audit Status", "text"],
    ["CSR Status", "CSR", "CSR Status", "CSR Received Status", "text"],
    ["CSR Max Cam", "CSR", "CSR Camera/VOIP", "Working Cam Count S1 | Working Cam Count S2 | Working Cam Count S3 | Working Cam Count S4", "num"],
    ["Rejected Param", "CSR", "Rejected Parameter", "Rejected Parameteres", "text"],
    ["Rejected Remarks", "CSR", "Rejected Remarks", "Remarks", "text"],
]
DEFAULT_MAPPING = pd.DataFrame(_ROWS, columns=MAP_COLS)

# CD column blank hona rule ke hisaab se theek hai: {CD field header: (CD status header, [keywords])}
RULE_BLANK = {"CSR Camera/VOIP": ("CSR Status", ["rejected", "not rec"])}

def tier(q):
    if pd.isna(q): return np.nan
    for lim, v in ((.95, 0), (.90, .05), (.85, .10), (.80, .15), (.70, .30)):
        if q >= lim: return v
    return 1.0 if q <= 0 else .50

# Rule re-check on day-1 rows. g(name) = numeric column, g(name, "text") = text column
def _txt(g, n): return g(n, "text").str.lower()
CHECKS = [
    ("Cam not Received", "Camera not Received in Data", lambda g: (g("Ideal Cam") - g("MAX Cam RCV as per Data")).clip(lower=0)),
    ("Not Rcv Deduction", "Not Received Camera Duduction", lambda g: (g("Ideal Cam") - g("MAX Cam RCV as per Data")).clip(lower=0) * g("Final Camera Rate")),
    ("Deduction %", "Deduction %", lambda g: g("Audit %").map(tier)),
    ("CSR Not Rcv", "CSR Not Recived", lambda g: _txt(g, "CSR Status").str.contains("not rec").astype(int) * 1000),
    ("CSR Rejected", "CSR Rejected Deduction", lambda g: _txt(g, "CSR Status").str.contains("rejected").astype(int) * 500),
]

_MAP_ALIASES = {
    "Field": ["field", "field name", "parameter", "name", "column name"],
    "Source": ["source", "source tab", "source sheet", "tab", "sheet", "source name"],
    "Center Data header": ["center data header", "centre data header", "center data column", "centre data column", "cd header", "cd column", "cd", "center data", "center header"],
    "Source header": ["source header", "source column", "source col", "src header", "src column", "source heading"],
    "Kind": ["kind", "type", "data type", "compare type"],
}

def canon_source(name):
    """'PM Compile' / 'pm' / 'Calling Sheet' -> 'PM' / 'Calling' (SOURCES ke key)."""
    n = nk(name)
    if not n: return ""
    for k in SOURCES:
        if nk(k) == n: return k
    for k, al in DEFAULT_TABS.items():
        if n in [nk(a) for a in al]: return k
    pool = {nk(a): k for k, al in DEFAULT_TABS.items() for a in al}
    mt = difflib.get_close_matches(n, list(pool), n=1, cutoff=0.75)
    return pool[mt[0]] if mt else str(name).strip()

def parse_mapping(raw):
    """Google Sheet ke Mapping tab (header=None frame) -> standard mapping DataFrame, ya None."""
    for i in range(min(15, len(raw))):
        hdr = [nk(v) if pd.notna(v) else "" for v in raw.iloc[i].tolist()]
        pos = {}
        for std, al in _MAP_ALIASES.items():
            for j, h in enumerate(hdr):
                if h in al and j not in pos.values(): pos[std] = j; break
        if {"Source", "Source header"} <= set(pos) and ("Center Data header" in pos or "Field" in pos):
            body = raw.iloc[i + 1:]
            m = pd.DataFrame({std: (body.iloc[:, pos[std]].values if std in pos else [""] * len(body)) for std in MAP_COLS})
            m = m.replace(r"^\s*$", np.nan, regex=True).dropna(how="all", subset=["Source", "Source header", "Center Data header", "Field"]).fillna("")
            m = m.astype(str).apply(lambda c: c.map(_c))
            if "Field" not in pos: m["Field"] = m["Center Data header"].where(m["Center Data header"] != "", m["Source header"])
            m.loc[m["Field"] == "", "Field"] = m["Center Data header"]
            return m.reset_index(drop=True)
    return None

def _c(s): return re.sub(r"\s+", " ", str(s).replace("\xa0", " ")).strip()
def nk(s): return _c(s).lower()
def split_alias(s):
    if s is None or (isinstance(s, float) and np.isnan(s)): return []
    return [a.strip() for a in re.split(r"[|;\n,]", str(s)) if a.strip()]

def num(v):
    if v is None or (isinstance(v, float) and np.isnan(v)): return np.nan
    if isinstance(v, (int, float, np.number)): return float(v)
    t = str(v).strip().replace(",", "")
    if t == "": return np.nan
    pct = t.endswith("%"); t = t.rstrip("%")
    try: x = float(t)
    except ValueError: return np.nan
    return x / 100 if pct else x

def txt(v):
    """Compare ke liye text: 5.0 -> '5', extra spaces hatao, lower."""
    if v is None or (isinstance(v, float) and np.isnan(v)): return ""
    if isinstance(v, (float, np.floating)) and float(v).is_integer(): v = int(v)
    return _c(v).lower()

def norm_tc(v):
    if isinstance(v, (float, np.floating)) and not np.isnan(v) and float(v).is_integer(): return str(int(v))
    return str(v).strip()

def join_uniq(s):
    out = []
    for v in s.dropna():
        v = str(v).strip()
        if v and v not in out: out.append(v)
    return " | ".join(out) if out else np.nan

def pick_tab(names, wanted):
    nn = {nk(n): n for n in names}
    for w in wanted:
        if nk(w) in nn: return nn[nk(w)]
    for w in wanted:
        m = difflib.get_close_matches(nk(w), list(nn), n=1, cutoff=0.6)
        if m: return nn[m[0]]
    return None

def frames_from_excel(xf, tabs=None):
    frames, warns = {}, []
    for k in SOURCES:
        t = pick_tab(xf.sheet_names, [(tabs or {}).get(k)] if (tabs or {}).get(k) else DEFAULT_TABS[k])
        if t is None: warns.append(f"{k}: tab nahi mila (available: {xf.sheet_names})"); continue
        frames[k] = xf.parse(t, header=None)
    return frames, warns

def prep(raw, tc_aliases):
    """Header row khud dhoondhta hai (TC column ke naam se), duplicate headers ko .1 .2 lagata hai."""
    al = {nk(a) for a in tc_aliases}; h = None
    for i in range(min(10, len(raw))):
        if any(nk(v) in al for v in raw.iloc[i].tolist() if pd.notna(v)): h = i; break
    if h is None: return None, None, f"TC column {tc_aliases} top 10 rows me nahi mila"
    heads, seen = [], {}
    for i, v in enumerate(raw.iloc[h].tolist()):
        name = f"Unnamed: {i}" if (pd.isna(v) or str(v).strip() == "") else _c(v)
        if name in seen: seen[name] += 1; name = f"{name}.{seen[name]}"
        else: seen[name] = 0
        heads.append(name)
    df = raw.iloc[h + 1:].reset_index(drop=True); df.columns = heads
    df = df.replace(r"^\s*$", np.nan, regex=True)
    lk = {}
    for c in df.columns: lk.setdefault(nk(c), c)
    tc = next(lk[nk(a)] for a in tc_aliases if nk(a) in lk)
    df = df[df[tc].notna()].copy(); df["__tc"] = df[tc].map(norm_tc)
    return df, lk, None

def field(df, lk, aliases, kind):
    cols = list(dict.fromkeys(lk[nk(a)] for a in aliases if nk(a) in lk))
    if not cols: return None
    uniq = not df["__tc"].duplicated().any()
    if kind in ("num", "pct"):
        v = pd.concat([df[c].map(num) for c in cols], axis=1).max(axis=1)
        s_ = pd.Series(v.values, index=df["__tc"].values)
        return s_ if uniq else s_.groupby(level=0, sort=False).max()
    v = df[cols].bfill(axis=1).iloc[:, 0]
    if uniq:
        return pd.Series(v.values, index=df["__tc"].values).map(lambda z: _c(int(z) if isinstance(z, (float, np.floating)) and float(z).is_integer() else z) if pd.notna(z) else np.nan)
    d = pd.DataFrame({"t": df["__tc"].values, "v": v.values}).dropna()
    d["v"] = d["v"].astype(str).str.strip(); d = d[d["v"] != ""].drop_duplicates()
    return d.groupby("t", sort=False)["v"].agg(" | ".join)

def gap(c, s, kind, tol):
    cb = txt(c) == ""; sb = txt(s) == ""
    if cb and sb: return "OK"
    if cb: return "MISSING IN CD"
    if sb: return "SOURCE BLANK"
    if kind in ("num", "pct"):
        a, b = num(c), num(s)
        if not (np.isnan(a) or np.isnan(b)): return "OK" if abs(a - b) < (tol if kind == "pct" else 1e-4) else "MISMATCH"
    return "OK" if txt(c) == txt(s) else "MISMATCH"

def run(frames, mapping, tol=0.006, auto=True):
    warns = []
    m = mapping.copy(); m.columns = [_c(c) for c in m.columns]
    for c in MAP_COLS:
        if c not in m.columns: m[c] = ""
    m = m.fillna("").astype(str)
    m["Source"] = m["Source"].map(canon_source)
    tcal = {k: ["TC Code"] for k in SOURCES}
    for _, r in m[m["Field"].str.lower() == "tc"].iterrows():
        if split_alias(r["Source header"]): tcal[r["Source"].strip()] = split_alias(r["Source header"])
    dfs = {}
    for k, raw in frames.items():
        df, lk, err = prep(raw, tcal.get(k, ["TC Code"]))
        if err: warns.append(f"{k}: {err}"); continue
        dfs[k] = (df, lk)
    if "Center Data" not in dfs: raise ValueError("Center Data load nahi hui: " + "; ".join(warns))
    cd_df, cd_lk = dfs["Center Data"]
    tcs = list(dict.fromkeys(cd_df["__tc"]))
    if auto:
        used = {nk(a) for _, r in m.iterrows() for a in split_alias(r["Center Data header"])}
        am = auto_map(dfs, used, tcs, tol, skip={nk(a) for a in tcal.get('Center Data', [])} | {nk(a) for _, r in m[m['Field'].str.lower() == 'day'].iterrows() for a in split_alias(r['Source header'])})
        if len(am): m = pd.concat([m, am.astype(str)], ignore_index=True)
        if "How" not in m.columns: m["How"] = ""
    wide = pd.DataFrame(index=tcs); long_rows = []; summary = []
    for _, r in m.iterrows():
        f, sk = r["Field"].strip(), r["Source"].strip()
        if f.lower() in ("tc", "day", "") or not sk: continue
        kind = (r["Kind"].strip().lower() or "text")
        cd_al, s_al = split_alias(r["Center Data header"]), split_alias(r["Source header"])
        cv = field(cd_df, cd_lk, cd_al, kind)
        if cv is None: warns.append(f"[{f}] Center Data me header nahi mila: {cd_al}"); continue
        if sk not in dfs: warns.append(f"[{f}] source '{sk}' load nahi hui"); continue
        sdf, slk = dfs[sk]
        sv = field(sdf, slk, s_al, kind)
        if sv is None:
            warns.append(f"[{f}] {sk} me header nahi mila: {s_al}. Available: {list(slk.values())[:40]}"); continue
        cv = cv.reindex(tcs); sv = sv.reindex(tcs); ins = set(sdf["__tc"])
        gp = [gap(cv[t], sv[t], kind, tol) if t in ins else "NOT IN SOURCE" for t in tcs]
        rb = next((v for k_, v in RULE_BLANK.items() if nk(k_) == nk(cd_al[0])), None) if cd_al else None
        if rb:
            st_ = field(cd_df, cd_lk, [rb[0]], "text")
            if st_ is not None:
                st_ = st_.reindex(tcs).fillna("").str.lower()
                gp = ["BLANK BY RULE" if g_ == "MISSING IN CD" and any(w in st_[t] for w in rb[1]) else g_ for g_, t in zip(gp, tcs)]
        gp = pd.Series(gp, index=tcs)
        wide[f"{f} | CD"] = cv.values; wide[f"{f} | SRC"] = sv.values; wide[f"{f} | GAP"] = gp.values
        how = r.get("How", "") if "How" in r.index else ""
        how = "" if (not isinstance(how, str) and pd.isna(how)) else str(how)
        summary.append(dict(Field=f, Source=f"{sk}: {s_al[0] if s_al else ''}" + (f" [{how}]" if how and how != "nan" else ""), **gp.value_counts().to_dict()))
        for t in gp[~gp.isin(["OK", "BLANK BY RULE"])].index:
            long_rows.append(dict(TC=t, Field=f, CenterData=cv[t], Source=sv[t], Gap=gp[t]))
    wide.index.name = "TC Code"
    summ = pd.DataFrame(summary).fillna(0)
    for k in ["OK", "MISMATCH", "MISSING IN CD", "SOURCE BLANK", "NOT IN SOURCE", "BLANK BY RULE"]:
        if k not in summ: summ[k] = 0
    summ = summ[["Field", "Source", "OK", "MISMATCH", "MISSING IN CD", "SOURCE BLANK", "NOT IN SOURCE", "BLANK BY RULE"]]
    num_cols = [c for c in summ.columns if c not in ("Field", "Source")]
    summ[num_cols] = summ[num_cols].astype(int)
    gaps = pd.DataFrame(long_rows, columns=["TC", "Field", "CenterData", "Source", "Gap"])
    # ---- rule re-check (day 1) ----
    day_al = []
    for _, r in m[m["Field"].str.lower() == "day"].iterrows(): day_al += split_alias(r["Source header"]) + split_alias(r["Center Data header"])
    day_al += ["Exam Day Number", "Day", "Exam Day"]
    dcol = next((cd_lk[nk(a)] for a in day_al if nk(a) in cd_lk), None)
    if dcol is None: warns.append("Day column nahi mila - rule check sab rows pe chala")
    d1 = cd_df[cd_df[dcol].map(num) == 1] if dcol else cd_df
    d1 = d1.drop_duplicates("__tc")
    def g(name, kind="num"):
        col = cd_lk.get(nk(name))
        if col is None: raise KeyError(name)
        s = d1[col]; s = s.map(num) if kind == "num" else s.fillna("").astype(str)
        s.index = d1["__tc"].values
        return s
    rr = []
    for label, col, fn in CHECKS:
        try: exp = pd.to_numeric(fn(g), errors="coerce").fillna(0); act = g(col).fillna(0)
        except KeyError as e: warns.append(f"Rule '{label}' skip: column nahi mila {e}"); continue
        for t in act.index[(act - exp).abs() > 0.01]:
            rr.append(dict(TC=t, Rule=label, Column=col, CenterData=act[t], Expected=exp[t], Diff=round(act[t] - exp[t], 2)))
    rules = pd.DataFrame(rr, columns=["TC", "Rule", "Column", "CenterData", "Expected", "Diff"])
    return dict(summary=summ, gaps=gaps, rules=rules, compare=wide, warns=warns, n=len(tcs))

def auto_map(dfs, used, tcs, tol=0.006, min_match=0.95, min_vals=30, skip=()):
    """Mapping me jo Center Data columns nahi hain, unke liye source columns auto dhoondhta hai
    (values >=95% match, ya same header naam + >=50% match)."""
    cd_df, cd_lk = dfs["Center Data"]; cache = {}
    def ser(key, col, kind):
        ck = (key, col, kind)
        if ck not in cache:
            df, lk = dfs[key]; cache[ck] = field(df, lk, [col], kind)
        return cache[ck]
    rows = []
    for col in list(cd_lk.values()):
        if nk(col) in used or nk(col) in skip or col == "__tc" or col.startswith("Unnamed"): continue
        raw = cd_df[col]; nn_ = raw.notna().sum()
        if nn_ < min_vals: continue
        kind = "pct" if "%" in col else ("num" if raw.map(num).notna().sum() >= 0.9 * nn_ else "text")
        a = ser("Center Data", col, kind)
        if a is None: continue
        a = a.reindex(tcs)
        if a.nunique() < 2: continue
        best = None
        for key, (sdf, slk) in dfs.items():
            if key == "Center Data": continue
            for sc in slk.values():
                if sc == "__tc": continue
                b = ser(key, sc, kind)
                if b is None: continue
                b = b.reindex(tcs)
                if b.notna().sum() < min_vals or b.nunique() < 2: continue
                if kind in ("num", "pct"):
                    av, bv = a.values.astype(float), b.values.astype(float)
                    eq = float(np.mean((np.isnan(av) & np.isnan(bv)) | (np.abs(av - bv) < (tol if kind == "pct" else 1e-4))))
                else:
                    eq = float(np.mean((a.map(txt) == b.map(txt)).values))
                ns = difflib.SequenceMatcher(None, nk(col), nk(sc)).ratio()
                if (eq >= min_match and ns >= 0.5) or (ns == 1 and eq >= 0.5):
                    sc_ = eq + 0.01 * ns
                    if best is None or sc_ > best[0]: best = (sc_, key, sc, eq)
        if best: rows.append([col, best[1], col, best[2], kind, f"auto {best[3]*100:.0f}% match"])
    return pd.DataFrame(rows, columns=MAP_COLS + ["How"])

def to_excel_bytes(res):
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as w:
        res["summary"].to_excel(w, sheet_name="Summary", index=False)
        res["gaps"].to_excel(w, sheet_name="Gaps_Only", index=False)
        res["rules"].to_excel(w, sheet_name="Rule_Errors", index=False)
        res["compare"].to_excel(w, sheet_name="Compare")
        red, amb, grn = (PatternFill("solid", start_color=h) for h in ("F8CBAD", "FFE699", "C6E0B4"))
        for row in w.sheets["Compare"].iter_rows(min_row=2):
            for cell in row:
                if cell.value in ("OK", "BLANK BY RULE"): cell.fill = grn
                elif cell.value in ("MISMATCH", "MISSING IN CD"): cell.fill = red
                elif cell.value in ("SOURCE BLANK", "NOT IN SOURCE"): cell.fill = amb
        for sh in w.sheets.values():
            sh.freeze_panes = "A2"; sh.auto_filter.ref = sh.dimensions
            for col in sh.columns: sh.column_dimensions[col[0].column_letter].width = 16
            for cell in sh[1]: cell.font = Font(bold=True)
    return bio.getvalue()
