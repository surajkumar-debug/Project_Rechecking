import io, re, difflib
import numpy as np, pandas as pd
from openpyxl.styles import PatternFill, Font

DEFAULT_CFG = {
    "PROJECT": "",
    "CD": {"sheet": "Center Data", "header": 0, "tc": "TC Code", "day_col": "Exam Day Number", "project": "Project Code"},
    "SRC": {
        "PM": {"sheet": "PM Compile", "header": 0, "tc": "TC Code", "project": "Project Code"},
        "Calling": {"sheet": "Calling", "header": 0, "tc": "TC Code", "project": "Project"},
        "Audit": {"sheet": "Audit", "header": 1, "tc": "Venue Code", "project": "Project Code", "rename": {"Unnamed: 38": "As per Audit Max"}},
        "AuditFinal": {"sheet": "Audit Final", "header": 0, "tc": "TC_CODE", "project": "Project Code"},
    },
    "CSR": {"sheet": "CSR", "header": 0, "tc": "TC Code", "project": "Project code", "day": "Exam Day",
            "status": "CSR Received Status",
            "cams": ["Working Cam Count S1", "Working Cam Count S2", "Working Cam Count S3", "Working Cam Count S4"],
            "param": "Rejected Parameteres", "remarks": "Remarks"},
    "MAP": [
        ["Actual Cam", "Actual Camera", "PM", "Final Cam Count", "num"],
        ["QC Cam", "QC Camera", "PM", "QC Cam Count", "num"],
        ["Monitor", "Monitor Count", "PM", "Final Monitor Count", "num"],
        ["UPS", "UPS Count", "PM", "Final UPS Count", "num"],
        ["HMS %", "HMS%", "PM", "HMS Avg Hit Per%", "pct"],
        ["DVR Installed", "Total DVR Count", "PM", "HMS DVR Installed Count", "num"],
        ["DVR Live", "Online DVR Count", "PM", "HMS No of DVR Hit", "num"],
        ["VOIP Status", "VOIP Status", "PM", "VOIP Final Installation Status", "text"],
        ["Common Center", "Common Center", "PM", "Common Centers", "text"],
        ["CCTV Installed", "CCTV Installation Status", "Calling", "CCTV Setup Installed", "text"],
        ["ISP", "Internet Installation Status", "Calling", "Internet Installation", "text"],
        ["Monitor Status", "Monitor Installation Status", "Calling", "Monitor Installed", "text"],
        ["UPS Status", "UPS.2", "Calling", "UPS Installed", "text"],
        ["Calling Status", "Calling Status", "Calling", "calling status", "text"],
        ["Remarks", "Remarks", "Calling", "remarks", "text"],
        ["Max Cam (Audit)", "MAX Cam RCV as per Data", "Audit", "As per Audit Max", "num"],
        ["Data %", "Data %", "Audit", "Data %", "pct"],
        ["Audit %", "Audit %", "Audit", "Audit %", "pct"],
        ["Audit Status", "Audit Status", "Audit", "Audit Status", "text"],
        ["CSR Status", "CSR Status", "CSR", "status", "text"],
        ["CSR Max Cam", "CSR Camera/VOIP", "CSR", "maxcam", "num"],
        ["Rejected Param", "Rejected Parameter", "CSR", "param", "text"],
        ["Rejected Remarks", "Rejected Remarks", "CSR", "remarks", "text"],
    ],
    "PCT_TOL": 0.006,
}

def _c(s): return re.sub(r"\s+", " ", str(s).replace("\xa0", " ")).strip()

def clean_cols(df):
    df.columns = [str(c) if str(c).startswith("Unnamed") else _c(c) for c in df.columns]
    return df

def load_sheet(x, cfg, project):
    df = x.parse(cfg["sheet"], header=int(cfg.get("header", 0)))
    df = clean_cols(df.rename(columns=cfg.get("rename") or {}))
    p = cfg.get("project")
    if project and p in df.columns: df = df[df[p].astype(str).str.strip() == project.strip()]
    return df

def tier(q):
    if pd.isna(q): return np.nan
    for lim, v in ((.95, 0), (.90, .05), (.85, .10), (.80, .15), (.70, .30)):
        if q >= lim: return v
    return 1.0 if q <= 0 else .50

# Rule re-check. (label, Center Data column, function(day1_df) -> expected)
CHECKS = [
    ("Cam not Received", "Camera not Received in Data", lambda d: (d["Ideal Cam"] - d["MAX Cam RCV as per Data"]).clip(lower=0)),
    ("Not Rcv Deduction", "Not Received Camera Duduction", lambda d: (d["Ideal Cam"] - d["MAX Cam RCV as per Data"]).clip(lower=0) * d["Final Camera Rate"]),
    ("Deduction %", "Deduction %", lambda d: d["Audit %"].map(tier)),
    ("CSR Not Rcv", "CSR Not Recived", lambda d: np.where(d["CSR Status"].fillna("").str.lower().str.contains("not rec"), 1000, 0)),
    ("CSR Rejected", "CSR Rejected Deduction", lambda d: np.where(d["CSR Status"].fillna("").str.lower().str.contains("rejected"), 500, 0)),
]

def _join_uniq(s):
    out = []
    for v in s.dropna():
        v = str(v).strip()
        if v and v not in out: out.append(v)
    return " | ".join(out) if out else np.nan

def build_frames(x, cfg):
    project = cfg.get("PROJECT", "")
    cdc = cfg["CD"]
    cd_all = load_sheet(x, cdc, project)
    cd1 = cd_all.drop_duplicates(cdc["tc"]).set_index(cdc["tc"])
    tcs = list(cd1.index)
    src = {}; warns = []
    for k, s in cfg["SRC"].items():
        if not s.get("sheet") or s["sheet"] not in x.sheet_names: continue
        d = load_sheet(x, s, project)
        if s["tc"] not in d.columns:
            warns.append(f"{k}: TC column {s['tc']!r} nahi mila. Available: {list(d.columns)[:30]}"); continue
        src[k] = d[~d[s["tc"]].duplicated()].set_index(s["tc"])
    cs = cfg.get("CSR")
    if cs and cs.get("sheet") in x.sheet_names:
        c = load_sheet(x, cs, project)
        miss = [m for m in (cs["tc"], cs["status"], cs["param"], cs["remarks"]) if m not in c.columns]
        if miss:
            warns.append(f"CSR sheet me ye columns nahi mile: {miss}. Available: {list(c.columns)[:40]}")
        else:
            c = c[c[cs["tc"]].isin(tcs)].copy()
            if cs.get("day") in c.columns:
                c["_d"] = c[cs["day"]].astype(str).str.extract(r"(\d+)")[0].astype(float)
                c = c.sort_values([cs["tc"], "_d"])
            cams = [m for m in cs["cams"] if m in c.columns]
            c["_w"] = c[cams].max(axis=1) if cams else np.nan
            g = c.groupby(cs["tc"])
            src["CSR"] = pd.DataFrame({"status": g[cs["status"]].apply(_join_uniq), "maxcam": g["_w"].max(),
                                       "param": g[cs["param"]].apply(_join_uniq), "remarks": g[cs["remarks"]].apply(_join_uniq)})
    return cd_all, cd1, tcs, src, warns

def gap(c, s, kind, tol):
    cb = pd.isna(c) or str(c).strip() == ""; sb = pd.isna(s) or str(s).strip() == ""
    if cb and sb: return "OK"
    if cb: return "MISSING IN CD"
    if sb: return "SOURCE BLANK"
    if kind in ("num", "pct"):
        try: return "OK" if abs(float(c) - float(s)) < (tol if kind == "pct" else 1e-4) else "MISMATCH"
        except Exception: return "MISMATCH"
    return "OK" if str(c).strip().lower() == str(s).strip().lower() else "MISMATCH"

def run(x, cfg):
    cd_all, cd1, tcs, src, bw = build_frames(x, cfg)
    tol = cfg.get("PCT_TOL", 0.006)
    wide = pd.DataFrame(index=tcs); long_rows = []; summary = []; warns = list(bw)
    for label, cdcol, sk, scol, kind in cfg["MAP"]:
        cdcol, scol = _c(cdcol), _c(scol) if sk != "CSR" else scol
        if cdcol not in cd1.columns: warns.append(f"Center Data me column nahi mila: {cdcol!r}"); continue
        if sk not in src: warns.append(f"Source sheet '{sk}' load nahi hui"); continue
        if scol not in src[sk].columns: warns.append(f"{sk} me column nahi mila: {scol!r}"); continue
        cv = cd1[cdcol]; sv = src[sk][scol].reindex(tcs); ins = pd.Index(tcs).isin(src[sk].index)
        gp = pd.Series([gap(cv[t], sv[t], kind, tol) if ins[i] else "NOT IN SOURCE" for i, t in enumerate(tcs)], index=tcs)
        wide[f"{label} | CD"] = cv.values; wide[f"{label} | SRC"] = sv.values; wide[f"{label} | GAP"] = gp.values
        summary.append(dict(Field=label, Source=f"{sk}: {scol}", **gp.value_counts().to_dict()))
        for t in gp[gp != "OK"].index:
            long_rows.append(dict(TC=t, Field=label, CenterData=cv[t], Source=sv[t], Gap=gp[t]))
    if "Center Name" in cd1.columns: wide.insert(0, "Center Name", cd1["Center Name"].values)
    wide.index.name = "TC Code"
    summ = pd.DataFrame(summary).fillna(0)
    for k in ["OK", "MISMATCH", "MISSING IN CD", "SOURCE BLANK", "NOT IN SOURCE"]:
        if k not in summ: summ[k] = 0
    summ = summ[["Field", "Source", "OK", "MISMATCH", "MISSING IN CD", "SOURCE BLANK", "NOT IN SOURCE"]] if len(summ) else summ
    gaps = pd.DataFrame(long_rows, columns=["TC", "Field", "CenterData", "Source", "Gap"])
    # rule re-check
    dc = cfg["CD"].get("day_col")
    day1 = cd_all[cd_all[dc] == 1].set_index(cfg["CD"]["tc"]) if dc in cd_all.columns else cd1
    rr = []
    for label, col, fn in CHECKS:
        if col not in day1.columns: continue
        try: exp = pd.Series(fn(day1), index=day1.index)
        except KeyError: continue
        act = pd.to_numeric(day1[col], errors="coerce").fillna(0); exp = pd.to_numeric(exp, errors="coerce").fillna(0)
        for t in day1.index[(act - exp).abs() > 0.01]:
            rr.append(dict(TC=t, Rule=label, Column=col, CenterData=act[t], Expected=exp[t], Diff=round(act[t] - exp[t], 2)))
    rules = pd.DataFrame(rr, columns=["TC", "Rule", "Column", "CenterData", "Expected", "Diff"])
    return dict(summary=summ, gaps=gaps, rules=rules, compare=wide, warns=warns, n=len(tcs))

def auto_detect(x, cfg, min_match=0.95, min_vals=30):
    """Center Data ke har column ko har source column se value-match karke best mapping suggest karta hai."""
    cd_all, cd1, tcs, src, _ = build_frames(x, cfg)
    rows = []
    def norm(s):
        n = pd.to_numeric(s, errors="coerce")
        return n.round(3) if n.notna().sum() >= s.notna().sum() * 0.9 and s.notna().any() else s.astype(str).str.strip().str.lower().where(s.notna())
    for col in cd1.columns:
        a = cd1[col]
        if a.notna().sum() < min_vals or a.nunique() < 2: continue
        na = norm(a); best = None
        for sk, sdf in src.items():
            for sc in sdf.columns:
                b = sdf[sc].reindex(tcs)
                if b.notna().sum() < min_vals: continue
                nb = norm(b)
                eq = ((na == nb) | (na.isna() & nb.isna())).mean()
                sc_ = eq + 0.01 * difflib.SequenceMatcher(None, _c(col).lower(), _c(sc).lower()).ratio()
                if eq >= min_match and (best is None or sc_ > best[3]): best = (eq, sk, sc, sc_)
        if best:
            kind = "pct" if "%" in col else ("num" if pd.to_numeric(a, errors="coerce").notna().sum() >= a.notna().sum() * 0.9 else "text")
            rows.append([col, col, best[1], best[2], kind, round(best[0] * 100, 1)])
    return pd.DataFrame(rows, columns=["Label", "CD column", "Source", "Source column", "Kind", "Match %"])

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
                if cell.value == "OK": cell.fill = grn
                elif cell.value in ("MISMATCH", "MISSING IN CD"): cell.fill = red
                elif cell.value in ("SOURCE BLANK", "NOT IN SOURCE"): cell.fill = amb
        for sh in w.sheets.values():
            sh.freeze_panes = "A2"; sh.auto_filter.ref = sh.dimensions
            for col in sh.columns: sh.column_dimensions[col[0].column_letter].width = 16
            for cell in sh[1]: cell.font = Font(bold=True)
    return bio.getvalue()
    
