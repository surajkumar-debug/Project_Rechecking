import datetime as dt, io, re
import numpy as np, pandas as pd, streamlit as st
import core, gsheet

st.set_page_config(page_title="Center Data Recheck", layout="wide")
st.markdown("""
<style>
header[data-testid="stHeader"] {background: transparent; height: 2.5rem;}
div[data-testid="stMainBlockContainer"], .block-container {padding-top: 1rem !important; padding-bottom: 7rem !important;}
html, body, [data-testid="stAppViewContainer"], section.main {overflow-anchor: none;}
h1.portal-title {font-size: 1.9rem; font-weight: 700; margin: 0 0 .6rem 0; padding: 0;}
</style>
<h1 class="portal-title">Center Data Recheck Portal</h1>
""", unsafe_allow_html=True)

def secret(k, d=""):
    try: return st.secrets.get(k, d)
    except Exception: return d

@st.cache_data(ttl=300, show_spinner=False)
def tabs_of(sid, key): return gsheet.list_tabs(sid, key)

@st.cache_data(ttl=300, show_spinner=False)
def tab_df(sid, tab, key): return gsheet.get_tab(sid, tab, key)

@st.cache_resource(show_spinner=False)
def xl(b): return pd.ExcelFile(io.BytesIO(b))

def read_mapping(sid, names, key):
    """Reads the mapping from the 'Mapping' tab of the control sheet. The header row may be anywhere in the first 15 rows
    and column names are matched flexibly. Returns (mapping_df, status_text, raw_df)."""
    t = core.pick_tab(names, ["Mapping"])
    if not t: return core.DEFAULT_MAPPING.copy(), "No 'Mapping' tab was found in the sheet; the default mapping is being used", None
    raw = tab_df(sid, t, key)
    if raw.empty: return core.DEFAULT_MAPPING.copy(), f"The '{t}' tab is empty; the default mapping is being used", raw
    m = core.parse_mapping(raw)
    if m is None: return core.DEFAULT_MAPPING.copy(), f"The headers in the '{t}' tab were not recognised; the default mapping is being used", raw
    return m, f"Mapping: loaded from the '{t}' tab of the Google Sheet ({len(m)} rows)", raw

SEC_URL, SEC_KEY = str(secret("control_url")).strip(), str(secret("api_key")).strip()
HAVE_SECRETS = bool(SEC_URL and SEC_KEY)

with st.sidebar:
    mode = "Control Google Sheet" if HAVE_SECRETS else st.radio("Data source", ["Control Google Sheet", "Excel upload"])
    auto = st.checkbox("Auto-match extra columns", True, help="Automatically detects columns that are not defined in the Mapping, using header names and value matching.")
    tol = st.number_input("% tolerance", 0.0, 0.1, 0.006, 0.001, format="%.3f")
    if st.button("Refresh Google Sheet data"):
        st.cache_data.clear(); st.session_state.pop("res", None); st.rerun()

frames, load_warns, mapping, label = {}, [], None, "project"

if mode == "Control Google Sheet":
    if HAVE_SECRETS: url, key = SEC_URL, SEC_KEY            # Credentials come from Secrets; nothing is shown in the UI
    else:
        url = st.text_input("Control sheet link or Apps Script Web App URL (/exec)")
        key = st.text_input("Google API key or Apps Script token", type="password")
        if not url or not key: st.info("Enter the Web App URL and token (or the Sheet link and API key), or configure them in Streamlit Secrets."); st.stop()
    key = key.strip(); url = url.strip()
    if key.lower().startswith("http"):
        st.error("The second field expects the access token, not a link. Use the value of `const TOKEN` from your Apps Script."); st.stop()
    sid = gsheet.sheet_id(url)
    try: names = tabs_of(sid, key)
    except Exception as e: st.error(str(e)); st.stop()
    mapping, status, mraw = read_mapping(sid, names, key)
    if status.startswith("Mapping:"): st.caption("✅ " + status)
    else:
        st.warning(status + ". The Mapping tab requires these columns: Field, Source, Center Data header, Source header, Kind (optional).")
        if mraw is not None and not mraw.empty: st.caption("Content found in the Mapping tab (first rows):"); st.dataframe(mraw.head(8), width="stretch")
    chosen = {}
    with st.expander("Advanced settings: tab selection and mapping", expanded=False):
        st.caption("Tabs: " + ", ".join(names))
        cols = st.columns(len(core.SOURCES))
        for c, k in zip(cols, core.SOURCES):
            d = core.pick_tab(names, core.DEFAULT_TABS[k])
            chosen[k] = c.selectbox(k, names, index=names.index(d) if d in names else 0, key=f"tab_{k}")
        st.caption("The mapping is loaded from the Mapping tab of the Google Sheet. Make permanent changes there, then use Refresh in the sidebar. Edits made here apply to the current run only.")
        mapping = st.data_editor(mapping, num_rows="dynamic", width="stretch", key=f"map_{sid[:6]}")
    label = sid[:6]
    def load():
        fr = {k: tab_df(sid, t, key) for k, t in chosen.items()}
        w = []
        for e in core.extra_sources(mapping):          # any additional tab named in the Mapping
            t = core.pick_tab(names, [e], 0.85)
            if t: fr[e] = tab_df(sid, t, key)
            else: w.append(f"{e}: tab not found in the Google Sheet")
        return fr, w + [f"{k}: tab is empty" for k, v in fr.items() if v.empty]
else:
    up = st.sidebar.file_uploader("Project Excel (.xlsx)", type=["xlsx"])
    mup = st.sidebar.file_uploader("Mapping (.csv) - optional", type=["csv"])
    if not up: st.info("Upload the project Excel file from the sidebar."); st.stop()
    mapping = pd.read_csv(mup) if mup else core.DEFAULT_MAPPING.copy()
    label = up.name.rsplit(".", 1)[0]
    load = lambda: core.frames_from_excel(xl(up.getvalue()), extras=core.extra_sources(mapping))
    with st.expander("Advanced settings: mapping", expanded=False):
        mapping = st.data_editor(mapping, num_rows="dynamic", width="stretch", key=f"map_{label}")

if st.button("Run Recheck", type="primary"):
    with st.spinner("Reading tabs and comparing data (auto-match may take around 30 seconds)..."):
        try:
            frames, load_warns = load()
            res = core.run(frames, mapping, tol, auto)
            res["warns"] = load_warns + res["warns"]; st.session_state["res"] = res; st.session_state["label"] = label
        except Exception as e: st.error(f"Recheck failed: {e}")

res = st.session_state.get("res")
if res:
    for w in res["warns"]: st.warning(w)
    s = res["summary"]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Centers", res["n"]); m2.metric("Gap cells", len(res["gaps"])); m3.metric("Rule errors", len(res["rules"]))
    m4.metric("Clean fields", int(((s["OK"] + s["BLANK BY RULE"]) == res["n"]).sum()))
    st.download_button("Download Excel report", core.to_excel_bytes(res), f"Recheck_{st.session_state['label']}.xlsx", type="primary")
    def tc_match(series, query):
        """Keep rows whose TC contains any of the comma/space separated search terms."""
        terms = [t.strip().lower() for t in re.split(r"[,\s]+", query) if t.strip()]
        if not terms: return pd.Series(True, index=series.index)
        v = series.astype(str).str.lower()
        return v.apply(lambda x: any(t in x for t in terms))

    def csv_btn(df, name, key):
        st.download_button("Download filtered data (CSV)", df.to_csv(index=False), f"{name}.csv", key=key)

    t1, t2, t3, t5, t4 = st.tabs(["Summary", "Gaps (center-wise)", "Rule errors", "Foldering (No)", "Full compare"])

    with t1:
        issue_cols = ["MISMATCH", "MISSING IN CD", "SOURCE BLANK"]
        only_issues = st.checkbox("Show only fields with issues", False, key="f_sum_issue")
        sv = s[(s[issue_cols].sum(axis=1) > 0)] if only_issues else s
        st.dataframe(sv, width="stretch", hide_index=True, height=35 * (len(sv) + 1) + 3)
        if (s[issue_cols].sum(axis=1) > 0).any():
            st.caption("Issues by field")
            st.bar_chart(s.set_index("Field")[issue_cols], height=300)

    with t2:
        g = res["gaps"]
        c1, c2, c3 = st.columns([2, 2, 2])
        f_field = c1.multiselect("Field", sorted(g["Field"].unique()), key="f_g_field")
        f_gap = c2.multiselect("Gap type", sorted(g["Gap"].unique()), key="f_g_gap")
        f_tc = c3.text_input("Search TC code (comma separated)", key="f_g_tc")
        view = st.radio("View", ["Detail", "Count per center", "Count per field"], horizontal=True, key="f_g_view")
        gv = g
        if f_field: gv = gv[gv["Field"].isin(f_field)]
        if f_gap: gv = gv[gv["Gap"].isin(f_gap)]
        gv = gv[tc_match(gv["TC"], f_tc)]
        st.caption(f"{len(gv):,} of {len(g):,} rows | {gv['TC'].nunique():,} centers")
        if view == "Count per center":
            gv = gv.groupby("TC").agg(Issues=("Field", "size"), Fields=("Field", lambda x: ", ".join(sorted(set(x))))).reset_index().sort_values("Issues", ascending=False)
        elif view == "Count per field":
            gv = gv.pivot_table(index="Field", columns="Gap", values="TC", aggfunc="count", fill_value=0).reset_index()
        st.dataframe(gv, width="stretch", hide_index=True, height=520)
        csv_btn(gv, "gaps_filtered", "dl_gaps")

    with t3:
        r = res["rules"]
        if r.empty: st.info("No rule errors were found.")
        else:
            c1, c2 = st.columns([2, 2])
            f_rule = c1.multiselect("Rule", sorted(r["Rule"].unique()), key="f_r_rule")
            f_tc = c2.text_input("Search TC code (comma separated)", key="f_r_tc")
            rv = r[r["Rule"].isin(f_rule)] if f_rule else r
            rv = rv[tc_match(rv["TC"], f_tc)]
            st.caption(f"{len(rv):,} of {len(r):,} rows | {rv['TC'].nunique():,} centers")
            st.dataframe(rv, width="stretch", hide_index=True, height=520)
            csv_btn(rv, "rule_errors_filtered", "dl_rules")

    with t5:
        fd = res.get("foldering")
        if fd is None or fd.empty: st.info("No foldering records were found, or the foldering columns could not be detected. Please check the warnings above.")
        else:
            st.caption(f"{len(fd)} centers listed (Folder Status 'No' in Audit Final, or a foldering entry present in Center Data). Check = result of the rule: the center code and deduction must appear in Center Data only where the status is 'No'.")
            c1, c2 = st.columns([2, 2])
            f_gap = c1.multiselect("Check result", sorted(fd["Check"].unique()), key="f_f_gap")
            f_tc = c2.text_input("Search TC code (comma separated)", key="f_f_tc")
            fv = fd[fd["Check"].isin(f_gap)] if f_gap else fd
            fv = fv[tc_match(fv["TC"], f_tc)]
            st.dataframe(fv, width="stretch", hide_index=True, height=520)
            csv_btn(fv, "foldering_filtered", "dl_fold")

    with t4:
        cmp_df = res["compare"]
        fields = sorted({c.rsplit(" | ", 1)[0] for c in cmp_df.columns if c.endswith((" | CD", " | SRC", " | GAP"))})
        issue_status = set(res["gaps"]["Gap"].unique())
        c1, c2, c3 = st.columns([2, 2, 2])
        f_fields = c1.multiselect("Fields to display", fields, key="f_c_fields")
        f_tc = c2.text_input("Search TC code (comma separated)", key="f_c_tc")
        only_bad = c3.checkbox("Only centers with issues", False, key="f_c_bad")
        shown = f_fields or fields
        cols = [f"{f} | {k}" for f in shown for k in ("CD", "SRC", "GAP") if f"{f} | {k}" in cmp_df.columns]
        cv = cmp_df[cols]
        gap_cols = [c for c in cols if c.endswith(" | GAP")]
        if only_bad and gap_cols: cv = cv[cv[gap_cols].isin(issue_status).any(axis=1)]
        cv = cv[tc_match(cv.index.to_series(), f_tc)]
        st.caption(f"{len(cv):,} of {len(cmp_df):,} centers")
        st.dataframe(cv, width="stretch", height=520)

st.markdown("<div style='height:80px'></div>", unsafe_allow_html=True)
