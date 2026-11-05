import datetime as dt, io
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
        return fr, [f"{k}: tab is empty" for k, v in fr.items() if v.empty]
else:
    up = st.sidebar.file_uploader("Project Excel (.xlsx)", type=["xlsx"])
    mup = st.sidebar.file_uploader("Mapping (.csv) - optional", type=["csv"])
    if not up: st.info("Upload the project Excel file from the sidebar."); st.stop()
    mapping = pd.read_csv(mup) if mup else core.DEFAULT_MAPPING.copy()
    label = up.name.rsplit(".", 1)[0]
    load = lambda: core.frames_from_excel(xl(up.getvalue()))
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
    t1, t2, t3, t5, t4 = st.tabs(["Summary", "Gaps (center-wise)", "Rule errors", "Foldering (No)", "Full compare"])
    with t1: st.dataframe(s, width="stretch", hide_index=True, height=35 * (len(s) + 1) + 3)
    with t2:
        f = st.multiselect("Field filter", sorted(res["gaps"]["Field"].unique()))
        st.dataframe(res["gaps"][res["gaps"]["Field"].isin(f)] if f else res["gaps"], width="stretch", hide_index=True, height=520)
    with t3: st.dataframe(res["rules"], width="stretch", hide_index=True, height=520)
    with t5:
        fd = res.get("foldering")
        if fd is None or fd.empty: st.info("No centers with Folder Status 'No' were found, or the foldering column could not be detected. Please check the warnings above.")
        else: st.caption(f"{len(fd)} centers with Folder Status 'No' (in Audit Final or Center Data). The Gap column shows whether both sources agree."); st.dataframe(fd, width="stretch", hide_index=True, height=520)
    with t4: st.dataframe(res["compare"], width="stretch", height=520)

st.markdown("<div style='height:80px'></div>", unsafe_allow_html=True)
