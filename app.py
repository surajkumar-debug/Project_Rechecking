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
    """Control sheet ke 'Mapping' tab se mapping padhta hai. Header row pehli 15 rows me kahin bhi ho sakti hai,
    column ke naam flexible hain. Returns (mapping_df, status_text, raw_df)."""
    t = core.pick_tab(names, ["Mapping"])
    if not t: return core.DEFAULT_MAPPING.copy(), "Sheet me 'Mapping' naam ka tab nahi mila - default mapping use ho rahi hai", None
    raw = tab_df(sid, t, key)
    if raw.empty: return core.DEFAULT_MAPPING.copy(), f"'{t}' tab khaali hai - default mapping use ho rahi hai", raw
    m = core.parse_mapping(raw)
    if m is None: return core.DEFAULT_MAPPING.copy(), f"'{t}' tab ke headers samajh nahi aaye - default mapping use ho rahi hai", raw
    return m, f"Mapping: Google Sheet ke '{t}' tab se ({len(m)} rows)", raw

SEC_URL, SEC_KEY = str(secret("control_url")).strip(), str(secret("api_key")).strip()
HAVE_SECRETS = bool(SEC_URL and SEC_KEY)

with st.sidebar:
    mode = "Control Google Sheet" if HAVE_SECRETS else st.radio("Data kahan se?", ["Control Google Sheet", "Excel upload"])
    auto = st.checkbox("Auto-match extra columns", True, help="Mapping me jo columns nahi hain unhe value / naam match se khud dhoondta hai")
    tol = st.number_input("% tolerance", 0.0, 0.1, 0.006, 0.001, format="%.3f")
    if st.button("Google Sheet data refresh"):
        st.cache_data.clear(); st.session_state.pop("res", None); st.rerun()

frames, load_warns, mapping, label = {}, [], None, "project"

if mode == "Control Google Sheet":
    if HAVE_SECRETS: url, key = SEC_URL, SEC_KEY            # Secrets se - UI me kuch nahi dikhega
    else:
        url = st.text_input("Control sheet link ya Apps Script Web App URL (/exec)")
        key = st.text_input("Google API key ya Apps Script TOKEN", type="password")
        if not url or not key: st.info("Web App URL + TOKEN (ya Sheet link + API key) daalo, ya Streamlit Secrets me daal do."); st.stop()
    key = key.strip(); url = url.strip()
    if key.lower().startswith("http"):
        st.error("TOKEN wali jagah link nahi, sirf TOKEN chahiye (Apps Script me `const TOKEN = '...'` wala text)."); st.stop()
    sid = gsheet.sheet_id(url)
    try: names = tabs_of(sid, key)
    except Exception as e: st.error(str(e)); st.stop()
    mapping, status, mraw = read_mapping(sid, names, key)
    if status.startswith("Mapping:"): st.caption("✅ " + status)
    else:
        st.warning(status + ". Mapping tab me ye columns chahiye: Field, Source, Center Data header, Source header, Kind (Kind optional).")
        if mraw is not None and not mraw.empty: st.caption("Mapping tab me jo mila (pehli rows):"); st.dataframe(mraw.head(8), width="stretch")
    chosen = {}
    with st.expander("⚙️ Advanced - tab selection aur mapping", expanded=False):
        st.caption("Tabs: " + ", ".join(names))
        cols = st.columns(len(core.SOURCES))
        for c, k in zip(cols, core.SOURCES):
            d = core.pick_tab(names, core.DEFAULT_TABS[k])
            chosen[k] = c.selectbox(k, names, index=names.index(d) if d in names else 0, key=f"tab_{k}")
        st.caption("Mapping Google Sheet ke Mapping tab se aati hai. Permanent badlav sheet me karo (phir sidebar me refresh dabao); yahan ka edit sirf is run ke liye hai.")
        mapping = st.data_editor(mapping, num_rows="dynamic", width="stretch", key=f"map_{sid[:6]}")
    label = sid[:6]
    def load():
        fr = {k: tab_df(sid, t, key) for k, t in chosen.items()}
        return fr, [f"{k}: tab khaali hai" for k, v in fr.items() if v.empty]
else:
    up = st.sidebar.file_uploader("Project Excel (.xlsx)", type=["xlsx"])
    mup = st.sidebar.file_uploader("Mapping (.csv) - optional", type=["csv"])
    if not up: st.info("Left side se Excel upload karo."); st.stop()
    mapping = pd.read_csv(mup) if mup else core.DEFAULT_MAPPING.copy()
    label = up.name.rsplit(".", 1)[0]
    load = lambda: core.frames_from_excel(xl(up.getvalue()))
    with st.expander("⚙️ Advanced - mapping", expanded=False):
        mapping = st.data_editor(mapping, num_rows="dynamic", width="stretch", key=f"map_{label}")

if st.button("Recheck chalao", type="primary"):
    with st.spinner("Tabs padh ke compare ho raha hai (auto-match me ~30 sec lag sakte hain)..."):
        try:
            frames, load_warns = load()
            res = core.run(frames, mapping, tol, auto)
            res["warns"] = load_warns + res["warns"]; st.session_state["res"] = res; st.session_state["label"] = label
        except Exception as e: st.error(f"Run fail: {e}")

res = st.session_state.get("res")
if res:
    for w in res["warns"]: st.warning(w)
    s = res["summary"]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Centers", res["n"]); m2.metric("Gap cells", len(res["gaps"])); m3.metric("Rule errors", len(res["rules"]))
    m4.metric("Clean fields", int(((s["OK"] + s["BLANK BY RULE"]) == res["n"]).sum()))
    st.download_button("Excel download", core.to_excel_bytes(res), f"Recheck_{st.session_state['label']}.xlsx", type="primary")
    t1, t2, t3, t5, t4 = st.tabs(["Summary", "Gaps (center-wise)", "Rule errors", "Foldering (No)", "Full compare"])
    with t1: st.dataframe(s, width="stretch", hide_index=True, height=35 * (len(s) + 1) + 3)
    with t2:
        f = st.multiselect("Field filter", sorted(res["gaps"]["Field"].unique()))
        st.dataframe(res["gaps"][res["gaps"]["Field"].isin(f)] if f else res["gaps"], width="stretch", hide_index=True, height=520)
    with t3: st.dataframe(res["rules"], width="stretch", hide_index=True, height=520)
    with t5:
        fd = res.get("foldering")
        if fd is None or fd.empty: st.info("Foldering me koi 'No' center nahi mila (ya Foldering column detect nahi hua - upar warning dekho).")
        else: st.caption(f"{len(fd)} centers jinka Folder Status 'No' hai (Audit Final ya Center Data me). Gap column dikhata hai ki dono me match hai ya nahi."); st.dataframe(fd, width="stretch", hide_index=True, height=520)
    with t4: st.dataframe(res["compare"], width="stretch", height=520)

st.markdown("<div style='height:80px'></div>", unsafe_allow_html=True)
