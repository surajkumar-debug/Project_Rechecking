import datetime as dt, io
import numpy as np, pandas as pd, streamlit as st
import core, gsheet

st.set_page_config(page_title="Center Data Recheck", layout="wide")
st.title("Center Data Recheck Portal")

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
    """Control sheet ke 'Mapping' tab me hamare format ke columns hon to wahi use hota hai, warna default."""
    t = core.pick_tab(names, ["Mapping"])
    if t:
        raw = tab_df(sid, t, key)
        if not raw.empty:
            hdr = [str(c).strip() for c in raw.iloc[0]]
            if {"Field", "Source", "Center Data header", "Source header"} <= set(hdr):
                m = raw.iloc[1:].reset_index(drop=True); m.columns = hdr
                return m.dropna(how="all"), True
    return core.DEFAULT_MAPPING.copy(), False

with st.sidebar:
    mode = st.radio("Data kahan se?", ["Control Google Sheet", "Excel upload"])
    auto = st.checkbox("Auto-match extra columns", True, help="Mapping me jo columns nahi hain unhe value / naam match se khud dhoondta hai")
    tol = st.number_input("% tolerance", 0.0, 0.1, 0.006, 0.001, format="%.3f")

frames, load_warns, mapping, label = {}, [], None, "project"

if mode == "Control Google Sheet":
    url = st.text_input("Control sheet link ya Apps Script Web App URL (/exec)", secret("control_url"))
    key = st.text_input("Google API key ya Apps Script TOKEN", secret("api_key"), type="password")
    if not url or not key: st.info("Web App URL + TOKEN (ya Sheet link + API key) daalo."); st.stop()
    sid = gsheet.sheet_id(url)
    try: names = tabs_of(sid, key)
    except Exception as e: st.error(str(e)); st.stop()
    st.caption("Tabs mile: " + ", ".join(names))
    cols = st.columns(len(core.SOURCES)); chosen = {}
    for c, k in zip(cols, core.SOURCES):
        d = core.pick_tab(names, core.DEFAULT_TABS[k])
        chosen[k] = c.selectbox(k, names, index=names.index(d) if d in names else 0, key=f"tab_{k}")
    mapping, custom = read_mapping(sid, names, key)
    st.caption("Mapping: Control sheet ke 'Mapping' tab se" if custom else "Mapping: default (Mapping tab hamare format me nahi mila)")
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

with st.expander("Mapping (Field -> header aliases)"):
    st.caption("Header ke naam se match hota hai, position se nahi. Ek field ke multiple naam | se alag karo.")
    mapping = st.data_editor(mapping, num_rows="dynamic", use_container_width=True, key="map")
    st.download_button("Mapping CSV download", mapping.to_csv(index=False), "mapping.csv")

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
    t1, t2, t3, t4 = st.tabs(["Summary", "Gaps (center-wise)", "Rule errors", "Full compare"])
    with t1: st.dataframe(s, use_container_width=True, hide_index=True)
    with t2:
        f = st.multiselect("Field filter", sorted(res["gaps"]["Field"].unique()))
        st.dataframe(res["gaps"][res["gaps"]["Field"].isin(f)] if f else res["gaps"], use_container_width=True, hide_index=True)
    with t3: st.dataframe(res["rules"], use_container_width=True, hide_index=True)
    with t4: st.dataframe(res["compare"], use_container_width=True)
