import io, json
import pandas as pd, streamlit as st
import core

st.set_page_config(page_title="Center Data Recheck", layout="wide")
st.title("Center Data Recheck Portal")
st.caption("Excel upload karo -> Center Data ko Calling / PM / CSR / Audit se auto-compare -> gaps + deduction errors + Excel download")

# ---------- Sidebar ----------
with st.sidebar:
    up = st.file_uploader("Project Excel (.xlsx)", type=["xlsx"])
    cfg_up = st.file_uploader("Saved config (.json)", type=["json"])
    cfg = json.loads(cfg_up.getvalue()) if cfg_up else json.loads(json.dumps(core.DEFAULT_CFG))
    project = st.text_input("Project Code filter (optional)", cfg.get("PROJECT", ""))
    cfg["PROJECT"] = project
    tol = st.number_input("% tolerance", 0.0, 0.1, float(cfg.get("PCT_TOL", 0.006)), 0.001, format="%.3f")
    cfg["PCT_TOL"] = tol

if not up:
    st.info("Left side se Excel upload karo."); st.stop()

@st.cache_resource(show_spinner=False)
def read(b): return pd.ExcelFile(io.BytesIO(b))
x = read(up.getvalue())
sheets = x.sheet_names

# ---------- Config ----------
with st.expander("1. Sheet settings (naye project me yahi badalna hai)", expanded=not cfg_up):
    cdc = cfg["CD"]
    c1, c2, c3, c4 = st.columns(4)
    cdc["sheet"] = c1.selectbox("Center Data sheet", sheets, index=sheets.index(cdc["sheet"]) if cdc["sheet"] in sheets else 0)
    cdc["tc"] = c2.text_input("TC column", cdc["tc"]); cdc["day_col"] = c3.text_input("Exam day column", cdc["day_col"]); cdc["project"] = c4.text_input("Project column", cdc.get("project", ""))
    rows = [dict(Key=k, Sheet=v["sheet"], Header=v["header"], TC=v["tc"], Project=v.get("project", "")) for k, v in cfg["SRC"].items()]
    ed = st.data_editor(pd.DataFrame(rows), num_rows="dynamic", use_container_width=True,
                        column_config={"Sheet": st.column_config.SelectboxColumn(options=sheets)}, key="src")
    new = {}
    for _, r in ed.dropna(subset=["Key"]).iterrows():
        old = cfg["SRC"].get(r.Key, {})
        new[r.Key] = {**old, "sheet": r.Sheet, "header": int(r.Header or 0), "tc": r.TC, "project": r.Project or ""}
    cfg["SRC"] = new
    cs = cfg["CSR"]
    st.markdown("**CSR sheet**")
    d1, d2, d3, d4 = st.columns(4)
    cs["sheet"] = d1.selectbox("CSR sheet", sheets, index=sheets.index(cs["sheet"]) if cs["sheet"] in sheets else 0)
    cs["tc"] = d2.text_input("CSR TC column", cs["tc"]); cs["day"] = d3.text_input("CSR day column (optional)", cs["day"]); cs["status"] = d4.text_input("CSR status column", cs["status"])
    e1, e2, e3 = st.columns(3)
    cs["param"] = e1.text_input("Rejected param column", cs["param"]); cs["remarks"] = e2.text_input("CSR remarks column", cs["remarks"])
    cs["cams"] = [v.strip() for v in e3.text_input("Working cam columns (comma se)", ", ".join(cs["cams"])).split(",") if v.strip()]
    st.caption("Header = header row ka number (0 = pehli row, Audit jaisi 2-row header me 1).")

with st.expander("2. Column mapping", expanded=not cfg_up):
    if st.button("Auto-detect mapping (value match se)"):
        with st.spinner("Columns match ho rahe hain..."):
            sug = core.auto_detect(x, cfg)
        st.session_state["sug"] = sug
    if "sug" in st.session_state:
        st.write("Suggestions (galat ho sakte hain jab 2 columns ki values same hon - review karo):")
        sug = st.data_editor(st.session_state["sug"], use_container_width=True, key="sugged")
        if st.button("Ye suggestions mapping me use karo"):
            cfg["MAP"] = [[r.Label, r["CD column"], r.Source, r["Source column"], r.Kind] for _, r in sug.iterrows()]
            st.session_state["map"] = pd.DataFrame(cfg["MAP"], columns=["Label", "CD column", "Source", "Source column", "Kind"])
    mdf = st.session_state.get("map", pd.DataFrame(cfg["MAP"], columns=["Label", "CD column", "Source", "Source column", "Kind"]))
    mdf = st.data_editor(mdf, num_rows="dynamic", use_container_width=True, key="mapedit",
                         column_config={"Kind": st.column_config.SelectboxColumn(options=["num", "pct", "text"])})
    cfg["MAP"] = mdf.dropna(subset=["CD column"]).values.tolist()
    st.download_button("Config save karo (.json)", json.dumps(cfg, indent=1), "recheck_config.json")

# ---------- Run ----------
if st.button("Recheck chalao", type="primary"):
    with st.spinner("Check ho raha hai..."):
        st.session_state["res"] = core.run(x, cfg)

res = st.session_state.get("res")
if res:
    for w in res["warns"]: st.warning(w)
    s = res["summary"]
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Centers", res["n"]); m2.metric("Gap cells", len(res["gaps"])); m3.metric("Rule errors", len(res["rules"]))
    m4.metric("Clean fields", int((s["OK"] == res["n"]).sum()) if len(s) else 0)
    st.download_button("Excel download", core.to_excel_bytes(res), "Recheck_Output.xlsx", type="primary")
    t1, t2, t3, t4 = st.tabs(["Summary", "Gaps (center-wise)", "Rule errors", "Full compare"])
    with t1:
        st.dataframe(s.style.map(lambda v: "background-color:#F8CBAD" if isinstance(v, (int, float)) and v > 0 else "", subset=["MISMATCH", "MISSING IN CD"]), use_container_width=True, hide_index=True)
    with t2:
        f = st.multiselect("Field filter", sorted(res["gaps"]["Field"].unique()))
        g = res["gaps"][res["gaps"]["Field"].isin(f)] if f else res["gaps"]
        st.dataframe(g, use_container_width=True, hide_index=True)
    with t3:
        st.dataframe(res["rules"], use_container_width=True, hide_index=True)
        st.caption("Rules core.py ke CHECKS list me edit hote hain.")
    with t4:
        st.dataframe(res["compare"], use_container_width=True)
