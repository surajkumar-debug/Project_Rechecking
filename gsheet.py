import re
import numpy as np, pandas as pd, requests
from urllib.parse import quote

API = "https://sheets.googleapis.com/v4/spreadsheets"


def sheet_id(url):
    """Google Sheet ka link (ya sirf ID) -> ID. Apps Script URL ho to waise ka waisa."""
    url = (url or "").strip()
    if "script.google.com" in url:
        return url
    m = re.search(r"/d/([\w-]+)", url)
    return m.group(1) if m else url


def is_script(src):
    return "script.google.com" in (src or "")


def _get(url, params):
    try:
        r = requests.get(url, params=params, timeout=120)
    except requests.RequestException as e:
        raise RuntimeError(f"Network error: {e}")
    if r.status_code != 200:
        try:
            msg = r.json()["error"]["message"]
        except Exception:
            msg = r.text[:200]
        hint = ""
        if r.status_code in (403, 404):
            hint = " | Sheet ko 'Anyone with the link - Viewer' karo aur Google Sheets API enabled key use karo (ya Apps Script Web App use karo)."
        if r.status_code == 400 and "API key" in msg:
            hint = " | API key galat hai."
        raise RuntimeError(f"Google API {r.status_code}: {msg}{hint}")
    try:
        return r.json()
    except Exception:
        raise RuntimeError("Link se JSON nahi aaya - Web App access 'Anyone' hai? URL /exec pe khatam hota hai? Deploy ke baad 'New version' banayi?")


def _script(src, token, **kw):
    j = _get(src.strip(), {"token": token, **kw})
    if isinstance(j, dict) and j.get("error"):
        raise RuntimeError("Apps Script: " + str(j["error"]))
    return j


def _frame(vals):
    if not vals:
        return pd.DataFrame()
    w = max(len(r) for r in vals)
    df = pd.DataFrame([list(r) + [""] * (w - len(r)) for r in vals], dtype=object)
    return df.replace("", np.nan)


def list_tabs(sid, key):
    if is_script(sid):
        return _script(sid, key, action="tabs")["tabs"]
    j = _get(f"{API}/{sid}", {"fields": "sheets.properties.title", "key": key})
    return [s["properties"]["title"] for s in j.get("sheets", [])]


def get_tab(sid, tab, key):
    if is_script(sid):
        return _frame(_script(sid, key, action="tab", name=tab).get("values", []))
    rng = quote("'" + tab.replace("'", "''") + "'", safe="")
    j = _get(f"{API}/{sid}/values/{rng}", {"valueRenderOption": "UNFORMATTED_VALUE", "key": key})
    return _frame(j.get("values", []))
