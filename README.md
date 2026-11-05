# Center Data Recheck Portal

Files (sab ek hi folder me, GitHub repo ke root pe):
app.py, core.py, gsheet.py, requirements.txt, mapping.csv, apps_script.gs

## Streamlit Cloud
1. Repo me ye saari files daalo (gsheet.py ka naam .py ke saath hona zaroori hai).
2. Main file = app.py
3. (Optional) App settings -> Secrets me secrets_example.toml jaisa daalo.

## Google Sheet se data (2 tareeke)
A) Apps Script (private sheet ke liye best)
   Sheet -> Extensions -> Apps Script -> apps_script.gs paste -> TOKEN badlo
   -> Deploy -> New deployment -> Web app -> Execute as: Me, Access: Anyone -> URL (/exec) copy
   Portal me URL + TOKEN daalo. Code badalne ke baad "New version" deploy karna.
B) API key: Sheet 'Anyone with link - Viewer' karo, Google Cloud me Sheets API enable karke key banao.
   Portal me sheet link + key daalo.
