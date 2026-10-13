# R4 Line-wise WIP Dashboard

Streamlit version of the "R4 WIP Dashboard (Standard for line WIP)" Excel workbook.

## Inputs (upload in the sidebar)
| File | Columns |
|---|---|
| Morning commitment (.xlsx/.csv) | `BAND`, `PLAN HOURS`, `DAY PLAN` (optional zone name in the first column, filled down) |
| System report (.xlsx/.csv) | `Workstation` (e.g. `R4-BAND-01`), `WIP Qty` (several rows per band are summed) |

Templates are in `templates/`. A full workbook's sheets can also be uploaded: the sheet name containing "commit" / "report" is picked automatically.

## Logic
- Hourly target = day plan / plan hours
- WIP hours = WIP / hourly target; WIP days = WIP / hourly target / plan hours
- Level by WIP days: >1.18 Red, >0.98 Green, >0.78 Yellow, >0.58 Gray, else Black
- Pcs dashboard: Target = day plan x 1.18, Actual = system WIP

## Run locally
```
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud
1. Create a GitHub repo and push this folder (see below).
2. Go to https://share.streamlit.io -> **Create app** -> pick the repo, branch `main`, main file `app.py`.
3. Click **Deploy**. Every `git push` redeploys automatically.

```
git init
git add .
git commit -m "Initial WIP dashboard"
git branch -M main
git remote add origin https://github.com/<you>/<repo>.git
git push -u origin main
```
Use a **private** repo if the code is internal (Community Cloud supports private repos). Do not commit real production data.
