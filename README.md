# 🎓 Muthur Zone O/L 2025 Dashboard

A Streamlit app that analyses the O/L 2025 Results Register for the Muthur Zone
(T/Mu): Pass I / Pass II qualification, school rankings, subject and basket
analysis, a student explorer, data-quality diagnostics and Excel/CSV exports.

## Features

| Page | What it does |
|---|---|
| Zone Overview | Key metrics, performance bands, grade distribution, top/bottom schools |
| School Performance | Filters (band, size, name), scatter, ranking chart, heatmap, table + CSV |
| School Drill-down | One school vs zone averages, grade distribution per subject |
| Compare Schools | Side-by-side comparison of 2–6 schools incl. radar chart |
| Subject Analysis | Credit/pass %, categories, grade distribution, SBA vs credit % |
| Basket & Religion Choices | Most chosen subjects with credit rates |
| Student Explorer | Filter/search candidates, download CSV |
| Data Quality & Diagnostics | Grade Count mismatches, Pass II delta, unparsed cells |
| Downloads | Summary workbook, student-level workbook, Markdown summary |

The register is **uploaded in the sidebar** (processed in memory, never stored),
or read from a local `Results_Register-OL2025.xlsx` during local development.

## Project layout

```
app.py                 Streamlit UI
ol_analysis.py         Analysis engine (also a CLI)
tests/                 pytest suite (synthetic data)
requirements.txt       Runtime dependencies
requirements-dev.txt   + pytest, ruff
.streamlit/config.toml Theme / server settings
.streamlit/secrets.toml.example   Optional password
.github/workflows/ci.yml          Tests on every push
```

## Run locally

```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Command-line mode (writes the Excel/Markdown outputs):

```bash
python ol_analysis.py Results_Register-OL2025.xlsx
```

Run tests: `pip install -r requirements-dev.txt && pytest`

## Deploy via GitHub → Streamlit Community Cloud

1. Create a GitHub repository and push this folder:
   ```bash
   git init
   git add .
   git commit -m "Muthur Zone O/L dashboard"
   git branch -M main
   git remote add origin https://github.com/<you>/<repo>.git
   git push -u origin main
   ```
2. Go to <https://share.streamlit.io>, sign in with GitHub, click **Create app**,
   pick the repo, branch `main`, main file `app.py`, then **Deploy**.
3. (Recommended) In **App → Settings → Secrets** add:
   ```toml
   APP_PASSWORD = "your-strong-password"
   ```
   The app then asks for the password before showing anything.
4. Open the app and upload the Results Register in the sidebar.

Every `git push` to `main` redeploys automatically.

## ⚠️ Privacy

The register contains candidate-level exam results. The `.gitignore` excludes
all `Results_Register-*.xlsx` and generated output files — **do not commit them**,
especially to a public repository. Use the upload box, make the repo private,
and set `APP_PASSWORD`.

## Configuration

Pass criteria, subject codes and categories live at the top of `ol_analysis.py`
(`PASS_I_MIN_CREDITS`, `SUBJECT_NAMES`, `CATEGORY_OF_CODE`, …). The excluded
Census ID (16179) can be toggled in the sidebar.
