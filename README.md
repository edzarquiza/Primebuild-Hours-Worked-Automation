# 🏗️ Primebuild Hours Worked Automation

Automated compliance analysis tool for Primebuild employee timesheet exports.  
Built with **Python + Streamlit** — replaces the manual Excel VBA macro workflow.

---

## 🚀 Live App

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://your-app-url.streamlit.app)

---

## 📋 What It Does

| Check | Threshold | Output |
|---|---|---|
| Long Shift Detection | > 14 hours | Yellow highlight in Long Shift sheet |
| Short Break Detection | < 10 hours between consecutive shifts | Yellow highlight |
| Fatigue Risk Detection | Short break + combined hours > 14h | Red highlight |
| Weekly Hours Summary | Aggregated per employee | Red if > 60h/week |

### Workflow
1. **Upload** one or more weekly timesheet `.xlsx` exports
2. **Auto-processing** — filters Shift Work entries, sorts by employee & time
3. **Download** individual or bulk ZIP reports instantly

---

## 📁 Output Excel File (per upload)

Each report contains three sheets:

- **Long Shift** — all compliance flagged rows with colour coding
- **Weekly Hours** — total shift hours per employee, red if > 60h
- **Summary** — dashboard with key counts at a glance

---

## 🛠️ Local Setup

```bash
# Clone the repo
git clone https://github.com/your-username/primebuild-timesheet.git
cd primebuild-timesheet

# Install dependencies
pip install -r requirements.txt

# Run the app
streamlit run app.py
```

Then open [http://localhost:8501](http://localhost:8501) in your browser.

---

## ☁️ Deploy to Streamlit Cloud

1. Push this repo to GitHub
2. Go to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app** → select your repo → set `app.py` as the main file
4. Click **Deploy**

---

## 📂 Expected Input Format

The tool expects the standard Primebuild timesheet export (`.xlsx`) with an **Export** sheet containing these columns:

| Column | Description |
|---|---|
| `First Name` / `Surname` | Employee name |
| `Start Date` / `Start Time` | Shift start |
| `End Date` / `End Time` | Shift end |
| `Duration` | Shift length (HH:MM:SS) |
| `Work Type` | Must be `Shift work` to be included |
| `Location` | Project/site description |

---

## 🏢 Business Value

| Metric | Before | After |
|---|---|---|
| Processing time | ~20–30 min/report | < 5 seconds |
| Manual errors | Possible | Eliminated |
| Files supported | One at a time | Multiple files at once |
| Output | Manual Excel | Formatted, colour-coded Excel |

---

## 👤 Author

**Edward Jon Arquiza** — Primebuild
