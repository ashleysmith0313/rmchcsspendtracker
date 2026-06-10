# RMCHCS Spend Tracker

Weekly spend tracking dashboard for Rehoboth McKinley Christian Health Care Services.
Built with Streamlit. Deployed to Streamlit Cloud. Internal use by Vista Staffing Solutions.

## Setup

### 1. Clone the repo
```
git clone https://github.com/YOUR_USERNAME/rmchcs-spend-tracker.git
cd rmchcs-spend-tracker
```

### 2. Install dependencies (local testing only)
```
pip install -r requirements.txt
```

### 3. Run locally
```
streamlit run app.py
```

### 4. Deploy to Streamlit Cloud
- Go to https://share.streamlit.io
- Connect your GitHub account
- Select this repo
- Set main file path to: app.py
- Click Deploy

## Data Storage

All spend entries are stored in `data/spend_log.json`. This file is committed to GitHub,
so your data persists across deployments. Do not delete it.

## Generating Reports

Use the "Generate Report" page to build a PDF for any logged week and download it directly.

## Project Structure

```
rmchcs-spend-tracker/
├── app.py                      Main Streamlit application
├── requirements.txt            Python dependencies
├── data/
│   └── spend_log.json          Persistent data store
├── reports/                    Generated PDFs (not committed to git)
├── utils/
│   ├── data_store.py           Read/write logic for spend_log.json
│   └── pdf_generator.py        ReportLab PDF report builder
└── .streamlit/
    └── config.toml             Streamlit server config
```

## Notes

- Data is stored in JSON in the repo itself. This works well for a single user or small team.
- PDFs are generated on demand and are not saved to the repo.
- For multi-user access with live data sync, a future upgrade would move to a database (Supabase or similar).
