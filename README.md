# Pisangboy Dashboard

This repository contains a simple HTML landing page and a Streamlit
application (`dashboard.py`) for exploring project data stored in an
Excel workbook.

## Streamlit app

The app expects a file named `Project_Dashboard_Data.xlsx` in the same
folder. You can also provide a URL to the workbook by setting the
`PROJECT_DASHBOARD_URL` environment variable. To start the app run:

```bash
streamlit run dashboard.py
```

