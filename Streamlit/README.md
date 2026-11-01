# Streamlit dashboard

Run locally from the project root:

```bash
python -m pip install -r Streamlit/requirements.txt
streamlit run Streamlit/app.py
```

Open the browser address shown by Streamlit, normally `http://localhost:8501`.

## Public demo deployment

This repository's workbook contains fictional mock data only. For Streamlit Community Cloud, select `Streamlit/app.py` as the entry point; the root `requirements.txt` installs its dependencies. Do not replace the public workbook with real participant data.

The dashboard reads `../Inputs/Step1_model_setup.xlsx` by default. Use the sidebar file-path control only for an approved local workbook. Do not expose Streamlit publicly or upload real participant data without authorization.

It provides the same six views as the PBIP project and includes data-quality checks. It does not provide editable data entry; update the input workbook, then use Streamlit’s **Clear cache / Reload** or restart the app.
