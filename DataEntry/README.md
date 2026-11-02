# Recce Performance Data Entry

Run from the project root:

```bash
python -m pip install -r DataEntry/requirements.txt
streamlit run DataEntry/app.py
```

Use this local/internal app to add or correct `FactResult` rows and add Participants or Cycles. Close `Inputs/Step1_model_setup.xlsx` in Excel before saving. After an operational batch, run `python validate_data.py` from the project root.
