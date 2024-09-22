# BSC-ops-admin
Ops admin automation code for Berkeley Student Cooperative. Code adapted from Ryan Frigo.

## Installation

```
pip install -e .
```

Other steps:
1. Create a `.env` folder in the root of the project.
2. Create a `.env/.env` file with the variable `EMAIL_PASSWORD` set to the gmail app-specific password.
3. Create a `credentials.json` file with the contents of the google api key, and also put it in the `.env` folder.
4. Make sure `SPREADSHEET_ID`, `POTENTIAL_TERMINATION_NOTICE_DOCUMENT_ID`, `CONDITIONAL_CONTRACT_DOCUMENT_ID`, `OPS_SUPERVISOR`, and `SEMESTER_YEAR` are up to date in `process_new_entries.py`.


## Running

```
python -m bsc_ops_admin.process_new_entries
```

# TODOs
- Add a test that checks all conditions run without errors