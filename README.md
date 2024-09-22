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
- Add caching of google docs based on last modified time
- Add cron job that runs this daily (and runs if missed a day). Maybe run every hour and check if the last time it ran was > 1 day ago. Maybe have a MacOS pop up appear when asking user input.
- Document what/how much of the script is MacOS specific. Definitely have a non-MacOS version that can be run as a python package without cron.
- Add a schedule-send email after the 15 day notice is up, to ops admin? Or maybe have it schedule a job for 15 days later that checks whether it was done
- Add better handling of people eligible for reinstatement (ask Alex for a list of terminated people and the reasons why they were terminated). Document how that was obtained.
