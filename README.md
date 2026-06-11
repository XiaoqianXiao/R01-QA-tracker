# R01-QA-tracker

## Run the QA Summary Script

Place the REDCap QA exports and codebooks in the project `data/` folder:

```text
data/ParticipantsQAtracker.csv
data/ClinicianQAtracker.csv
data/Participants_REDCap.pdf
data/Clini_REDCap.pdf
```

From the project root, run:

```bash
python code/qa_summary.py --input-dir data --output-dir results
```

The script writes:

```text
results/QA_summary.xlsx
results/QA_summary_validation_log.txt
```

`QA_summary.xlsx` contains two sheets:

```text
subject_wise
group_wise
```

Before using the workbook for final QA decisions, review the `CONFIG` section near the top of `code/qa_summary.py` and fill in any project-specific rules, including expected sessions, expected questionnaires, MRI QC fields, behavioral QC fields, dropout fields, ASAP fields, and session timing windows. The validation log should also be reviewed after each run for unknown events, missing fields, duplicate records, date parsing failures, and subjects needing follow-up.
