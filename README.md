# I-FOCUS QA Summary Generator

By Xiaoqian Xiao (xxqian@uw.edu)

`qa_summary.py` builds arm-specific QA summary Excel workbooks from REDCap QA exports, study design workbooks, and MRI QC files.

The script is intentionally **design-file-driven**: required sessions are read from `Required_Sessions_for_each_Arm.xlsx`, expected instruments are read from `Instruments_in_each_Session_each_Arm.xlsx`, and only instruments marked with `Included_in_summary == 1` are included in the final outputs and denominators.

## What the script does

For each configured study arm, the script:

1. Reads the latest participant and clinician REDCap QA export files.
2. Standardizes subject IDs to three digits.
3. Infers study arm from the standardized subject ID.
4. Reads required sessions and expected instruments from the two design workbooks.
5. Builds subject-level questionnaire, behavioral task, self/other task QC, MRI QC, interval, and readiness fields.
6. Filters group-level summaries to subjects with `dropout_status` of `Active` or `Completed`.
7. Writes one Excel workbook per active arm, plus a validation log.

## Requirements

Use Python 3.10 or newer.

Install dependencies:

```bash
pip install pandas openpyxl
```

## Expected input folder

By default, the script reads inputs from a folder named `data/`.

```text
data/
├── IFOCUSStudyParticipa-QAtracker_DATA_*.csv
├── IFOCUSStudyClinician-QAtracker_DATA_*.csv
├── Required_Sessions_for_each_Arm.xlsx
├── Instruments_in_each_Session_each_Arm.xlsx
├── QC_anat.csv
├── QC_func.csv
├── Participants_REDCap.pdf        # optional sidecar; checked but not parsed
└── Clini_REDCap.pdf               # optional sidecar; checked but not parsed
```

### Required files

| File | Purpose |
|---|---|
| `IFOCUSStudyParticipa-QAtracker_DATA_*.csv` | Participant REDCap QA export. The latest matching file is used. |
| `IFOCUSStudyClinician-QAtracker_DATA_*.csv` | Clinician REDCap QA export. The latest matching file is used. |
| `Required_Sessions_for_each_Arm.xlsx` | Defines the required sessions for each arm. |
| `Instruments_in_each_Session_each_Arm.xlsx` | Defines expected instruments by arm/session. Only rows with `Included_in_summary == 1` are included. |
| `QC_anat.csv` | Anatomical MRI QC file. |
| `QC_func.csv` | Functional MRI QC file. |

The REDCap codebook PDFs are optional sidecars. If present, their presence is recorded in the validation log; if missing, the script records a warning.

## Running the script

Run with default folders:

```bash
python qa_summary.py
```

This reads from `data/` and writes to `results/`.

Run with custom folders:

```bash
python qa_summary.py --input-dir /path/to/data --output-dir /path/to/results
```

Command-line options:

| Option | Default | Description |
|---|---:|---|
| `--input-dir` | `data` | Directory containing REDCap exports, design workbooks, and MRI QC files. |
| `--output-dir` | `results` | Directory where Excel workbooks and the validation log are written. |

## Outputs

The script writes one workbook per active arm:

```text
results/
├── QA_summary_arm1.xlsx
├── QA_summary_arm2.xlsx
├── QA_summary_arm3.xlsx
└── QA_summary_validation_log.txt
```

Each workbook contains two sheets.

### `subject_wise`

One row per subject. This sheet includes:

- `subject_id`
- `arm`
- `dropout_status`
- session-level questionnaire date/status columns
- ANT completion status
- self/other task run status and QC fields
- MRI QC status fields
- session completion summaries
- interval-from-baseline fields
- overall subject-level QA flags

Important summary columns include:

| Column | Meaning |
|---|---|
| `complete_all_expected_experiment_sessions` | Whether all required sessions are complete. |
| `Nof_missing_expected_experiment_sessions` | Number of required sessions not complete. |
| `missing_expected_experiment_sessions` | Names of missing required sessions. |
| `complete_all_experiment_sessions` | Same value as `complete_all_expected_experiment_sessions`. |
| `complete_all_instrument` | Whether all expected questionnaire/instrument rows are complete. |
| `complete_all_ANT` | Whether all expected ANT rows are complete. |
| `all_MRI_QC_passed` | Whether all expected MRI QC rows pass. |
| `all_selfOther_QC_passed` | Whether all expected self/other task QC rows pass. |
| `subject_QC_pass` | Overall subject-level QA pass flag. |

### `group_wise`

A stacked set of group-level tables. The current script writes:

1. `Table 1. Session-wise Summary`
2. `Table 2. Instrument-wise Summary`
3. `Table 3-1. ANT task complete rate by session`
4. `Table 3-2. ANT task complete rate by subject`
5. `Table 4-1. SelfOthers QC pass rate by session`
6. `Table 4-2. SelfOthers QC pass rate by subject`
7. `Table 5. MRI QC pass rate by session`
8. `Table 5-2. MRI QC pass rate by subject`
9. `Table 6. Session interval summary after Baseline`
10. `Table 7. ASAP summary`
11. `Table 8. Participant-level QA readiness summary`

Group-level tables are generated after filtering to subjects whose `dropout_status` is `Active` or `Completed`.

### `QA_summary_validation_log.txt`

The validation log records:

- loaded files
- active arms
- row/subject counters
- excluded design rows
- missing optional sidecars
- unmapped REDCap events or MRI sessions
- missing instrument mappings
- other warnings generated during processing

## Key assumptions

### Subject IDs and arm mapping

The script standardizes subject IDs by extracting digits and keeping the last three digits.

Current arm inference is:

| First digit of standardized subject ID | Script arm label |
|---:|---|
| `1` | `arm1` |
| `2` | `arm3` |
| `3` | `arm2` |

If the study arm coding changes, update the `infer_arm()` function before running the script.

### Session handling

The script normalizes several session labels, for example:

| Input label | Normalized label |
|---|---|
| `Baseline 1`, `Baseline1`, `baseline_1`, `baseline` | `Baseline` |
| `Repeat baseline` | `Repeat Baseline` |
| `T3 IE` | `IE T3` |
| `T6 IE` | `IE T6` |
| `T9 IE` | `IE T9` |
| `T12 IE` | `IE T12` |

Final outputs exclude `Screening` and `ASAP` sessions from expected session columns, while still reporting `total_ASAP_count`.

### REDCap completion status

For standard REDCap completion fields:

| Value | Status |
|---:|---|
| `2` | `complete` |
| `1` | `unverified` |
| `0` | `incomplete` |
| blank | `missing` |
| other nonblank value | `review_required` |

For ANT-related checklist fields, the script applies special logic where values such as `1`, `yes`, `true`, or `checked` are treated as complete. Values such as `0`, `NA`, `N/A`, `not applicable`, or `none` are also treated as complete so the ANT QA check can pass when the field is not applicable.

### MRI QC status

MRI QC is based on the `Poor_Quality` column:

| `Poor_Quality` value | MRI QC status |
|---|---|
| true-like value | `fail` |
| false-like value | `pass` |
| other / unclear value | `review_required` |

The script expects:

- `QC_anat.csv` to contain `subID`, `sesID`, `modality`, and `Poor_Quality`
- `QC_func.csv` to contain `subID`, `sesID`, `modality`, `Poor_Quality`, `taskID`, and `runID`

## Configuration

Most project-specific choices are controlled near the top of the script in `QA_CONFIG`.

Default configuration:

```python
QA_CONFIG = {
    "active_arms": ["arm1", "arm2", "arm3"],
    "required_sessions": {
        "include_only": [],
        "exclude": {},
        "add": [],
    },
    "expected_instruments": {
        "include_only": {},
        "exclude": [],
        "add": [],
    },
    "subjects": {
        "include_only": [],
        "exclude": [],
    },
}
```

### Generate only selected arms

```python
QA_CONFIG["active_arms"] = ["arm1", "arm3"]
```

### Keep only selected sessions

For all active arms:

```python
QA_CONFIG["required_sessions"]["include_only"] = ["Baseline", "T1", "T2"]
```

For one arm:

```python
QA_CONFIG["required_sessions"]["include_only"] = {
    "arm1": ["Baseline", "T1", "T2"]
}
```

### Exclude a session

```python
QA_CONFIG["required_sessions"]["exclude"] = {
    "arm1": ["Repeat Baseline"]
}
```

### Keep only selected instruments

```python
QA_CONFIG["expected_instruments"]["include_only"] = {
    "arm1": {
        "Baseline": ["PHQ-9", "GAD-7", "SCID-5"]
    }
}
```

### Exclude a specific expected instrument

```python
QA_CONFIG["expected_instruments"]["exclude"] = [
    {"arm": "arm1", "session": "T12", "instrument": "SCID-5 T12"}
]
```

### Run only selected subjects

```python
QA_CONFIG["subjects"]["include_only"] = ["101", "102"]
```

Or by arm:

```python
QA_CONFIG["subjects"]["include_only"] = {
    "arm1": ["101", "102"],
    "arm3": ["201"]
}
```

### Exclude selected subjects

```python
QA_CONFIG["subjects"]["exclude"] = ["105"]
```

## Maintenance notes

Update these dictionaries when the REDCap project changes:

| Script object | When to update |
|---|---|
| `PARTICIPANT_INSTRUMENTS` | When participant instrument field names change or new participant instruments are added. |
| `CLINICIAN_INSTRUMENTS` | When clinician instrument field names change or new clinician instruments are added. |
| `ANT_FIELDS_BY_SESSION` | When ANT checklist fields differ by session. |
| `MRI_SESSION_MAP` | When MRI `sesID` labels change. |
| `SESSION_ORDER` | When final output session order changes. |
| `infer_arm()` | When subject-ID-to-arm mapping changes. |

## Troubleshooting

### `No file found matching pattern`

Check that the input directory contains REDCap exports matching:

```text
IFOCUSStudyParticipa-QAtracker_DATA_*.csv
IFOCUSStudyClinician-QAtracker_DATA_*.csv
```

The script uses the most recently modified matching file for each export type.

### `Required design file not found`

Confirm that both design workbooks are in the input directory:

```text
Required_Sessions_for_each_Arm.xlsx
Instruments_in_each_Session_each_Arm.xlsx
```

### Missing questionnaire columns in the output

Check `Instruments_in_each_Session_each_Arm.xlsx` and confirm the instrument row has:

```text
Included_in_summary = 1
```

Also check that the instrument name matches one of the names in `PARTICIPANT_INSTRUMENTS` or `CLINICIAN_INSTRUMENTS`.

### Unexpected missing MRI QC rows

Check that MRI sessions in `QC_anat.csv` and `QC_func.csv` are represented in `MRI_SESSION_MAP`, and that the normalized session is required for that arm.

### Unexpected missing or unmapped REDCap events

Check the `participant_event_name` and `clinician_event_name` columns in `Required_Sessions_for_each_Arm.xlsx`. These values are used to map REDCap events into final session labels.

### Output validation errors

Before saving each workbook, the script checks that:

- `subject_wise` has no duplicate `subject_id` rows
- `subject_wise` has no duplicate columns
- final column names do not contain forbidden tokens such as `Screening`, `Baseline1`, or `UNMAPPED_*`
- `group_wise` contains the required set of summary tables

If validation fails, inspect the error message and the validation log.

## Recommended workflow

1. Export the newest participant and clinician REDCap QA files.
2. Place them in the input directory with the two design workbooks and MRI QC files.
3. Confirm `QA_CONFIG` is set correctly.
4. Run:

   ```bash
   python qa_summary.py --input-dir data --output-dir results
   ```

5. Review `QA_summary_validation_log.txt` first.
6. Review each arm-specific Excel workbook.
7. If warnings indicate missing mappings, update the relevant dictionary or design workbook and rerun.
