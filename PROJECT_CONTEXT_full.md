# PROJECT_CONTEXT.md — R01-QA-tracker

## Project purpose

The R01-QA-tracker project uses REDCap QA tracker exports, REDCap codebooks, external study-design Excel files, and external MRI QC metric files to generate automatic QA summary workbooks.

The goal is to reduce manual checking of REDCap records by producing arm-specific standardized Excel files:

```text
QA_summary_arm1.xlsx
QA_summary_arm2.xlsx
QA_summary_arm3.xlsx
```

Each workbook should help the team quickly evaluate, within one study arm:

- subject-level questionnaire completion
- task completion
- behavioral QC status
- MRI QC status from external QC CSV files
- post-baseline session timing
- ASAP burden
- dropout/withdrawal status
- overall readiness for analysis
- group-level missingness, completion rates, QC pass rates, and readiness rates

The script should also produce a validation log, either one shared log or one log per arm:

```text
QA_summary_validation_log.txt
QA_summary_arm1_validation_log.txt
QA_summary_arm2_validation_log.txt
QA_summary_arm3_validation_log.txt
```

The main script is:

```text
qa_summary.py
```

---

## Dataset

### Input files overview

The script uses the following input files:

```text
ParticipantsQAtracker.csv
ClinicianQAtracker.csv
Required_Sessions_for_each_Arm.xlsx
Instruments_in_each_Session_each_Arm.xlsx
QC_anat.csv
QC_func.csv
Participants_REDCap.pdf
Clini_REDCap.pdf
```

The subject ID source and normalization rules should be defined after this complete input-file overview, before the detailed processing rules for each file type.

### Subject ID source, normalization, and arm assignment rule

The script must create one standardized `subject_id` field before merging or summarizing records. The source column differs by input file:

```text
ParticipantsQAtracker.csv -> subject_id comes from record_id
ClinicianQAtracker.csv    -> subject_id comes from preescreen_id
QC_anat.csv               -> subject_id comes from subID
QC_func.csv               -> subject_id comes from subID
```

For REDCap QA exports:

- In `ParticipantsQAtracker.csv`, use `record_id` as the subject ID.
- In `ClinicianQAtracker.csv`, use `preescreen_id` as the subject ID.
- The first digit of this three-digit subject ID determines the arm.

For external MRI QC files:

- In `QC_anat.csv` and `QC_func.csv`, use `subID` as the subject ID source.
- `subID` is expected to have the BIDS-style prefix `sub-`.
- The standardized `subject_id` should be the three digits after `sub-`.
- The first digit of those three digits determines the arm.

Examples:

```text
ParticipantsQAtracker.csv: record_id     = 101     -> subject_id = 101 -> Arm 1
ClinicianQAtracker.csv:    preescreen_id = 205     -> subject_id = 205 -> Arm 2
QC_anat.csv/QC_func.csv:   subID         = sub-312 -> subject_id = 312 -> Arm 3
```

Arm assignment rule:

```text
subject_id first digit == 1 -> Arm 1
subject_id first digit == 2 -> Arm 2
subject_id first digit == 3 -> Arm 3
```

Across `ParticipantsQAtracker.csv`, `ClinicianQAtracker.csv`, `QC_anat.csv`, and `QC_func.csv`, records with the same standardized `subject_id` belong to the same subject and must be merged into one subject-level row.

Recommended implementation:

```python
def standardize_subject_id(raw_value: str, source_file: str) -> str:
    raw = str(raw_value).strip()

    if source_file in {"QC_anat.csv", "QC_func.csv"}:
        # MRI QC files store subject IDs as BIDS-style values such as sub-101.
        if raw.startswith("sub-"):
            raw = raw[4:]

    digits = ''.join(ch for ch in raw if ch.isdigit())

    # Expected final subject_id is the three-digit study ID.
    if len(digits) >= 3:
        return digits[-3:]
    return digits


def infer_arm_from_subject_id(subject_id: str) -> str:
    sid = standardize_subject_id(subject_id, source_file="standardized")
    if len(sid) < 3:
        return "unknown"

    arm_digit = sid[0]
    if arm_digit == "1":
        return "arm1"
    if arm_digit == "2":
        return "arm2"
    if arm_digit == "3":
        return "arm3"
    return "unknown"
```

Unknown or invalid subject IDs or arm assignments should be logged and excluded from arm-specific workbooks unless an explicit debug option is used.

---

### REDCap QA exports

The script reads two REDCap QA tracker CSV files:

```text
ParticipantsQAtracker.csv
ClinicianQAtracker.csv
```

Both files may contain records for the same standardized `subject_id`. These should be treated as records for the same participant. The file-specific subject ID rules are defined in the subject ID section above.

### Required session and instrument design files

The script must read two study-design Excel files from the project data/input directory:

```text
Required_Sessions_for_each_Arm.xlsx
Instruments_in_each_Session_each_Arm.xlsx
```

These files define the expected protocol structure and must be treated as the source of truth for required sessions and expected instruments.

`Required_Sessions_for_each_Arm.xlsx` should be used to determine:

- which sessions are required for Arm 1
- which sessions are required for Arm 2
- which sessions are required for Arm 3
- which sessions should be considered expected for each subject based on the subject's arm

`Instruments_in_each_Session_each_Arm.xlsx` should be used to determine:

- which instruments/questionnaires/tasks/QC items are required in each session for each arm
- which instruments should be included in `subject_wise`
- which instruments should be included in group-level expected denominators
- which instruments should be treated as expected for subject-level completeness/readiness summaries

Critical instrument-inclusion rule:

```text
Only instruments with Included_in_summary == 1 should be included in expected instruments.
```

This rule must be applied at the arm + session + instrument level.

For a given subject, an instrument is expected only if all of the following are true:

1. the subject belongs to that arm
2. the session is required for that arm in `Required_Sessions_for_each_Arm.xlsx`
3. the instrument appears for that arm/session in `Instruments_in_each_Session_each_Arm.xlsx`
4. `Included_in_summary` for that arm/session/instrument is marked as `1`

Rows in `Instruments_in_each_Session_each_Arm.xlsx` with `Included_in_summary` blank, `0`, `no`, `false`, or any value other than `1` should not create expected instruments, should not create subject-level columns, and should not contribute to group-level expected denominators.

Do not infer required sessions or expected instruments from observed REDCap rows alone. Observed rows can confirm completion, but the Excel design files define what is required and what should be included in the summary.

### External MRI QC CSV files

MRI QC should **not** be read from REDCap QA tracker columns.

MRI QC should be read from these two external CSV files in the project data directory:

```text
/Users/xiaoqianxiao/PycharmProjects/R01-QA-tracker/data/QC_anat.csv
/Users/xiaoqianxiao/PycharmProjects/R01-QA-tracker/data/QC_func.csv
```

The script should resolve these files relative to the selected input/data directory when possible. If the script is run from the project root or with `--input-dir data`, the default lookup should find:

```text
data/QC_anat.csv
data/QC_func.csv
```

`QC_anat.csv` contains anatomical MRI QC rows. Required columns:

```text
subID
sesID
modality
Poor_Quality
```

`QC_func.csv` contains functional MRI QC rows. Required columns:

```text
subID
sesID
modality
taskID
runID
Poor_Quality
```

The MRI QC flag rule is:

```text
Poor_Quality == True  -> qc_status = fail, qc_pass = False
Poor_Quality == False -> qc_status = pass, qc_pass = True
```

Accepted raw boolean values should include at least:

```text
True, true, TRUE, 1, yes, y
False, false, FALSE, 0, no, n
```

Missing or unrecognized `Poor_Quality` values should become:

```text
qc_status = review_required
qc_pass = False
missingness_reason = review_required
```

### REDCap codebooks

The script also uses two REDCap codebook PDFs:

```text
Participants_REDCap.pdf
Clini_REDCap.pdf
```

The codebooks should be used to confirm or extract:

- REDCap unique event names
- human-readable Event Name labels from the codebook `Events` section
- form/instrument names
- field names and field labels
- coded choices and readable values
- REDCap completion fields
- date fields
- task completion fields
- behavioral QC fields
- dropout/withdrawal fields
- ASAP-related fields

MRI QC should come from `QC_anat.csv` and `QC_func.csv`, not from the REDCap codebooks or REDCap QA exports.

If PDF parsing is incomplete or unreliable, the script should use manual mappings in the `CONFIG` section of `qa_summary.py`.


## Architecture

The pipeline should be configuration-first and design-file-driven.

Do not infer expectedness from missing REDCap rows alone. A missing row does not automatically mean an error, and an observed row does not automatically mean the form was expected.

Expectedness should come from:

1. `Required_Sessions_for_each_Arm.xlsx`
2. `Instruments_in_each_Session_each_Arm.xlsx`
3. the `Included_in_summary` column in `Instruments_in_each_Session_each_Arm.xlsx`
4. clearly labeled fallback mappings in the `CONFIG` section only when the design files are missing or incomplete

The script should build internal long-format tables first, then pivot or summarize them into arm-specific final Excel outputs.

Recommended internal tables:

```text
subject_arm_map
required_sessions_by_arm
expected_instruments_by_arm_session
questionnaire_long
task_completion_long
mri_qc_long
behavioral_qc_long
session_long
```

Important naming rule:

```text
expected_instruments_by_arm_session
```

is preferred over `required_instruments_by_arm_session`, because an instrument should only be expected/included when `Included_in_summary == 1`.

If the code keeps the old internal table name `required_instruments_by_arm_session`, it must still apply the `Included_in_summary == 1` filter before using that table for expectedness, subject-wise columns, or group-wise denominators.

Final outputs should be written separately by arm:

```text
QA_summary_arm1.xlsx
QA_summary_arm2.xlsx
QA_summary_arm3.xlsx
```

Each workbook should contain:

```text
subject_wise
group_wise
```

---

## Core design rules

### 1. Subject merge rule

Use only the standardized three-digit:

```text
subject_id
```

as the subject-level merge key.

Create `subject_id` from the correct file-specific source column before merging:

```text
ParticipantsQAtracker.csv -> record_id
ClinicianQAtracker.csv    -> preescreen_id
QC_anat.csv               -> subID, remove sub- and keep the three digits after it
QC_func.csv               -> subID, remove sub- and keep the three digits after it
```

If the same standardized `subject_id` appears in `ParticipantsQAtracker.csv`, `ClinicianQAtracker.csv`, `QC_anat.csv`, and/or `QC_func.csv`, merge all relevant records into one row in the `subject_wise` sheet for that subject's arm.

Do not include this column in the final output:

```text
source_group
```

It is acceptable to keep these internally:

```text
source_file
instrument_source
```

These internal fields are useful for debugging, duplicate detection, group-level summaries, and preventing output column overwriting.

For MRI QC rows, use:

```text
source_file = mriqc_anat or mriqc_func
instrument_source = mriqc
```

### 2. Arm-specific output rule

Generate one output workbook per arm.

Each arm-specific workbook should include only subjects assigned to that arm by the subject ID rule.

Each arm-specific `subject_wise` sheet should include only:

1. sessions required for that arm according to `Required_Sessions_for_each_Arm.xlsx`
2. instruments for that arm/session according to `Instruments_in_each_Session_each_Arm.xlsx`
3. instruments with `Included_in_summary == 1`

Do not include non-required sessions or instruments with `Included_in_summary != 1` in the subject-level sheet, even if such records are observed in the REDCap export.

Observed but not expected records should be logged as unexpected observations, not written as standard subject-level columns unless a debug output is explicitly enabled.

### 3. Required-session and included-instrument subject-level output rule

The `subject_wise` sheet should be protocol-driven.

For each subject:

1. infer arm from `subject_id`
2. get required sessions for that arm from `Required_Sessions_for_each_Arm.xlsx`
3. get candidate instruments for each required session from `Instruments_in_each_Session_each_Arm.xlsx`
4. keep only candidate instruments where `Included_in_summary == 1`
5. generate columns only for those required sessions and included instruments
6. fill observed values when present
7. mark missing expected records as `missing`
8. do not create columns for non-required sessions/instruments or for instruments where `Included_in_summary != 1`

This applies to:

- questionnaire columns
- task completion columns
- behavioral QC columns
- MRI QC columns, when MRI expectedness is defined by arm/session/instrument design
- session timing columns
- subject-level completeness/readiness summaries

### 4. Instrument disambiguation rule

Subjects are merged by `subject_id`, but instruments from different source files must not overwrite each other.

If a questionnaire/task/scan appears in only one source file, use the clean instrument name directly.

Examples:

```text
ses-Baseline_qn-PHQ9_status
ses-T6Scan_mri-anat_T1w_qc_status
ses-T6Scan_mri-func_selfother_run02_qc_status
```

If the same `session + questionnaire`, `session + task`, or `session + scan_or_run` appears in more than one source file, prefix the output name with the source.

Examples:

```text
ses-Baseline_qn-participants_VisitStatus_status
ses-Baseline_qn-clinician_VisitStatus_status

ses-T6_task-participants_SelfOther_status
ses-T6_task-clinician_SelfOther_status

ses-T6_beh-participants_SelfOther_qc_pass
ses-T6_beh-clinician_SelfOther_qc_pass
```

MRI QC output names should be generated from external MRI QC file metadata, not REDCap column names.

Recommended MRI scan/run output names:

```text
anat_T1w
func_rest_run01
func_selfother_run01
func_selfother_run02
```

### 5. REDCap event name rule

Do not allow REDCap internal event names in final Excel outputs.

Do not show names like:

```text
baseline_1_arm_1
t1_arm_1
```

Use the human-readable `Event Name` from the codebook `Events` section.

Examples:

```text
baseline_1_arm_1 -> Baseline
t1_arm_1 -> T1
t6_scan_arm_1 -> T6 Scan
```

The final canonical baseline label is:

```text
Baseline
```

Do not use these final labels:

```text
Baseline 1
Baseline1
```

If an event cannot be mapped, log it and raise an error by default. A debugging flag such as `--allow-unmapped-events` may be added, but unmapped events should not be allowed in routine output.

### 6. MRI QC session mapping rule

External MRI QC files use BIDS-style session IDs in `sesID`.

The script must map `sesID` values to the same canonical session names used in the rest of the project.

Recommended default mapping:

```python
"mri_qc_session_mapping": {
    "ses-baseline": "Baseline",
    "ses-repeatbaseline": "Repeat Baseline",
    "ses-T1": "T1",
    "ses-T2": "T2",
    "ses-T3": "T3",
    "ses-IE-T3": "IE T3",
    "ses-T4": "T4",
    "ses-T5": "T5",
    "ses-T6": "T6 Scan",
    "ses-IE-T6": "IE T6",
    "ses-T7": "T7",
    "ses-T8": "T8",
    "ses-T9": "T9",
    "ses-IE-T9": "IE T9",
    "ses-T10": "T10",
    "ses-T11": "T11",
    "ses-T12": "T12 Scan",
    "ses-IE-T12": "IE T12",
}
```

Important MRI scan-session rule:

```text
ses-T6  in MRI QC files should map to T6 Scan.
ses-T12 in MRI QC files should map to T12 Scan.
```

This avoids mixing the treatment-session behavioral/questionnaire records with the MRI scan records.

### 7. Screening exclusion rule

Screening can be recognized internally only for filtering, validation, or debugging.

Do not include Screening in final Excel outputs.

No final output should contain:

```text
Screening
ses-Screening_*
```

This applies to:

- `subject_wise` questionnaire columns
- `subject_wise` task completion columns
- `subject_wise` behavioral QC columns
- `subject_wise` MRI QC columns
- `subject_wise` session interval columns
- every `group_wise` table

### 8. Expectedness rule

All rates should use expected records as the denominator.

Expectedness should be arm-specific and should come from the design Excel files.

A record is expected only when:

1. the session is required for that arm in `Required_Sessions_for_each_Arm.xlsx`
2. the instrument is listed for that arm/session in `Instruments_in_each_Session_each_Arm.xlsx`
3. `Included_in_summary == 1` for that arm/session/instrument

Do not count `not_expected` records in denominators.

Do not count instruments with `Included_in_summary != 1` in denominators.

Examples:

```text
complete_rate = complete_NofSubjects / expected_NofSubjects
missing_rate = missing_NofSubjects / expected_NofSubjects
qc_pass_rate = qc_pass_NofSubjects / expected_NofSubjects
```

For MRI QC from external CSV files:

- Observed rows in `QC_anat.csv` and `QC_func.csv` are expected by default only if MRI expectedness is not explicitly available in the design files.
- If the arm/session/instrument design files define MRI expectedness, use only included MRI instruments/items where `Included_in_summary == 1` as the denominator source.
- Missing MRI QC rows for expected MRI items should become `qc_status = missing`.

### 9. Session interval rule

Only calculate intervals from Baseline for required sessions after Baseline.

Do not create interval columns for:

```text
Screening
Baseline
```

Do create interval columns only for required post-baseline sessions for that subject's arm.

Session interval logic should be used only for timing/window summaries. It should not define whether all experiment sessions are complete.

### 10. `complete_all_experiment_sessions` rule

`complete_all_experiment_sessions` must be calculated from the required session design, not from interval-window rules.

Correct logic:

```text
complete_all_experiment_sessions = True only if every required experiment session for that subject's arm has evidence of completion.
```

Use `Required_Sessions_for_each_Arm.xlsx` as the source of required sessions.

Do not calculate `complete_all_experiment_sessions` using:

```text
session_long[session_long["interval_valid"] != "not_expected"]
```

because sessions without timing-window rules may be incorrectly excluded from the all-session completion check.

Recommended implementation concept:

```python
required_sessions = get_required_sessions_for_arm(subject_arm)
required_sessions = required_sessions - {"Screening"}

complete_all_experiment_sessions = all(
    session_completed(subject_id, session)
    for session in required_sessions
)
```

Recommended diagnostic output:

```text
missing_required_experiment_sessions
```

This should list required sessions that do not have completion evidence for that subject.

---

## Configuration expectations

The `CONFIG` section should define project-specific rules and mappings.

Recommended components:

```python
CONFIG = {
    "subject_id_source_fields_by_file": {
        "ParticipantsQAtracker.csv": "record_id",
        "ClinicianQAtracker.csv": "preescreen_id",
        "QC_anat.csv": "subID",
        "QC_func.csv": "subID",
    },

    "mri_qc_subject_id_prefix": "sub-",
    "standard_subject_id_digits": 3,

    "design_files": {
        "required_sessions_by_arm": "Required_Sessions_for_each_Arm.xlsx",
        "instruments_by_session_by_arm": "Instruments_in_each_Session_each_Arm.xlsx",
    },

    "instrument_inclusion_field": "Included_in_summary",
    "instrument_inclusion_values": ["1", 1, True, "true", "yes", "y"],

    "arm_assignment": {
        "method": "standardized_three_digit_subject_id_first_digit",
        "arm_digit_map": {"1": "arm1", "2": "arm2", "3": "arm3"},
    },

    "output_by_arm": True,
    "output_file_template": "QA_summary_{arm}.xlsx",

    "event_field_candidates": [
        "redcap_event_name",
        "event_name",
        "event",
    ],

    "event_mapping": {
        "screening_arm_1": "Screening",
        "baseline_1_arm_1": "Baseline",
        "repeat_baseline_arm_1": "Repeat Baseline",
        "t1_arm_1": "T1",
        "t2_arm_1": "T2",
        "t3_arm_1": "T3",
        "ie_t3_arm_1": "IE T3",
        "t4_arm_1": "T4",
        "t5_arm_1": "T5",
        "t6_arm_1": "T6",
        "ie_t6_arm_1": "IE T6",
        "t6_scan_arm_1": "T6 Scan",
        "t7_arm_1": "T7",
        "t8_arm_1": "T8",
        "t9_arm_1": "T9",
        "ie_t9_arm_1": "IE T9",
        "t10_arm_1": "T10",
        "t11_arm_1": "T11",
        "t12_arm_1": "T12",
        "ie_t12_arm_1": "IE T12",
        "t12_scan_arm_1": "T12 Scan",
    },

    "session_aliases": {
        "Baseline 1": "Baseline",
        "Baseline1": "Baseline",
        "Treatment Session 1": "T1",
        "Treatment Session 2": "T2",
        "Treatment Session 3": "T3",
        "Treatment Session 3 (IE)": "IE T3",
        "Treatment Session 4": "T4",
        "Treatment Session 5": "T5",
        "Treatment Session 6": "T6",
        "Treatment Session 6 (IE)": "IE T6",
        "Treatment Session 7": "T7",
        "Treatment Session 8": "T8",
        "Treatment Session 9": "T9",
        "Treatment Session 9 (IE)": "IE T9",
        "Treatment Session 10": "T10",
        "Treatment Session 11": "T11",
        "Treatment Session 12": "T12",
        "Treatment Session 12 (IE)": "IE T12",
    },

    "mri_qc_files": {
        "anat": "QC_anat.csv",
        "func": "QC_func.csv",
    },

    "mri_qc_session_mapping": {
        "ses-baseline": "Baseline",
        "ses-repeatbaseline": "Repeat Baseline",
        "ses-T1": "T1",
        "ses-T2": "T2",
        "ses-T3": "T3",
        "ses-IE-T3": "IE T3",
        "ses-T4": "T4",
        "ses-T5": "T5",
        "ses-T6": "T6 Scan",
        "ses-IE-T6": "IE T6",
        "ses-T7": "T7",
        "ses-T8": "T8",
        "ses-T9": "T9",
        "ses-IE-T9": "IE T9",
        "ses-T10": "T10",
        "ses-T11": "T11",
        "ses-T12": "T12 Scan",
        "ses-IE-T12": "IE T12",
    },

    "mri_qc_poor_quality_field": "Poor_Quality",
    "mri_qc_true_values": ["true", "1", "yes", "y"],
    "mri_qc_false_values": ["false", "0", "no", "n"],

    "excluded_output_sessions": [
        "Screening",
    ],

    "session_order": [
        "Baseline",
        "Repeat Baseline",
        "T1",
        "T2",
        "T3",
        "IE T3",
        "T4",
        "T5",
        "T6",
        "IE T6",
        "T6 Scan",
        "T7",
        "T8",
        "T9",
        "IE T9",
        "T10",
        "T11",
        "T12",
        "IE T12",
        "T12 Scan",
    ],

    "baseline_session_name": "Baseline",

    "exclude_from_interval_sessions": [
        "Screening",
        "Baseline",
    ],

    "task_completion_complete_values": ["complete", "completed", "done", "yes", "y", "2"],
    "task_completion_incomplete_values": ["incomplete", "not_complete", "not completed", "no", "n", "0"],
    "task_completion_unverified_values": ["partial", "partially complete", "unverified", "1"],

    "behavioral_qc_pass_values": ["pass", "passed", "usable", "1", "yes"],
    "behavioral_qc_fail_values": ["fail", "failed", "unusable", "0", "no"],

    "dropout_fields": [],
    "dropout_values": ["dropout", "withdrawn", "withdrew", "yes", "1"],

    "asap_fields": [],
    "session_timing_windows": {},
    "session_date_fields": {},
    "date_fields_by_questionnaire": {},
    "completion_status_fields_by_form": {},

    "treat_observed_as_expected_when_no_design_config": False,
}
```

Important note for REDCap completion codes: for standard REDCap `_complete` fields, `2` usually means complete, `1` means unverified, and `0` means incomplete. Do not place the same coded value in both complete and unverified lists for the same field type.

Do not keep or use `total_score_fields_by_questionnaire` for final output. Questionnaire total score columns should not be written to `subject_wise`.

Do not use `mri_qc_fields` for MRI QC final output. MRI QC should be loaded from `QC_anat.csv` and `QC_func.csv`.

---

## Data processing strategy

### Step 1. Read input files

Read all raw REDCap fields, study-design Excel files, and MRI QC CSV fields as strings when possible:

```python
pd.read_csv(..., dtype=str, keep_default_na=False)
pd.read_excel(..., dtype=str, keep_default_na=False)
```

This avoids losing subject ID formatting, leading zeros, coded values, and raw date text.

### Step 2. Load design files

Load:

```text
Required_Sessions_for_each_Arm.xlsx
Instruments_in_each_Session_each_Arm.xlsx
```

Create normalized internal design tables:

```text
required_sessions_by_arm
expected_instruments_by_arm_session
```

Recommended `required_sessions_by_arm` columns:

```text
arm
session
required
session_order_index
design_order
participant_event_name
clinician_event_name
```

Recommended `expected_instruments_by_arm_session` columns:

```text
arm
session
instrument
instrument_key
instrument_type
instrument_source
included_in_summary
expected
session_order_index
instrument_order_index
```

Rules for `expected_instruments_by_arm_session`:

- read the `Included_in_summary` column from `Instruments_in_each_Session_each_Arm.xlsx`
- normalize the column name so that `Included_in_summary`, `included in summary`, and `included_in_summary` are recognized
- keep only rows where `Included_in_summary == 1`
- treat kept rows as expected instruments
- exclude rows where `Included_in_summary` is blank, `0`, `no`, `false`, or any value other than `1`
- log the number of included and excluded instrument-design rows for validation

`instrument_type` should use controlled values such as:

```text
questionnaire
task_completion
behavioral_qc
mri_qc
session_status
```

Normalize all session names to the canonical session labels used in `CONFIG["session_order"]`.

### Step 3. Normalize columns

Normalize REDCap export column names to stable snake_case names for internal processing.

For external MRI QC CSVs, either preserve known column names before mapping or normalize them with a controlled mapping:

```text
subID -> subject_id
sesID -> raw_mri_session
Poor_Quality -> poor_quality
```

### Step 4. Standardize subject IDs and infer arm

Create one standardized three-digit field before any merge:

```text
subject_id
```

Use file-specific subject ID source columns:

```text
ParticipantsQAtracker.csv -> record_id
ClinicianQAtracker.csv    -> preescreen_id
QC_anat.csv               -> subID
QC_func.csv               -> subID
```

For `ParticipantsQAtracker.csv`, set `subject_id` from `record_id`.

For `ClinicianQAtracker.csv`, set `subject_id` from `preescreen_id`.

For `QC_anat.csv` and `QC_func.csv`, set `subject_id` from `subID` by removing the `sub-` prefix and keeping the three digits after it. For example, `sub-101` becomes `101`.

Then infer:

```text
arm
```

from the first digit of the standardized three-digit `subject_id`.

All downstream subject-level and group-level summaries should use this standardized `subject_id` and this arm assignment.

### Step 5. Standardize events and sessions

For REDCap exports, create:

```text
raw_event_name
session
```

`session` must always use readable, canonical labels.

For MRI QC CSVs, create:

```text
raw_mri_session = sesID
session = canonical mapped value from CONFIG["mri_qc_session_mapping"]
```

Unknown MRI QC `sesID` values should be logged and should become `UNMAPPED_MRI_SESSION_<sesID>` or raise an error, depending on the same validation policy used for REDCap events.

### Step 6. Build long-format internal tables

Build these internal tables before creating the Excel output.

#### `questionnaire_long`

Recommended columns:

```text
subject_id
arm
source_file
instrument_source
session
questionnaire
questionnaire_output_name
date
status
expected
missingness_reason
raw_form_name
raw_completion_value
included_in_summary
```

Allowed statuses:

```text
complete
incomplete
missing
unverified
not_expected
review_required
```

Do not write total score fields to final output.

Only rows matching expected instruments with `Included_in_summary == 1` should be used for subject-wise expectedness and group-wise denominators.

#### `task_completion_long`

Recommended columns:

```text
subject_id
arm
source_file
instrument_source
session
task
task_output_name
status
task_completed
expected
missingness_reason
raw_task_completion_value
included_in_summary
```

Allowed statuses:

```text
complete
incomplete
missing
unverified
not_expected
review_required
```

Task completion is separate from behavioral QC. Task completion means the task/session item was completed or available. Behavioral QC means the completed task passed usability/QC rules.

Only rows matching expected instruments with `Included_in_summary == 1` should be used for subject-wise expectedness and group-wise denominators.

#### `behavioral_qc_long`

Recommended columns:

```text
subject_id
arm
source_file
instrument_source
session
task
task_output_name
qc_status
qc_pass
expected
missingness_reason
raw_qc_value
included_in_summary
```

Allowed QC statuses:

```text
pass
fail
missing
unverified
not_expected
review_required
```

Only rows matching expected instruments with `Included_in_summary == 1` should be used for subject-wise expectedness and group-wise denominators.

#### `mri_qc_long`

Build `mri_qc_long` from `QC_anat.csv` and `QC_func.csv`.

Recommended columns:

```text
subject_id
arm
source_file
instrument_source
session
scan_or_run
scan_or_run_output_name
qc_status
qc_pass
expected
missingness_reason
raw_qc_value
raw_mri_session
modality
taskID
runID
included_in_summary
```

Allowed QC statuses:

```text
pass
fail
missing
unverified
not_expected
review_required
```

For `QC_anat.csv` rows:

```text
source_file = mriqc_anat
instrument_source = mriqc
scan_or_run = anat_<clean modality>
```

Example:

```text
modality = T1w.html -> scan_or_run = anat_T1w
```

For `QC_func.csv` rows:

```text
source_file = mriqc_func
instrument_source = mriqc
scan_or_run = func_<taskID>_run<runID>
```

Examples:

```text
taskID = rest, runID = 01      -> func_rest_run01
taskID = selfother, runID = 02 -> func_selfother_run02
```

If MRI expectedness is represented in `Instruments_in_each_Session_each_Arm.xlsx`, only MRI rows/items with `Included_in_summary == 1` should be used for expectedness and denominators.

#### `session_long`

Recommended columns:

```text
subject_id
arm
session
session_date
session_completed
intervalFromBaseline_weeks
interval_valid
expected_target_weeks
allowed_min_weeks
allowed_max_weeks
missingness_reason
```

`session_long` should contain required post-baseline sessions for interval calculations. Do not create Baseline or Screening interval rows.

For `complete_all_experiment_sessions`, use the full required session list from `Required_Sessions_for_each_Arm.xlsx`, excluding Screening, not just the interval rows with timing-window rules.

---

## `subject_wise` sheet

Each arm-specific `subject_wise` sheet should have one row per unique `subject_id` assigned to that arm.

### Base columns

```text
subject_id
arm
dropout_status
overall_QA_status
```

Do not include:

```text
source_group
```

### Required-only and included-only columns

The subject-level sheet should include only required sessions and included instruments for that subject's arm.

Use the design files as the source of truth:

```text
Required_Sessions_for_each_Arm.xlsx
Instruments_in_each_Session_each_Arm.xlsx
```

A subject-level instrument column should be created only when:

```text
Included_in_summary == 1
```

for that arm/session/instrument in `Instruments_in_each_Session_each_Arm.xlsx`.

Do not include non-required sessions, non-required instruments, or instruments where `Included_in_summary != 1` in `subject_wise`.

### Questionnaire columns

Only for questionnaire instruments where `Included_in_summary == 1`:

```text
ses-<Session>_qn-<Questionnaire>_date
ses-<Session>_qn-<Questionnaire>_status
```

Do not include questionnaire total score columns.

Do not include:

```text
ses-<Session>_qn-<Questionnaire>_totalScore
ses-<Session>_qn-<Questionnaire>_total_score
```

Baseline questionnaire date/status columns must be included when Baseline questionnaire records are required and included for that arm:

```text
ses-Baseline_qn-<Questionnaire>_date
ses-Baseline_qn-<Questionnaire>_status
```

### Task completion columns

Only for required tasks in required sessions where `Included_in_summary == 1`:

```text
ses-<Session>_task-<Task>_status
ses-<Session>_task-<Task>_completed
```

### Behavioral QC columns

Only for required behavioral QC items in required sessions where `Included_in_summary == 1`:

```text
ses-<Session>_beh-<Task>_qc_status
ses-<Session>_beh-<Task>_qc_pass
```

### MRI QC columns

MRI QC columns should come from `QC_anat.csv` and `QC_func.csv`.

Only include MRI QC columns that are required for the subject's arm/session and have `Included_in_summary == 1` when the design files define MRI expectedness. If MRI expectedness is not defined, include observed MRI QC rows by arm and log that MRI expectedness was observation-based.

```text
ses-<Session>_mri-<ScanOrRunType>_qc_status
ses-<Session>_mri-<ScanOrRunType>_qc_pass
```

Examples:

```text
ses-Baseline_mri-anat_T1w_qc_status
ses-Baseline_mri-anat_T1w_qc_pass
ses-T6Scan_mri-func_selfother_run01_qc_status
ses-T6Scan_mri-func_selfother_run01_qc_pass
ses-T12Scan_mri-func_rest_run01_qc_status
ses-T12Scan_mri-func_rest_run01_qc_pass
```

### Session timing columns

Create session interval columns only for required post-baseline sessions:

```text
ses-<Session>_session_completed
ses-<Session>_intervalFromBaseline_weeks
ses-<Session>_interval_valid
```

Do not create:

```text
ses-Screening_intervalFromBaseline_weeks
ses-Baseline_intervalFromBaseline_weeks
```

### ASAP columns

```text
total_ASAP_count
has_ASAP
```

### Subject-level summary columns

Use these summary columns:

```text
complete_all_experiment_sessions
missing_required_experiment_sessions
all_mustHave_questionnaires_complete_perSession
all_mustHave_questionnaires_complete
all_required_task_completion_complete_perSession
all_required_task_completion_complete
all_required_MRI_QC_passed_perSession
all_required_MRI_QC_passed
all_required_selfOther_QC_passed_perSession
all_required_selfOther_QC_passed
all_required_criteria_passed_perSession
all_required_criteria_passed
ready_for_analysis
needs_followup
overall_QA_status
```

`complete_all_experiment_sessions` should mean:

```text
All required sessions for the subject's arm, excluding Screening if Screening is excluded from output, have evidence of completion.
```

It should not depend on whether timing-window rules exist for those sessions.

`missing_required_experiment_sessions` should list required sessions that are missing completion evidence for the subject.

Do not keep both old and new names for the same concept.

Use:

```text
complete_all_experiment_sessions
withdrawn_or_dropout
```

Do not use:

```text
completed_experiment_sessions
dropout
withdrawn_or_dropout plus dropout as duplicate columns
```

### Subject-wise column order

Columns should be organized in the configured time sequence when applicable and filtered to each arm's required sessions:

```text
Baseline -> Repeat Baseline -> T1 -> T2 -> T3 -> IE T3 -> T4 -> T5 -> T6 -> IE T6 -> T6 Scan -> T7 -> T8 -> T9 -> IE T9 -> T10 -> T11 -> T12 -> IE T12 -> T12 Scan
```

Do not sort sessions alphabetically.

---

## `group_wise` sheet

Each arm-specific `group_wise` sheet should summarize only subjects from that arm and only required sessions/instruments for that arm.

The sheet should use only instruments where:

```text
Included_in_summary == 1
```

The sheet should contain seven separate tables, stacked vertically with clear table titles and blank rows between them.

### Table 1. Questionnaire completion and missingness by session

Rows should be ordered by:

```text
session_order
questionnaire
instrument_source
```

Recommended columns:

```text
arm
session
questionnaire
instrument_source
expected_NofSubjects
complete_NofSubjects
incomplete_NofSubjects
missing_NofSubjects
unverified_NofSubjects
review_required_NofSubjects
complete_rate
missing_rate
```

Denominator rule:

```text
expected_NofSubjects counts only subjects expected to complete that questionnaire in that arm/session where Included_in_summary == 1.
```

### Table 2. Instrument-wise summary

This table summarizes completion/pass rate across QA domains.

Required domains:

```text
questionnaire
task_completion
behavioral_qc
mri_qc
```

Recommended columns:

```text
arm
qa_domain
instrument_or_item
instrument_source
expected_NofRecords
passed_or_complete_NofRecords
failed_or_incomplete_NofRecords
missing_NofRecords
review_required_NofRecords
passed_or_complete_rate
missing_rate
```

For MRI QC, `instrument_or_item` should use the external MRI QC `scan_or_run` name.

Denominator rule:

```text
expected_NofRecords uses only instruments/items where Included_in_summary == 1.
```

### Table 3. Task completion by session

Rows should be ordered by:

```text
session_order
task
instrument_source
```

Recommended columns:

```text
arm
session
task
instrument_source
expected_NofSubjects
complete_NofSubjects
incomplete_NofSubjects
missing_NofSubjects
unverified_NofSubjects
review_required_NofSubjects
complete_rate
missing_rate
```

Use only task-completion items where `Included_in_summary == 1`.

### Table 4. Behavioral QC pass rate by session

Rows should be ordered by:

```text
session_order
task
instrument_source
```

Recommended columns:

```text
arm
session
task
instrument_source
expected_NofSubjects
qc_pass_NofSubjects
qc_fail_NofSubjects
missing_NofSubjects
review_required_NofSubjects
qc_pass_rate
missing_rate
```

Use only behavioral-QC items where `Included_in_summary == 1`.

### Table 5. MRI QC pass rate by session

Rows should be ordered by:

```text
session_order
scan_or_run
instrument_source
```

MRI QC rows must come from `QC_anat.csv` and `QC_func.csv` using `Poor_Quality` as the flag.

Recommended columns:

```text
arm
session
scan_or_run
instrument_source
expected_NofSubjects
qc_pass_NofSubjects
qc_fail_NofSubjects
missing_NofSubjects
review_required_NofSubjects
qc_pass_rate
missing_rate
```

Use only MRI QC items where `Included_in_summary == 1` when MRI expectedness is defined by the instrument design file.

### Table 6. Session interval summary after Baseline

Rows should be ordered by:

```text
session_order
```

Do not include Screening or Baseline interval rows.

Recommended columns:

```text
arm
session
expected_NofSubjects
completed_NofSubjects
valid_interval_NofSubjects
invalid_interval_NofSubjects
missing_interval_NofSubjects
mean_intervalFromBaseline_weeks
sd_intervalFromBaseline_weeks
min_intervalFromBaseline_weeks
max_intervalFromBaseline_weeks
```

### Table 7. Participant-level QA readiness summary

This table should not duplicate items.

Count each readiness/status category once.

Recommended rows or columns should include:

```text
arm
total_NofSubjects
withdrawn_or_dropout_NofSubjects
complete_all_experiment_sessions_NofSubjects
all_mustHave_questionnaires_complete_NofSubjects
all_required_task_completion_complete_NofSubjects
all_required_MRI_QC_passed_NofSubjects
all_required_selfOther_QC_passed_NofSubjects
all_required_criteria_passed_NofSubjects
ready_for_analysis_NofSubjects
needs_followup_NofSubjects
```

Do not double-count `ready_for_analysis` or `needs_followup` both as booleans and again from `overall_QA_status`.

---

## Output files

The script should generate output workbooks per arm:

```text
QA_summary_arm1.xlsx
QA_summary_arm2.xlsx
QA_summary_arm3.xlsx
```

Each workbook should contain:

```text
subject_wise
group_wise
```

Optional combined output may be added only as an extra file, not as a replacement for arm-specific files:

```text
QA_summary_all_arms.xlsx
```

The arm-specific files are required.

---

## Validation requirements

The script should fail clearly or log warnings for data problems that could affect interpretation.

Required validations:

### Design files present

Validate that these files exist:

```text
Required_Sessions_for_each_Arm.xlsx
Instruments_in_each_Session_each_Arm.xlsx
```

If either is missing, raise a clear error unless an explicit fallback config is provided.

### Instrument inclusion column present

Validate that `Instruments_in_each_Session_each_Arm.xlsx` contains an inclusion column that normalizes to:

```text
Included_in_summary
```

or:

```text
included_in_summary
```

If the inclusion column is missing, raise a clear error by default because instrument expectedness depends on this field.

The validation log should report:

```text
included_in_summary_design_rows
excluded_by_included_in_summary_design_rows
```

### Instrument inclusion filter applied

Validate that no final `subject_wise` columns and no group-wise denominators are created from instrument-design rows where:

```text
Included_in_summary != 1
```

This should apply to questionnaire, task completion, behavioral QC, and MRI QC expectedness.

### Subject ID source columns valid

Validate that each input file contains the expected subject ID source column:

```text
ParticipantsQAtracker.csv -> record_id
ClinicianQAtracker.csv    -> preescreen_id
QC_anat.csv               -> subID
QC_func.csv               -> subID
```

For MRI QC files, validate that `subID` values can be standardized by removing the `sub-` prefix and keeping the three digits after it.

Log any missing, malformed, duplicated, or conflicting subject IDs before merging.

### Arm assignment valid

Validate that every subject included in final outputs has an inferred arm:

```text
arm1
arm2
arm3
```

Arm should be inferred from the first digit of the standardized three-digit `subject_id`.

Subjects with unknown arms should be logged and excluded from arm-specific outputs unless a debug option is enabled.

### Required-only subject-level columns

Validate that each arm-specific `subject_wise` sheet contains only sessions required for that arm and only instruments included for that arm/session.

Observed but non-required or non-included records should be logged as unexpected observations.

### `complete_all_experiment_sessions` source validation

Validate that `complete_all_experiment_sessions` is calculated from `Required_Sessions_for_each_Arm.xlsx`, not from timing-window or interval-validity filtering.

The logic should include required sessions even if no `session_timing_windows` rule exists.

The validation log should include, for each arm:

```text
required_sessions_checked_for_complete_all_experiment_sessions_<arm>
```

Recommended diagnostic column:

```text
missing_required_experiment_sessions
```

### No Screening in final output

```python
if any("Screening" in str(col) for col in subject_wise.columns):
    raise ValueError("Screening columns should not appear in subject_wise.")
```

Also validate every group-wise table contains no Screening values.

### No total score columns

```python
total_score_cols = [
    col for col in subject_wise.columns
    if "totalscore" in col.lower() or "total_score" in col.lower()
]
if total_score_cols:
    raise ValueError(f"Total score columns should not be in final output: {total_score_cols}")
```

### Baseline questionnaire columns present when required and included

If Baseline questionnaires are required for an arm and have `Included_in_summary == 1`, that arm's `subject_wise` sheet should contain date/status columns beginning with:

```text
ses-Baseline_qn-
```

### MRI QC files present

If `QC_anat.csv` or `QC_func.csv` is missing, log a warning.

If both files are missing, `mri_qc_long` may be empty, but the validation log must clearly state that no external MRI QC files were loaded.

### MRI QC required columns present

For `QC_anat.csv`, validate:

```text
subID
sesID
modality
Poor_Quality
```

For `QC_func.csv`, validate:

```text
subID
sesID
modality
taskID
runID
Poor_Quality
```

If required columns are missing, log a clear warning and do not silently create misleading MRI QC summaries.

### MRI QC session mapping valid

Validate that all observed `sesID` values in external MRI QC files either map to canonical sessions or are logged as unmapped.

Final outputs should not contain:

```text
UNMAPPED_MRI_SESSION_
```

unless a debug option is explicitly enabled.

### Final-output integrity checks

Before saving final workbooks, validate:

- `subject_wise` has one row per subject per arm.
- `subject_wise` does not contain duplicate columns.
- Final outputs do not contain REDCap internal event strings such as `_arm_`.
- Final outputs do not contain `UNMAPPED_EVENT_` unless an explicit debug flag allows it.
- Final outputs do not contain `UNMAPPED_MRI_SESSION_` unless an explicit debug flag allows it.
- Final outputs do not contain `Screening`.
- Final outputs do not contain `Baseline 1` or `Baseline1`.
- `subject_wise` contains Baseline questionnaire date/status columns when Baseline questionnaire records are required and included.
- `subject_wise` does not contain total score columns.
- `subject_wise` does not contain `completed_experiment_sessions`.
- `subject_wise` contains `complete_all_experiment_sessions`.
- `complete_all_experiment_sessions` was calculated against required sessions from `Required_Sessions_for_each_Arm.xlsx`.
- `subject_wise` contains `missing_required_experiment_sessions`.
- `ready_for_analysis`, `needs_followup`, and `overall_QA_status` each appear only once.
- `dropout_status` exists and is the only subject-level dropout/withdrawal status column.
- `group_wise` contains seven separate tables.
- `group_wise` contains the instrument-wise summary table.
- `group_wise` contains the task completion table.
- Participant-level QA readiness summary has no duplicated `status_metric` rows.
- Participant-level QA readiness summary does not contain `dropout`.
- Participant-level QA readiness summary does not contain `completed_experiment_sessions`.
- No interval columns are created for Screening or Baseline.
- No subject-wise instrument columns are created for rows where `Included_in_summary != 1`.

---

## Coding standards

- Keep the script configuration-first.
- Keep the script design-file-driven.
- Keep helper functions small and testable.
- Use readable function names.
- Use stable column names in final Excel output.
- Read raw CSV and Excel design values as strings.
- Log assumptions, missing fields, unknown status codes, excluded design rows, and unmapped events.
- Raise errors for final-output violations instead of silently producing misleading workbooks.
- Prefer explicit mappings over fragile inference.
- Do not overwrite same-subject records from different source files.
- Do not mix multiple unrelated group summaries into one wide dataframe.
- Keep expectedness logic separate from observed-data parsing.
- Apply `Included_in_summary == 1` before constructing expected instruments, subject-level columns, or group-level denominators.

---

## Known assumptions

- The two QA tracker CSVs and the two MRI QC CSVs may contain overlapping standardized `subject_id` values and should be merged by subject.
- In `ParticipantsQAtracker.csv`, `subject_id` comes from `record_id`.
- In `ClinicianQAtracker.csv`, `subject_id` comes from `preescreen_id`.
- In `QC_anat.csv` and `QC_func.csv`, `subject_id` comes from `subID`; remove the `sub-` prefix and use the three digits after it.
- Across `ParticipantsQAtracker.csv`, `ClinicianQAtracker.csv`, `QC_anat.csv`, and `QC_func.csv`, records with the same standardized `subject_id` belong to the same subject.
- `source_file` and `instrument_source` are internal metadata, not subject-level grouping variables.
- Subject arm is inferred from the first digit of the standardized three-digit `subject_id`.
- `Required_Sessions_for_each_Arm.xlsx` is the source of truth for required sessions by arm.
- `Instruments_in_each_Session_each_Arm.xlsx` is the source of truth for candidate instruments by arm/session.
- `Included_in_summary == 1` is required for an instrument to be treated as expected, included in `subject_wise`, or counted in group-level denominators.
- REDCap `_complete` fields usually use `2 = complete`, `1 = unverified`, and `0 = incomplete`.
- Task completion and behavioral QC are separate QA domains.
- Baseline is the anchor for post-baseline session intervals.
- Screening is not needed in the final QA workbook.
- Questionnaire total score columns are not needed in the final QA workbook.
- MRI QC should be read from `QC_anat.csv` and `QC_func.csv`, with `Poor_Quality` as the fail/pass flag.
- `ses-T6` and `ses-T12` in MRI QC files should map to `T6 Scan` and `T12 Scan`, respectively.
- `complete_all_experiment_sessions` is a boolean indicating whether all required experiment sessions from `Required_Sessions_for_each_Arm.xlsx` are complete for that subject's arm.
- `complete_all_experiment_sessions` should not be calculated from `interval_valid != "not_expected"` because sessions without interval-window rules may otherwise be missed.
