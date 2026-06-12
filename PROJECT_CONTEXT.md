# PROJECT_CONTEXT.md — R01-QA-tracker

## 1. Project goal

Create `qa_summary.py` to read REDCap QA exports, study-design Excel files, MRI QC CSV files, and REDCap codebooks, then generate arm-specific QA summary workbooks.

Required output files:

```text
QA_summary_arm1.xlsx
QA_summary_arm2.xlsx
QA_summary_arm3.xlsx
```

Each workbook must contain two sheets:

```text
subject_wise
group_wise
```

The script should also create a validation log:

```text
QA_summary_validation_log.txt
```

or one log per arm:

```text
QA_summary_arm1_validation_log.txt
QA_summary_arm2_validation_log.txt
QA_summary_arm3_validation_log.txt
```

The goal is to reduce manual REDCap checking and make QA status clear by subject and by arm.

---

## 2. Input files

The script should use these files:

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

### Main role of each file

| File | Role |
|---|---|
| `ParticipantsQAtracker.csv` | Participant REDCap QA records |
| `ClinicianQAtracker.csv` | Clinician REDCap QA records |
| `Required_Sessions_for_each_Arm.xlsx` | Required sessions by arm |
| `Instruments_in_each_Session_each_Arm.xlsx` | Expected instruments by arm/session |
| `QC_anat.csv` | Anatomical MRI QC |
| `QC_func.csv` | Functional MRI QC |
| `Participants_REDCap.pdf` | Participant REDCap codebook |
| `Clini_REDCap.pdf` | Clinician REDCap codebook |

---

## 3. Subject ID and arm rules

Create one standardized `subject_id` before merging any data.

| File | Raw subject ID column | Standardization rule |
|---|---|---|
| `ParticipantsQAtracker.csv` | `record_id` | Use as the three-digit `subject_id` |
| `ClinicianQAtracker.csv` | `preescreen_id` | Use as the three-digit `subject_id` |
| `QC_anat.csv` | `subID` | Remove `sub-`; keep the three digits after it |
| `QC_func.csv` | `subID` | Remove `sub-`; keep the three digits after it |

Examples:

```text
record_id = 101      -> subject_id = 101 -> arm1
preescreen_id = 205  -> subject_id = 205 -> arm2
subID = sub-312      -> subject_id = 312 -> arm3
```

Arm assignment:

```text
subject_id starts with 1 -> arm1
subject_id starts with 2 -> arm2
subject_id starts with 3 -> arm3
```

Records with the same standardized `subject_id` across the participant export, clinician export, and MRI QC files belong to the same subject and must be merged into one subject-level row.

Do not use `source_group` in the final output.

It is okay to keep these internally for debugging:

```text
source_file
instrument_source
```

---

## 4. Required sessions and expected instruments

The output must be design-file-driven.

Use:

```text
Required_Sessions_for_each_Arm.xlsx
```

to decide which sessions are required for each arm.

Use:

```text
Instruments_in_each_Session_each_Arm.xlsx
```

to decide which instruments are expected for each arm/session.

Critical rule:

```text
Only instruments with Included_in_summary == 1 are expected and included.
```

Rows with blank, `0`, `no`, `false`, or any value other than `1` in `Included_in_summary` must be excluded from:

- `subject_wise` columns
- group-level denominators
- subject-level completeness checks
- readiness checks

Do not infer required sessions or expected instruments from observed REDCap rows alone.

---

## 5. Event and session naming

Final Excel outputs must use readable session names.

Do not show REDCap internal event names such as:

```text
baseline_1_arm_1
t1_arm_1
```

Use the human-readable `Event Name` from the REDCap codebook `Events` section.

Examples:

```text
baseline_1_arm_1 -> Baseline
t1_arm_1         -> T1
t6_scan_arm_1   -> T6 Scan
```

Final session labels should use this order:

```text
Baseline
Repeat Baseline
T1
T2
T3
IE T3
T4
T5
T6
IE T6
T6 Scan
T7
T8
T9
IE T9
T10
T11
T12
IE T12
T12 Scan
```

Do not use these final labels:

```text
Baseline 1
Baseline1
```

Screening may be recognized internally, but it must not appear in the final Excel output.

---

## 6. MRI QC rules

MRI QC must come from external CSV files, not REDCap QA tracker columns.

Default paths:

```text
data/QC_anat.csv
data/QC_func.csv
```

### `QC_anat.csv` required columns

```text
subID
sesID
modality
Poor_Quality
```

### `QC_func.csv` required columns

```text
subID
sesID
modality
taskID
runID
Poor_Quality
```

### MRI QC pass/fail rule

```text
Poor_Quality == True  -> qc_status = fail, qc_pass = False
Poor_Quality == False -> qc_status = pass, qc_pass = True
```

Accept these true values:

```text
True, true, TRUE, 1, yes, y
```

Accept these false values:

```text
False, false, FALSE, 0, no, n
```

Missing or unclear values should become:

```text
qc_status = review_required
qc_pass = False
missingness_reason = review_required
```

### MRI session mapping

Map MRI `sesID` values to canonical session names.

Important mappings:

```text
ses-baseline -> Baseline
ses-repeatbaseline -> Repeat Baseline
ses-T6 -> T6 Scan
ses-T12 -> T12 Scan
```


Recommended scan/run names:

```text
anat_T1w
func_rest_run1
func_rest_run2
func_selfother_run1
func_selfother_run2
```

---

## 7. Internal processing strategy

Build long-format internal tables first, then create the Excel outputs.

Recommended internal tables:

```text
subject_arm_map
expected_sessions_by_arm
expected_instruments_by_arm_session
questionnaire_long
behavioral_qc_long
mri_qc_long
session_long
```

Read raw files as strings when possible:

```python
pd.read_csv(..., dtype=str, keep_default_na=False)
pd.read_excel(..., dtype=str, keep_default_na=False)
```

This avoids losing subject ID formatting, REDCap codes, and raw date values.

---

## 8. Subject-wise output

Each `subject_wise` sheet should have one row per subject in that arm.

Base columns:

```text
subject_id
arm
dropout_status
```

### Include only required and included items

For each subject:

1. infer arm from `subject_id`
2. get expected sessions for that arm
3. get expected instruments for each expected session
4. keep only instruments with `Included_in_summary == 1`
5. create columns only for those sessions and instruments
6. fill observed values when present
7. mark missing expected records as `missing`
8. log observed-but-not-expected records

### Questionnaire columns

Use:

```text
ses-<Session>_qn-<Questionnaire>_date
ses-<Session>_qn-<Questionnaire>_status # complete, incomplete or missing
```

### Behavioral QC columns

Use:
```text
two tasks: ANT and self_others
for ANT use: ses-<Session>_beh-ANT_status #complete, incomplete, or missing
for self_others use: 
	- ses-<Session>_beh-selfOthers_Npractice #number of practices before passing the threshold
	- ses-<Session>_beh-selfOthers_run1_status #complete, incomplete, or missing
	- ses-<Session>_beh-selfOthers_run1_acc #
	- ses-<Session>_beh-selfOthers_run1_missingRate #
	- ses-<Session>_beh-selfOthers_run2_status #complete, incomplete, or missing
	- ses-<Session>_beh-selfOthers_run2_acc #
	- ses-<Session>_beh-selfOthers_run2_missingRate #
```

Source
```text
self other task: ClinicianQAtracker.csv, scan_run_ (hold this for REDCap updates)
ANT task: ClinicianQAtracker.csv
bl_during_other_2 (this is for Baseline)
t3t9_checklist_other (this is for T3, T9)
t6t12_checklist_other_visit_2 (this is for T6, T12)
t12_other_2 (this is for T12)
```
### MRI QC columns

Use:

```text
ses-<Session>_mri-<ScanOrRun>_qc_status
```

Examples:

```text
ses-Baseline_mri-anat_T1w_qc_status
ses-T6Scan_mri-func_selfother_run01_qc_status
```

### Session columns

Only create timing interval columns for required post-baseline sessions.

Use:

```text
ses-<Session>_session_completed #only true if all requried instrument and behavioral task have been done for that session
ses-<Session>_missingTask #list missing required instrument or task for that session 
ses-<Session>_selfOtherQC_passed #only true if all selfOther QC passed for that session
ses-<Session>_selfOtherQC_failed #list run of selfOther task which failed the QC cretira for that session
ses-<Session>_MRIQC_passed #only true if all MRI QC passed for that session
ses-<Session>_MRIQC_failed #list MRI which failed the QC cretira for that session
ses-<Session>_QC_pass #only true if session complete and all QC passed
ses-<Session>_intervalFromBaseline_weeks
ses-<Session>_interval_valid
```

Do not create interval columns for:

```text
Screening
Baseline
```

### ASAP columns

Use:

```text
total_ASAP_count
```


### Subject-level summary columns

Use these columns:

```text
complete_all_expected_experiment_sessions #true if all requried session marked as complete
Nof_missing_expected_experiment_sessions #number of missing or incompleted expected sessions
missing_expected_experiment_sessions #list missing or incompleted expected sessions
complete_all_instrument
complete_all_ANT
all_MRI_QC_passed #true if all required sessions' MRI QC passed
all_selfOther_QC_passed #true if all required sessions' selfOther QC passed
subject_QC_pass #only true if all sessions complete and all QC passed
```

---

## 9. `complete_all_experiment_sessions` rule

This field must be based on expected sessions.

Correct meaning:

```text
complete_all_experiment_sessions = True only if every expected experiment session for that subject's arm has completion evidence.
```

Use:

```text
Required_Sessions_for_each_Arm.xlsx
```

as the source of expected sessions.

Exclude Screening and ASAP

---

## 10. Group-wise output

Each arm-specific `group_wise` sheet should summarize only subjects from that arm.

Use only expected sessions and instruments with:

```text
Included_in_summary == 1
```

The sheet should contain seven separate tables stacked vertically, with clear titles and blank rows between tables.

### Table 1. Session-wise Summary

Columns:

```text
arm
session
instrument
instrument_source
expected_NofSubjects
complete_NofSubjects
review_required_NofSubjects
complete_rate
missing_rate
```

Order rows by:

```text
session_order
questionnaire
instrument_source
```

### Table 2. Instrument-wise Summary
Columns:

```
arm
instrument_or_item
instrument_source
expected_NofSubjects
complete_NofRecords #only subject completed all required sessions for that instrument
review_required_NofRecords
complete_rate
missing_rate
```


### Table 3-1. ANT task complete rate by session

Columns:

```text
arm
session
expected_NofSubjects
complete_NofSubjects
missing_NofSubjects
review_required_NofSubjects
complete_rate
missing_rate
```

### Table 3-2. ANT task complete rate by subject

Columns:

```text
arm
expected_NofSubjects
complete_NofSubjects
missing_NofSubjects
review_required_NofSubjects
complete_rate
missing_rate
```
### Table 4-1. SelfOthers QC pass rate by session

Columns:

```text
arm
session
expected_NofSubjects
qc_pass_NofSubjects
qc_fail_NofSubjects
missing_NofSubjects
review_required_NofSubjects
qc_pass_rate
missing_rate
```
### Table 4-2. SelfOthers QC pass rate by subject

Columns:

```text
arm
expected_NofSubjects
qc_pass_NofSubjects
qc_fail_NofSubjects
missing_NofSubjects
review_required_NofSubjects
qc_pass_rate
missing_rate
```
### Table 5. MRI QC pass rate by session

Columns:

```text
arm
session
expected_NofSubjects
qc_pass_NofSubjects #only subject with all scan runs
qc_fail_NofSubjects
missing_NofSubjects
review_required_NofSubjects
qc_pass_rate
missing_rate
```
### Table 5-2. MRI QC pass rate by subject

Columns:

```text
arm
scan_or_run
expected_NofSubjects
qc_pass_NofSubjects # only subject with all sessions passed
qc_fail_NofSubjects
missing_NofSubjects
review_required_NofSubjects
qc_pass_rate
missing_rate
```


### Table 6. Session interval summary after Baseline

Do not include Screening or Baseline.

Columns:

```text
arm
session
completed_NofSubjects
valid_interval_NofSubjects
invalid_interval_NofSubjects
mean_intervalFromBaseline_weeks
sd_intervalFromBaseline_weeks
min_intervalFromBaseline_weeks
max_intervalFromBaseline_weeks
```

### Table 7. Participant-level QA readiness summary
Columns:

```text
arm
total_NofSubjects
withdrawn_or_dropout_NofSubjects
withdrawn_or_dropout_rate
complete_all_experiment_sessions_NofSubjects
complete_all_experiment_sessions_rate
complete_all_instrument_NofSubjects
complete_all_instrument_rate
complete_all_ANT_NofSubjects
complete_all_ANT_rate
all_MRI_QC_passed_NofSubjects
all_MRI_QC_passed_rate
all_selfOther_QC_passed_NofSubjects
all_selfOther_QC_passed_rate
QC_pass_NofSubjects_rate
QC_passrate
```

---

## 11. Denominator rules

All rates must use expected records as the denominator.

Expected means:

1. the session is required for that arm
2. the instrument is listed for that arm/session
3. `Included_in_summary == 1`

Examples:

```text
complete_rate = complete_NofSubjects / expected_NofSubjects
missing_rate = missing_NofSubjects / expected_NofSubjects
qc_pass_rate = qc_pass_NofSubjects / expected_NofSubjects
```

Do not count `not_expected` records in denominators.

Do not count instruments with `Included_in_summary != 1` in denominators.

---

## 12. Validation requirements

The script should fail clearly or log warnings when important inputs or mappings are missing.

Validate these points:

1. Required design files exist.
2. `Instruments_in_each_Session_each_Arm.xlsx` contains `Included_in_summary` or a normalized equivalent.
3. The `Included_in_summary == 1` filter is applied before creating subject columns or group denominators.
4. Each file contains its expected subject ID column.
5. Subject IDs can be standardized to three digits.
6. Arm can be inferred as `arm1`, `arm2`, or `arm3`.
7. Final outputs contain only required sessions and included instruments for that arm.
8. Observed-but-not-expected records are logged.
9. `complete_all_experiment_sessions` uses `Required_Sessions_for_each_Arm.xlsx`, not interval validity.
10. Screening does not appear in final outputs.
11. Total score columns do not appear in final outputs.
12. Baseline questionnaire date/status columns appear when Baseline questionnaires are required and included.
13. MRI QC files are found or clearly logged as missing.
14. MRI QC files contain required columns.
15. MRI `sesID` values map to canonical sessions.
16. Final outputs do not contain REDCap internal event names such as `_arm_`.
17. Final outputs do not contain `UNMAPPED_EVENT_` or `UNMAPPED_MRI_SESSION_` unless debug mode is explicitly enabled.
18. `subject_wise` has one row per subject per arm.
19. `subject_wise` has no duplicate columns.
20. `group_wise` contains all seven required tables.
21. Participant-level readiness summary has no duplicated status rows.

---

## 13. Coding standards

- Keep the script configuration-first.
- Keep the script design-file-driven.
- Keep expectedness logic separate from observed-data parsing.
- Use small, testable helper functions.
- Use readable function names.
- Read CSV and Excel files as strings when possible.
- Use stable final Excel column names.
- Log assumptions, missing fields, unknown status codes, excluded design rows, and unmapped events.
- Raise errors for final-output violations instead of silently producing misleading workbooks.
- Prefer explicit mappings over fragile inference.
- Do not overwrite same-subject records from different source files.
- Do not combine unrelated group summaries into one wide dataframe.

---

## 14. Known assumptions
- REDCap `_complete` fields usually mean: `2 = complete`, `1 = unverified`, `0 = incomplete`.
- Task completion and behavioral QC are separate QA domains.
- Baseline is the anchor for post-baseline session intervals.
- `complete_all_experiment_sessions` means all required non-screening and non-asap sessions are complete.
