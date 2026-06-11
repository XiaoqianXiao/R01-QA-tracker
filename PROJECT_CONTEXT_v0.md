Goal: 
    Create one Python script: qa_summary.py.
    The script should read REDCap QA exports and REDCap codebooks, reformat subject-level QA information, and generate a project-level QA summary workbook.

Inputs:
    - REDCap QA exports:
        - ParticipantsQAtracker.csv
        - ClinicianQAtracker.csv
    - REDCap codebooks:
        - Participants_REDCap.pdf
        - Clini_REDCap.pdf

Outputs: QA_summary.xlsx with two sheets
    - subject_wise: One row per subject. Each row should summarize that subject’s questionnaire completion, behavioral QA, MRI QA, session timing, and subject-level eligibility/completeness status.
        - description of status. 
            - Recommended column pattern:
                - subject_id
                - ses-<Session>_qn-<Questionnaire>_date
                - ses-<Session>_qn-<Questionnaire>_status
                - ses-<Session>_qn-<Questionnaire>_totalScore
            - Status Definitions:
                - complete: Required fields or REDCap completion status indicate that the questionnaire/session item is complete.
                - incomplete: The record exists, but completion status or required fields indicate it is not complete.
                - missing: The questionnaire/session is expected for the subject, but no usable record or date is present. 
                - unverified: The record exists, but QA confirmation/checking field is missing or unclear. 
                - not_expected: The questionnaire/session is not expected for that subject based on study design, withdrawal, group, or session schedule.

        - added subject level summary
            - session interval (unit: weeks), For each session after baseline, calculate the interval from baseline:
                - ses-<Session>_intervalFromBaseline_weeks
            - number of total ASAP
        - Subject-level overall summary columns:
            - dropout_status
            - completed_experiment_sessions
            - all_mustHave_questionnaires_complete_perSession
            - all_mustHave_questionnaires_complete
            - all_required_MRI_QC_passed_perSession
            - all_required_MRI_QC_passed
            - all_required_selfOther_QC_passed_perSession
            - all_required_selfOther_QC_passed
            - all_required_criteria_passed_perSession
            - all_required_criteria_passed
            - overall_QA_status

    - group_wise
        - Session-by-questionnaire summary: sumamry for each instrument each session across subjects
            - source_file
            - instrument_type
            - session
            - questionnaire
            - expected_NofSubjects
            - complete_NofSubjects
            - missing_NofSubjects
            - incomplete_NofSubjects
            - unverified_NofSubjects
            - complete_rate
            - missing_rate
            - qc_pass_NofSubjects
            - qc_fail_NofSubjects
            - qc_pass_rate

        - Instrument-level summary across sessions: summary for each instrument 
            - source_file
            - instrument_type
            - questionnaire
            - total_number_of_sessions
            - expected_NofSubjects
            - complete_NofSubjects
            - missing_rate

        - summary session interval (unit: weeks; from each session to baseline)
            - expected_target_weeks
            - allowed_min_weeks
            - allowed_max_weeks
            - mean and std for group-wise interval from setted standard
            - number and rate of number of subjects meet the valid for each session
            - number and rate of number of subjects meet the valid for all sessions
            - early_NofSubjects
            - early_rate
            - late_NofSubjects
            - late_rate

        - summary for ASAP
            -  how many of subject has ASAP, range
            - subjects_with_0_ASAP 
            - subjects_with_1_ASAP 
            - subjects_with_2plus_ASAP
        
        - summary for paticipants status, group level number and ration for:
            - drop-out
            - complete the experiment session
            - with all sessions meet all cretia:
                - 1. Must-have questionnaires are complete.
                - 2. Required MRI QC passed.
                - 3. Required self/other behavioral QC passed.
                - 4. Required sessions are completed.
                - 5. Subject is not marked as withdrawn/dropout unless project rules say otherwise.
            - ready_for_analysis
            - needs_followup

Coding rule:
 - Do not use REDCap internal unique event names such as baseline_1_arm_1 in the final Excel output. Use the human-readable Event Name from the codebook section "Events".

 Data Processing Logic:
 - Step 1: Read input files.
    All raw REDCap fields should initially be read as strings to avoid losing ID formatting or coded values
 - Step 2: Parse or manually map codebook information.
    Use the PDFs to create mappings for:
        - REDCap event unique name -> human-readable Event Name
        - form/instrument name -> readable instrument label
        - field name -> field label
        - coded choices -> readable values
        - completion status fields
        - total score fields
        - date fields
        - QA/QC status fields
    If PDF parsing is unreliable, create a manual configuration block inside qa_summary.py.
 - Step 3: Standardize subject IDs
    Create one standardized subject ID field: subject_id.
    The script should handle possible variations such as:
        - record_id
        - participant_id
        - subject_id
        - study_id
    The exact source field should be determined from the REDCap exports/codebook.
 - Step 4: Standardize events and sessions
    Use the codebook Event Name rather than REDCap internal event name.
    For example:
        - baseline_1_arm_1 -> Baseline
        - t1_arm_1 -> T1
    The final Excel output should use readable session names.
 - Step 5: Standardize status values
 - Step 6: Build subject-wise table
    For each subject:
        1. Gather all rows/forms across both REDCap exports.
        2. Identify expected sessions and expected questionnaires.
        3. Fill questionnaire date/status/totalScore columns.
        4. Fill behavioral QC columns.
        5. Fill MRI QC columns.
        6. Calculate session intervals from baseline.
        7. Count ASAP flags.
        8. Derive subject-level summary columns.
 - Step 7: Build group-wise summary tables
    Using the subject-wise table and expected-session configuration:
        1. Compute session-by-questionnaire missingness and completion rates.
        2. Compute instrument-level summaries across sessions.
        3. Compute MRI QC pass/fail rates.
        4. Compute self/other behavioral QC pass/fail rates.
        5. Compute session interval summaries.
        6. Compute ASAP summaries.
        7. Compute participant status summaries.
 - Step 8: Write Excel output
    Write to: QA_summary.xlsx
    Sheets:
        subject_wise
        group_wise
    Use openpyxl or xlsxwriter.
    Recommended formatting:
        - Freeze first row
        - Apply autofilter
        - Auto-adjust column widths
        - Bold header row
        - Use clear section headers in group_wise
        - Keep all percentage/rate columns numeric, not strings
 - Step 9: Add validation checks
    Before saving the final workbook, print or log:
        - Number of subjects in ParticipantsQAtracker.csv
        - Number of subjects in ClinicianQAtracker.csv
        - Number of unique subject IDs in final subject_wise table
        - Number of sessions detected
        - Number of questionnaires detected
        - Number of missing baseline dates
        - Number of subjects with missing required MRI QC
        - Number of subjects with missing required self/other QC
    Also create warning messages for:
        - Unknown REDCap event names
        - Unknown questionnaire/status codes
        - Duplicate subject-session-questionnaire records
        - Missing subject IDs
        - Missing baseline dates
        - Date parsing failures
        - Expected fields not found in CSV

