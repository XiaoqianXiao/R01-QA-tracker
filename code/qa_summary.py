#!/usr/bin/env python3
"""Build arm-specific QA summary workbooks from REDCap, design, and MRI QC files."""

from __future__ import annotations

import argparse
import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import pandas as pd


SESSION_ORDER = [
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
]

ARMS = ["arm1", "arm2", "arm3"]
EXCLUDED_FINAL_SESSIONS = {"Screening"}
EXCLUDED_COMPLETENESS_SESSIONS = {"Screening"}
TRUE_VALUES = {"true", "1", "yes", "y"}
FALSE_VALUES = {"false", "0", "no", "n"}

INPUT_FILES = {
    "participants": "ParticipantsQAtracker.csv",
    "clinician": "ClinicianQAtracker.csv",
}

DESIGN_FILES = {
    "required_sessions": "Required_Sessions_for_each_Arm.xlsx",
    "expected_instruments": "Instruments_in_each_Session_each_Arm.xlsx",
}

MRI_FILES = {
    "anat": "QC_anat.csv",
    "func": "QC_func.csv",
}

PARTICIPANT_INSTRUMENTS: dict[str, dict[str, str | list[str]]] = {
    "PHQ-9": {"date": "phq_9_date", "complete": "phq9_complete"},
    "GAD-7": {"date": "gad_7_date", "complete": "gad7_complete"},
    "LSAS-SR": {"date": "lsas_sr_date", "complete": "lsas_sr_complete"},
    "BDD-YBOCS-SR": {"date": "bddybocs_sr_date", "complete": "bddybocs_sr_complete"},
    "Demographics Form": {"date": "demo_date", "complete": "demographics_form_complete"},
    "SCS-R": {"date": "scs_date", "complete": "scsr_complete"},
    "RRS": {"date": "rrs_date", "complete": "rrs_complete"},
    "PSWQ": {"date": "pswq_date", "complete": "pswq_complete"},
    "QIDS": {"date": "qids_sr_date", "complete": "qids_complete"},
    "SDS": {"date": "sds_date", "complete": "sds_complete"},
    "CEQ": {"date": "ceq_date", "complete": "ceq_complete"},
    "WHODAS": {"date": "whodas_date", "complete": "whodas_complete"},
    "DSM5 Cross Cutting": {"date": "dsm5cc_date", "complete": "dsm5_cross_cutting_complete"},
    "MRI Exit Questionnaire": {"date": "mri_exit_date", "complete": "mri_exit_questionnaire_complete"},
}

CLINICIAN_INSTRUMENTS: dict[str, dict[str, str | list[str]]] = {
    "SCID-5": {"date": "scid5_axis1_date", "complete": "scid5_complete"},
    "SCID-5 T12": {"date": "scid5_axis1_date", "complete": "scid5_complete"},
    "WRAT-5": {"date": "wrat5_date", "complete": "wrat5_complete"},
    "Diagnostic Summary": {"date": "diagnostic_summary_date", "complete": "diagnostic_summary_complete"},
    "Eligibility Form": {"date": "eligibility_date", "complete": "eligibility_form_complete"},
    "PEAS": {"date": "peas_date_of_visit", "complete": "peas_complete"},
    "Life Changes Form": {"date": "lcf", "complete": "life_changes_form_complete"},
    "LSAS - Clinician Rated": {"date": "lsas_cr_date", "complete": "lsas_clinician_rated_complete"},
    "BDD-YBOCS - Clinician Rated": {
        "date": "bddybocs_cr_date",
        "complete": "bddybocs_clinician_rated_complete",
    },
    "BABS": {"date": "babs_date", "complete": "babs_complete"},
    "SIGH-A": {"date": "sigh_a_date", "complete": "sigha_complete"},
    "BDD Data Form": {"complete": "bdd_data_form_complete"},
    "Scan Run Sheet": {"date": "scan_date", "complete": "scan_run_sheet_complete"},
    "Baseline/Repeat Baseline Visit Checklist": {"complete": "bl_during_other_2"},
    "Treatment Visit Checklist": {"complete": ["t3t9_checklist_other", "t6t12_checklist_other_visit_2", "t12_other_2"]},
    "T3/T9 IE Visit Checklist": {"complete": "t3t9_checklist_other"},
    "T6/T12 IE Visit Checklist": {"complete": ["t6t12_checklist_other_visit_2", "t12_other_2"]},
}

EVENT_OVERRIDES = {
    "baseline_1": "Baseline",
    "baseline": "Baseline",
    "repeat_baseline": "Repeat Baseline",
    "screening": "Screening",
}

MRI_SESSION_MAP = {
    "ses-baseline": "Baseline",
    "ses-repeatbaseline": "Repeat Baseline",
    "ses-T1": "T1",
    "ses-T2": "T2",
    "ses-T3": "T3",
    "ses-T4": "T4",
    "ses-T5": "T5",
    "ses-T6": "T6 Scan",
    "ses-T7": "T7",
    "ses-T8": "T8",
    "ses-T9": "T9",
    "ses-T10": "T10",
    "ses-T11": "T11",
    "ses-T12": "T12 Scan",
}


@dataclass
class ValidationLog:
    info_lines: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    counters: dict[str, Any] = field(default_factory=dict)

    def info(self, message: str) -> None:
        self.info_lines.append(message)

    def warn(self, message: str) -> None:
        self.warnings.append(message)

    def count(self, key: str, value: Any) -> None:
        self.counters[key] = value

    def write(self, path: Path) -> None:
        lines = ["QA Summary Validation Log", "=" * 25, ""]
        lines.extend(self.info_lines)
        lines.extend(["", "Counters", "-" * 8])
        for key in sorted(self.counters):
            lines.append(f"{key}: {self.counters[key]}")
        lines.extend(["", "Warnings", "-" * 8])
        lines.extend(self.warnings or ["No warnings recorded."])
        path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def clean(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def slug(text: str) -> str:
    value = clean(text)
    value = value.replace("/", "_").replace(" ", "")
    value = re.sub(r"[^0-9A-Za-z_-]+", "", value)
    return value


def normalize_column(name: str) -> str:
    value = clean(name).lower()
    value = re.sub(r"[^0-9a-z]+", "_", value)
    return re.sub(r"_+", "_", value).strip("_")


def canonical_session(value: str) -> str:
    text = clean(value)
    aliases = {
        "Baseline 1": "Baseline",
        "Baseline1": "Baseline",
        "Treatment Session 1": "T1",
        "Treatment Session 2": "T2",
        "Treatment Session 3": "T3",
        "Treatment Session 4": "T4",
        "Treatment Session 5": "T5",
        "Treatment Session 6": "T6",
        "Treatment Session 7": "T7",
        "Treatment Session 8": "T8",
        "Treatment Session 9": "T9",
        "Treatment Session 10": "T10",
        "Treatment Session 11": "T11",
        "Treatment Session 12": "T12",
        "T3 IE": "IE T3",
        "T6 IE": "IE T6",
        "T9 IE": "IE T9",
        "T12 IE": "IE T12",
    }
    return aliases.get(text, text)


def session_sort_key(session: str) -> int:
    session = canonical_session(session)
    return SESSION_ORDER.index(session) if session in SESSION_ORDER else len(SESSION_ORDER) + 1


def is_asap_session(session: str) -> bool:
    return canonical_session(session).upper().startswith("ASAP")


def final_session_allowed(session: str) -> bool:
    session = canonical_session(session)
    return session not in EXCLUDED_FINAL_SESSIONS and not is_asap_session(session)


def standardize_subject_id(raw: Any) -> str:
    digits = "".join(re.findall(r"\d", clean(raw)))
    if len(digits) >= 3:
        return digits[-3:]
    return digits


def infer_arm(subject_id: str) -> str:
    return {"1": "arm1", "2": "arm2", "3": "arm3"}.get(clean(subject_id)[:1], "unknown")


def status_from_complete_code(value: Any) -> str:
    text = clean(value).lower()
    if text == "2":
        return "complete"
    if text == "1":
        return "unverified"
    if text == "0":
        return "incomplete"
    if text:
        return "review_required"
    return "missing"


def status_from_presence(value: Any) -> str:
    return "complete" if clean(value) else "missing"


def status_from_fields(row: pd.Series, fields: str | list[str] | None) -> str:
    if not fields:
        return "missing"
    field_list = [fields] if isinstance(fields, str) else fields
    statuses = []
    for field in field_list:
        if field not in row.index:
            continue
        value = row[field]
        statuses.append(status_from_complete_code(value) if clean(value) in {"0", "1", "2"} else status_from_presence(value))
    if not statuses:
        return "missing"
    if "complete" in statuses:
        return "complete"
    if "unverified" in statuses:
        return "unverified"
    if "review_required" in statuses:
        return "review_required"
    if "incomplete" in statuses:
        return "incomplete"
    return "missing"


def qc_status_from_poor_quality(value: Any) -> str:
    text = clean(value).lower()
    if text in TRUE_VALUES:
        return "fail"
    if text in FALSE_VALUES:
        return "pass"
    return "review_required"


def read_table(path: Path, log: ValidationLog) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Missing required input file: {path}")
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
    else:
        df = pd.read_excel(path, dtype=str, keep_default_na=False)
    df.columns = [normalize_column(col) for col in df.columns]
    log.info(f"Loaded {path.name}: {df.shape[0]} rows, {df.shape[1]} columns")
    return df.fillna("")


def expand_design_session_label(label: str) -> list[str]:
    value = canonical_session(label)
    explicit = {
        "T3 / T3 IE": ["T3", "IE T3"],
        "T6 / T6 IE": ["T6", "IE T6"],
        "T9 / T9 IE": ["T9", "IE T9"],
        "T12 / T12 IE": ["T12", "IE T12"],
        "Treatment Session 1, 2, 4, 5, 7, 8, 10, and 11": ["T1", "T2", "T4", "T5", "T7", "T8", "T10", "T11"],
        "Treatment Session 7, 8, 10, and 11": ["T7", "T8", "T10", "T11"],
    }
    if clean(label) in explicit:
        return explicit[clean(label)]
    return [value]


def load_required_sessions(input_dir: Path, log: ValidationLog) -> pd.DataFrame:
    path = input_dir / DESIGN_FILES["required_sessions"]
    if not path.exists():
        raise FileNotFoundError(f"Required design file not found: {path}")
    rows = []
    workbook = pd.ExcelFile(path)
    for sheet in workbook.sheet_names:
        arm = "arm" + re.search(r"(\d+)", sheet).group(1)
        df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False)
        df.columns = [normalize_column(col) for col in df.columns]
        for _, row in df.iterrows():
            session = canonical_session(row.get("required_session", ""))
            rows.append(
                {
                    "arm": arm,
                    "session": session,
                    "order": int(clean(row.get("order", "999")) or 999),
                    "participant_event_name": clean(row.get("participant_event_name", "")),
                    "clinician_event_name": clean(row.get("clinician_event_name", "")),
                }
            )
    out = pd.DataFrame(rows)
    log.count("required_session_rows_loaded", len(out))
    return out


def load_expected_instruments(input_dir: Path, required_sessions: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    path = input_dir / DESIGN_FILES["expected_instruments"]
    if not path.exists():
        raise FileNotFoundError(f"Required design file not found: {path}")
    rows = []
    workbook = pd.ExcelFile(path)
    required_by_arm = {
        arm: set(group["session"])
        for arm, group in required_sessions.groupby("arm", sort=False)
    }
    for sheet in workbook.sheet_names:
        arm = "arm" + re.search(r"(\d+)", sheet).group(1)
        df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False)
        df.columns = [normalize_column(col) for col in df.columns]
        if "included_in_summary" not in df.columns:
            raise ValueError(f"{path.name} sheet {sheet} is missing Included_in_summary")
        excluded = 0
        for _, row in df.iterrows():
            if clean(row.get("included_in_summary", "")) != "1":
                excluded += 1
                continue
            source = clean(row.get("source_of_instrument", ""))
            instrument = clean(row.get("instrument", ""))
            for session in expand_design_session_label(row.get("session", "")):
                if not final_session_allowed(session):
                    continue
                if session not in required_by_arm.get(arm, set()):
                    log.warn(f"Included design row is not required for {arm}: {session} / {instrument}")
                    continue
                rows.append(
                    {
                        "arm": arm,
                        "session": session,
                        "instrument": instrument,
                        "instrument_source": source,
                    }
                )
        log.info(f"{path.name} {sheet}: excluded {excluded} rows where Included_in_summary != 1")
    out = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
    log.count("included_expected_instrument_rows", len(out))
    return out


def build_event_map(required_sessions: pd.DataFrame) -> dict[tuple[str, str], str]:
    mapping: dict[tuple[str, str], str] = {}
    for _, row in required_sessions.iterrows():
        for source, col in [("participants", "participant_event_name"), ("clinician", "clinician_event_name")]:
            event = clean(row[col])
            if event and not event.startswith("—"):
                mapping[(source, event)] = row["session"]
    return mapping


def event_to_session(source: str, event_name: str, event_map: dict[tuple[str, str], str], log: ValidationLog) -> str:
    event = clean(event_name)
    if (source, event) in event_map:
        return canonical_session(event_map[(source, event)])
    base = re.sub(r"_arm_\d+$", "", event)
    if base in EVENT_OVERRIDES:
        return EVENT_OVERRIDES[base]
    match = re.match(r"treatment_session(?:_arm_\d+)?([a-p])?$", event)
    if match:
        # Fallback for source-specific treatment event shifts when the design file uses notes.
        letter = match.group(1) or ""
        index = "" if not letter else chr(ord(letter) - 48)
        log.warn(f"Observed event not explicitly mapped by design: {source} {event}")
        return f"UNMAPPED_EVENT_{base}{index}"
    log.warn(f"Unknown REDCap event: {source} {event}")
    return f"UNMAPPED_EVENT_{event}"


def load_redcap(input_dir: Path, required_sessions: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    frames = []
    event_map = build_event_map(required_sessions)
    for source, filename in INPUT_FILES.items():
        df = read_table(input_dir / filename, log)
        id_col = "record_id" if source == "participants" else "preescreen_id"
        if id_col not in df.columns:
            raise ValueError(f"{filename} is missing required subject ID column {id_col}")
        if "redcap_event_name" not in df.columns:
            raise ValueError(f"{filename} is missing redcap_event_name")
        df["source_file"] = filename
        df["source"] = source
        df["subject_id"] = df[id_col].map(standardize_subject_id)
        df["arm"] = df["subject_id"].map(infer_arm)
        df["session"] = df["redcap_event_name"].map(lambda event: event_to_session(source, event, event_map, log))
        frames.append(df)
    out = pd.concat(frames, ignore_index=True, sort=False).fillna("")
    bad = out[(out["subject_id"].str.len() != 3) | (out["arm"] == "unknown")]
    if not bad.empty:
        log.warn(f"{len(bad)} REDCap rows have nonstandard subject IDs or unknown arms")
    log.count("redcap_rows_loaded", len(out))
    return out


def source_instrument_config(source_label: str, instrument: str) -> dict[str, str | list[str]]:
    if source_label.lower().startswith("participant"):
        return PARTICIPANT_INSTRUMENTS.get(instrument, {})
    return CLINICIAN_INSTRUMENTS.get(instrument, {})


def build_questionnaire_long(redcap: pd.DataFrame, expected: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    rows = []
    expected_q = expected.copy()
    for _, exp in expected_q.iterrows():
        source_name = exp["instrument_source"]
        source = "participants" if source_name.lower().startswith("participant") else "clinician"
        cfg = source_instrument_config(source_name, exp["instrument"])
        if not cfg:
            log.warn(f"No field mapping for expected instrument: {exp['arm']} {exp['session']} {exp['instrument']}")
        matches = redcap[
            (redcap["arm"] == exp["arm"])
            & (redcap["session"] == exp["session"])
            & (redcap["source"] == source)
        ]
        subjects = sorted(set(redcap.loc[redcap["arm"] == exp["arm"], "subject_id"]))
        observed_by_subject = {sid: group.iloc[0] for sid, group in matches.groupby("subject_id", sort=False)}
        for subject_id in subjects:
            row = observed_by_subject.get(subject_id)
            date_field = clean(cfg.get("date", ""))
            complete_field = cfg.get("complete")
            if row is None:
                status = "missing"
                date = ""
            else:
                status = status_from_fields(row, complete_field)
                date = clean(row.get(date_field, "")) if date_field else ""
                if status == "missing" and date:
                    status = "complete"
            rows.append(
                {
                    "subject_id": subject_id,
                    "arm": exp["arm"],
                    "session": exp["session"],
                    "instrument": exp["instrument"],
                    "instrument_source": exp["instrument_source"],
                    "date": date,
                    "status": status,
                    "expected": True,
                }
            )
    out = pd.DataFrame(rows)
    log.count("questionnaire_expected_rows", len(out))
    return out


def build_session_long(redcap: pd.DataFrame, required_sessions: pd.DataFrame) -> pd.DataFrame:
    subjects = redcap[["subject_id", "arm"]].drop_duplicates()
    rows = []
    date_candidates = [
        "phq_9_date",
        "gad_7_date",
        "scid5_axis1_date",
        "eligibility_date",
        "peas_date_of_visit",
        "scan_date",
        "asap_date",
    ]
    observed_dates = {}
    for (subject_id, session), group in redcap.groupby(["subject_id", "session"], sort=False):
        dates = []
        for _, record in group.iterrows():
            for field in date_candidates:
                if field in record.index and clean(record[field]):
                    dates.append(clean(record[field]))
        observed_dates[(subject_id, session)] = dates[0] if dates else ""
    for _, subject in subjects.iterrows():
        arm = subject["arm"]
        required = required_sessions[
            (required_sessions["arm"] == arm)
            & (required_sessions["session"].map(final_session_allowed))
        ]
        for _, req in required.iterrows():
            session = req["session"]
            date = observed_dates.get((subject["subject_id"], session), "")
            rows.append(
                {
                    "subject_id": subject["subject_id"],
                    "arm": arm,
                    "session": session,
                    "session_completed": bool(date),
                    "session_date": date,
                }
            )
    session_long = pd.DataFrame(rows)
    if session_long.empty:
        return session_long
    baseline_dates = (
        session_long[session_long["session"] == "Baseline"][["subject_id", "session_date"]]
        .rename(columns={"session_date": "baseline_date"})
        .drop_duplicates("subject_id")
    )
    session_long = session_long.merge(baseline_dates, on="subject_id", how="left")
    session_long["intervalFromBaseline_weeks"] = session_long.apply(interval_weeks, axis=1)
    session_long["interval_valid"] = session_long.apply(
        lambda row: "" if row["session"] == "Baseline" else (pd.notna(row["intervalFromBaseline_weeks"]) and row["intervalFromBaseline_weeks"] >= 0),
        axis=1,
    )
    return session_long


def interval_weeks(row: pd.Series) -> float | None:
    if row["session"] == "Baseline":
        return None
    baseline = pd.to_datetime(clean(row.get("baseline_date", "")), errors="coerce")
    current = pd.to_datetime(clean(row.get("session_date", "")), errors="coerce")
    if pd.isna(baseline) or pd.isna(current):
        return None
    return round((current - baseline).days / 7, 2)


def load_mri_long(input_dir: Path, redcap: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    subject_arms = redcap[["subject_id", "arm"]].drop_duplicates().set_index("subject_id")["arm"].to_dict()
    rows = []
    for kind, filename in MRI_FILES.items():
        path = input_dir / filename
        if not path.exists():
            log.warn(f"MRI QC file missing: {path}")
            continue
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        required = ["subID", "sesID", "modality", "Poor_Quality"]
        if kind == "func":
            required.extend(["taskID", "runID"])
        missing = [col for col in required if col not in df.columns]
        if missing:
            log.warn(f"{filename} missing required columns: {missing}")
            continue
        for _, row in df.iterrows():
            subject_id = standardize_subject_id(row["subID"])
            session = MRI_SESSION_MAP.get(clean(row["sesID"]), f"UNMAPPED_MRI_SESSION_{clean(row['sesID'])}")
            if session.startswith("UNMAPPED"):
                log.warn(f"Unmapped MRI sesID: {row['sesID']}")
            if kind == "anat":
                scan = "anat_T1w"
            else:
                run = clean(row.get("runID", "")).zfill(2)
                scan = f"func_{clean(row.get('taskID', 'task'))}_run{run}"
            rows.append(
                {
                    "subject_id": subject_id,
                    "arm": subject_arms.get(subject_id, infer_arm(subject_id)),
                    "session": session,
                    "scan_or_run": scan,
                    "qc_status": qc_status_from_poor_quality(row["Poor_Quality"]),
                    "expected": True,
                    "source_file": filename,
                }
            )
    out = pd.DataFrame(rows)
    log.count("mri_qc_rows_loaded", len(out))
    return out


def add_missing_mri_expectations(mri: pd.DataFrame, redcap: pd.DataFrame) -> pd.DataFrame:
    if mri.empty:
        return mri
    subjects = redcap[["subject_id", "arm"]].drop_duplicates()
    rows = []
    for (arm, session, scan), group in mri.groupby(["arm", "session", "scan_or_run"], sort=False):
        arm_subjects = subjects[subjects["arm"] == arm]["subject_id"].tolist()
        observed = set(group["subject_id"])
        for subject_id in arm_subjects:
            if subject_id in observed:
                rows.extend(group[group["subject_id"] == subject_id].to_dict("records"))
            else:
                rows.append(
                    {
                        "subject_id": subject_id,
                        "arm": arm,
                        "session": session,
                        "scan_or_run": scan,
                        "qc_status": "missing",
                        "expected": True,
                        "source_file": "expected_from_observed_mri_design",
                    }
                )
    return pd.DataFrame(rows).drop_duplicates(["subject_id", "arm", "session", "scan_or_run"])


def build_task_long(redcap: pd.DataFrame) -> pd.DataFrame:
    scan_fields = [col for col in redcap.columns if col.startswith("scan_run_")]
    rows = []
    for _, row in redcap.iterrows():
        if not final_session_allowed(row["session"]):
            continue
        for field in scan_fields:
            value = clean(row.get(field, ""))
            if not value:
                continue
            rows.append(
                {
                    "subject_id": row["subject_id"],
                    "arm": row["arm"],
                    "session": row["session"],
                    "task": field.replace("scan_run_", "scan_run"),
                    "status": "complete" if value in {"100", "2", "1", "yes", "Yes"} else "review_required",
                    "completed": value,
                    "expected": True,
                }
            )
    return pd.DataFrame(rows)


def build_behavioral_qc_long(task_long: pd.DataFrame) -> pd.DataFrame:
    if task_long.empty:
        return pd.DataFrame(columns=["subject_id", "arm", "session", "instrument", "instrument_source", "qc_status", "expected"])
    out = task_long.copy()
    out["instrument"] = out["task"]
    out["instrument_source"] = "behavioral_qc"
    out["qc_status"] = out["status"].map(lambda value: "pass" if value == "complete" else "review_required")
    return out[["subject_id", "arm", "session", "instrument", "instrument_source", "qc_status", "expected"]]


def aggregate_status(values: Iterable[str], good: str) -> bool:
    vals = [clean(value) for value in values]
    return bool(vals) and all(value == good for value in vals)


def pivot_subject_wise(
    arm: str,
    redcap: pd.DataFrame,
    qlong: pd.DataFrame,
    task_long: pd.DataFrame,
    behavioral_long: pd.DataFrame,
    mri_long: pd.DataFrame,
    session_long: pd.DataFrame,
) -> pd.DataFrame:
    subjects = sorted(redcap.loc[redcap["arm"] == arm, "subject_id"].dropna().unique())
    if not subjects:
        return empty_subject_wise()
    rows = []
    for subject_id in subjects:
        row: dict[str, Any] = {
            "subject_id": subject_id,
            "arm": arm,
            "dropout_status": "active",
        }
        subj_q = qlong[(qlong["arm"] == arm) & (qlong["subject_id"] == subject_id)]
        for _, item in subj_q.iterrows():
            prefix = f"ses-{slug(item['session'])}_qn-{slug(item['instrument'])}"
            row[f"{prefix}_date"] = item["date"]
            row[f"{prefix}_status"] = item["status"]
        subj_task = task_long[(task_long["arm"] == arm) & (task_long["subject_id"] == subject_id)]
        for _, item in subj_task.iterrows():
            prefix = f"ses-{slug(item['session'])}_task-{slug(item['task'])}"
            row[f"{prefix}_status"] = item["status"]
            row[f"{prefix}_completed"] = item["completed"]
        subj_beh = behavioral_long[(behavioral_long["arm"] == arm) & (behavioral_long["subject_id"] == subject_id)]
        for _, item in subj_beh.iterrows():
            row[f"ses-{slug(item['session'])}_beh-{slug(item['instrument'])}_qc_status"] = item["qc_status"]
        subj_mri = mri_long[(mri_long["arm"] == arm) & (mri_long["subject_id"] == subject_id)]
        for _, item in subj_mri.iterrows():
            row[f"ses-{slug(item['session'])}_mri-{slug(item['scan_or_run'])}_qc_status"] = item["qc_status"]
        subj_sessions = session_long[(session_long["arm"] == arm) & (session_long["subject_id"] == subject_id)]
        missing_sessions = []
        for _, item in subj_sessions.iterrows():
            session = item["session"]
            if session == "Baseline":
                continue
            row[f"ses-{slug(session)}_session_completed"] = bool(item["session_completed"])
            row[f"ses-{slug(session)}_intervalFromBaseline_weeks"] = item["intervalFromBaseline_weeks"]
            row[f"ses-{slug(session)}_interval_valid"] = item["interval_valid"]
        for _, item in subj_sessions.iterrows():
            if item["session"] not in EXCLUDED_COMPLETENESS_SESSIONS and not bool(item["session_completed"]):
                missing_sessions.append(item["session"])
        row["total_ASAP_count"] = int(redcap[(redcap["subject_id"] == subject_id) & (redcap["session"].str.contains("ASAP", na=False))].shape[0])
        row["has_ASAP"] = row["total_ASAP_count"] > 0
        row["complete_all_experiment_sessions"] = len(missing_sessions) == 0
        row["missing_expected_experiment_sessions"] = "; ".join(missing_sessions)
        row["missing_required_experiment_sessions"] = "; ".join(missing_sessions)
        row["all_mustHave_questionnaires_complete_perSession"] = per_session_summary(subj_q, "status", "complete")
        row["all_mustHave_questionnaires_complete"] = aggregate_status(subj_q["status"], "complete")
        row["all_required_task_complete_perSession"] = per_session_summary(subj_task, "status", "complete")
        row["all_required_task_complete"] = True if subj_task.empty else aggregate_status(subj_task["status"], "complete")
        row["all_required_MRI_QC_passed_perSession"] = per_session_summary(subj_mri, "qc_status", "pass")
        row["all_required_MRI_QC_passed"] = True if subj_mri.empty else aggregate_status(subj_mri["qc_status"], "pass")
        row["all_required_selfOther_QC_passed_perSession"] = per_session_summary(subj_beh, "qc_status", "pass")
        row["all_required_selfOther_QC_passed"] = True if subj_beh.empty else aggregate_status(subj_beh["qc_status"], "pass")
        criteria = [
            row["complete_all_experiment_sessions"],
            row["all_mustHave_questionnaires_complete"],
            row["all_required_task_complete"],
            row["all_required_MRI_QC_passed"],
            row["all_required_selfOther_QC_passed"],
        ]
        row["all_required_criteria_passed_perSession"] = ""
        row["all_required_criteria_passed"] = all(criteria)
        row["ready_for_analysis"] = row["all_required_criteria_passed"]
        row["needs_followup"] = not row["ready_for_analysis"]
        row["overall_QA_status"] = "ready_for_analysis" if row["ready_for_analysis"] else "needs_followup"
        rows.append(row)
    out = pd.DataFrame(rows)
    return order_subject_columns(out)


def subject_wise_base_columns() -> list[str]:
    return [
        "subject_id",
        "arm",
        "dropout_status",
        "overall_QA_status",
    ]


def subject_wise_asap_columns() -> list[str]:
    return ["total_ASAP_count", "has_ASAP"]


def subject_wise_summary_columns() -> list[str]:
    return [
        "complete_all_experiment_sessions",
        "missing_expected_experiment_sessions",
        "missing_required_experiment_sessions",
        "all_mustHave_questionnaires_complete_perSession",
        "all_mustHave_questionnaires_complete",
        "all_required_task_complete_perSession",
        "all_required_task_complete",
        "all_required_MRI_QC_passed_perSession",
        "all_required_MRI_QC_passed",
        "all_required_selfOther_QC_passed_perSession",
        "all_required_selfOther_QC_passed",
        "all_required_criteria_passed_perSession",
        "all_required_criteria_passed",
        "ready_for_analysis",
        "needs_followup",
    ]


def per_session_summary(df: pd.DataFrame, status_col: str, pass_value: str) -> str:
    if df.empty:
        return ""
    parts = []
    for session, group in df.groupby("session", sort=False):
        ok = aggregate_status(group[status_col], pass_value)
        parts.append(f"{session}:{ok}")
    return "; ".join(parts)


def order_subject_columns(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return empty_subject_wise()
    base = subject_wise_base_columns()
    asap = subject_wise_asap_columns()
    summary = subject_wise_summary_columns()
    fixed = set(base + asap + summary)
    dynamic = [col for col in df.columns if col not in fixed]
    ordered = (
        [col for col in base if col in df.columns]
        + sorted(dynamic, key=subject_dynamic_column_sort_key)
        + [col for col in asap if col in df.columns]
        + [col for col in summary if col in df.columns]
    )
    return df[ordered]


def empty_subject_wise() -> pd.DataFrame:
    return pd.DataFrame(
        columns=subject_wise_base_columns()
        + subject_wise_asap_columns()
        + subject_wise_summary_columns()
    )


def subject_dynamic_column_sort_key(column: str) -> tuple[int, str, int, str]:
    match = re.match(r"ses-([^_]+)_(qn|task|beh|mri)-(.+)", column)
    if not match:
        timing = re.match(r"ses-([^_]+)_(session_completed|intervalFromBaseline_weeks|interval_valid)$", column)
        if timing:
            session = unslug_session(timing.group(1))
            timing_order = {
                "session_completed": 0,
                "intervalFromBaseline_weeks": 1,
                "interval_valid": 2,
            }.get(timing.group(2), 9)
            return (session_sort_key(session), "z_timing", timing_order, column)
        return (999, "zz", 999, column)
    session = unslug_session(match.group(1))
    section_order = {"qn": "a_qn", "task": "b_task", "beh": "c_beh", "mri": "d_mri"}[match.group(2)]
    return (session_sort_key(session), section_order, 0, column)


def unslug_session(value: str) -> str:
    compact = clean(value)
    for session in SESSION_ORDER:
        if slug(session) == compact:
            return session
    return compact


def summarize_questionnaires(arm: str, qlong: pd.DataFrame) -> pd.DataFrame:
    rows = []
    data = qlong[qlong["arm"] == arm]
    for (session, instrument, source), group in data.groupby(["session", "instrument", "instrument_source"], sort=False):
        expected = len(group)
        complete = int((group["status"] == "complete").sum())
        review = int(group["status"].isin(["unverified", "review_required"]).sum())
        missing = int(group["status"].isin(["missing", "incomplete"]).sum())
        rows.append(rate_row(arm, session, instrument, source, expected, complete, review, missing, "complete"))
    return sort_group(pd.DataFrame(rows))


def summarize_instruments(arm: str, qlong: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "arm",
        "instrument_or_item",
        "instrument_source",
        "expected_NofSubjects",
        "complete_NofRecords",
        "review_required_NofRecords",
        "complete_rate",
        "missing_rate",
    ]
    rows = []
    data = qlong[qlong["arm"] == arm]
    for (instrument, source), group in data.groupby(["instrument", "instrument_source"], sort=False):
        subject_statuses = group.groupby("subject_id")["status"].apply(list)
        expected = int(subject_statuses.index.nunique())
        complete = int(subject_statuses.map(lambda statuses: all(status == "complete" for status in statuses)).sum())
        review = int(subject_statuses.map(lambda statuses: any(status in {"unverified", "review_required"} for status in statuses)).sum())
        missing = int(subject_statuses.map(lambda statuses: any(status in {"missing", "incomplete"} for status in statuses)).sum())
        rows.append(
            {
                "arm": arm,
                "instrument_or_item": instrument,
                "instrument_source": source,
                "expected_NofSubjects": expected,
                "complete_NofRecords": complete,
                "review_required_NofRecords": review,
                "complete_rate": safe_rate(complete, expected),
                "missing_rate": safe_rate(missing, expected),
            }
        )
    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows, columns=columns).sort_values(["instrument_source", "instrument_or_item"]).reset_index(drop=True)


def rate_row(
    arm: str,
    session: str,
    instrument: str,
    source: str,
    expected: int,
    good: int,
    review: int,
    missing: int,
    good_label: str,
) -> dict[str, Any]:
    return {
        "arm": arm,
        "session": session,
        "instrument": instrument,
        "instrument_source": source,
        "expected_NofSubjects": expected,
        f"{good_label}_NofSubjects": good,
        "review_required_NofSubjects": review,
        f"{good_label}_rate": safe_rate(good, expected),
        "missing_rate": safe_rate(missing, expected),
    }


def safe_rate(numerator: int, denominator: int) -> float:
    return round(numerator / denominator, 3) if denominator else 0.0


def sort_group(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "session" not in df.columns:
        return df
    df = df.copy()
    df["_sort"] = df["session"].map(session_sort_key)
    return df.sort_values(["_sort"] + [c for c in ["instrument", "instrument_source"] if c in df.columns]).drop(columns="_sort").reset_index(drop=True)


def summarize_behavioral(arm: str, behavioral: pd.DataFrame) -> pd.DataFrame:
    rows = []
    data = behavioral[behavioral["arm"] == arm]
    for (session, instrument, source), group in data.groupby(["session", "instrument", "instrument_source"], sort=False):
        expected = len(group)
        passed = int((group["qc_status"] == "pass").sum())
        failed = int((group["qc_status"] == "fail").sum())
        missing = int((group["qc_status"] == "missing").sum())
        review = int(group["qc_status"].isin(["unverified", "review_required"]).sum())
        rows.append(
            {
                "arm": arm,
                "session": session,
                "instrument": instrument,
                "instrument_source": source,
                "expected_NofSubjects": expected,
                "qc_pass_NofSubjects": passed,
                "qc_fail_NofSubjects": failed,
                "missing_NofSubjects": missing,
                "review_required_NofSubjects": review,
                "qc_pass_rate": safe_rate(passed, expected),
                "missing_rate": safe_rate(missing, expected),
            }
        )
    return sort_group(pd.DataFrame(rows))


def summarize_mri_by_session(arm: str, mri: pd.DataFrame) -> pd.DataFrame:
    rows = []
    data = mri[mri["arm"] == arm]
    for session, group in data.groupby("session", sort=False):
        by_subject = group.groupby("subject_id")["qc_status"].apply(list)
        expected = len(by_subject)
        passed = int(by_subject.map(lambda values: all(value == "pass" for value in values)).sum())
        failed = int(by_subject.map(lambda values: any(value == "fail" for value in values)).sum())
        missing = int(by_subject.map(lambda values: any(value == "missing" for value in values)).sum())
        review = int(by_subject.map(lambda values: any(value == "review_required" for value in values)).sum())
        rows.append(
            {
                "arm": arm,
                "session": session,
                "expected_NofSubjects": expected,
                "qc_pass_NofSubjects": passed,
                "qc_fail_NofSubjects": failed,
                "missing_NofSubjects": missing,
                "review_required_NofSubjects": review,
                "qc_pass_rate": safe_rate(passed, expected),
                "missing_rate": safe_rate(missing, expected),
            }
        )
    return sort_group(pd.DataFrame(rows))


def summarize_mri_by_scan(arm: str, mri: pd.DataFrame) -> pd.DataFrame:
    rows = []
    data = mri[mri["arm"] == arm]
    for (session, scan), group in data.groupby(["session", "scan_or_run"], sort=False):
        expected = len(group)
        passed = int((group["qc_status"] == "pass").sum())
        failed = int((group["qc_status"] == "fail").sum())
        missing = int((group["qc_status"] == "missing").sum())
        review = int((group["qc_status"] == "review_required").sum())
        rows.append(
            {
                "arm": arm,
                "session": session,
                "scan_or_run": scan,
                "expected_NofSubjects": expected,
                "qc_pass_NofSubjects": passed,
                "qc_fail_NofSubjects": failed,
                "missing_NofSubjects": missing,
                "review_required_NofSubjects": review,
                "qc_pass_rate": safe_rate(passed, expected),
                "missing_rate": safe_rate(missing, expected),
            }
        )
    return sort_group(pd.DataFrame(rows))


def summarize_intervals(arm: str, session_long: pd.DataFrame) -> pd.DataFrame:
    rows = []
    data = session_long[(session_long["arm"] == arm) & (session_long["session"] != "Baseline")]
    for session, group in data.groupby("session", sort=False):
        intervals = pd.to_numeric(group["intervalFromBaseline_weeks"], errors="coerce").dropna()
        rows.append(
            {
                "arm": arm,
                "session": session,
                "completed_NofSubjects": int(group["session_completed"].sum()),
                "valid_interval_NofSubjects": int((group["interval_valid"] == True).sum()),
                "invalid_interval_NofSubjects": int((group["interval_valid"] == False).sum()),
                "mean_intervalFromBaseline_weeks": round(float(intervals.mean()), 2) if len(intervals) else "",
                "sd_intervalFromBaseline_weeks": round(float(intervals.std()), 2) if len(intervals) > 1 else "",
                "min_intervalFromBaseline_weeks": round(float(intervals.min()), 2) if len(intervals) else "",
                "max_intervalFromBaseline_weeks": round(float(intervals.max()), 2) if len(intervals) else "",
            }
        )
    return sort_group(pd.DataFrame(rows))


def summarize_readiness(arm: str, subject_wise: pd.DataFrame) -> pd.DataFrame:
    if subject_wise.empty:
        total = 0
        sw = pd.DataFrame()
    else:
        total = len(subject_wise)
        sw = subject_wise
    bool_count = lambda col: int(sw[col].sum()) if col in sw.columns and total else 0
    return pd.DataFrame(
        [
            {
                "arm": arm,
                "total_NofSubjects": total,
                "withdrawn_or_dropout_NofSubjects": int((sw.get("dropout_status", pd.Series(dtype=str)) != "active").sum()) if total else 0,
                "complete_all_experiment_sessions_NofSubjects": bool_count("complete_all_experiment_sessions"),
                "all_mustHave_questionnaires_complete_NofSubjects": bool_count("all_mustHave_questionnaires_complete"),
                "all_required_MRI_QC_passed_NofSubjects": bool_count("all_required_MRI_QC_passed"),
                "all_required_selfOther_QC_passed_NofSubjects": bool_count("all_required_selfOther_QC_passed"),
                "all_required_criteria_passed_NofSubjects": bool_count("all_required_criteria_passed"),
                "ready_for_analysis_NofSubjects": bool_count("ready_for_analysis"),
                "needs_followup_NofSubjects": bool_count("needs_followup"),
            }
        ]
    )


def build_group_tables(
    arm: str,
    qlong: pd.DataFrame,
    task_long: pd.DataFrame,
    behavioral_long: pd.DataFrame,
    mri_long: pd.DataFrame,
    session_long: pd.DataFrame,
    subject_wise: pd.DataFrame,
) -> list[tuple[str, pd.DataFrame]]:
    task_as_behavioral = behavioral_long.copy()
    return [
        ("Table 1. Questionnaire completion by session", summarize_questionnaires(arm, qlong)),
        ("Table 2. Instrument-wise summary", summarize_instruments(arm, qlong)),
        ("Table 3. Behavioral QC pass rate by session", summarize_behavioral(arm, task_as_behavioral)),
        ("Table 4. MRI QC pass rate by session", summarize_mri_by_session(arm, mri_long)),
        ("Table 5. MRI QC pass rate by subject", summarize_mri_by_scan(arm, mri_long)),
        ("Table 6. Session interval summary after Baseline", summarize_intervals(arm, session_long)),
        ("Table 7. Participant-level QA readiness summary", summarize_readiness(arm, subject_wise)),
    ]


def validate_outputs(subject_wise: pd.DataFrame, group_tables: list[tuple[str, pd.DataFrame]], log: ValidationLog) -> None:
    if subject_wise.columns.duplicated().any():
        raise ValueError("subject_wise has duplicate columns")
    forbidden_tokens = ["Screening", "Baseline 1", "Baseline1", "_arm_", "UNMAPPED_EVENT_", "UNMAPPED_MRI_SESSION_"]
    for token in forbidden_tokens:
        if any(token in str(col) for col in subject_wise.columns):
            raise ValueError(f"Forbidden token in subject_wise columns: {token}")
        if not subject_wise.empty and subject_wise.astype(str).apply(lambda col: col.str.contains(token, regex=False, na=False)).any().any():
            raise ValueError(f"Forbidden token in subject_wise values: {token}")
    if any("total_score" in col.lower() or "totalscore" in col.lower() for col in subject_wise.columns):
        raise ValueError("Total score columns must not appear in subject_wise")
    if len(group_tables) != 7:
        raise ValueError("group_wise must contain seven tables")
    log.count("group_wise_tables_written", len(group_tables))


def write_workbook(path: Path, subject_wise: pd.DataFrame, group_tables: list[tuple[str, pd.DataFrame]]) -> None:
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        subject_wise.to_excel(writer, sheet_name="subject_wise", index=False)
        start = 0
        for title, table in group_tables:
            pd.DataFrame([[title]]).to_excel(writer, sheet_name="group_wise", startrow=start, index=False, header=False)
            start += 1
            table.to_excel(writer, sheet_name="group_wise", startrow=start, index=False)
            start += len(table) + 3
    format_workbook(path)


def format_workbook(path: Path) -> None:
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Font
    except Exception:
        return
    workbook = load_workbook(path)
    for sheet in workbook.worksheets:
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        for column in sheet.columns:
            width = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column)
            sheet.column_dimensions[column[0].column_letter].width = min(max(width + 2, 10), 45)
    workbook.save(path)


def build(input_dir: Path, output_dir: Path) -> tuple[list[Path], Path]:
    log = ValidationLog()
    log.info(f"Input directory: {input_dir}")
    log.info(f"Output directory: {output_dir}")
    required_sessions = load_required_sessions(input_dir, log)
    expected = load_expected_instruments(input_dir, required_sessions, log)
    redcap = load_redcap(input_dir, required_sessions, log)
    redcap_final = redcap[~redcap["session"].isin(EXCLUDED_FINAL_SESSIONS)].copy()
    qlong = build_questionnaire_long(redcap_final, expected, log)
    session_long = build_session_long(redcap_final, required_sessions)
    mri_long = add_missing_mri_expectations(load_mri_long(input_dir, redcap_final, log), redcap_final)
    mri_long = mri_long[~mri_long["session"].isin(EXCLUDED_FINAL_SESSIONS)].copy() if not mri_long.empty else mri_long
    task_long = build_task_long(redcap_final)
    behavioral_long = build_behavioral_qc_long(task_long)
    output_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    for arm in ARMS:
        subject_wise = pivot_subject_wise(
            arm,
            redcap_final,
            qlong,
            task_long,
            behavioral_long,
            mri_long,
            session_long,
        )
        tables = build_group_tables(arm, qlong, task_long, behavioral_long, mri_long, session_long, subject_wise)
        validate_outputs(subject_wise, tables, log)
        path = output_dir / f"QA_summary_{arm}.xlsx"
        write_workbook(path, subject_wise, tables)
        paths.append(path)
        log.count(f"{arm}_subjects", len(subject_wise))
        log.count(f"{arm}_ready_for_analysis", int(subject_wise["ready_for_analysis"].sum()) if not subject_wise.empty else 0)
    log_path = output_dir / "QA_summary_validation_log.txt"
    log.write(log_path)
    return paths, log_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate arm-specific QA summary workbooks.")
    parser.add_argument("--input-dir", type=Path, default=Path("data"))
    parser.add_argument("--output-dir", type=Path, default=Path("results"))
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    paths, log_path = build(args.input_dir.resolve(), args.output_dir.resolve())
    for path in paths:
        print(f"Wrote {path}")
    print(f"Wrote {log_path}")


if __name__ == "__main__":
    main()
