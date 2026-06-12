#!/usr/bin/env python3
"""Build arm-specific QA summary workbooks from REDCap, design, and MRI QC files.

This script is intentionally design-file-driven:

* Required sessions come from Required_Sessions_for_each_Arm.xlsx.
* Expected instruments come from Instruments_in_each_Session_each_Arm.xlsx.
* Only rows with Included_in_summary == 1 are included in final columns and
  denominators.
"""

from __future__ import annotations

import argparse
import math
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

import pandas as pd


ARMS = ("arm1", "arm2", "arm3")
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
CODEBOOK_FILES = ("Participants_REDCap.pdf", "Clini_REDCap.pdf")

TRUE_VALUES = {"true", "1", "yes", "y"}
FALSE_VALUES = {"false", "0", "no", "n"}
COMPLETE_STATUSES = {"complete"}
REVIEW_STATUSES = {"review_required", "unverified"}
MISSING_STATUSES = {"missing", "incomplete"}

MRI_SESSION_MAP = {
    "ses-baseline": "Baseline",
    "ses-repeatbaseline": "Repeat Baseline",
    "ses-t1": "T1",
    "ses-t2": "T2",
    "ses-t3": "T3",
    "ses-t4": "T4",
    "ses-t5": "T5",
    "ses-t6": "T6 Scan",
    "ses-t7": "T7",
    "ses-t8": "T8",
    "ses-t9": "T9",
    "ses-t10": "T10",
    "ses-t11": "T11",
    "ses-t12": "T12 Scan",
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
    "Treatment Visit Checklist": {
        "complete": ["t3t9_checklist_other", "t6t12_checklist_other_visit_2", "t12_other_2"]
    },
    "T3/T9 IE Visit Checklist": {"complete": "t3t9_checklist_other"},
    "T6/T12 IE Visit Checklist": {"complete": ["t6t12_checklist_other_visit_2", "t12_other_2"]},
    "T12 Healthy Control Only": {"complete": "t12_other_2"},
}

ANT_FIELDS_BY_SESSION = {
    "Baseline": "bl_during_other_2",
    "Repeat Baseline": "bl_during_other_2",
    "T3": "t3t9_checklist_other",
    "IE T3": "t3t9_checklist_other",
    "T6": "t6t12_checklist_other_visit_2",
    "IE T6": "t6t12_checklist_other_visit_2",
    "T9": "t3t9_checklist_other",
    "IE T9": "t3t9_checklist_other",
    "T12": "t12_other_2",
    "IE T12": "t12_other_2",
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


def normalize_column(name: str) -> str:
    value = clean(name).lower()
    value = re.sub(r"[^0-9a-z]+", "_", value)
    return re.sub(r"_+", "_", value).strip("_")


def excel_safe_label(value: str) -> str:
    return re.sub(r"[^0-9A-Za-z]+", "", canonical_session(value))


def item_label(value: str) -> str:
    text = clean(value).replace("&", "and")
    text = re.sub(r"[^0-9A-Za-z_-]+", "", text)
    return text


def canonical_session(value: Any) -> str:
    text = clean(value)
    aliases = {
        "Baseline 1": "Baseline",
        "Baseline1": "Baseline",
        "baseline_1": "Baseline",
        "baseline": "Baseline",
        "Repeat baseline": "Repeat Baseline",
        "T3 IE": "IE T3",
        "T6 IE": "IE T6",
        "T9 IE": "IE T9",
        "T12 IE": "IE T12",
    }
    if text in aliases:
        return aliases[text]
    match = re.fullmatch(r"Treatment Session\s+(\d+)", text, flags=re.IGNORECASE)
    if match:
        return f"T{int(match.group(1))}"
    match = re.fullmatch(r"ASAP Session\s+(\d+)", text, flags=re.IGNORECASE)
    if match:
        return f"ASAP {int(match.group(1))}"
    return text


def session_sort_key(session: str) -> int:
    session = canonical_session(session)
    if session in SESSION_ORDER:
        return SESSION_ORDER.index(session)
    if session.upper().startswith("ASAP"):
        return len(SESSION_ORDER) + 10
    return len(SESSION_ORDER) + 20


def final_session_allowed(session: str) -> bool:
    session = canonical_session(session)
    return session != "Screening" and not session.upper().startswith("ASAP")


def standardize_subject_id(raw: Any) -> str:
    digits = "".join(re.findall(r"\d", clean(raw)))
    return digits[-3:] if len(digits) >= 3 else digits


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
    if text in TRUE_VALUES:
        return "complete"
    if text in FALSE_VALUES:
        return "incomplete"
    if text:
        return "review_required"
    return "missing"


def combined_status(statuses: Iterable[str]) -> str:
    values = [s for s in statuses if s]
    if not values:
        return "missing"
    if any(s == "complete" for s in values):
        return "complete"
    if any(s in REVIEW_STATUSES for s in values):
        return "review_required"
    if any(s == "incomplete" for s in values):
        return "incomplete"
    return "missing"


def status_from_fields(row: pd.Series | None, fields: str | list[str] | None) -> str:
    if row is None or not fields:
        return "missing"
    field_list = [fields] if isinstance(fields, str) else fields
    statuses = [status_from_complete_code(row.get(field, "")) for field in field_list if field in row.index]
    return combined_status(statuses)


def qc_status_from_poor_quality(value: Any) -> str:
    text = clean(value).lower()
    if text in TRUE_VALUES:
        return "fail"
    if text in FALSE_VALUES:
        return "pass"
    return "review_required"


def read_csv_strings(path: Path, log: ValidationLog) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Missing required input file: {path}")
    df = pd.read_csv(path, dtype=str, keep_default_na=False).fillna("")
    log.info(f"Loaded {path.name}: {df.shape[0]} rows, {df.shape[1]} columns")
    return df


def load_required_sessions(input_dir: Path, log: ValidationLog) -> pd.DataFrame:
    path = input_dir / DESIGN_FILES["required_sessions"]
    if not path.exists():
        raise FileNotFoundError(f"Required design file not found: {path}")
    rows: list[dict[str, Any]] = []
    workbook = pd.ExcelFile(path)
    for sheet in workbook.sheet_names:
        arm_match = re.search(r"(\d+)", sheet)
        if not arm_match:
            log.warn(f"Skipping required-session sheet without arm number: {sheet}")
            continue
        arm = f"arm{arm_match.group(1)}"
        df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False).fillna("")
        df.columns = [normalize_column(col) for col in df.columns]
        for required_col in ("order", "required_session", "participant_event_name", "clinician_event_name"):
            if required_col not in df.columns:
                raise ValueError(f"{path.name} sheet {sheet} is missing {required_col}")
        for _, row in df.iterrows():
            session = canonical_session(row["required_session"])
            rows.append(
                {
                    "arm": arm,
                    "session": session,
                    "order": int(clean(row["order"]) or 999),
                    "participant_event_name": clean(row["participant_event_name"]),
                    "clinician_event_name": clean(row["clinician_event_name"]),
                }
            )
    out = pd.DataFrame(rows).drop_duplicates(["arm", "session"]).reset_index(drop=True)
    log.count("required_session_rows_loaded", len(out))
    return out


def expand_design_session_label(value: Any) -> list[str]:
    text = clean(value)
    explicit = {
        "T3 / T3 IE": ["T3", "IE T3"],
        "T6 / T6 IE": ["T6", "IE T6"],
        "T9 / T9 IE": ["T9", "IE T9"],
        "T12 / T12 IE": ["T12", "IE T12"],
        "Treatment Session 1, 2, 4, 5, 7, 8, 10, and 11": [
            "T1",
            "T2",
            "T4",
            "T5",
            "T7",
            "T8",
            "T10",
            "T11",
        ],
        "Treatment Session 7, 8, 10, and 11": ["T7", "T8", "T10", "T11"],
    }
    if text in explicit:
        return explicit[text]
    return [canonical_session(text)]


def load_expected_instruments(input_dir: Path, required_sessions: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    path = input_dir / DESIGN_FILES["expected_instruments"]
    if not path.exists():
        raise FileNotFoundError(f"Required design file not found: {path}")
    required_by_arm = {arm: set(group["session"]) for arm, group in required_sessions.groupby("arm", sort=False)}
    rows: list[dict[str, str]] = []
    workbook = pd.ExcelFile(path)
    for sheet in workbook.sheet_names:
        arm_match = re.search(r"(\d+)", sheet)
        if not arm_match:
            log.warn(f"Skipping instrument sheet without arm number: {sheet}")
            continue
        arm = f"arm{arm_match.group(1)}"
        df = pd.read_excel(path, sheet_name=sheet, dtype=str, keep_default_na=False).fillna("")
        df.columns = [normalize_column(col) for col in df.columns]
        for required_col in ("session", "source_of_instrument", "instrument", "included_in_summary"):
            if required_col not in df.columns:
                raise ValueError(f"{path.name} sheet {sheet} is missing {required_col}")
        excluded = 0
        for _, row in df.iterrows():
            if clean(row["included_in_summary"]) != "1":
                excluded += 1
                continue
            for session in expand_design_session_label(row["session"]):
                if not final_session_allowed(session):
                    continue
                if session not in required_by_arm.get(arm, set()):
                    log.warn(f"Included instrument not in required sessions: {arm} {session} {row['instrument']}")
                    continue
                rows.append(
                    {
                        "arm": arm,
                        "session": session,
                        "instrument": clean(row["instrument"]),
                        "instrument_source": clean(row["source_of_instrument"]),
                    }
                )
        log.info(f"{path.name} {sheet}: excluded {excluded} rows where Included_in_summary != 1")
    out = pd.DataFrame(rows).drop_duplicates().reset_index(drop=True)
    log.count("included_expected_instrument_rows", len(out))
    return out


def source_for_instrument(source_label: str) -> str:
    return "participants" if clean(source_label).lower().startswith("participant") else "clinician"


def event_is_direct(event: str) -> bool:
    value = clean(event)
    return bool(value and value != "-" and not value.startswith("—"))


def build_event_map(required_sessions: pd.DataFrame) -> dict[tuple[str, str], str]:
    mapping: dict[tuple[str, str], str] = {}
    for _, row in required_sessions.iterrows():
        for source, column in (("participants", "participant_event_name"), ("clinician", "clinician_event_name")):
            event = clean(row[column])
            if event_is_direct(event):
                mapping[(source, event)] = row["session"]
    return mapping


def event_to_session(source: str, event_name: Any, event_map: dict[tuple[str, str], str], log: ValidationLog) -> str:
    event = clean(event_name)
    if (source, event) in event_map:
        return canonical_session(event_map[(source, event)])
    base = re.sub(r"_arm_\d+$", "", event)
    if base in {"baseline", "baseline_1"}:
        return "Baseline"
    if base == "repeat_baseline":
        return "Repeat Baseline"
    if base == "screening":
        return "Screening"
    if base.startswith("asap_session"):
        return "ASAP"
    log.warn(f"Unknown REDCap event mapping: {source} {event}")
    return f"UNMAPPED_EVENT_{event}"


def load_redcap(input_dir: Path, required_sessions: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    frames = []
    event_map = build_event_map(required_sessions)
    for source, filename in INPUT_FILES.items():
        df = read_csv_strings(input_dir / filename, log)
        df.columns = [normalize_column(col) for col in df.columns]
        id_col = "record_id" if source == "participants" else "preescreen_id"
        if id_col not in df.columns:
            raise ValueError(f"{filename} is missing required subject ID column {id_col}")
        if "redcap_event_name" not in df.columns:
            raise ValueError(f"{filename} is missing redcap_event_name")
        df["source"] = source
        df["source_file"] = filename
        df["subject_id"] = df[id_col].map(standardize_subject_id)
        df["arm"] = df["subject_id"].map(infer_arm)
        df["session"] = df["redcap_event_name"].map(lambda event: event_to_session(source, event, event_map, log))
        frames.append(df)
    out = pd.concat(frames, ignore_index=True, sort=False).fillna("")
    bad_subjects = out[(out["subject_id"].str.len() != 3) | (out["arm"] == "unknown")]
    if not bad_subjects.empty:
        log.warn(f"{len(bad_subjects)} REDCap rows have nonstandard subject IDs or unknown arms")
    log.count("redcap_rows_loaded", len(out))
    return out


def build_subject_arm_map(redcap: pd.DataFrame, mri_subjects: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    redcap_subjects = redcap[["subject_id", "arm"]].drop_duplicates()
    all_subjects = pd.concat([redcap_subjects, mri_subjects], ignore_index=True).drop_duplicates()
    all_subjects = all_subjects[all_subjects["arm"].isin(ARMS)].copy()
    all_subjects = all_subjects.sort_values(["arm", "subject_id"]).reset_index(drop=True)
    log.count("subjects_loaded", len(all_subjects))
    return all_subjects


def instrument_config(source_label: str, instrument: str) -> dict[str, str | list[str]]:
    if source_for_instrument(source_label) == "participants":
        return PARTICIPANT_INSTRUMENTS.get(instrument, {})
    return CLINICIAN_INSTRUMENTS.get(instrument, {})


def build_questionnaire_long(
    redcap: pd.DataFrame,
    subjects: pd.DataFrame,
    expected: pd.DataFrame,
    log: ValidationLog,
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for _, exp in expected.iterrows():
        source = source_for_instrument(exp["instrument_source"])
        cfg = instrument_config(exp["instrument_source"], exp["instrument"])
        if not cfg:
            log.warn(f"No REDCap field mapping for expected instrument: {exp['arm']} {exp['session']} {exp['instrument']}")
        matches = redcap[
            (redcap["arm"] == exp["arm"])
            & (redcap["session"] == exp["session"])
            & (redcap["source"] == source)
        ]
        by_subject = {sid: group.iloc[0] for sid, group in matches.groupby("subject_id", sort=False)}
        for subject_id in subjects.loc[subjects["arm"] == exp["arm"], "subject_id"]:
            observed = by_subject.get(subject_id)
            date_field = clean(cfg.get("date", ""))
            date = clean(observed.get(date_field, "")) if observed is not None and date_field else ""
            status = status_from_fields(observed, cfg.get("complete"))
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


def build_session_dates(redcap: pd.DataFrame) -> dict[tuple[str, str], str]:
    date_fields = [
        "phq_9_date",
        "gad_7_date",
        "bddybocs_sr_date",
        "scid5_axis1_date",
        "eligibility_date",
        "peas_date_of_visit",
        "scan_date",
        "mri_exit_date",
    ]
    observed: dict[tuple[str, str], str] = {}
    for (subject_id, session), group in redcap.groupby(["subject_id", "session"], sort=False):
        dates = []
        for _, row in group.iterrows():
            for field in date_fields:
                value = clean(row.get(field, ""))
                if value:
                    dates.append(value)
        if dates:
            observed[(subject_id, session)] = dates[0]
    return observed


def interval_weeks(baseline_date: str, current_date: str) -> float | None:
    baseline = pd.to_datetime(clean(baseline_date), errors="coerce")
    current = pd.to_datetime(clean(current_date), errors="coerce")
    if pd.isna(baseline) or pd.isna(current):
        return None
    return round((current - baseline).days / 7, 2)


def build_behavioral_long(redcap: pd.DataFrame, subjects: pd.DataFrame, required_sessions: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    clinician = redcap[redcap["source"] == "clinician"].copy()
    scan_fields = sorted([col for col in clinician.columns if col.startswith("scan_run_") and col != "scan_run_sheet_complete"])
    for _, req in required_sessions.iterrows():
        arm = req["arm"]
        session = req["session"]
        if not final_session_allowed(session):
            continue
        matches = clinician[(clinician["arm"] == arm) & (clinician["session"] == session)]
        by_subject = {sid: group.iloc[0] for sid, group in matches.groupby("subject_id", sort=False)}
        for subject_id in subjects.loc[subjects["arm"] == arm, "subject_id"]:
            record = by_subject.get(subject_id)
            ant_field = ANT_FIELDS_BY_SESSION.get(session)
            if ant_field:
                rows.append(
                    {
                        "subject_id": subject_id,
                        "arm": arm,
                        "session": session,
                        "domain": "ANT",
                        "item": "ANT",
                        "status": status_from_fields(record, ant_field),
                        "qc_status": "",
                        "qc_pass": "",
                        "value": clean(record.get(ant_field, "")) if record is not None else "",
                    }
                )
            if record is not None:
                populated = [field for field in scan_fields if clean(record.get(field, ""))]
            else:
                populated = []
            run_fields = (populated or scan_fields)[:2]
            for idx in (1, 2):
                field = run_fields[idx - 1] if idx <= len(run_fields) else ""
                value = clean(record.get(field, "")) if record is not None and field else ""
                status = "missing" if not value or value.lower() in {"na", "n/a", "nd"} else "complete"
                if value and status != "complete":
                    status = "review_required"
                qc_status = "pass" if status == "complete" else ("missing" if status == "missing" else "review_required")
                rows.append(
                    {
                        "subject_id": subject_id,
                        "arm": arm,
                        "session": session,
                        "domain": "selfOthers",
                        "item": f"run{idx}",
                        "status": status,
                        "qc_status": qc_status,
                        "qc_pass": qc_status == "pass",
                        "value": value,
                    }
                )
            rows.append(
                {
                    "subject_id": subject_id,
                    "arm": arm,
                    "session": session,
                    "domain": "selfOthers",
                    "item": "Npractice",
                    "status": "",
                    "qc_status": "",
                    "qc_pass": "",
                    "value": "",
                }
            )
    return pd.DataFrame(rows)


def read_mri_subjects(input_dir: Path, log: ValidationLog) -> pd.DataFrame:
    rows = []
    for filename in MRI_FILES.values():
        path = input_dir / filename
        if not path.exists():
            continue
        df = pd.read_csv(path, dtype=str, keep_default_na=False)
        if "subID" not in df.columns:
            continue
        for raw in df["subID"]:
            subject_id = standardize_subject_id(raw)
            rows.append({"subject_id": subject_id, "arm": infer_arm(subject_id)})
    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame(columns=["subject_id", "arm"])
    return out.drop_duplicates()


def load_mri_long(input_dir: Path, subjects: pd.DataFrame, required_sessions: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for kind, filename in MRI_FILES.items():
        path = input_dir / filename
        if not path.exists():
            log.warn(f"MRI QC file missing: {path}")
            continue
        df = pd.read_csv(path, dtype=str, keep_default_na=False).fillna("")
        required = ["subID", "sesID", "modality", "Poor_Quality"]
        if kind == "func":
            required.extend(["taskID", "runID"])
        missing = [col for col in required if col not in df.columns]
        if missing:
            log.warn(f"{filename} missing required columns: {', '.join(missing)}")
            continue
        for _, row in df.iterrows():
            subject_id = standardize_subject_id(row["subID"])
            session = MRI_SESSION_MAP.get(clean(row["sesID"]).lower(), f"UNMAPPED_MRI_SESSION_{clean(row['sesID'])}")
            if session.startswith("UNMAPPED"):
                log.warn(f"Unmapped MRI sesID: {row['sesID']}")
            if kind == "anat":
                scan_or_run = "anat_T1w"
            else:
                task = clean(row["taskID"]).lower() or "task"
                run = clean(row["runID"]).zfill(2)
                scan_or_run = f"func_{task}_run{run}"
            qc_status = qc_status_from_poor_quality(row["Poor_Quality"])
            rows.append(
                {
                    "subject_id": subject_id,
                    "arm": infer_arm(subject_id),
                    "session": session,
                    "scan_or_run": scan_or_run,
                    "qc_status": qc_status,
                    "qc_pass": qc_status == "pass",
                }
            )
    observed = pd.DataFrame(rows)
    if observed.empty:
        log.count("mri_qc_rows_loaded", 0)
        return pd.DataFrame(columns=["subject_id", "arm", "session", "scan_or_run", "qc_status", "qc_pass"])
    log.count("mri_qc_rows_loaded", len(observed))
    allowed_sessions = set(required_sessions[required_sessions["session"].map(final_session_allowed)]["session"])
    observed = observed[observed["session"].isin(allowed_sessions)].copy()
    expectation_rows = []
    for (arm, session, scan_or_run), group in observed.groupby(["arm", "session", "scan_or_run"], sort=False):
        for subject_id in subjects.loc[subjects["arm"] == arm, "subject_id"]:
            match = group[group["subject_id"] == subject_id]
            if match.empty:
                expectation_rows.append(
                    {
                        "subject_id": subject_id,
                        "arm": arm,
                        "session": session,
                        "scan_or_run": scan_or_run,
                        "qc_status": "missing",
                        "qc_pass": False,
                    }
                )
            else:
                expectation_rows.extend(match.to_dict("records"))
    return pd.DataFrame(expectation_rows).drop_duplicates(["subject_id", "arm", "session", "scan_or_run"])


def status_counts(df: pd.DataFrame, status_col: str = "status") -> dict[str, int]:
    statuses = df[status_col].fillna("").map(clean)
    complete = int(statuses.isin(COMPLETE_STATUSES).sum())
    review = int(statuses.isin(REVIEW_STATUSES).sum())
    missing = int(statuses.isin(MISSING_STATUSES).sum())
    return {"complete": complete, "review_required": review, "missing": missing}


def rate(num: int, denom: int) -> float:
    return round(num / denom, 4) if denom else 0.0


def bool_text(value: bool) -> str:
    return "True" if value else "False"


def session_summary_values(
    subject_id: str,
    session: str,
    qn: pd.DataFrame,
    beh: pd.DataFrame,
    mri: pd.DataFrame,
) -> dict[str, Any]:
    qn_s = qn[(qn["subject_id"] == subject_id) & (qn["session"] == session)]
    ant_s = beh[(beh["subject_id"] == subject_id) & (beh["session"] == session) & (beh["domain"] == "ANT")]
    self_s = beh[(beh["subject_id"] == subject_id) & (beh["session"] == session) & (beh["domain"] == "selfOthers")]
    mri_s = mri[(mri["subject_id"] == subject_id) & (mri["session"] == session)]
    missing_tasks = []
    for _, row in qn_s.iterrows():
        if row["status"] != "complete":
            missing_tasks.append(row["instrument"])
    for _, row in ant_s.iterrows():
        if row["status"] != "complete":
            missing_tasks.append("ANT")
    session_completed = len(missing_tasks) == 0
    self_fail = [
        row["item"]
        for _, row in self_s.iterrows()
        if row["item"].startswith("run") and row["qc_status"] in {"fail", "missing", "review_required"}
    ]
    mri_fail = [
        row["scan_or_run"]
        for _, row in mri_s.iterrows()
        if row["qc_status"] in {"fail", "missing", "review_required"}
    ]
    self_expected = not self_s.empty
    mri_expected = not mri_s.empty
    self_pass = self_expected and not self_fail
    mri_pass = mri_expected and not mri_fail
    return {
        "session_completed": session_completed,
        "missingTask": "; ".join(missing_tasks),
        "selfOtherQC_passed": self_pass,
        "selfOtherQC_failed": "; ".join(self_fail),
        "MRIQC_passed": mri_pass,
        "MRIQC_failed": "; ".join(mri_fail),
        "QC_pass": session_completed and (not self_expected or self_pass) and (not mri_expected or mri_pass),
    }


def build_subject_wise(
    arm: str,
    subjects: pd.DataFrame,
    required_sessions: pd.DataFrame,
    qn: pd.DataFrame,
    beh: pd.DataFrame,
    mri: pd.DataFrame,
    redcap: pd.DataFrame,
) -> pd.DataFrame:
    arm_subjects = subjects[subjects["arm"] == arm]["subject_id"].tolist()
    arm_sessions = required_sessions[
        (required_sessions["arm"] == arm) & (required_sessions["session"].map(final_session_allowed))
    ].sort_values("order")
    observed_dates = build_session_dates(redcap)
    rows: list[dict[str, Any]] = []
    for subject_id in arm_subjects:
        row: dict[str, Any] = {"subject_id": subject_id, "arm": arm, "dropout_status": "active"}
        for session in arm_sessions["session"]:
            s_label = excel_safe_label(session)
            qn_s = qn[(qn["subject_id"] == subject_id) & (qn["session"] == session)].sort_values("instrument")
            for _, item in qn_s.iterrows():
                q_label = item_label(item["instrument"])
                row[f"ses-{s_label}_qn-{q_label}_date"] = item["date"]
                row[f"ses-{s_label}_qn-{q_label}_status"] = item["status"]
            ant_s = beh[
                (beh["subject_id"] == subject_id)
                & (beh["session"] == session)
                & (beh["domain"] == "ANT")
            ]
            if not ant_s.empty:
                row[f"ses-{s_label}_beh-ANT_status"] = ant_s.iloc[0]["status"]
            self_s = beh[
                (beh["subject_id"] == subject_id)
                & (beh["session"] == session)
                & (beh["domain"] == "selfOthers")
            ]
            npractice = self_s[self_s["item"] == "Npractice"]
            if not npractice.empty:
                row[f"ses-{s_label}_beh-selfOthers_Npractice"] = npractice.iloc[0]["value"]
            for run in ("run1", "run2"):
                run_s = self_s[self_s["item"] == run]
                if not run_s.empty:
                    run_row = run_s.iloc[0]
                    row[f"ses-{s_label}_beh-selfOthers_{run}_status"] = run_row["status"]
                    row[f"ses-{s_label}_beh-selfOthers_{run}_acc"] = ""
                    row[f"ses-{s_label}_beh-selfOthers_{run}_missingRate"] = ""
            mri_s = mri[(mri["subject_id"] == subject_id) & (mri["session"] == session)].sort_values("scan_or_run")
            for _, item in mri_s.iterrows():
                row[f"ses-{s_label}_mri-{item_label(item['scan_or_run'])}_qc_status"] = item["qc_status"]
            summary = session_summary_values(subject_id, session, qn, beh, mri)
            for key, value in summary.items():
                row[f"ses-{s_label}_{key}"] = bool_text(value) if isinstance(value, bool) else value
            if session != "Baseline":
                weeks = interval_weeks(
                    observed_dates.get((subject_id, "Baseline"), ""),
                    observed_dates.get((subject_id, session), ""),
                )
                row[f"ses-{s_label}_intervalFromBaseline_weeks"] = "" if weeks is None else weeks
                row[f"ses-{s_label}_interval_valid"] = "" if weeks is None else bool_text(weeks >= 0)
        subject_redcap = redcap[redcap["subject_id"] == subject_id]
        row["total_ASAP_count"] = int(subject_redcap["session"].map(lambda x: clean(x).upper().startswith("ASAP")).sum())
        session_flags = [
            row.get(f"ses-{excel_safe_label(session)}_session_completed") == "True"
            for session in arm_sessions["session"]
        ]
        missing_sessions = [
            session
            for session in arm_sessions["session"]
            if row.get(f"ses-{excel_safe_label(session)}_session_completed") != "True"
        ]
        row["complete_all_expected_experiment_sessions"] = bool_text(all(session_flags))
        row["Nof_missing_expected_experiment_sessions"] = len(missing_sessions)
        row["missing_expected_experiment_sessions"] = "; ".join(missing_sessions)
        row["complete_all_experiment_sessions"] = row["complete_all_expected_experiment_sessions"]
        qn_subject = qn[qn["subject_id"] == subject_id]
        ant_subject = beh[(beh["subject_id"] == subject_id) & (beh["domain"] == "ANT")]
        self_subject = beh[(beh["subject_id"] == subject_id) & (beh["domain"] == "selfOthers") & (beh["item"].str.startswith("run"))]
        mri_subject = mri[mri["subject_id"] == subject_id]
        row["complete_all_instrument"] = bool_text(not qn_subject.empty and (qn_subject["status"] == "complete").all())
        row["complete_all_ANT"] = bool_text(not ant_subject.empty and (ant_subject["status"] == "complete").all())
        row["all_MRI_QC_passed"] = bool_text(not mri_subject.empty and (mri_subject["qc_status"] == "pass").all())
        row["all_selfOther_QC_passed"] = bool_text(not self_subject.empty and (self_subject["qc_status"] == "pass").all())
        row["subject_QC_pass"] = bool_text(
            row["complete_all_expected_experiment_sessions"] == "True"
            and row["complete_all_instrument"] == "True"
            and row["complete_all_ANT"] == "True"
            and row["all_MRI_QC_passed"] == "True"
            and row["all_selfOther_QC_passed"] == "True"
        )
        rows.append(row)
    base_cols = ["subject_id", "arm", "dropout_status"]
    summary_cols = [
        "total_ASAP_count",
        "complete_all_expected_experiment_sessions",
        "Nof_missing_expected_experiment_sessions",
        "missing_expected_experiment_sessions",
        "complete_all_experiment_sessions",
        "complete_all_instrument",
        "complete_all_ANT",
        "all_MRI_QC_passed",
        "all_selfOther_QC_passed",
        "subject_QC_pass",
    ]
    if not rows:
        return pd.DataFrame(columns=base_cols + summary_cols)
    out = pd.DataFrame(rows).fillna("")
    other_cols = [col for col in out.columns if col not in base_cols + summary_cols]
    return out[base_cols + other_cols + summary_cols]


def table_session_instrument_summary(arm: str, qn: pd.DataFrame) -> pd.DataFrame:
    rows = []
    arm_qn = qn[qn["arm"] == arm]
    for (session, instrument, source), group in arm_qn.groupby(["session", "instrument", "instrument_source"], sort=False):
        counts = status_counts(group)
        expected = len(group)
        rows.append(
            {
                "arm": arm,
                "session": session,
                "instrument": instrument,
                "instrument_source": source,
                "expected_NofSubjects": expected,
                "complete_NofSubjects": counts["complete"],
                "review_required_NofSubjects": counts["review_required"],
                "complete_rate": rate(counts["complete"], expected),
                "missing_rate": rate(counts["missing"], expected),
            }
        )
    columns = [
        "arm",
        "session",
        "instrument",
        "instrument_source",
        "expected_NofSubjects",
        "complete_NofSubjects",
        "review_required_NofSubjects",
        "complete_rate",
        "missing_rate",
    ]
    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows, columns=columns).sort_values(
        ["session", "instrument"],
        key=lambda col: col.map(session_sort_key) if col.name == "session" else col,
    )


def table_instrument_summary(arm: str, qn: pd.DataFrame) -> pd.DataFrame:
    rows = []
    arm_qn = qn[qn["arm"] == arm]
    for (instrument, source), group in arm_qn.groupby(["instrument", "instrument_source"], sort=False):
        per_subject = group.groupby("subject_id")["status"].apply(lambda s: "complete" if (s == "complete").all() else combined_status(s))
        complete = int((per_subject == "complete").sum())
        review = int(per_subject.isin(REVIEW_STATUSES).sum())
        expected = len(per_subject)
        missing = expected - complete - review
        rows.append(
            {
                "arm": arm,
                "instrument_or_item": instrument,
                "instrument_source": source,
                "expected_NofSubjects": expected,
                "complete_NofRecords": complete,
                "review_required_NofRecords": review,
                "complete_rate": rate(complete, expected),
                "missing_rate": rate(missing, expected),
            }
        )
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
    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows, columns=columns).sort_values(["instrument_or_item"])


def table_ant_by_session(arm: str, beh: pd.DataFrame) -> pd.DataFrame:
    rows = []
    ant = beh[(beh["arm"] == arm) & (beh["domain"] == "ANT")]
    for session, group in ant.groupby("session", sort=False):
        counts = status_counts(group)
        expected = len(group)
        rows.append(
            {
                "arm": arm,
                "session": session,
                "expected_NofSubjects": expected,
                "complete_NofSubjects": counts["complete"],
                "missing_NofSubjects": counts["missing"],
                "review_required_NofSubjects": counts["review_required"],
                "complete_rate": rate(counts["complete"], expected),
                "missing_rate": rate(counts["missing"], expected),
            }
        )
    columns = [
        "arm",
        "session",
        "expected_NofSubjects",
        "complete_NofSubjects",
        "missing_NofSubjects",
        "review_required_NofSubjects",
        "complete_rate",
        "missing_rate",
    ]
    if not rows:
        return pd.DataFrame(columns=columns)
    return pd.DataFrame(rows, columns=columns).sort_values("session", key=lambda col: col.map(session_sort_key))


def table_ant_by_subject(arm: str, beh: pd.DataFrame) -> pd.DataFrame:
    ant = beh[(beh["arm"] == arm) & (beh["domain"] == "ANT")]
    per_subject = ant.groupby("subject_id")["status"].apply(lambda s: "complete" if (s == "complete").all() else combined_status(s))
    expected = len(per_subject)
    complete = int((per_subject == "complete").sum())
    review = int(per_subject.isin(REVIEW_STATUSES).sum())
    missing = expected - complete - review
    return pd.DataFrame(
        [
            {
                "arm": arm,
                "expected_NofSubjects": expected,
                "complete_NofSubjects": complete,
                "missing_NofSubjects": missing,
                "review_required_NofSubjects": review,
                "complete_rate": rate(complete, expected),
                "missing_rate": rate(missing, expected),
            }
        ]
    )


def qc_table_by_session(arm: str, df: pd.DataFrame, source: str) -> pd.DataFrame:
    rows = []
    data = df[df["arm"] == arm]
    for session, session_group in data.groupby("session", sort=False):
        per_subject = session_group.groupby("subject_id")["qc_status"].apply(
            lambda s: "pass" if (s == "pass").all() else ("review_required" if (s == "review_required").any() else ("missing" if (s == "missing").any() else "fail"))
        )
        expected = len(per_subject)
        passed = int((per_subject == "pass").sum())
        failed = int((per_subject == "fail").sum())
        missing = int((per_subject == "missing").sum())
        review = int((per_subject == "review_required").sum())
        rows.append(
            {
                "arm": arm,
                "session": session,
                "expected_NofSubjects": expected,
                "qc_pass_NofSubjects": passed,
                "qc_fail_NofSubjects": failed,
                "missing_NofSubjects": missing,
                "review_required_NofSubjects": review,
                "qc_pass_rate": rate(passed, expected),
                "missing_rate": rate(missing, expected),
            }
        )
    columns = [
        "arm",
        "session",
        "expected_NofSubjects",
        "qc_pass_NofSubjects",
        "qc_fail_NofSubjects",
        "missing_NofSubjects",
        "review_required_NofSubjects",
        "qc_pass_rate",
        "missing_rate",
    ]
    out = pd.DataFrame(rows, columns=columns)
    if out.empty:
        return out
    return out.sort_values("session", key=lambda col: col.map(session_sort_key))


def qc_table_by_subject(arm: str, df: pd.DataFrame, label_col: str | None = None) -> pd.DataFrame:
    rows = []
    data = df[df["arm"] == arm]
    group_cols = [label_col] if label_col else []
    for key, group in data.groupby(group_cols, sort=False) if group_cols else [(None, data)]:
        per_subject = group.groupby("subject_id")["qc_status"].apply(
            lambda s: "pass" if (s == "pass").all() else ("review_required" if (s == "review_required").any() else ("missing" if (s == "missing").any() else "fail"))
        )
        expected = len(per_subject)
        passed = int((per_subject == "pass").sum())
        failed = int((per_subject == "fail").sum())
        missing = int((per_subject == "missing").sum())
        review = int((per_subject == "review_required").sum())
        row = {
            "arm": arm,
            "expected_NofSubjects": expected,
            "qc_pass_NofSubjects": passed,
            "qc_fail_NofSubjects": failed,
            "missing_NofSubjects": missing,
            "review_required_NofSubjects": review,
            "qc_pass_rate": rate(passed, expected),
            "missing_rate": rate(missing, expected),
        }
        if label_col:
            row[label_col] = key
        rows.append(row)
    columns = [
        "arm",
        "expected_NofSubjects",
        "qc_pass_NofSubjects",
        "qc_fail_NofSubjects",
        "missing_NofSubjects",
        "review_required_NofSubjects",
        "qc_pass_rate",
        "missing_rate",
    ]
    if label_col:
        columns.insert(1, label_col)
    out = pd.DataFrame(rows, columns=columns)
    if label_col and not out.empty:
        cols = ["arm", label_col] + [col for col in out.columns if col not in {"arm", label_col}]
        out = out[cols]
    return out


def table_interval_summary(arm: str, subject_wise: pd.DataFrame, sessions: Iterable[str]) -> pd.DataFrame:
    rows = []
    for session in sessions:
        if session == "Baseline":
            continue
        label = excel_safe_label(session)
        completed_col = f"ses-{label}_session_completed"
        interval_col = f"ses-{label}_intervalFromBaseline_weeks"
        valid_col = f"ses-{label}_interval_valid"
        if completed_col not in subject_wise.columns:
            continue
        intervals = pd.to_numeric(subject_wise.get(interval_col, pd.Series(dtype=str)), errors="coerce")
        valid = subject_wise.get(valid_col, pd.Series([""] * len(subject_wise))).map(clean) == "True"
        completed = subject_wise[completed_col].map(clean) == "True"
        rows.append(
            {
                "arm": arm,
                "session": session,
                "completed_NofSubjects": int(completed.sum()),
                "valid_interval_NofSubjects": int(valid.sum()),
                "invalid_interval_NofSubjects": int((completed & ~valid).sum()),
                "mean_intervalFromBaseline_weeks": round(float(intervals[valid].mean()), 2) if valid.any() else "",
                "sd_intervalFromBaseline_weeks": round(float(intervals[valid].std()), 2) if valid.sum() > 1 else "",
                "min_intervalFromBaseline_weeks": round(float(intervals[valid].min()), 2) if valid.any() else "",
                "max_intervalFromBaseline_weeks": round(float(intervals[valid].max()), 2) if valid.any() else "",
            }
        )
    return pd.DataFrame(rows)


def table_readiness(arm: str, subject_wise: pd.DataFrame) -> pd.DataFrame:
    total = len(subject_wise)
    count_true = lambda col: int((subject_wise[col].map(clean) == "True").sum()) if col in subject_wise else 0
    qc_pass = count_true("subject_QC_pass")
    return pd.DataFrame(
        [
            {
                "arm": arm,
                "total_NofSubjects": total,
                "withdrawn_or_dropout_NofSubjects": 0,
                "withdrawn_or_dropout_rate": 0,
                "complete_all_experiment_sessions_NofSubjects": count_true("complete_all_experiment_sessions"),
                "complete_all_experiment_sessions_rate": rate(count_true("complete_all_experiment_sessions"), total),
                "complete_all_instrument_NofSubjects": count_true("complete_all_instrument"),
                "complete_all_instrument_rate": rate(count_true("complete_all_instrument"), total),
                "complete_all_ANT_NofSubjects": count_true("complete_all_ANT"),
                "complete_all_ANT_rate": rate(count_true("complete_all_ANT"), total),
                "all_MRI_QC_passed_NofSubjects": count_true("all_MRI_QC_passed"),
                "all_MRI_QC_passed_rate": rate(count_true("all_MRI_QC_passed"), total),
                "all_selfOther_QC_passed_NofSubjects": count_true("all_selfOther_QC_passed"),
                "all_selfOther_QC_passed_rate": rate(count_true("all_selfOther_QC_passed"), total),
                "QC_pass_NofSubjects_rate": qc_pass,
                "QC_passrate": rate(qc_pass, total),
            }
        ]
    )


def build_group_tables(
    arm: str,
    qn: pd.DataFrame,
    beh: pd.DataFrame,
    mri: pd.DataFrame,
    subject_wise: pd.DataFrame,
    arm_sessions: list[str],
) -> list[tuple[str, pd.DataFrame]]:
    self_qc = beh[(beh["domain"] == "selfOthers") & (beh["item"].str.startswith("run"))].copy()
    return [
        ("Table 1. Session-wise Summary", table_session_instrument_summary(arm, qn)),
        ("Table 2. Instrument-wise Summary", table_instrument_summary(arm, qn)),
        ("Table 3-1. ANT task complete rate by session", table_ant_by_session(arm, beh)),
        ("Table 3-2. ANT task complete rate by subject", table_ant_by_subject(arm, beh)),
        ("Table 4-1. SelfOthers QC pass rate by session", qc_table_by_session(arm, self_qc, "selfOthers")),
        ("Table 4-2. SelfOthers QC pass rate by subject", qc_table_by_subject(arm, self_qc)),
        ("Table 5. MRI QC pass rate by session", qc_table_by_session(arm, mri, "MRI")),
        ("Table 5-2. MRI QC pass rate by subject", qc_table_by_subject(arm, mri, "scan_or_run")),
        ("Table 6. Session interval summary after Baseline", table_interval_summary(arm, subject_wise, arm_sessions)),
        ("Table 7. Participant-level QA readiness summary", table_readiness(arm, subject_wise)),
    ]


def write_group_sheet(writer: pd.ExcelWriter, tables: list[tuple[str, pd.DataFrame]]) -> None:
    sheet_name = "group_wise"
    startrow = 0
    for title, table in tables:
        pd.DataFrame([[title]]).to_excel(writer, sheet_name=sheet_name, index=False, header=False, startrow=startrow)
        startrow += 1
        table.to_excel(writer, sheet_name=sheet_name, index=False, startrow=startrow)
        startrow += len(table) + 3


def validate_final_outputs(path: Path, subject_wise: pd.DataFrame, tables: list[tuple[str, pd.DataFrame]]) -> None:
    if subject_wise["subject_id"].duplicated().any():
        raise ValueError(f"{path.name}: subject_wise has duplicate subject_id rows")
    if len(subject_wise.columns) != len(set(subject_wise.columns)):
        raise ValueError(f"{path.name}: subject_wise has duplicate columns")
    bad_tokens = ["_arm_", "UNMAPPED_EVENT_", "UNMAPPED_MRI_SESSION_", "Baseline1", "Baseline 1", "Screening"]
    joined_cols = "\n".join(subject_wise.columns)
    for token in bad_tokens:
        if token in joined_cols:
            raise ValueError(f"{path.name}: final subject_wise columns contain forbidden token {token}")
    if len(tables) < 10:
        raise ValueError(f"{path.name}: group_wise is missing required tables")


def check_input_sidecars(input_dir: Path, log: ValidationLog) -> None:
    for filename in CODEBOOK_FILES:
        path = input_dir / filename
        if path.exists():
            log.info(f"Found REDCap codebook: {filename}")
        else:
            log.warn(f"REDCap codebook missing: {filename}")


def build_outputs(input_dir: Path, output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    log = ValidationLog()
    check_input_sidecars(input_dir, log)
    required_sessions = load_required_sessions(input_dir, log)
    expected = load_expected_instruments(input_dir, required_sessions, log)
    redcap = load_redcap(input_dir, required_sessions, log)
    mri_subjects = read_mri_subjects(input_dir, log)
    subjects = build_subject_arm_map(redcap, mri_subjects, log)
    qn = build_questionnaire_long(redcap, subjects, expected, log)
    beh = build_behavioral_long(redcap, subjects, required_sessions)
    mri = load_mri_long(input_dir, subjects, required_sessions, log)
    log.count("behavioral_rows_built", len(beh))
    for arm in ARMS:
        arm_sessions = (
            required_sessions[
                (required_sessions["arm"] == arm) & (required_sessions["session"].map(final_session_allowed))
            ]
            .sort_values("order")["session"]
            .tolist()
        )
        subject_wise = build_subject_wise(arm, subjects, required_sessions, qn, beh, mri, redcap)
        tables = build_group_tables(arm, qn, beh, mri, subject_wise, arm_sessions)
        out_path = output_dir / f"QA_summary_{arm}.xlsx"
        validate_final_outputs(out_path, subject_wise, tables)
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            subject_wise.to_excel(writer, sheet_name="subject_wise", index=False)
            write_group_sheet(writer, tables)
        log.info(f"Wrote {out_path}")
    log.write(output_dir / "QA_summary_validation_log.txt")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input-dir", type=Path, default=Path("data"), help="Directory containing QA inputs")
    parser.add_argument("--output-dir", type=Path, default=Path("results"), help="Directory for generated workbooks")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    build_outputs(args.input_dir, args.output_dir)


if __name__ == "__main__":
    main()
