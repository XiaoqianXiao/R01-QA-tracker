#!/usr/bin/env python3
"""Create project-level QA summaries from REDCap QA tracker exports.

This script is intentionally configuration-first. REDCap exports often contain
rows or missing rows that are not meaningful without study-design rules, so
expectedness is defined in CONFIG and never inferred from absence alone.
"""

from __future__ import annotations

import argparse
import math
import re
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd


# =============================================================================
# CONFIG
# =============================================================================

CONFIG: dict[str, Any] = {
    "design_files": {
        "required_sessions_by_arm": "Required_Sessions_for_each_Arm.xlsx",
        "instruments_by_session_by_arm": "Instruments_in_each_Session_each_Arm.xlsx",
    },
    "output_by_arm": True,
    "output_file_template": "QA_summary_{arm}.xlsx",
    "arms": ["arm1", "arm2", "arm3"],
    "subject_id_source_columns": {
        "ParticipantsQAtracker.csv": "record_id",
        "ClinicianQAtracker.csv": "preescreen_id",
        "QC_anat.csv": "subID",
        "QC_func.csv": "subID",
        "participants": "record_id",
        "clinician": "preescreen_id",
        "mriqc_anat": "subID",
        "mriqc_func": "subID",
    },
    "subject_id_standardization": {
        "mri_prefix": "sub-",
        "expected_digits": 3,
    },
    "arm_assignment": {
        "method": "standardized_subject_id_first_digit",
        "arm_digit_map": {"1": "arm1", "2": "arm2", "3": "arm3"},
    },
    "debug_subject_id_candidate_fields": [
        "record_id",
        "participant_id",
        "subject_id",
        "study_id",
        "preescreen_id",
    ],
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
    "event_mapping_by_source": {
        "participants": {
            "treatment_session_arm_1": "Treatment Session 1",
            "treatment_session_arm_1b": "Treatment Session 2",
            "treatment_session_arm_1c": "Treatment Session 3 (IE)",
            "treatment_session_arm_1d": "Treatment Session 4",
            "treatment_session_arm_1e": "Treatment Session 5",
            "treatment_session_arm_1f": "Treatment Session 6 (IE)",
            "treatment_session_arm_1g": "Treatment Session 7",
            "treatment_session_arm_1h": "Treatment Session 8",
            "treatment_session_arm_1i": "Treatment Session 9 (IE)",
            "treatment_session_arm_1j": "Treatment Session 10",
            "treatment_session_arm_1k": "Treatment Session 11",
            "treatment_session_arm_1l": "Treatment Session 12 (IE)",
        },
        "clinician": {
            "treatment_session_arm_1": "Treatment Session 1",
            "treatment_session_arm_1b": "Treatment Session 2",
            "treatment_session_arm_1c": "Treatment Session 3",
            "treatment_session_arm_1d": "Treatment Session 3 (IE)",
            "treatment_session_arm_1e": "Treatment Session 4",
            "treatment_session_arm_1f": "Treatment Session 5",
            "treatment_session_arm_1g": "Treatment Session 6",
            "treatment_session_arm_1h": "Treatment Session 6 (IE)",
            "treatment_session_arm_1i": "Treatment Session 7",
            "treatment_session_arm_1j": "Treatment Session 8",
            "treatment_session_arm_1k": "Treatment Session 9",
            "treatment_session_arm_1l": "Treatment Session 9 (IE)",
            "treatment_session_arm_1m": "Treatment Session 10",
            "treatment_session_arm_1n": "Treatment Session 11",
            "treatment_session_arm_1o": "Treatment Session 12",
            "treatment_session_arm_1p": "Treatment Session 12 (IE)",
            "asap_session_arm_1": "ASAP Session",
            "asap_session_2_arm_1": "ASAP Session 2",
            "asap_session_3_arm_1": "ASAP Session 3",
            "asap_session_4_arm_1": "ASAP Session 4",
            "asap_session_5_arm_1": "ASAP Session 5",
        },
    },
    "session_aliases": {
        "Baseline 1": "Baseline",
        "Baseline1": "Baseline",
        "baseline 1": "Baseline",
        "T3 IE": "IE T3",
        "T6 IE": "IE T6",
        "T9 IE": "IE T9",
        "T12 IE": "IE T12",
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
    "expected_sessions": [
        # "Baseline",
        # "T1",
        # "T2",
    ],
    "expected_questionnaires_by_session": {
        # "Baseline": ["QuestionnaireA", "QuestionnaireB"],
    },
    "must_have_questionnaires_by_session": {
        # "Baseline": ["QuestionnaireA"],
    },
    "optional_questionnaires_by_session": {
        # "Baseline": ["QuestionnaireB"],
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
    "mri_qc_fields": {
        # "T6 Scan": {
        #     "anatomical": "anat_qc_field",
        #     "selfother_run1": "selfother_run1_qc_field",
        # }
    },
    "mri_qc_pass_values": [
        "pass",
        "passed",
        "usable",
        "1",
        "yes",
    ],
    "mri_qc_fail_values": [
        "fail",
        "failed",
        "unusable",
        "0",
        "no",
    ],
    "task_completion_fields": {
        # "T6": {
        #     "selfOther": "selfother_task_complete",
        # }
    },
    "task_completion_complete_values": [
        "complete",
        "completed",
        "done",
        "yes",
        "y",
        "2",
    ],
    "task_completion_incomplete_values": [
        "incomplete",
        "not_complete",
        "not completed",
        "no",
        "n",
        "0",
    ],
    "task_completion_unverified_values": [
        "partial",
        "partially complete",
        "unverified",
        "1",
    ],
    "behavioral_qc_fields": {
        # "T6": {
        #     "selfOther": "selfother_behavioral_qc_field",
        # }
    },
    "behavioral_qc_pass_values": [
        "pass",
        "passed",
        "usable",
        "1",
        "yes",
    ],
    "behavioral_qc_fail_values": [
        "fail",
        "failed",
        "unusable",
        "0",
        "no",
    ],
    "dropout_fields": [
        # "dropout_status",
        # "withdrawal_status",
    ],
    "dropout_values": [
        "dropout",
        "withdrawn",
        "withdrew",
        "yes",
        "1",
    ],
    "asap_fields": [
        # "asap_flag",
        # "asap_count",
    ],
    "session_timing_windows": {
        # "T1": {
        #     "target_weeks": 1,
        #     "allowed_min_weeks": 0,
        #     "allowed_max_weeks": 2,
        # }
    },
    "date_fields_by_questionnaire": {
        # "QuestionnaireA": "questionnaire_a_date",
    },
    "completion_status_fields_by_form": {
        # "questionnaire_a": "questionnaire_a_complete",
    },
    "baseline_session_name": "Baseline",
    "exclude_from_interval_sessions": [
        "Screening",
        "Baseline",
    ],
    "interval_sessions": [
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
    "session_date_fields": {
        # "Baseline": "baseline_date",
        # "T1": "t1_date",
    },
    "treat_observed_as_expected_when_no_design_config": False,
}


INPUT_FILES = {
    "participants": "ParticipantsQAtracker.csv",
    "clinician": "ClinicianQAtracker.csv",
}

CODEBOOK_FILES = {
    "participants": "Participants_REDCap.pdf",
    "clinician": "Clini_REDCap.pdf",
}

OUTPUT_XLSX = "QA_summary.xlsx"
VALIDATION_LOG = "QA_summary_validation_log.txt"

ALLOWED_QUESTIONNAIRE_STATUSES = {
    "complete",
    "incomplete",
    "missing",
    "unverified",
    "not_expected",
    "review_required",
}

ALLOWED_QC_STATUSES = {
    "pass",
    "fail",
    "missing",
    "unverified",
    "not_expected",
    "review_required",
}

ALLOWED_TASK_COMPLETION_STATUSES = {
    "complete",
    "incomplete",
    "missing",
    "unverified",
    "not_expected",
    "review_required",
}


@dataclass
class ValidationLog:
    lines: list[str] = field(default_factory=list)
    warnings: dict[str, list[str]] = field(default_factory=lambda: defaultdict(list))
    counters: dict[str, Any] = field(default_factory=dict)

    def info(self, message: str) -> None:
        self.lines.append(message)

    def warn(self, category: str, message: str) -> None:
        self.warnings[category].append(message)

    def set_counter(self, key: str, value: Any) -> None:
        self.counters[key] = value

    def write(self, path: Path) -> None:
        content: list[str] = ["QA Summary Validation Log", "=" * 25, ""]
        content.extend(self.lines)
        content.append("")
        content.append("Counters")
        content.append("-" * 8)
        for key in sorted(self.counters):
            content.append(f"{key}: {self.counters[key]}")
        content.append("")
        content.append("Warnings")
        content.append("-" * 8)
        if not self.warnings:
            content.append("No warnings recorded.")
        for category in sorted(self.warnings):
            content.append(f"[{category}]")
            for message in self.warnings[category]:
                content.append(f"- {message}")
        path.write_text("\n".join(content) + "\n", encoding="utf-8")


def clean_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def normalize_column_name(name: str) -> str:
    normalized = str(name).strip().lower()
    normalized = re.sub(r"[^0-9a-zA-Z]+", "_", normalized)
    normalized = re.sub(r"_+", "_", normalized).strip("_")
    return normalized


def normalize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_column_name(col) for col in df.columns]
    return df


def canonical_session_name(session: str) -> str:
    value = clean_value(session)
    aliases = CONFIG.get("session_aliases", {})
    return aliases.get(value, value)


def canonical_session_mapping(mapping: dict[str, Any]) -> dict[str, Any]:
    return {canonical_session_name(session): value for session, value in mapping.items()}


def filter_excluded_output_sessions(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty or "session" not in df.columns:
        return df
    excluded = {canonical_session_name(session) for session in CONFIG.get("excluded_output_sessions", [])}
    return df[~df["session"].map(canonical_session_name).isin(excluded)].copy()


def session_sort_key(session: str) -> int:
    session = canonical_session_name(session)
    order = CONFIG.get("session_order", [])
    try:
        return order.index(session)
    except ValueError:
        return len(order) + 999


def sort_by_session_order(df: pd.DataFrame, sort_cols: list[str]) -> pd.DataFrame:
    if df.empty or "session" not in df.columns:
        return df
    df = df.copy()
    df["_session_order"] = df["session"].map(session_sort_key)
    actual_sort_cols = ["_session_order"] + [
        col for col in sort_cols if col in df.columns and col != "session"
    ]
    df = df.sort_values(actual_sort_cols, kind="mergesort").drop(columns=["_session_order"])
    return df.reset_index(drop=True)


def read_csv_as_string(path: Path, source_file: str, log: ValidationLog) -> pd.DataFrame:
    if not path.exists():
        log.warn("missing_input_file", f"{path} was not found.")
        return pd.DataFrame()
    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    df = normalize_column_names(df)
    df["source_file"] = source_file
    df["instrument_source"] = source_file
    log.info(f"Input file: {path}")
    log.set_counter(f"rows_loaded_{source_file}", len(df))
    return df


def read_mri_qc_csv(
    path: Path,
    source_file: str,
    required_columns: list[str],
    log: ValidationLog,
) -> pd.DataFrame:
    if not path.exists():
        log.warn("missing_mri_qc_file", f"{path} was not found.")
        return pd.DataFrame()

    df = pd.read_csv(path, dtype=str, keep_default_na=False)
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        log.warn(
            "missing_mri_qc_columns",
            f"{path.name} is missing required columns: {missing}",
        )
        return pd.DataFrame()

    df = df.copy()
    df["source_file"] = source_file
    df["instrument_source"] = "mriqc"
    log.info(f"MRI QC input file: {path}")
    log.set_counter(f"rows_loaded_{source_file}", len(df))
    return df


def choose_most_complete_field(
    df: pd.DataFrame,
    candidate_fields: list[str],
    label: str,
    log: ValidationLog,
) -> str | None:
    available = [field for field in candidate_fields if field in df.columns]
    if not available:
        log.warn("missing_expected_fields", f"No {label} candidate fields found: {candidate_fields}")
        return None
    completeness = {
        field: df[field].map(lambda value: clean_value(value) != "").sum()
        for field in available
    }
    chosen = max(completeness, key=completeness.get)
    log.info(f"Detected {label} field: {chosen} from candidates {available}")
    if len(available) > 1:
        log.info(f"{label} field completeness: {completeness}")
    return chosen


def get_subject_id_source_column(source_file: str) -> str:
    mapping = CONFIG.get("subject_id_source_columns", {})
    if source_file in mapping:
        return normalize_column_name(mapping[source_file])
    raise KeyError(f"No configured subject ID source column for {source_file}")


def standardize_subject_id(raw_value: Any, source_file: str, log: ValidationLog | None = None) -> str:
    raw = clean_value(raw_value)
    if not raw:
        return ""

    source_label = clean_value(source_file)
    mri_sources = {"QC_anat.csv", "QC_func.csv", "mriqc_anat", "mriqc_func"}
    if source_label in mri_sources:
        prefix = CONFIG.get("subject_id_standardization", {}).get("mri_prefix", "sub-")
        if raw.lower().startswith(str(prefix).lower()):
            raw = raw[len(str(prefix)):]
        elif log is not None:
            log.warn("nonstandard_mri_subID", f"{source_label}: {raw}")

    digits = "".join(ch for ch in raw if ch.isdigit())
    expected_digits = int(CONFIG.get("subject_id_standardization", {}).get("expected_digits", 3))
    if len(digits) >= expected_digits:
        return digits[-expected_digits:]
    return digits


def standardize_subject_ids_for_source(df: pd.DataFrame, source_file: str, log: ValidationLog) -> tuple[pd.DataFrame, str | None]:
    df = df.copy()
    try:
        field_name = get_subject_id_source_column(source_file)
    except KeyError as exc:
        log.warn("missing_subject_id_source_config", str(exc))
        df["subject_id"] = ""
        return df, None

    if field_name not in df.columns:
        log.warn(
            "missing_subject_id_source_columns",
            f"{source_file} is missing required subject ID source column: {field_name}",
        )
        df["subject_id"] = ""
        return df, field_name

    df["subject_id"] = df[field_name].map(lambda value: standardize_subject_id(value, source_file, log))
    missing_rows = df.index[df["subject_id"] == ""].tolist()
    if missing_rows:
        log.warn("missing_subject_ids", f"{source_file}: {len(missing_rows)} rows have missing subject IDs.")
    if (df["subject_id"].map(lambda value: clean_value(value).lower().startswith("sub-"))).any():
        log.warn("subject_id_standardization", f"{source_file}: standardized subject_id still contains sub- prefix.")
    examples = (
        df[[field_name, "subject_id"]]
        .drop_duplicates()
        .head(5)
        .to_dict("records")
    )
    log.info(f"subject_id_source_columns_valid: {source_file} uses {field_name}")
    log.info(f"standardized_subject_id_examples_{source_file}: {examples}")
    log.set_counter(f"detected_subject_id_field_{source_file}", field_name)
    return df, field_name


def standardize_subject_ids(df: pd.DataFrame, log: ValidationLog) -> tuple[pd.DataFrame, str | None]:
    if df.empty:
        df = df.copy()
        df["subject_id"] = ""
        return df, None
    frames: list[pd.DataFrame] = []
    fields: list[str] = []
    for source_file, group in df.groupby("source_file", dropna=False, sort=False):
        standardized, field_name = standardize_subject_ids_for_source(group, clean_value(source_file), log)
        frames.append(standardized)
        if field_name:
            fields.append(f"{source_file}:{field_name}")
    return pd.concat(frames, ignore_index=True, sort=False).fillna(""), ";".join(fields) or None

def infer_arm_from_subject_id(subject_id: str) -> str:
    sid = clean_value(subject_id)
    expected_digits = int(CONFIG.get("subject_id_standardization", {}).get("expected_digits", 3))
    if len(sid) < expected_digits:
        return "unknown"
    arm_digit = sid[0]
    arm_map = CONFIG.get("arm_assignment", {}).get("arm_digit_map", {})
    return arm_map.get(arm_digit, "unknown")


def add_arm_assignment(df: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    df["arm"] = df["subject_id"].map(infer_arm_from_subject_id)
    unknown_subjects = sorted(set(df.loc[df["arm"] == "unknown", "subject_id"].map(clean_value)) - {""})
    for subject_id in unknown_subjects:
        log.warn("unknown_subject_arm", f"subject_id={subject_id}")
    for arm in CONFIG.get("arms", []):
        log.set_counter(f"subjects_assigned_{arm}", df.loc[df["arm"] == arm, "subject_id"].nunique())
    log.set_counter("subjects_assigned_unknown_arm", len(unknown_subjects))
    return df


def design_file_path(input_dir: Path, filename: str) -> Path:
    direct_path = input_dir / filename
    if direct_path.exists():
        return direct_path
    data_path = input_dir / "data" / filename
    if data_path.exists():
        return data_path
    return direct_path


def normalize_design_arm(value: str) -> str:
    normalized = clean_value(value).lower().replace(" ", "")
    if normalized in {"arm1", "1"}:
        return "arm1"
    if normalized in {"arm2", "2"}:
        return "arm2"
    if normalized in {"arm3", "3"}:
        return "arm3"
    match = re.search(r"arm\s*([123])", clean_value(value), flags=re.I)
    return f"arm{match.group(1)}" if match else "unknown"


def normalize_design_session(value: str) -> str:
    session = clean_value(value)
    if not session:
        return ""
    session = re.sub(r"\s+", " ", session).strip()
    aliases = CONFIG.get("session_aliases", {})
    if session in aliases:
        return aliases[session]
    for alias, canonical in aliases.items():
        if session.lower() == str(alias).lower():
            return canonical
    return canonical_session_name(session)


def normalize_design_instrument(value: str) -> str:
    instrument = clean_value(value)
    instrument = re.sub(r"\s+", " ", instrument).strip()
    return instrument


def included_in_summary(value: Any) -> bool:
    return clean_value(value).lower() in {"1", "true", "yes", "y"}


def instrument_match_key(value: str) -> str:
    return re.sub(r"[^0-9a-z]+", "", clean_value(value).lower())


def infer_instrument_type(instrument: str, instrument_source: str, log: ValidationLog) -> str:
    normalized = instrument_match_key(instrument)
    if normalized in {"scanrunsheet", "mriexitquestionnaire"}:
        log.warn(
            "unknown_instrument_type_assumptions",
            f"{instrument} treated as questionnaire/session-status design item, not external MRI QC.",
        )
        return "questionnaire"
    if "task" in normalized and "checklist" not in normalized:
        return "task_completion"
    if "qc" in normalized:
        return "behavioral_qc"
    return "questionnaire"


def load_required_sessions_by_arm(path: Path, log: ValidationLog) -> pd.DataFrame:
    if not path.exists():
        log.warn("missing_design_files", f"{path} was not found.")
        raise FileNotFoundError(f"Required design file not found: {path}")

    rows: list[dict[str, Any]] = []
    workbook = pd.ExcelFile(path)
    for sheet_name in workbook.sheet_names:
        arm = normalize_design_arm(sheet_name)
        df = pd.read_excel(path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
        df = df.fillna("")
        session_col = next((col for col in df.columns if normalize_column_name(col) in {"required_session", "session"}), None)
        if session_col is None:
            log.warn("missing_design_columns", f"{path.name}/{sheet_name}: no required session column found.")
            continue
        order_col = next((col for col in df.columns if normalize_column_name(col) == "order"), None)
        for row_index, row in df.iterrows():
            session = normalize_design_session(row[session_col])
            if not session:
                continue
            if session in CONFIG.get("excluded_output_sessions", []):
                continue
            if session not in CONFIG.get("session_order", []):
                log.warn("unknown_design_sessions", f"{path.name}/{sheet_name}: {session}")
            order_value = pd.to_numeric(clean_value(row[order_col]), errors="coerce") if order_col else pd.NA
            rows.append(
                {
                    "arm": arm,
                    "session": session,
                    "required": True,
                    "session_order_index": session_sort_key(session),
                    "design_order": int(order_value) if pd.notna(order_value) else row_index + 1,
                    "participant_event_name": clean_value(row.get("Participant event name", "")),
                    "clinician_event_name": clean_value(row.get("Clinician event name", "")),
                }
            )
    out = pd.DataFrame(rows).drop_duplicates(["arm", "session"]).reset_index(drop=True)
    log.info(f"Loaded required sessions design: {path}")
    log.set_counter("required_session_design_rows", len(out))
    return sort_by_session_order(out, ["session", "arm"])


def expand_design_session_label(label: str, required_sessions_for_arm: set[str]) -> list[str]:
    label = clean_value(label)
    if not label:
        return []
    normalized = normalize_design_session(label)
    if normalized in required_sessions_for_arm:
        return [normalized]

    expanded: list[str] = []
    for part in re.split(r"\s*/\s*", label):
        part_norm = normalize_design_session(part)
        if part_norm in required_sessions_for_arm:
            expanded.append(part_norm)

    treatment_match = re.search(r"Treatment Session\s+(.+)$", label, flags=re.I)
    if treatment_match:
        number_text = treatment_match.group(1)
        for number in re.findall(r"\d+", number_text):
            session = f"T{number}"
            if session in required_sessions_for_arm:
                expanded.append(session)

    for token in re.findall(r"\bT(?:1[0-2]|[1-9])\s*IE\b|\bT(?:1[0-2]|[1-9])\b", label, flags=re.I):
        token = re.sub(r"\s+", " ", token.upper()).replace(" IE", " IE")
        if token.endswith(" IE"):
            session = f"IE {token[:-3]}"
        else:
            session = token
        if session in required_sessions_for_arm:
            expanded.append(session)

    return sorted(set(expanded), key=session_sort_key)


def load_required_instruments_by_arm_session(
    path: Path,
    required_sessions_by_arm: pd.DataFrame,
    log: ValidationLog,
) -> pd.DataFrame:
    if not path.exists():
        log.warn("missing_design_files", f"{path} was not found.")
        raise FileNotFoundError(f"Required design file not found: {path}")

    rows: list[dict[str, Any]] = []
    required_lookup = {
        arm: set(group["session"])
        for arm, group in required_sessions_by_arm.groupby("arm", dropna=False)
    }
    workbook = pd.ExcelFile(path)
    for sheet_name in workbook.sheet_names:
        arm = normalize_design_arm(sheet_name)
        required_sessions_for_arm = required_lookup.get(arm, set())
        df = pd.read_excel(path, sheet_name=sheet_name, dtype=str, keep_default_na=False)
        df = df.fillna("")
        if len(df.columns) < 3:
            log.warn("missing_design_columns", f"{path.name}/{sheet_name}: expected at least 3 columns.")
            continue
        session_col, source_col, instrument_col = list(df.columns[:3])
        included_col = next(
            (col for col in df.columns if normalize_column_name(col) == "included_in_summary"),
            None,
        )
        if included_col is None:
            log.warn(
                "missing_design_columns",
                f"{path.name}/{sheet_name}: Included_in_summary column not found; no instruments from this sheet will be expected.",
            )
            continue
        excluded_by_included_flag = 0
        current_sessions = expand_design_session_label(session_col, required_sessions_for_arm)
        current_source = clean_value(source_col).lower()

        first_instrument = normalize_design_instrument(instrument_col)
        if first_instrument:
            for session in current_sessions:
                instrument_type = infer_instrument_type(first_instrument, current_source, log)
                rows.append(
                    {
                        "arm": arm,
                        "session": session,
                        "instrument": first_instrument,
                        "instrument_key": instrument_match_key(first_instrument),
                        "instrument_type": instrument_type,
                        "instrument_source": "participants" if "participant" in current_source else "clinician" if "clinician" in current_source else current_source,
                        "required": True,
                        "included_in_summary": True,
                        "session_order_index": session_sort_key(session),
                        "instrument_order_index": 0,
                    }
                )

        for row_index, row in df.iterrows():
            session_value = clean_value(row[session_col])
            source_value = clean_value(row[source_col])
            instrument = normalize_design_instrument(row[instrument_col])
            if session_value:
                current_sessions = expand_design_session_label(session_value, required_sessions_for_arm)
            if source_value:
                current_source = source_value.lower()
            if not instrument or not current_sessions:
                continue
            if not included_in_summary(row[included_col]):
                excluded_by_included_flag += 1
                continue
            source = "participants" if "participant" in current_source else "clinician" if "clinician" in current_source else current_source
            instrument_type = infer_instrument_type(instrument, source, log)
            for session in current_sessions:
                rows.append(
                    {
                        "arm": arm,
                        "session": session,
                        "instrument": instrument,
                        "instrument_key": instrument_match_key(instrument),
                        "instrument_type": instrument_type,
                        "instrument_source": source,
                        "required": True,
                        "included_in_summary": True,
                        "session_order_index": session_sort_key(session),
                        "instrument_order_index": row_index + 1,
                    }
                )
        if excluded_by_included_flag:
            log.info(
                f"{path.name}/{sheet_name}: excluded {excluded_by_included_flag} instruments with Included_in_summary != 1"
            )
    out = pd.DataFrame(rows)
    if out.empty:
        log.warn("missing_design_columns", f"{path.name}: no required instruments parsed.")
        return out
    out = out[~out["session"].isin(CONFIG.get("excluded_output_sessions", []))].copy()
    out = out.drop_duplicates(["arm", "session", "instrument_source", "instrument_key"]).reset_index(drop=True)
    out = out.sort_values(["arm", "session_order_index", "instrument_order_index", "instrument"], kind="mergesort").reset_index(drop=True)
    log.info(f"Loaded required instruments design: {path}")
    log.set_counter("required_instrument_design_rows", len(out))
    return out


def load_design_tables(input_dir: Path, log: ValidationLog) -> dict[str, pd.DataFrame]:
    design_files = CONFIG.get("design_files", {})
    sessions_path = design_file_path(input_dir, design_files.get("required_sessions_by_arm", "Required_Sessions_for_each_Arm.xlsx"))
    instruments_path = design_file_path(input_dir, design_files.get("instruments_by_session_by_arm", "Instruments_in_each_Session_each_Arm.xlsx"))
    required_sessions = load_required_sessions_by_arm(sessions_path, log)
    required_instruments = load_required_instruments_by_arm_session(instruments_path, required_sessions, log)
    return {
        "required_sessions_by_arm": required_sessions,
        "required_instruments_by_arm_session": required_instruments,
    }


def add_design_event_mappings(event_mapping: dict[str, str], required_sessions_by_arm: pd.DataFrame) -> dict[str, str]:
    event_mapping = dict(event_mapping)
    if required_sessions_by_arm.empty:
        return event_mapping
    for _, row in required_sessions_by_arm.iterrows():
        session = clean_value(row["session"])
        for col in ("participant_event_name", "clinician_event_name"):
            raw = clean_value(row.get(col, ""))
            if not raw or raw.startswith("—"):
                continue
            event_mapping[normalize_column_name(raw)] = session
    return event_mapping


def parse_pdf_text(path: Path, log: ValidationLog) -> str:
    if not path.exists():
        log.warn("missing_codebook", f"{path} was not found.")
        return ""
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception as exc:
        try:
            from PyPDF2 import PdfReader  # type: ignore
        except Exception:
            log.warn("codebook_parse", f"PDF parser unavailable; using CONFIG fallback for {path.name}: {exc}")
            return ""
    try:
        reader = PdfReader(str(path))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception as exc:
        log.warn("codebook_parse", f"Could not parse {path.name}; using CONFIG fallback: {exc}")
        return ""


def parse_event_mapping_from_pdf(text: str) -> dict[str, str]:
    mapping: dict[str, str] = {}
    for line in text.splitlines():
        compact = " ".join(line.split())
        match = re.search(r"\b([a-z0-9]+(?:_[a-z0-9]+)*_arm_\d+)\b\s+(.+)$", compact, flags=re.I)
        if match:
            event_key = normalize_column_name(match.group(1))
            event_label = match.group(2).strip()
            if event_label and len(event_label) <= 80:
                mapping[event_key] = event_label
    return mapping


def get_codebook_event_mapping(input_dir: Path, log: ValidationLog) -> dict[str, str]:
    event_mapping = dict(CONFIG["event_mapping"])
    for source_file, filename in CODEBOOK_FILES.items():
        text = parse_pdf_text(input_dir / filename, log)
        parsed = parse_event_mapping_from_pdf(text)
        if parsed:
            log.info(f"Parsed {len(parsed)} event mappings from {filename}.")
            event_mapping.update({key: event_mapping.get(key, value) for key, value in parsed.items()})
    return {normalize_column_name(key): value for key, value in event_mapping.items()}


def map_event_names(df: pd.DataFrame, event_mapping: dict[str, str], log: ValidationLog) -> tuple[pd.DataFrame, str | None]:
    event_field = choose_most_complete_field(
        df,
        CONFIG["event_field_candidates"],
        "event",
        log,
    )
    df = df.copy()
    if event_field is None:
        df["raw_event_name"] = ""
        df["session"] = CONFIG["baseline_session_name"]
    else:
        df["raw_event_name"] = df[event_field].map(clean_value).map(normalize_column_name)

        source_mappings = {
            source: {
                normalize_column_name(raw_event): label
                for raw_event, label in mapping.items()
            }
            for source, mapping in CONFIG.get("event_mapping_by_source", {}).items()
        }

        def map_one_event(row: pd.Series) -> str:
            raw_event = row["raw_event_name"]
            source = row["source_file"]
            if source in source_mappings and raw_event in source_mappings[source]:
                return source_mappings[source][raw_event]
            if raw_event in event_mapping:
                return event_mapping[raw_event]
            if raw_event:
                log.warn("unknown_redcap_event_names", raw_event)
                return f"UNMAPPED_EVENT_{raw_event}"
            return CONFIG["baseline_session_name"]

        df["session"] = df.apply(map_one_event, axis=1)
    df["session"] = df["session"].map(canonical_session_name)
    log.set_counter("detected_event_field", event_field or "not_found")
    log.set_counter("number_of_sessions_detected", df["session"].nunique(dropna=True))
    return df, event_field


def readable_name(raw_name: str) -> str:
    parts = [part for part in normalize_column_name(raw_name).split("_") if part]
    return "".join(part.capitalize() for part in parts) or raw_name


def infer_completion_fields(columns: list[str]) -> dict[str, str]:
    configured = {
        normalize_column_name(form): normalize_column_name(field)
        for form, field in CONFIG["completion_status_fields_by_form"].items()
    }
    inferred = {
        col.removesuffix("_complete"): col
        for col in columns
        if col.endswith("_complete")
    }
    inferred.update(configured)
    return {form: field for form, field in inferred.items() if field in columns}


def first_present(row: pd.Series, candidates: list[str]) -> str:
    for candidate in candidates:
        if candidate in row.index:
            value = clean_value(row[candidate])
            if value:
                return value
    return ""


def infer_date_field(questionnaire: str, columns: list[str]) -> str | None:
    configured = CONFIG["date_fields_by_questionnaire"].get(questionnaire)
    if configured and normalize_column_name(configured) in columns:
        return normalize_column_name(configured)
    raw = normalize_column_name(questionnaire)
    candidates = [
        f"{raw}_date",
        f"{raw}_dt",
        f"{raw}_completed_date",
        f"{raw}_completion_date",
    ]
    for candidate in candidates:
        if candidate in columns:
            return candidate
    matches = [col for col in columns if col.startswith(raw) and "date" in col]
    return matches[0] if matches else None


def determine_questionnaire_status(raw_completion_value: str, date_value: str, expected: bool) -> tuple[str, str]:
    value = clean_value(raw_completion_value).lower()
    if not expected:
        if value == "" and date_value == "":
            return "not_expected", "not_expected"
    if value in {"2", "complete", "completed", "yes", "y", "done"}:
        return "complete", ""
    if value in {"0", "incomplete", "not_complete", "not completed", "no", "n"}:
        return "incomplete", "partially_completed"
    if value in {"1", "unverified", "partial", "partially complete"}:
        return "unverified", "no_qc_confirmation"
    if value == "":
        if date_value:
            return "unverified", "no_qc_confirmation"
        if expected:
            return "missing", "not_started"
        return "not_expected", "not_expected"
    return "review_required", "review_required"


def expected_questionnaires_for_session(session: str) -> set[str]:
    session = canonical_session_name(session)
    expected_by_session = canonical_session_mapping(CONFIG["expected_questionnaires_by_session"])
    must_have_by_session = canonical_session_mapping(CONFIG["must_have_questionnaires_by_session"])
    expected = set(expected_by_session.get(session, []))
    expected.update(must_have_by_session.get(session, []))
    return {readable_name(item) for item in expected}


def is_questionnaire_expected(session: str, questionnaire: str) -> bool:
    expected = expected_questionnaires_for_session(session)
    if expected:
        return questionnaire in expected
    if CONFIG["treat_observed_as_expected_when_no_design_config"]:
        return True
    return False


def build_questionnaire_long(df: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    completion_fields = infer_completion_fields(list(df.columns))
    if not completion_fields:
        log.warn("missing_expected_fields", "No REDCap completion status fields ending in _complete were found.")
    rows: list[dict[str, Any]] = []
    unknown_status_values: set[str] = set()
    for _, row in df.iterrows():
        for form_name, completion_field in completion_fields.items():
            session = canonical_session_name(clean_value(row["session"]))
            questionnaire = readable_name(form_name)
            date_field = infer_date_field(questionnaire, list(df.columns)) or infer_date_field(form_name, list(df.columns))
            date_value = clean_value(row[date_field]) if date_field else ""
            raw_completion = clean_value(row[completion_field])
            expected = is_questionnaire_expected(session, questionnaire)
            status, reason = determine_questionnaire_status(raw_completion, date_value, expected)
            if status == "review_required" and raw_completion:
                unknown_status_values.add(f"{completion_field}={raw_completion}")
            rows.append(
                {
                    "subject_id": row["subject_id"],
                    "source_file": row["source_file"],
                    "instrument_source": row["instrument_source"],
                    "session": session,
                    "questionnaire": questionnaire,
                    "date": date_value,
                    "status": status,
                    "expected": bool(expected),
                    "missingness_reason": reason,
                    "raw_form_name": form_name,
                    "raw_completion_value": raw_completion,
                }
            )
    for item in sorted(unknown_status_values):
        log.warn("unknown_questionnaire_status_codes", item)
    qlong = pd.DataFrame(rows)
    log.set_counter("number_of_questionnaires_detected", qlong["questionnaire"].nunique() if not qlong.empty else 0)
    return add_questionnaire_output_names(resolve_duplicate_questionnaires(qlong, log))


def add_questionnaire_output_names(qlong: pd.DataFrame) -> pd.DataFrame:
    """Add source-aware questionnaire names for wide subject-level output."""
    if qlong.empty:
        return qlong
    qlong = qlong.copy()
    source_counts = qlong.groupby(["session", "questionnaire"])["source_file"].transform("nunique")
    qlong["questionnaire_output_name"] = qlong["questionnaire"]
    needs_prefix = source_counts > 1
    qlong.loc[needs_prefix, "questionnaire_output_name"] = (
        qlong.loc[needs_prefix, "instrument_source"].map(clean_value)
        + "_"
        + qlong.loc[needs_prefix, "questionnaire"].map(clean_value)
    )
    return qlong


def add_qc_output_names(qc_long: pd.DataFrame, item_column: str, output_column: str) -> pd.DataFrame:
    """Add source-aware QC item names for wide subject-level output."""
    if qc_long.empty:
        return qc_long
    qc_long = qc_long.copy()
    source_counts = qc_long.groupby(["session", item_column])["source_file"].transform("nunique")
    qc_long[output_column] = qc_long[item_column]
    needs_prefix = source_counts > 1
    qc_long.loc[needs_prefix, output_column] = (
        qc_long.loc[needs_prefix, "instrument_source"].map(clean_value)
        + "_"
        + qc_long.loc[needs_prefix, item_column].map(clean_value)
    )
    return qc_long


def status_priority(status: str) -> int:
    return {
        "complete": 5,
        "incomplete": 4,
        "unverified": 3,
        "review_required": 2,
        "missing": 1,
        "not_expected": 0,
    }.get(status, -1)


def resolve_duplicate_questionnaires(qlong: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    if qlong.empty:
        return qlong
    keys = ["subject_id", "source_file", "session", "questionnaire"]
    duplicates = qlong[qlong.duplicated(keys, keep=False)].copy()
    if not duplicates.empty:
        log.warn(
            "duplicate_subject_session_questionnaire_records",
            f"{len(duplicates)} duplicate rows found before resolving by completeness.",
        )
    resolved_rows: list[dict[str, Any]] = []
    for _, group in qlong.groupby(keys, dropna=False):
        group = group.copy()
        if len(group) == 1:
            resolved_rows.append(group.iloc[0].to_dict())
            continue
        group["_priority"] = group["status"].map(status_priority)
        best = group[group["_priority"] == group["_priority"].max()].drop(columns=["_priority"])
        comparable_cols = ["date", "status", "raw_completion_value"]
        if len(best[comparable_cols].drop_duplicates()) > 1:
            row = best.iloc[0].to_dict()
            row["status"] = "review_required"
            row["missingness_reason"] = "review_required"
            log.warn(
                "conflicting_duplicate_records",
                "Conflict for "
                + ", ".join(f"{key}={row[key]}" for key in keys),
            )
            resolved_rows.append(row)
        else:
            resolved_rows.append(best.iloc[0].to_dict())
    return pd.DataFrame(resolved_rows)


def normalize_task_completion_status(raw_value: str, expected: bool) -> tuple[str, str]:
    value = clean_value(raw_value).lower()
    complete_values = {str(v).strip().lower() for v in CONFIG["task_completion_complete_values"]}
    incomplete_values = {str(v).strip().lower() for v in CONFIG["task_completion_incomplete_values"]}
    unverified_values = {str(v).strip().lower() for v in CONFIG["task_completion_unverified_values"]}

    if not expected and value == "":
        return "not_expected", "not_expected"
    if value in complete_values:
        return "complete", ""
    if value in incomplete_values:
        return "incomplete", "partially_completed"
    if value in unverified_values:
        return "unverified", "no_completion_confirmation"
    if value == "":
        return ("missing", "not_started") if expected else ("not_expected", "not_expected")
    return "review_required", "review_required"


def build_task_completion_long(df: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    task_fields_by_session = canonical_session_mapping(CONFIG.get("task_completion_fields", {}))
    unknown_values: set[str] = set()

    for _, row in df.iterrows():
        session = canonical_session_name(clean_value(row["session"]))
        fields = task_fields_by_session.get(session, {})
        for task_name, raw_field in fields.items():
            field_name = normalize_column_name(raw_field)
            expected = True
            if field_name not in df.columns:
                log.warn("missing_expected_fields", f"task_completion_fields: {field_name} not found in CSV.")
                raw_value = ""
            else:
                raw_value = clean_value(row[field_name])
            status, reason = normalize_task_completion_status(raw_value, expected)
            if status == "review_required":
                unknown_values.add(f"{field_name}={raw_value}")
            rows.append(
                {
                    "subject_id": row["subject_id"],
                    "source_file": row["source_file"],
                    "instrument_source": row["instrument_source"],
                    "session": session,
                    "task": task_name,
                    "status": status,
                    "task_completed": status == "complete",
                    "expected": expected,
                    "missingness_reason": reason,
                    "raw_task_completion_value": raw_value,
                }
            )

    for item in sorted(unknown_values):
        log.warn("unknown_task_completion_values", item)
    task_long = pd.DataFrame(rows)
    return add_qc_output_names(task_long, "task", "task_output_name")


def normalize_qc_status(raw_value: str, pass_values: list[str], fail_values: list[str], expected: bool) -> tuple[str, bool | None, str]:
    value = clean_value(raw_value).lower()
    pass_set = {str(item).strip().lower() for item in pass_values}
    fail_set = {str(item).strip().lower() for item in fail_values}
    if not expected and value == "":
        return "not_expected", None, "not_expected"
    if value in pass_set:
        return "pass", True, ""
    if value in fail_set:
        return "fail", False, "failed_qc"
    if value == "":
        return "missing", False, "no_qc_confirmation" if expected else "not_expected"
    return "review_required", False, "review_required"


def build_qc_long(
    df: pd.DataFrame,
    qc_config_key: str,
    pass_values_key: str,
    fail_values_key: str,
    item_column: str,
    log: ValidationLog,
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    qc_fields_by_session = canonical_session_mapping(CONFIG[qc_config_key])
    unknown_values: set[str] = set()
    for _, row in df.iterrows():
        session = canonical_session_name(clean_value(row["session"]))
        fields = qc_fields_by_session.get(session, {})
        for item_name, raw_field in fields.items():
            field_name = normalize_column_name(raw_field)
            expected = True
            if field_name not in df.columns:
                log.warn("missing_expected_fields", f"{qc_config_key}: {field_name} not found in CSV.")
                raw_value = ""
            else:
                raw_value = clean_value(row[field_name])
            status, qc_pass, reason = normalize_qc_status(
                raw_value,
                CONFIG[pass_values_key],
                CONFIG[fail_values_key],
                expected,
            )
            if status == "review_required":
                unknown_values.add(f"{field_name}={raw_value}")
            rows.append(
                {
                    "subject_id": row["subject_id"],
                    "source_file": row["source_file"],
                    "instrument_source": row["instrument_source"],
                    "session": session,
                    item_column: item_name,
                    "qc_status": status,
                    "qc_pass": qc_pass,
                    "expected": expected,
                    "missingness_reason": reason,
                    "raw_qc_value": raw_value,
                }
            )
    for item in sorted(unknown_values):
        log.warn("unknown_qc_values", item)
    return pd.DataFrame(rows)


def map_mri_qc_session(raw_session: str, log: ValidationLog) -> str:
    raw = clean_value(raw_session)
    mapping = CONFIG.get("mri_qc_session_mapping", {})

    if raw in mapping:
        return canonical_session_name(mapping[raw])

    normalized_raw = raw.strip()
    if normalized_raw in mapping:
        return canonical_session_name(mapping[normalized_raw])

    lower_mapping = {str(key).lower(): value for key, value in mapping.items()}
    if raw.lower() in lower_mapping:
        return canonical_session_name(lower_mapping[raw.lower()])

    if raw:
        log.warn("unknown_mri_qc_sessions", raw)
        return f"UNMAPPED_MRI_SESSION_{raw}"

    log.warn("unknown_mri_qc_sessions", "missing sesID")
    return "UNMAPPED_MRI_SESSION_MISSING"


def normalize_poor_quality_flag(
    raw_value: str,
    log: ValidationLog,
    context: str,
) -> tuple[str, bool, str]:
    value = clean_value(raw_value).lower()
    true_values = {str(item).strip().lower() for item in CONFIG.get("mri_qc_true_values", [])}
    false_values = {str(item).strip().lower() for item in CONFIG.get("mri_qc_false_values", [])}

    if value in true_values:
        return "fail", False, "failed_qc"
    if value in false_values:
        return "pass", True, ""
    if value == "":
        log.warn("unknown_mri_qc_poor_quality_values", f"{context}: missing Poor_Quality")
        return "review_required", False, "review_required"

    log.warn("unknown_mri_qc_poor_quality_values", f"{context}: Poor_Quality={raw_value}")
    return "review_required", False, "review_required"


def clean_mri_label(value: str) -> str:
    label = clean_value(value)
    label = re.sub(r"\.html$", "", label, flags=re.I)
    label = re.sub(r"[^0-9a-zA-Z]+", "_", label)
    label = re.sub(r"_+", "_", label).strip("_")
    return label or "unknown"


def qc_status_priority(status: str) -> int:
    return {
        "fail": 5,
        "review_required": 4,
        "missing": 3,
        "unverified": 2,
        "pass": 1,
        "not_expected": 0,
    }.get(status, -1)


def resolve_duplicate_mri_qc_records(mri_long: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    if mri_long.empty:
        return mri_long

    keys = ["subject_id", "session", "scan_or_run"]
    duplicates = mri_long[mri_long.duplicated(keys, keep=False)].copy()
    if not duplicates.empty:
        log.warn(
            "duplicate_mri_qc_records",
            f"{len(duplicates)} duplicate MRI QC rows found before resolving.",
        )

    resolved_rows: list[dict[str, Any]] = []
    for _, group in mri_long.groupby(keys, dropna=False):
        group = group.copy()
        if len(group) == 1:
            resolved_rows.append(group.iloc[0].to_dict())
            continue

        group["_priority"] = group["qc_status"].map(qc_status_priority)
        best = group[group["_priority"] == group["_priority"].max()].drop(columns=["_priority"])
        comparable_cols = ["qc_status", "qc_pass", "raw_qc_value"]
        if len(best[comparable_cols].drop_duplicates()) > 1:
            row = best.iloc[0].to_dict()
            row["qc_status"] = "review_required"
            row["qc_pass"] = False
            row["missingness_reason"] = "review_required"
            log.warn(
                "conflicting_duplicate_mri_qc_records",
                "Conflict for " + ", ".join(f"{key}={row[key]}" for key in keys),
            )
            resolved_rows.append(row)
        else:
            resolved_rows.append(best.iloc[0].to_dict())

    return pd.DataFrame(resolved_rows)


def mri_qc_file_path(input_dir: Path, filename: str) -> Path:
    direct_path = input_dir / filename
    if direct_path.exists():
        return direct_path
    data_path = input_dir / "data" / filename
    if data_path.exists():
        return data_path
    return direct_path


def build_mri_qc_long(input_dir: Path, log: ValidationLog) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    files = CONFIG.get("mri_qc_files", {})
    poor_quality_field = CONFIG.get("mri_qc_poor_quality_field", "Poor_Quality")

    anat = read_mri_qc_csv(
        mri_qc_file_path(input_dir, files.get("anat", "QC_anat.csv")),
        source_file="mriqc_anat",
        required_columns=["subID", "sesID", "modality", poor_quality_field],
        log=log,
    )
    func = read_mri_qc_csv(
        mri_qc_file_path(input_dir, files.get("func", "QC_func.csv")),
        source_file="mriqc_func",
        required_columns=["subID", "sesID", "modality", "taskID", "runID", poor_quality_field],
        log=log,
    )

    for _, row in anat.iterrows():
        raw_subid = clean_value(row["subID"])
        subject_id = standardize_subject_id(raw_subid, "QC_anat.csv", log)
        raw_session = clean_value(row["sesID"])
        session = map_mri_qc_session(raw_session, log)
        scan_or_run = f"anat_{clean_mri_label(row['modality'])}"
        context = f"QC_anat.csv raw_subID={raw_subid}, subject_id={subject_id}, sesID={raw_session}, scan_or_run={scan_or_run}"
        qc_status, qc_pass, reason = normalize_poor_quality_flag(row[poor_quality_field], log, context)
        rows.append(
            {
                "raw_subID": raw_subid,
                "subject_id": subject_id,
                "source_file": "mriqc_anat",
                "instrument_source": "mriqc",
                "session": session,
                "scan_or_run": scan_or_run,
                "qc_status": qc_status,
                "qc_pass": qc_pass,
                "expected": True,
                "missingness_reason": reason,
                "raw_qc_value": clean_value(row[poor_quality_field]),
                "raw_mri_session": raw_session,
                "modality": clean_value(row["modality"]),
                "taskID": "",
                "runID": "",
            }
        )

    for _, row in func.iterrows():
        raw_subid = clean_value(row["subID"])
        subject_id = standardize_subject_id(raw_subid, "QC_func.csv", log)
        raw_session = clean_value(row["sesID"])
        session = map_mri_qc_session(raw_session, log)
        task = clean_mri_label(row["taskID"])
        run = clean_mri_label(row["runID"])
        scan_or_run = f"func_{task}_run{run}"
        context = f"QC_func.csv raw_subID={raw_subid}, subject_id={subject_id}, sesID={raw_session}, scan_or_run={scan_or_run}"
        qc_status, qc_pass, reason = normalize_poor_quality_flag(row[poor_quality_field], log, context)
        rows.append(
            {
                "raw_subID": raw_subid,
                "subject_id": subject_id,
                "source_file": "mriqc_func",
                "instrument_source": "mriqc",
                "session": session,
                "scan_or_run": scan_or_run,
                "qc_status": qc_status,
                "qc_pass": qc_pass,
                "expected": True,
                "missingness_reason": reason,
                "raw_qc_value": clean_value(row[poor_quality_field]),
                "raw_mri_session": raw_session,
                "modality": clean_value(row["modality"]),
                "taskID": clean_value(row["taskID"]),
                "runID": clean_value(row["runID"]),
            }
        )

    mri_long = pd.DataFrame(rows)
    if mri_long.empty:
        log.warn("mri_qc_empty", "No MRI QC rows were loaded from QC_anat.csv or QC_func.csv.")
        return mri_long

    if (mri_long["subject_id"].map(lambda value: clean_value(value).lower().startswith("sub-"))).any():
        log.warn("subject_id_standardization", "MRI QC standardized subject_id still contains sub- prefix.")
    examples = mri_long[["raw_subID", "subject_id"]].drop_duplicates().head(10).to_dict("records")
    log.info(f"mri_subID_standardization_examples: {examples}")
    mri_long = filter_excluded_output_sessions(mri_long)
    mri_long = resolve_duplicate_mri_qc_records(mri_long, log)
    mri_long["arm"] = mri_long["subject_id"].map(infer_arm_from_subject_id)
    unknown_mri_subjects = sorted(set(mri_long.loc[mri_long["arm"] == "unknown", "subject_id"].map(clean_value)) - {""})
    for subject_id in unknown_mri_subjects:
        log.warn("unknown_subject_arm", f"MRI QC subject_id={subject_id}")
    mri_long = add_qc_output_names(mri_long, "scan_or_run", "scan_or_run_output_name")
    log.set_counter("number_of_mri_qc_rows", len(mri_long))
    log.set_counter("number_of_mri_qc_subjects", mri_long["subject_id"].nunique())
    log.set_counter("mri_qc_pass_rows", int((mri_long["qc_status"] == "pass").sum()))
    log.set_counter("mri_qc_fail_rows", int((mri_long["qc_status"] == "fail").sum()))
    log.set_counter("mri_qc_review_required_rows", int((mri_long["qc_status"] == "review_required").sum()))
    return mri_long


def add_arm_to_long_table(df: pd.DataFrame, subject_arm_map: dict[str, str]) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    if "arm" not in df.columns:
        df["arm"] = df["subject_id"].map(subject_arm_map).fillna("unknown")
    return df


def required_session_set(required_sessions_by_arm: pd.DataFrame, arm: str) -> set[str]:
    if required_sessions_by_arm.empty:
        return set()
    return set(required_sessions_by_arm.loc[required_sessions_by_arm["arm"] == arm, "session"])


def filter_long_to_required_sessions(
    df: pd.DataFrame,
    required_sessions_by_arm: pd.DataFrame,
    log: ValidationLog,
    table_name: str,
) -> pd.DataFrame:
    if df.empty:
        return df
    df = df.copy()
    keep_mask = []
    unexpected = 0
    for _, row in df.iterrows():
        arm = clean_value(row.get("arm", "unknown"))
        session = canonical_session_name(clean_value(row.get("session", "")))
        required = session in required_session_set(required_sessions_by_arm, arm)
        keep_mask.append(required)
        if not required and arm != "unknown":
            unexpected += 1
    if unexpected:
        log.warn(
            "unexpected_observed_records_not_required_by_design",
            f"{table_name}: {unexpected} observed rows excluded because their arm/session is not required.",
        )
    return df[pd.Series(keep_mask, index=df.index)].copy().reset_index(drop=True)


def apply_required_questionnaire_design(
    qlong: pd.DataFrame,
    subject_table: pd.DataFrame,
    required_instruments: pd.DataFrame,
    log: ValidationLog,
) -> pd.DataFrame:
    questionnaire_design = required_instruments[
        required_instruments["instrument_type"].isin(["questionnaire", "session_status"])
    ].copy() if not required_instruments.empty else pd.DataFrame()
    if questionnaire_design.empty:
        log.warn("missing_design_columns", "No questionnaire/session-status instruments were parsed from the design file.")
        return qlong

    observed = qlong.copy()
    if observed.empty:
        observed = pd.DataFrame(
            columns=[
                "subject_id", "arm", "source_file", "instrument_source", "session", "questionnaire",
                "date", "status", "expected", "missingness_reason", "raw_form_name", "raw_completion_value",
            ]
        )
    observed["instrument_key"] = observed["questionnaire"].map(instrument_match_key)
    observed["source_key"] = observed["source_file"].map(clean_value)

    used_observed_indexes: set[int] = set()
    rows: list[dict[str, Any]] = []
    subjects = subject_table[subject_table["arm"].isin(CONFIG.get("arms", []))]
    for _, subject in subjects.iterrows():
        subject_id = subject["subject_id"]
        arm = subject["arm"]
        arm_design = questionnaire_design[questionnaire_design["arm"] == arm]
        for _, design_row in arm_design.iterrows():
            session = design_row["session"]
            instrument = design_row["instrument"]
            source = clean_value(design_row["instrument_source"])
            source_candidates = [source]
            if source == "participants":
                source_candidates.append("participant")
            match = observed[
                (observed["subject_id"] == subject_id)
                & (observed["arm"] == arm)
                & (observed["session"] == session)
                & (observed["instrument_key"] == design_row["instrument_key"])
                & (observed["source_key"].isin(source_candidates))
            ]
            if match.empty:
                rows.append(
                    {
                        "subject_id": subject_id,
                        "arm": arm,
                        "source_file": source,
                        "instrument_source": source,
                        "session": session,
                        "questionnaire": instrument,
                        "date": "",
                        "status": "missing",
                        "expected": True,
                        "missingness_reason": "not_started",
                        "raw_form_name": "",
                        "raw_completion_value": "",
                        "questionnaire_output_name": instrument,
                    }
                )
                log.warn("missing_required_records", f"{subject_id}/{arm}/{session}/{source}/{instrument}")
                continue
            best = match.iloc[0].to_dict()
            used_observed_indexes.add(int(match.index[0]))
            best.update(
                {
                    "questionnaire": instrument,
                    "expected": True,
                    "questionnaire_output_name": instrument,
                    "instrument_source": source,
                    "source_file": source,
                }
            )
            rows.append(best)

    unused = observed[
        (~observed.index.isin(used_observed_indexes))
        & (observed["arm"].isin(CONFIG.get("arms", [])))
        & (observed["status"] != "not_expected")
    ]
    if not unused.empty:
        log.warn(
            "unexpected_observed_records_not_required_by_design",
            f"questionnaire_long: {len(unused)} observed rows did not match required arm/session/instrument design.",
        )

    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return add_questionnaire_output_names(out.drop(columns=["instrument_key", "source_key"], errors="ignore"))


def apply_observation_based_mri_expectedness(
    mri_long: pd.DataFrame,
    required_instruments: pd.DataFrame,
    log: ValidationLog,
) -> pd.DataFrame:
    if mri_long.empty:
        return mri_long
    mri_design = required_instruments[
        required_instruments["instrument_type"].eq("mri_qc")
    ] if not required_instruments.empty else pd.DataFrame()
    if mri_design.empty:
        log.warn(
            "mri_expectedness_observation_based",
            "No explicit MRI QC expectedness found in instrument design; observed external MRI QC rows define MRI denominators.",
        )
        return mri_long
    return mri_long


def arm_filter(df: pd.DataFrame, arm: str) -> pd.DataFrame:
    if df.empty:
        return df
    if "arm" not in df.columns:
        return df
    return df[df["arm"] == arm].copy().reset_index(drop=True)


def add_arm_column_to_group_tables(group_tables: dict[str, pd.DataFrame], arm: str) -> dict[str, pd.DataFrame]:
    updated: dict[str, pd.DataFrame] = {}
    for title, table in group_tables.items():
        table = table.copy()
        if "arm" not in table.columns:
            table.insert(0, "arm", arm)
        updated[title] = table
    return updated


def build_behavioral_qc_long(df: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    return add_qc_output_names(
        build_qc_long(df, "behavioral_qc_fields", "behavioral_qc_pass_values", "behavioral_qc_fail_values", "task", log),
        "task",
        "task_output_name",
    )


def parse_date(value: str, context: str, log: ValidationLog) -> pd.Timestamp | pd.NaT:
    raw = clean_value(value)
    if not raw:
        return pd.NaT
    parsed = pd.to_datetime(raw, errors="coerce")
    if pd.isna(parsed):
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%-m/%-d/%Y", "%Y/%m/%d"):
            try:
                return pd.to_datetime(raw, format=fmt)
            except Exception:
                continue
        log.warn("date_parsing_failures", f"{context}: {raw}")
        return pd.NaT
    return parsed


def calculate_interval_weeks(session_date: pd.Timestamp, baseline_date: pd.Timestamp) -> float | None:
    if pd.isna(session_date) or pd.isna(baseline_date):
        return None
    return round((session_date - baseline_date).days / 7, 3)


def classify_interval_validity(session: str, interval_weeks: float | None, baseline_date: Any, session_date: Any) -> str:
    session = canonical_session_name(session)
    timing_windows = canonical_session_mapping(CONFIG["session_timing_windows"])
    if session not in timing_windows:
        return "not_expected"
    if pd.isna(baseline_date):
        return "missing_baseline_date"
    if pd.isna(session_date):
        return "missing_session_date"
    if interval_weeks is None:
        return "review_required"
    window = timing_windows[session]
    if interval_weeks < window["allowed_min_weeks"]:
        return "early"
    if interval_weeks > window["allowed_max_weeks"]:
        return "late"
    return "valid"


def build_session_long(
    df: pd.DataFrame,
    qlong: pd.DataFrame,
    required_sessions_by_arm: pd.DataFrame,
    log: ValidationLog,
) -> pd.DataFrame:
    session_date_fields = {
        canonical_session_name(session): normalize_column_name(field)
        for session, field in CONFIG["session_date_fields"].items()
    }
    dates: dict[tuple[str, str], str] = {}
    for _, row in df.iterrows():
        session = canonical_session_name(clean_value(row["session"]))
        field = session_date_fields.get(session)
        date_value = clean_value(row[field]) if field and field in df.columns else ""
        if not date_value and not qlong.empty:
            candidates = qlong[
                (qlong["subject_id"] == row["subject_id"])
                & (qlong["session"] == session)
                & (qlong["date"].map(clean_value) != "")
            ]
            if not candidates.empty:
                date_value = clean_value(candidates.iloc[0]["date"])
        if date_value:
            dates[(row["subject_id"], session)] = date_value

    baseline_session = CONFIG["baseline_session_name"]
    rows: list[dict[str, Any]] = []
    subject_ids = sorted(set(df["subject_id"]) - {""})
    subject_arm_lookup = (
        df[["subject_id", "arm"]]
        .drop_duplicates("subject_id")
        .set_index("subject_id")["arm"]
        .to_dict()
        if "arm" in df.columns else {}
    )
    exclude_from_interval = {
        canonical_session_name(session)
        for session in CONFIG.get("exclude_from_interval_sessions", [])
    }
    configured_interval_sessions = {
        canonical_session_name(session)
        for session in CONFIG.get("interval_sessions", [])
    }
    for subject_id in subject_ids:
        subject_arm = subject_arm_lookup.get(subject_id, infer_arm_from_subject_id(subject_id))
        required_sessions_for_subject = sorted(
            [
                canonical_session_name(session)
                for session in required_session_set(required_sessions_by_arm, subject_arm)
                if canonical_session_name(session) not in exclude_from_interval
            ],
            key=session_sort_key,
        )
        if configured_interval_sessions:
            required_sessions_for_subject = [
                session for session in required_sessions_for_subject
                if session in configured_interval_sessions
            ]
        baseline_raw = dates.get((subject_id, baseline_session), "")
        baseline_date = parse_date(baseline_raw, f"{subject_id}/{baseline_session}", log)
        for session in required_sessions_for_subject:
            session_raw = dates.get((subject_id, session), "")
            session_date = parse_date(session_raw, f"{subject_id}/{session}", log)
            interval = calculate_interval_weeks(session_date, baseline_date)
            validity = classify_interval_validity(session, interval, baseline_date, session_date)
            window = canonical_session_mapping(CONFIG["session_timing_windows"]).get(session, {})
            rows.append(
                {
                    "subject_id": subject_id,
                    "arm": subject_arm,
                    "session": session,
                    "session_date": session_raw,
                    "session_completed": bool(session_raw),
                    "intervalFromBaseline_weeks": interval,
                    "interval_valid": validity,
                    "expected_target_weeks": window.get("target_weeks", ""),
                    "allowed_min_weeks": window.get("allowed_min_weeks", ""),
                    "allowed_max_weeks": window.get("allowed_max_weeks", ""),
                    "missingness_reason": validity if validity != "valid" else "",
                }
            )
    session_long = pd.DataFrame(rows)
    missing_baseline = sum(
        1 for subject_id in subject_ids
        if clean_value(dates.get((subject_id, baseline_session), "")) == ""
    )
    log.set_counter("number_of_missing_baseline_dates", missing_baseline)
    return session_long


def value_in_config(row: pd.Series, fields: list[str], values: list[str]) -> bool:
    configured_fields = [normalize_column_name(field) for field in fields]
    configured_values = {str(value).strip().lower() for value in values}
    for field in configured_fields:
        if field in row.index and clean_value(row[field]).lower() in configured_values:
            return True
    return False


def summarize_asap_for_subject(df_subject: pd.DataFrame) -> int:
    fields = [normalize_column_name(field) for field in CONFIG["asap_fields"]]
    count = 0
    for _, row in df_subject.iterrows():
        for field in fields:
            if field not in row.index:
                continue
            value = clean_value(row[field])
            if not value:
                continue
            numeric = pd.to_numeric(value, errors="coerce")
            if pd.notna(numeric):
                count += int(numeric)
            elif value.lower() in {"yes", "y", "true", "1", "asap"}:
                count += 1
    return count


def safe_column_fragment(value: str) -> str:
    fragment = re.sub(r"[^0-9a-zA-Z_]+", "", str(value))
    return fragment or "Unknown"


def pivot_long_tables_to_subject_wise(
    df: pd.DataFrame,
    qlong: pd.DataFrame,
    task_long: pd.DataFrame,
    mri_long: pd.DataFrame,
    behavioral_long: pd.DataFrame,
    session_long: pd.DataFrame,
    required_sessions_by_arm: pd.DataFrame,
) -> pd.DataFrame:
    subject_ids = set(df["subject_id"].dropna().map(clean_value)) - {""}
    if not mri_long.empty:
        subject_ids.update(set(mri_long["subject_id"].dropna().map(clean_value)) - {""})
    subject_ids = sorted(subject_ids)
    rows: list[dict[str, Any]] = []
    for subject_id in subject_ids:
        df_subject = df[df["subject_id"] == subject_id]
        dropout = any(
            value_in_config(row, CONFIG["dropout_fields"], CONFIG["dropout_values"])
            for _, row in df_subject.iterrows()
        )
        subject_arm = first_nonblank(df_subject.get("arm", pd.Series(dtype=str)))
        if not subject_arm:
            subject_arm = infer_arm_from_subject_id(subject_id)
        out: dict[str, Any] = {
            "subject_id": subject_id,
            "arm": subject_arm,
            "dropout_status": "withdrawn_or_dropout" if dropout else "active_or_unknown",
        }
        for _, row in qlong[qlong["subject_id"] == subject_id].iterrows() if not qlong.empty else []:
            questionnaire_name = row.get("questionnaire_output_name", row["questionnaire"])
            prefix = f"ses-{safe_column_fragment(row['session'])}_qn-{safe_column_fragment(questionnaire_name)}"
            out[f"{prefix}_date"] = row["date"]
            out[f"{prefix}_status"] = row["status"]
        for _, row in task_long[task_long["subject_id"] == subject_id].iterrows() if not task_long.empty else []:
            task_name = row.get("task_output_name", row["task"])
            prefix = f"ses-{safe_column_fragment(row['session'])}_task-{safe_column_fragment(task_name)}"
            out[f"{prefix}_status"] = row["status"]
            out[f"{prefix}_completed"] = row["task_completed"]
        for _, row in mri_long[mri_long["subject_id"] == subject_id].iterrows() if not mri_long.empty else []:
            scan_name = row.get("scan_or_run_output_name", row["scan_or_run"])
            prefix = f"ses-{safe_column_fragment(row['session'])}_mri-{safe_column_fragment(scan_name)}"
            out[f"{prefix}_qc_status"] = row["qc_status"]
            out[f"{prefix}_qc_pass"] = row["qc_pass"]
        for _, row in behavioral_long[behavioral_long["subject_id"] == subject_id].iterrows() if not behavioral_long.empty else []:
            task_name = row.get("task_output_name", row["task"])
            prefix = f"ses-{safe_column_fragment(row['session'])}_beh-{safe_column_fragment(task_name)}"
            out[f"{prefix}_qc_status"] = row["qc_status"]
            out[f"{prefix}_qc_pass"] = row["qc_pass"]
        for _, row in session_long[session_long["subject_id"] == subject_id].iterrows() if not session_long.empty else []:
            prefix = f"ses-{safe_column_fragment(row['session'])}"
            out[f"{prefix}_session_completed"] = row["session_completed"]
            out[f"{prefix}_intervalFromBaseline_weeks"] = row["intervalFromBaseline_weeks"]
            out[f"{prefix}_interval_valid"] = row["interval_valid"]
        asap_count = summarize_asap_for_subject(df_subject)
        out["total_ASAP_count"] = asap_count
        out["has_ASAP"] = asap_count > 0
        rows.append(out)
    subject_wise = pd.DataFrame(rows)
    return add_subject_level_summary_columns(
        subject_wise,
        qlong,
        task_long,
        mri_long,
        behavioral_long,
        session_long,
        required_sessions_by_arm,
    )


def all_expected_complete(records: pd.DataFrame, complete_status: str, empty_value: bool = True) -> bool:
    if records.empty:
        return empty_value
    expected = records[records["expected"].astype(bool)]
    if expected.empty:
        return empty_value
    return bool((expected["status"] == complete_status).all())


def all_expected_qc_pass(records: pd.DataFrame, empty_value: bool = True) -> bool:
    if records.empty:
        return empty_value
    expected = records[records["expected"].astype(bool)]
    if expected.empty:
        return empty_value
    return bool((expected["qc_status"] == "pass").all())


def all_expected_task_complete(task_records: pd.DataFrame, empty_value: bool = True) -> bool:
    if task_records.empty:
        return empty_value
    expected = task_records[task_records["expected"].astype(bool)]
    if expected.empty:
        return empty_value
    return bool((expected["status"] == "complete").all())


def compute_complete_all_experiment_sessions(
    subject_id: str,
    arm: str,
    required_sessions_by_arm: pd.DataFrame,
    session_completion_lookup: dict[tuple[str, str], bool],
    excluded_sessions: set[str] | None = None,
) -> tuple[bool, str]:
    excluded = {canonical_session_name(session) for session in (excluded_sessions or {"Screening"})}
    required_sessions = sorted(
        [
            canonical_session_name(session)
            for session in required_session_set(required_sessions_by_arm, arm)
            if canonical_session_name(session) not in excluded
        ],
        key=session_sort_key,
    )
    if not required_sessions:
        return False, ""
    missing = [
        session for session in required_sessions
        if not session_completion_lookup.get((subject_id, session), False)
    ]
    return len(missing) == 0, "; ".join(missing)


def build_session_completion_lookup(
    qlong: pd.DataFrame,
    task_long: pd.DataFrame,
    mri_long: pd.DataFrame,
    behavioral_long: pd.DataFrame,
    session_long: pd.DataFrame,
) -> dict[tuple[str, str], bool]:
    lookup: dict[tuple[str, str], bool] = {}

    def mark(subject_id: Any, session: Any, complete: bool) -> None:
        sid = clean_value(subject_id)
        sess = canonical_session_name(clean_value(session))
        if not sid or not sess:
            return
        lookup[(sid, sess)] = lookup.get((sid, sess), False) or bool(complete)

    if not qlong.empty:
        for _, item in qlong[qlong["expected"].astype(bool)].iterrows():
            mark(item["subject_id"], item["session"], item["status"] == "complete")
    if not task_long.empty:
        for _, item in task_long[task_long["expected"].astype(bool)].iterrows():
            mark(item["subject_id"], item["session"], item["status"] == "complete")
    if not mri_long.empty:
        for _, item in mri_long[mri_long["expected"].astype(bool)].iterrows():
            mark(item["subject_id"], item["session"], item["qc_status"] == "pass")
    if not behavioral_long.empty:
        for _, item in behavioral_long[behavioral_long["expected"].astype(bool)].iterrows():
            mark(item["subject_id"], item["session"], item["qc_status"] == "pass")
    if not session_long.empty and "session_completed" in session_long.columns:
        for _, item in session_long.iterrows():
            mark(item["subject_id"], item["session"], bool(item["session_completed"]))
    return lookup


def add_subject_level_summary_columns(
    subject_wise: pd.DataFrame,
    qlong: pd.DataFrame,
    task_long: pd.DataFrame,
    mri_long: pd.DataFrame,
    behavioral_long: pd.DataFrame,
    session_long: pd.DataFrame,
    required_sessions_by_arm: pd.DataFrame,
) -> pd.DataFrame:
    if subject_wise.empty:
        return subject_wise
    rows: list[dict[str, Any]] = []
    for _, row in subject_wise.iterrows():
        subject_id = row["subject_id"]
        qsub = qlong[qlong["subject_id"] == subject_id] if not qlong.empty else pd.DataFrame()
        tsub = task_long[task_long["subject_id"] == subject_id] if not task_long.empty else pd.DataFrame()
        msub = mri_long[mri_long["subject_id"] == subject_id] if not mri_long.empty else pd.DataFrame()
        bsub = behavioral_long[behavioral_long["subject_id"] == subject_id] if not behavioral_long.empty else pd.DataFrame()
        ssub = session_long[session_long["subject_id"] == subject_id] if not session_long.empty else pd.DataFrame()
        subject_arm = clean_value(row.get("arm", "")) or infer_arm_from_subject_id(subject_id)
        required_sessions_for_subject = {
            canonical_session_name(session)
            for session in required_session_set(required_sessions_by_arm, subject_arm)
            if canonical_session_name(session) not in CONFIG.get("excluded_output_sessions", [])
        }
        session_completion_lookup = build_session_completion_lookup(qlong, task_long, mri_long, behavioral_long, session_long)
        complete_all_experiment_sessions, missing_required_sessions_text = compute_complete_all_experiment_sessions(
            subject_id,
            subject_arm,
            required_sessions_by_arm,
            session_completion_lookup,
            excluded_sessions={"Screening"},
        )
        required_intervals = ssub[ssub["interval_valid"] != "not_expected"] if not ssub.empty else pd.DataFrame()
        intervals_ok = True if required_intervals.empty else bool((required_intervals["interval_valid"] == "valid").all())
        q_complete = all_expected_complete(qsub, "complete")
        task_complete = all_expected_task_complete(tsub)
        mri_pass = all_expected_qc_pass(msub)
        behavioral_pass = all_expected_qc_pass(bsub)
        all_required = q_complete and task_complete and mri_pass and behavioral_pass and intervals_ok
        dropout = row.get("dropout_status") == "withdrawn_or_dropout"
        statuses = set()
        if not qsub.empty:
            statuses.update(qsub[qsub["expected"].astype(bool)]["status"].tolist())
        if not tsub.empty:
            statuses.update(tsub[tsub["expected"].astype(bool)]["status"].tolist())
        if not msub.empty:
            statuses.update(msub[msub["expected"].astype(bool)]["qc_status"].tolist())
        if not bsub.empty:
            statuses.update(bsub[bsub["expected"].astype(bool)]["qc_status"].tolist())
        if not required_intervals.empty:
            statuses.update(required_intervals["interval_valid"].tolist())
        if not complete_all_experiment_sessions:
            statuses.add("missing_required_session")
        needs_followup = bool(statuses & {"missing", "incomplete", "unverified", "fail", "early", "late", "missing_baseline_date", "missing_session_date", "missing_required_session"})
        review_required = bool(statuses & {"review_required"})
        ready = all_required and not dropout and not review_required and not needs_followup
        updated = row.to_dict()
        updated.update(
            {
                "complete_all_experiment_sessions": complete_all_experiment_sessions,
                "missing_required_experiment_sessions": missing_required_sessions_text,
                "all_mustHave_questionnaires_complete_perSession": q_complete,
                "all_mustHave_questionnaires_complete": q_complete,
                "all_required_task_completion_complete_perSession": task_complete,
                "all_required_task_completion_complete": task_complete,
                "all_required_MRI_QC_passed_perSession": mri_pass,
                "all_required_MRI_QC_passed": mri_pass,
                "all_required_selfOther_QC_passed_perSession": behavioral_pass,
                "all_required_selfOther_QC_passed": behavioral_pass,
                "all_required_criteria_passed_perSession": all_required,
                "all_required_criteria_passed": all_required,
                "ready_for_analysis": ready,
                "needs_followup": needs_followup or review_required,
                "overall_QA_status": derive_overall_QA_status(dropout, ready, review_required, needs_followup),
            }
        )
        rows.append(updated)
    return pd.DataFrame(rows)


def reorder_subject_wise_columns(subject_wise: pd.DataFrame) -> pd.DataFrame:
    if subject_wise.empty:
        return subject_wise

    base_cols = [
        "subject_id",
        "arm",
        "dropout_status",
        "total_ASAP_count",
        "has_ASAP",
        "complete_all_experiment_sessions",
        "missing_required_experiment_sessions",
        "all_mustHave_questionnaires_complete_perSession",
        "all_mustHave_questionnaires_complete",
        "all_required_task_completion_complete_perSession",
        "all_required_task_completion_complete",
        "all_required_MRI_QC_passed_perSession",
        "all_required_MRI_QC_passed",
        "all_required_selfOther_QC_passed_perSession",
        "all_required_selfOther_QC_passed",
        "all_required_criteria_passed_perSession",
        "all_required_criteria_passed",
        "ready_for_analysis",
        "needs_followup",
        "overall_QA_status",
    ]
    base_cols = [col for col in base_cols if col in subject_wise.columns]
    remaining = [col for col in subject_wise.columns if col not in base_cols]
    session_lookup = {safe_column_fragment(session): idx for idx, session in enumerate(CONFIG.get("session_order", []))}

    def col_key(col: str) -> tuple[int, int, str]:
        match = re.match(r"ses-([^_]+)_", col)
        if not match:
            return (9999, 9999, col)
        session_rank = session_lookup.get(match.group(1), 9999)
        domain_rank = 999
        if "_qn-" in col:
            domain_rank = 1
        elif "_task-" in col:
            domain_rank = 2
        elif "_beh-" in col:
            domain_rank = 3
        elif "_mri-" in col:
            domain_rank = 4
        elif col.endswith("_session_completed") or "intervalFromBaseline" in col or col.endswith("_interval_valid"):
            domain_rank = 5
        return (session_rank, domain_rank, col)

    return subject_wise[base_cols + sorted(remaining, key=col_key)]


def empty_subject_wise_schema() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "subject_id",
            "arm",
            "dropout_status",
            "total_ASAP_count",
            "has_ASAP",
            "complete_all_experiment_sessions",
            "missing_required_experiment_sessions",
            "all_mustHave_questionnaires_complete_perSession",
            "all_mustHave_questionnaires_complete",
            "all_required_task_completion_complete_perSession",
            "all_required_task_completion_complete",
            "all_required_MRI_QC_passed_perSession",
            "all_required_MRI_QC_passed",
            "all_required_selfOther_QC_passed_perSession",
            "all_required_selfOther_QC_passed",
            "all_required_criteria_passed_perSession",
            "all_required_criteria_passed",
            "ready_for_analysis",
            "needs_followup",
            "overall_QA_status",
        ]
    )


def derive_overall_QA_status(dropout: bool, ready: bool, review_required: bool, needs_followup: bool) -> str:
    if dropout:
        return "withdrawn_or_dropout"
    if review_required:
        return "review_required"
    if ready:
        return "ready_for_analysis"
    if needs_followup:
        return "needs_followup"
    return "not_expected"


def rate(numerator: int | float, denominator: int | float) -> float | None:
    if denominator == 0:
        return None
    return numerator / denominator


def summarize_session_questionnaires(qlong: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    if qlong.empty:
        return pd.DataFrame(rows)
    source_counts = qlong.groupby(["session", "questionnaire"])["source_file"].nunique()
    for keys, group in qlong.groupby(["source_file", "session", "questionnaire"], dropna=False):
        source_file, session, questionnaire = keys
        duplicate_across_sources = source_counts.loc[(session, questionnaire)] > 1
        instrument_source = source_file if duplicate_across_sources else ""
        expected = group[group["expected"].astype(bool)]
        denominator = expected["subject_id"].nunique()
        complete_n = expected[expected["status"] == "complete"]["subject_id"].nunique()
        missing_n = expected[expected["status"] == "missing"]["subject_id"].nunique()
        incomplete_n = expected[expected["status"] == "incomplete"]["subject_id"].nunique()
        unverified_n = expected[expected["status"] == "unverified"]["subject_id"].nunique()
        rows.append(
            {
                "session": session,
                "questionnaire": questionnaire,
                "instrument_source": instrument_source,
                "expected_NofSubjects": denominator,
                "complete_NofSubjects": complete_n,
                "missing_NofSubjects": missing_n,
                "incomplete_NofSubjects": incomplete_n,
                "unverified_NofSubjects": unverified_n,
                "complete_rate": rate(complete_n, denominator),
                "missing_rate": rate(missing_n, denominator),
            }
        )
    return sort_by_session_order(pd.DataFrame(rows), ["session", "questionnaire", "instrument_source"])


def summarize_task_completion(task_long: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    task_long = filter_excluded_output_sessions(task_long)
    if task_long.empty:
        return pd.DataFrame(rows)
    source_counts = task_long.groupby(["session", "task"])["source_file"].nunique()
    for keys, group in task_long.groupby(["source_file", "session", "task"], dropna=False):
        source_file, session, task = keys
        duplicate_across_sources = source_counts.loc[(session, task)] > 1
        instrument_source = source_file if duplicate_across_sources else ""
        expected = group[group["expected"].astype(bool)]
        denominator = expected["subject_id"].nunique()
        complete_n = expected[expected["status"] == "complete"]["subject_id"].nunique()
        missing_n = expected[expected["status"] == "missing"]["subject_id"].nunique()
        incomplete_n = expected[expected["status"] == "incomplete"]["subject_id"].nunique()
        unverified_n = expected[expected["status"] == "unverified"]["subject_id"].nunique()
        rows.append(
            {
                "session": session,
                "task": task,
                "instrument_source": instrument_source,
                "expected_NofSubjects": denominator,
                "complete_NofSubjects": complete_n,
                "missing_NofSubjects": missing_n,
                "incomplete_NofSubjects": incomplete_n,
                "unverified_NofSubjects": unverified_n,
                "complete_rate": rate(complete_n, denominator),
                "missing_rate": rate(missing_n, denominator),
            }
        )
    return sort_by_session_order(pd.DataFrame(rows), ["session", "task", "instrument_source"])


def summarize_instrument_wise(
    qlong: pd.DataFrame,
    task_long: pd.DataFrame,
    behavioral_long: pd.DataFrame,
    mri_long: pd.DataFrame,
) -> pd.DataFrame:
    records: list[dict[str, Any]] = []

    def add_records(
        df: pd.DataFrame,
        domain: str,
        name_col: str,
        status_col: str,
        complete_value: str,
        missing_or_fail_values: set[str],
    ) -> None:
        df = filter_excluded_output_sessions(df)
        if df.empty:
            return
        expected = df[df["expected"].astype(bool)].copy()
        if expected.empty:
            return
        source_counts = expected.groupby([name_col])["source_file"].transform("nunique")
        expected["instrument_source_summary"] = ""
        expected.loc[source_counts > 1, "instrument_source_summary"] = expected.loc[source_counts > 1, "source_file"]
        for _, row in expected.iterrows():
            status = clean_value(row[status_col])
            records.append(
                {
                    "instrument_domain": domain,
                    "instrument_name": row[name_col],
                    "instrument_source": row["instrument_source_summary"],
                    "session": row["session"],
                    "subject_id": row["subject_id"],
                    "complete_or_pass": status == complete_value,
                    "missing_or_fail": status in missing_or_fail_values,
                    "unverified": status == "unverified",
                    "review_required": status == "review_required",
                }
            )

    add_records(qlong, "questionnaire", "questionnaire", "status", "complete", {"missing", "incomplete"})
    add_records(task_long, "task_completion", "task", "status", "complete", {"missing", "incomplete"})
    add_records(behavioral_long, "behavioral_qc", "task", "qc_status", "pass", {"missing", "fail"})
    add_records(mri_long, "mri_qc", "scan_or_run", "qc_status", "pass", {"missing", "fail"})
    if not records:
        return pd.DataFrame([])

    all_records = pd.DataFrame(records)
    rows: list[dict[str, Any]] = []
    for keys, group in all_records.groupby(["instrument_domain", "instrument_name", "instrument_source"], dropna=False):
        instrument_domain, instrument_name, instrument_source = keys
        ordered_sessions = sorted(group["session"].unique(), key=session_sort_key)
        expected_records = len(group[["subject_id", "session"]].drop_duplicates())
        expected_subjects = group["subject_id"].nunique()
        complete_n = int(group["complete_or_pass"].sum())
        missing_fail_n = int(group["missing_or_fail"].sum())
        unverified_n = int(group["unverified"].sum())
        review_required_n = int(group["review_required"].sum())
        rows.append(
            {
                "instrument_domain": instrument_domain,
                "instrument_name": instrument_name,
                "instrument_source": instrument_source,
                "first_session": ordered_sessions[0] if ordered_sessions else "",
                "last_session": ordered_sessions[-1] if ordered_sessions else "",
                "expected_NofRecords": expected_records,
                "expected_NofSubjects": expected_subjects,
                "complete_or_pass_NofRecords": complete_n,
                "missing_or_fail_NofRecords": missing_fail_n,
                "unverified_NofRecords": unverified_n,
                "review_required_NofRecords": review_required_n,
                "complete_or_pass_rate": rate(complete_n, expected_records),
                "missing_or_fail_rate": rate(missing_fail_n, expected_records),
            }
        )
    out = pd.DataFrame(rows)
    domain_order = {
        "questionnaire": 1,
        "task_completion": 2,
        "behavioral_qc": 3,
        "mri_qc": 4,
    }
    out["_domain_order"] = out["instrument_domain"].map(domain_order).fillna(999)
    out = out.sort_values(["_domain_order", "instrument_name", "instrument_source"], kind="mergesort")
    return out.drop(columns=["_domain_order"]).reset_index(drop=True)


def summarize_qc(qc_long: pd.DataFrame, section: str, item_column: str) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    if qc_long.empty:
        return pd.DataFrame(rows)
    source_counts = qc_long.groupby(["session", item_column])["source_file"].nunique()
    for keys, group in qc_long.groupby(["source_file", "session", item_column], dropna=False):
        source_file, session, item_name = keys
        duplicate_across_sources = source_counts.loc[(session, item_name)] > 1
        if section == "mri_qc_summary":
            instrument_source = first_nonblank(group.get("instrument_source", pd.Series(dtype=str)))
        else:
            instrument_source = source_file if duplicate_across_sources else ""
        expected = group[group["expected"].astype(bool)]
        denominator = expected["subject_id"].nunique()
        pass_n = expected[expected["qc_status"] == "pass"]["subject_id"].nunique()
        fail_n = expected[expected["qc_status"] == "fail"]["subject_id"].nunique()
        missing_n = expected[expected["qc_status"] == "missing"]["subject_id"].nunique()
        unverified_n = expected[expected["qc_status"] == "unverified"]["subject_id"].nunique()
        rows.append(
            {
                "session": session,
                item_column: item_name,
                "instrument_source": instrument_source,
                "expected_NofSubjects": denominator,
                "qc_pass_NofSubjects": pass_n,
                "qc_fail_NofSubjects": fail_n,
                "qc_missing_NofSubjects": missing_n,
                "qc_unverified_NofSubjects": unverified_n,
                "qc_pass_rate": rate(pass_n, denominator),
            }
        )
    return sort_by_session_order(pd.DataFrame(rows), ["session", item_column, "instrument_source"])


def summarize_mri_qc(mri_long: pd.DataFrame) -> pd.DataFrame:
    return summarize_qc(mri_long, "mri_qc_summary", "scan_or_run")


def summarize_behavioral_qc(behavioral_long: pd.DataFrame) -> pd.DataFrame:
    return summarize_qc(behavioral_long, "behavioral_qc_summary", "task")


def summarize_session_intervals(session_long: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    if session_long.empty:
        return pd.DataFrame(rows)
    for session, group in session_long.groupby("session", dropna=False):
        expected = group[group["interval_valid"] != "not_expected"]
        denominator = expected["subject_id"].nunique()
        valid_n = expected[expected["interval_valid"] == "valid"]["subject_id"].nunique()
        early_n = expected[expected["interval_valid"] == "early"]["subject_id"].nunique()
        late_n = expected[expected["interval_valid"] == "late"]["subject_id"].nunique()
        missing_baseline_n = expected[expected["interval_valid"] == "missing_baseline_date"]["subject_id"].nunique()
        missing_session_n = expected[expected["interval_valid"] == "missing_session_date"]["subject_id"].nunique()
        numeric_intervals = pd.to_numeric(expected["intervalFromBaseline_weeks"], errors="coerce")
        target = expected["expected_target_weeks"].replace("", pd.NA).dropna()
        target_value = pd.to_numeric(target, errors="coerce").dropna()
        target_float = float(target_value.iloc[0]) if not target_value.empty else None
        deviations = numeric_intervals - target_float if target_float is not None else pd.Series(dtype=float)
        rows.append(
            {
                "session": session,
                "expected_target_weeks": target_float,
                "allowed_min_weeks": first_nonblank(expected.get("allowed_min_weeks", pd.Series(dtype=str))),
                "allowed_max_weeks": first_nonblank(expected.get("allowed_max_weeks", pd.Series(dtype=str))),
                "expected_NofSubjects": denominator,
                "interval_valid_NofSubjects": valid_n,
                "interval_valid_rate": rate(valid_n, denominator),
                "mean_intervalFromBaseline_weeks": numeric_intervals.mean(),
                "std_intervalFromBaseline_weeks": numeric_intervals.std(),
                "early_NofSubjects": early_n,
                "early_rate": rate(early_n, denominator),
                "late_NofSubjects": late_n,
                "late_rate": rate(late_n, denominator),
                "missing_baseline_date_NofSubjects": missing_baseline_n,
                "missing_session_date_NofSubjects": missing_session_n,
            }
        )
    return sort_by_session_order(pd.DataFrame(rows), ["session"])


def first_nonblank(series: pd.Series) -> Any:
    for value in series:
        cleaned = clean_value(value)
        if cleaned:
            return cleaned
    return ""


def summarize_participant_status_and_asap(subject_wise: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    if subject_wise.empty:
        return pd.DataFrame(rows)
    denominator = len(subject_wise)

    def add_row(metric: str, count: int) -> None:
        rows.append(
            {
                "status_metric": metric,
                "NofSubjects": count,
                "denominator_NofSubjects": denominator,
                "rate": rate(count, denominator),
            }
        )

    withdrawn_n = int((subject_wise["dropout_status"] == "withdrawn_or_dropout").sum())
    add_row("withdrawn_or_dropout", withdrawn_n)

    bool_metrics = [
        "complete_all_experiment_sessions",
        "all_mustHave_questionnaires_complete",
        "all_required_task_completion_complete",
        "all_required_MRI_QC_passed",
        "all_required_selfOther_QC_passed",
        "all_required_criteria_passed",
        "ready_for_analysis",
        "needs_followup",
    ]
    for metric in bool_metrics:
        if metric not in subject_wise.columns:
            continue
        values = subject_wise[metric]
        count = int(values.sum()) if values.dtype == bool else int((values.astype(str).str.lower() == "true").sum())
        add_row(metric, count)

    review_required_n = int((subject_wise["overall_QA_status"] == "review_required").sum())
    add_row("review_required", review_required_n)

    counts = pd.to_numeric(subject_wise["total_ASAP_count"], errors="coerce").fillna(0)
    add_row("subjects_with_0_ASAP", int((counts == 0).sum()))
    add_row("subjects_with_1_ASAP", int((counts == 1).sum()))
    add_row("subjects_with_2plus_ASAP", int((counts >= 2).sum()))

    out = pd.DataFrame(rows)
    duplicated = out["status_metric"][out["status_metric"].duplicated()].tolist()
    if duplicated:
        raise ValueError(f"Duplicated participant-level summary metrics: {duplicated}")
    forbidden = {"dropout", "completed_experiment_sessions"}
    present_forbidden = sorted(set(out["status_metric"]) & forbidden)
    if present_forbidden:
        raise ValueError(f"Forbidden participant-level summary metrics: {present_forbidden}")
    return out


def build_group_wise_tables(
    qlong: pd.DataFrame,
    task_long: pd.DataFrame,
    mri_long: pd.DataFrame,
    behavioral_long: pd.DataFrame,
    session_long: pd.DataFrame,
    subject_wise: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    return {
        "Table 1. Questionnaire completion and missingness by session": summarize_session_questionnaires(qlong),
        "Table 2. Instrument-wise summary": summarize_instrument_wise(qlong, task_long, behavioral_long, mri_long),
        "Table 3. Task completion by session": summarize_task_completion(task_long),
        "Table 4. Behavioral QC pass rate by session": summarize_behavioral_qc(behavioral_long),
        "Table 5. MRI QC pass rate by session": summarize_mri_qc(mri_long),
        "Table 6. Session interval summary after Baseline": summarize_session_intervals(session_long),
        "Table 7. Participant-level QA readiness summary": summarize_participant_status_and_asap(subject_wise),
    }


def run_validation_checks(
    df: pd.DataFrame,
    subject_wise: pd.DataFrame,
    qlong: pd.DataFrame,
    task_long: pd.DataFrame,
    mri_long: pd.DataFrame,
    behavioral_long: pd.DataFrame,
    log: ValidationLog,
    counter_prefix: str = "",
) -> None:
    for source_file, group in df.groupby("source_file", dropna=False):
        log.set_counter(f"{counter_prefix}subjects_in_{source_file}", group["subject_id"].replace("", pd.NA).dropna().nunique())
    log.set_counter(f"{counter_prefix}unique_subject_ids_in_subject_wise", len(subject_wise))
    if "subject_id" in subject_wise.columns:
        log.info(
            f"{counter_prefix}standardized_subject_id_examples: "
            f"{subject_wise['subject_id'].drop_duplicates().head(10).tolist()}"
        )
    missing_mri = 0
    missing_task = 0
    if not task_long.empty:
        missing_task = task_long[
            (task_long["expected"].astype(bool))
            & (task_long["status"].isin(["missing", "incomplete", "unverified", "review_required"]))
        ]["subject_id"].nunique()
    if not mri_long.empty:
        missing_mri = mri_long[
            (mri_long["expected"].astype(bool))
            & (mri_long["qc_status"].isin(["missing", "unverified", "review_required", "fail"]))
        ]["subject_id"].nunique()
    missing_behavioral = 0
    if not behavioral_long.empty:
        missing_behavioral = behavioral_long[
            (behavioral_long["expected"].astype(bool))
            & (behavioral_long["qc_status"].isin(["missing", "unverified", "review_required", "fail"]))
        ]["subject_id"].nunique()
    log.set_counter(f"{counter_prefix}subjects_with_missing_required_task_completion", missing_task)
    log.set_counter(f"{counter_prefix}subjects_with_missing_required_MRI_QC", missing_mri)
    log.set_counter(f"{counter_prefix}subjects_with_missing_required_self_other_QC", missing_behavioral)
    if "needs_followup" in subject_wise.columns:
        log.set_counter(f"{counter_prefix}subjects_needing_followup", int(subject_wise["needs_followup"].sum()))
    if qlong.empty:
        log.warn("unknown_form_or_questionnaire_names", "No questionnaire/form completion fields were detected.")


def dataframe_contains_token(df: pd.DataFrame, token: str) -> bool:
    if df.empty:
        return False
    if any(token.lower() in str(col).lower() for col in df.columns):
        return True
    values = df.astype(str)
    return bool(values.apply(lambda col: col.str.contains(token, case=False, regex=False, na=False)).any().any())


def validate_final_output(
    subject_wise: pd.DataFrame,
    group_tables: dict[str, pd.DataFrame],
    qlong: pd.DataFrame,
    task_long: pd.DataFrame,
    mri_long: pd.DataFrame,
    behavioral_long: pd.DataFrame,
    session_long: pd.DataFrame,
    log: ValidationLog,
    allow_unmapped_events: bool = False,
) -> None:
    if "source_group" in subject_wise.columns:
        raise ValueError("source_group should not appear in subject_wise output.")
    duplicated_subjects = subject_wise["subject_id"].duplicated().sum() if "subject_id" in subject_wise.columns else 0
    if duplicated_subjects:
        raise ValueError(f"subject_wise has duplicated subject_id rows: {duplicated_subjects}")
    duplicated_columns = subject_wise.columns[subject_wise.columns.duplicated()].tolist()
    if duplicated_columns:
        raise ValueError(f"Duplicate columns in subject_wise: {duplicated_columns}")
    if "completed_experiment_sessions" in subject_wise.columns:
        raise ValueError("Use complete_all_experiment_sessions, not completed_experiment_sessions.")
    if "complete_all_experiment_sessions" not in subject_wise.columns:
        raise ValueError("Missing required subject-wise column: complete_all_experiment_sessions.")
    if "missing_required_experiment_sessions" not in subject_wise.columns:
        raise ValueError("Missing required subject-wise column: missing_required_experiment_sessions.")
    if "subject_id" in subject_wise.columns:
        prefixed_subjects = subject_wise["subject_id"].map(lambda value: clean_value(value).lower().startswith("sub-"))
        if prefixed_subjects.any():
            raise ValueError("Final subject_wise subject_id values must not start with sub-.")
        invalid_arms = sorted(set(subject_wise.get("arm", pd.Series(dtype=str)).map(clean_value)) - set(CONFIG.get("arms", [])) - {""})
        if invalid_arms:
            raise ValueError(f"Final subject_wise contains invalid arms: {invalid_arms}")
    total_score_cols = [
        col for col in subject_wise.columns
        if "totalscore" in col.lower() or "total_score" in col.lower()
    ]
    if total_score_cols:
        raise ValueError(f"Total score columns should not be in final output: {total_score_cols}")
    bad_columns = [col for col in subject_wise.columns if "_arm_" in col.lower()]
    if bad_columns:
        raise ValueError(f"Final subject_wise columns contain REDCap unique event names: {bad_columns}")
    for bad_text in ["Screening", "Baseline 1", "Baseline1"]:
        if any(bad_text in str(col) for col in subject_wise.columns):
            raise ValueError(f"Forbidden text found in subject_wise columns: {bad_text}")
        if dataframe_contains_token(subject_wise, bad_text):
            raise ValueError(f"Forbidden text found in subject_wise values: {bad_text}")
    baseline_interval_cols = [
        col for col in subject_wise.columns
        if "ses-Baseline" in col and "intervalFromBaseline" in col
    ]
    if baseline_interval_cols:
        raise ValueError(f"Baseline interval columns should not exist: {baseline_interval_cols}")
    screening_interval_cols = [
        col for col in subject_wise.columns
        if "ses-Screening" in col and "intervalFromBaseline" in col
    ]
    if screening_interval_cols:
        raise ValueError(f"Screening interval columns should not exist: {screening_interval_cols}")
    screening_cols = [col for col in subject_wise.columns if "ses-Screening" in col]
    if screening_cols:
        raise ValueError(f"Screening columns should not appear in subject_wise: {screening_cols}")
    if not mri_long.empty:
        mri_cols = [col for col in subject_wise.columns if "_mri-" in col]
        if not mri_cols:
            raise ValueError("MRI QC rows were loaded, but no MRI QC columns were written to subject_wise.")
        if (mri_long["session"] == "Screening").any():
            raise ValueError("Screening should not appear in mri_qc_long final output.")
    if not allow_unmapped_events and dataframe_contains_token(subject_wise, "UNMAPPED_EVENT_"):
        raise ValueError("Unmapped REDCap events remain in subject_wise output.")

    baseline_q = qlong[qlong["session"] == "Baseline"] if not qlong.empty else pd.DataFrame()
    if not baseline_q.empty:
        baseline_q_cols = [
            col for col in subject_wise.columns
            if col.startswith("ses-Baseline_qn-") and (col.endswith("_date") or col.endswith("_status"))
        ]
        if not baseline_q_cols:
            raise ValueError("Expected Baseline questionnaire date/status columns are missing from subject_wise.")

    expected_table_names = [
        "Table 1. Questionnaire completion and missingness by session",
        "Table 2. Instrument-wise summary",
        "Table 3. Task completion by session",
        "Table 4. Behavioral QC pass rate by session",
        "Table 5. MRI QC pass rate by session",
        "Table 6. Session interval summary after Baseline",
        "Table 7. Participant-level QA readiness summary",
    ]
    missing_tables = [name for name in expected_table_names if name not in group_tables]
    if missing_tables:
        raise ValueError(f"Missing group-wise tables: {missing_tables}")
    if len(group_tables) != len(expected_table_names):
        raise ValueError(f"Expected {len(expected_table_names)} group-wise tables, found {len(group_tables)}")
    for table_title, table_df in group_tables.items():
        if "_arm_" in table_title.lower() or dataframe_contains_token(table_df, "_arm_"):
            raise ValueError(f"Final group_wise table contains REDCap unique event names: {table_title}")
        for bad_text in ["Screening", "Baseline 1", "Baseline1"]:
            if bad_text in table_title or dataframe_contains_token(table_df, bad_text):
                raise ValueError(f"Forbidden text found in {table_title}: {bad_text}")
        if not allow_unmapped_events and (
            "UNMAPPED_EVENT_" in table_title or dataframe_contains_token(table_df, "UNMAPPED_EVENT_")
        ):
            raise ValueError(f"Unmapped REDCap events remain in final output: {table_title}")
    participant_table = group_tables.get("Table 7. Participant-level QA readiness summary")
    if participant_table is not None and not participant_table.empty:
        duplicated = participant_table["status_metric"][participant_table["status_metric"].duplicated()].tolist()
        if duplicated:
            raise ValueError(f"Duplicated participant-level metrics: {duplicated}")
        forbidden = {"dropout", "completed_experiment_sessions"}
        present = sorted(set(participant_table["status_metric"]) & forbidden)
        if present:
            raise ValueError(f"Forbidden participant-level metrics: {present}")
    for name, table_df in {
        "questionnaire_long": qlong,
        "task_completion_long": task_long,
        "mri_qc_long": mri_long,
        "behavioral_qc_long": behavioral_long,
        "session_long": session_long,
    }.items():
        if dataframe_contains_token(table_df, "Screening"):
            raise ValueError(f"Screening should not appear in final-use {name}.")
    log.set_counter("group_wise_tables_written", len(group_tables))


def apply_excel_formatting(xlsx_path: Path) -> None:
    try:
        from openpyxl import load_workbook
        from openpyxl.formatting.rule import CellIsRule
        from openpyxl.styles import Font, PatternFill
    except Exception:
        return
    workbook = load_workbook(xlsx_path)
    fill_red = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    fill_yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_green = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    fill_gray = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    for worksheet in workbook.worksheets:
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        for cell in worksheet[1]:
            cell.font = Font(bold=True)
        for column_cells in worksheet.columns:
            max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_length + 2, 10), 45)
        max_row = worksheet.max_row
        max_col = worksheet.max_column
        if max_row > 1 and max_col > 0:
            ref = f"A2:{worksheet.cell(max_row, max_col).coordinate}"
            for text in ["missing", "fail", "needs_followup"]:
                worksheet.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill_red))
            for text in ["incomplete", "unverified", "review_required"]:
                worksheet.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill_yellow))
            for text in ["complete", "pass", "ready_for_analysis"]:
                worksheet.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=[f'"{text}"'], fill=fill_green))
            worksheet.conditional_formatting.add(ref, CellIsRule(operator="equal", formula=['"not_expected"'], fill=fill_gray))
    workbook.save(xlsx_path)


def write_group_wise_tables(writer: pd.ExcelWriter, group_tables: dict[str, pd.DataFrame]) -> None:
    sheet_name = "group_wise"
    startrow = 0
    for table_title, table_df in group_tables.items():
        pd.DataFrame([[table_title]]).to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=startrow,
            index=False,
            header=False,
        )
        startrow += 1
        table_df.to_excel(
            writer,
            sheet_name=sheet_name,
            startrow=startrow,
            index=False,
        )
        startrow += len(table_df) + 3


def write_QA_summary_xlsx(
    subject_wise: pd.DataFrame,
    group_tables: dict[str, pd.DataFrame],
    output_path: Path,
) -> None:
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        subject_wise.to_excel(writer, sheet_name="subject_wise", index=False)
        write_group_wise_tables(writer, group_tables)
    apply_excel_formatting(output_path)


def load_inputs(input_dir: Path, log: ValidationLog) -> pd.DataFrame:
    frames = [
        read_csv_as_string(input_dir / filename, source_file, log)
        for source_file, filename in INPUT_FILES.items()
    ]
    frames = [frame for frame in frames if not frame.empty]
    if not frames:
        raise FileNotFoundError("No REDCap QA tracker CSV files were found.")
    return pd.concat(frames, ignore_index=True, sort=False).fillna("")


def validate_output_files(output_paths: list[Path], log: ValidationLog) -> None:
    for output_path in output_paths:
        if not output_path.exists():
            raise FileNotFoundError(f"Expected output file was not written: {output_path}")
        try:
            workbook = pd.ExcelFile(output_path)
            missing_sheets = [sheet for sheet in ["subject_wise", "group_wise"] if sheet not in workbook.sheet_names]
            if missing_sheets:
                raise ValueError(f"{output_path.name} is missing sheets: {missing_sheets}")
        except Exception as exc:
            raise ValueError(f"Could not validate output workbook {output_path}: {exc}") from exc
        log.warn("output_file_validation", f"validated {output_path.name}")


def build_outputs(input_dir: Path, output_dir: Path, allow_unmapped_events: bool = False) -> tuple[list[Path], Path]:
    log = ValidationLog()
    log.info(f"Input directory: {input_dir}")
    log.info(f"Output directory: {output_dir}")
    design_tables = load_design_tables(input_dir, log)
    required_sessions = design_tables["required_sessions_by_arm"]
    required_instruments = design_tables["required_instruments_by_arm_session"]
    df = load_inputs(input_dir, log)
    df, _ = standardize_subject_ids(df, log)
    df = add_arm_assignment(df, log)
    event_mapping = add_design_event_mappings(get_codebook_event_mapping(input_dir, log), required_sessions)
    df, _ = map_event_names(df, event_mapping, log)
    qlong = build_questionnaire_long(df, log)
    task_long = build_task_completion_long(df, log)
    mri_long = build_mri_qc_long(input_dir, log)
    behavioral_long = build_behavioral_qc_long(df, log)
    session_long = build_session_long(df, qlong, required_sessions, log)

    subject_arm_map = (
        df[["subject_id", "arm"]]
        .drop_duplicates("subject_id")
        .set_index("subject_id")["arm"]
        .to_dict()
    )
    qlong = add_arm_to_long_table(qlong, subject_arm_map)
    task_long = add_arm_to_long_table(task_long, subject_arm_map)
    behavioral_long = add_arm_to_long_table(behavioral_long, subject_arm_map)
    session_long = add_arm_to_long_table(session_long, subject_arm_map)

    qlong = filter_excluded_output_sessions(qlong)
    task_long = filter_excluded_output_sessions(task_long)
    mri_long = filter_excluded_output_sessions(mri_long)
    behavioral_long = filter_excluded_output_sessions(behavioral_long)
    session_long = filter_excluded_output_sessions(session_long)

    qlong = apply_required_questionnaire_design(
        qlong,
        df[["subject_id", "arm"]].drop_duplicates("subject_id"),
        required_instruments,
        log,
    )
    task_long = filter_long_to_required_sessions(task_long, required_sessions, log, "task_completion_long")
    mri_long = filter_long_to_required_sessions(mri_long, required_sessions, log, "mri_qc_long")
    mri_long = apply_observation_based_mri_expectedness(mri_long, required_instruments, log)
    behavioral_long = filter_long_to_required_sessions(behavioral_long, required_sessions, log, "behavioral_qc_long")
    session_long = filter_long_to_required_sessions(session_long, required_sessions, log, "session_long")

    source_presence_frames = [df[["subject_id", "source_file"]].copy()]
    if not mri_long.empty:
        source_presence_frames.append(mri_long[["subject_id", "source_file"]].copy())
    source_presence = pd.concat(source_presence_frames, ignore_index=True, sort=False).drop_duplicates()
    if not source_presence.empty:
        presence = (
            source_presence.assign(present=True)
            .pivot_table(index="subject_id", columns="source_file", values="present", aggfunc="any", fill_value=False)
            .reset_index()
        )
        log.info(f"cross_file_merge_subject_presence_examples: {presence.head(10).to_dict('records')}")
        log.set_counter("subjects_present_in_multiple_sources_after_standardization", int((presence.drop(columns=["subject_id"]).sum(axis=1) > 1).sum()))

    output_dir.mkdir(parents=True, exist_ok=True)
    xlsx_paths: list[Path] = []
    output_template = CONFIG.get("output_file_template", "QA_summary_{arm}.xlsx")
    for arm in CONFIG.get("arms", ["arm1", "arm2", "arm3"]):
        arm_df = arm_filter(df, arm)
        arm_qlong = arm_filter(qlong, arm)
        arm_task_long = arm_filter(task_long, arm)
        arm_mri_long = arm_filter(mri_long, arm)
        arm_behavioral_long = arm_filter(behavioral_long, arm)
        arm_session_long = arm_filter(session_long, arm)
        subject_wise = pivot_long_tables_to_subject_wise(
            arm_df,
            arm_qlong,
            arm_task_long,
            arm_mri_long,
            arm_behavioral_long,
            arm_session_long,
            required_sessions,
        )
        if subject_wise.empty:
            subject_wise = empty_subject_wise_schema()
        subject_wise = reorder_subject_wise_columns(subject_wise)
        required_sessions_checked = sorted(
            [session for session in required_session_set(required_sessions, arm) if session != "Screening"],
            key=session_sort_key,
        )
        log.info(f"required_sessions_checked_for_complete_all_experiment_sessions_{arm}: {required_sessions_checked}")
        log.set_counter(f"{arm}_required_session_count", len(required_session_set(required_sessions, arm)))
        log.set_counter(
            f"{arm}_subjects_missing_required_experiment_sessions",
            int((subject_wise["complete_all_experiment_sessions"] == False).sum()),
        )
        group_tables = build_group_wise_tables(
            arm_qlong,
            arm_task_long,
            arm_mri_long,
            arm_behavioral_long,
            arm_session_long,
            subject_wise,
        )
        group_tables = add_arm_column_to_group_tables(group_tables, arm)
        run_validation_checks(
            arm_df,
            subject_wise,
            arm_qlong,
            arm_task_long,
            arm_mri_long,
            arm_behavioral_long,
            log,
            counter_prefix=f"{arm}_",
        )
        validate_final_output(
            subject_wise,
            group_tables,
            arm_qlong,
            arm_task_long,
            arm_mri_long,
            arm_behavioral_long,
            arm_session_long,
            log,
            allow_unmapped_events=allow_unmapped_events,
        )
        xlsx_path = output_dir / output_template.format(arm=arm)
        write_QA_summary_xlsx(subject_wise, group_tables, xlsx_path)
        xlsx_paths.append(xlsx_path)

    validate_output_files(xlsx_paths, log)
    log_path = output_dir / VALIDATION_LOG
    log.write(log_path)
    return xlsx_paths, log_path


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate QA_summary.xlsx from REDCap QA tracker exports.")
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=Path("."),
        help="Directory containing REDCap CSV exports and PDF codebooks.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("."),
        help="Directory where QA_summary.xlsx and QA_summary_validation_log.txt should be written.",
    )
    parser.add_argument(
        "--allow-unmapped-events",
        action="store_true",
        help="Allow final workbook output with UNMAPPED_EVENT_ session labels for debugging.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    xlsx_paths, log_path = build_outputs(
        args.input_dir.resolve(),
        args.output_dir.resolve(),
        allow_unmapped_events=args.allow_unmapped_events,
    )
    for xlsx_path in xlsx_paths:
        print(f"Wrote {xlsx_path}")
    print(f"Wrote {log_path}")


if __name__ == "__main__":
    main()
