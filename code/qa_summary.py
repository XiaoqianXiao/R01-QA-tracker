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
    "participants": "IFOCUSStudyParticipa-QAtracker_DATA_*.csv",
    "clinician": "IFOCUSStudyClinician-QAtracker_DATA_*.csv",
}
DESIGN_FILES = {
    "required_sessions": "Required_Sessions_for_each_Arm.xlsx",
    "expected_instruments": "Instruments_in_each_Session_each_Arm.xlsx",
}
MRI_FILES = {
    #"anat": "QC_anat.csv",
    "func": "QC_func.csv",
}
CODEBOOK_FILES = ("Participants_REDCap.pdf", "Clini_REDCap.pdf")
STANDARD_INTERVAL_WEEKS = {
    "Repeat Baseline": 4,
    "T12": 16,
}
INTERVAL_SUMMARY_SESSIONS = tuple(STANDARD_INTERVAL_WEEKS)
TE_SCAN_PAIRS = (
    ("T6", ("IE T6",)),
    ("T12", ("IE T12", "T12")),
)
STANDARD_TE_SCAN_INTERVAL_DAYS = 10

# ---------------------------------------------------------------------------
# Configure expected sessions and instruments here.
#
# Default behavior follows PROJECT_CONTEXT.md exactly: required sessions and
# expected instruments are read from the two design workbooks, and only
# Included_in_summary == 1 instruments are included. The optional settings below
# let a user narrow, exclude, or add expected items without changing processing
# logic elsewhere in the script.
#
# Common examples:
#   Keep only Baseline for every active arm:
#       "required_sessions": {"include_only": ["Baseline"], ...}
#   Keep only selected sessions for one arm:
#       "required_sessions": {"include_only": {"arm1": ["Baseline", "T1"]}, ...}
#   Exclude one expected instrument:
#       "expected_instruments": {
#           "exclude": [
#               {"arm": "arm1", "session": "T12", "instrument": "SCID-5 T12"},
#           ],
#           ...
#       }
#   Include or exclude subjects:
#       "subjects": {"include_only": ["101", "102"], "exclude": ["105"]}
#       "subjects": {"include_only": {"arm1": ["101"], "arm3": ["301"]}}
# ---------------------------------------------------------------------------
QA_CONFIG: dict[str, Any] = {
    # Which output arms to build. Leave as all three arms unless intentionally
    # generating a subset.
    "active_arms": ["arm1", "arm2", "arm3"],
    "required_sessions": {
        # If an arm is listed here, only these sessions are retained for that arm.
        # Session names use final labels, e.g. "Baseline", "T1", "IE T3".
        "include_only": [],
        # Remove sessions after reading the design workbook.
        "exclude": {},
        # Add custom sessions. Each item may include arm, session, order,
        # participant_event_name, and clinician_event_name.
        "add": [],
    },
    "expected_instruments": {
        # If an arm is listed here, only these configured instruments are retained
        # for that arm. Shape: {"arm1": {"Baseline": ["PHQ-9", "GAD-7"]}}.
        "include_only": {},
        # Remove expected instruments after Included_in_summary == 1 filtering.
        # Items may specify arm, session, instrument, and/or instrument_source.
        "exclude": [],
        # Add expected instruments. Each item requires arm, session, instrument,
        # and instrument_source.
        "add": [],
    },
    "subjects": {
        # Optional subject filters are applied after subject IDs are standardized.
        # Accepted shapes:
        #   [] or {}                  -> no filtering
        #   ["101", "102"]           -> apply to all active arms
        #   {"arm1": ["101", "102"]} -> apply to specific arms
        "include_only": [],
        "exclude": [],
    },
}

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
    #"ses-t6": "T6 Scan",
    "ses-t6": "IE T6",
    "ses-t7": "T7",
    "ses-t8": "T8",
    "ses-t9": "T9",
    "ses-t10": "T10",
    "ses-t11": "T11",
    #"ses-t12": "T12 Scan",
    "ses-t12": "IE T12",
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


def design_flag_enabled(value: Any) -> bool:
    """Interpret design-workbook flags such as 1, 1.0, True, or Yes."""
    return clean(value).lower() in TRUE_VALUES | {"1.0"}


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


def configured_arms(config: dict[str, Any]) -> list[str]:
    arms = [clean(arm) for arm in config.get("active_arms", ARMS)]
    return [arm for arm in arms if arm in ARMS]


def configured_session_set(values: Iterable[Any]) -> set[str]:
    if isinstance(values, str):
        values = [values]
    return {canonical_session(value) for value in values if clean(value)}


def normalize_arm_session_config(value: Any, active_arms: Iterable[str]) -> dict[str, set[str]]:
    if not value:
        return {}
    if isinstance(value, dict):
        return {clean(arm): configured_session_set(sessions) for arm, sessions in value.items()}
    sessions = configured_session_set(value)
    return {arm: sessions for arm in active_arms}


def normalize_subject_filter(value: Any, active_arms: Iterable[str]) -> dict[str, set[str]]:
    if not value:
        return {}
    if isinstance(value, dict):
        return {
            clean(arm): {standardize_subject_id(subject) for subject in subjects if standardize_subject_id(subject)}
            for arm, subjects in value.items()
        }
    if isinstance(value, str):
        value = [value]
    subjects = {standardize_subject_id(subject) for subject in value if standardize_subject_id(subject)}
    return {arm: subjects for arm in active_arms}


def normalize_expected_include_config(value: Any, active_arms: Iterable[str]) -> dict[str, dict[str, set[str]]]:
    if not value:
        return {}
    if not isinstance(value, dict):
        raise ValueError(
            "QA_CONFIG['expected_instruments']['include_only'] must be a dict, "
            'for example {"arm1": {"Baseline": ["PHQ-9"]}}'
        )
    active = set(active_arms)
    has_arm_keys = any(clean(key) in active for key in value)
    if not has_arm_keys:
        return {
            arm: {canonical_session(session): {clean(item) for item in instruments} for session, instruments in value.items()}
            for arm in active
        }
    normalized: dict[str, dict[str, set[str]]] = {}
    for arm, by_session in value.items():
        arm = clean(arm)
        if not isinstance(by_session, dict):
            raise ValueError(
                f"QA_CONFIG expected_instruments.include_only for {arm} must map sessions to instrument lists"
            )
        normalized[arm] = {
            canonical_session(session): {clean(item) for item in instruments}
            for session, instruments in by_session.items()
        }
    return normalized


def apply_subject_config(subjects: pd.DataFrame, config: dict[str, Any], log: ValidationLog) -> pd.DataFrame:
    subject_config = config.get("subjects", {})
    out = subjects.copy()
    active = set(configured_arms(config))
    out = out[out["arm"].isin(active)].copy()
    include_only = normalize_subject_filter(subject_config.get("include_only", []), active)
    if include_only:
        before = len(out)
        keep_mask = pd.Series(False, index=out.index)
        for arm, keep_subjects in include_only.items():
            keep_mask |= (out["arm"] == arm) & (out["subject_id"].isin(keep_subjects))
        out = out[keep_mask].copy()
        log.info(f"Config subjects.include_only: retained {len(out)} of {before} subjects")
    exclude = normalize_subject_filter(subject_config.get("exclude", []), active)
    for arm, drop_subjects in exclude.items():
        before = len(out)
        out = out[~((out["arm"] == arm) & (out["subject_id"].isin(drop_subjects)))].copy()
        log.info(f"Config subjects.exclude {arm}: removed {before - len(out)} subjects")
    return out.sort_values(["arm", "subject_id"]).reset_index(drop=True)


def apply_required_session_config(
    required_sessions: pd.DataFrame,
    config: dict[str, Any],
    log: ValidationLog,
) -> pd.DataFrame:
    session_config = config.get("required_sessions", {})
    out = required_sessions.copy()
    active = set(configured_arms(config))
    out = out[out["arm"].isin(active)].copy()
    include_only = normalize_arm_session_config(session_config.get("include_only", {}), active)
    for arm, keep_sessions in include_only.items():
        before = len(out)
        out = out[(out["arm"] != arm) | (out["session"].isin(keep_sessions))].copy()
        log.info(f"Config required_sessions.include_only {arm}: retained {len(out)} of {before} rows")
    exclude = normalize_arm_session_config(session_config.get("exclude", {}), active)
    for arm, drop_sessions in exclude.items():
        before = len(out)
        out = out[~((out["arm"] == arm) & (out["session"].isin(drop_sessions)))].copy()
        log.info(f"Config required_sessions.exclude {arm}: removed {before - len(out)} rows")
    additions = []
    for item in session_config.get("add", []):
        arm = clean(item.get("arm", ""))
        session = canonical_session(item.get("session", ""))
        if arm not in active or not session:
            log.warn(f"Skipping invalid configured required session: {item}")
            continue
        additions.append(
            {
                "arm": arm,
                "session": session,
                "order": int(clean(item.get("order", "999")) or 999),
                "participant_event_name": clean(item.get("participant_event_name", "")),
                "clinician_event_name": clean(item.get("clinician_event_name", "")),
            }
        )
    if additions:
        out = pd.concat([out, pd.DataFrame(additions)], ignore_index=True)
        log.info(f"Config required_sessions.add: added {len(additions)} rows")
    out = out.drop_duplicates(["arm", "session"], keep="last").sort_values(["arm", "order"]).reset_index(drop=True)
    return out


def instrument_config_matches(row: pd.Series, rule: dict[str, Any]) -> bool:
    for key in ("arm", "session", "instrument", "instrument_source"):
        value = clean(rule.get(key, ""))
        if not value:
            continue
        row_value = canonical_session(row[key]) if key == "session" else clean(row[key])
        rule_value = canonical_session(value) if key == "session" else value
        if row_value != rule_value:
            return False
    return True


def apply_expected_instrument_config(
    expected: pd.DataFrame,
    required_sessions: pd.DataFrame,
    config: dict[str, Any],
    log: ValidationLog,
) -> pd.DataFrame:
    instrument_config = config.get("expected_instruments", {})
    out = expected.copy()
    active = set(configured_arms(config))
    out = out[out["arm"].isin(active)].copy()
    include_only = normalize_expected_include_config(instrument_config.get("include_only", {}), active)
    for arm, by_session in include_only.items():
        allowed: set[tuple[str, str]] = set()
        for session, instruments in by_session.items():
            for instrument in instruments:
                allowed.add((session, clean(instrument)))
        before = len(out)
        out = out[
            (out["arm"] != arm)
            | out.apply(lambda row: (row["session"], row["instrument"]) in allowed, axis=1)
        ].copy()
        log.info(f"Config expected_instruments.include_only {arm}: retained {len(out)} of {before} rows")
    for rule in instrument_config.get("exclude", []):
        before = len(out)
        out = out[~out.apply(lambda row: instrument_config_matches(row, rule), axis=1)].copy()
        log.info(f"Config expected_instruments.exclude {rule}: removed {before - len(out)} rows")
    required_by_arm = {arm: set(group["session"]) for arm, group in required_sessions.groupby("arm", sort=False)}
    additions = []
    for item in instrument_config.get("add", []):
        arm = clean(item.get("arm", ""))
        session = canonical_session(item.get("session", ""))
        instrument = clean(item.get("instrument", ""))
        source = clean(item.get("instrument_source", ""))
        if arm not in active or session not in required_by_arm.get(arm, set()) or not instrument or not source:
            log.warn(f"Skipping invalid configured expected instrument: {item}")
            continue
        additions.append(
            {
                "arm": arm,
                "session": session,
                "instrument": instrument,
                "instrument_source": source,
                "total_score_required": design_flag_enabled(item.get("total_score", "")),
            }
        )
    if additions:
        out = pd.concat([out, pd.DataFrame(additions)], ignore_index=True)
        log.info(f"Config expected_instruments.add: added {len(additions)} rows")
    if "total_score_required" not in out.columns:
        out["total_score_required"] = False
    out["total_score_required"] = out["total_score_required"].map(
        lambda value: design_flag_enabled(value)
    )
    key_cols = ["arm", "session", "instrument", "instrument_source"]
    return (
        out.sort_values("total_score_required", ascending=False)
        .drop_duplicates(key_cols, keep="first")
        .reset_index(drop=True)
    )


def final_session_allowed(session: str) -> bool:
    session = canonical_session(session)
    return session != "Screening" and not session.upper().startswith("ASAP")


def standardize_subject_id(raw: Any) -> str:
    digits = "".join(re.findall(r"\d", clean(raw)))
    return digits[-3:] if len(digits) >= 3 else digits


def infer_arm(subject_id: str) -> str:
    return {"1": "arm1", "2": "arm3", "3": "arm2"}.get(clean(subject_id)[:1], "unknown")


def status_from_complete_code(value: Any, field_name: str = "") -> str:
    # Protect against Pandas Series
    if type(value).__name__ in ('Series', 'DataFrame'):
        value = value.iloc[0] if not value.empty else ""
        
    text = clean(value).lower()

    # --- ANT EXCEPTION LOGIC ---
    # Dynamically check if the field is one of our ANT fields
    if field_name in ANT_FIELDS_BY_SESSION.values():
        if text in {"1", "1.0", "yes", "y", "true", "checked"}:
            return "complete"
        if text in {"0", "0.0", "na", "n/a", "not applicable", "none"}:
            return "complete"  # Treat NA as complete so QA passes
        if text in {"2", "2.0", "no", "n", "false", "unchecked"}:
            return "incomplete"
            
        if text:
            print(f"⚠️ [QA Alert] Unrecognized ANT value for '{field_name}': '{text}'")
            return "review_required"
        return "missing"

    # --- STANDARD REDCAP LOGIC ---
    if text == "2":
        return "complete"
    if text == "1":
        return "unverified"
    if text == "0":
        return "incomplete"
    if text in {"true", "1", "yes", "y", "checked"}:
        return "complete"
    if text in {"false", "0", "no", "n", "unchecked"}:
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
    rows: list[dict[str, Any]] = []
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
        if "total_score" not in df.columns:
            log.warn(
                f"{path.name} sheet {sheet} has no Total_score column; "
                "all instruments on this sheet will be treated as Total_score = 0"
            )
        excluded = 0
        for _, row in df.iterrows():
            if clean(row["included_in_summary"]) != "1":
                excluded += 1
                continue
            total_score_required = design_flag_enabled(row.get("total_score", ""))
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
                        "total_score_required": total_score_required,
                    }
                )
        log.info(f"{path.name} {sheet}: excluded {excluded} rows where Included_in_summary != 1")
    out = pd.DataFrame(rows)
    if out.empty:
        out = pd.DataFrame(
            columns=["arm", "session", "instrument", "instrument_source", "total_score_required"]
        )
    else:
        # If duplicated design rows disagree, Total_score = 1 takes priority.
        key_cols = ["arm", "session", "instrument", "instrument_source"]
        out["total_score_required"] = out["total_score_required"].fillna(False).astype(bool)
        out = (
            out.sort_values("total_score_required", ascending=False)
            .drop_duplicates(key_cols, keep="first")
            .reset_index(drop=True)
        )
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


def get_latest_file(input_dir: Path, pattern: str) -> Path:
    files = list(input_dir.glob(pattern))
    if not files:
        raise FileNotFoundError(f"No file found matching pattern: {pattern}")
    # Return the file with the most recent modification time
    return max(files, key=lambda f: f.stat().st_mtime)


def load_redcap(input_dir: Path, required_sessions: pd.DataFrame, log: ValidationLog) -> pd.DataFrame:
    frames = []
    event_map = build_event_map(required_sessions)
    for source, pattern in INPUT_FILES.items():
        file_path = get_latest_file(input_dir, pattern)
        log.info(f"Using {source} file: {file_path.name}")

        df = read_csv_strings(file_path, log)
        df.columns = [normalize_column(col) for col in df.columns]
        id_col = "record_id" if source == "participants" else "preescreen_id"
        if id_col not in df.columns:
            raise ValueError(f"{file_path} is missing required subject ID column {id_col}")
        if "redcap_event_name" not in df.columns:
            raise ValueError(f"{file_path} is missing redcap_event_name")
        df["source"] = source
        df["source_file"] = file_path
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


def numeric_score_present(value: Any) -> bool:
    """Return True for a finite numeric total, including zero."""
    text = clean(value).replace(",", "")
    if not text:
        return False
    try:
        return math.isfinite(float(text))
    except (TypeError, ValueError):
        return False


def resolve_total_score_field(
    columns: Iterable[Any],
    cfg: dict[str, str | list[str]],
    instrument: str,
) -> str:
    """Resolve the instrument-specific REDCap *_total column.

    An explicit cfg["total"] mapping is preferred when present. Otherwise, the
    function matches the *_total field to the instrument's completion-field stem
    (for example phq9_complete -> phq9_total), allowing underscore differences.
    """
    available = [normalize_column(col) for col in columns]
    total_fields = [field for field in available if field.endswith("_total")]
    if not total_fields:
        return ""

    explicit = cfg.get("total")
    explicit_fields = [explicit] if isinstance(explicit, str) else (explicit or [])
    for field in explicit_fields:
        normalized = normalize_column(field)
        if normalized in total_fields:
            return normalized

    complete = cfg.get("complete")
    complete_fields = [complete] if isinstance(complete, str) else (complete or [])
    stems: list[str] = []
    for field in complete_fields:
        stem = re.sub(r"_complete$", "", normalize_column(field))
        if stem:
            stems.append(stem)

    instrument_stem = normalize_column(instrument)
    if instrument_stem:
        stems.append(instrument_stem)

    # First try direct names such as gad7_total or lsas_sr_total.
    direct_candidates: list[str] = []
    for stem in stems:
        direct_candidates.append(f"{stem}_total")
        direct_candidates.append(f"{stem.replace('_', '')}_total")
    for candidate in direct_candidates:
        if candidate in total_fields:
            return candidate

    # Then allow formatting differences such as phq9 versus phq_9.
    compact_stems = {
        re.sub(r"[^0-9a-z]+", "", stem.lower())
        for stem in stems
        if stem
    }
    exact_matches = [
        field
        for field in total_fields
        if re.sub(r"[^0-9a-z]+", "", field[:-6].lower()) in compact_stems
    ]
    if len(exact_matches) == 1:
        return exact_matches[0]

    # A unique prefix match supports names such as qids_sr_total for QIDS.
    prefix_matches = []
    for field in total_fields:
        field_stem = re.sub(r"[^0-9a-z]+", "", field[:-6].lower())
        if any(
            field_stem.startswith(stem) or stem.startswith(field_stem)
            for stem in compact_stems
            if len(stem) >= 4
        ):
            prefix_matches.append(field)
    return prefix_matches[0] if len(prefix_matches) == 1 else ""


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
        total_score_required = design_flag_enabled(exp.get("total_score_required", ""))
        enforce_total_score = exp["session"] == "Baseline" and total_score_required
        total_score_field = (
            resolve_total_score_field(matches.columns, cfg, exp["instrument"])
            if enforce_total_score
            else ""
        )
        if enforce_total_score and not total_score_field:
            log.warn(
                "Could not resolve a unique *_total REDCap field for "
                f"{exp['arm']} Baseline {exp['instrument']}; its status cannot be complete"
            )

        for subject_id in subjects.loc[subjects["arm"] == exp["arm"], "subject_id"]:
            observed = by_subject.get(subject_id)
            date_field = clean(cfg.get("date", ""))
            date = clean(observed.get(date_field, "")) if observed is not None and date_field else ""
            status = status_from_fields(observed, cfg.get("complete"))

            total_score_value = (
                clean(observed.get(total_score_field, ""))
                if observed is not None and total_score_field
                else ""
            )
            total_score_numeric = numeric_score_present(total_score_value)

            if enforce_total_score:
                # A Baseline instrument marked Total_score = 1 is complete only
                # when its REDCap completion field is complete AND its own
                # instrument-specific *_total field contains a number.
                if status == "complete" and not total_score_numeric:
                    status = "incomplete"
            elif status == "missing" and date:
                # Preserve the existing date fallback for instruments that do
                # not require a numeric Baseline total score.
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
                    "total_score_required": enforce_total_score,
                    "total_score_field": total_score_field,
                    "total_score_value": total_score_value,
                    "total_score_numeric": total_score_numeric,
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


def build_session_field_dates(redcap: pd.DataFrame, field: str, source: str | None = None) -> dict[tuple[str, str], str]:
    data = redcap.copy()
    if source is not None:
        data = data[data["source"] == source].copy()
    observed: dict[tuple[str, str], str] = {}
    if field not in data.columns:
        return observed
    for _, row in data.iterrows():
        date = clean(row.get(field, ""))
        if date:
            observed.setdefault((row["subject_id"], row["session"]), date)
    return observed


def interval_weeks(baseline_date: str, current_date: str) -> float | None:
    baseline = pd.to_datetime(clean(baseline_date), errors="coerce")
    current = pd.to_datetime(clean(current_date), errors="coerce")
    if pd.isna(baseline) or pd.isna(current):
        return None
    return round((current - baseline).days / 7, 2)


def interval_days(start_date: str, current_date: str) -> int | None:
    start = pd.to_datetime(clean(start_date), errors="coerce")
    current = pd.to_datetime(clean(current_date), errors="coerce")
    if pd.isna(start) or pd.isna(current):
        return None
    return int((current - start).days)


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
                if record is not None:
                    # Define your exact REDCap column names here
                    acc_field = f"selfother_run{idx}_accuracy"
                    miss_field = f"selfother_run{idx}_missing_rate"
                    
                    acc_val = clean(record.get(acc_field, ""))
                    miss_val = clean(record.get(miss_field, ""))
                    
                    # If fields are entirely empty, mark missing
                    if not acc_val and not miss_val:
                        status = "missing"
                        qc_status = "missing"
                    else:
                        try:
                            # Parse values to floats for comparison
                            acc_float = float(acc_val)
                            miss_float = float(miss_val)
                            
                            status = "complete" # Data exists
                            
                            # Threshold Logic: > 90% Acc, < 10% Missing
                            if acc_float > 90.0 and miss_float < 10.0:
                                qc_status = "pass"
                            else:
                                qc_status = "fail"
                                
                        except ValueError:
                            # Catch cases where text like "NA" or "ND" was entered instead of numbers
                            status = "missing" if acc_val.lower() in {"na", "n/a", "nd"} else "review_required"
                            qc_status = "missing" if status == "missing" else "review_required"
                else:
                    acc_val, miss_val = "", ""
                    status = "missing"
                    qc_status = "missing"

                # Append to the long-format dataframe list
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
                        # Store both metrics in the value column so they appear in outputs for debugging
                        "value": f"Acc: {acc_val} | Miss: {miss_val}" if (acc_val or miss_val) else "", 
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
            #if kind == "anat":
            #    scan_or_run = "anat_T1w"
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


def subject_list(values: Iterable[Any]) -> str:
    subjects = sorted({clean(value) for value in values if clean(value)})
    return "; ".join(subjects)


def subjects_with_status(df: pd.DataFrame, statuses: set[str], status_col: str = "status") -> str:
    if df.empty:
        return ""
    mask = df[status_col].fillna("").map(clean).isin(statuses)
    return subject_list(df.loc[mask, "subject_id"])


def subjects_from_series(series: pd.Series, statuses: set[str]) -> str:
    if series.empty:
        return ""
    return subject_list(series[series.isin(statuses)].index)
    

def status_from_fields(row: pd.Series | None, fields: str | list[str] | None) -> str:
    if row is None or not fields:
        return "missing"
        
    field_list = [fields] if isinstance(fields, str) else fields
    
    statuses = []
    for field in field_list:
        if field in row.index:
            # CRITICAL FIX: We MUST pass 'field' as the second argument so the exception triggers
            statuses.append(status_from_complete_code(row.get(field, ""), field))
            
    return combined_status(statuses)


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

from datetime import datetime, timedelta
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
    clinician_data = redcap[redcap["source"] == "clinician"]
    # status_map = {"1": "Active", "2": "Withdrawn", "3": "Ineligible", "4": "Completed", "5":"Other"}
    # for subject_id in arm_subjects:
    #     subject_record = clinician_data[clinician_data["subject_id"] == subject_id]
    #     date_cols = [c for c in qn.columns if 'date' in c]
    #     qn_dates = qn.copy()
    #     for col in date_cols:
    #         qn_dates[col] = pd.to_datetime(qn_dates[col], errors='coerce')
    #     last_activity = qn_dates.groupby('subject_id')[date_cols].max().max(axis=1)
    #     current_time = datetime.now()
    #     three_months = timedelta(days=90)
    #     raw_status = ""
    #     if "participant_status" in subject_record.columns:
    #         raw_status = str(subject_record["participant_status"].iloc[0])
    #     if raw_status in status_map:
    #         dropout_status = status_map[raw_status]
    #     subject_status_record = subject_record[subject_record['redcap_event_name'].str.contains('subject_status', case=False, na=False)]
    #     if subject_status_record.empty:
    #         subject_status_complete_val = False
    #         elig_val = None
    #     else:
    #         # Safely get the status and group from the first screening record
    #         subject_status_complete_val = str(subject_status_record['subject_status_complete'].iloc[0]).strip() == '2'
    #         elig_val = subject_status_record['elig_group'].iloc[0]
    #     current_event = subject_record['redcap_event_name'] if not subject_record.empty else ""
    #     if not subject_status_complete_val:
    #         dropout_status = 'subject_status_not_complete'
    #     elif str(elig_val) not in ["1", "2", "4"]:
    #         dropout_status = 'subject_status_not_Eligible'
    #     #elif (len(current_event) == 19 and arm == 'arm1') or (len(current_event) == 4 and arm == 'arm2'):
    #     #    dropout_status = 'Completed'
    #     else:
    #         # Fallback logic: check for last activity
    #         last_date = last_activity.get(subject_id)
    #         if pd.notna(last_date) and (current_time - last_date) <= three_months:
    #             dropout_status = "Active"
    #         else:
    #             dropout_status = "need to review"
    #     row: dict[str, Any] = {"subject_id": subject_id, "arm": arm, "dropout_status": dropout_status}
    #     for session in arm_sessions["session"]:
    #         s_label = excel_safe_label(session)
    #         qn_s = qn[(qn["subject_id"] == subject_id) & (qn["session"] == session)].sort_values("instrument")
    #         for _, item in qn_s.iterrows():
    #             q_label = item_label(item["instrument"])
    #             row[f"ses-{s_label}_qn-{q_label}_date"] = item["date"]
    #             row[f"ses-{s_label}_qn-{q_label}_status"] = item["status"]
    #         ant_s = beh[
    #             (beh["subject_id"] == subject_id)
    #             & (beh["session"] == session)
    #             & (beh["domain"] == "ANT")
    #         ]
    #         if not ant_s.empty:
    #             row[f"ses-{s_label}_beh-ANT_status"] = ant_s.iloc[0]["status"]
    #         self_s = beh[
    #             (beh["subject_id"] == subject_id)
    #             & (beh["session"] == session)
    #             & (beh["domain"] == "selfOthers")
    #         ]
    #         npractice = self_s[self_s["item"] == "Npractice"]
    #         if not npractice.empty:
    #             row[f"ses-{s_label}_beh-selfOthers_Npractice"] = npractice.iloc[0]["value"]
    #         for run in ("run1", "run2"):
    #             run_s = self_s[self_s["item"] == run]
    #             if not run_s.empty:
    #                 run_row = run_s.iloc[0]
    #                 row[f"ses-{s_label}_beh-selfOthers_{run}_status"] = run_row["status"]
    #                 row[f"ses-{s_label}_beh-selfOthers_{run}_acc"] = ""
    #                 row[f"ses-{s_label}_beh-selfOthers_{run}_missingRate"] = ""
    #         mri_s = mri[(mri["subject_id"] == subject_id) & (mri["session"] == session)].sort_values("scan_or_run")
    #         for _, item in mri_s.iterrows():
    #             row[f"ses-{s_label}_mri-{item_label(item['scan_or_run'])}_qc_status"] = item["qc_status"]
    #         summary = session_summary_values(subject_id, session, qn, beh, mri)
    #         for key, value in summary.items():
    #             row[f"ses-{s_label}_{key}"] = bool_text(value) if isinstance(value, bool) else value
    #         if session in INTERVAL_SUMMARY_SESSIONS:
    #             reference_session = "Baseline" if session == "Repeat Baseline" else "Repeat Baseline"
    #             weeks = interval_weeks(
    #                 observed_dates.get((subject_id, reference_session), ""),
    #                 observed_dates.get((subject_id, session), ""),
    #             )
    #             row[f"ses-{s_label}_intervalFromRepeatBaseline_weeks"] = "" if weeks is None else weeks
    #             row[f"ses-{s_label}_interval_valid"] = "" if weeks is None else bool_text(weeks >= 0)
    #             standard_weeks = STANDARD_INTERVAL_WEEKS.get(session)
    #             if weeks is not None and standard_weeks is not None:
    #                 row[f"ses-{s_label}_intervalDeviationFromStandard_weeks"] = round(weeks - standard_weeks, 2)
    #             else:
    #                 row[f"ses-{s_label}_intervalDeviationFromStandard_weeks"] = ""
    #     subject_redcap = redcap[redcap["subject_id"] == subject_id]
    #     row["total_ASAP_count"] = int(subject_redcap["session"].map(lambda x: clean(x).upper().startswith("ASAP")).sum())
    #     session_flags = [
    #         row.get(f"ses-{excel_safe_label(session)}_session_completed") == "True"
    #         for session in arm_sessions["session"]
    #     ]
    #     missing_sessions = [
    #         session
    #         for session in arm_sessions["session"]
    #         if row.get(f"ses-{excel_safe_label(session)}_session_completed") != "True"
    #     ]
    #     row["complete_all_expected_experiment_sessions"] = bool_text(all(session_flags))
    #     row["Nof_missing_expected_experiment_sessions"] = len(missing_sessions)
    #     row["missing_expected_experiment_sessions"] = "; ".join(missing_sessions)
    #     row["complete_all_experiment_sessions"] = row["complete_all_expected_experiment_sessions"]
    #     qn_subject = qn[qn["subject_id"] == subject_id]
    #     ant_subject = beh[(beh["subject_id"] == subject_id) & (beh["domain"] == "ANT")]
    #     self_subject = beh[(beh["subject_id"] == subject_id) & (beh["domain"] == "selfOthers") & (beh["item"].str.startswith("run"))]
    #     mri_subject = mri[mri["subject_id"] == subject_id]
    #     row["complete_all_instrument"] = bool_text(not qn_subject.empty and (qn_subject["status"] == "complete").all())
    #     row["complete_all_ANT"] = bool_text(not ant_subject.empty and (ant_subject["status"] == "complete").all())
    #     row["all_MRI_QC_passed"] = bool_text(not mri_subject.empty and (mri_subject["qc_status"] == "pass").all())
    #     row["all_selfOther_QC_passed"] = bool_text(not self_subject.empty and (self_subject["qc_status"] == "pass").all())
    #     row["subject_QC_pass"] = bool_text(
    #         row["complete_all_expected_experiment_sessions"] == "True"
    #         and row["complete_all_instrument"] == "True"
    #         and row["complete_all_ANT"] == "True"
    #         and row["all_MRI_QC_passed"] == "True"
    #         and row["all_selfOther_QC_passed"] == "True"
    #     )
    #     rows.append(row)
    # base_cols = ["subject_id", "arm", "dropout_status"]
    status_map = {
        1: "Active",
        2: "Withdrawn",
        3: "Ineligible",
        4: "Completed",
        5: "Other",
    }

    eligible_groups = {1, 2, 4}

    # Calculate questionnaire activity dates once, outside the subject loop
    date_cols = [c for c in qn.columns if "date" in c.lower()]

    qn_dates = qn.copy()

    for col in date_cols:
        qn_dates[col] = pd.to_datetime(qn_dates[col], errors="coerce")

    if date_cols:
        last_activity = (
            qn_dates
            .groupby("subject_id")[date_cols]
            .max()
            .max(axis=1)
        )
    else:
        last_activity = pd.Series(dtype="datetime64[ns]")

    current_time = pd.Timestamp.now()
    three_months = pd.Timedelta(days=90)


    for subject_id in arm_subjects:

        subject_record = clinician_data[
            clinician_data["subject_id"] == subject_id
        ]

        # Restrict status values to the subject_status REDCap event
        subject_status_record = subject_record[
            subject_record["redcap_event_name"]
            .str.contains("subject_status", case=False, na=False)
        ]

        if subject_status_record.empty:
            participant_status_val = None
            subject_status_complete_val = False
            elig_val = None

        else:
            # Get the last nonmissing participant_status value
            participant_status_values = pd.to_numeric(
                subject_status_record["participant_status"],
                errors="coerce",
            ).dropna()

            participant_status_val = (
                int(participant_status_values.iloc[-1])
                if not participant_status_values.empty
                else None
            )

            # Get the last nonmissing subject_status_complete value
            complete_values = pd.to_numeric(
                subject_status_record["subject_status_complete"],
                errors="coerce",
            ).dropna()

            subject_status_complete_val = (
                not complete_values.empty
                and int(complete_values.iloc[-1]) == 2
            )

            # Get the last nonmissing eligibility-group value
            elig_values = pd.to_numeric(
                subject_status_record["elig_group"],
                errors="coerce",
            ).dropna()

            elig_val = (
                int(elig_values.iloc[-1])
                if not elig_values.empty
                else None
            )

        # Determine participant/dropout status
        if not subject_status_complete_val:
            dropout_status = "subject_status_not_complete"

        # A recorded participant_status takes priority
        elif participant_status_val in status_map:
            dropout_status = status_map[participant_status_val]

        # Use eligibility only when participant_status is missing or invalid
        elif elig_val is None:
            dropout_status = "eligibility_missing"

        elif elig_val not in eligible_groups:
            dropout_status = "subject_status_not_Eligible"

        else:
            # Fallback: determine status from most recent activity
            last_date = last_activity.get(subject_id)

            if (
                pd.notna(last_date)
                and current_time - last_date <= three_months
            ):
                dropout_status = "Active"
            else:
                dropout_status = "need to review"

        row: dict[str, Any] = {
            "subject_id": subject_id,
            "arm": arm,
            "dropout_status": dropout_status,
        }

        for session in arm_sessions["session"]:
            s_label = excel_safe_label(session)

            qn_s = qn[
                (qn["subject_id"] == subject_id)
                & (qn["session"] == session)
            ].sort_values("instrument")

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
                row[f"ses-{s_label}_beh-selfOthers_Npractice"] = (
                    npractice.iloc[0]["value"]
                )

            for run in ("run1", "run2"):
                run_s = self_s[self_s["item"] == run]

                if not run_s.empty:
                    run_row = run_s.iloc[0]
                    row[f"ses-{s_label}_beh-selfOthers_{run}_status"] = (
                        run_row["status"]
                    )
                    row[f"ses-{s_label}_beh-selfOthers_{run}_acc"] = ""
                    row[f"ses-{s_label}_beh-selfOthers_{run}_missingRate"] = ""

            mri_s = mri[
                (mri["subject_id"] == subject_id)
                & (mri["session"] == session)
            ].sort_values("scan_or_run")

            for _, item in mri_s.iterrows():
                row[
                    f"ses-{s_label}_mri-"
                    f"{item_label(item['scan_or_run'])}_qc_status"
                ] = item["qc_status"]

            summary = session_summary_values(
                subject_id,
                session,
                qn,
                beh,
                mri,
            )

            for key, value in summary.items():
                row[f"ses-{s_label}_{key}"] = (
                    bool_text(value)
                    if isinstance(value, bool)
                    else value
                )

            if session in INTERVAL_SUMMARY_SESSIONS:
                reference_session = (
                    "Baseline"
                    if session == "Repeat Baseline"
                    else "Repeat Baseline"
                )

                weeks = interval_weeks(
                    observed_dates.get(
                        (subject_id, reference_session),
                        "",
                    ),
                    observed_dates.get(
                        (subject_id, session),
                        "",
                    ),
                )

                row[f"ses-{s_label}_intervalFromRepeatBaseline_weeks"] = (
                    ""
                    if weeks is None
                    else weeks
                )

                row[f"ses-{s_label}_interval_valid"] = (
                    ""
                    if weeks is None
                    else bool_text(weeks >= 0)
                )

                standard_weeks = STANDARD_INTERVAL_WEEKS.get(session)

                if weeks is not None and standard_weeks is not None:
                    row[
                        f"ses-{s_label}_"
                        "intervalDeviationFromStandard_weeks"
                    ] = round(weeks - standard_weeks, 2)
                else:
                    row[
                        f"ses-{s_label}_"
                        "intervalDeviationFromStandard_weeks"
                    ] = ""

        subject_redcap = redcap[
            redcap["subject_id"] == subject_id
        ]

        row["total_ASAP_count"] = int(
            subject_redcap["session"]
            .map(lambda x: clean(x).upper().startswith("ASAP"))
            .sum()
        )

        session_flags = [
            row.get(
                f"ses-{excel_safe_label(session)}_session_completed"
            ) == "True"
            for session in arm_sessions["session"]
        ]

        missing_sessions = [
            session
            for session in arm_sessions["session"]
            if row.get(
                f"ses-{excel_safe_label(session)}_session_completed"
            ) != "True"
        ]

        row["complete_all_expected_experiment_sessions"] = bool_text(
            all(session_flags)
        )

        row["Nof_missing_expected_experiment_sessions"] = len(
            missing_sessions
        )

        row["missing_expected_experiment_sessions"] = "; ".join(
            missing_sessions
        )

        row["complete_all_experiment_sessions"] = (
            row["complete_all_expected_experiment_sessions"]
        )

        qn_subject = qn[
            qn["subject_id"] == subject_id
        ]

        ant_subject = beh[
            (beh["subject_id"] == subject_id)
            & (beh["domain"] == "ANT")
        ]

        self_subject = beh[
            (beh["subject_id"] == subject_id)
            & (beh["domain"] == "selfOthers")
            & (beh["item"].str.startswith("run", na=False))
        ]

        mri_subject = mri[
            mri["subject_id"] == subject_id
        ]

        row["complete_all_instrument"] = bool_text(
            not qn_subject.empty
            and (qn_subject["status"] == "complete").all()
        )

        row["complete_all_ANT"] = bool_text(
            not ant_subject.empty
            and (ant_subject["status"] == "complete").all()
        )

        row["all_MRI_QC_passed"] = bool_text(
            not mri_subject.empty
            and (mri_subject["qc_status"] == "pass").all()
        )

        row["all_selfOther_QC_passed"] = bool_text(
            not self_subject.empty
            and (self_subject["qc_status"] == "pass").all()
        )

        row["subject_QC_pass"] = bool_text(
            row["complete_all_expected_experiment_sessions"] == "True"
            and row["complete_all_instrument"] == "True"
            and row["complete_all_ANT"] == "True"
            and row["all_MRI_QC_passed"] == "True"
            and row["all_selfOther_QC_passed"] == "True"
        )

        rows.append(row)


    base_cols = [
        "subject_id",
        "arm",
        "dropout_status",
    ]
##############################
#replace before
################################

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
                "missing_subjects": subjects_with_status(group, MISSING_STATUSES),
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
        "missing_subjects",
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
                "missing_subjects": subjects_from_series(per_subject, MISSING_STATUSES),
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
        "missing_subjects",
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
                "missing_subjects": subjects_with_status(group, MISSING_STATUSES),
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
        "missing_subjects",
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
                "missing_subjects": subjects_from_series(per_subject, MISSING_STATUSES),
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
                "missing_subjects": subjects_from_series(per_subject, {"missing"}),
                "qc_fail_subjects": subjects_from_series(per_subject, {"fail"}),
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
        "missing_subjects",
        "qc_fail_subjects",
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
            "missing_subjects": subjects_from_series(per_subject, {"missing"}),
            "qc_fail_subjects": subjects_from_series(per_subject, {"fail"}),
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
        "missing_subjects",
        "qc_fail_subjects",
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
        if session not in INTERVAL_SUMMARY_SESSIONS:
            continue
        label = excel_safe_label(session)
        completed_col = f"ses-{label}_session_completed"
        interval_col = f"ses-{label}_intervalFromRepeatBaseline_weeks"
        valid_col = f"ses-{label}_interval_valid"
        if completed_col not in subject_wise.columns:
            continue
        intervals = pd.to_numeric(subject_wise.get(interval_col, pd.Series(dtype=str)), errors="coerce")
        completed = intervals.notna()
        valid = completed & (subject_wise.get(valid_col, pd.Series([""] * len(subject_wise))).map(clean) == "True")
        standard_weeks = STANDARD_INTERVAL_WEEKS.get(session, "")
        deviations = intervals - standard_weeks if standard_weeks != "" else pd.Series([pd.NA] * len(subject_wise))
        rows.append(
            {
                "arm": arm,
                "session": session,
                "reference_session": "Baseline" if session == "Repeat Baseline" else "Repeat Baseline",
                "standard_interval_weeks": standard_weeks,
                "completed_NofSubjects": int(completed.sum()),
                "valid_interval_NofSubjects": int(valid.sum()),
                "invalid_interval_NofSubjects": int((completed & ~valid).sum()),
                "invalid_or_missing_interval_subjects": subject_list(subject_wise.loc[completed & ~valid, "subject_id"]),
                "mean_intervalFromRepeatBaseline_weeks": round(float(intervals[valid].mean()), 2) if valid.any() else "",
                "sd_intervalFromRepeatBaseline_weeks": round(float(intervals[valid].std()), 2) if valid.sum() > 1 else "",
                "min_intervalFromRepeatBaseline_weeks": round(float(intervals[valid].min()), 2) if valid.any() else "",
                "max_intervalFromRepeatBaseline_weeks": round(float(intervals[valid].max()), 2) if valid.any() else "",
                "mean_deviationFromStandard_weeks": round(float(deviations[valid].mean()), 2) if standard_weeks != "" and valid.any() else "",
            }
        )
    columns = [
        "arm",
        "session",
        "reference_session",
        "standard_interval_weeks",
        "completed_NofSubjects",
        "valid_interval_NofSubjects",
        "invalid_interval_NofSubjects",
        "invalid_or_missing_interval_subjects",
        "mean_intervalFromRepeatBaseline_weeks",
        "sd_intervalFromRepeatBaseline_weeks",
        "min_intervalFromRepeatBaseline_weeks",
        "max_intervalFromRepeatBaseline_weeks",
        "mean_deviationFromStandard_weeks",
    ]
    return pd.DataFrame(rows, columns=columns)


def table_te_scan_interval_summary(arm: str, subject_wise: pd.DataFrame, redcap: pd.DataFrame) -> pd.DataFrame:
    rows = []
    te_dates = build_session_field_dates(redcap, "peas_date_of_visit", source="clinician")
    scan_dates = build_session_field_dates(redcap, "scan_date", source="clinician")
    arm_subject_wise = subject_wise[subject_wise["arm"] == arm]
    for te_session, scan_sessions in TE_SCAN_PAIRS:
        te_label = excel_safe_label(te_session)
        te_completed_col = f"ses-{te_label}_session_completed"
        if te_completed_col not in arm_subject_wise.columns:
            continue
        intervals = []
        scan_session_values = []
        te_date_values = []
        for _, row in arm_subject_wise.iterrows():
            subject_id = row["subject_id"]
            scan_session = next(
                (candidate for candidate in scan_sessions if scan_dates.get((subject_id, candidate), "")),
                "",
            )
            te_date = te_dates.get((subject_id, te_session), "")
            days = interval_days(
                te_date,
                scan_dates.get((subject_id, scan_session), "") if scan_session else "",
            )
            intervals.append(days)
            scan_session_values.append(scan_session)
            te_date_values.append(te_date)
        interval_series = pd.to_numeric(pd.Series(intervals, index=arm_subject_wise.index), errors="coerce")
        scan_session_series = pd.Series(scan_session_values, index=arm_subject_wise.index)
        te_date_series = pd.Series(te_date_values, index=arm_subject_wise.index)
        te_present = te_date_series.map(clean) != ""
        scan_present = scan_session_series.map(clean) != ""
        completed = te_present & scan_present
        valid = completed & interval_series.notna() & (interval_series >= 0)
        deviations = interval_series - STANDARD_TE_SCAN_INTERVAL_DAYS
        rows.append(
            {
                "arm": arm,
                "te_session": te_session,
                "scan_session": " or ".join(scan_sessions),
                "standard_interval_days": STANDARD_TE_SCAN_INTERVAL_DAYS,
                "completed_pair_NofSubjects": int(completed.sum()),
                "valid_interval_NofSubjects": int(valid.sum()),
                "invalid_interval_NofSubjects": int((completed & ~valid).sum()),
                "invalid_or_missing_interval_subjects": subject_list(arm_subject_wise.loc[completed & ~valid, "subject_id"]),
                "mean_interval_days": round(float(interval_series[valid].mean()), 2) if valid.any() else "",
                "sd_interval_days": round(float(interval_series[valid].std()), 2) if valid.sum() > 1 else "",
                "min_interval_days": round(float(interval_series[valid].min()), 2) if valid.any() else "",
                "max_interval_days": round(float(interval_series[valid].max()), 2) if valid.any() else "",
                "mean_deviationFromStandard_days": round(float(deviations[valid].mean()), 2) if valid.any() else "",
            }
        )
    columns = [
        "arm",
        "te_session",
        "scan_session",
        "standard_interval_days",
        "completed_pair_NofSubjects",
        "valid_interval_NofSubjects",
        "invalid_interval_NofSubjects",
        "invalid_or_missing_interval_subjects",
        "mean_interval_days",
        "sd_interval_days",
        "min_interval_days",
        "max_interval_days",
        "mean_deviationFromStandard_days",
    ]
    return pd.DataFrame(rows, columns=columns)


def table_readiness(arm: str, subject_wise: pd.DataFrame) -> pd.DataFrame:
    total = len(subject_wise)
    count_true = lambda col: int((subject_wise[col].map(clean) == "True").sum()) if col in subject_wise else 0
    missing_true = lambda col: subject_list(subject_wise.loc[subject_wise[col].map(clean) != "True", "subject_id"]) if col in subject_wise else ""
    qc_pass = count_true("subject_QC_pass")
    return pd.DataFrame(
        [
            {
                "arm": arm,
                "total_NofSubjects": total,
                "withdrawn_or_dropout_NofSubjects": 0,
                "withdrawn_or_dropout_rate": 0,
                "complete_all_experiment_sessions_NofSubjects": count_true("complete_all_experiment_sessions"),
                "complete_all_experiment_sessions_missing_subjects": missing_true("complete_all_experiment_sessions"),
                "complete_all_experiment_sessions_rate": rate(count_true("complete_all_experiment_sessions"), total),
                "complete_all_instrument_NofSubjects": count_true("complete_all_instrument"),
                "complete_all_instrument_missing_subjects": missing_true("complete_all_instrument"),
                "complete_all_instrument_rate": rate(count_true("complete_all_instrument"), total),
                "complete_all_ANT_NofSubjects": count_true("complete_all_ANT"),
                "complete_all_ANT_missing_subjects": missing_true("complete_all_ANT"),
                "complete_all_ANT_rate": rate(count_true("complete_all_ANT"), total),
                "all_MRI_QC_passed_NofSubjects": count_true("all_MRI_QC_passed"),
                "all_MRI_QC_passed_missing_subjects": missing_true("all_MRI_QC_passed"),
                "all_MRI_QC_passed_rate": rate(count_true("all_MRI_QC_passed"), total),
                "all_selfOther_QC_passed_NofSubjects": count_true("all_selfOther_QC_passed"),
                "all_selfOther_QC_passed_missing_subjects": missing_true("all_selfOther_QC_passed"),
                "all_selfOther_QC_passed_rate": rate(count_true("all_selfOther_QC_passed"), total),
                "QC_pass_NofSubjects_rate": qc_pass,
                "QC_pass_missing_subjects": missing_true("subject_QC_pass"),
                "QC_passrate": rate(qc_pass, total),
            }
        ]
    )


def table_asap_summary(arm: str, subject_wise: pd.DataFrame) -> pd.DataFrame:
    total = len(subject_wise)
    asap_counts = pd.to_numeric(
        subject_wise.get("total_ASAP_count", pd.Series([0] * total)),
        errors="coerce",
    ).fillna(0).astype(int)
    row: dict[str, Any] = {"arm": arm, "total_NofSubjects": total}
    for count in range(6):
        row[f"{count}_ASAP_NofSubjects"] = int((asap_counts == count).sum())
    return pd.DataFrame(
        [row],
        columns=[
            "arm",
            "total_NofSubjects",
            "0_ASAP_NofSubjects",
            "1_ASAP_NofSubjects",
            "2_ASAP_NofSubjects",
            "3_ASAP_NofSubjects",
            "4_ASAP_NofSubjects",
            "5_ASAP_NofSubjects",
        ],
    )


# def build_group_tables(
#     arm: str,
#     qn: pd.DataFrame,
#     beh: pd.DataFrame,
#     mri: pd.DataFrame,
#     subject_wise: pd.DataFrame,
#     arm_sessions: list[str],
# ) -> list[tuple[str, pd.DataFrame]]:
#     self_qc = beh[(beh["domain"] == "selfOthers") & (beh["item"].str.startswith("run"))].copy()
#     return [
#         ("Table 1. Session-wise Summary", table_session_instrument_summary(arm, qn)),
#         ("Table 2. Instrument-wise Summary", table_instrument_summary(arm, qn)),
#         ("Table 3-1. ANT task complete rate by session", table_ant_by_session(arm, beh)),
#         ("Table 3-2. ANT task complete rate by subject", table_ant_by_subject(arm, beh)),
#         ("Table 4-1. SelfOthers QC pass rate by session", qc_table_by_session(arm, self_qc, "selfOthers")),
#         ("Table 4-2. SelfOthers QC pass rate by subject", qc_table_by_subject(arm, self_qc)),
#         ("Table 5. MRI QC pass rate by session", qc_table_by_session(arm, mri, "MRI")),
#         ("Table 5-2. MRI QC pass rate by subject", qc_table_by_subject(arm, mri, "scan_or_run")),
#         ("Table 6. Session interval summary from Repeat Baseline", table_interval_summary(arm, subject_wise, arm_sessions)),
#         ("Table 7. ASAP summary", table_asap_summary(arm, subject_wise)),
#         ("Table 8. Participant-level QA readiness summary", table_readiness(arm, subject_wise)),
#     ]
def build_group_tables(
    arm: str,
    qn: pd.DataFrame,
    beh: pd.DataFrame,
    mri: pd.DataFrame,
    subject_wise: pd.DataFrame,
    arm_sessions: list[str],
    redcap: pd.DataFrame,
) -> list[tuple[str, pd.DataFrame]]:
    
    # 1. Extract valid subjects based on dropout_status
    valid_statuses = {"Active", "Completed"}
    valid_subjects = set(
        subject_wise[subject_wise["dropout_status"].isin(valid_statuses)]["subject_id"]
    )

    # 2. Filter all long-format DataFrames to include only valid subjects
    qn_filtered = qn[qn["subject_id"].isin(valid_subjects)].copy()
    beh_filtered = beh[beh["subject_id"].isin(valid_subjects)].copy()
    mri_filtered = mri[mri["subject_id"].isin(valid_subjects)].copy()
    subject_wise_filtered = subject_wise[subject_wise["subject_id"].isin(valid_subjects)].copy()

    # 3. Create selfOthers subset from the filtered behavioral data
    self_qc = beh_filtered[
        (beh_filtered["domain"] == "selfOthers") & 
        (beh_filtered["item"].str.startswith("run"))
    ].copy()

    # 4. Generate tables using the filtered DataFrames
    return [
        ("Table 1. Session-wise Summary", table_session_instrument_summary(arm, qn_filtered)),
        ("Table 2. Instrument-wise Summary", table_instrument_summary(arm, qn_filtered)),
        ("Table 3-1. ANT task complete rate by session", table_ant_by_session(arm, beh_filtered)),
        ("Table 3-2. ANT task complete rate by subject", table_ant_by_subject(arm, beh_filtered)),
        ("Table 4-1. SelfOthers QC pass rate by session", qc_table_by_session(arm, self_qc, "selfOthers")),
        ("Table 4-2. SelfOthers QC pass rate by subject", qc_table_by_subject(arm, self_qc)),
        ("Table 5. MRI QC pass rate by session", qc_table_by_session(arm, mri_filtered, "MRI")),
        ("Table 5-2. MRI QC pass rate by subject", qc_table_by_subject(arm, mri_filtered, "scan_or_run")),
        ("Table 6. Session interval summary from Repeat Baseline", table_interval_summary(arm, subject_wise_filtered, arm_sessions)),
        ("Table 6-2. TE- scan Session interval summary", table_te_scan_interval_summary(arm, subject_wise, redcap)),
        ("Table 7. ASAP summary", table_asap_summary(arm, subject_wise_filtered)),
        ("Table 8. Participant-level QA readiness summary", table_readiness(arm, subject_wise_filtered)),
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
    active_arms = configured_arms(QA_CONFIG)
    if not active_arms:
        raise ValueError("QA_CONFIG active_arms did not include any valid arms")
    log.info(f"Configured active arms: {', '.join(active_arms)}")
    required_sessions = load_required_sessions(input_dir, log)
    required_sessions = apply_required_session_config(required_sessions, QA_CONFIG, log)
    log.count("required_session_rows_after_config", len(required_sessions))
    expected = load_expected_instruments(input_dir, required_sessions, log)
    expected = apply_expected_instrument_config(expected, required_sessions, QA_CONFIG, log)
    log.count("included_expected_instrument_rows_after_config", len(expected))
    redcap = load_redcap(input_dir, required_sessions, log)
    mri_subjects = read_mri_subjects(input_dir, log)
    subjects = build_subject_arm_map(redcap, mri_subjects, log)
    subjects = apply_subject_config(subjects, QA_CONFIG, log)
    log.count("subjects_after_config", len(subjects))
    qn = build_questionnaire_long(redcap, subjects, expected, log)
    beh = build_behavioral_long(redcap, subjects, required_sessions)
    mri = load_mri_long(input_dir, subjects, required_sessions, log)
    log.count("behavioral_rows_built", len(beh))
    for arm in active_arms:
        arm_sessions = (
            required_sessions[
                (required_sessions["arm"] == arm) & (required_sessions["session"].map(final_session_allowed))
            ]
            .sort_values("order")["session"]
            .tolist()
        )
        subject_wise = build_subject_wise(arm, subjects, required_sessions, qn, beh, mri, redcap)
        tables = build_group_tables(arm, qn, beh, mri, subject_wise, arm_sessions, redcap)
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
