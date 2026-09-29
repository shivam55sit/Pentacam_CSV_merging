"""
Pentacam CSV Pipeline — Streamlit App
======================================

End-to-end pipeline that:
  1. Scans a folder tree for school date-folders and filters only CSV files.
  2. Standardises school names and groups date-folders under a single school.
  3. Auto-detects delimiters (;  ,  tab) and reads CSVs correctly.
  4. Extracts specific columns per the Pentacam_PARAMETER.xlsx mapping.
  5. Merges per-patient across CSV file types (on Pat-ID + DOB + Eye).
  6. Outputs downloadable Excel files (per-school + combined + validation).

Run with:
    streamlit run app.py
"""

import io
import os
import re
import tempfile
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st

# ═══════════════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

# Candidate delimiters (tried in this order)
CSV_DELIMITER_CANDIDATES = [";", ",", "\t"]

# Identifier columns present in every Pentacam CSV export
IDENTIFIER_COLUMNS = [
    "Last Name:",
    "First Name:",
    "Pat-ID:",
    "D.o.Birth:",
    "Exam Eye:",
]

# Columns used as the actual merge key
PRIMARY_MATCH_COLUMNS = ["Pat-ID:", "D.o.Birth:", "Exam Eye:"]

# Regex that matches folder names like  SCHOOL NAME_DD-MM-YYYY
# Also handles optional spaces, underscores, or commas before the date
# Supports dates with -, _, ., or space as separators, and 2 or 4 digit years
# Also allows optional trailing text after the date (e.g., 'topo 1', '(topo-1)', '-1')
DATE_FOLDER_PATTERN = re.compile(
    r"^(?P<school>.+?)[_ \-,]*(?P<date>\d{1,2}[-_ .]\d{1,2}[-_ .]\d{2,4}).*$",
    re.IGNORECASE
)

# Regex to strip trailing date from a filename stem
# e.g. "BADisplay-LOAD_01-09-2025" → "BADisplay-LOAD"
DATE_SUFFIX_RE = re.compile(r"[_\-\s]*\d{1,2}[-_.]\d{1,2}[-_.]\d{2,4}\s*$")

# Default mapping (used when no Pentacam_PARAMETER.xlsx is uploaded)
DEFAULT_COLUMN_MAP = {
    "INDEX-LOAD.CSV": [
        "K Max (Front):",
        "K Max X (Front):",
        "K Max Y (Front):",
    ],
    "CorneoScleral-LOAD.CSV": [
        "HWTW:",
    ],
    "CHAMBER-LOAD.CSV": [
        "C.Volume:",
        "C.Angle:",
        "Pupil:",
    ],
    "PACHY-LOAD.CSV": [
        "AC Depth",
    ],
    "BADisplay-LOAD.CSV": [
        "Pachy Prog Index Min.:",
        "Pachy Prog Index Max.:",
        "Pachy Prog Index Avg.:",
        "ART Max.:",
        "ART Min.:",
        "ART Avg.:",
        "Dist. Apex-Thin.Loc. [mm]:",
        "Status:",
        "BAD Df:",
        "BAD Db:",
        "BAD Dp:",
        "BAD Dt:",
        "BAD Dam:",
        "BAD D:",
    ],
}

# Non-CSV extensions to filter out
NON_CSV_EXTENSIONS = {
    ".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".tif",
    ".pdf", ".doc", ".docx", ".ppt", ".pptx",
    ".xls", ".xlsx", ".zip", ".rar", ".7z",
    ".mp4", ".avi", ".mov", ".wmv",
    ".exe", ".dll", ".bat", ".sh",
}


# ═══════════════════════════════════════════════════════════════════════════════
#  HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════


def normalize_school_name(name: str) -> str:
    """
    Normalise a school name for grouping.
    - Strip leading/trailing whitespace and underscores
    - Collapse multiple spaces into one
    - Convert to UPPER CASE
    """
    name = name.strip().strip("_").strip()
    name = re.sub(r"\s+", " ", name)
    return name.upper()


def sniff_delimiter(first_line: str) -> str:
    """Pick whichever candidate delimiter appears most often in the header."""
    counts = {d: first_line.count(d) for d in CSV_DELIMITER_CANDIDATES}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else CSV_DELIMITER_CANDIDATES[0]


def read_csv_smart(path: Path) -> pd.DataFrame:
    """
    Read a Pentacam CSV with auto-detected delimiter and encoding.

    Pentacam exports are typically semicolon-delimited.  Some header names
    contain literal commas (e.g. "Chord Mu (polar vector, rel. to pupil
    center) [mm]:"), which can fool the sniffer into picking comma and
    mangling the data.  So semicolon is tried first explicitly.
    """
    last_error = None
    for encoding in ("utf-8-sig", "latin1", "cp1252", "utf-16"):
        try:
            # 1. Try semicolon explicitly first
            df = pd.read_csv(
                path, sep=";", engine="python", encoding=encoding, dtype=str
            )
            if df.shape[1] <= 1:
                # 2. Fall back to comma
                df_comma = pd.read_csv(
                    path, sep=",", engine="python", encoding=encoding, dtype=str
                )
                if df_comma.shape[1] > 1:
                    df = df_comma
                else:
                    # 3. Last resort: let pandas guess
                    df = pd.read_csv(
                        path, sep=None, engine="python", encoding=encoding, dtype=str
                    )
            df.columns = [c.strip() for c in df.columns]
            return df
        except Exception as e:
            last_error = e
            continue
    raise RuntimeError(f"Could not read {path}: {last_error}")


def read_csv_smart_from_bytes(file_bytes: bytes, filename: str) -> pd.DataFrame:
    """Same as read_csv_smart but works with in-memory bytes (from ZIP upload)."""
    last_error = None
    for encoding in ("utf-8-sig", "latin1", "cp1252", "utf-16"):
        try:
            text = file_bytes.decode(encoding)
            break
        except (UnicodeDecodeError, Exception) as e:
            last_error = e
            continue
    else:
        raise RuntimeError(f"Could not decode {filename}: {last_error}")

    # Try semicolon first
    df = pd.read_csv(io.StringIO(text), sep=";", engine="python", dtype=str)
    if df.shape[1] <= 1:
        df_comma = pd.read_csv(io.StringIO(text), sep=",", engine="python", dtype=str)
        if df_comma.shape[1] > 1:
            df = df_comma
        else:
            df = pd.read_csv(io.StringIO(text), sep=None, engine="python", dtype=str)
    df.columns = [c.strip() for c in df.columns]
    return df


def normalize_column_name(col: str) -> str:
    """Standardise column names for comparison."""
    col = str(col).strip()
    col = re.sub(r"\s+", " ", col)
    return col


def normalize_loose(name: str) -> str:
    """Lowercase, remove all non-alphanumeric chars for fuzzy matching."""
    return re.sub(r"[^a-z0-9]", "", name.lower())


def find_column(df: pd.DataFrame, target: str) -> str | None:
    """
    Find target column in df using tiered matching:
      1. Exact match
      2. Loose match (ignore case/spacing/punctuation)
      3. Prefix match (handles columns with extra units appended)
    """
    if target in df.columns:
        return target

    norm_target = normalize_loose(target)

    for col in df.columns:
        if normalize_loose(col) == norm_target:
            return col

    candidates = [
        col for col in df.columns if normalize_loose(col).startswith(norm_target)
    ]
    if len(candidates) == 1:
        return candidates[0]
    elif len(candidates) > 1:
        candidates.sort(key=lambda c: len(normalize_loose(c)))
        return candidates[0]

    return None


def strip_date_suffix(stem: str) -> str:
    """Remove a trailing date pattern from a filename stem."""
    return DATE_SUFFIX_RE.sub("", stem).strip(" _-")


def normalize_file_for_matching(filename: str) -> str:
    """
    Convert actual Pentacam filenames into their base name for matching.
    e.g. "INDEX-LOAD_04-09-2025.CSV" → "index-load.csv"
    """
    stem = Path(filename).stem
    stem = DATE_SUFFIX_RE.sub("", stem)
    stem = stem.strip().lower()
    stem = re.sub(r"\s+", " ", stem)
    return stem + ".csv"


def normalize_text_identifier(value) -> str:
    """Normalise text identifiers for matching (uppercase, strip spaces)."""
    if pd.isna(value):
        return ""
    value = str(value).strip().upper()
    value = re.sub(r"\s+", " ", value)
    return value


def normalize_dob(value) -> str:
    """Normalise DOB to YYYY-MM-DD for consistent matching."""
    if pd.isna(value):
        return ""
    value_string = str(value).strip()
    if value_string == "":
        return ""
    try:
        date_value = pd.to_datetime(value_string, errors="coerce", dayfirst=True)
        if not pd.isna(date_value):
            return date_value.strftime("%Y-%m-%d")
    except Exception:
        pass
    return normalize_text_identifier(value)


def normalize_eye(value) -> str:
    """Standardise eye notation: R/Right/RE/OD → R, L/Left/LE/OS → L."""
    if pd.isna(value):
        return ""
    value = str(value).strip().upper()
    value = re.sub(r"\s+", "", value)
    mapping = {
        "RIGHT": "R", "RE": "R", "OD": "R", "R": "R",
        "LEFT": "L", "LE": "L", "OS": "L", "L": "L",
        "BOTH": "B", "OU": "B", "B": "B",
    }
    return mapping.get(value, value)


def normalize_identifier_value(series: pd.Series, column_name: str) -> pd.Series:
    """Normalise one identifier column's values."""
    if column_name == "D.o.Birth:":
        return series.apply(normalize_dob)
    elif column_name == "Exam Eye:":
        return series.apply(normalize_eye)
    else:
        return series.apply(normalize_text_identifier)


def load_column_mapping(mapping_file) -> dict[str, list[str]]:
    """
    Read the Pentacam_PARAMETER.xlsx mapping file.
    Returns dict: {filename: [column_headers]}
    """
    if mapping_file is None:
        return DEFAULT_COLUMN_MAP

    mapping_df = pd.read_excel(mapping_file, sheet_name="Full Re-check")
    mapping_df.columns = [normalize_column_name(c) for c in mapping_df.columns]

    required = ["Exact Export File", "Exact Column Header"]
    for col in required:
        if col not in mapping_df.columns:
            st.error(
                f"Column '{col}' not found in sheet 'Full Re-check'. "
                f"Available: {mapping_df.columns.tolist()}"
            )
            return DEFAULT_COLUMN_MAP

    mapping_df = mapping_df.dropna(subset=required).copy()
    mapping_df["Exact Export File"] = mapping_df["Exact Export File"].astype(str).str.strip()
    mapping_df["Exact Column Header"] = mapping_df["Exact Column Header"].astype(str).str.strip()

    return (
        mapping_df.groupby("Exact Export File")["Exact Column Header"]
        .apply(list)
        .to_dict()
    )


# ═══════════════════════════════════════════════════════════════════════════════
#  STEP 1 — SCAN & CLASSIFY FOLDERS
# ═══════════════════════════════════════════════════════════════════════════════


def scan_folder(root_path: Path) -> dict:
    """
    Walk the folder tree and return structured info about discovered folders.

    Returns:
        {
            "school_folders": [
                {
                    "original_folder_name": str,
                    "school_name_raw": str,
                    "school_name_normalised": str,
                    "date": str,
                    "folder_path": Path,
                    "csv_files": [str, ...],
                    "non_csv_files": [str, ...],
                }
            ],
            "unmatched_folders": [str, ...],
        }
    """
    result = {"school_folders": [], "unmatched_folders": []}

    if not root_path.exists():
        return result

    for entry in sorted(root_path.iterdir()):
        if not entry.is_dir():
            continue

        m = DATE_FOLDER_PATTERN.match(entry.name)
        if not m:
            # Check if this is a "Pentacam X SSP HYD" container folder
            # (the original structure had an extra nesting level)
            pentacam_pattern = re.compile(r"^Pentacam \d+ SSP HYD$", re.IGNORECASE)
            if pentacam_pattern.match(entry.name):
                # Recurse into this container
                sub_result = scan_folder(entry)
                result["school_folders"].extend(sub_result["school_folders"])
                result["unmatched_folders"].extend(sub_result["unmatched_folders"])
            else:
                result["unmatched_folders"].append(entry.name)
            continue

        school_raw = m.group("school").strip()
        date = m.group("date")

        csv_files = []
        non_csv_files = []
        for f in sorted(entry.iterdir()):
            if f.is_file():
                if f.suffix.lower() == ".csv":
                    csv_files.append(f.name)
                else:
                    non_csv_files.append(f.name)

        result["school_folders"].append({
            "original_folder_name": entry.name,
            "school_name_raw": school_raw,
            "school_name_normalised": normalize_school_name(school_raw),
            "date": date,
            "folder_path": entry,
            "csv_files": csv_files,
            "non_csv_files": non_csv_files,
        })

    return result


def scan_zip(zip_buffer: io.BytesIO) -> tuple[dict, dict[str, bytes]]:
    """
    Scan a ZIP file and return the same structure as scan_folder,
    plus a dict of {relative_path: file_bytes} for all CSV files.
    """
    result = {"school_folders": [], "unmatched_folders": []}
    file_contents: dict[str, bytes] = {}

    with zipfile.ZipFile(zip_buffer, "r") as zf:
        # Build a mapping of top-level folder names → their CSV/non-CSV files
        folder_map: dict[str, dict] = {}

        for info in zf.infolist():
            if info.is_dir():
                continue

            parts = Path(info.filename).parts
            if len(parts) < 2:
                continue

            # Determine the school folder — could be at level 1 or level 2
            # depending on whether there's a container like "Pentacam 1 SSP HYD"
            pentacam_pattern = re.compile(r"^Pentacam \d+ SSP HYD$", re.IGNORECASE)

            if len(parts) >= 3 and pentacam_pattern.match(parts[0]):
                folder_name = parts[1]
                file_name = parts[-1]
            elif len(parts) >= 3 and pentacam_pattern.match(parts[1]):
                folder_name = parts[2] if len(parts) > 2 else parts[1]
                file_name = parts[-1]
            else:
                # Standard case: root/SCHOOL_DATE/file.csv
                # Skip the root folder if it's just a wrapper
                if len(parts) == 2:
                    folder_name = parts[0]
                    file_name = parts[1]
                elif len(parts) >= 3:
                    # Might be root_wrapper/SCHOOL_DATE/file.csv
                    folder_name = parts[1] if not DATE_FOLDER_PATTERN.match(parts[0]) else parts[0]
                    file_name = parts[-1]
                else:
                    continue

            if folder_name not in folder_map:
                folder_map[folder_name] = {"csv": [], "non_csv": [], "paths": {}}

            if file_name.lower().endswith(".csv"):
                folder_map[folder_name]["csv"].append(file_name)
                # Store file content
                key = f"{folder_name}/{file_name}"
                folder_map[folder_name]["paths"][file_name] = info.filename
                file_contents[key] = zf.read(info.filename)
            else:
                folder_map[folder_name]["non_csv"].append(file_name)

        # Classify folders
        for folder_name, files in folder_map.items():
            m = DATE_FOLDER_PATTERN.match(folder_name)
            if not m:
                result["unmatched_folders"].append(folder_name)
                continue

            school_raw = m.group("school").strip()
            date = m.group("date")

            result["school_folders"].append({
                "original_folder_name": folder_name,
                "school_name_raw": school_raw,
                "school_name_normalised": normalize_school_name(school_raw),
                "date": date,
                "folder_path": None,  # No filesystem path for ZIP entries
                "csv_files": files["csv"],
                "non_csv_files": files["non_csv"],
            })

    return result, file_contents


# ═══════════════════════════════════════════════════════════════════════════════
#  STEP 2 — GROUP SCHOOLS
# ═══════════════════════════════════════════════════════════════════════════════


def group_by_school(scan_result: dict) -> dict[str, list[dict]]:
    """
    Group scanned folders by normalised school name.
    Returns: {normalised_school_name: [folder_info, ...]}
    """
    grouped: dict[str, list[dict]] = {}
    for folder_info in scan_result["school_folders"]:
        key = folder_info["school_name_normalised"]
        grouped.setdefault(key, []).append(folder_info)
    return dict(sorted(grouped.items()))


# ═══════════════════════════════════════════════════════════════════════════════
#  STEP 3 & 4 — PROCESS CSVs + EXTRACT COLUMNS + MERGE
# ═══════════════════════════════════════════════════════════════════════════════


def process_school(
    school_name: str,
    date_folders: list[dict],
    column_map: dict[str, list[str]],
    file_contents: dict[str, bytes] | None = None,
) -> tuple[pd.DataFrame | None, list[dict]]:
    """
    Process all CSV files for a single school.

    Args:
        school_name: normalised school name
        date_folders: list of folder_info dicts for this school
        column_map: {csv_file_base: [columns_to_extract]}
        file_contents: if from ZIP, dict of {folder/file: bytes}

    Returns:
        (merged_dataframe, validation_records)
    """
    validation: list[dict] = []

    # ── Collect all CSVs across dates, grouping by base name ────────────
    # base_name → list of (dataframe, source_info)
    base_groups: dict[str, list[tuple[pd.DataFrame, str]]] = {}

    for folder_info in date_folders:
        folder_name = folder_info["original_folder_name"]
        date = folder_info["date"]

        for csv_name in folder_info["csv_files"]:
            # Read the CSV
            try:
                if file_contents is not None:
                    key = f"{folder_name}/{csv_name}"
                    if key not in file_contents:
                        continue
                    df = read_csv_smart_from_bytes(file_contents[key], csv_name)
                else:
                    csv_path = folder_info["folder_path"] / csv_name
                    df = read_csv_smart(csv_path)

                # Normalise column names
                df.columns = [normalize_column_name(c) for c in df.columns]

                # Determine base name
                base = normalize_file_for_matching(csv_name)
                source_tag = f"{csv_name} (from {folder_name})"
                base_groups.setdefault(base, []).append((df, source_tag))

            except Exception as e:
                validation.append({
                    "School": school_name,
                    "Folder": folder_name,
                    "File": csv_name,
                    "Status": "Read Error",
                    "Column": "",
                    "Details": str(e),
                })

    if not base_groups:
        return None, validation

    # ── For each required file type, extract columns ────────────────────
    per_file_tables: dict[str, pd.DataFrame] = {}
    name_lookup: dict[str, dict] = {}

    for required_file, required_columns in column_map.items():
        normalised_required = normalize_file_for_matching(required_file)

        if normalised_required not in base_groups:
            validation.append({
                "School": school_name,
                "Folder": "ALL",
                "File": required_file,
                "Status": "Missing File",
                "Column": "",
                "Details": f"No CSVs matching '{required_file}' found",
            })
            continue

        # Concatenate all dated versions of this file type
        dfs_and_sources = base_groups[normalised_required]
        combined_parts = []
        for df, source_tag in dfs_and_sources:
            df_copy = df.copy()
            df_copy["__source__"] = source_tag
            combined_parts.append(df_copy)

        full_df = pd.concat(combined_parts, ignore_index=True, sort=False)

        # ── Find identifier columns ─────────────────────────────────
        id_cols_found: dict[str, str] = {}
        missing_ids: list[str] = []

        for id_col in IDENTIFIER_COLUMNS:
            found = find_column(full_df, id_col)
            if found:
                id_cols_found[id_col] = found
            else:
                missing_ids.append(id_col)

        if missing_ids:
            for mid in missing_ids:
                validation.append({
                    "School": school_name,
                    "Folder": "ALL",
                    "File": required_file,
                    "Status": "Missing Identifier",
                    "Column": mid,
                    "Details": "",
                })

        # Check if we have at least the primary match columns
        primary_available = all(
            pk in id_cols_found for pk in PRIMARY_MATCH_COLUMNS
        )
        if not primary_available:
            validation.append({
                "School": school_name,
                "Folder": "ALL",
                "File": required_file,
                "Status": "Skipped",
                "Column": "",
                "Details": "Primary identifier columns (Pat-ID, DOB, Eye) missing",
            })
            continue

        # ── Find target data columns ─────────────────────────────────
        target_cols_found: list[str] = []
        for target_col in required_columns:
            if target_col in IDENTIFIER_COLUMNS:
                continue
            found = find_column(full_df, target_col)
            if found:
                target_cols_found.append(found)
            else:
                validation.append({
                    "School": school_name,
                    "Folder": "ALL",
                    "File": required_file,
                    "Status": "Missing Column",
                    "Column": target_col,
                    "Details": "",
                })

        # ── Build extracted DataFrame ─────────────────────────────────
        extracted = pd.DataFrame()

        # Add identifiers
        for canon_name, actual_name in id_cols_found.items():
            extracted[canon_name] = full_df[actual_name].copy()

        # Add target data columns
        for col in target_cols_found:
            extracted[normalize_column_name(col)] = full_df[col].copy()

        # ── Create normalised match key ────────────────────────────────
        for id_col in IDENTIFIER_COLUMNS:
            if id_col in extracted.columns:
                extracted[f"_MATCH_{id_col}"] = normalize_identifier_value(
                    extracted[id_col], id_col
                )

        extracted["_MATCH_KEY"] = (
            extracted.get("_MATCH_Pat-ID:", pd.Series([""] * len(extracted))).astype(str)
            + "|"
            + extracted.get("_MATCH_D.o.Birth:", pd.Series([""] * len(extracted))).astype(str)
            + "|"
            + extracted.get("_MATCH_Exam Eye:", pd.Series([""] * len(extracted))).astype(str)
        )

        # Build name lookup from this file
        for _, row in extracted.iterrows():
            key = row.get("_MATCH_KEY", "")
            if key and key not in name_lookup:
                name_lookup[key] = {
                    id_col: row.get(id_col, "")
                    for id_col in IDENTIFIER_COLUMNS
                    if id_col in extracted.columns
                }

        # Drop duplicate match records
        extracted = extracted.drop_duplicates(subset="_MATCH_KEY", keep="first")

        # Keep only match key + parameter columns
        param_cols = [
            c for c in extracted.columns
            if c not in IDENTIFIER_COLUMNS and not c.startswith("_MATCH_")
        ]
        # Prefix parameter columns with file stem to avoid collisions
        file_stem = Path(required_file).stem
        rename_map = {}
        for c in param_cols:
            # Check if this column already exists in a previously processed table
            existing = set()
            for prev_df in per_file_tables.values():
                existing.update(prev_df.columns)
            if c in existing:
                rename_map[c] = f"{file_stem}_{c}"

        extracted = extracted.rename(columns=rename_map)

        keep_cols = ["_MATCH_KEY"] + [
            c for c in extracted.columns
            if c not in IDENTIFIER_COLUMNS and not c.startswith("_MATCH_")
        ]
        extracted = extracted[keep_cols]

        per_file_tables[required_file] = extracted

    if not per_file_tables:
        return None, validation

    # ── Merge all file-type tables together ──────────────────────────────
    merged = None
    for key, df in per_file_tables.items():
        if merged is None:
            merged = df
            continue
        merged = pd.merge(
            merged, df, on="_MATCH_KEY", how="outer", suffixes=("", f"_dup_{key}")
        )

    if merged is None or merged.empty:
        return None, validation

    # ── Reconstruct identifiers from name_lookup ─────────────────────────
    identifier_rows = []
    for match_key in merged["_MATCH_KEY"]:
        if match_key in name_lookup:
            identifier_rows.append(name_lookup[match_key])
        else:
            parts = str(match_key).split("|")
            identifier_rows.append({
                "Last Name:": "",
                "First Name:": "",
                "Pat-ID:": parts[0] if len(parts) > 0 else "",
                "D.o.Birth:": parts[1] if len(parts) > 1 else "",
                "Exam Eye:": parts[2] if len(parts) > 2 else "",
            })

    identifier_df = pd.DataFrame(identifier_rows)

    # Remove internal columns
    parameter_df = merged.drop(columns=["_MATCH_KEY"])

    # Remove any _dup_ columns created by merge
    dup_cols = [c for c in parameter_df.columns if "_dup_" in str(c)]
    if dup_cols:
        parameter_df = parameter_df.drop(columns=dup_cols)

    # Combine identifiers + parameters
    final_df = pd.concat(
        [identifier_df.reset_index(drop=True), parameter_df.reset_index(drop=True)],
        axis=1,
    )

    # Remove completely empty rows
    final_df = final_df.dropna(how="all")

    return final_df, validation


# ═══════════════════════════════════════════════════════════════════════════════
#  EXCEL WRITERS
# ═══════════════════════════════════════════════════════════════════════════════


def dataframe_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convert a DataFrame to an in-memory Excel file."""
    buf = io.BytesIO()
    df.to_excel(buf, index=False, engine="openpyxl")
    buf.seek(0)
    return buf.getvalue()


def multi_sheet_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    """Write multiple DataFrames to different sheets in one Excel file."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            # Excel sheet names max 31 chars
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
    buf.seek(0)
    return buf.getvalue()


# ═══════════════════════════════════════════════════════════════════════════════
#  STREAMLIT APP
# ═══════════════════════════════════════════════════════════════════════════════


def main():
    st.set_page_config(
        page_title="Pentacam CSV Pipeline",
        page_icon="👁️",
        layout="wide",
    )

    st.title("👁️ Pentacam CSV Pipeline")
    st.caption(
        "Standardise school folders → Extract CSVs → Auto-detect delimiters → "
        "Extract columns → Merge per-patient → Download Excel"
    )

    # ── Sidebar ──────────────────────────────────────────────────────────
    with st.sidebar:
        st.header("⚙️ Configuration")

        input_mode = st.radio(
            "Input Mode",
            ["📁 Local Folder Path", "📦 Upload ZIP File"],
            help=(
                "**Local Folder Path**: Paste the path to the folder on this machine.\n\n"
                "**Upload ZIP**: Upload a ZIP containing the school folders."
            ),
        )

        folder_path = None
        zip_file = None

        if "folder_path" not in st.session_state:
            st.session_state.folder_path = ""

        if input_mode == "📁 Local Folder Path":
            folder_path = st.text_input(
                "Folder Path",
                placeholder=r"C:\Pentacam 1 SSP HYD",
                help="Full path to the root folder on the server/local machine.",
            )
            st.warning("⚠️ **Note:** This option only works if the app is running on your local computer. Since you deployed to a server, the server cannot access your computer's C: drive. Please use the **Upload ZIP File** option instead.")
        else:
            zip_file = st.file_uploader(
                "Upload ZIP",
                type=["zip"],
                help="ZIP file containing school folders (e.g. SCHOOL_DD-MM-YYYY/...CSV)",
            )

        st.divider()

        mapping_file = st.file_uploader(
            "📊 Column Mapping File (optional)",
            type=["xlsx"],
            help=(
                "Upload Pentacam_PARAMETER.xlsx to customise which columns to extract. "
                "Leave empty to use the default mapping."
            ),
        )

        st.divider()

        run_button = st.button("🚀 Run Pipeline", type="primary", use_container_width=True)

    # ── Main Area ────────────────────────────────────────────────────────
    if not run_button:
        st.info(
            "Configure the input source in the sidebar and click **Run Pipeline** to start."
        )

        with st.expander("ℹ️ Expected Folder Structure", expanded=False):
            st.code(
                """
Source_Folder/
├── AHAD HIGH SCHOOL_25-03-2026/
│   ├── BADisplay-LOAD.CSV
│   ├── CHAMBER-LOAD.CSV
│   ├── INDEX-LOAD.CSV
│   ├── PACHY-LOAD.CSV
│   ├── CorneoScleral-LOAD.CSV
│   ├── some_image.jpg          ← ignored
│   └── report.pdf              ← ignored
├── AHAD HIGH SCHOOL_26-03-2026/
│   ├── BADisplay-LOAD.CSV
│   └── ...
├── Alif high school _ 05-01-2026/
│   └── ...  (same school, different casing)
└── Army Public School_10-12-2025/
    └── ...
                """,
                language=None,
            )

        with st.expander("📊 Default Column Mapping", expanded=False):
            for csv_file, columns in DEFAULT_COLUMN_MAP.items():
                st.markdown(f"**{csv_file}**")
                for col in columns:
                    st.markdown(f"  - `{col}`")
        return

    # ── Validate Input ───────────────────────────────────────────────────
    if input_mode == "📁 Local Folder Path" and not folder_path:
        st.error("Please enter a folder path.")
        return
    if input_mode == "📦 Upload ZIP File" and zip_file is None:
        st.error("Please upload a ZIP file.")
        return

    # ── Load Column Mapping ──────────────────────────────────────────────
    column_map = load_column_mapping(mapping_file)

    # ══════════════════════════════════════════════════════════════════════
    #  STEP 1: SCAN & CLASSIFY
    # ══════════════════════════════════════════════════════════════════════
    step1 = st.expander("📂 Step 1: Scan & Classify Folders", expanded=True)
    file_contents = None

    with step1:
        with st.spinner("Scanning folders..."):
            if input_mode == "📁 Local Folder Path":
                root = Path(folder_path)
                if not root.exists():
                    st.error(f"Folder not found: `{folder_path}`")
                    return
                scan_result = scan_folder(root)
            else:
                zip_buffer = io.BytesIO(zip_file.getvalue())
                scan_result, file_contents = scan_zip(zip_buffer)

        total_folders = len(scan_result["school_folders"])
        total_csvs = sum(len(f["csv_files"]) for f in scan_result["school_folders"])
        total_non_csv = sum(len(f["non_csv_files"]) for f in scan_result["school_folders"])

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("School-Date Folders", total_folders)
        col2.metric("CSV Files", total_csvs)
        col3.metric("Non-CSV Files (ignored)", total_non_csv)
        col4.metric("Unmatched Folders", len(scan_result["unmatched_folders"]))

        if total_folders == 0:
            st.error(
                "No folders matching the `SCHOOLNAME_DD-MM-YYYY` pattern were found. "
                "Please check your folder structure."
            )
            return

        # Show folder details
        folder_summary = []
        for f in scan_result["school_folders"]:
            folder_summary.append({
                "Folder": f["original_folder_name"],
                "School (raw)": f["school_name_raw"],
                "Date": f["date"],
                "CSVs": len(f["csv_files"]),
                "Other Files": len(f["non_csv_files"]),
            })
        st.dataframe(pd.DataFrame(folder_summary), use_container_width=True, hide_index=True)

        if scan_result["unmatched_folders"]:
            st.warning(
                f"**{len(scan_result['unmatched_folders'])} folder(s) skipped** "
                f"(don't match SCHOOL_DATE pattern): "
                f"{', '.join(scan_result['unmatched_folders'][:10])}"
            )

    # ══════════════════════════════════════════════════════════════════════
    #  STEP 2: STANDARDISE & GROUP SCHOOLS
    # ══════════════════════════════════════════════════════════════════════
    step2 = st.expander("🏫 Step 2: Standardise & Group Schools", expanded=True)

    with step2:
        grouped = group_by_school(scan_result)
        st.metric("Unique Schools (after normalisation)", len(grouped))

        group_summary = []
        for norm_name, folders in grouped.items():
            raw_names = sorted(set(f["school_name_raw"] for f in folders))
            dates = sorted(f["date"] for f in folders)
            total_csv = sum(len(f["csv_files"]) for f in folders)
            group_summary.append({
                "Normalised Name": norm_name,
                "Original Name(s)": " / ".join(raw_names),
                "Date(s)": ", ".join(dates),
                "Total CSVs": total_csv,
                "Visits": len(folders),
            })

        st.dataframe(
            pd.DataFrame(group_summary),
            use_container_width=True,
            hide_index=True,
        )

    # ══════════════════════════════════════════════════════════════════════
    #  STEP 3: PROCESS CSVs & EXTRACT COLUMNS
    # ══════════════════════════════════════════════════════════════════════
    step3 = st.expander("⚙️ Step 3: Process CSVs & Extract Columns", expanded=True)

    with step3:
        progress_bar = st.progress(0, text="Processing schools...")
        all_school_results: dict[str, pd.DataFrame] = {}
        all_validation: list[dict] = []

        for idx, (school_name, date_folders) in enumerate(grouped.items()):
            progress = (idx + 1) / len(grouped)
            progress_bar.progress(progress, text=f"Processing: {school_name}")

            result_df, val_records = process_school(
                school_name, date_folders, column_map, file_contents
            )
            all_validation.extend(val_records)

            if result_df is not None and not result_df.empty:
                all_school_results[school_name] = result_df

        progress_bar.progress(1.0, text="✅ Processing complete!")

        # Summary
        col1, col2, col3 = st.columns(3)
        col1.metric("Schools Processed", len(all_school_results))
        col2.metric(
            "Schools with No Data",
            len(grouped) - len(all_school_results),
        )
        total_patients = sum(len(df) for df in all_school_results.values())
        col3.metric("Total Patient Records", total_patients)

        # Show per-school summary
        if all_school_results:
            school_stats = []
            for name, df in all_school_results.items():
                school_stats.append({
                    "School": name,
                    "Records": len(df),
                    "Columns": len(df.columns),
                })
            st.dataframe(
                pd.DataFrame(school_stats),
                use_container_width=True,
                hide_index=True,
            )

        # Show validation issues
        if all_validation:
            val_df = pd.DataFrame(all_validation)
            issues_count = len(val_df[val_df["Status"] != "OK"]) if "OK" in val_df["Status"].values else len(val_df)
            st.warning(f"⚠️ {issues_count} validation issue(s) found")
            with st.expander("View Validation Details"):
                st.dataframe(val_df, use_container_width=True, hide_index=True)

    # ══════════════════════════════════════════════════════════════════════
    #  STEP 4: DOWNLOAD RESULTS
    # ══════════════════════════════════════════════════════════════════════
    step4 = st.expander("📥 Step 4: Download Results", expanded=True)

    with step4:
        if not all_school_results:
            st.error(
                "No data was extracted from any school. Check the validation "
                "issues above for details."
            )
            return

        st.success(
            f"✅ Successfully processed **{len(all_school_results)} school(s)** "
            f"with **{total_patients} total patient records**."
        )

        # ── Combined All-Schools Excel ────────────────────────────────
        st.subheader("📊 Combined Output (All Schools)")
        combined_parts = []
        for school_name, df in all_school_results.items():
            df_with_school = df.copy()
            df_with_school.insert(0, "School", school_name)
            combined_parts.append(df_with_school)

        combined_all = pd.concat(combined_parts, ignore_index=True, sort=False)

        st.dataframe(combined_all.head(50), use_container_width=True, hide_index=True)
        if len(combined_all) > 50:
            st.caption(f"Showing first 50 of {len(combined_all)} rows. Download for full data.")

        combined_bytes = dataframe_to_excel_bytes(combined_all)
        st.download_button(
            "⬇️ Download All Schools Combined (Excel)",
            data=combined_bytes,
            file_name="All_Schools_Combined.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

        # ── Per-School Excel (multi-sheet workbook) ───────────────────
        st.subheader("📁 Per-School Output")

        sheets = {
            name[:31]: df for name, df in all_school_results.items()
        }
        per_school_bytes = multi_sheet_excel_bytes(sheets)

        st.download_button(
            "⬇️ Download Per-School Workbook (one sheet per school)",
            data=per_school_bytes,
            file_name="Per_School_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # ── Validation Report ─────────────────────────────────────────
        if all_validation:
            st.subheader("📋 Validation Report")
            val_df = pd.DataFrame(all_validation)
            val_bytes = dataframe_to_excel_bytes(val_df)
            st.download_button(
                "⬇️ Download Validation Report (Excel)",
                data=val_bytes,
                file_name="Validation_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        # ── Individual School Downloads ───────────────────────────────
        st.subheader("📄 Individual School Files")
        with st.expander("Download individual school files"):
            for school_name, df in all_school_results.items():
                safe_name = re.sub(r'[<>:"/\\|?*]', '_', school_name)
                school_bytes = dataframe_to_excel_bytes(df)
                st.download_button(
                    f"⬇️ {school_name} ({len(df)} records)",
                    data=school_bytes,
                    file_name=f"{safe_name}_Consolidated.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_{safe_name}",
                )


if __name__ == "__main__":
    main()
