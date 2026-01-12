#!/usr/bin/env python3
"""
Denodo Extract v2 (enhanced)
----------------------------
Pulls JSON data from a Denodo-backed HTML page, aggregates rows by `identifier`
into one row with per-field distinct lists, and writes JSON/CSV/XLSX outputs.
Lists are stringified for CSV/Excel using " | " as the separator.

Key features:
- Structured logging: console + rotating file handler (logs roll @ 1 MB, keep last 3 files)
- Config validation: clear, actionable errors if required keys are missing/invalid
- Excel output: writes .xlsx with an actual Excel Table (styled) in addition to CSV
- TLS control: --verify-ca supports true/false or a CA bundle path
- Output placement: --output-dir to pick a folder, or --use-config-dir to write next to the YAML
- Dual CLI/GUI: run headless with CLI flags or interactively with --gui prompts
- No-arg start: if launched without flags, you'll see a brief instruction prompt and then GUI mode

How to run (CLI):
    python Denodo_Extract.py --config /path/to/config.yml --verify-ca true
    # with custom output folder:
    python Denodo_Extract.py --config /path/to/config.yml --output-dir D:/Exports
    # use CA bundle:
    python Denodo_Extract.py --config /path/to/config.yml --verify-ca C:/corp/ca-bundle.pem

How to run (GUI):
    python Denodo_Extract.py --gui
    # prompts for config, output directory, and optional lookback

YAML expectations (minimal):
STORAGE:
  LOCAL_SYNC:
    ROOT_DIR: C:\\Data\\Denodo
PATHS:
  OUTPUT_DIR: exports
  SUBSET_DIR: output\\subsets
  LOG_DIR: logs  # optional, defaults to <output_dir>/logs
FILES:
  EAMS_RESULT_DATA: denodo_raw.json
  EAMS_CSV_OUTPUT: denodo_export.csv
  EAMS_RUN_DATE: last_run.txt
  EAMS_CONTACT_INFO: contact_info.csv
  EAMS_SERVER_INFO: server_info.csv
  # EAMS_SPLUNK_SERVER_FEED: optional (defaults to splunk_server_feed.csv)
  # EAMS_XLSX_OUTPUT: optional (defaults to same stem as CSV + .xlsx)
DATA_SOURCE:  # or legacy DATASOURCE
  URL: https://denodo.example.local/report123
AUTHENTICATION:
  SERVICE_NAME: denodo-prod
  USERNAME: user
SUBSETS:
  CONTACT_INFO: [field1, field2]
  SERVICE_IPS: [field3, field4]
  # SPLUNK_SERVER_FEED: [field5, field6]
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
from logging.handlers import RotatingFileHandler
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import re
_ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
_EXCEL_CELL_LIMIT = 32767

import requests
import yaml
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    import keyring  # type: ignore
except Exception:  # pragma: no cover
    keyring = None

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


def normalize_cell_value(val: Any) -> str:
    if val is None:
        s = ""
    elif isinstance(val, (list, tuple, set)):
        # Lists are flattened for CSV/Excel using a pipe separator to avoid clashing with commas in values.
        s = " | ".join("" if v is None else str(v) for v in val)
    elif isinstance(val, dict):
        import json
        try:
            s = json.dumps(val, ensure_ascii=False, separators=(",", ":"))
        except Exception:
            s = str(val)
    else:
        s = str(val)
    s = _ILLEGAL_XML_RE.sub("", s)
    return s[:_EXCEL_CELL_LIMIT - 1] + "…" if len(s) > _EXCEL_CELL_LIMIT else s


# ---------------------------
# Logging
# ---------------------------

def setup_logging(log_dir: Path, level: str = "INFO") -> None:
    """
    Configure application logging.

    Creates the log directory if needed and configures the root logger with:
    - a console handler, and
    - a rotating file handler (denodo_extract.log).

    Args:
        log_dir: Directory where log files will be written.
        level: Log level name (e.g., "INFO", "DEBUG"). Defaults to "INFO".
    """
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "denodo_extract.log"

    # Parse level safely
    level_value = getattr(logging, level.upper(), logging.INFO)

    fmt = logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")

    # Console handler
    ch = logging.StreamHandler()
    ch.setLevel(level_value)
    ch.setFormatter(fmt)

    # Rotating file handler (1 MB, keep last 3 files)
    fh = RotatingFileHandler(log_path, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    fh.setLevel(level_value)
    fh.setFormatter(fmt)

    root = logging.getLogger()
    root.setLevel(level_value)
    # Avoid duplicate handlers if re-run in the same interpreter
    root.handlers.clear()
    root.addHandler(ch)
    root.addHandler(fh)

# ---------------------------
# Config
# ---------------------------

class ConfigError(Exception):
    """Raised when the YAML configuration is missing required keys or invalid."""


def choose_config_file() -> Path:
    """
    Show a GUI file picker to select a YAML configuration file.

    Returns:
        Path to the selected YAML file.

    Raises:
        SystemExit: If the user cancels the dialog.
    """
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    filepath = filedialog.askopenfilename(
        title="Select YAML configuration file",
        filetypes=[("YAML files", "*.yml *.yaml"), ("All files", "*.*")],
    )
    if not filepath:
        messagebox.showerror("No file selected", "No configuration file was selected. Exiting.")
        sys.exit(1)
    return Path(filepath)


def choose_output_dir(default: Optional[Path] = None) -> Optional[Path]:
    """
    Show a GUI folder picker to select an output directory.

    Args:
        default: Optional default directory to preselect.

    Returns:
        Path to the selected directory, or None if the user cancels.
    """
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    initial = default.as_posix() if default else None
    path = filedialog.askdirectory(title="Select output directory", initialdir=initial)
    return Path(path) if path else None


def show_start_instructions() -> None:
    """
    Display basic start/run instructions for interactive (GUI) usage.

    Returns:
        None
    """
    message = (
        "Denodo Extract can run either way:\n\n"
        "  CLI (headless):\n"
        "    python Denodo_Extract.py --config <path/to/config.yml>\n"
        "      [--output-dir <path>] [--verify-ca true|false|/path]\n\n"
        "  GUI (default when no flags): prompts for config/output and optional lookback.\n\n"
        "Continuing with GUI..."
    )
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("Denodo Extract", message)
    except Exception:
        print(message)


def prompt_lookback_days() -> Optional[float]:
    """
    Prompt the user (GUI) for a lookback window in days.

    Returns:
        Number of days as a float, or None if the user cancels or leaves it blank.
    """
    import tkinter as tk
    from tkinter import messagebox, simpledialog

    root = tk.Tk()
    root.withdraw()
    value = simpledialog.askstring(
        "Lookback window",
        "Enter lookback in days (e.g., 2, 0.5 for 12h, 1/24 ≈ 0.0417 for 1h). Leave blank to skip:",
    )
    if value is None:
        return None  # cancelled
    value = value.strip()
    if not value:
        return None
    try:
        days = float(value)
        if days < 0:
            raise ValueError("Lookback must be non-negative")
        return days
    except Exception:
        messagebox.showerror("Invalid input", "Please enter a non-negative number (e.g., 2 or 12).")
        return None


def load_config(path: Path) -> Dict[str, Any]:
    """
    Load a YAML configuration file.

    Args:
        path: Path to a YAML configuration file.

    Returns:
        Parsed configuration as a dictionary.

    Raises:
        FileNotFoundError: If the YAML file does not exist.
        yaml.YAMLError: If the YAML cannot be parsed.
    """
    try:
        with path.open("r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
    except Exception as exc:  # pragma: no cover
        raise ConfigError(f"Failed to read YAML config: {exc}") from exc
    if not isinstance(cfg, dict):
        raise ConfigError("Top-level YAML must be a mapping (dictionary).")
    return cfg


def validate_config(cfg: Dict[str, Any]) -> Dict[str, Any]:
    """
    Validate and normalize configuration.

    Validates that required sections/keys exist and normalizes legacy aliases
    (e.g., DATASOURCE -> DATA_SOURCE). Intended to fail fast with actionable errors.

    Args:
        cfg: Raw configuration dictionary loaded from YAML.

    Returns:
        Validated configuration dictionary (may include normalized keys).

    Raises:
        ValueError: If required sections/keys are missing or invalid.
    """
    missing = []
    # Required sections
    for sect in ("STORAGE", "PATHS", "FILES", "AUTHENTICATION", "SUBSETS"):
        if sect not in cfg or not isinstance(cfg[sect], dict):
            missing.append(sect)
    if "DATA_SOURCE" not in cfg and "DATASOURCE" not in cfg:
        missing.append("DATA_SOURCE")

    if missing:
        raise ConfigError(f"Missing or invalid sections in YAML: {', '.join(missing)}")

    # STORAGE
    storage = cfg["STORAGE"]
    if "LOCAL_SYNC" not in storage or not isinstance(storage["LOCAL_SYNC"], dict):
        raise ConfigError("STORAGE.LOCAL_SYNC is required and must be a mapping")
    local_sync = storage["LOCAL_SYNC"]
    if "ROOT_DIR" not in local_sync or not isinstance(local_sync["ROOT_DIR"], str) or not local_sync["ROOT_DIR"].strip():
        raise ConfigError("STORAGE.LOCAL_SYNC.ROOT_DIR is required and must be a non-empty string")

    # PATHS
    paths = cfg["PATHS"]
    for key in ("OUTPUT_DIR", "SUBSET_DIR"):
        if key not in paths or not isinstance(paths[key], str) or not paths[key].strip():
            raise ConfigError(f"PATHS.{key} is required and must be a non-empty string")
    if "LOG_DIR" in paths and (not isinstance(paths["LOG_DIR"], str) or not paths["LOG_DIR"].strip()):
        raise ConfigError("PATHS.LOG_DIR must be a non-empty string if provided")

    # FILES
    files = cfg["FILES"]
    data_source = cfg.get("DATA_SOURCE") or cfg.get("DATASOURCE")
    if not isinstance(data_source, dict):
        raise ConfigError("DATA_SOURCE (or legacy DATASOURCE) must be a mapping")
    # normalize for downstream use
    cfg["DATA_SOURCE"] = data_source

    for key in ("EAMS_RESULT_DATA", "EAMS_CSV_OUTPUT", "EAMS_RUN_DATE", "EAMS_SERVER_INFO", "EAMS_CONTACT_INFO"):
        if key not in files or not isinstance(files[key], str) or not files[key].strip():
            raise ConfigError(f"FILES.{key} is required and must be a non-empty string")

    url_value = data_source.get("URL")
    if not isinstance(url_value, str) or not url_value.strip():
        legacy_url = files.get("URL")
        if isinstance(legacy_url, str) and legacy_url.strip():
            url_value = legacy_url
            data_source["URL"] = legacy_url  # mirror for downstream use
        else:
            raise ConfigError("DATA_SOURCE.URL is required and must be a non-empty string (legacy configs may use FILES.URL)")

    # Allow optional keys but ensure strings if present
    for optional_key, value in files.items():
        if not isinstance(value, str):
            raise ConfigError(f"FILES.{optional_key} must be a string")

    # AUTHENTICATION
    auth = cfg["AUTHENTICATION"]
    for key in ("SERVICE_NAME", "USERNAME"):
        if key not in auth or not isinstance(auth[key], str) or not auth[key].strip():
            raise ConfigError(f"AUTHENTICATION.{key} is required and must be a non-empty string")

    subsets = cfg["SUBSETS"]
    if "CONTACT_INFO" not in subsets or not isinstance(subsets["CONTACT_INFO"], list) or not subsets["CONTACT_INFO"]:
        raise ConfigError("SUBSETS.CONTACT_INFO is required and must be a non-empty list")
    if "SERVICE_IPS" not in subsets or not isinstance(subsets["SERVICE_IPS"], list) or not subsets["SERVICE_IPS"]:
        raise ConfigError("SUBSETS.SERVICE_IPS is required and must be a non-empty list")
    for key, values in subsets.items():
        if not isinstance(values, list):
            raise ConfigError(f"SUBSETS.{key} must be a list of field names")
        for v in values:
            if not isinstance(v, str) or not v.strip():
                raise ConfigError(f"SUBSETS.{key} entries must be non-empty strings")

    return cfg


# ---------------------------
# HTTP
# ---------------------------

def create_session() -> requests.Session:
    """
    Create a requests session with retry/backoff.

    Returns:
        A configured requests.Session that retries transient HTTP failures.
    """
    session = requests.Session()
    retries = Retry(
        total=3,
        backoff_factor=0.3,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["HEAD", "GET", "OPTIONS"],
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retries)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


def interpret_verify_ca(value: str | None) -> str | bool:
    """
    Interpret the TLS verification option from CLI.

    Args:
        value: One of:
            - "true"/"false" (case-insensitive) to enable/disable TLS verification
            - a filesystem path to a CA bundle
            - None to use default verification (enabled)

    Returns:
        True/False, or a string path to a CA bundle file.
    """
    if value is None:
        return True
    v = value.strip().lower()
    if v in ("true", "1", "yes"):
        return True
    if v in ("false", "0", "no"):
        return False
    return value  # presumed path to CA bundle


def fetch_page(session: requests.Session, url: str, username: str, password: str, verify_ca: str | bool) -> str:
    """
    Fetch the Denodo-backed HTML page using HTTP Basic Authentication.

    Args:
        session: Requests session to use.
        url: Target URL to fetch.
        username: HTTP Basic Auth username.
        password: HTTP Basic Auth password.
        verify_ca: TLS verification setting (True/False or CA bundle path).

    Returns:
        Response body as HTML text.

    Raises:
        requests.RequestException: On network/HTTP failures.
    """
    import urllib3
    if verify_ca is False:
        # Warn about disabling TLS verification
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    logging.info("Fetching URL: %s (verify=%s)", url, verify_ca if verify_ca is not True else "True")
    resp = session.get(url, auth=(username, password), timeout=30, verify=verify_ca)
    try:
        resp.raise_for_status()
    except requests.HTTPError:
        logging.error("HTTP error %s: %s", resp.status_code, resp.text[:500])
        raise
    return resp.text


# ---------------------------
# Parsing & Output
# ---------------------------

def extract_json_data(html: str) -> List[Dict[str, Any]]:
    """
    Extract embedded JSON rows from the HTML page.

    The source page is expected to contain a JavaScript assignment like:
        var data = [ ... ];

    Args:
        html: Full HTML response body.

    Returns:
        Parsed list of row dictionaries.

    Raises:
        ValueError: If the embedded JSON payload cannot be found or parsed.
    """
    match = re.search(r"var\s+data\s*=\s*(\[\s*{.*?}\s*\])", html, re.DOTALL)
    if not match:
        raise ValueError("Expected JSON pattern not found in HTML (looked for 'var data = [...]').")
    payload = match.group(1)
    try:
        return json.loads(payload)
    except json.JSONDecodeError as exc:
        snippet = payload[:2000]
        logging.debug("Failed payload snippet: %s", snippet)
        raise ValueError(f"Failed to parse embedded JSON: {exc}")


def ensure_output_dirs(out_dir: Path) -> None:
    """
    Ensure the output directory exists.

    Args:
        out_dir: Output directory path.

    Returns:
        None
    """
    print(f"Ensuring output directory exists: {out_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)


def parse_run_date(path: Path) -> Optional[datetime]:
    """
    Parse a stored run timestamp from a text file.

    Args:
        path: Path to the run date file.

    Returns:
        Parsed datetime if present and valid; otherwise None.
    """
    if not path.exists():
        return None
    try:
        text = path.read_text(encoding="utf-8").strip()
        if not text:
            return None
        # Support trailing 'Z' and fractional seconds.
        cleaned = text.replace("Z", "+00:00") if text.endswith("Z") else text
        return datetime.fromisoformat(cleaned)
    except Exception as exc:
        logging.warning("Could not parse previous run date in %s: %s", path.as_posix(), exc)
        return None


def write_run_date(run_date_path: Path, when: datetime) -> None:
    """
    Write a run timestamp to the run date file.

    Args:
        run_date_path: Path to the run date file.
        when: Timestamp to record.

    Returns:
        None
    """
    run_date_path.parent.mkdir(parents=True, exist_ok=True)
    stamp = when.isoformat(timespec="seconds")
    run_date_path.write_text(stamp + "\n", encoding="utf-8")
    logging.info("Run date recorded: %s -> %s", run_date_path.as_posix(), stamp)


def parse_last_modified(value: Any) -> Optional[datetime]:
    """
    Parse a last-modified timestamp value from the dataset.

    Args:
        value: Input value which may be a string, datetime, or None.

    Returns:
        A datetime if parsable; otherwise None.
    """
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    try:
        cleaned = s.replace("Z", "+00:00") if s.endswith("Z") else s
        return datetime.fromisoformat(cleaned)
    except Exception:
        return None


def filter_modified_rows(rows: Iterable[Dict[str, Any]], since: datetime) -> List[Dict[str, Any]]:
    """
    Return rows whose lastmodifieddate is newer than 'since'.
    """
    newer = []
    for row in rows:
        modified = parse_last_modified(row.get("lastmodifieddate"))
        if modified and modified > since:
            newer.append(row)
    return newer


def aggregate_by_identifier(rows: Iterable[Dict[str, Any]], key: str = "identifier") -> List[Dict[str, Any]]:
    """
    Aggregate source rows by identifier.

    For each identifier, collects distinct values per field into lists. Empty/None
    values are ignored. Produces one normalized record per identifier.

    Args:
        rows: Iterable of source row dictionaries.
        key: Field name to group by. Defaults to "identifier".

    Returns:
        List of aggregated row dictionaries.
    """
    grouped: Dict[Any, Dict[str, Any]] = {}
    for row in rows:
        identifier_value = row.get(key)
        bucket = grouped.setdefault(identifier_value, {key: identifier_value, "_seen": {}, "_data": {}})
        seen = bucket["_seen"]
        data = bucket["_data"]

        for field, val in row.items():
            if field == key:
                continue
            if val is None:
                continue
            if isinstance(val, str) and not val.strip():
                continue

            field_seen = seen.setdefault(field, set())
            try:
                signature = json.dumps(val, sort_keys=True, default=str)
            except Exception:
                signature = repr(val)
            if signature in field_seen:
                continue

            field_seen.add(signature)
            data.setdefault(field, []).append(val)

    aggregated_rows: List[Dict[str, Any]] = []
    for bucket in grouped.values():
        row_out = {key: bucket.get(key)}
        for field, vals in bucket["_data"].items():
            row_out[field] = vals
        aggregated_rows.append(row_out)
    return aggregated_rows


def marker_name(base_name: str, marker: str) -> str:
    """
    Create a derived filename by inserting a marker before the extension.

    Args:
        base_name: Base filename (e.g., "denodo_export.csv").
        marker: Marker string (e.g., "changes", "modified").

    Returns:
        New filename including the marker.
    """
    p = Path(base_name)
    return f"{p.stem}_{marker}{p.suffix or ''}"


def timestamped_name(base_name: str, ts: str) -> str:
    """
    Create a timestamped filename.

    Args:
        base_name: Base filename.
        ts: Timestamp string (already formatted).

    Returns:
        Timestamped filename.
    """
    p = Path(base_name)
    return f"{ts}_{p.name}"


def ensure_named_file(base_name: str, default_suffix: str) -> str:
    """
    Ensure a filename has an extension.

    Args:
        base_name: Input filename (may be missing extension).
        default_suffix: Extension to apply if missing (e.g., ".json").

    Returns:
        Filename with an extension.
    """
    p = Path(base_name)
    return p.name if p.suffix else f"{p.name}{default_suffix}"


def union_fieldnames(rows: Iterable[Dict[str, Any]]) -> List[str]:
    """
    Compute the union of all keys across rows.

    Args:
        rows: Iterable of row dictionaries.

    Returns:
        Sorted list of unique field names.
    """
    fields = set()
    order = []
    for row in rows:
        for k in row.keys():
            if k not in fields:
                fields.add(k)
                order.append(k)
    return order


def write_json_snapshot(rows: List[Dict[str, Any]], out_path: Path) -> None:
    """
    Write a JSON snapshot to disk.

    Args:
        rows: Row dictionaries to write.
        out_path: Destination path.

    Returns:
        None
    """
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    logging.info("JSON snapshot saved: %s", out_path.as_posix())


def write_csv(rows: List[Dict[str, Any]], out_path: Path) -> None:
    """
    Write a CSV file with headers derived from the union of row keys.

    Args:
        rows: Row dictionaries to write.
        out_path: Destination path.

    Returns:
        None
    """
    if not rows:
        logging.warning("No rows to write to CSV.")
        return
    headers = union_fieldnames(rows)
    with out_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        for r in rows:
            writer.writerow(r)
    logging.info("CSV saved: %s", out_path.as_posix())


def write_csv_with_headers(rows: List[Dict[str, Any]], headers: List[str], out_path: Path) -> None:
    """
    Write a CSV file using an explicit header order.

    Args:
        rows: Row dictionaries to write.
        headers: Header field order to use.
        out_path: Destination path.

    Returns:
        None
    """
    if not rows:
        logging.warning("No rows to write to CSV.")
        return
    with out_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        for r in rows:
            row_out = {h: r.get(h, "") for h in headers}
            writer.writerow(row_out)
    logging.info("CSV saved: %s", out_path.as_posix())


def write_excel_table_with_headers(rows: List[Dict[str, Any]], headers: List[str], out_path: Path, table_name: str) -> None:
    """
    Write an XLSX file containing a styled Excel Table using explicit headers.

    Args:
        rows: Row dictionaries to write.
        headers: Header field order to use.
        out_path: Destination path.
        table_name: Excel table name.

    Returns:
        None
    """
    if not rows:
        logging.warning("No rows to write to Excel.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "data"
    ws.append(headers)
    for r in rows:
        ws.append([normalize_cell_value(r.get(h, "")) for h in headers])

    def col_letter(n: int) -> str:
        letters = ""
        while n:
            n, rem = divmod(n - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

    last_row = ws.max_row
    last_col = ws.max_column
    safe_name = re.sub(r"\W+", "", table_name)[:254] or "T1"
    ref = f"A1:{col_letter(last_col)}{last_row}"
    table = Table(displayName=safe_name, ref=ref)
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True, showColumnStripes=False)
    table.tableStyleInfo = style
    ws.add_table(table)

    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        col_letter_name = col[0].column_letter
        ws.column_dimensions[col_letter_name].width = min(max(12, max_len + 2), 60)

    wb.save(out_path)
    logging.info("Excel saved with table: %s", out_path.as_posix())


def subset_rows(rows: List[Dict[str, Any]], headers: List[str]) -> List[Dict[str, Any]]:
    """
    Project rows onto a subset of headers.

    Args:
        rows: Source rows.
        headers: Headers to keep.

    Returns:
        Subset rows with only the requested headers.
    """
    return [{h: r.get(h, "") for h in headers} for r in rows]


def write_excel_table(rows, out_path: Path, table_name: str = "DenodoData") -> None:
    """
    Write an XLSX file containing a styled Excel Table.

    Args:
        rows: Row dictionaries to write.
        out_path: Destination path.
        table_name: Excel table name. Defaults to "DenodoData".

    Returns:
        None
    """
    if not rows:
        logging.warning("No rows to write to Excel.")
        return

    headers = union_fieldnames(rows)
    wb = Workbook()
    ws = wb.active
    ws.title = "data"

    # Write header
    ws.append(headers)
    # Write rows
    for r in rows:
        ws.append([normalize_cell_value(r.get(h, "")) for h in headers])

    # Add table
    last_row = ws.max_row
    last_col = ws.max_column

    # Build Excel range like A1:Z999
    def col_letter(n: int) -> str:
        letters = ""
        while n:
            n, rem = divmod(n - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

    safe_name = re.sub(r"\W+", "", table_name)[:254] or "T1"
    ref = f"A1:{col_letter(last_col)}{last_row}"
    table = Table(displayName=safe_name, ref=ref)
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True, showColumnStripes=False)
    table.tableStyleInfo = style
    ws.add_table(table)

    # Autosize columns (basic)
    for col in ws.columns:
        max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
        col_letter_name = col[0].column_letter
        ws.column_dimensions[col_letter_name].width = min(max(12, max_len + 2), 60)

    wb.save(out_path)
    logging.info("Excel saved with table: %s", out_path.as_posix())


def find_latest_snapshot(output_dir: Path, base_name: str) -> Optional[Path]:
    """
    Find the most recent JSON snapshot matching the configured base name.

    Args:
        output_dir: Directory to search.
        base_name: Base JSON filename (without timestamp).

    Returns:
        Path to the latest matching snapshot, or None if not found.
    """
    pattern = f"*_{Path(base_name).name}"
    candidates = sorted(p for p in output_dir.glob(pattern) if p.is_file())
    return candidates[-1] if candidates else None


def load_json_rows(path: Path) -> List[Dict[str, Any]]:
    """
    Load rows from a JSON snapshot file.

    Args:
        path: JSON file path.

    Returns:
        List of row dictionaries.
    """
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return data
    except Exception as exc:
        logging.warning("Could not load JSON rows from %s: %s", path.as_posix(), exc)
    return []


def diff_rows_by_identifier(prev_rows: List[Dict[str, Any]], curr_rows: List[Dict[str, Any]], id_key: str = "identifier") -> List[Dict[str, Any]]:
    """
    Compute a field-level diff between snapshots.

    Compares previous vs current aggregated rows keyed by an identifier field. The
    result is a list of change records describing added/removed/changed fields.

    Args:
        prev_rows: Previous snapshot rows.
        curr_rows: Current snapshot rows.
        id_key: Field used as the unique identifier. Defaults to "identifier".

    Returns:
        List of change dictionaries with keys like identifier/field/old/new/change_type.
    """
    prev_map = {r.get(id_key): r for r in prev_rows if r.get(id_key) is not None}
    curr_ids = set()
    changes: List[Dict[str, Any]] = []

    for curr in curr_rows:
        ident = curr.get(id_key)
        curr_ids.add(ident)
        prev = prev_map.get(ident)
        if prev is None:
            for field, new_val in curr.items():
                if field == id_key:
                    continue
                changes.append({"identifier": ident, "field": field, "old": "", "new": new_val, "change_type": "added"})
            continue

        all_fields = set(prev.keys()) | set(curr.keys())
        for field in all_fields:
            if field == id_key:
                continue
            old_val = prev.get(field, "")
            new_val = curr.get(field, "")
            if old_val != new_val:
                changes.append({"identifier": ident, "field": field, "old": old_val, "new": new_val, "change_type": "changed"})

    for ident, prev in prev_map.items():
        if ident in curr_ids:
            continue
        for field, old_val in prev.items():
            if field == id_key:
                continue
            changes.append({"identifier": ident, "field": field, "old": old_val, "new": "", "change_type": "removed"})

    return changes


# ---------------------------
# Main
# ---------------------------

def main(argv: Optional[List[str]] = None) -> int:
    raw_args = list(argv) if argv is not None else sys.argv[1:]
    if not raw_args:
        show_start_instructions()
        raw_args = ["--gui"]

    parser = argparse.ArgumentParser(description="Denodo Extract (enhanced)")
    parser.add_argument("--config", help="Path to YAML config file")
    parser.add_argument("--log-level", default="INFO", help="Logging level (DEBUG, INFO, WARNING, ERROR)")
    parser.add_argument("--verify-ca", default="true",
                        help="TLS verification: 'true'/'false' or a path to a CA bundle file")
    parser.add_argument("--lookback-days", type=float, default=None,
                        help="Override: number of days back to collect modified rows (e.g., 2, 12).")
    parser.add_argument("--gui", action="store_true",
                        help="Run with GUI prompts (file pickers/message boxes).")
    out_group = parser.add_mutually_exclusive_group()
    out_group.add_argument("--output-dir", help="Override output directory (absolute or relative).")
    out_group.add_argument("--use-config-dir", action="store_true",
                           help="Write outputs next to the YAML config file.")
    args = parser.parse_args(raw_args)

    if args.gui:
        cfg_path = Path(args.config) if args.config else choose_config_file()
    else:
        if not args.config:
            print("Error: --config is required in CLI mode (omit to use --gui).", file=sys.stderr)
            return 2
        cfg_path = Path(args.config)
    cfg_path = cfg_path.resolve()
    cfg = validate_config(load_config(cfg_path))

    # Resolve directories
    storage_cfg = cfg["STORAGE"]["LOCAL_SYNC"]
    root_dir = Path(storage_cfg["ROOT_DIR"])
    paths_cfg = cfg["PATHS"]
    default_output_dir = Path(paths_cfg["OUTPUT_DIR"])
    if not default_output_dir.is_absolute():
        default_output_dir = root_dir / default_output_dir

    if args.use_config_dir:
        output_dir = cfg_path.parent
    elif args.output_dir:
        output_dir = Path(args.output_dir)
        if not output_dir.is_absolute():
            output_dir = output_dir.resolve()
    elif args.gui:
        picked = choose_output_dir(default_output_dir)
        if picked:
            output_dir = picked
        else:
            print("No output directory selected. Exiting.")
            return 1
    else:
        output_dir = default_output_dir

    if not output_dir.is_absolute():
        output_dir = output_dir.resolve()

    log_dir_cfg = paths_cfg.get("LOG_DIR")
    if log_dir_cfg:
        log_dir = Path(log_dir_cfg)
        if not log_dir.is_absolute():
            log_dir = output_dir / log_dir
    else:
        log_dir = output_dir / "logs"

    subset_dir_cfg = paths_cfg["SUBSET_DIR"]
    subset_dir = Path(subset_dir_cfg)
    if not subset_dir.is_absolute():
        subset_dir = output_dir / subset_dir

    # Initialize logging
    setup_logging(log_dir, level=args.log_level)

    # Filenames
    files_cfg = cfg["FILES"]
    json_base = files_cfg["EAMS_RESULT_DATA"]
    csv_base = files_cfg["EAMS_CSV_OUTPUT"]
    run_date_cfg = files_cfg["EAMS_RUN_DATE"]
    xlsx_base = files_cfg.get("EAMS_XLSX_OUTPUT") or (Path(csv_base).with_suffix(".xlsx").name)
    data_source_cfg = cfg.get("DATA_SOURCE") or cfg.get("DATASOURCE") or {}
    url = data_source_cfg.get("URL") or files_cfg.get("URL")

    run_date_path = Path(run_date_cfg)
    if not run_date_path.is_absolute():
        # Keep run-date alongside other outputs
        run_date_path = output_dir / run_date_path

    # Output dirs
    ensure_output_dirs(output_dir)
    ensure_output_dirs(subset_dir)

    run_at = datetime.now()
    ts = run_at.strftime("%Y%m%d_%H%M%S")
    json_name = timestamped_name(json_base, ts)
    csv_name = timestamped_name(csv_base, ts)
    xlsx_name = timestamped_name(xlsx_base, ts)
    contact_base = ensure_named_file(files_cfg["EAMS_CONTACT_INFO"], ".csv")
    server_base = ensure_named_file(files_cfg["EAMS_SERVER_INFO"], ".csv")

    json_path = output_dir / json_name
    csv_path = output_dir / csv_name
    xlsx_path = output_dir / xlsx_name
    contact_csv_path = subset_dir / timestamped_name(contact_base, ts)
    server_csv_path = subset_dir / timestamped_name(server_base, ts)
    contact_json_path = subset_dir / timestamped_name(Path(contact_base).with_suffix(".json").name, ts)
    contact_xlsx_path = subset_dir / timestamped_name(Path(contact_base).with_suffix(".xlsx").name, ts)
    server_json_path = subset_dir / timestamped_name(Path(server_base).with_suffix(".json").name, ts)
    server_xlsx_path = subset_dir / timestamped_name(Path(server_base).with_suffix(".xlsx").name, ts)
    changes_json_path = output_dir / timestamped_name(marker_name(json_base, "changes"), ts)
    changes_csv_path = output_dir / timestamped_name(marker_name(csv_base, "changes"), ts)
    changes_xlsx_path = output_dir / timestamped_name(marker_name(xlsx_base, "changes"), ts)

    subsets_cfg = cfg["SUBSETS"]
    contact_fields = subsets_cfg["CONTACT_INFO"]
    service_fields = subsets_cfg["SERVICE_IPS"]
    splunk_fields = subsets_cfg.get("SPLUNK_SERVER_FEED")

    splunk_base = None
    splunk_json_path = None
    splunk_csv_path = None
    splunk_xlsx_path = None
    if splunk_fields:
        splunk_base = ensure_named_file(files_cfg.get("EAMS_SPLUNK_SERVER_FEED", "splunk_server_feed.csv"), ".csv")
        splunk_json_path = subset_dir / timestamped_name(Path(splunk_base).with_suffix(".json").name, ts)
        splunk_csv_path = subset_dir / timestamped_name(splunk_base, ts)
        splunk_xlsx_path = subset_dir / timestamped_name(Path(splunk_base).with_suffix(".xlsx").name, ts)

    auth = cfg["AUTHENTICATION"]
    service = auth["SERVICE_NAME"]
    username = auth["USERNAME"]
    if keyring is None:
        logging.error("The 'keyring' package is unavailable. Install it or provide another secret mechanism.")
        return 1
    password = keyring.get_password(service, username)
    if not password:
        logging.error("No password found in keyring for service=%s user=%s", service, username)
        logging.info("Set it via: python -c \"import keyring; keyring.set_password('%s','%s','YOUR_SECRET')\"", service, username)
        return 1

    session = create_session()
    verify_ca = interpret_verify_ca(args.verify_ca)
    previous_run_at = parse_run_date(run_date_path)
    modified_anchor: Optional[datetime] = None

    lookback_days: Optional[float] = args.lookback_days
    if lookback_days is None and args.gui:
        lookback_days = prompt_lookback_days()

    if lookback_days is not None:
        modified_anchor = run_at - timedelta(days=lookback_days)
        logging.info("Using lookback window: %s days (since %s)", lookback_days, modified_anchor.isoformat())
    elif previous_run_at:
        modified_anchor = previous_run_at
        logging.info("Using previous run date for modified rows: %s", modified_anchor.isoformat())

    previous_json_path = find_latest_snapshot(output_dir, json_base)
    previous_rows: List[Dict[str, Any]] = []
    if previous_json_path:
        logging.info("Previous snapshot detected for diff: %s", previous_json_path.as_posix())
        previous_rows = load_json_rows(previous_json_path)
    else:
        logging.info("No prior snapshot found for %s; skipping change diff.", json_base)

    try:
        html = fetch_page(session, url, username, password, verify_ca)
        rows = extract_json_data(html)
        aggregated_rows = aggregate_by_identifier(rows, key="identifier")
        logging.info("Aggregated %d source rows into %d identifier rows", len(rows), len(aggregated_rows))

        contact_subset_rows = subset_rows(aggregated_rows, contact_fields)
        service_subset_rows = subset_rows(aggregated_rows, service_fields)
        splunk_subset_rows = subset_rows(aggregated_rows, splunk_fields) if splunk_fields else None

        write_json_snapshot(aggregated_rows, json_path)
        write_csv(aggregated_rows, csv_path)
        write_excel_table(aggregated_rows, xlsx_path, table_name="DenodoData")
        write_json_snapshot(contact_subset_rows, contact_json_path)
        write_csv_with_headers(aggregated_rows, contact_fields, contact_csv_path)
        write_excel_table_with_headers(contact_subset_rows, contact_fields, contact_xlsx_path, table_name="ContactInfo")
        write_json_snapshot(service_subset_rows, server_json_path)
        write_csv_with_headers(aggregated_rows, service_fields, server_csv_path)
        write_excel_table_with_headers(service_subset_rows, service_fields, server_xlsx_path, table_name="ServiceIps")
        if splunk_subset_rows is not None:
            write_json_snapshot(splunk_subset_rows, splunk_json_path)
            write_csv_with_headers(aggregated_rows, splunk_fields, splunk_csv_path)
            write_excel_table_with_headers(splunk_subset_rows, splunk_fields, splunk_xlsx_path, table_name="SplunkServerFeed")

        if previous_rows:
            changes_rows = diff_rows_by_identifier(previous_rows, aggregated_rows, id_key="identifier")
            if changes_rows:
                change_headers = ["identifier", "field", "old", "new", "change_type"]
                write_json_snapshot(changes_rows, changes_json_path)
                write_csv_with_headers(changes_rows, change_headers, changes_csv_path)
                write_excel_table_with_headers(changes_rows, change_headers, changes_xlsx_path, table_name="ChangedFields")
                logging.info("Detected %d changed fields versus previous snapshot.", len(changes_rows))
            else:
                logging.info("No field-level changes detected versus previous snapshot.")

        if modified_anchor:
            modified_rows = filter_modified_rows(rows, modified_anchor)
            if modified_rows:
                aggregated_modified_rows = aggregate_by_identifier(modified_rows, key="identifier")
                modified_json_name = timestamped_name(marker_name(json_base, "modified"), ts)
                modified_csv_name = timestamped_name(marker_name(csv_base, "modified"), ts)
                modified_xlsx_name = timestamped_name(marker_name(xlsx_base, "modified"), ts)
                write_json_snapshot(aggregated_modified_rows, output_dir / modified_json_name)
                write_csv(aggregated_modified_rows, output_dir / modified_csv_name)
                write_excel_table(aggregated_modified_rows, output_dir / modified_xlsx_name, table_name="DenodoDataModified")
                logging.info(
                    "Modified rows since %s written (%d rows source -> %d identifiers).",
                    modified_anchor.isoformat(),
                    len(modified_rows),
                    len(aggregated_modified_rows),
                )
            else:
                logging.info("No rows with lastmodifieddate newer than %s", modified_anchor.isoformat())
        write_run_date(run_date_path, run_at)
        logging.info("All done ✔")
        logging.info(
            "Outputs:\n"
            "  JSON: %s\n"
            "  CSV:  %s\n"
            "  XLSX: %s\n"
            "  Contact subset JSON: %s\n"
            "  Contact subset CSV:  %s\n"
            "  Contact subset XLSX: %s\n"
            "  Server subset JSON: %s\n"
            "  Server subset CSV:  %s\n"
            "  Server subset XLSX: %s",
            json_path,
            csv_path,
            xlsx_path,
            contact_json_path,
            contact_csv_path,
            contact_xlsx_path,
            server_json_path,
            server_csv_path,
            server_xlsx_path,
        )
        if splunk_subset_rows is not None:
            logging.info(
                "Splunk subset outputs:\n"
                "  Splunk subset JSON: %s\n"
                "  Splunk subset CSV:  %s\n"
                "  Splunk subset XLSX: %s",
                splunk_json_path,
                splunk_csv_path,
                splunk_xlsx_path,
            )
        if args.gui:
            try:
                import tkinter as tk
                from tkinter import messagebox
                root = tk.Tk()
                root.withdraw()
                message_lines = [
                    "Completed successfully.",
                    "",
                    "Outputs:",
                    f"JSON: {json_path}",
                    f"CSV: {csv_path}",
                    f"XLSX: {xlsx_path}",
                    f"Contact subset JSON: {contact_json_path}",
                    f"Contact subset CSV: {contact_csv_path}",
                    f"Contact subset XLSX: {contact_xlsx_path}",
                    f"Server subset JSON: {server_json_path}",
                    f"Server subset CSV: {server_csv_path}",
                    f"Server subset XLSX: {server_xlsx_path}",
                ]
                if splunk_subset_rows is not None:
                    message_lines.extend(
                        [
                            f"Splunk subset JSON: {splunk_json_path}",
                            f"Splunk subset CSV: {splunk_csv_path}",
                            f"Splunk subset XLSX: {splunk_xlsx_path}",
                        ]
                    )
                messagebox.showinfo(
                    "Denodo Extract",
                    "\n".join(message_lines),
                )
            except Exception:
                # GUI messaging is best-effort
                pass
        return 0
    except Exception as exc:
        logging.exception("Fatal error: %s", exc)
        if args.gui:
            try:
                import tkinter as tk
                from tkinter import messagebox
                root = tk.Tk()
                root.withdraw()
                messagebox.showerror("Denodo Extract - Error", str(exc))
            except Exception:
                pass
        return 1


if __name__ == "__main__":
    sys.exit(main())
