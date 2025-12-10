#!/usr/bin/env python3
"""
Denodo Extract v2 (enhanced)
----------------------------
Additions in this version:
- Structured logging: console + rotating file handler (logs roll @ 1 MB, keep last 3 files)
- Config validation: clear, actionable errors if required keys are missing/invalid
- Excel output: writes .xlsx with an actual Excel Table (styled) in addition to CSV
- TLS control: --verify-ca supports true/false or a CA bundle path
- Output placement: --output-dir to pick a folder, or --use-config-dir to write next to the YAML
- Dual CLI/GUI: run headless with CLI flags or interactively with --gui prompts
- No-arg start: if launched without flags, you’ll see a brief instruction prompt and then GUI mode

Usage:
    python Denodo_Extract_v2_enhanced.py --config /path/to/config.yml --verify-ca true
    # or provide a CA bundle path:
    python Denodo_Extract_v2_enhanced.py --config /path/to/config.yml --verify-ca C:/corp/ca-bundle.pem
    # choose where outputs land:
    python Denodo_Extract_v2_enhanced.py --config /path/to/config.yml --output-dir D:/Exports
    # GUI mode (pick config/output via dialogs):
    python Denodo_Extract_v2_enhanced.py --gui

YAML expectations (minimal):
DATALOC:
  ROOT_DIR: C:\\Data\\Denodo
  OUTPUT_DIR: exports
  # LOG_DIR: optional, defaults to <ROOT_DIR>/<OUTPUT_DIR>/logs

DATA:
  EAMS_RESULT_DATA: denodo_raw.json
  EAMS_CSV_OUTPUT: denodo_export.csv
  EAMS_RUN_DATE: last_run.txt
  # EAMS_XLSX_OUTPUT: optional (defaults to same stem as CSV + .xlsx)
  URL: https://denodo.example.local/report123

AUTHENTICATION:
  SERVICE_NAME: denodo-prod
  USERNAME: user
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


def normalize_cell_value(val) -> str:
    if val is None:
        s = ""
    elif isinstance(val, (list, tuple, set)):
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
    Configure root logger with both console and rotating file handler.
    Creates the log directory if missing.
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
    Prompt the user to select a YAML config if --config is not provided.
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
    Prompt the user to select an output directory. Returns None if cancelled.
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
    Present a brief how-to when starting without CLI flags.
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
    Prompt for lookback days via a simple Tkinter dialog. Returns None if cancelled/empty.
    """
    import tkinter as tk
    from tkinter import messagebox, simpledialog

    root = tk.Tk()
    root.withdraw()
    value = simpledialog.askstring("Lookback window", "Enter lookback in days (e.g., 2 or 12). Leave blank to skip:")
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
    Load YAML into a dictionary. Raise ConfigError on parse issues.
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
    Validate presence and basic types of required keys.
    Returns the (possibly updated) config dict.
    Raises ConfigError with actionable messages on failure.
    """
    missing = []
    # Required sections
    for sect in ("DATALOC", "DATA", "AUTHENTICATION"):
        if sect not in cfg or not isinstance(cfg[sect], dict):
            missing.append(sect)

    if missing:
        raise ConfigError(f"Missing or invalid sections in YAML: {', '.join(missing)}")

    # DATALOC
    dataloc = cfg["DATALOC"]
    for key in ("ROOT_DIR", "OUTPUT_DIR"):
        if key not in dataloc or not isinstance(dataloc[key], str) or not dataloc[key].strip():
            raise ConfigError(f"DATALOC.{key} is required and must be a non-empty string")

    # DATA
    data = cfg["DATA"]
    for key in ("EAMS_RESULT_DATA", "EAMS_CSV_OUTPUT", "EAMS_RUN_DATE", "URL"):
        if key not in data or not isinstance(data[key], str) or not data[key].strip():
            raise ConfigError(f"DATA.{key} is required and must be a non-empty string")

    # AUTHENTICATION
    auth = cfg["AUTHENTICATION"]
    for key in ("SERVICE_NAME", "USERNAME"):
        if key not in auth or not isinstance(auth[key], str) or not auth[key].strip():
            raise ConfigError(f"AUTHENTICATION.{key} is required and must be a non-empty string")

    return cfg


# ---------------------------
# HTTP
# ---------------------------

def create_session() -> requests.Session:
    """
    Create a requests session with retry/backoff for transient errors.
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
    Interpret --verify-ca option:
    - None or 'true' (default)  -> True
    - 'false'                   -> False
    - otherwise                 -> treat as a file path and return the string as-is
    """
    if value is None:
        return True
    v = value.strip().lower()
    if v in ("true", "1", "yes"):
        return True
    if v in ("false", "0", "no"):
        return False
    return value  # presumed path to CA bundle


def fetch_page(session: requests.Session, url: str, username: str, password: str, verify_ca) -> str:
    """
    GET the target URL using Basic Auth and return the HTML text.
    Raises requests exceptions on error.
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

def extract_json_data(html: str):
    """
    Locate 'var data = [ {...}, ... ]' in the HTML and parse it into a list of dicts.
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
    Ensure output directories exist.
    """
    print(f"Ensuring output directory exists: {out_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)


def parse_run_date(path: Path) -> Optional[datetime]:
    """
    Read prior run timestamp from EAMS_RUN_DATE. Returns None if missing or invalid.
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
    Record the execution time to the configured run date file.
    """
    run_date_path.parent.mkdir(parents=True, exist_ok=True)
    stamp = when.isoformat(timespec="seconds")
    run_date_path.write_text(stamp + "\n", encoding="utf-8")
    logging.info("Run date recorded: %s -> %s", run_date_path.as_posix(), stamp)


def parse_last_modified(value: Any) -> Optional[datetime]:
    """
    Attempt to parse a lastmodifieddate field into a datetime.
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


def marker_name(base_name: str, marker: str) -> str:
    """
    Insert a marker before the file extension (or at end if none).
    """
    p = Path(base_name)
    return f"{p.stem}_{marker}{p.suffix or ''}"


def timestamped_name(base_name: str, ts: str) -> str:
    """
    Insert _YYYYMMDD_HHMMSS before file extension.
    """
    p = Path(base_name)
    stem = p.stem
    suffix = p.suffix or ""
    return f"{stem}_{ts}{suffix}"


def union_fieldnames(rows):
    """
    Return a stable, first-seen-order list of unioned keys across all dict rows.
    """
    fields = set()
    order = []
    for row in rows:
        for k in row.keys():
            if k not in fields:
                fields.add(k)
                order.append(k)
    return order


def write_json_snapshot(rows, out_path: Path) -> None:
    """
    Pretty-print JSON for human inspection.
    """
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    logging.info("JSON snapshot saved: %s", out_path.as_posix())


def write_csv(rows, out_path: Path) -> None:
    """
    Write CSV with unioned headers across all rows.
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


def write_excel_table(rows, out_path: Path, table_name: str = "DenodoData") -> None:
    """
    Write an .xlsx file with a styled Excel Table containing the rows.
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


# ---------------------------
# Main
# ---------------------------

def main(argv=None) -> int:
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
    dataloc_cfg = cfg["DATALOC"]
    root_dir = Path(dataloc_cfg["ROOT_DIR"])
    default_output_dir = root_dir / dataloc_cfg["OUTPUT_DIR"]

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

    log_dir_cfg = dataloc_cfg.get("LOG_DIR")
    if log_dir_cfg:
        log_dir = Path(log_dir_cfg)
        if not log_dir.is_absolute():
            log_dir = output_dir / log_dir
    else:
        log_dir = output_dir / "logs"

    # Initialize logging
    setup_logging(log_dir, level=args.log_level)

    # Filenames
    data_cfg = cfg["DATA"]
    json_base = data_cfg["EAMS_RESULT_DATA"]
    csv_base = data_cfg["EAMS_CSV_OUTPUT"]
    run_date_cfg = data_cfg["EAMS_RUN_DATE"]
    xlsx_base = data_cfg.get("EAMS_XLSX_OUTPUT") or (Path(csv_base).with_suffix(".xlsx").name)
    url = data_cfg["URL"]

    run_date_path = Path(run_date_cfg)
    if not run_date_path.is_absolute():
        # Keep run-date alongside other outputs
        run_date_path = output_dir / run_date_path

    # Output dirs
    ensure_output_dirs(output_dir)

    run_at = datetime.now()
    ts = run_at.strftime("%Y%m%d_%H%M%S")
    json_name = timestamped_name(json_base, ts)
    csv_name = timestamped_name(csv_base, ts)
    xlsx_name = timestamped_name(xlsx_base, ts)

    json_path = output_dir / json_name
    csv_path = output_dir / csv_name
    xlsx_path = output_dir / xlsx_name

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

    try:
        html = fetch_page(session, url, username, password, verify_ca)
        rows = extract_json_data(html)
        write_json_snapshot(rows, json_path)
        write_csv(rows, csv_path)
        write_excel_table(rows, xlsx_path, table_name="DenodoData")
        if modified_anchor:
            modified_rows = filter_modified_rows(rows, modified_anchor)
            if modified_rows:
                modified_json_name = timestamped_name(marker_name(json_base, "modified"), ts)
                modified_csv_name = timestamped_name(marker_name(csv_base, "modified"), ts)
                modified_xlsx_name = timestamped_name(marker_name(xlsx_base, "modified"), ts)
                write_json_snapshot(modified_rows, output_dir / modified_json_name)
                write_csv(modified_rows, output_dir / modified_csv_name)
                write_excel_table(modified_rows, output_dir / modified_xlsx_name, table_name="DenodoDataModified")
                logging.info(
                    "Modified rows since %s written (%d rows).",
                    modified_anchor.isoformat(),
                    len(modified_rows),
                )
            else:
                logging.info("No rows with lastmodifieddate newer than %s", modified_anchor.isoformat())
        write_run_date(run_date_path, run_at)
        logging.info("All done ✔")
        logging.info("Outputs:\n  JSON: %s\n  CSV:  %s\n  XLSX: %s", json_path, csv_path, xlsx_path)
        if args.gui:
            try:
                import tkinter as tk
                from tkinter import messagebox
                root = tk.Tk()
                root.withdraw()
                messagebox.showinfo(
                    "Denodo Extract",
                    f"Completed successfully.\n\nOutputs:\nJSON: {json_path}\nCSV: {csv_path}\nXLSX: {xlsx_path}",
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
