#!/usr/bin/env python3
"""
Denodo Extract + SharePoint (Microsoft Graph) - merged

This file merges the original Denodo_Extract.py pipeline with the SharePoint
connectivity logic proven in ms_graph_test.py.

Highlights
- Storage backend: local_sync OR sharepoint (app-only Graph)
- SharePoint backend stages outputs locally, then uploads to SharePoint paths
- Optional connectivity test: --sp-test (prints resolved webUrl; can open browser)

Notes
- Prefer storing the Graph client secret in an environment variable and reference
  it via STORAGE.SHAREPOINT.CLIENT_SECRET_ENV in the YAML.
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
from logging.handlers import RotatingFileHandler
import os
import posixpath
import re
import sys
import tempfile
import time
import webbrowser

from datetime import datetime, timedelta, timezone
from pathlib import Path, PurePosixPath
from typing import Any, Dict, Iterable, List, Optional, Tuple
from urllib.parse import urlparse

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


# ---------------------------
# Helpers
# ---------------------------

_ILLEGAL_XML_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
_EXCEL_CELL_LIMIT = 32767
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def normalize_cell_value(val: Any) -> str:
    if val is None:
        s = ""
    elif isinstance(val, (list, tuple, set)):
        s = " | ".join("" if v is None else str(v) for v in val)
    elif isinstance(val, dict):
        try:
            s = json.dumps(val, ensure_ascii=False, separators=(",", ":"))
        except Exception:
            s = str(val)
    else:
        s = str(val)
    s = _ILLEGAL_XML_RE.sub("", s)
    return s[:_EXCEL_CELL_LIMIT - 1] + "…" if len(s) > _EXCEL_CELL_LIMIT else s


def parse_graph_datetime(value: str) -> Optional[datetime]:
    if not value:
        return None
    cleaned = value.replace("Z", "+00:00") if value.endswith("Z") else value
    try:
        return datetime.fromisoformat(cleaned)
    except ValueError:
        return None


def relative_posix_path(path: Path, base_dir: Path) -> str:
    try:
        return PurePosixPath(path.relative_to(base_dir)).as_posix()
    except ValueError:
        return path.name


# ---------------------------
# Logging
# ---------------------------

def setup_logging(log_dir: Path, level: str = "INFO") -> None:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "denodo_extract.log"

    level_value = getattr(logging, level.upper(), logging.INFO)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s",
                            datefmt="%Y-%m-%d %H:%M:%S")

    ch = logging.StreamHandler()
    ch.setLevel(level_value)
    ch.setFormatter(fmt)

    fh = RotatingFileHandler(log_path, maxBytes=1_000_000, backupCount=3, encoding="utf-8")
    fh.setLevel(level_value)
    fh.setFormatter(fmt)

    root = logging.getLogger()
    root.setLevel(level_value)
    root.handlers.clear()
    root.addHandler(ch)
    root.addHandler(fh)


# ---------------------------
# Config
# ---------------------------

class ConfigError(Exception):
    """Raised when the YAML configuration is missing required keys or invalid."""


def load_config(path: Path) -> Dict[str, Any]:
    try:
        with path.open("r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
    except Exception as exc:
        raise ConfigError(f"Failed to read YAML config: {exc}") from exc
    if not isinstance(cfg, dict):
        raise ConfigError("Top-level YAML must be a mapping (dictionary).")
    return cfg


def validate_config(cfg: Dict[str, Any]) -> Dict[str, Any]:
    missing = []
    for sect in ("STORAGE", "PATHS", "FILES", "AUTHENTICATION", "SUBSETS"):
        if sect not in cfg or not isinstance(cfg[sect], dict):
            missing.append(sect)
    if "DATA_SOURCE" not in cfg and "DATASOURCE" not in cfg:
        missing.append("DATA_SOURCE")

    if missing:
        raise ConfigError(f"Missing or invalid sections in YAML: {', '.join(missing)}")

    # STORAGE
    storage = cfg["STORAGE"]
    backend = str(storage.get("BACKEND", "local_sync")).strip().lower()
    if backend not in {"local_sync", "sharepoint"}:
        raise ConfigError("STORAGE.BACKEND must be 'local_sync' or 'sharepoint'")
    storage["BACKEND"] = backend

    if backend == "local_sync":
        if "LOCAL_SYNC" not in storage or not isinstance(storage["LOCAL_SYNC"], dict):
            raise ConfigError("STORAGE.LOCAL_SYNC is required and must be a mapping")
        local_sync = storage["LOCAL_SYNC"]
        if not isinstance(local_sync.get("ROOT_DIR"), str) or not local_sync["ROOT_DIR"].strip():
            raise ConfigError("STORAGE.LOCAL_SYNC.ROOT_DIR is required and must be a non-empty string")
    else:
        if "SHAREPOINT" not in storage or not isinstance(storage["SHAREPOINT"], dict):
            raise ConfigError("STORAGE.SHAREPOINT is required and must be a mapping when BACKEND=sharepoint")
        sp = storage["SHAREPOINT"]
        # Required for SharePoint backend
        for key in ("SITE_URL", "DOCUMENT_LIBRARY"):
            if not isinstance(sp.get(key), str) or not sp[key].strip():
                raise ConfigError(f"STORAGE.SHAREPOINT.{key} is required and must be a non-empty string")
        # Optional but validated
        for key in ("BASE_FOLDER", "TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "CLIENT_SECRET_ENV", "OPEN_IN_BROWSER"):
            if key in sp and sp[key] is not None and key != "OPEN_IN_BROWSER" and not isinstance(sp[key], str):
                raise ConfigError(f"STORAGE.SHAREPOINT.{key} must be a string if provided")

    # PATHS
    paths = cfg["PATHS"]
    for key in ("OUTPUT_DIR", "SUBSET_DIR"):
        if not isinstance(paths.get(key), str) or not paths[key].strip():
            raise ConfigError(f"PATHS.{key} is required and must be a non-empty string")
    if "LOG_DIR" in paths and (not isinstance(paths["LOG_DIR"], str) or not paths["LOG_DIR"].strip()):
        raise ConfigError("PATHS.LOG_DIR must be a non-empty string if provided")

    # FILES
    files = cfg["FILES"]
    data_source = cfg.get("DATA_SOURCE") or cfg.get("DATASOURCE")
    if not isinstance(data_source, dict):
        raise ConfigError("DATA_SOURCE (or legacy DATASOURCE) must be a mapping")
    cfg["DATA_SOURCE"] = data_source

    for key in ("EAMS_RESULT_DATA", "EAMS_CSV_OUTPUT", "EAMS_RUN_DATE", "EAMS_SERVER_INFO", "EAMS_CONTACT_INFO"):
        if not isinstance(files.get(key), str) or not files[key].strip():
            raise ConfigError(f"FILES.{key} is required and must be a non-empty string")

    url_value = data_source.get("URL")
    if not isinstance(url_value, str) or not url_value.strip():
        legacy_url = files.get("URL")
        if isinstance(legacy_url, str) and legacy_url.strip():
            data_source["URL"] = legacy_url
        else:
            raise ConfigError("DATA_SOURCE.URL is required and must be a non-empty string")

    for k, v in files.items():
        if not isinstance(v, str):
            raise ConfigError(f"FILES.{k} must be a string")

    # AUTHENTICATION
    auth = cfg["AUTHENTICATION"]
    for key in ("SERVICE_NAME", "USERNAME"):
        if not isinstance(auth.get(key), str) or not auth[key].strip():
            raise ConfigError(f"AUTHENTICATION.{key} is required and must be a non-empty string")

    # SUBSETS
    subsets = cfg["SUBSETS"]
    if not isinstance(subsets.get("CONTACT_INFO"), list) or not subsets["CONTACT_INFO"]:
        raise ConfigError("SUBSETS.CONTACT_INFO must be a non-empty list")
    if not isinstance(subsets.get("SERVICE_IPS"), list) or not subsets["SERVICE_IPS"]:
        raise ConfigError("SUBSETS.SERVICE_IPS must be a non-empty list")
    for key, values in subsets.items():
        if not isinstance(values, list):
            raise ConfigError(f"SUBSETS.{key} must be a list of field names")
        for v in values:
            if not isinstance(v, str) or not v.strip():
                raise ConfigError(f"SUBSETS.{key} entries must be non-empty strings")

    return cfg


def load_graph_credentials(cfg: Dict[str, Any]) -> Tuple[str, str, str]:
    """
    Resolve Microsoft Graph credentials for app-only access.

    Order of preference:
    1) STORAGE.SHAREPOINT.TENANT_ID / CLIENT_ID plus secret from:
       - STORAGE.SHAREPOINT.CLIENT_SECRET_ENV (env var), else
       - STORAGE.SHAREPOINT.CLIENT_SECRET (inline), else
       - GRAPH_CLIENT_SECRET env var
    2) GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET env vars
    """
    sp = cfg["STORAGE"]["SHAREPOINT"]
    tenant_id = str(sp.get("TENANT_ID") or os.getenv("GRAPH_TENANT_ID", "")).strip()
    client_id = str(sp.get("CLIENT_ID") or os.getenv("GRAPH_CLIENT_ID", "")).strip()

    secret = ""
    if sp.get("CLIENT_SECRET_ENV"):
        secret = os.getenv(str(sp.get("CLIENT_SECRET_ENV")), "").strip()
    if not secret and sp.get("CLIENT_SECRET"):
        secret = str(sp.get("CLIENT_SECRET")).strip()
    if not secret:
        secret = os.getenv("GRAPH_CLIENT_SECRET", "").strip()

    if not tenant_id or not client_id or not secret:
        raise ConfigError(
            "SharePoint backend requires Graph credentials. Provide TENANT_ID and CLIENT_ID, and "
            "either CLIENT_SECRET_ENV (preferred) or CLIENT_SECRET, or set env vars "
            "GRAPH_TENANT_ID / GRAPH_CLIENT_ID / GRAPH_CLIENT_SECRET."
        )
    return tenant_id, client_id, secret


# ---------------------------
# SharePoint / Graph client
# ---------------------------

class SharePointClient:
    def __init__(
        self,
        *,
        site_url: str,
        document_library: str,
        base_folder: str,
        tenant_id: str,
        client_id: str,
        client_secret: str,
    ) -> None:
        self.site_url = site_url
        self.document_library = document_library
        self.base_folder = base_folder.strip().strip("/")
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret

        self.session = requests.Session()
        self._access_token: Optional[str] = None
        self._token_expires_at = 0.0
        self._site_id_cache: Optional[str] = None
        self._drive_id_cache: Optional[str] = None

    def _token(self) -> str:
        if self._access_token and time.time() < self._token_expires_at:
            return self._access_token

        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"
        data = {
            "client_id": self.client_id,
            "client_secret": self.client_secret,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        }
        resp = self.session.post(token_url, data=data, timeout=30)
        resp.raise_for_status()
        payload = resp.json()
        self._access_token = payload["access_token"]
        expires_in = int(payload.get("expires_in", 3600))
        self._token_expires_at = time.time() + max(expires_in - 60, 0)
        return self._access_token

    def _headers(self) -> Dict[str, str]:
        return {"Authorization": f"Bearer {self._token()}"}

    def resolve_site_id(self) -> str:
        if self._site_id_cache:
            return self._site_id_cache

        parsed = urlparse(self.site_url)
        if not parsed.netloc or not parsed.path.strip("/"):
            raise ConfigError("STORAGE.SHAREPOINT.SITE_URL must include a site path (e.g., /sites/EOC).")

        url = f"{GRAPH_BASE}/sites/{parsed.netloc}:{parsed.path}"
        resp = self.session.get(url, headers=self._headers(), timeout=30)
        resp.raise_for_status()
        self._site_id_cache = resp.json()["id"]
        return self._site_id_cache

    def resolve_drive_id(self) -> str:
        if self._drive_id_cache:
            return self._drive_id_cache

        site_id = self.resolve_site_id()
        url = f"{GRAPH_BASE}/sites/{site_id}/drives"
        resp = self.session.get(url, headers=self._headers(), timeout=30)
        resp.raise_for_status()
        for drive in resp.json().get("value", []):
            if drive.get("name") == self.document_library:
                self._drive_id_cache = drive["id"]
                return self._drive_id_cache
        raise ConfigError(f"Document library '{self.document_library}' not found in site {self.site_url}.")

    def get_item(self, remote_path: str) -> Dict[str, Any]:
        remote_path = remote_path.strip().strip("/")
        drive_id = self.resolve_drive_id()
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_path}"
        params = {"$select": "id,name,webUrl,lastModifiedDateTime,parentReference"}
        resp = self.session.get(url, headers=self._headers(), params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def item_exists(self, remote_path: str) -> bool:
        remote_path = remote_path.strip().strip("/")
        drive_id = self.resolve_drive_id()
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_path}"
        resp = self.session.get(url, headers=self._headers(), timeout=30)
        if resp.status_code == 404:
            return False
        resp.raise_for_status()
        return True

    def ensure_folder_path(self, folder_path: str) -> None:
        folder_path = folder_path.strip().strip("/")
        if not folder_path:
            return
        drive_id = self.resolve_drive_id()
        current = ""
        for segment in folder_path.split("/"):
            current = f"{current}/{segment}" if current else segment
            if self.item_exists(current):
                continue
            parent = posixpath.dirname(current)
            if parent:
                url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{parent}:/children"
            else:
                url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"
            payload = {"name": segment, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"}
            resp = self.session.post(url, headers=self._headers(), json=payload, timeout=30)
            if resp.status_code == 409:
                continue
            resp.raise_for_status()

    def list_children(self, folder_path: str) -> List[Dict[str, Any]]:
        folder_path = folder_path.strip().strip("/")
        drive_id = self.resolve_drive_id()
        if folder_path:
            url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder_path}:/children"
        else:
            url = f"{GRAPH_BASE}/drives/{drive_id}/root/children"
        params = {"$select": "name,lastModifiedDateTime,folder,file,webUrl"}
        out: List[Dict[str, Any]] = []
        while url:
            resp = self.session.get(url, headers=self._headers(), params=params, timeout=30)
            resp.raise_for_status()
            payload = resp.json()
            out.extend(payload.get("value", []))
            url = payload.get("@odata.nextLink")
            params = None  # nextLink already has query string
        return out

    def download_file(self, remote_path: str, local_path: Path) -> bool:
        remote_path = remote_path.strip().strip("/")
        drive_id = self.resolve_drive_id()
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_path}:/content"
        resp = self.session.get(url, headers=self._headers(), timeout=60)
        if resp.status_code == 404:
            return False
        resp.raise_for_status()
        local_path.parent.mkdir(parents=True, exist_ok=True)
        local_path.write_bytes(resp.content)
        return True

    def upload_file(self, local_path: Path, remote_path: str) -> None:
        remote_path = remote_path.strip().strip("/")
        folder = posixpath.dirname(remote_path)
        if folder:
            self.ensure_folder_path(folder)
        drive_id = self.resolve_drive_id()
        url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{remote_path}:/content"
        resp = self.session.put(url, headers=self._headers(), data=local_path.read_bytes(), timeout=120)
        resp.raise_for_status()


def find_latest_remote_snapshot(client: SharePointClient, remote_folder: str, base_name: str) -> Optional[str]:
    suffix = f"_{Path(base_name).name}"
    candidates: List[Tuple[datetime, str]] = []
    for item in client.list_children(remote_folder):
        name = item.get("name", "")
        if not name.endswith(suffix):
            continue
        modified = parse_graph_datetime(item.get("lastModifiedDateTime", "")) or datetime.min.replace(tzinfo=timezone.utc)
        candidates.append((modified, name))
    if not candidates:
        return None
    candidates.sort(key=lambda x: x[0])
    return posixpath.join(remote_folder, candidates[-1][1])


# ---------------------------
# HTTP / Denodo fetch
# ---------------------------

def create_session() -> requests.Session:
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
    if value is None:
        return True
    v = value.strip().lower()
    if v in ("true", "1", "yes"):
        return True
    if v in ("false", "0", "no"):
        return False
    return value


def fetch_page(session: requests.Session, url: str, username: str, password: str, verify_ca: str | bool) -> str:
    import urllib3
    if verify_ca is False:
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    logging.info("Fetching URL: %s (verify=%s)", url, verify_ca if verify_ca is not True else "True")
    resp = session.get(url, auth=(username, password), timeout=30, verify=verify_ca)
    resp.raise_for_status()
    return resp.text


def extract_json_data(html: str) -> List[Dict[str, Any]]:
    match = re.search(r"var\s+data\s*=\s*(\[\s*{.*?}\s*\])", html, re.DOTALL)
    if not match:
        raise ValueError("Expected JSON pattern not found in HTML (looked for 'var data = [...]').")
    payload = match.group(1)
    return json.loads(payload)


# ---------------------------
# Output helpers
# ---------------------------

def ensure_output_dirs(out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)


def union_fieldnames(rows: Iterable[Dict[str, Any]]) -> List[str]:
    fields = set()
    order: List[str] = []
    for row in rows:
        for k in row.keys():
            if k not in fields:
                fields.add(k)
                order.append(k)
    return order


def write_json_snapshot(rows: List[Dict[str, Any]], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(rows, f, ensure_ascii=False, indent=2)
    logging.info("JSON saved: %s", out_path.as_posix())


def write_csv(rows: List[Dict[str, Any]], out_path: Path) -> None:
    if not rows:
        logging.warning("No rows to write to CSV: %s", out_path.as_posix())
        return
    out_path.parent.mkdir(parents=True, exist_ok=True)
    headers = union_fieldnames(rows)
    with out_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        for r in rows:
            writer.writerow(r)
    logging.info("CSV saved: %s", out_path.as_posix())


def write_csv_with_headers(rows: List[Dict[str, Any]], headers: List[str], out_path: Path) -> None:
    if not rows:
        logging.warning("No rows to write to CSV: %s", out_path.as_posix())
        return
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        for r in rows:
            row_out = {h: r.get(h, "") for h in headers}
            writer.writerow(row_out)
    logging.info("CSV saved: %s", out_path.as_posix())


def write_excel_table_with_headers(rows: List[Dict[str, Any]], headers: List[str], out_path: Path, table_name: str) -> None:
    if not rows:
        logging.warning("No rows to write to Excel: %s", out_path.as_posix())
        return

    out_path.parent.mkdir(parents=True, exist_ok=True)
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

    wb.save(out_path)
    logging.info("XLSX saved: %s", out_path.as_posix())


def write_excel_table(rows: List[Dict[str, Any]], out_path: Path, table_name: str = "DenodoData") -> None:
    headers = union_fieldnames(rows)
    write_excel_table_with_headers(rows, headers, out_path, table_name)


def subset_rows(rows: List[Dict[str, Any]], headers: List[str]) -> List[Dict[str, Any]]:
    return [{h: r.get(h, "") for h in headers} for r in rows]


def aggregate_by_identifier(rows: Iterable[Dict[str, Any]], key: str = "identifier") -> List[Dict[str, Any]]:
    grouped: Dict[Any, Dict[str, Any]] = {}
    for row in rows:
        ident = row.get(key)
        bucket = grouped.setdefault(ident, {key: ident, "_seen": {}, "_data": {}})
        seen = bucket["_seen"]
        data = bucket["_data"]

        for field, val in row.items():
            if field == key or val is None:
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

    out: List[Dict[str, Any]] = []
    for bucket in grouped.values():
        row_out = {key: bucket.get(key)}
        for field, vals in bucket["_data"].items():
            row_out[field] = vals
        out.append(row_out)
    return out


def marker_name(base_name: str, marker: str) -> str:
    p = Path(base_name)
    return f"{p.stem}_{marker}{p.suffix or ''}"


def timestamped_name(base_name: str, ts: str) -> str:
    p = Path(base_name)
    return f"{ts}_{p.name}"


def ensure_named_file(base_name: str, default_suffix: str) -> str:
    p = Path(base_name)
    return p.name if p.suffix else f"{p.name}{default_suffix}"


def parse_run_date(path: Path) -> Optional[datetime]:
    if not path.exists():
        return None
    try:
        text = path.read_text(encoding="utf-8").strip()
        if not text:
            return None
        cleaned = text.replace("Z", "+00:00") if text.endswith("Z") else text
        return datetime.fromisoformat(cleaned)
    except Exception:
        return None


def write_run_date(run_date_path: Path, when: datetime) -> None:
    run_date_path.parent.mkdir(parents=True, exist_ok=True)
    stamp = when.isoformat(timespec="seconds")
    run_date_path.write_text(stamp + "\n", encoding="utf-8")
    logging.info("Run date recorded: %s -> %s", run_date_path.as_posix(), stamp)


def load_json_rows(path: Path) -> List[Dict[str, Any]]:
    try:
        with path.open("r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, list) else []
    except Exception:
        return []


def diff_rows_by_identifier(prev_rows: List[Dict[str, Any]], curr_rows: List[Dict[str, Any]], id_key: str = "identifier") -> List[Dict[str, Any]]:
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

        for field in set(prev.keys()) | set(curr.keys()):
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


def find_latest_snapshot(output_dir: Path, base_name: str) -> Optional[Path]:
    pattern = f"*_{Path(base_name).name}"
    candidates = sorted(p for p in output_dir.glob(pattern) if p.is_file())
    return candidates[-1] if candidates else None


# ---------------------------
# Main
# ---------------------------

def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Denodo Extract (merged SharePoint/Graph)")
    parser.add_argument("--config", required=True, help="Path to YAML config file")
    parser.add_argument("--log-level", default="INFO")
    parser.add_argument("--verify-ca", default="true", help="TLS verify: true/false or CA bundle path")
    parser.add_argument("--lookback-days", type=float, default=None)
    parser.add_argument("--output-dir", help="Override output directory")
    parser.add_argument("--use-config-dir", action="store_true", help="Write outputs next to the YAML")

    # SharePoint test options (from ms_graph_test.py)
    parser.add_argument("--sp-test", action="store_true", help="Only test SharePoint connectivity and exit")
    parser.add_argument("--open-sp", action="store_true", help="When used with --sp-test, open folder webUrl")

    args = parser.parse_args(argv if argv is not None else sys.argv[1:])

    cfg_path = Path(args.config).resolve()
    cfg = validate_config(load_config(cfg_path))

    storage_backend = cfg["STORAGE"]["BACKEND"]
    paths_cfg = cfg["PATHS"]
    files_cfg = cfg["FILES"]
    subsets_cfg = cfg["SUBSETS"]
    sharepoint_client: Optional[SharePointClient] = None
    remote_output_prefix = ""
    remote_subset_prefix = ""

    # Output dir resolution (local or staging)
    if storage_backend == "local_sync":
        root_dir = Path(cfg["STORAGE"]["LOCAL_SYNC"]["ROOT_DIR"])
        default_output_dir = Path(paths_cfg["OUTPUT_DIR"])
        if not default_output_dir.is_absolute():
            default_output_dir = root_dir / default_output_dir
    else:
        sp_cfg = cfg["STORAGE"]["SHAREPOINT"]
        tenant_id, client_id, client_secret = load_graph_credentials(cfg)
        sharepoint_client = SharePointClient(
            site_url=sp_cfg["SITE_URL"],
            document_library=sp_cfg["DOCUMENT_LIBRARY"],
            base_folder=sp_cfg.get("BASE_FOLDER", ""),
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=client_secret,
        )

        staging_root = Path(tempfile.mkdtemp(prefix="denodo_extract_"))
        default_output_dir = Path(paths_cfg["OUTPUT_DIR"])
        if not default_output_dir.is_absolute():
            default_output_dir = staging_root / default_output_dir

        base_folder = sp_cfg.get("BASE_FOLDER", "").strip().strip("/")
        output_suffix = PurePosixPath(paths_cfg["OUTPUT_DIR"]).as_posix().lstrip("/")
        subset_suffix = PurePosixPath(paths_cfg["SUBSET_DIR"]).as_posix().lstrip("/")
        remote_output_prefix = posixpath.join(base_folder, output_suffix) if base_folder else output_suffix
        remote_subset_prefix = posixpath.join(base_folder, subset_suffix) if base_folder else subset_suffix

    # CLI output override
    if args.use_config_dir:
        output_dir = cfg_path.parent
    elif args.output_dir:
        output_dir = Path(args.output_dir).resolve()
    else:
        output_dir = default_output_dir.resolve()

    # log dir
    log_dir_cfg = paths_cfg.get("LOG_DIR")
    if log_dir_cfg:
        log_dir = Path(log_dir_cfg)
        if not log_dir.is_absolute():
            log_dir = output_dir / log_dir
    else:
        log_dir = output_dir / "logs"

    # subset dir
    subset_dir = Path(paths_cfg["SUBSET_DIR"])
    if not subset_dir.is_absolute():
        subset_dir = output_dir / subset_dir

    setup_logging(log_dir, level=args.log_level)
    logging.info("Storage backend: %s", storage_backend)

    # Optional SharePoint connectivity test and exit
    if args.sp_test:
        if not sharepoint_client:
            logging.error("--sp-test requires STORAGE.BACKEND=sharepoint")
            return 2
        base_folder = cfg["STORAGE"]["SHAREPOINT"].get("BASE_FOLDER", "").strip().strip("/")
        # If BASE_FOLDER provided, test that folder; else test document library root
        test_path = base_folder or ""
        try:
            if test_path:
                item = sharepoint_client.get_item(test_path)
            else:
                # Root item is /root (special): use BASE_FOLDER absent -> list children
                item = {"name": sharepoint_client.document_library, "webUrl": None}
            logging.info("Resolved site_id: %s", sharepoint_client.resolve_site_id())
            logging.info("Resolved drive_id: %s", sharepoint_client.resolve_drive_id())
            if test_path:
                logging.info("Resolved folder: %s", test_path)
                logging.info("webUrl: %s", item.get("webUrl"))
                if (args.open_sp or bool(cfg["STORAGE"]["SHAREPOINT"].get("OPEN_IN_BROWSER"))) and item.get("webUrl"):
                    webbrowser.open(item["webUrl"], new=2)
            else:
                logging.info("BASE_FOLDER not set; SharePoint root access OK.")
            print("=== CONNECTED SUCCESSFULLY ===")
            return 0
        except Exception as exc:
            logging.exception("SharePoint test failed: %s", exc)
            return 1

    # Filenames
    json_base = files_cfg["EAMS_RESULT_DATA"]
    csv_base = files_cfg["EAMS_CSV_OUTPUT"]
    run_date_cfg = files_cfg["EAMS_RUN_DATE"]
    xlsx_base = files_cfg.get("EAMS_XLSX_OUTPUT") or (Path(csv_base).with_suffix(".xlsx").name)

    url = (cfg["DATA_SOURCE"].get("URL") or "").strip()
    if not url:
        raise ConfigError("DATA_SOURCE.URL is empty after validation.")

    run_date_path = Path(run_date_cfg)
    if not run_date_path.is_absolute():
        run_date_path = output_dir / run_date_path

    ensure_output_dirs(output_dir)
    ensure_output_dirs(subset_dir)

    # If SharePoint: pull down run-date and latest prior snapshot for diff
    if sharepoint_client and remote_output_prefix:
        remote_run_date_path = posixpath.join(remote_output_prefix, relative_posix_path(run_date_path, output_dir))
        try:
            if sharepoint_client.download_file(remote_run_date_path, run_date_path):
                logging.info("Downloaded run date from SharePoint: %s", remote_run_date_path)
        except Exception as exc:
            logging.warning("Could not download run date from SharePoint: %s", exc)

        remote_snapshot_path = find_latest_remote_snapshot(sharepoint_client, remote_output_prefix, json_base)
        if remote_snapshot_path:
            local_snapshot_path = output_dir / Path(remote_snapshot_path).name
            try:
                if sharepoint_client.download_file(remote_snapshot_path, local_snapshot_path):
                    logging.info("Downloaded prior snapshot for diff: %s", remote_snapshot_path)
            except Exception as exc:
                logging.warning("Could not download prior snapshot: %s", exc)

    previous_run_at = parse_run_date(run_date_path)
    lookback_days = args.lookback_days
    modified_anchor: Optional[datetime] = None
    run_at = datetime.now()
    if lookback_days is not None:
        modified_anchor = run_at - timedelta(days=lookback_days)
        logging.info("Using lookback window: %s days (since %s)", lookback_days, modified_anchor.isoformat())
    elif previous_run_at:
        modified_anchor = previous_run_at
        logging.info("Using previous run date for modified rows: %s", modified_anchor.isoformat())

    # Find local prior snapshot (either already local, or downloaded from SharePoint above)
    previous_json_path = find_latest_snapshot(output_dir, json_base)
    previous_rows: List[Dict[str, Any]] = load_json_rows(previous_json_path) if previous_json_path else []

    # Credentials for Denodo
    auth = cfg["AUTHENTICATION"]
    if keyring is None:
        logging.error("Missing dependency 'keyring'. Install it (pip install keyring) and ensure your secret is stored.")
        return 1
    password = keyring.get_password(auth["SERVICE_NAME"], auth["USERNAME"])
    if not password:
        logging.error("No password found in keyring for service=%s user=%s", auth["SERVICE_NAME"], auth["USERNAME"])
        return 1

    # Run
    ts = run_at.strftime("%Y%m%d_%H%M%S")
    json_path = output_dir / timestamped_name(json_base, ts)
    csv_path = output_dir / timestamped_name(csv_base, ts)
    xlsx_path = output_dir / timestamped_name(xlsx_base, ts)

    contact_fields = subsets_cfg["CONTACT_INFO"]
    service_fields = subsets_cfg["SERVICE_IPS"]
    splunk_fields = subsets_cfg.get("SPLUNK_SERVER_FEED")

    contact_base = ensure_named_file(files_cfg["EAMS_CONTACT_INFO"], ".csv")
    server_base = ensure_named_file(files_cfg["EAMS_SERVER_INFO"], ".csv")

    contact_csv_path = subset_dir / timestamped_name(contact_base, ts)
    server_csv_path = subset_dir / timestamped_name(server_base, ts)
    contact_json_path = subset_dir / timestamped_name(Path(contact_base).with_suffix(".json").name, ts)
    server_json_path = subset_dir / timestamped_name(Path(server_base).with_suffix(".json").name, ts)
    contact_xlsx_path = subset_dir / timestamped_name(Path(contact_base).with_suffix(".xlsx").name, ts)
    server_xlsx_path = subset_dir / timestamped_name(Path(server_base).with_suffix(".xlsx").name, ts)

    changes_json_path = output_dir / timestamped_name(marker_name(json_base, "changes"), ts)
    changes_csv_path = output_dir / timestamped_name(marker_name(csv_base, "changes"), ts)
    changes_xlsx_path = output_dir / timestamped_name(marker_name(xlsx_base, "changes"), ts)

    session = create_session()
    verify_ca = interpret_verify_ca(args.verify_ca)

    try:
        html = fetch_page(session, url, auth["USERNAME"], password, verify_ca)
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
            splunk_base = ensure_named_file(files_cfg.get("EAMS_SPLUNK_SERVER_FEED", "splunk_server_feed.csv"), ".csv")
            splunk_csv_path = subset_dir / timestamped_name(splunk_base, ts)
            splunk_json_path = subset_dir / timestamped_name(Path(splunk_base).with_suffix(".json").name, ts)
            splunk_xlsx_path = subset_dir / timestamped_name(Path(splunk_base).with_suffix(".xlsx").name, ts)
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
                logging.info("Detected %d changed fields vs previous snapshot.", len(changes_rows))
            else:
                logging.info("No field-level changes detected vs previous snapshot.")

        # NOTE: The "modified rows" feature depends on having a parseable lastmodifieddate in the source.
        # Your prior script included a parse_last_modified() tailored to your data; if you still need this,
        # re-add it here and write _modified outputs similarly.

        write_run_date(run_date_path, run_at)

        # Upload to SharePoint
        if sharepoint_client and remote_output_prefix and remote_subset_prefix:
            uploads: List[Tuple[Path, str]] = []
            for local_file in [json_path, csv_path, xlsx_path, changes_json_path, changes_csv_path, changes_xlsx_path, run_date_path]:
                if local_file.exists():
                    rel_path = relative_posix_path(local_file, output_dir)
                    uploads.append((local_file, posixpath.join(remote_output_prefix, rel_path)))

            for local_file in [contact_csv_path, contact_json_path, contact_xlsx_path, server_csv_path, server_json_path, server_xlsx_path]:
                if local_file.exists():
                    rel_path = relative_posix_path(local_file, subset_dir)
                    uploads.append((local_file, posixpath.join(remote_subset_prefix, rel_path)))

            # include splunk files if created
            if splunk_fields:
                for p in subset_dir.glob(f"{ts}_splunk_server_feed.*"):
                    rel_path = relative_posix_path(p, subset_dir)
                    uploads.append((p, posixpath.join(remote_subset_prefix, rel_path)))

            for local_path, remote_path in uploads:
                try:
                    sharepoint_client.upload_file(local_path, remote_path)
                    logging.info("Uploaded to SharePoint: %s", remote_path)
                except Exception as exc:
                    logging.error("Upload failed (%s -> %s): %s", local_path, remote_path, exc)

        logging.info("All done ✔")
        return 0

    except Exception as exc:
        logging.exception("Fatal error: %s", exc)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
