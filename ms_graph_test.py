#!/usr/bin/env python3
"""
SharePoint/Teams folder connectivity test (Microsoft Graph)
-----------------------------------------------------------
- Authenticates via client credentials (app-only) using MSAL
- Resolves a SharePoint site by hostname + site path
- Resolves a folder (Teams channel folder or any folder) by drive path
- Prints and optionally opens the folder's SharePoint webUrl

Install:
  pip install msal requests pyyaml
Run:
  python sp_connect_test.py --config MS_Graph_Test.yaml
"""

from __future__ import annotations

import argparse
import sys
import webbrowser
from dataclasses import dataclass

import requests
import yaml
import msal


GRAPH_BASE = "https://graph.microsoft.com/v1.0"


@dataclass(frozen=True)
class SPConfig:
    tenant_id: str
    client_id: str
    client_secret: str
    tenant_hostname: str
    site_path: str
    folder_path: str
    open_in_browser: bool = False


def load_config(path: str) -> SPConfig:
    with open(path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    missing = [k for k in ["TENANT_ID", "CLIENT_ID", "CLIENT_SECRET", "TENANT_HOSTNAME", "SITE_PATH", "FOLDER_PATH"]
               if not raw.get(k)]
    if missing:
        raise ValueError(f"Missing required config keys: {', '.join(missing)}")

    return SPConfig(
        tenant_id=str(raw["TENANT_ID"]).strip(),
        client_id=str(raw["CLIENT_ID"]).strip(),
        client_secret=str(raw["CLIENT_SECRET"]).strip(),
        tenant_hostname=str(raw["TENANT_HOSTNAME"]).strip(),
        site_path=str(raw["SITE_PATH"]).strip(),
        folder_path=str(raw["FOLDER_PATH"]).strip().lstrip("/"),
        open_in_browser=bool(raw.get("OPEN_IN_BROWSER", False)),
    )


def get_access_token(cfg: SPConfig) -> str:
    authority = f"https://login.microsoftonline.com/{cfg.tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id=cfg.client_id,
        authority=authority,
        client_credential=cfg.client_secret,
    )

    # For app-only Graph:
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"Token acquisition failed: {result.get('error')} - {result.get('error_description')}")
    return result["access_token"]


def graph_get(url: str, token: str, *, params: dict | None = None) -> dict:
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, params=params, timeout=60)
    if not resp.ok:
        raise RuntimeError(f"Graph GET failed {resp.status_code}: {resp.text}")
    return resp.json()


def resolve_site_id(cfg: SPConfig, token: str) -> str:
    # GET /sites/{hostname}:/sites/{site-path}
    url = f"{GRAPH_BASE}/sites/{cfg.tenant_hostname}:{cfg.site_path}"
    data = graph_get(url, token)
    return data["id"]


def resolve_drive_id(site_id: str, token: str) -> str:
    # GET /sites/{site-id}/drive (default document library)
    url = f"{GRAPH_BASE}/sites/{site_id}/drive"
    data = graph_get(url, token)
    return data["id"]


def resolve_folder_item(drive_id: str, folder_path: str, token: str) -> dict:
    # GET /drives/{drive-id}/root:/{path}
    # Important: keep spaces; requests will URL-encode them.
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder_path}"
    # Ask for webUrl as well (usually included, but expand/select makes intent clear)
    params = {"$select": "id,name,webUrl,parentReference"}
    return graph_get(url, token, params=params)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", required=True, help="Path to YAML config (MS_Graph_Test.yaml)")
    args = ap.parse_args()

    cfg = load_config(args.config)

    print("1) Getting access token...")
    token = get_access_token(cfg)
    print("   OK")

    print("2) Resolving SharePoint site...")
    site_id = resolve_site_id(cfg, token)
    print(f"   site_id: {site_id}")

    print("3) Resolving default document library (drive)...")
    drive_id = resolve_drive_id(site_id, token)
    print(f"   drive_id: {drive_id}")

    print("4) Resolving folder item...")
    folder_item = resolve_folder_item(drive_id, cfg.folder_path, token)

    print("\n=== CONNECTED SUCCESSFULLY ===")
    print(f"Folder name: {folder_item.get('name')}")
    print(f"Folder itemId: {folder_item.get('id')}")
    print(f"Web URL (SharePoint view): {folder_item.get('webUrl')}")

    # This opens a browser only when you run it in an environment with a GUI.
    if cfg.open_in_browser:
        web_url = folder_item.get("webUrl")
        if web_url:
            print("\nOpening browser to folder webUrl...")
            webbrowser.open(web_url, new=2)

    print("\nNext step: store drive_id + folder_itemId for stable automation addressing.")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as e:
        print(f"\nERROR: {e}", file=sys.stderr)
        raise
