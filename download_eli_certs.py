"""
Download the "ELI CERTS 2024 - 2026" folder from SharePoint via Microsoft Graph.

Source (decoded from the Teams/SharePoint link):
    https://aldingerco.sharepoint.com
    /sites/IntegrationTeam
    Shared Documents / ELI / Files / ELI CERTS 2024 - 2026

Auth: Microsoft device-code flow (MSAL public client) -- delegated, so it
downloads exactly what the signed-in user can see. Same tenant + app
registration the Holts pipeline already uses; only the site/folder differ.

A persistent token cache (.msal_token_cache.json next to this script) means you
only sign in once; later runs reuse the cached refresh token silently.

A manifest of every file found (_eli_manifest.json) is written after each scan,
so you can inspect/filter without re-hitting SharePoint.

Usage:
    python download_eli_certs.py --list-only          # scan + show a breakdown
    python download_eli_certs.py                       # download everything
    python download_eli_certs.py --ext pdf             # only .pdf files
    python download_eli_certs.py --ext pdf doc docx    # only these extensions
    python download_eli_certs.py --workers 16
    python download_eli_certs.py --output "D:\\dir"
"""

import os
import sys
import time
import json
import atexit
import argparse
import requests
import msal
from collections import defaultdict
from urllib.parse import quote
from concurrent.futures import ThreadPoolExecutor, as_completed
from threading import Lock

# Windows consoles default to cp1252, which chokes on some SharePoint filenames.
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

PRINT_LOCK = Lock()

# ======================== Configuration ========================
# Same tenant + app registration as the Holts pipeline (aldingerco tenant).
# These are NOT secrets: a tenant id and a public-client app id.
TENANT_ID = "87c497fe-0035-452c-9e7c-663075b2dd6b"
CLIENT_ID = "01e3d669-ef6e-4b03-8e5d-00ef699706ef"

SHAREPOINT_HOST = "aldingerco.sharepoint.com"
SITE_PATH = "/sites/IntegrationTeam"

# Path INSIDE the default "Documents" (Shared Documents) drive, drive-root relative.
FOLDER_PATH = "ELI/Files/ELI CERTS 2024 - 2026"

SCOPES = [
    "https://graph.microsoft.com/Files.Read.All",
    "https://graph.microsoft.com/Sites.Read.All",
]

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TOKEN_CACHE_PATH = os.path.join(SCRIPT_DIR, ".msal_token_cache.json")
MANIFEST_PATH = os.path.join(SCRIPT_DIR, "_eli_manifest.json")
DEFAULT_OUTPUT = os.path.join(SCRIPT_DIR, "ELI CERTS 2024 - 2026")


def _long(path):
    r"""Return a Windows long-path (\\?\) prefixed absolute path to dodge MAX_PATH."""
    p = os.path.abspath(path)
    if os.name == "nt" and not p.startswith("\\\\?\\"):
        return "\\\\?\\" + p
    return p


def human(n):
    n = float(n)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{n:.1f} {unit}"
        n /= 1024


# ======================== Token Management ========================

_app = None
_token = None
_token_time = 0
TOKEN_REFRESH_INTERVAL = 2700  # refresh every 45 min


def _build_token_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_PATH):
        try:
            cache.deserialize(open(TOKEN_CACHE_PATH, "r", encoding="utf-8").read())
        except Exception:
            pass

    def _save():
        if cache.has_state_changed:
            with open(TOKEN_CACHE_PATH, "w", encoding="utf-8") as f:
                f.write(cache.serialize())

    atexit.register(_save)
    return cache


def get_app():
    global _app
    if _app is None:
        _app = msal.PublicClientApplication(
            CLIENT_ID,
            authority=f"https://login.microsoftonline.com/{TENANT_ID}",
            token_cache=_build_token_cache(),
        )
    return _app


def authenticate(force_refresh=False):
    """Acquire a token: silent from cache if possible, else device-code flow."""
    global _token, _token_time

    if not force_refresh and _token and (time.time() - _token_time) < TOKEN_REFRESH_INTERVAL:
        return _token

    app = get_app()

    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            _token = result["access_token"]
            _token_time = time.time()
            return _token

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception(f"Failed to create device flow: {flow.get('error_description')}")

    print("\n" + "=" * 60, flush=True)
    print("ACTION REQUIRED: Sign in to Microsoft", flush=True)
    print("=" * 60, flush=True)
    print(f"\n1. Open your browser to: {flow['verification_uri']}", flush=True)
    print(f"2. Enter this code:      {flow['user_code']}", flush=True)
    print("\nWaiting for you to sign in...", flush=True)
    print("=" * 60 + "\n", flush=True)

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise Exception(f"Authentication failed: {result.get('error_description')}")

    _token = result["access_token"]
    _token_time = time.time()
    print("Signed in successfully.\n", flush=True)
    return _token


def get_token():
    return authenticate()


# ======================== Graph API ========================

def get_site_id(token):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOST}:{SITE_PATH}"
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.json()["id"]


def get_drive_id(token, site_id):
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    drives = r.json()["value"]
    for d in drives:
        if d["name"] == "Documents":
            return d["id"]
    return drives[0]["id"]


def get_folder_id_by_path(token, drive_id, folder_path):
    headers = {"Authorization": f"Bearer {token}"}
    encoded = quote(folder_path, safe="/")
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded}"
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.json()["id"]


def list_all_files_recursive(drive_id, item_id, relative_path="", retries=4):
    """Recursively list every file under a folder. Returns list of
    {name, path, size, id}. Handles pagination + 401 refresh + throttling."""
    files = []
    for attempt in range(1, retries + 1):
        try:
            token = get_token()
            headers = {"Authorization": f"Bearer {token}"}
            url = (f"https://graph.microsoft.com/v1.0/drives/{drive_id}"
                   f"/items/{item_id}/children?$top=200")

            while url:
                resp = requests.get(url, headers=headers)
                if resp.status_code == 401:
                    authenticate(force_refresh=True)
                    break
                if resp.status_code == 429:
                    wait = int(resp.headers.get("Retry-After", 10))
                    time.sleep(wait)
                    continue
                resp.raise_for_status()
                data = resp.json()

                for item in data.get("value", []):
                    if "file" in item:
                        fp = f"{relative_path}/{item['name']}" if relative_path else item["name"]
                        files.append({
                            "name": item["name"],
                            "path": fp,
                            "size": item["size"],
                            "id": item["id"],
                        })
                    elif "folder" in item:
                        sp = f"{relative_path}/{item['name']}" if relative_path else item["name"]
                        files.extend(list_all_files_recursive(drive_id, item["id"], sp))

                url = data.get("@odata.nextLink")
            else:
                return files
        except requests.HTTPError as e:
            code = e.response.status_code if e.response is not None else None
            if attempt < retries and code in (401, 429, 500, 503):
                time.sleep(5 * attempt)
            else:
                raise
    return files


# ======================== Download ========================

def download_file(drive_id, item_id, local_path, retries=4):
    os.makedirs(_long(os.path.dirname(local_path)), exist_ok=True)
    for attempt in range(1, retries + 1):
        try:
            token = get_token()
            headers = {"Authorization": f"Bearer {token}"}
            url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
            with requests.get(url, headers=headers, stream=True, timeout=180, allow_redirects=True) as r:
                if r.status_code == 401:
                    authenticate(force_refresh=True)
                    continue
                if r.status_code == 429:
                    time.sleep(int(r.headers.get("Retry-After", 10)))
                    continue
                r.raise_for_status()
                with open(_long(local_path), "wb") as f:
                    for chunk in r.iter_content(chunk_size=1024 * 1024):
                        if chunk:
                            f.write(chunk)
            return "ok"
        except Exception as e:
            if attempt < retries:
                time.sleep(5 * attempt)
            else:
                return f"error: {e}"
    return "error: max retries"


def process_file(drive_id, file_info, output_dir):
    local_path = os.path.join(output_dir, file_info["path"].replace("/", os.sep))
    lp = _long(local_path)
    if os.path.exists(lp) and os.path.getsize(lp) == file_info["size"]:
        return "skip"
    status = download_file(drive_id, file_info["id"], local_path)
    if status != "ok":
        with PRINT_LOCK:
            print(f"  x {file_info['path']}  [{status}]", flush=True)
    return status


# ======================== Manifest / breakdown ========================

def save_manifest(files):
    with open(MANIFEST_PATH, "w", encoding="utf-8") as f:
        json.dump(files, f, indent=2)


def print_breakdown(files):
    total = sum(f["size"] for f in files)
    print(f"\nTotal: {len(files)} files, {human(total)}\n", flush=True)

    by_ext = defaultdict(lambda: [0, 0])
    by_top = defaultdict(lambda: [0, 0])
    for f in files:
        ext = (os.path.splitext(f["name"])[1] or "(none)").lower()
        by_ext[ext][0] += 1
        by_ext[ext][1] += f["size"]
        top = f["path"].split("/", 1)[0] if "/" in f["path"] else "(root)"
        by_top[top][0] += 1
        by_top[top][1] += f["size"]

    print("By file type:", flush=True)
    for ext, (c, sz) in sorted(by_ext.items(), key=lambda kv: -kv[1][1]):
        print(f"  {ext:<10} {c:>7} files   {human(sz):>10}", flush=True)

    print("\nBy top-level subfolder (top 30 by size):", flush=True)
    rows = sorted(by_top.items(), key=lambda kv: -kv[1][1])[:30]
    for top, (c, sz) in rows:
        print(f"  {human(sz):>10}  {c:>7} files   {top}", flush=True)


# ======================== Main ========================

def main():
    ap = argparse.ArgumentParser(description="Download ELI CERTS 2024-2026 from SharePoint")
    ap.add_argument("--output", default=DEFAULT_OUTPUT, help="Local destination directory")
    ap.add_argument("--workers", type=int, default=12, help="Parallel download workers")
    ap.add_argument("--list-only", action="store_true", help="Scan + show breakdown; do not download")
    ap.add_argument("--ext", nargs="+", default=None,
                    help="Only download files with these extensions, e.g. --ext pdf doc")
    ap.add_argument("--use-manifest", action="store_true",
                    help="Reuse _eli_manifest.json instead of re-scanning SharePoint")
    args = ap.parse_args()

    output_dir = os.path.abspath(args.output)

    print("ELI CERTS 2024 - 2026", flush=True)
    print("=" * 60, flush=True)
    print(f"Site:   {SHAREPOINT_HOST}{SITE_PATH}", flush=True)
    print(f"Folder: {FOLDER_PATH}", flush=True)
    print(f"Output: {output_dir}", flush=True)
    print("=" * 60, flush=True)

    # Get file list -- from manifest or a fresh scan.
    if args.use_manifest and os.path.exists(MANIFEST_PATH):
        with open(MANIFEST_PATH, "r", encoding="utf-8") as f:
            files = json.load(f)
        print(f"\nUsing cached manifest: {len(files)} files.", flush=True)
    else:
        token = authenticate()
        print("\nResolving site / drive / folder...", flush=True)
        site_id = get_site_id(token)
        drive_id = get_drive_id(token, site_id)
        folder_id = get_folder_id_by_path(token, drive_id, FOLDER_PATH)
        print("Connected. Scanning folder (recursive)...", flush=True)
        files = list_all_files_recursive(drive_id, folder_id)
        save_manifest(files)
        print(f"Scan complete. Manifest saved to {MANIFEST_PATH}", flush=True)

    print_breakdown(files)

    # Apply extension filter if requested.
    if args.ext:
        wanted = {("." + e.lower().lstrip(".")) for e in args.ext}
        before = len(files)
        files = [f for f in files if os.path.splitext(f["name"])[1].lower() in wanted]
        print(f"\nFilter --ext {sorted(wanted)}: {len(files)} of {before} files selected, "
              f"{human(sum(f['size'] for f in files))}.", flush=True)

    if args.list_only:
        print("\n(list-only) Nothing downloaded.", flush=True)
        return

    if not files:
        print("\nNothing to download.", flush=True)
        return

    # Need a live drive_id for downloading (manifest path may have skipped resolve).
    if args.use_manifest:
        token = authenticate()
        site_id = get_site_id(token)
        drive_id = get_drive_id(token, site_id)

    os.makedirs(_long(output_dir), exist_ok=True)
    total = len(files)
    print(f"\nDownloading {total} files to {output_dir} ({args.workers} workers)...\n", flush=True)

    ok = skip = err = done = 0
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futures = [ex.submit(process_file, drive_id, fi, output_dir) for fi in files]
        for fut in as_completed(futures):
            try:
                s = fut.result()
            except Exception:
                s = "error"
            if s == "ok":
                ok += 1
            elif s == "skip":
                skip += 1
            else:
                err += 1
            done += 1
            if done % 250 == 0 or done == total:
                with PRINT_LOCK:
                    print(f"  [{done}/{total}]  downloaded={ok}  already-had={skip}  errors={err}", flush=True)

    print("\n" + "=" * 60, flush=True)
    print("DONE", flush=True)
    print(f"  downloaded:  {ok}", flush=True)
    print(f"  already had: {skip}", flush=True)
    print(f"  errors:      {err}", flush=True)
    print(f"  saved to:    {output_dir}", flush=True)


if __name__ == "__main__":
    main()
