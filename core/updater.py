import os
import sys
import json
import urllib.request
import urllib.error
import subprocess
import tempfile
import threading
from typing import Optional, Dict, Any, Callable

try:
    from config import APP_VERSION
except ImportError:
    APP_VERSION = "1.0.0"

GITHUB_REPO = "monkdluffy/Me2D"
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def parse_version_tuple(v_str: str) -> tuple:
    """Normalize version string like 'v1.2.3' or '1.2.0' to tuple (1, 2, 3) for accurate semver comparison."""
    if not v_str:
        return (0, 0, 0)
    clean = v_str.strip().lstrip("vV")
    parts = []
    for part in clean.split("."):
        try:
            parts.append(int(part))
        except ValueError:
            # Handle alpha/beta suffix if present
            digits = "".join([c for c in part if c.isdigit()])
            parts.append(int(digits) if digits else 0)
    return tuple(parts)


def is_newer_version(remote_ver: str, current_ver: str) -> bool:
    """Check if remote version is strictly newer than current version."""
    return parse_version_tuple(remote_ver) > parse_version_tuple(current_ver)


def check_for_updates(current_version: str = APP_VERSION, timeout: int = 5) -> Optional[Dict[str, Any]]:
    """
    Check GitHub Releases API for new version.
    Returns None if no newer version or network error/offline.
    Returns dict if an update is available:
    {
        "tag_name": "v1.1.0",
        "name": "Canido Doc Manager v1.1.0",
        "body": "Catatan update...",
        "html_url": "https://github.com/monkdluffy/Me2D/releases/tag/v1.1.0",
        "download_url": "...",
        "asset_name": "Canido_Doc_Manager_Setup.exe",
        "asset_size": 12345678,
        "is_direct_download": True
    }
    """
    try:
        req = urllib.request.Request(
            GITHUB_API_URL,
            headers={
                "User-Agent": f"CanidoDocManager-App/{current_version}",
                "Accept": "application/vnd.github.v3+json"
            }
        )
        with urllib.request.urlopen(req, timeout=timeout) as response:
            if response.status != 200:
                return None
            data = json.loads(response.read().decode("utf-8"))

        remote_tag = data.get("tag_name", "")
        if not remote_tag:
            return None

        if is_newer_version(remote_tag, current_version):
            download_url = None
            asset_name = None
            asset_size = 0

            # Prefer installer binary (.exe or .msi or .zip)
            assets = data.get("assets", [])
            for asset in assets:
                name = asset.get("name", "")
                if name.lower().endswith((".exe", ".msi", ".zip")):
                    download_url = asset.get("browser_download_url")
                    asset_name = name
                    asset_size = asset.get("size", 0)
                    break

            is_direct = True
            if not download_url:
                # Fallback to browser release URL
                download_url = data.get("html_url", f"https://github.com/{GITHUB_REPO}/releases/latest")
                asset_name = "Kunjungi Halaman Rilis GitHub"
                is_direct = False

            return {
                "tag_name": remote_tag,
                "name": data.get("name") or remote_tag,
                "body": data.get("body", "Pembaruan fitur dan stabilitas sistem.").strip(),
                "html_url": data.get("html_url", f"https://github.com/{GITHUB_REPO}/releases/latest"),
                "download_url": download_url,
                "asset_name": asset_name,
                "asset_size": asset_size,
                "is_direct_download": is_direct
            }
        return None
    except Exception:
        # Silently fail on 404 (no releases yet), timeout, or offline
        return None


def download_update_file(
    url: str,
    dest_path: str,
    progress_callback: Optional[Callable[[float, int, int], None]] = None
) -> bool:
    """
    Streams file download with progress updates.
    progress_callback(fraction_0_to_1, downloaded_bytes, total_bytes)
    """
    try:
        req = urllib.request.Request(
            url,
            headers={"User-Agent": "CanidoDocManager-Updater"}
        )
        with urllib.request.urlopen(req, timeout=30) as resp, open(dest_path, "wb") as out_file:
            total_size = int(resp.headers.get("Content-Length", 0))
            downloaded = 0
            chunk_size = 64 * 1024  # 64 KB

            while True:
                chunk = resp.read(chunk_size)
                if not chunk:
                    break
                out_file.write(chunk)
                downloaded += len(chunk)
                if progress_callback and total_size > 0:
                    fraction = min(1.0, downloaded / total_size)
                    progress_callback(fraction, downloaded, total_size)

        return True
    except Exception:
        if os.path.exists(dest_path):
            try:
                os.remove(dest_path)
            except Exception:
                pass
        return False


def launch_installer_and_exit(installer_path: str):
    """
    Launches downloaded installer executable and terminates the running application.
    """
    try:
        if sys.platform == "win32":
            subprocess.Popen([installer_path], shell=True)
        else:
            subprocess.Popen(["open", installer_path])
        # Exit running Python application
        sys.exit(0)
    except Exception:
        pass
