"""Opt-in removal of inactive App Service deployment ZIPs; never touches DBs."""
from __future__ import annotations

import base64
import json
import os
from pathlib import Path
import shutil
import sys
import urllib.request
import xml.etree.ElementTree as ET


def prune_inactive_packages(root: Path) -> dict:
    if root.is_symlink():
        raise ValueError("Package directory must not be a symbolic link.")
    root = root.resolve(strict=True)
    active_name = (root / "packagename.txt").read_text(encoding="utf-8").strip()
    if Path(active_name).name != active_name or not active_name.endswith(".zip"):
        raise ValueError("Invalid active deployment package; no files removed.")
    active = root / active_name
    if not active.is_file() or active.is_symlink():
        raise ValueError("Active deployment package is missing or unsafe; no files removed.")
    candidates = sorted(path for path in root.glob("*.zip") if path.name != active_name)
    # Validate every target before performing any deletion.
    for path in candidates:
        if path.is_symlink() or path.resolve().parent != root or not path.is_file():
            raise ValueError("Unsafe inactive package path; no files removed.")
    before = shutil.disk_usage(root).free
    removed = []
    for path in candidates:
        if (root / "packagename.txt").read_text(encoding="utf-8").strip() != active_name:
            raise ValueError("Active deployment changed; cleanup stopped.")
        size = path.stat().st_size
        path.unlink()
        removed.append({"name": path.name, "bytes": size})
    return {"active_package": active_name, "removed": removed,
            "free_bytes_before": before, "free_bytes_after": shutil.disk_usage(root).free}


def main() -> None:
    if sys.argv[1:] == ["--remote"]:
        print(json.dumps(prune_inactive_packages(Path("/home/data/SitePackages"))))
        return
    if sys.argv[1:]:
        raise ValueError("Unexpected cleanup arguments.")
    profile_text = os.environ["AZURE_PUBLISH_PROFILE"]
    profiles = ET.fromstring(profile_text).findall("publishProfile")
    profile = next((p for p in profiles if p.get("publishMethod") == "ZipDeploy"), profiles[0])
    username, password = profile.get("userName", ""), profile.get("userPWD", "")
    if not username or not password:
        raise ValueError("Missing SCM publishing credentials.")
    encoded = base64.b64encode(Path(__file__).read_bytes()).decode()
    command = (
        'python3 -c "import base64; exec(compile(base64.b64decode(\''
        + encoded + "'), '<package cleanup>', 'exec'))\" --remote"
    )
    auth = base64.b64encode(f"{username}:{password}".encode()).decode()
    req = urllib.request.Request(
        "https://" + os.environ["AZURE_WEBAPP_NAME"] + ".scm.azurewebsites.net/api/command",
        data=json.dumps({"command": command, "dir": "/home"}).encode(),
        headers={"Authorization": "Basic " + auth, "Content-Type": "application/json"},
    )
    with urllib.request.urlopen(req, timeout=90) as response:
        result = json.load(response)
    if result.get("ExitCode") != 0:
        raise RuntimeError(str(result.get("Error") or result.get("Output") or "SCM cleanup failed"))
    print(result.get("Output", ""))


if __name__ == "__main__":
    main()
