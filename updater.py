"""Keep Roof CRM up to date: no git, no zip juggling.

start.bat runs this before every start (`--auto`), so opening the app is
also how it updates. update.bat runs it on demand.

How it works: ask GitHub which version of the main branch is current, and
if it isn't the one installed here, download that branch as a zip (about
300 KB) and copy its files over this folder. Only files that actually
changed are written. Your data lives outside this folder (Documents/RoofCRM)
and is never touched.

A folder that is a git clone is updated with git instead, so its history
stays intact.

Standard library only: it has to run even when the app's own requirements
are missing or out of date.
"""

import io
import os
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

REPO = "morales451/CRM"
BRANCH = "main"
SHA_URL = f"https://api.github.com/repos/{REPO}/commits/{BRANCH}"
ZIP_URL = f"https://codeload.github.com/{REPO}/zip/refs/heads/{BRANCH}"

APP_DIR = Path(__file__).resolve().parent
STAMP = ".installed_version"      # the commit this folder was last updated to


def latest_sha(timeout: float = 5) -> str | None:
    """The newest commit on the branch, or None when GitHub can't be reached."""
    req = urllib.request.Request(SHA_URL, headers={
        "Accept": "application/vnd.github.sha", "User-Agent": "roof-crm-updater"})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            sha = r.read().decode().strip()
    except Exception:
        return None
    return sha if len(sha) == 40 and all(c in "0123456789abcdef" for c in sha) else None


def download(timeout: float = 30) -> bytes:
    req = urllib.request.Request(ZIP_URL, headers={"User-Agent": "roof-crm-updater"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def apply_zip(data: bytes, app_dir: Path) -> list[str]:
    """Copy a GitHub branch zip over app_dir. Returns the files it changed.

    The zip holds everything under one top folder ("CRM-main/"), which is
    dropped. Every file is read and checked BEFORE anything is written, so a
    bad download changes nothing. Unchanged files aren't rewritten. That
    matters for start.bat, which is running while this runs.
    """
    files = {}
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        if z.testzip() is not None:
            raise ValueError("the download is damaged")
        for info in z.infolist():
            if info.is_dir():
                continue
            parts = Path(info.filename).parts[1:]
            if not parts or ".." in parts or Path(info.filename).is_absolute():
                continue
            files[Path(*parts)] = z.read(info)
    if not any(p.name == "app.py" and len(p.parts) == 1 for p in files):
        raise ValueError("the download doesn't look like Roof CRM")

    changed = []
    root = app_dir.resolve()
    for rel, content in files.items():
        dest = (root / rel).resolve()
        if root not in dest.parents:
            continue
        if dest.exists() and dest.read_bytes() == content:
            continue
        dest.parent.mkdir(parents=True, exist_ok=True)
        tmp = dest.with_name(dest.name + ".updating")
        tmp.write_bytes(content)
        os.replace(tmp, dest)      # never leaves a half-written file behind
        changed.append(rel.as_posix())
    return sorted(changed)


def _install_requirements(app_dir: Path) -> None:
    print("  New requirements - installing...")
    subprocess.call([sys.executable, "-m", "pip", "install", "-q", "-r",
                     str(app_dir / "requirements.txt")])


def _git_update(app_dir: Path) -> int:
    before = subprocess.run(["git", "rev-parse", "HEAD"], cwd=app_dir,
                            capture_output=True, text=True).stdout.strip()
    if subprocess.call(["git", "pull", "--ff-only", "origin", BRANCH], cwd=app_dir) != 0:
        print("  Update failed (see above). Nothing was changed; starting the version you have.")
        return 1
    after = subprocess.run(["git", "rev-parse", "HEAD"], cwd=app_dir,
                           capture_output=True, text=True).stdout.strip()
    if before != after:
        diff = subprocess.run(["git", "diff", "--name-only", before, after],
                              cwd=app_dir, capture_output=True, text=True).stdout
        if "requirements.txt" in diff.split():
            _install_requirements(app_dir)
    return 0


def update(app_dir: Path = APP_DIR, auto: bool = False) -> int:
    """Bring app_dir up to date. In auto mode (on start) every problem just
    means "start the version you have", so it never blocks the app."""
    if auto and os.environ.get("ROOF_CRM_NO_AUTO_UPDATE"):
        return 0
    print("Checking for updates...")
    if (app_dir / ".git").exists():
        try:
            return _git_update(app_dir)
        except FileNotFoundError:
            pass                                  # no git on this machine: use the zip

    stamp = app_dir / STAMP
    installed = stamp.read_text().strip() if stamp.exists() else ""
    sha = latest_sha()
    if sha is None:
        print("  Couldn't reach GitHub (offline?) - starting the version you have."
              if auto else "  Couldn't reach GitHub. Check the internet connection.")
        if auto:
            return 0
    elif sha == installed:
        print("  Already up to date.")
        return 0

    try:
        changed = apply_zip(download(), app_dir)
    except Exception as e:
        print(f"  Update failed ({e}). Nothing was changed"
              + ("; starting the version you have." if auto else "."))
        return 0 if auto else 1
    if sha:
        stamp.write_text(sha + "\n")
    if not changed:
        print("  Already up to date.")
        return 0
    print(f"  Updated {len(changed)} file(s).")
    if "requirements.txt" in changed:
        _install_requirements(app_dir)
    if not auto:
        print("  If the app is open, close its window and start it again.")
    return 0


if __name__ == "__main__":
    sys.exit(update(auto="--auto" in sys.argv[1:]))
