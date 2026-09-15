#!/usr/bin/env python3
"""Update a portable AddressMend folder from the official master branch.

Copyright (C) 2026 Connor Baird
SPDX-License-Identifier: GPL-3.0-or-later
"""

import argparse
from datetime import datetime, timezone
import io
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import urllib.request
import zipfile


REPOSITORY = "https://github.com/cbstatsoft/AddressMend.git"
COMMIT_URL = "https://api.github.com/repos/cbstatsoft/AddressMend/commits/master"
# Only application files are replaced. Never extract an archive over user data.
APP_FILES = (
    "addressmend.py", "start.cmd", "start.sh", "start.command", "update.py",
    "update.cmd", "README.md", "CHANGELOG.md", "LICENSE", "SECURITY.md",
    "pyproject.toml", "INSTALL", "UNINSTALL",
)
MAX_DOWNLOAD = 20 * 1024 * 1024


def download(url: str) -> bytes:
    request = urllib.request.Request(url, headers={"User-Agent": "AddressMend-updater"})
    with urllib.request.urlopen(request, timeout=60) as response:
        data = response.read(MAX_DOWNLOAD + 1)
    if len(data) > MAX_DOWNLOAD:
        raise RuntimeError("The update exceeds the download size limit.")
    return data


def latest_files() -> tuple[str, dict[str, bytes]]:
    commit = json.loads(download(COMMIT_URL))["sha"]
    if not isinstance(commit, str) or len(commit) != 40 or any(c not in "0123456789abcdef" for c in commit):
        raise RuntimeError("GitHub returned an invalid commit identifier.")
    archive = download(f"https://codeload.github.com/cbstatsoft/AddressMend/zip/{commit}")
    payload = {}
    with zipfile.ZipFile(io.BytesIO(archive)) as bundle:
        root = f"AddressMend-{commit}/"
        for name in APP_FILES:
            try:
                member = bundle.getinfo(root + name)
            except KeyError:
                raise RuntimeError(f"The update is missing {name}; nothing was changed.") from None
            if member.file_size > MAX_DOWNLOAD or sum(map(len, payload.values())) + member.file_size > MAX_DOWNLOAD:
                raise RuntimeError("The unpacked update exceeds the size limit.")
            payload[name] = bundle.read(member)
    # Check source syntax with this interpreter before touching the installation.
    for name in ("addressmend.py", "update.py"):
        compile(payload[name], name, "exec")
    return commit, payload


def confirm(question: str) -> bool:
    return input(f"{question} [y/N]: ").strip().casefold() in {"y", "yes"}


def install_files(folder: Path, payload: dict[str, bytes], commit: str) -> Path:
    """Stage first; preserve originals and roll back a failed replacement."""
    if set(payload) - set(APP_FILES):
        raise RuntimeError("Update contains an unexpected application filename.")
    for name in payload:
        target = folder / name
        if target.is_symlink() or (target.exists() and not target.is_file()):
            raise RuntimeError(f"Cannot update {name}: it is not a regular file.")
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S-%f")
    backup = folder / "update-backups" / stamp
    backup.mkdir(parents=True, exist_ok=False)
    replaced = []
    with tempfile.TemporaryDirectory(prefix="addressmend-update-", dir=folder) as temporary:
        staging = Path(temporary)
        for name, contents in payload.items():
            target = folder / name
            staged = staging / name
            staged.write_bytes(contents)
            if target.exists():
                shutil.copy2(target, backup / name)
                shutil.copymode(target, staged)
        (backup / "update-info.json").write_text(
            json.dumps({"target_commit": commit, "files": list(payload)}, indent=2), encoding="utf-8"
        )
        try:
            for name in payload:
                os.replace(staging / name, folder / name)
                replaced.append(name)
        except BaseException:
            failures = []
            for name in reversed(replaced):
                try:
                    if (backup / name).exists():
                        shutil.copy2(backup / name, folder / name)
                    else:
                        (folder / name).unlink()
                except OSError:
                    failures.append(name)
            if failures:
                raise RuntimeError(f"Update and rollback failed for {', '.join(failures)}. Restore files from {backup}.") from None
            raise
    return backup


def git_update(folder: Path, check: bool) -> None:
    git = shutil.which("git")
    if not git:
        raise RuntimeError("This is a Git checkout, but Git is unavailable. Install Git or use a separate ZIP download.")

    def run(*arguments: str) -> str:
        result = subprocess.run([git, "-C", str(folder), *arguments], capture_output=True, text=True, timeout=120)
        if result.returncode:
            raise RuntimeError(result.stderr.strip() or f"Git command failed: {arguments[0]}")
        return result.stdout.strip()

    if Path(run("rev-parse", "--show-toplevel")).resolve() != folder.resolve():
        raise RuntimeError("The application folder must be the root of its Git checkout.")
    if run("branch", "--show-current") != "master":
        raise RuntimeError("Switch to master before updating. Feature branches are left unchanged.")
    if run("status", "--porcelain", "--untracked-files=no"):
        raise RuntimeError("There are local changes or conflicts. Commit or resolve them before updating.")
    for marker in ("MERGE_HEAD", "rebase-merge", "rebase-apply", "CHERRY_PICK_HEAD", "REVERT_HEAD"):
        marker_path = Path(run("rev-parse", "--git-path", marker))
        if not marker_path.is_absolute():
            marker_path = folder / marker_path
        if marker_path.exists():
            raise RuntimeError("Finish or abort the current Git operation before updating.")
    print("Checking the official GitHub master branch...")
    run("fetch", "--no-tags", REPOSITORY, "master")
    target = run("rev-parse", "FETCH_HEAD")
    current = run("rev-parse", "HEAD")
    if current == target:
        print("AddressMend is already up to date.")
        return
    if run("merge-base", current, target) != current:
        raise RuntimeError("Local history differs from GitHub. Nothing was overwritten; reconcile the branches manually.")
    print(f"Update available: {current[:12]} -> {target[:12]}")
    if check or not confirm("Close AddressMend, then update this checkout?"):
        return
    backup = "backup/before-update-" + datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S-%f")
    run("branch", backup, current)
    run("merge", "--ff-only", target)
    print(f"Updated successfully. Previous code is preserved in Git branch {backup}.")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="check for updates without replacing application files")
    parser.add_argument("--pause", action="store_true", help="keep the Windows launcher open after completion")
    args = parser.parse_args()
    folder = Path(__file__).resolve().parent
    result = 0
    try:
        print("AddressMend updater — official GitHub master branch")
        print(f"Application folder: {folder}")
        print("Close AddressMend before installing an update. No administrator access is needed for a writable portable folder.")
        if not (folder / "addressmend.py").is_file():
            raise RuntimeError("Keep update.py beside addressmend.py in the application folder.")
        if (folder / ".git").exists():
            git_update(folder, args.check)
        else:
            print("Checking GitHub (no Git installation needed)...")
            commit, payload = latest_files()
            changed = {name: value for name, value in payload.items() if not (folder / name).is_file() or (folder / name).read_bytes() != value}
            if not changed:
                print("AddressMend is already up to date.")
            else:
                print(f"GitHub revision: {commit[:12]}; application files to replace: {', '.join(changed)}")
                print("Local edits to these files will be backed up, then replaced. Address data, results and saved keys are not updated.")
                if not args.check and confirm("Back up and update these application files?"):
                    backup = install_files(folder, changed, commit)
                    print(f"Updated successfully. Previous files: {backup}")
        print("You can now launch AddressMend normally.")
    except (Exception, KeyboardInterrupt) as exc:
        print(f"Update stopped: {exc or 'cancelled'}", file=sys.stderr)
        result = 1
    finally:
        if args.pause:
            try:
                input("\nPress Enter to close this updater...")
            except (EOFError, KeyboardInterrupt):
                pass
    return result


if __name__ == "__main__":
    raise SystemExit(main())
