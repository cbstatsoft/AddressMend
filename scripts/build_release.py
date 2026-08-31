#!/usr/bin/env python3
# Copyright (C) 2026 Connor Baird
# SPDX-License-Identifier: GPL-3.0-or-later
"""Validate and build deterministic AddressMend release archives."""

from __future__ import annotations

import argparse
import hashlib
import re
import tomllib
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DIST = ROOT / "dist"
FIXED_ZIP_TIME = (1980, 1, 1, 0, 0, 0)

DESKTOP_FILES = (
    "addressmend.py",
    "Start_AddressMend.cmd",
    "Start_AddressMend.sh",
    "READ_ME_FIRST.txt",
    "LICENSE.txt",
    "README.md",
    "CHANGELOG.md",
    "RELEASE_NOTES.md",
    "SECURITY.md",
)

REPOSITORY_FILES = (
    ".editorconfig",
    ".gitattributes",
    ".github/workflows/release.yml",
    ".github/workflows/tests.yml",
    ".github/ISSUE_TEMPLATE/bug_report.yml",
    ".github/ISSUE_TEMPLATE/config.yml",
    ".github/pull_request_template.md",
    ".gitignore",
    "CHANGELOG.md",
    "CONTRIBUTING.md",
    "LICENSE.txt",
    "MANIFEST.in",
    "README.md",
    "READ_ME_FIRST.txt",
    "RELEASE_NOTES.md",
    "SECURITY.md",
    "Start_AddressMend.cmd",
    "Start_AddressMend.sh",
    "generated_address_reference.tsv",
    "generated_test_input.tsv",
    "pyproject.toml",
    "scripts/build_release.py",
    "tests/test_generated_tsv_integration.py",
    "tests/test_addressmend.py",
    "addressmend.py",
)


def project_version() -> str:
    metadata = tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))
    return str(metadata["project"]["version"])


def application_version() -> str:
    source = (ROOT / "addressmend.py").read_text(encoding="utf-8")
    match = re.search(r'^VERSION\s*=\s*"([^"]+)"', source, re.MULTILINE)
    if not match:
        raise SystemExit("VERSION was not found in addressmend.py")
    return match.group(1)


def release_title() -> str:
    first_line = (ROOT / "RELEASE_NOTES.md").read_text(encoding="utf-8").splitlines()[0]
    if not first_line.startswith("# "):
        raise SystemExit("RELEASE_NOTES.md must begin with one '# ' title")
    return first_line[2:].strip()


def validate(tag: str | None = None) -> str:
    version = project_version()
    if application_version() != version:
        raise SystemExit("pyproject.toml and application VERSION do not match")
    if tag and tag != f"v{version}":
        raise SystemExit(f"tag {tag!r} does not match version v{version}")
    expected_title = f"AddressMend {version}"
    if release_title() != expected_title:
        raise SystemExit(f"release title must be {expected_title!r}")
    missing = [
        name
        for name in dict.fromkeys(DESKTOP_FILES + REPOSITORY_FILES)
        if not (ROOT / name).is_file()
    ]
    if missing:
        raise SystemExit("missing release files: " + ", ".join(missing))
    for forbidden in ("upload", "outputs", "real_results", "corrections_and_online_cache"):
        if any(forbidden in name.casefold() for name in REPOSITORY_FILES):
            raise SystemExit(f"private/runtime path included in source archive: {forbidden}")
    return version


def file_bytes(path: Path) -> bytes:
    data = path.read_bytes()
    if path.suffix.casefold() == ".cmd":
        data = data.replace(b"\r\n", b"\n").replace(b"\n", b"\r\n")
    return data


def write_zip(target: Path, names: tuple[str, ...], root_name: str) -> None:
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        for name in names:
            source = ROOT / name
            info = zipfile.ZipInfo(f"{root_name}/{name}", FIXED_ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            mode = 0o755 if source.suffix == ".sh" or name.startswith("scripts/") else 0o644
            info.external_attr = (mode & 0xFFFF) << 16
            info.create_system = 3
            archive.writestr(info, file_bytes(source), compresslevel=9)


def digest(path: Path) -> str:
    checksum = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            checksum.update(block)
    return checksum.hexdigest()


def build(tag: str | None = None) -> list[Path]:
    version = validate(tag)
    DIST.mkdir(exist_ok=True)
    desktop = DIST / f"AddressMend_{version}_Desktop.zip"
    repository = DIST / f"AddressMend_{version}_Repository.zip"
    root_name = f"AddressMend-{version}"
    write_zip(desktop, DESKTOP_FILES, root_name)
    write_zip(repository, REPOSITORY_FILES, root_name)
    checksums = DIST / "SHA256SUMS.txt"
    checksums.write_text(
        "".join(
            f"{digest(path)}  {path.name}\n" for path in (desktop, repository)
        ),
        encoding="utf-8",
        newline="\n",
    )
    return [desktop, repository, checksums]


def parser() -> argparse.ArgumentParser:
    result = argparse.ArgumentParser(description=__doc__)
    result.add_argument("--check", action="store_true", help="validate without creating archives")
    result.add_argument("--tag", help="require an exact vX.Y.Z Git tag")
    return result


def main() -> int:
    args = parser().parse_args()
    if args.check:
        version = validate(args.tag)
        print(f"release metadata ready for v{version}: {release_title()}")
        return 0
    for path in build(args.tag):
        print(path.relative_to(ROOT))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
