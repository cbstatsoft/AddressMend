# Changelog

All notable changes to AddressMend are recorded here. The project uses
semantic versioning.

## 1.3.1 — 31 August 2026

- Replaced unrestricted fuzzy address rewriting with a two-stage
  detector/corrector policy: number-only and flat-only input can be completed
  automatically from an exact bare premise or close same-parity neighbours,
  while weaker partial, fuzzy and third-party candidates remain review-only.
- Reject address candidates with a conflicting premise number anywhere in the
  candidate, including flat and building prefixes.
- Preserve explicit postcode boundaries, remove deletion-based variants and
  avoid fuzzy replacement of already valid postcodes.
- Preserve valid uncommon email domains, stop inferring unbracketed surnames
  from concatenated email local-parts, and use only delimiter-separated identity
  evidence to resolve bracketed name alternatives.
- Preserve the supplied field when postcode, email or bracket evidence remains
  unresolved, while retaining the candidate and explanation in the audit.
- Consolidate sequential applied edits into a single accurate audit event.
- Added built-in regressions for exact-premise and neighbour-supported address
  completion, premise conflicts, postcode truncation, uncommon domains,
  shared-email surnames and review-only fuzzy changes.

## 1.3.0 — 31 August 2026

- Added native macOS clipboard integration through the built-in `pbcopy` and
  `pbpaste` commands.
- Added a Finder/Terminal `start.command` launcher with common Homebrew and
  python.org paths available when Finder supplies a restricted environment.
- Added DTABNK-style `INSTALL` and `UNINSTALL` POSIX scripts: user-local on
  macOS and system-wide under `/opt/addressmend` on Linux.
- Added macOS-specific paste guidance, diagnostics coverage, package metadata
  and full installation/uninstallation documentation.

## 1.2.1 — 31 August 2026

- Changed the friendly desktop output folder to `Documents/AddressMend` and
  added a safe rename attempt for folders created under former names.
- Expanded square-bracket OCR choices with two or more alternatives, including
  `[x/y]` and `[r/t/k]`, with a 128-variant safety limit.
- Added field-aware resolution for bracketed names, postcodes, addresses and
  email addresses; ambiguous choices remain unchanged and are marked for review.
- Updated the README, security information, package version, in-code help and
  built-in self-test to describe and check the current behaviour.

## 1.2.0 — 31 August 2026

- Renamed the former UK Address Harmoniser software **AddressMend** and aligned
  its module, launchers, package metadata, documentation and release archives.
- Kept the Linux/X11 clipboard owner alive so TSV can be pasted into
  LibreOffice Calc under i3, with Wayland, xclip, xsel and Tk fallbacks.
- Added a Linux launcher and platform-specific paste instructions.
- Added general Doogal street-consensus and shared-building completion for
  known-address lists that omit a supplied premise or flat.
- Added a cautious OpenStreetMap/Nominatim lookup for missing postcodes, with
  caching, disclosure and a policy-compliant request delay.
- Reused approved person spellings by exact email instead of leaving the
  existing correction-memory table unused.
- Separated harmless Markdown removal and provisional harmonisation from
  substantive high-confidence changes in the audit report.
- Added a guided, resumable offline-data downloader and official access pages
  for sources requiring sign-in or licence acceptance.
- Added deterministic release building, checksums, GitHub tag automation and
  privacy-safe synthetic fixtures for the public repository.
- Passed the automated unit/integration suite and a real 30-row regression run.

## 1.1.1 — 29 August 2026

- Corrected the final summary so unchanged DNS verifications are not counted as
  field changes.
- Added successful uncommon-domain verification to the per-row explanation.
- Passed a generated 12-row TSV integration test covering full, partial,
  ambiguous and OCR-damaged addresses plus postcode and email OCR cases.

## 1.1.0 — 29 August 2026

- Simplified the desktop banner to the programme name and version.
- Made the full online-assisted cleaning procedure the desktop default.
- Added main-menu option 8 to switch explicitly between standard and local-only
  processing.
- Retained clear disclosure of the limited postcode/domain values sent to each
  lookup service.

## 1.0.0 — 29 August 2026

- First release-ready Windows edition.
- Added a double-click Command Prompt launcher and beginner-friendly menu.
- Added direct multi-line table pasting terminated by `DONE`.
- Added Markdown, CSV, TSV, clipboard and file input.
- Added Excel-ready TSV output and a field-level review report.
- Added conservative postcode, email, name and OCR normalisation.
- Added offline SQLite/FTS5 address indexing and guided source imports.
- Added optional sequential Doogal and postcodes.io lookup.
- Added postcode-constrained harmonisation of full, partial and OCR-damaged addresses.
- Added conservative email OCR repair and cached uncommon-domain MX/DNS checking.
- Added optional licensed getAddress.io support.
- Added approved-correction learning and an online-result cache.
- Added environment diagnostics, explicit privacy explanations and self-tests.
- Added Git project metadata, automated tests and Windows/Linux CI configuration.
- Licensed under GPL-3.0-or-later; copyright Connor Baird.
