# AddressMend

[![Tests](https://github.com/cbstatsoft/addressmend/actions/workflows/tests.yml/badge.svg)](https://github.com/cbstatsoft/addressmend/actions/workflows/tests.yml)
[![Release 1.2.0](https://img.shields.io/badge/release-1.2.0-2ea44f)](https://github.com/cbstatsoft/addressmend/releases/tag/v1.2.0)
![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)
[![GPL-3.0-or-later](https://img.shields.io/badge/licence-GPL--3.0--or--later-blue)](LICENSE)
![Windows and Linux](https://img.shields.io/badge/platform-Windows%20%7C%20Linux-555)

AddressMend is a deterministic, rule-based desktop programme for cleaning
six-column UK contact tables. It repairs common OCR damage, normalises postcodes
and email addresses, harmonises partial addresses with postcode evidence and
produces an audit trail explaining every decision.

The ordinary desktop workflow requires only Python 3.10 or newer. It is built
primarily for restricted Windows work environments without administrator
rights, while also supporting GNU/Linux and LibreOffice.

**Current release:** `v1.2.0`

**Licence:** GNU GPL version 3 or later (`GPL-3.0-or-later`)

**Copyright:** Copyright (C) 2026 Connor Baird

## Why use it?

- Paste a Markdown table directly into the friendly desktop menu.
- Read six-column Markdown, TSV and CSV files without installing packages.
- Preserve the original row order and six output columns.
- Repair postcode spacing, Markdown email wrappers and conservative OCR errors.
- Complete partial addresses from offline data or optional online lookups.
- Reuse corrections that a person has previously approved.
- Mark uncertain suggestions for review instead of silently presenting guesses
  as facts.
- Export plain tab-separated text for Excel and LibreOffice Calc.

The six columns are:

```text
title    first_name    last_name    address    postcode    email
```

## Quick start

### Windows

1. Download and extract `AddressMend_1.2.0_Desktop.zip` from the release.
2. Keep the extracted files together.
3. Double-click `Start_AddressMend.cmd`.
4. Choose **1 — Paste entries into this window**.
5. Paste the complete table, type `DONE` on a new line and press Enter.
6. Select cell A1 in Excel or Calc and press Ctrl+V.

No administrator rights or Python package installation are required. Python
3.10 or newer must already be available on the computer.

### Linux, i3 and LibreOffice Calc

Open a terminal in the extracted folder and run:

```sh
chmod +x Start_AddressMend.sh
./Start_AddressMend.sh
```

Use Ctrl+Shift+V to paste into most Linux terminals. When cleaning finishes,
select cell A1 in Calc and press Ctrl+V. The programme keeps an X11 clipboard
owner alive so the copied data does not disappear when the cleaning operation
ends. It automatically uses `wl-copy`, `xclip`, `xsel` or a persistent Tk
fallback according to what is already available.

If clipboard integration is unavailable, open the saved `.tsv` file in Calc
and select **Separated by: Tab** and **Character set: Unicode (UTF-8)**.

## What the programme creates

Results are written to `Documents/AddressMend Results`:

- `cleaned_entries_<date-time>.tsv` — cleaned six-column output;
- `review_report_<date-time>.tsv` — field-level explanations and review flags;
- `corrections_and_online_cache.sqlite` — approved corrections and cached
  lookups;
- `uk_addresses.sqlite` — optional offline address/postcode index.

Confidence labels in the review report have specific meanings:

| Label | Meaning |
| --- | --- |
| `high`, numeric score, or `learned` | Evidence met the automatic correction threshold. |
| `verified` | The supplied value was supported and did not need changing. |
| `formatting` | Only wrappers, spacing or escaping changed. |
| `review` | A provisional harmonisation was useful but still needs confirmation. |
| `unresolved` | The programme deliberately declined to guess. |

Always inspect `review` and `unresolved` items before relying on the output.

## How address harmonisation works

The cleaner does not contain person-specific or address-specific `if` rules.
It applies reusable evidence rules:

1. Normalise the supplied postcode and try credible OCR variants.
2. Check approved correction memory.
3. Match the premise or fragment within that postcode using the offline index.
4. Optionally consult Doogal's postcode-constrained known-address list.
5. Accept a full match only when the score and uniqueness threshold are met.
6. Infer a street provisionally only when all usable postcode candidates agree
   on that street.
7. Collapse flat listings to a shared base address only when they agree on one
   building.
8. Leave the value unresolved when evidence conflicts.

Approved batch-specific corrections live in the local SQLite memory, not in
the public source code.

## Offline data

Desktop option **4** downloads direct official files with progress and resume
support. Option **5** imports a downloaded file into the local SQLite/FTS5
index.

| Source | Coverage and use | Access |
| --- | --- | --- |
| [HM Land Registry Price Paid Data](https://www.gov.uk/government/statistical-data-sets/price-paid-data-downloads) | England and Wales sale addresses since 1995; incomplete for properties not sold for value. | Direct discovery/download; very large. |
| [Energy Performance Certificates](https://get-energy-performance-data.communities.gov.uk/) | Broad England and Wales residential coverage, with historical duplicates. | Free sign-in required. |
| [Companies House](https://download.companieshouse.gov.uk/) | UK registered-office addresses and named buildings. | Latest monthly snapshot can be discovered automatically. |
| [Food Standards Agency](https://ratings.food.gov.uk/open-data) | UK food-establishment addresses. | Official selection/download page. |
| [OpenStreetMap](https://download.geofabrik.de/europe/united-kingdom.html) | UK-wide community data; address coverage varies. | ODbL extracts; PBF import needs optional `pyosmium`. |
| [Code-Point Open](https://osdatahub.os.uk/downloads/open/CodePointOpen) | Authoritative live postcode reference for Great Britain; no premise addresses. | Free OS access conditions apply. |
| Generic CSV/TSV/XML/ZIP | Any lawfully obtained file with postcode and address fields. | Local import. |

No complete, authoritative UK premise-address register is both free and open.
Combining sources improves coverage but does not remove the need for review.

## Optional online services and privacy

The desktop menu starts with the standard online-assisted procedure enabled.
Option **9** switches to local-only processing.

| Service | Data sent | Purpose |
| --- | --- | --- |
| Doogal | Postcode only | Postcode-constrained roads and known addresses. |
| postcodes.io | Postcode or plausible postcode variants | Validation and canonical formatting. |
| Google Public DNS | Domain after `@` only | Check an uncommon email domain; never verifies a mailbox. |
| OpenStreetMap/Nominatim | Address only when its postcode is missing or invalid | Provisional missing-postcode recovery. |
| getAddress.io | Partial address and postcode, only when separately configured | Licensed premise lookup. |

Names and complete email addresses are not sent to Doogal, postcodes.io or DNS.
Nominatim calls are sequential, limited to at most one per second, cached and
identified with an application user agent in accordance with its
[public usage policy](https://operations.osmfoundation.org/policies/nominatim/).

Input files, outputs, downloaded datasets and SQLite databases are ignored by
Git because they can contain personal data or be very large. Do not add real
contact data to bug reports, tests or commits.

## Command-line use

The menu is recommended for ordinary use. The command line is available for
repeatable or advanced work.

```sh
# Environment and feature check
python addressmend.py doctor

# Clean a saved table locally
python addressmend.py clean input.tsv -o output.tsv --audit review.tsv

# Use the standard free online helpers explicitly
python addressmend.py clean input.tsv -o output.tsv --audit review.tsv \
  --online-validate --doogal --nominatim --validate-email-domains \
  --memory corrections.sqlite

# Build an offline address index from a headed TSV
python addressmend.py build-index addresses.tsv --db uk_addresses.sqlite \
  --postcode-column postcode --address-columns address

# Teach row-aligned, manually approved corrections
python addressmend.py learn original.tsv approved.tsv \
  --memory corrections.sqlite

# Download a direct official data URL with safe resume support
python addressmend.py download https://example.invalid/data.zip \
  --destination "Downloaded address data"
```

Run `python addressmend.py --help` and the individual command's `--help`
for every option.

## Repository layout

```text
addressmend.py                     standalone application
Start_AddressMend.cmd              Windows desktop launcher
Start_AddressMend.sh               Linux launcher
tests/                             standard-library automated tests
generated_test_input.tsv           synthetic integration-test input
generated_address_reference.tsv    synthetic offline test reference
scripts/build_release.py           reproducible release packager
.github/workflows/tests.yml        Windows/Linux test matrix
.github/workflows/release.yml      tagged GitHub release automation
```

The single-file application is intentional: a user in a locked-down work
environment can copy and run it without installing this project as a package.

## Licence

AddressMend is free software licensed under the GNU General Public
License version 3 or, at your option, any later version. See
[LICENSE](LICENSE) for the complete terms.
---

**🄯 Connor Baird, 2026**
