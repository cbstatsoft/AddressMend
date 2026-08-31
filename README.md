# AddressMend

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/downloads/)
[![GPL-3.0-or-later](https://img.shields.io/badge/licence-GPL--3.0--or--later-A42E2B?logo=gnu&logoColor=white)](LICENSE)
![Windows](https://img.shields.io/badge/Windows-supported-0078D4?logo=windows11&logoColor=white)
![Linux](https://img.shields.io/badge/Linux-supported-FCC624?logo=linux&logoColor=black)
![macOS](https://img.shields.io/badge/macOS-supported-000000?logo=apple&logoColor=white)

AddressMend is a deterministic, rule-based desktop programme for cleaning
six-column UK contact tables. It repairs common OCR damage, normalises postcodes
and email addresses, harmonises partial addresses with postcode evidence and
produces an audit trail explaining every decision.

The ordinary desktop workflow requires only Python 3.10 or newer. It is built
primarily for restricted Windows work environments without administrator
rights, while also supporting macOS, GNU/Linux, Excel, Numbers and LibreOffice.

## Features

- Paste a Markdown table directly into the friendly desktop menu.
- Read six-column Markdown, TSV and CSV files without installing packages.
- Preserve the original row order and six output columns.
- Repair postcode spacing, Markdown email wrappers and conservative OCR errors.
- Resolve bracketed OCR alternatives such as `[x/y]` when field evidence
  selects one choice.
- Complete partial addresses from offline data or optional online lookups.
- Reuse corrections that a person has previously approved.
- Mark uncertain suggestions for review instead of silently presenting guesses
  as facts.
- Use the native Windows, macOS, Wayland or X11 clipboard when available.
- Export plain tab-separated text for Excel, Numbers and LibreOffice Calc.

The six columns are:

```text
title    first_name    last_name    address    postcode    email
```

OCR transcriptions may put alternatives in square brackets, for example `[x/y]`
or `[r/t/k]`. AddressMend expands these choices with a safety limit, then uses the
appropriate evidence for the field: email text for names, UK syntax and postcode
sources for postcodes, postcode-constrained candidates for addresses, and email
syntax for email addresses. It changes the value only when one choice is uniquely
supported; otherwise the original bracketed value is retained and marked
`unresolved` in the review report.

## What the programme creates

Results are written to `Documents/AddressMend`:

- `cleaned_entries_<date-time>.tsv` — cleaned six-column output;
- `review_report_<date-time>.tsv` — field-level explanations and review flags;
- `corrections_and_online_cache.sqlite` — approved corrections and cached
  lookups;
- `uk_addresses.sqlite` — optional offline address/postcode index.

On first use, AddressMend tries to rename an existing `AddressMend Results`,
`UK Address Harmoniser Results` or `UK Address Cleaner Results` folder to the new
`AddressMend` name. If the operating system prevents the rename, the old folder
is left untouched and the new folder is created safely.

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

1. Expand bracketed alternatives, normalise the supplied postcode and try
   credible OCR variants.
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

## Install, launch and uninstall

No third-party Python packages are required. Windows runs portably without
administrator rights, macOS installs for the current user, and the Linux
installer uses root or `sudo` for a conventional system-wide installation.
Python 3.10 or newer must already be available.

### Windows

Download or clone the repository, then double-click `start.cmd`. AddressMend
runs from that folder and saves its working files under
`Documents/AddressMend`.

### macOS

In Terminal, enter the downloaded AddressMend folder and run:

```sh
chmod +x INSTALL UNINSTALL start.command
./INSTALL
```

The installer copies the programme into
`~/Library/Application Support/AddressMend`, creates the terminal command at
`~/.local/bin/addressmend`, and creates
`~/Applications/AddressMend.command` for launching from Finder. It uses
macOS's built-in `pbcopy` and `pbpaste` clipboard tools. If macOS displays a
security warning for the launcher, right-click it and choose **Open** once.

### Linux

From the downloaded AddressMend folder, run:

```sh
chmod +x INSTALL UNINSTALL
./INSTALL
```

The installer requests `sudo` when necessary, places the programme in
`/opt/addressmend`, and creates `/usr/local/bin/addressmend`. If `sudo` is
not installed, run it through `su` as instructed. The existing `start.sh`
launcher also continues to work directly from the repository.

### Uninstall

Run `./UNINSTALL` from a downloaded repository copy. On Linux it requests
`sudo` when necessary. It removes only the installed application and launcher
files. Results, downloaded datasets and the correction database under
`Documents/AddressMend` are deliberately retained.

## Command-line use

The menu is recommended for ordinary use. The command line is available for
repeatable or advanced work.

```sh
chmod a+x addressmend.py

# Environment and feature check
./addressmend.py doctor

# Clean a saved table locally
./addressmend.py clean input.tsv -o output.tsv --audit review.tsv

# Use the standard free online helpers explicitly
./addressmend.py clean input.tsv -o output.tsv --audit review.tsv \
  --online-validate --doogal --nominatim --validate-email-domains \
  --memory corrections.sqlite

# Build an offline address index from a headed TSV
./addressmend.py build-index addresses.tsv --db uk_addresses.sqlite \
  --postcode-column postcode --address-columns address

# Teach row-aligned, manually approved corrections
./addressmend.py learn original.tsv approved.tsv \
  --memory corrections.sqlite

# Download a direct official data URL with safe resume support
./addressmend.py download https://example.invalid/data.zip \
  --destination "Downloaded address data"
```

Run `./addressmend.py --help` and the individual command's `--help`
for every option.

## Repository layout

```text
addressmend.py         standalone application
start.cmd              Windows desktop launcher
start.sh               Linux launcher
start.command          macOS Finder/Terminal launcher
INSTALL                 macOS user/Linux system installer
UNINSTALL               macOS user/Linux system uninstaller
```

The single-file application is intentional: a user in a locked-down work
environment can copy and run it without installing this project as a package.

## Licence

AddressMend is free software licensed under the GNU General Public
License version 3 or, at your option, any later version. See
[LICENSE](LICENSE) for the complete terms.

---


**🄯 Connor Baird, 2026**