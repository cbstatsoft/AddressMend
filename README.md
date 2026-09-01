# AddressMend

[![Python 3.10+](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/downloads/)
[![GPL-3.0-or-later](https://img.shields.io/badge/licence-GPL--3.0--or--later-A42E2B?logo=gnu&logoColor=white)](LICENSE)
![Windows](https://img.shields.io/badge/Windows-supported-0078D4?logo=windows11&logoColor=white)
![Linux](https://img.shields.io/badge/Linux-supported-FCC624?logo=linux&logoColor=black)
![macOS](https://img.shields.io/badge/macOS-supported-000000?logo=apple&logoColor=white)

AddressMend is a deterministic, rule-based desktop programme for cleaning
six-column UK contact tables. It repairs common OCR damage, normalises postcodes
and email addresses, evaluates partial addresses against postcode evidence and
produces an audit trail explaining every decision.

The ordinary desktop workflow requires only Python 3.10 or newer. It is built
primarily for restricted Windows work environments without administrator
rights, while also supporting macOS, GNU/Linux, Microsoft Excel, Apple Numbers
and LibreOffice Calc.

## Features

- Paste a Markdown table directly into the friendly desktop menu.
- Read six-column Markdown, TSV and CSV files without installing packages.
- Preserve the original row order and six output columns.
- Repair postcode spacing, Markdown email wrappers and conservative OCR errors.
- Resolve bracketed OCR alternatives such as `[x/y]` when field evidence
  selects one choice.
- Automatically apply strongly corroborated incomplete-address and unique
  one-character OCR corrections; retain weaker or ambiguous candidates for review.
- Reuse corrections that a person has previously approved.
- Search Homedata for unresolved UK address fragments without
  sending the contact's name or email address.
- Optionally send only unresolved rows to OpenAI, another OpenAI-compatible API
  or a local Ollama model for a constrained second-stage review.
- Mark uncertain suggestions for review instead of silently presenting guesses
  as facts.
- Copy the completed TSV to the native Windows, macOS, Wayland or X11 clipboard
  after every enabled lookup and LLM stage has finished.
- Export plain tab-separated text for Microsoft Excel, Apple Numbers and
  LibreOffice Calc.

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

## OCR correction procedure

AddressMend separates **error detection** from **correction**. A different
lookup result is not, by itself, proof that the transcription is wrong. The
programme first decides whether a field is malformed, ambiguous or demonstrably
incomplete; only then does it generate and assess replacements.

1. **Preserve and normalise.** Trim repeated whitespace, remove recognised
   Markdown email wrappers and normalise harmless case or postcode spacing.
   Keep the supplied text available for both the audit and conservative
   fallback.
2. **Detect a possible OCR problem.** Detection is field-specific: square
   brackets indicate explicit alternatives; invalid syntax can flag a postcode
   or email; and a bare house/flat number is an incomplete address. Merely
   having text in the address field does not prove that the address is complete
   or correctly spelt. A single street name is kept but marked unresolved
   because its premise or named-property identifier is missing. A complete,
   syntactically valid name or address is changed only when a constrained source
   supplies a unique one-character correction with no premise contradiction.
3. **Generate restricted candidates.** Expand at most 128 bracket combinations.
   Postcodes use UK structure and length-preserving OCR character confusions,
   not arbitrary character deletion. Automatic address candidates must be
   constrained by the supplied postcode. Homedata may look beyond an incorrect
   postcode. Exact premise/property completions, unique one-character street
   corrections and a sole one-character bare-premise correction can pass its
   automatic gate. A broad Nominatim search is review-only. Name
   evidence comes only from delimiter-separated words in the email local part.
   Valid uncommon email domains remain valid candidates rather than being
   changed to a popular provider by edit distance.
4. **Reject contradictions.** Discard address candidates with a conflicting
   house, flat or unit number anywhere in the address. Reject ambiguous bracket
   choices, non-unique postcode variants and candidates that fail the relevant
   field syntax.
5. **Apply only a supported winner.** Previously approved exact corrections take
   priority. A new correction is automatic only when there is one exact bare
   premise, one postcode-constrained candidate that strictly extends the same
   premise/property text, close same-parity neighbours on the sole street, or
   one compatible address separated by exactly one OCR insertion, deletion or
   substitution. A postcode's numbered properties may establish the spelling
   of a supplied street name, but AddressMend returns only that corrected street
   name and still flags the missing premise; it never borrows a house number.
   The numeric rule confidence must also meet `--address-threshold` where
   applicable.
6. **Abstain and explain.** Fuzzy changes to complete text, broad online matches
   and conflicting or incomplete evidence are written as `review` or
   `unresolved`. They do not replace the field in the cleaned TSV.
7. **Optionally review unresolved fields with an LLM.** This stage runs only
   with `--llm-provider`. The model sees the unresolved record and deterministic
   evidence, returns structured proposals and may abstain. A `high` proposal is
   still checked locally for field syntax and premise preservation. It can be
   applied automatically only where the detector already found incomplete,
   invalid or explicitly bracketed input; changes to complete text remain
   `review` suggestions.
8. **Learn only after approval.** The `learn` command compares a raw batch with
   a row-aligned approved batch and stores exact row, person and address
   corrections in the local SQLite memory. Future exact matches are labelled
   `learned`; review suggestions never teach themselves.

The evidence used for each field is deliberately different:

| Field | Candidate evidence | Automatic gate |
| --- | --- | --- |
| Title/name | Explicit bracket alternatives; approved person matched by exact email | Unique delimiter-separated email evidence, or approved memory |
| Address | Offline index, Doogal, Homedata, optional getAddress.io or Nominatim | Approved memory; harmless formatting; a unique strict completion preserving the premise; or one unique one-character correction with no premise contradiction. Street-only values stay unresolved. |
| Postcode | UK syntax, bracket alternatives, offline postcode index and exact postcodes.io validation | Unique valid bracket/index result or exact canonical formatting; no fuzzy change to an already valid postcode |
| Email | Wrapper removal, email syntax and explicit bracket alternatives | Harmless formatting or one unique syntactically valid bracket choice; uncommon domains are not similarity-rewritten |

This design follows the established detector-then-corrector pattern for reducing
false OCR changes and uses field-specific confusion sets rather than unrestricted
edit distance. See [Schaefer and Neudecker's two-step approach](https://aclanthology.org/2020.latechclfl-1.6.pdf)
and [Hassan, Noeman and Hassan's discussion of symbol confusion matrices](https://aclanthology.org/I08-2131.pdf).

## What the programme creates

Results are written to `Documents/AddressMend`:

- `cleaned_entries_<date-time>.tsv` — cleaned six-column output;
- `review_report_<date-time>.tsv` — field-level explanations and review flags;
- `corrections_and_online_cache.sqlite` — approved corrections and cached
  lookups;
- `uk_addresses.sqlite` — optional offline address/postcode index.

After a desktop-menu batch succeeds, the exact contents of the completed
`cleaned_entries_…tsv` are also placed on the system clipboard. This happens
after deterministic lookups and the optional LLM fallback, so the saved file and
the pasted six-column table are the same final result. Command-line users can
request the same behaviour with `-o @clipboard`. On Windows, AddressMend uses the
native Unicode clipboard API and reads the value back before reporting success.
Paste the result into the first spreadsheet cell with Ctrl+V; Ctrl+C in the
AddressMend console is an interrupt, not a clipboard-copy command.

Every non-quiet batch ends with a row-level completion report. It assigns each
row to exactly one final category—no correction needed, deterministic
rules/lookups, approved correction memory, LLM automatic completion, or still
review/unresolved—and prints the count and percentage of the full batch. When an
LLM is enabled, a second line reports the percentage sent to it and how many of
those rows were automatically changed, retained only as review suggestions, or
left without a usable change. Rows that received both deterministic work and a
final LLM completion are attributed to the LLM, so the five main categories do
not double-count rows.

On first use, AddressMend tries to rename an existing `AddressMend Results`,
`UK Address Harmoniser Results` or `UK Address Cleaner Results` folder to the new
`AddressMend` name. If the operating system prevents the rename, the old folder
is left untouched and the new folder is created safely.

Confidence labels in the review report have specific meanings:

| Label | Meaning |
| --- | --- |
| `high` or numeric score | Deterministic field evidence met a calibrated automatic correction rule. |
| `llm-high` | An opt-in LLM proposal passed both the model's high-confidence gate and AddressMend's local detector/validator gates. |
| `learned` | An exact correction previously approved by a person was reused. |
| `verified` | The supplied value was supported and did not need changing. |
| `formatting` | Only wrappers, spacing or escaping changed. |
| `review` | A candidate is recorded for confirmation; the cleaned TSV retains the supplied value. |
| `unresolved` | The programme deliberately declined to guess and retained the supplied value. |

Always inspect `review` and `unresolved` items before relying on the output.

## How address harmonisation works

The cleaner does not contain person-specific or address-specific `if` rules.
It applies reusable evidence rules:

1. Preserve an explicitly supplied postcode boundary while expanding bracketed
   alternatives and trying only length-preserving OCR substitutions.
2. Reuse an exact correction that a person previously approved.
3. Validate the exact postcode and match the supplied premise or fragment within
   it using the offline index.
4. Consult Doogal's postcode-constrained known-address list.
5. Search Homedata with only the unresolved address
   fragment and postcode. This can recover a mistyped postcode or complete an
   exact named property.
6. Reject every candidate containing a conflicting premise number, including
   numbers in flat or building prefixes.
7. Treat number-only, flat-only and street-only text as detected incomplete
   input even though the address field is non-empty. Apply a completion only
   when postcode-constrained data contains one exact bare premise, one strict
   extension with identical premise identifiers, or close same-parity
   neighbours around the missing premise on the sole street.
8. Apply a unique one-character address correction automatically when its
   premise is unchanged. This includes a misspelt street suffix or named
   property. Repeated numbered candidates can verify a street's spelling, but
   the result stays street-only and unresolved instead of acquiring an
   unsupported number. A sole one-character bare-premise OCR correction can
   also be applied when no competing compatible address exists. Broader fuzzy
   changes remain `review` suggestions without changing the cleaned TSV.
9. Retain and flag the supplied value when evidence conflicts or the address
   remains incomplete.

This detector-then-corrector split prevents a plausible lookup from rewriting a
field that was not demonstrably incomplete. Numeric confidence labels for the
automatic incomplete-address rules are calibrated against the maintained
regression batch; `--address-threshold` can raise the required level.

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

## Online services and privacy

The desktop menu and `clean` command start with the standard free online-assisted
procedure enabled. Option **9** switches the desktop workflow to local-only
processing. Command-line users can disable individual services with
`--no-doogal`, `--no-online-validate`, `--no-validate-email-domains`,
`--no-nominatim` and `--no-homedata`.

| Service | Data sent | Purpose |
| --- | --- | --- |
| Doogal | Postcode only | Postcode-constrained roads and known addresses. |
| postcodes.io | Exact postcode after conservative local normalisation | Validation and canonical formatting. |
| Google Public DNS | Domain after `@` only | Check an uncommon email domain; never verifies a mailbox. |
| OpenStreetMap/Nominatim | Address only when its postcode is missing or invalid | Provisional missing-postcode recovery. |
| Homedata | Unresolved address fragment and postcode only | Wider UK address search, including recovery from an incorrect postcode. |
| getAddress.io | Partial address and postcode, only when separately configured | Licensed premise lookup. |
| Ollama Web Search | Unresolved address fragment and postcode only, when Ollama fallback and `OLLAMA_API_KEY` are configured | Bounded internet evidence for the local Ollama reviewer. |
| Remote LLM API | The complete six-field record for a row already marked `review` or `unresolved`, plus its deterministic evidence | Constrained second-stage proposal. |

Names and complete email addresses are not sent to Doogal, postcodes.io or DNS.
Valid uncommon email domains are never changed merely because they resemble a
common provider; DNS results can flag a domain for review but cannot prove or
rewrite a mailbox.
Nominatim calls are sequential, limited to at most one per second, cached and
identified with an application user agent in accordance with its
[public usage policy](https://operations.osmfoundation.org/policies/nominatim/).

Homedata search is on by default and can be disabled with desktop option **11**
or `--no-homedata`. AddressMend uses only the documented unauthenticated
address-find endpoint, makes sequential requests and caches results when
correction memory is enabled. The service states that find requests need no API
key, but it is a third-party service with no AddressMend availability guarantee.
Its [terms](https://homedata.co.uk/terms) apply.

AddressMend does not automate Rightmove or similar property-listing websites.
[Rightmove's terms](https://www.rightmove.co.uk/c/terms-of-use/) prohibit bots,
crawlers and scrapers. Such sites may still be checked manually when reviewing a
provisional result, but they are not a safe or stable programme dependency.

### Optional LLM fallback

LLM processing is off by default. It is a fallback after the deterministic
detector and lookup stages, not a replacement for them. The output is constrained
to the six known fields, checked locally and included in the audit. The cleaner
accepts three provider modes:

All high-probability deterministic corrections described above run when LLM
fallback is off. An LLM is not required for automatic correction.

Desktop users can select menu option **10** to configure a provider. Option
**12** detects installed memory, recommends a local model, installs Ollama through
Windows Package Manager with confirmation when available, and downloads the
chosen model with confirmation. Before downloading, it runs `ollama list`, shows
already-installed models from known general-purpose reviewer families and prefers
a suitable installed model. On Windows, the OpenAI option can accept the key
through a hidden prompt and persist
`OPENAI_API_KEY` for the current user through PowerShell `setx`. AddressMend also
loads it into the current process, so no restart is needed. On other platforms,
set the environment variable before starting the programme.

| Provider | API used | Defaults |
| --- | --- | --- |
| `openai` | OpenAI Responses API with strict Structured Outputs and `store: false` | `https://api.openai.com/v1`, `gpt-5.6-terra`, key in `OPENAI_API_KEY` |
| `compatible` | OpenAI-compatible `/chat/completions` with JSON output | Base URL and model required; optional key in `LLM_API_KEY` |
| `ollama` | Native Ollama `/api/chat` with a JSON schema | `http://localhost:11434`; local review needs no key |

The guided setup recommends `gpt-oss:20b` (14 GB download) when at least 24 GB
of system memory is available, `qwen3:8b` (5.2 GB) for 12–23 GB, and
`qwen3:4b` (2.5 GB) on smaller systems. The extra headroom is deliberate because
Windows and the rest of the application also need memory. The user may type a
different Ollama model instead.

```sh
# OpenAI (API billing is separate from a ChatGPT subscription)
export OPENAI_API_KEY='...'
./addressmend.py clean input.tsv -o output.tsv --audit review.tsv \
  --memory corrections.sqlite --llm-provider openai

# A different OpenAI-compatible service
export LLM_API_KEY='...'
./addressmend.py clean input.tsv -o output.tsv --audit review.tsv \
  --llm-provider compatible --llm-base-url https://llm.example/v1 \
  --llm-model provider-model

# Local Ollama; menu option 12 can install and select this automatically
ollama pull gpt-oss:20b
./addressmend.py clean input.tsv -o output.tsv --audit review.tsv \
  --memory corrections.sqlite --llm-provider ollama --llm-model gpt-oss:20b

# Optional controlled internet evidence for Ollama (enabled when this key exists)
export OLLAMA_API_KEY='...'
# Add --no-ollama-web-search to keep the Ollama run completely local.
```

AddressMend teaches every selected model the same conservative reviewer procedure
in its system instruction. It tells the model to change only fields detected as
problematic, rank exact and deterministic evidence above web snippets, preserve
premise identifiers, avoid invented personal data, abstain when evidence conflicts,
and return the required JSON only. This instruction is sent with every request, so a
separate Ollama `Modelfile` is not required.

For Ollama only, `--ollama-web-search` is on by default but does nothing unless
`OLLAMA_API_KEY` exists. When available, AddressMend calls Ollama's official
[Web Search API](https://docs.ollama.com/capabilities/web-search) for unresolved
address or postcode issues. This hosted service requires a free Ollama account and
API key; Ollama describes a generous free tier for individuals in its
[web-search announcement](https://ollama.com/blog/web-search). Only the address
fragment and postcode form the query. The contact's title, name and email are never
sent to web search. Results are truncated, attached as untrusted evidence and cannot
instruct the model or bypass AddressMend's local validation gates. A web-search
failure simply falls back to the local Ollama review.

Remote endpoints must use HTTPS; unencrypted HTTP is accepted only for a
loopback address. API keys are read from environment variables and are never
written to the audit or cache. When `--memory` is supplied, the request is
identified in SQLite by a hash and the structured result is cached to avoid
repeat cost; the cached proposal can itself contain personal data. Local Ollama
keeps the request on the computer only when its base URL is loopback. Any other
endpoint sends the full unresolved record to that provider, so check its terms,
retention and data-location policy first.

On Windows, OpenAI and Ollama web-search keys entered through the guided menus are
stored in the current user's Windows environment, just as if `setx OPENAI_API_KEY`
or `setx OLLAMA_API_KEY` had been run manually. They are not written to
AddressMend's files, but remain retrievable by programmes running under that Windows
account. Entry is hidden and passed to PowerShell through standard input, rather
than being placed in the AddressMend command line or shell history.

The API has no access to this project's previous ChatGPT conversations or
ChatGPT memory. Reuse approved results through AddressMend's explicit `learn`
command and local correction database.

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
`Documents/AddressMend`. After a batch, press Enter to return to the main menu;
choose **Q** when you intend to close it. The launcher pauses only after an error,
so a normal **Q** exit now closes cleanly without a second “Press any key” prompt.

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
addressmend.py  standalone application
start.cmd       Windows desktop launcher
start.sh        Linux launcher
start.command   macOS Finder/Terminal launcher
INSTALL         macOS user/Linux system installer
UNINSTALL       macOS user/Linux system uninstaller
```

The single-file application is intentional: a user in a locked-down work
environment can copy and run it without installing this project as a package.

## Licence

AddressMend is free software licensed under the GNU General Public
License version 3 or, at your option, any later version. See
[LICENSE](LICENSE) for the complete terms.

---


**🄯 Connor Baird, 2026**
