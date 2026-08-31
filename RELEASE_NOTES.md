# AddressMend 1.2.0

This release makes the desktop workflow reliable across Windows, Linux/i3,
Excel and LibreOffice Calc, while improving conservative address completion
without adding person-specific rules.

## Highlights

- Keeps an X11 clipboard owner alive so TSV remains pasteable into Calc under
  i3; also supports Wayland, xclip, xsel and a built-in Tk fallback.
- Adds postcode-wide street consensus for incomplete Doogal results.
- Collapses flat listings to a shared base address only when every relevant
  candidate agrees on the building.
- Marks inferred street/number combinations as provisional when the exact
  premise is absent from the source.
- Reuses manually approved person spellings through exact-email correction
  memory, rather than embedding batch-specific data in source code.
- Adds a rate-limited, cached OpenStreetMap/Nominatim fallback for addresses
  with missing or invalid postcodes.
- Adds a guided desktop downloader for offline address sources, including
  progress and resumable partial files.
- Separates formatting, verification, provisional changes and unresolved items
  clearly in the audit report.

## Platforms

- Windows with Python 3.10 or newer; no administrator rights required.
- Linux with Python 3.10 or newer, including i3 and LibreOffice Calc.

## Verification

- Standard-library unit and integration tests pass on Windows and Linux.
- Release packaging checks version consistency and excludes personal/runtime
  data.
- The supplied 30-row regression batch remained six-column TSV and gained ten
  additional corrected or harmonised rows; uncertain cases remained marked for
  review.

## Upgrade notes

Extract the desktop ZIP into a new folder. To preserve previously approved
corrections, keep a backup of `corrections_and_online_cache.sqlite` in
`Documents/AddressMend Results`. Existing installations using the
former `UK Address Harmoniser Results` or `UK Address Cleaner Results` folder
continue to use it automatically so their correction memory is not lost.

Review the online-service disclosure when enabling internet-assisted cleaning.
No complete free/open UK premise register exists, so provisional and unresolved
items still require human confirmation.
