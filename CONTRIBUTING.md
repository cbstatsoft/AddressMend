# Contributing

AddressMend is maintained through small, reviewable Git commits.

## Before committing

1. Keep the standard Windows workflow free of mandatory third-party packages.
2. Preserve the six output columns and original row order.
3. Never silently guess an ambiguous personal name, address or postcode.
4. Explain every automatic correction in the audit trail.
5. Do not commit contact data, result files, SQLite databases or downloaded
   address datasets.
6. Use invented identities and reserved example domains in tests, examples and
   issue reports. Never turn a user's contact list into a public fixture.
7. Run:

   ```text
   python -m unittest discover -s tests -v
   python addressmend.py self-test
   python -m py_compile addressmend.py
   python scripts/build_release.py --check
   ```

8. Update `CHANGELOG.md`, `RELEASE_NOTES.md` and both version declarations when
   making a release.

## Commit scope

Prefer one coherent change per commit. Parser, normalisation, lookup, indexing
and desktop-interface changes should remain in their corresponding functions
and receive focused tests. Network-dependent tests must not be required for a
normal test run.

The single-file deployment is intentional for restricted Windows computers; do
not replace its focused functions with one large procedural workflow. If the
module eventually needs splitting, retain the standalone release launcher and
avoid making package installation mandatory for ordinary users.

Contributions are accepted under the project's GPL-3.0-or-later licence.

## Release procedure

1. Confirm the working tree contains no personal or downloaded data.
2. Run the four checks above.
3. Build reproducible archives with `python scripts/build_release.py`.
4. Inspect `dist/SHA256SUMS.txt` and both ZIP contents.
5. Commit the release, then create and push the matching annotated tag, such as
   `v1.2.0`. The GitHub workflow validates the tag and publishes the archives.
