## Summary

Explain the user-visible change and why it is needed.

## Verification

- [ ] `python -m unittest discover -s tests -v`
- [ ] `python addressmend.py self-test`
- [ ] `python -m py_compile addressmend.py scripts/build_release.py`
- [ ] `python scripts/build_release.py --check`

## Privacy and compatibility

- [ ] Tests and examples use invented identities and reserved example domains.
- [ ] No input, output, database, API key or downloaded dataset is included.
- [ ] The no-admin Windows workflow and six-column TSV format still work.
- [ ] User-facing wording uses British English.
