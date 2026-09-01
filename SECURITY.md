# Security and privacy

## Supported version

Security and privacy fixes are applied to the latest release of AddressMend.
At present that is version 1.6.x.

## Reporting a vulnerability

Use the repository's private security-advisory facility when available. If that
is unavailable, open an issue containing only enough information to request a
private contact route.

Do not include real names, postal addresses, email addresses, API keys,
correction databases, result files or downloaded address datasets in a public
issue.

## Data-handling expectations

- Keep live input, output and SQLite files outside Git.
- Do not commit test cases copied from real contact records.
- Store API keys in environment variables or through AddressMend's hidden guided
  entry, not source files or command history.
- On Windows, guided provider setup and menu option 13 may persist API keys in
  the current user's environment through PowerShell `setx`. Entry is hidden and
  passed through standard input, but software running as that Windows user can
  retrieve it.
- On macOS and Linux, guided entry stores keys in AddressMend's per-user
  `api_keys.json` configuration with an owner-only directory (`0700`) and file
  (`0600`). The file is plain text, is never committed automatically, and remains
  readable by software running as the same OS user. AddressMend refuses to load
  it if group or other permissions are present.
- Treat `--llm-provider` as disclosure of an entire unresolved contact row to
  the configured endpoint. Use loopback Ollama when records must remain local.
- Remember that the local LLM cache can contain names, addresses and email
  addresses even though its lookup key is hashed.
- The correction-memory SQLite database and `approval_decisions_…tsv` files can
  contain current and suggested addresses plus the operator's decisions. Treat
  them as contact data and do not attach them to public issues or commits.
- Treat third-party address results as evidence, not guaranteed truth.
- Confirm substantive address, postcode, name and email suggestions before
  approving them into correction memory; review-only candidates do not alter
  the cleaned output.
- Review all `review` and `unresolved` audit entries before operational use.
- Check the relevant provider's current terms, attribution and rate limits.

The programme performs no telemetry and has no automatic update mechanism.
Network lookups occur only when the corresponding online options are enabled.
