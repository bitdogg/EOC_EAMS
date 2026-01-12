# EOC_EAMS — Denodo Extract Scripts

Scripts used by the EOC to extract, normalize, and archive EAMS data from a
Denodo V5 web endpoint.

The extractor pulls structured JSON data embedded in Denodo HTML pages and
produces timestamped **JSON snapshots** and **CSV exports** for analysis,
auditing, and downstream automation.

## What this project does
- YAML-driven configuration (CLI or GUI picker)
- Secure credential retrieval via OS keyring
- Resilient HTTP retrieval with retries
- Extracts embedded JSON from HTML
- Writes timestamped JSON and CSV outputs

## How to run
```bash
python Denodo_Extract.py --config C:\path\to\config.yml
```

See `docs/Developer_User_Guide.md` for full documentation.
