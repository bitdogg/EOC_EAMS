# Denodo Extract — Developer & User Guide

This guide documents the Denodo extract script end-to-end and provides
operational and teaching detail for new and experienced users.

## High-level workflow
1. Select YAML config (CLI or GUI)
2. Load paths, URL, auth metadata
3. Retrieve password from keyring
4. Fetch Denodo HTML with retries
5. Extract embedded JSON
6. Write timestamped JSON and CSV

## YAML configuration
```yaml
DATALOC:
  ROOT_DIR: C:\path\to\root
  OUTPUT_DIR: output
DATA:
  EAMS_RESULT_DATA: eams_result_data.json
  EAMS_CSV_OUTPUT: eams_output.csv
  URL: https://denodo.host.example/report
AUTHENTICATION:
  SERVICE_NAME: denodo-prod
  USERNAME: myuser
```

## Security notes
- No secrets in code or YAML
- Prefer verified TLS when possible
- Scope Denodo accounts to least privilege
