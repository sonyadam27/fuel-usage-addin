# Security Policy

## Supported Versions

| Add-In | Status |
|---|---|
| Fuel Harness Report | Active |
| Fuel Usage Per Day | Active |
| Fuel Theft Report | Active |
| Fuel Fill-Up Report | Active |

## Reporting a Vulnerability

If you find a security issue in this add-in source:

1. **Do not open a public GitHub issue.**
2. Email **sonyadam@geotab.com** with the subject line `[SECURITY] fuel-usage-addin`.
3. Include a description of the issue, steps to reproduce, and potential impact.
4. You will receive a response within 5 business days.

## What Counts as a Vulnerability

- Exposure of MyGeotab API credentials or session tokens in committed files
- Cross-site scripting (XSS) in add-in HTML/JS that could affect MyGeotab users
- Logic errors in theft detection that could produce false reports leading to incorrect action
- Unauthorised modification of add-in source that causes malicious behaviour in MyGeotab

## What Does NOT Belong in This Repo

- MyGeotab session tokens, bearer tokens, or API keys
- Raw fleet data (CSV exports, device maps)
- Personal access tokens for any service
- Database credentials of any kind

These are enforced via `.gitignore` and GitHub secret scanning.
