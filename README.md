# Outlook Classic Arial 11 Intune Remediation

A GitHub-ready PowerShell package for standardising the default font in
**Microsoft Outlook (classic)** to Arial 11 by using Microsoft Intune
Remediations.

The detection and remediation scripts run as **SYSTEM** and enumerate eligible
loaded user hives beneath `HKEY_USERS`. They manage the six
`Office\16.0\Common\MailSettings` `REG_BINARY` values used for new messages,
replies/forwards and plain text.

## Repository contents

- `scripts/Capture-OutlookFontBaseline.ps1` captures a known-good reference
  configuration as comma-separated hexadecimal.
- `scripts/Detect-OutlookArial11.ps1` returns `0` when every loaded eligible
  user hive matches the approved baseline, otherwise `1`.
- `scripts/Remediate-OutlookArial11.ps1` writes and verifies the approved
  values.
- `scripts/Rollback-OutlookArial11.ps1` removes only the six values managed by
  the package.

## Intune settings

Create an Intune Remediation package and use:

- **Run this script using the logged-on credentials:** No
- **Run script in 64-bit PowerShell:** Yes
- Detection script: `Detect-OutlookArial11.ps1`
- Remediation script: `Remediate-OutlookArial11.ps1`

Pilot before broad assignment.

## Baseline note

The bundled detection/remediation scripts contain the Arial 11 values recovered
from the final reference scripts. Office builds and language settings can affect
the binary payload. For a new tenant or build, use the capture script on a
reference workstation, compare the results, and update the six expected values
if required.

## Rollback

First remove/disable the Intune assignment, then run the rollback script. The
rollback deliberately does **not** attempt to recreate prior user-customised
fonts.

## Scope

This package is for Outlook Classic only. It does not configure Word, Excel,
PowerPoint, New Outlook, Outlook on the web, or mobile Outlook.
