# Security Policy

## Supported versions

Security fixes are provided for the latest published release.

## Reporting a vulnerability

Use GitHub's private “Report a vulnerability” function when it is available for this repository.

Do not attach confidential presentations, workbooks, logs, manifests, access tokens, personal information, or client data to a public Issue. If private vulnerability reporting is temporarily unavailable, open a minimal Issue without sensitive details and ask the maintainer for a private contact method.

Include the affected version, Windows and Office versions, the failing command, expected behavior, and a minimal synthetic reproduction when possible.

## Local-processing boundary

The project runs locally and does not upload Office files. A malicious or malformed local PPTX may still target ZIP, XML, filesystem, or Office automation behavior. Keep Python, `lxml`, `openpyxl`, Windows, and Microsoft Office updated, and do not bypass archive or XML safety checks.
