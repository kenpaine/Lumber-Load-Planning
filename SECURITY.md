# Security Policy

## Scope

This project is a **static, client-side planning tool** (an HTML file, two Excel workbooks, and a VBA macro module).
It contains no server, no database, no authentication, and no user-submitted data.
The attack surface is therefore limited, but issues such as stored XSS in the HTML app or macro-level risks in
the Excel workbooks are still in scope.

## Reporting a Vulnerability

**Do not open a public GitHub issue for security vulnerabilities.**

Report security concerns privately by emailing:

> **kenthepaine@gmail.com**

Please include:

- A clear description of the vulnerability and its potential impact
- Steps to reproduce (file name, browser/Excel version, OS)
- Any proof-of-concept code or screenshots, if applicable

You can expect an acknowledgement within **72 hours** and a status update within **7 days**.

## Supported Versions

Only the latest code on the `master` branch is supported.
No back-ported patches are issued for older commits.

## Disclosure Policy

- We will work with you to understand and confirm the issue.
- We will prepare a fix and release it on `master`.
- We ask that you allow **14 days** from confirmation before any public disclosure.
- Credit will be given in the release notes if you wish.

## Out of Scope

- Issues in third-party software (Excel, browser engines, operating system).
- Social-engineering or phishing attacks unrelated to this project's code.
- Theoretical vulnerabilities without a realistic attack path.
