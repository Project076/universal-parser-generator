# Bank Statement Parser

Local web application for normalizing bank statements to Excel and verifying their financial integrity.

## Run

Use the bundled Python runtime:

```powershell
& 'C:\Users\Hp\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' app.py
```

Then open `http://localhost:8080` in a browser.

## AI cost guard

When `OPENAI_API_KEY` is configured, UPG uses the frontier parser-planning
model only for unfamiliar layouts. It first reuses a certified profile or
deterministic PDF geometry where possible; those paths make no AI request.
For a new layout, one job is capped at four evidence-led AI decisions by
default (classify, build, diagnose, and one material repair). Set
`UPG_MAX_AI_CALLS_PER_JOB` between `2` and `6` only when an operator needs a
different limit. Validation is never relaxed by this limit.

## Validation

The app extracts declared opening and closing balances from the source separately from transactions. It validates:

`opening balance - parsed withdrawals + parsed deposits = declared closing balance`

It also performs narration validation: each parsed narration is normalized for harmless case, spacing, and punctuation differences, then checked against the original statement source. The output Excel workbook is offered only when both validation gates pass. It contains a `Transactions` sheet and a `Validation` sheet.

After both checks pass, the app saves a reusable parser profile that contains only the validated header layout and column mapping. A later statement with the same layout is extracted using that existing parser; statement values, narration text, and balances are never saved in the profile.

## Supported first-pass inputs

- CSV and Excel column-based statements
- TXT statements with delimited transaction rows
- Machine-readable PDFs with recognizable transaction tables, including Axis- and PNB-style layouts

Scanned PDFs and image-only documents require an OCR connector. Password-protected PDFs require their password for that upload only; it is held in memory and is never saved. DOCX files need a document-text extraction profile. The application does not export a result for any source that fails either validation gate.
