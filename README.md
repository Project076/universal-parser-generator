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
For a new layout, one job uses at most two evidence-led AI decisions by
default: one measured layout blueprint and, only after a concrete failed
candidate, one targeted repaired layout map. Deterministic source evidence,
certified capabilities, and validation do not consume this allowance. Set
`UPG_MAX_AI_CALLS_PER_JOB` between `1` and `3` only when an operator needs a
different limit. Validation is never relaxed by this limit.

## Specialist pipeline architecture

UPG defines exactly 50 named logical specialist workers. They are pipeline
stages, not 50 operating-system processes. Each specialist can retrieve only
the certified knowledge relevant to its domain. For example, the furniture
removal specialist uses page-header, footer, summary, disclaimer, and
cross-page-boundary knowledge; it does not retrieve narration or amount rules.

If a stage is blocked, UPG produces a structured escalation containing the
blocked step, measured source evidence, applicable certified rules, attempted
candidates, and exact validation failures. The supervisor may propose only a
versioned addendum for that specialist. It cannot overwrite an earlier
certified rule, import another layout's geometry, invent statement values, or
weaken validation. Processing then replays from the earliest affected step and
must pass the complete independent certification gates before a parser is
saved.

Authenticated operators can inspect the registry with
`GET /pipeline-specialists`. The response contains worker names, domains,
scoped libraries, and repair policy, but never statement data or parser code.

The current deployment continues to use the configured OpenAI model for
agentic decisions. A local Qwen model should be connected as a separate
inference service before making it the primary model; the existing 1 GB web
container is intentionally not used to host the model.

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
