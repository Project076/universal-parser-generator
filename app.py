"""Local, dependency-light bank-statement normalization web application."""
from __future__ import annotations

import csv
from concurrent.futures import ThreadPoolExecutor
import io
import json
import hashlib
import hmac
import os
import re
import shutil
import threading
import time
from difflib import SequenceMatcher
import urllib.error
import urllib.request
import uuid
from datetime import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import openpyxl
from openpyxl.styles import Font, PatternFill
import pdfplumber
from pypdf import PdfReader

ROOT = Path(__file__).parent
# Railway's container filesystem is replaced on a restart.  When a persistent
# volume is mounted at /data, keep mutable job state, uploads, exports and
# learned profiles there.  Local development continues to use the project
# folder without any setup.
DATA_ROOT = Path(os.environ.get("UPG_DATA_DIR") or ("/data" if Path("/data").exists() else ROOT))
UPLOADS = DATA_ROOT / "uploads"
EXPORTS = DATA_ROOT / "exports"
PROFILES = DATA_ROOT / "profiles"
JOBS_DIR = DATA_ROOT / "jobs"
LEARNING_LEDGER = PROFILES / "validated_learning.json"
UPG_API_KEY = os.environ.get("UPG_API_KEY", "")
UPG_WEBHOOK_URL = os.environ.get("UPG_WEBHOOK_URL", "")
UPG_WEBHOOK_SECRET = os.environ.get("UPG_WEBHOOK_SECRET", "")
for folder in (UPLOADS, EXPORTS, PROFILES, JOBS_DIR):
    folder.mkdir(exist_ok=True)

# Seed a new persistent volume with the validated profiles shipped in the
# repository. Never overwrite a learned profile already present on the volume.
if DATA_ROOT != ROOT:
    bundled_profiles = ROOT / "profiles"
    if bundled_profiles.exists():
        for source in bundled_profiles.glob("*.json"):
            target = PROFILES / source.name
            if not target.exists():
                shutil.copy2(source, target)

JOBS: dict[str, dict] = {}
JOBS_LOCK = threading.RLock()
# Heavy PDF work is deliberately bounded.  The service accepts many requests
# immediately, but only a sized number may run in this container at once.
# Railway's 1 GB starter container must remain at one worker.  A larger service
# can opt in to 2-4 workers with UPG_WORKERS; we cap the value to avoid a typo
# turning 30 simultaneous uploads into an out-of-memory restart.
try:
    WORKER_CAPACITY = min(4, max(1, int(os.environ.get("UPG_WORKERS", "1"))))
except ValueError:
    WORKER_CAPACITY = 1
try:
    # A parser must keep retrying until it validates, but no single difficult
    # statement may occupy a multi-tenant worker forever.  One full retry
    # round per lease gives every queued tenant a fair turn; the job resumes
    # automatically from its persisted investigation state.
    WORKER_LEASE_ROUNDS = min(3, max(1, int(os.environ.get("UPG_WORKER_LEASE_ROUNDS", "1"))))
except ValueError:
    WORKER_LEASE_ROUNDS = 1
try:
    # A client that closes/cancels Fix with AI must not reserve a queue turn
    # forever.  Normal browser/API polling refreshes this lease automatically.
    JOB_CLIENT_LEASE_SECONDS = max(60, int(os.environ.get("UPG_CLIENT_LEASE_SECONDS", "300")))
except ValueError:
    JOB_CLIENT_LEASE_SECONDS = 300
JOB_EXECUTOR = ThreadPoolExecutor(max_workers=WORKER_CAPACITY, thread_name_prefix="upg-parser")
ACTIVE_JOB_IDS: set[str] = set()
# Per-upload, in-memory evidence cache. This speeds candidate retries without
# persisting statement text or transactions as long-term learning data.
EXTRACTION_CACHE: dict[tuple[str, str], tuple[list[list[object]], str]] = {}
EXTRACTION_CACHE_LOCK = threading.Lock()
PDF_TEXT_CACHE: dict[str, str] = {}
PDF_SAMPLE_CACHE: dict[str, str] = {}
# Passwords are request-scoped, held only in memory, and are never written to
# profiles, learning, exports, logs, or webhook payloads.
PDF_PASSWORD_CACHE: dict[str, str] = {}

def job_file(job_id: str) -> Path:
    return JOBS_DIR / f"{job_id}.json"

def persist_job_locked(job_id: str) -> None:
    """Atomically save non-secret job state for restart recovery."""
    job = dict(JOBS.get(job_id, {}))
    # Passwords are deliberately never part of a job record.
    job.pop("password", None)
    job.pop("pdf_password", None)
    job["updated_at"] = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    temporary = job_file(job_id).with_suffix(".tmp")
    temporary.write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
    temporary.replace(job_file(job_id))

def checkpoint_job(job_id: str) -> None:
    with JOBS_LOCK:
        persist_job_locked(job_id)

def replace_job(job_id: str, value: dict) -> None:
    with JOBS_LOCK:
        JOBS[job_id] = value
        persist_job_locked(job_id)

def patch_job(job_id: str, **changes: object) -> dict:
    with JOBS_LOCK:
        job = JOBS.get(job_id, {})
        job.update(changes)
        JOBS[job_id] = job
        persist_job_locked(job_id)
        return job

def queue_snapshot_locked() -> tuple[int, int]:
    """Return this job's FIFO position and the total queued depth.

    The executor is intentionally bounded, so a truthful position is more
    useful than showing every waiting tenant an unqualified "processing".
    """
    queued = sorted(
        (item for item in JOBS.items() if item[1].get("status") == "queued"),
        key=lambda item: (str(item[1].get("queued_at") or item[1].get("submitted_at") or ""), item[0]),
    )
    return len(queued), len(ACTIVE_JOB_IDS)

def refresh_queue_positions_locked() -> None:
    queued = sorted(
        (item for item in JOBS.items() if item[1].get("status") == "queued"),
        key=lambda item: (str(item[1].get("queued_at") or item[1].get("submitted_at") or ""), item[0]),
    )
    for position, (queued_id, job) in enumerate(queued, start=1):
        job["queue_position"] = position
        job["queue_depth"] = len(queued)
        job["worker_capacity"] = WORKER_CAPACITY
        JOBS[queued_id] = job
        persist_job_locked(queued_id)

def touch_worker(job_id: str, **changes: object) -> None:
    """Persist a liveness heartbeat for long PDF and AI operations."""
    changes["worker_heartbeat_at"] = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    patch_job(job_id, **changes)

def utc_now() -> datetime:
    return datetime.utcnow()

def timestamp_now() -> str:
    return utc_now().isoformat(timespec="seconds") + "Z"

def timestamp_age_seconds(value: object) -> float:
    try:
        return max(0.0, (utc_now() - datetime.fromisoformat(str(value).replace("Z", "+00:00")).replace(tzinfo=None)).total_seconds())
    except (TypeError, ValueError):
        return float("inf")

def touch_client(job_id: str) -> None:
    """Refresh a job lease, throttled to avoid writing on every poll."""
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job or job.get("status") in {"completed", "failed", "cancelled"}:
            return
        if timestamp_age_seconds(job.get("client_heartbeat_at")) < 30:
            return
        job["client_heartbeat_at"] = timestamp_now()
        JOBS[job_id] = job
        persist_job_locked(job_id)

def cancel_job(job_id: str, reason: str = "Cancelled by the requesting client.") -> bool:
    """Cancel queued work now, or request a safe stop between active steps."""
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return False
        if job.get("status") in {"completed", "failed", "cancelled"}:
            return True
        job.update({"cancel_requested": True, "cancel_reason": reason,
                    "client_heartbeat_at": timestamp_now()})
        if job.get("status") == "queued":
            job.update({"processing": False, "valid": False, "status": "cancelled",
                        "message": "UPG job was cancelled before a worker started it."})
        else:
            job["message"] = "UPG cancellation requested; the active worker will stop safely after its current extraction step."
        JOBS[job_id] = job
        persist_job_locked(job_id)
        refresh_queue_positions_locked()
        return True

def job_cancel_requested(job_id: str) -> bool:
    with JOBS_LOCK:
        job = JOBS.get(job_id, {})
        return bool(job.get("cancel_requested") or job.get("status") == "cancelled")

def expire_abandoned_jobs() -> None:
    """Remove work whose requesting browser/API client has stopped polling."""
    candidates: list[str] = []
    with JOBS_LOCK:
        for job_id, job in JOBS.items():
            if job.get("status") not in {"queued", "processing", "pending"}:
                continue
            last_seen = job.get("client_heartbeat_at") or job.get("submitted_at")
            if timestamp_age_seconds(last_seen) > JOB_CLIENT_LEASE_SECONDS:
                candidates.append(job_id)
    for job_id in candidates:
        cancel_job(job_id, "Cancelled automatically because the requesting client stopped monitoring this job.")

def queue_supervisor() -> None:
    while True:
        try:
            expire_abandoned_jobs()
        except Exception:
            pass
        time.sleep(15)

CANONICAL = ["date", "narration", "withdrawal", "deposit", "instrument_number", "balance"]
FIVE_MINUTES_MS = 300000
AI_MODEL = "gpt-5.6-sol"
DIAGNOSTIC_RULE_LIBRARY = {
    "value_date": "Use Value Date as the output date when both posting and value dates exist.",
    "dual_date_running_balance": "For layouts with both posting and Value Date, begin records only at the posting-date column but export the Value Date and infer amounts from running-balance changes.",
    "summary_endpoints": "Use the printed statement-summary opening and closing balances as endpoints.",
    "balance_delta": "Classify a single unsigned amount column from running-balance movement.",
    "continuation_merge": "Join narration and transaction fragments across rows or pages.",
    "footer_exclusion": "Exclude totals, closing labels, disclaimers, and page furniture.",
    "reverse_order": "Reverse newest-first statements before reconciliation.",
    "source_coverage": "Reject partial extracts and require all detectable source records.",
    "truncated_table_date": "Repair a date cell only when the original source proves its missing final year digit; retain the row's actual amount and balance, never discard it.",
    "signed_balance_text": "Use dated text blocks with Dr/Cr running balances, then classify debit or credit from each balance movement.",
    "headerless_layout": "Treat a repeated or missing table header as layout evidence, not as a transaction; infer columns only from dated rows and running balances.",
    "multi_page_continuation": "Preserve a dated transaction whose narration or amount cells continue across a page boundary, excluding page headers and footers between its parts.",
    "summary_total_warning": "Keep inconsistent printed debit or credit totals as a warning when transaction count, balance chain, and endpoint reconciliation independently pass.",
    "amount_balance_consistency": "When a source row visibly prints its transaction amount, require that amount to agree with the running-balance movement; reject a layout that only reconciles after replacing source amounts.",
    "unordered_balance_chain": "When a statement prints valid dated rows but their on-page order is not the running-balance order, reconstruct direction and order only from unique amount-and-balance links; reject ambiguity or any incomplete chain.",
}
PARSER_GENERATOR_POLICY = """
Bank statement extraction policy:
- Use the original PDF's word coordinates, column x-ranges, and row y-ranges as the primary evidence for creating and reusing a parser profile. Use extracted text only to join narration continuations, provide validation evidence, or as a fallback when usable PDF geometry is absent.
- B/F, opening balance, and brought-forward entries are statement metadata, never transactions.
- A row without a valid transaction date is statement furniture, a page/transaction total, or a balance label; never treat it as a transaction merely because it contains amounts.
- Use the statement opening balance when printed. If it is absent, derive it from the first real transaction's signed running balance minus its deposit plus its withdrawal.
- A printed statement-level opening or closing balance overrides any inferred value. Otherwise, the closing balance is the signed running balance of the last real transaction, never a page total, grand total, available amount, or other footer balance.
- Normalize Cr balances as positive and Dr balances as negative. A signed increase is a deposit; a signed decrease is a withdrawal.
- Balance-chain validation is mandatory for every transaction with a running balance: previous balance = current balance + current withdrawal - current deposit. Equivalently, current balance = previous balance - withdrawal + deposit. Do not release a parser when any transaction balance is missing or breaks this chain.
- Exception: if any transaction proves that the source running-balance chain is unreliable, do not use that column for a normal balance-chain pass or transaction classification. Certify only if parsed withdrawals and deposits exactly equal the printed statement totals, while narration coverage and transaction count pass. Form assumed endpoints from one available transaction balance and the verified totals, label them assumed, and require manual source review. If totals or independent evidence are inconsistent, withhold the parser.
- Transaction-count validation is mandatory: independently count source records that have a transaction date plus amount/running-balance evidence, and require exactly that many parsed transactions. More or fewer parsed rows is a failure even if balances reconcile.
- Particulars must contain only actual transaction narration. Do not put monetary amounts, blank-field substitutes, page headers, account-holder text, totals, or statement furniture in it. If the source Particulars is blank, output a blank narration.
- Join continuation fragments of the same transaction across pages and ignore repeated headers/footers.
- When a transaction crosses a page boundary, remove the intervening page total, disclaimer, bank header, account-holder block, and repeated statement title before joining the remaining fragments. The date, narration, amount, and running balance must remain one transaction.
- For text-layout statements with no reliable table geometry, detect whether each dated transaction is a multi-line block: a narration/reference line followed by an amount and running-balance line. Do not treat a statement-period date before the transaction heading as a transaction. If both posting Date and Value Date columns exist, only the posting-date column starts a record; output the Value Date without treating it as a second transaction.
- If a combined text layer is unreliable, also test a page-by-page text candidate and join its dated rows before validation. This candidate must still reconcile across the complete statement.
- When a layout has one unsigned amount column rather than separate debit/credit columns, infer withdrawal or deposit only from the signed change between consecutive running balances. Ignore long reference IDs, account numbers, dates, timestamps, page numbers, and footer postal codes as monetary values.
- A running balance of exactly zero is valid even when its usual Dr/Cr suffix is omitted. Treat a dated row ending in an explicit zero amount as a real transaction only when it has the row's amount and running-balance evidence; derive its direction from the balance chain.
- Some statements group or print same-date transactions out of running-balance order. If each row visibly has an amount and a balance, reconstruct the order only when each next row is uniquely proved by `next balance = previous balance - withdrawal + deposit`, from the declared opening balance through the declared closing balance. Preserve each row's actual source amount and narration; reject any ambiguous, partial, or disconnected chain.
- If an undated narration/reference fragment is immediately followed by a date, time, amount and signed running balance, join that fragment to the dated row. The date/time line is a continuation of that transaction, not a separate blank-narration transaction.
- A table extractor can truncate or misplace a date while still correctly reading the amount and running balance. For a date-like cell with a missing final year digit, consult the original source text for the immediately following digit and repair it only when that exact completion is present. Keep that transaction; do not discard it merely because the table cell is malformed. If its source Particulars is blank, export a blank Particulars field rather than inventing text.
- Cut a transaction block before closing-balance labels, transaction totals, grand totals, available-balance labels, disclaimers, and other footer furniture. The printed closing balance is validation evidence, never a transaction amount.
- Before generating a new parser, compare the layout with saved validated profiles. For a related layout, create an addendum that inherits the stable mapping and changes only the differing fields. Never overwrite or regress the older parser.
- For every upload, including a large PDF, run the exact existing validated parser/profile first. For a large PDF, combine that saved mapping with sampled original-PDF geometry. Create a new parser or addendum only after the existing parser has been fully extracted and has failed a release validation.
- Keep trying safe candidate strategies - saved parser, related-profile addendum, an AI-generated source-layout profile, detected table layout, signed or unsigned text running-balance layout, and chronological/reverse-chronological order - until one passes every validation gate. Do not stop after the first failed candidate and never export a partial or unreconciled result.
- A profile may be saved or Excel released only after every narration is traceable to the source and both the full financial reconciliation and each running-balance step pass. Printed debit/credit summary totals are an additional check when present, but a discrepancy is a warning rather than a release gate because some source statements print incorrect totals.
- A printed-total mismatch may be reported as a warning only when it is small enough to plausibly be a source-summary error. A materially divergent parsed total is a hard failure: it indicates that references, dates, or other non-monetary text may have been read as money.
- A partial extraction is never valid. A candidate must account for every detectable source transaction record; a shorter subset that happens to reconcile is a failure. Narration verification is one-to-one source coverage, not merely a loose text substring check.
- Self-healing is constrained to parser-profile addenda: use failed validation evidence to propose a revised header/column layout, test it from the original source, and retain it only after every release gate passes. Never modify application code, invent transactions, or weaken a validation to make a result pass.
- For long PDFs, create or repair a layout profile from a representative sample: first seven pages, seven pages centered around the middle, and last seven pages, plus adjacent boundary pages so transactions split across sampled-page edges remain visible. Apply the resulting candidate to the complete statement and validate the whole source before learning or export.
"""
ALIASES = {
    "date": ["date", "transaction date", "txn date", "value date"],
    "narration": ["narration", "particular", "particulars", "description", "remarks", "details"],
    "withdrawal": ["withdrawal", "withdrawal amount", "debit", "debit amount", "dr", "amount debited"],
    "deposit": ["deposit", "deposit amount", "credit", "credit amount", "cr", "amount credited"],
    "instrument_number": ["cheque number", "check number", "instrument number", "instrument", "cheque no", "chq no", "ref no", "chq. / ref no."],
    "balance": ["balance", "closing balance", "closing balance*", "running balance", "available balance"],
    "amount": ["amount", "amount(inr)", "transaction amount"],
    "transaction_type": ["type", "dr/cr", "transaction type"],
}

HTML = r'''<!doctype html><html><head><meta charset="utf-8"><title>Statement Normalizer</title><style>
body{font-family:system-ui;max-width:860px;margin:50px auto;color:#172033;background:#f5f7fb}.card{background:white;padding:30px;border-radius:16px;box-shadow:0 4px 22px #1223}h1{margin-top:0}label{display:block;margin:16px 0 5px;font-weight:650}input,button{font:inherit;padding:10px}input{width:100%;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:7px}button{margin-top:22px;background:#0f766e;color:white;border:0;border-radius:8px;cursor:pointer}.hint{color:#52606d}.result{margin-top:20px;padding:16px;border-radius:8px}.ok{background:#dcfce7}.fail{background:#fee2e2}.field{display:grid;grid-template-columns:1fr 1fr;gap:15px}</style></head><body><main class="card"><h1>Bank Statement Normalizer</h1><p class="hint">Upload a statement. Excel is created only after the declared balances reconcile with parsed transactions. For unfamiliar layouts, the configured AI parser generator may inspect the layout to create a profile; no export is released unless both checks pass.</p><form id="form"><label>Statement file</label><input name="file" type="file" accept=".csv,.xlsx,.xls,.txt,.pdf,.doc,.docx" required><div class="field"><div><label>Opening balance (optional fallback)</label><input name="opening" placeholder="Extracted from source when present"></div><div><label>Closing balance (optional fallback)</label><input name="closing" placeholder="Extracted from source when present"></div></div><label>PDF password (only if protected)</label><input name="password" type="password" autocomplete="off" placeholder="Used only in memory for this upload"><button>Parse and validate</button></form><section id="result"></section></main><script>
const f=document.querySelector('#form'), r=document.querySelector('#result'), submit=f.querySelector('button');let activeJob=null;
function show(d){const label=d.valid?'Validated':d.processing?'UPG is retrying':d.interrupted?'UPG job interrupted':'Not validated';r.className='result '+(d.valid?'ok':d.processing||d.interrupted?'':'fail');r.innerHTML=`<strong>${label}</strong><br>${d.message}`+(d.download?`<br><br><a href="${d.download}">Download validated Excel</a>`:'')}
async function poll(job){const d=await (await fetch('/status/'+job)).json();if(job!==activeJob)return;show(d);if(d.processing)setTimeout(()=>poll(job),2500);else{activeJob=null;submit.disabled=false;submit.textContent='Parse and validate'}}
f.onsubmit=async e=>{e.preventDefault();if(activeJob)return;r.className='result';r.innerHTML='<strong>UPG is retrying</strong><br>Creating and validating parser candidates.';submit.disabled=true;submit.textContent='UPG is working...';try{const d=await (await fetch('/parse',{method:'POST',body:new FormData(f)})).json();activeJob=d.job||null;show(d);if(d.processing)poll(d.job);else{submit.disabled=false;submit.textContent='Parse and validate'}}catch(err){activeJob=null;submit.disabled=false;submit.textContent='Parse and validate';r.className='result fail';r.innerHTML='<strong>Unable to start UPG</strong><br>The parser retry job could not start.'}};
</script></body></html>'''

def money(value: object) -> Decimal | None:
    if value is None or str(value).strip() == "": return None
    s = str(value).strip().replace(",", "").replace("₹", "").replace("$", "")
    s = re.sub(r"\s+", "", s)
    # Coordinate extraction can leave a transaction value followed by a page
    # total in the same cell. The first monetary token belongs to the row.
    token = re.match(r"-?\d+(?:\.\d{1,2})?(?:DR|CR)?", s, re.I)
    if token:
        s = token.group()
    suffix = re.search(r"(DR|CR)$", s, re.I)
    neg = (s.startswith("(") and s.endswith(")")) or bool(suffix and suffix.group(1).upper() == "DR")
    s = re.sub(r"(?:DR|CR)$", "", s.strip("()"), flags=re.I).strip()
    try: return (-1 if neg else 1) * Decimal(s)
    except InvalidOperation: return None

def indian_amount(value: Decimal | int | float) -> str:
    """Format monetary values with Indian lakh/crore grouping."""
    amount = Decimal(value).quantize(Decimal(".01"), rounding=ROUND_HALF_UP)
    sign = "-" if amount < 0 else ""
    whole, fraction = f"{abs(amount):.2f}".split(".")
    if len(whole) <= 3:
        grouped = whole
    else:
        tail, rest = whole[-3:], whole[:-3]
        pairs = []
        while rest:
            pairs.append(rest[-2:]); rest = rest[:-2]
        grouped = ",".join(reversed(pairs)) + "," + tail
    return sign + grouped + "." + fraction

def norm(s: object) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).lower())

def normalize_narration(s: object) -> str:
    """Comparison form: ignores harmless whitespace/case/punctuation variations only."""
    return norm(s)

def clean_narration(s: str) -> str:
    """Keep only the statement's Particulars, never amounts or page furniture."""
    lines = []
    furniture = ("jammu and kashmir bank", "statement of account", "page total", "grand total", "printed by", "ifsc code", "micr code", "unless the constituent", "customer id", "currency code", "a/c no", "interest rate", "no nomination", "c kyc", "ckyc")
    for line in s.splitlines():
        compact = re.sub(r"\s+", " ", line).strip()
        if compact and not any(marker in compact.lower() for marker in furniture): lines.append(compact)
    result = " ".join(lines)
    result = re.sub(r"-{5,}", " ", result)
    # Some PDF text layers collapse a multi-line page header into a single
    # string immediately before the continuation. Remove that header prefix.
    result = re.sub(r"(?is)\ba/c no\s*:.*?(?:no nomination available for the account|currency code\s*:\s*inr)\s*", "", result)
    result = re.sub(r"(?is)\bto:\s*.*?\.\s*,\s*", "", result)
    # Time is a separate statement column in several compact PDF text layers,
    # not part of Particulars.
    result = re.sub(r"\b\d{2}:\d{2}:\d{1,2}\b", "", result)
    result = re.sub(r"\s+-?[\d,]+(?:\.\d{1,2})?\s*$", "", result).strip()
    return "" if re.fullmatch(r"-?[\d,]+(?:\.\d{1,2})?", result) else result

def map_headers(headers: list[object]) -> dict[str, int]:
    mapped = {}
    for canonical, possibilities in ALIASES.items():
        for i, h in enumerate(headers):
            header = norm(h)
            if header in {norm(x) for x in possibilities}:
                mapped[canonical] = i; break
    # Header wording varies by bank. These semantic fallbacks allow the parser
    # generator to discover a new layout without a bank-specific code change.
    for i, h in enumerate(headers):
        header = norm(h)
        if "balance" in header: mapped.setdefault("balance", i)
        elif "withdraw" in header or "debit" in header or header.endswith("dr"): mapped.setdefault("withdrawal", i)
        elif "deposit" in header or "credit" in header or header.endswith("cr"): mapped.setdefault("deposit", i)
        elif any(word in header for word in ("narration", "particular", "remark", "description", "detail")): mapped.setdefault("narration", i)
        elif any(word in header for word in ("cheque", "check", "instrument", "refno")): mapped.setdefault("instrument_number", i)
        elif header in ("date", "trandate", "transactiondate", "valuedate"): mapped.setdefault("date", i)
        elif "amount" in header: mapped.setdefault("amount", i)
        elif header == "type": mapped.setdefault("transaction_type", i)
    value_date = next((i for i, h in enumerate(headers) if norm(h) in {"valuedate", "valuedt"}), None)
    if value_date is not None:
        mapped["date"] = value_date
    return mapped

def text_layout_fingerprint(raw: str) -> str:
    """Privacy-safe structural signature for text-only statement layouts."""
    header = re.search(r"(?im)^\s*date\s+.{0,120}(?:balance|amount)\s*$", raw)
    features = {
        "header": norm(header.group(0)) if header else "",
        "opening": bool(re.search(r"\b(?:opening balance|b/f)\b", raw, re.I)),
        "closing": bool(re.search(r"\bclosing balance\b", raw, re.I)),
        "month_dates": len(re.findall(r"\b\d{2}-[A-Za-z]{3}-\d{4}\b", raw)) > 3,
        "numeric_dates": len(re.findall(r"\b\d{2}[-/]\d{2}[-/]\d{4}\b", raw)) > 3,
    }
    return hashlib.sha256(json.dumps(features, sort_keys=True).encode()).hexdigest()[:16]

def profile_id(headers: list[object], layout_fingerprint: str = "") -> str:
    """Stable, privacy-safe layout signature: header labels only, never statement data."""
    signature = "|".join(norm(h) for h in headers) + "|" + layout_fingerprint
    return hashlib.sha256(signature.encode()).hexdigest()[:16]

def generated_canonical_headers(headers: list[object]) -> bool:
    """Do not learn a profile from our own synthetic fallback header."""
    return [norm(h) for h in headers] == [norm(h) for h in ["Date", "Narration", "Withdrawal", "Deposit", "Instrument Number", "Balance"]]

def load_profile(headers: list[object], layout_fingerprint: str = "") -> dict[str, int] | None:
    if generated_canonical_headers(headers) and not layout_fingerprint: return None
    profile = PROFILES / f"{profile_id(headers, layout_fingerprint)}.json"
    if not profile.exists(): return None
    try:
        data = json.loads(profile.read_text(encoding="utf-8"))
        return {k: int(v) for k, v in data["columns"].items()}
    except (OSError, ValueError, KeyError): return None

def certified_javascript_code(headers: list[object], strategy: str | None) -> tuple[str, str]:
    """Build the safe, self-contained JavaScript contract consumed by BS Analyzer.

    This is deliberately a deterministic generator, not AI-authored executable
    code. It emits only a reviewed line parser and a conservative layout
    detector. The source transaction extraction has already passed UPG's
    financial, narration, balance-chain, and count gates before this function
    is called.
    """
    header_text = " ".join(norm(header) for header in headers)
    anchors = []
    for group in (("date",), ("narration", "particular", "description"), ("balance",), ("withdrawal", "debit", "dr"), ("deposit", "credit", "cr")):
        found = next((word for word in group if word in header_text), None)
        if found:
            anchors.append(found)
    # These generic anchors still require all core transaction concepts, so a
    # profile cannot accidentally claim a non-statement document.
    anchors = anchors or ["date", "balance"]
    anchors_json = json.dumps(anchors)
    detection = """function detect(text) {
  const normalized = String(text || '').toLowerCase().replace(/\\s+/g, ' ');
  const anchors = __ANCHORS__;
  const hasNarration = normalized.includes('narration') || normalized.includes('particular') || normalized.includes('description');
  const hasDebit = normalized.includes('withdrawal') || normalized.includes('debit') || normalized.includes('dr');
  const hasCredit = normalized.includes('deposit') || normalized.includes('credit') || normalized.includes('cr');
  return normalized.includes('date') && normalized.includes('balance') && hasNarration && hasDebit && hasCredit && anchors.length >= 2;
}""".replace("__ANCHORS__", anchors_json)
    if strategy in {"running_balance_text", "unsigned_running_balance_text", "value_date_unsigned", "page_text_unsigned"}:
        parser = """function parse(text, options) {
  const blocks = String(text || '').split(/(?=^\\s*\\d{1,2}[\\/-]\\d{1,2}[\\/-]\\d{2,4}\\b)/m);
  const dateRe = /^\\s*(\\d{1,2}[\\/-]\\d{1,2}[\\/-]\\d{2,4})\\b/;
  const moneyRe = /-?\\d[\\d,]*\\.\\d{2}/g;
  const asNumber = (value) => Number(String(value).replace(/,/g, ''));
  const rows = []; let previous = Number(options && options.openingBalance); if (!Number.isFinite(previous)) previous = null;
  for (const block of blocks) {
    const match = block.match(dateRe); if (!match || /\\bB\\/F\\b/i.test(block) || /^(?:page total|grand total)/im.test(block)) continue;
    const signed = [...block.matchAll(/(-?[\\d,]+\\.\\d{2})\\s*(Dr|Cr)\\b/ig)]; const values = [...block.matchAll(moneyRe)];
    if (!signed.length || values.length < 2) continue;
    const balanceToken = signed[signed.length - 1], balance = asNumber(balanceToken[1]) * (balanceToken[2].toLowerCase() === 'dr' ? -1 : 1);
    const amountToken = values.filter((value) => value.index < balanceToken.index).pop(); if (!amountToken) continue;
    const amount = Math.abs(asNumber(amountToken[0])); const particulars = block.slice(match[0].length, amountToken.index).replace(/\\s+/g, ' ').trim();
    const refs = particulars.match(/\\b\\d{6,}\\b/g) || []; const p = match[1].split(/[\\/-]/); const year = p[2].length === 2 ? `20${p[2]}` : p[2];
    const delta = previous == null ? amount : balance - previous;
    rows.push({date:`${p[0].padStart(2,'0')}/${p[1].padStart(2,'0')}/${year}`, particulars, withdrawal:delta < 0 ? -delta : 0, deposit:delta > 0 ? delta : 0, balance, chqNo:refs.length ? refs[refs.length - 1] : ''}); previous = balance;
  } return rows;
}"""
        return detection, parser
    parser = """function parse(text, options) {
  const lines = String(text || '').split(/\\r?\\n/);
  const dateRe = /^\\s*(\\d{1,2}[\\/-]\\d{1,2}[\\/-]\\d{2,4})\\b/;
  const moneyRe = /-?\\d[\\d,]*\\.\\d{2}\\b/g;
  const ignored = /^(page\\s+\\d+|generation date|hdfc bank|statement summary|\\*\\*end of statement|date\\s+.*(?:balance|deposit|credit)|account branch|address\\s*:|contents of this statement)/i;
  const amount = (value) => {
    if (value == null) return null;
    const n = Number(String(value).replace(/,/g, ''));
    return Number.isFinite(n) && n !== 0 ? Math.abs(n) : null;
  };
  const date = (value) => {
    const p = String(value).split(/[\\/-]/);
    if (p.length !== 3) return String(value);
    const year = p[2].length === 2 ? `20${p[2]}` : p[2];
    return `${p[0].padStart(2, '0')}/${p[1].padStart(2, '0')}/${year}`;
  };
  const rows = [];
  let prefix = [];
  for (const rawLine of lines) {
    const line = String(rawLine || '').trim();
    if (!line || ignored.test(line)) { if (ignored.test(line)) prefix = []; continue; }
    const match = line.match(dateRe);
    if (!match) {
      // Preserve only transaction-like wrapped narration, never page furniture.
      if (/^(RTGS|NEFT|IMPS|UPI|CHQ|CASH|ADHOC|EMP|EVP|POS|CARD)/i.test(line)) prefix.push(line);
      continue;
    }
    const rest = line.slice(match[0].length);
    const money = [...rest.matchAll(moneyRe)];
    if (money.length < 3) { prefix = []; continue; }
    const values = money.slice(-3).map((item) => amount(item[0]));
    const firstAmountAt = money[money.length - 3].index;
    let detail = rest.slice(0, firstAmountAt)
      .replace(/\\b\\d{1,2}[\\/-]\\d{1,2}[\\/-]\\d{2,4}\\b/g, ' ')
      .replace(/\\s+/g, ' ').trim();
    const refs = detail.match(/\\b[A-Z0-9]{8,}\\b/g) || [];
    const chqNo = refs.length ? refs[refs.length - 1] : '';
    if (chqNo) detail = detail.replace(chqNo, ' ').replace(/\\s+/g, ' ').trim();
    const particulars = [...prefix, detail].join(' ').replace(/\\s+/g, ' ').trim();
    rows.push({ date: date(match[1]), particulars, withdrawal: values[0], deposit: values[1], balance: values[2], chqNo });
    prefix = [];
  }
  return rows;
}"""
    return detection, parser

def save_profile(headers: list[object], columns: dict[str, int], parent_profile: str | None = None, strategy: str | None = None, self_healed: bool = False, layout_fingerprint: str = "", diagnostic_rules: list[str] | None = None, validation: dict | None = None, bank_name: str = "Unknown", format_name: str = "PDF Statement") -> str:
    """Persist only validated, privacy-safe layout learning; never source rows."""
    if generated_canonical_headers(headers) and not layout_fingerprint: return ""
    ident = profile_id(headers, layout_fingerprint)
    profile_path = PROFILES / f"{ident}.json"
    try: prior = json.loads(profile_path.read_text(encoding="utf-8"))
    except (OSError, ValueError): prior = {}
    observations = int(prior.get("validated_observations", 0)) + 1
    detection_code, parser_code = certified_javascript_code(headers, strategy)
    data = {"version": int(prior.get("version", 0)) + 1, "header_signature": [str(h) for h in headers], "layout_fingerprint": layout_fingerprint, "columns": columns, "parent_profile": parent_profile, "validated_observations": observations, "last_validated_strategy": strategy or "detected_table", "self_healed_addendum": bool(self_healed), "diagnostic_rules": diagnostic_rules or [], "bank_name": bank_name or prior.get("bank_name", "Unknown"), "format_name": format_name or prior.get("format_name", "PDF Statement"), "detection_code": detection_code, "parser_code": parser_code, "validation": validation or {"status": "pass", "financial_pass": True, "narration_pass": True, "balance_chain_pass": True}, "certification": {"status": "certified", "source": "upg_native", "certified_at": datetime.utcnow().replace(microsecond=0).isoformat() + "Z"}}
    profile_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    # Aggregate learning intentionally contains only layout signatures and
    # validation outcomes, never account, narration, balances, or transactions.
    try: ledger = json.loads(LEARNING_LEDGER.read_text(encoding="utf-8"))
    except (OSError, ValueError): ledger = {"validated_profiles": {}}
    ledger.setdefault("validated_profiles", {})[ident] = {"observations": observations, "strategy": data["last_validated_strategy"], "self_healed_addendum": data["self_healed_addendum"], "diagnostic_rules": data["diagnostic_rules"], "parent_profile": parent_profile}
    LEARNING_LEDGER.write_text(json.dumps(ledger, indent=2), encoding="utf-8")
    return ident

def find_related_profile(headers: list[object]) -> tuple[str | None, dict[str, int]]:
    """Find a near-match without changing the original validated profile."""
    if generated_canonical_headers(headers): return None, {}
    current = "|".join(norm(h) for h in headers)
    best_id, best_columns, best_score = None, {}, 0.0
    for profile in PROFILES.glob("*.json"):
        try:
            data = json.loads(profile.read_text(encoding="utf-8"))
            if generated_canonical_headers(data.get("header_signature", [])):
                continue
            prior = "|".join(norm(h) for h in data.get("header_signature", []))
            score = SequenceMatcher(None, current, prior).ratio()
            if score > best_score:
                best_id, best_columns, best_score = profile.stem, {k: int(v) for k, v in data["columns"].items()}, score
        except (OSError, ValueError, KeyError):
            continue
    return (best_id, best_columns) if best_score >= 0.62 else (None, {})

def saved_text_strategy(path: Path) -> str | None:
    """Find the validated extraction strategy before new generation begins."""
    if path.suffix.lower() != ".pdf": return None
    try:
        # Profiles are saved from the normalized source used for extraction.
        # Look up with that same representation; using uncleaned PDF text here
        # made repeated page furniture change the signature and hid a valid
        # parser for an otherwise matching statement.
        fingerprint = text_layout_fingerprint(remove_page_furniture(cached_pdf_text(path)))
        for profile in PROFILES.glob("*.json"):
            data = json.loads(profile.read_text(encoding="utf-8"))
            if data.get("layout_fingerprint") == fingerprint:
                return data.get("last_validated_strategy")
    except (OSError, ValueError, KeyError):
        pass
    return None

def pdf_password(path: Path) -> str:
    with EXTRACTION_CACHE_LOCK:
        return PDF_PASSWORD_CACHE.get(str(path.resolve()), "")

def register_pdf_password(path: Path, password: str) -> None:
    if path.suffix.lower() != ".pdf":
        return
    key = str(path.resolve())
    with EXTRACTION_CACHE_LOCK:
        if password:
            PDF_PASSWORD_CACHE[key] = password
        else:
            PDF_PASSWORD_CACHE.pop(key, None)

def clear_pdf_password(path: Path) -> None:
    with EXTRACTION_CACHE_LOCK:
        PDF_PASSWORD_CACHE.pop(str(path.resolve()), None)

def open_pdf_reader(path: Path) -> PdfReader:
    reader = PdfReader(str(path))
    if reader.is_encrypted:
        password = pdf_password(path)
        if not password or not reader.decrypt(password):
            raise ValueError("PASSWORD_REQUIRED: This PDF is password protected. Enter its password and submit it again; UPG will not retry unreadable encrypted files.")
    return reader

def open_pdfplumber(path: Path):
    password = pdf_password(path)
    try:
        return pdfplumber.open(path, password=password or None)
    except Exception as error:
        if "password" in str(error).lower() or "encrypt" in str(error).lower():
            raise ValueError("PASSWORD_REQUIRED: This PDF is password protected. Enter its password and submit it again; UPG will not retry unreadable encrypted files.") from error
        raise

def is_large_pdf(path: Path) -> bool:
    try:
        return path.suffix.lower() == ".pdf" and len(open_pdf_reader(path).pages) > 60
    except Exception:
        return False

def prefers_running_balance_text(path: Path) -> bool:
    """Identify long text-layer statements before full-page geometry work."""
    if path.suffix.lower() != ".pdf":
        return False
    raw = remove_page_furniture(cached_pdf_text(path))
    heading = bool(re.search(r"(?i)(?:particulars|narration|description).*?(?:withdrawals?|debits?|deposits?|credits?).*?balance", raw[:6000]))
    dated_rows = len(re.findall(r"(?m)^\s*\d{2}[-/]\d{2}[-/]\d{2,4}\b", raw))
    signed_balances = bool(re.search(r"\b(?:cr|dr)\b", raw, re.I))
    return heading and dated_rows >= 10 and signed_balances

def sampled_pdf_text(path: Path) -> str:
    """Representative evidence for profile generation on very long PDFs."""
    key = str(path.resolve())
    with EXTRACTION_CACHE_LOCK:
        cached = PDF_SAMPLE_CACHE.get(key)
    if cached is not None:
        return cached
    reader = open_pdf_reader(path)
    count = len(reader.pages)
    indices = sampled_page_indices(count)
    sample = "\n".join(f"[PAGE {index + 1}]\n{reader.pages[index].extract_text() or ''}" for index in indices)
    with EXTRACTION_CACHE_LOCK:
        if len(PDF_SAMPLE_CACHE) >= 12:
            PDF_SAMPLE_CACHE.pop(next(iter(PDF_SAMPLE_CACHE)))
        PDF_SAMPLE_CACHE[key] = sample
    return sample

def sampled_pdf_geometry_evidence(path: Path) -> list[dict]:
    """Structured original-PDF layout evidence for the parser-generator AI.

    This intentionally sends table geometry, not a long character dump. Small
    PDFs contribute every page; large PDFs contribute representative first,
    middle, and last page groups (including boundary pages).
    """
    evidence: list[dict] = []
    with open_pdfplumber(path) as pdf:
        page_numbers = list(range(len(pdf.pages))) if len(pdf.pages) <= 60 else sampled_page_indices(len(pdf.pages))
        for page_number in page_numbers:
            page = pdf.pages[page_number]
            tables = page.find_tables()
            page_tables = []
            for table in tables[:3]:
                extracted = table.extract()
                if not extracted:
                    continue
                header_index = next((i for i, row in enumerate(extracted[:12]) if len(map_headers(row or [])) >= 3), 0)
                header_cells = table.rows[header_index].cells if header_index < len(table.rows) else []
                bands = [[round(float(cell[0]), 1), round(float(cell[2]), 1)] for cell in header_cells if cell]
                page_tables.append({
                    "bbox": [round(float(value), 1) for value in table.bbox],
                    "header_row_index": header_index,
                    "header": [str(value or "") for value in (extracted[header_index] if header_index < len(extracted) else [])],
                    "column_x_ranges": bands,
                    "row_count": len(extracted),
                })
            # Borderless statements (including many bank exports) have no
            # `find_tables()` result. Preserve their actual PDF geometry as
            # coordinate evidence without exposing transaction wording.
            words = page.extract_words(x_tolerance=1, y_tolerance=2)
            lines: dict[int, list[dict]] = {}
            for word in words:
                lines.setdefault(round(float(word["top"]) / 3), []).append(word)
            header_words: list[dict] = []
            for line in lines.values():
                labels = " ".join(str(word["text"]) for word in line).lower()
                if "date" in labels and "balance" in labels and any(label in labels for label in ("particular", "narration", "description")):
                    header_words = [{"label": str(word["text"]), "x0": round(float(word["x0"]), 1), "x1": round(float(word["x1"]), 1)} for word in sorted(line, key=lambda item: float(item["x0"]))]
                    break
            numeric_bands: dict[int, int] = {}
            for word in words:
                if re.fullmatch(r"-?[\d,]+(?:\.\d{1,2})?(?:Cr|Dr)?", str(word["text"]), re.I):
                    band = round(float(word["x0"]) / 10) * 10
                    numeric_bands[band] = numeric_bands.get(band, 0) + 1
            coordinate_fallback = {
                "header_word_positions": header_words,
                "numeric_column_x_ranges": [{"x0": x0, "x1": x0 + 10, "observations": count} for x0, count in sorted(numeric_bands.items(), key=lambda item: item[1], reverse=True)[:10]],
            }
            evidence.append({"page": page_number + 1, "width": round(float(page.width), 1), "height": round(float(page.height), 1), "tables": page_tables, "borderless_coordinate_evidence": coordinate_fallback})
    return evidence

def sampled_page_indices(count: int) -> list[int]:
    """Seven-page regions plus boundary context for large PDF layout learning."""
    if count <= 21:
        return list(range(count))
    middle = count // 2
    first = set(range(0, min(count, 8)))
    middle_window = set(range(max(0, middle - 4), min(count, middle + 5)))
    last = set(range(max(0, count - 8), count))
    return sorted(first | middle_window | last)

def ai_generated_profile(rows: list[list[object]], raw: str, repair_context: str = "", source_path: Path | None = None) -> tuple[int, dict[str, int]] | None:
    """Ask the embedded parser-generator AI for a new table layout, not transactions."""
    key = os.environ.get("OPENAI_API_KEY")
    if not key: return None
    schema = {
        "type": "object", "additionalProperties": False,
        "properties": {
            "header_row": {"type": "integer"},
            "columns": {
                "type": "object", "additionalProperties": False,
                "properties": {name: {"type": "integer"} for name in CANONICAL},
                "required": CANONICAL,
            },
        }, "required": ["header_row", "columns"],
    }
    geometry = sampled_pdf_geometry_evidence(source_path) if source_path and source_path.suffix.lower() == ".pdf" else []
    evidence = {"rows": rows[:35], "original_pdf_geometry_samples": geometry, "failed_validation_evidence": repair_context}
    instruction = (PARSER_GENERATOR_POLICY + "\nYou are a bank-statement parser generator and controlled self-healing planner. Identify one transaction-table header row and map "
        "its zero-based column positions to date, narration, withdrawal, deposit, instrument_number, "
        "and balance. Use the original_pdf_geometry_samples as primary evidence; do not infer a column from character order alone. Use -1 when a field is absent. If failure evidence is supplied, propose only a safe addendum to the source layout mapping; do not extract transactions, invent values, or change validation rules."
    )
    payload = {
        "model": AI_MODEL,
        "input": [{"role": "system", "content": [{"type": "input_text", "text": instruction}]},
                  {"role": "user", "content": [{"type": "input_text", "text": json.dumps(evidence)}]}],
        "text": {"format": {"type": "json_schema", "name": "bank_layout", "strict": True, "schema": schema}},
    }
    request = urllib.request.Request(
        "https://api.openai.com/v1/responses", data=json.dumps(payload).encode(),
        headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"}, method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            result = json.loads(response.read().decode())
        text = next((item["text"] for output in result.get("output", []) for item in output.get("content", []) if item.get("type") == "output_text"), "")
        generated = json.loads(text)
        header_row = int(generated["header_row"])
        columns = {name: int(index) for name, index in generated["columns"].items() if int(index) >= 0}
        if not (0 <= header_row < len(rows)) or len(columns) < 3: return None
        return header_row, columns
    except (urllib.error.URLError, urllib.error.HTTPError, KeyError, ValueError, json.JSONDecodeError):
        return None

def ai_choose_text_strategy(raw: str) -> str | None:
    """Let the parser-generator select a supported extraction path for a new layout."""
    key = os.environ.get("OPENAI_API_KEY")
    if not key: return None
    schema = {
        "type": "object", "additionalProperties": False,
        "properties": {"strategy": {"type": "string", "enum": ["running_balance_text", "unsigned_running_balance_text", "value_date_unsigned", "needs_ocr", "unsupported"]}},
        "required": ["strategy"],
    }
    payload = {
        "model": AI_MODEL,
        "input": (PARSER_GENERATOR_POLICY + "\nClassify this bank statement layout. Choose running_balance_text when dated entries have Dr/Cr running balances. Choose unsigned_running_balance_text when dated entries have unsigned running balances whose changes can infer debit or credit; choose "
            "value_date_unsigned when there are both posting Date and Value Date columns plus unsigned running balances; choose needs_ocr for image/scanned text; otherwise choose unsupported.\n\n" + raw[:50000]
        ),
        "text": {"format": {"type": "json_schema", "name": "extraction_strategy", "strict": True, "schema": schema}},
    }
    request = urllib.request.Request("https://api.openai.com/v1/responses", data=json.dumps(payload).encode(), headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=90) as response: result = json.loads(response.read().decode())
        text = next((item["text"] for output in result.get("output", []) for item in output.get("content", []) if item.get("type") == "output_text"), "")
        return json.loads(text)["strategy"]
    except (urllib.error.URLError, urllib.error.HTTPError, KeyError, ValueError, json.JSONDecodeError):
        return None

def ai_diagnose_failure(raw: str, failure: str) -> dict[str, list[str]]:
    """Create a privacy-safe, bounded investigation plan for the next retry."""
    key = os.environ.get("OPENAI_API_KEY")
    if not key: return {"rules": [], "strategies": []}
    safe_strategies = ["geometry_profile", "value_date_unsigned", "unsigned_running_balance_text", "running_balance_text", "page_text_unsigned", "detected_table", "ai_layout_addendum"]
    schema = {"type": "object", "additionalProperties": False, "properties": {
        "rules": {"type": "array", "items": {"type": "string", "enum": list(DIAGNOSTIC_RULE_LIBRARY)}, "maxItems": 5},
        "strategies": {"type": "array", "items": {"type": "string", "enum": safe_strategies}, "maxItems": 4},
    }, "required": ["rules", "strategies"]}
    prompt = PARSER_GENERATOR_POLICY + "\nAct as a senior bank-statement parser investigator. Study the source sample and failed validation evidence. Select only safe rule IDs and a priority order of already-supported candidate strategies for the next retry. Do not write executable code, invent transactions, expose source data, or weaken any validation.\nRules: " + json.dumps(DIAGNOSTIC_RULE_LIBRARY) + "\nStrategies: " + json.dumps(safe_strategies) + "\nFailure evidence: " + failure + "\nSource excerpt: " + raw[:12000]
    payload = {"model": AI_MODEL, "input": prompt, "text": {"format": {"type": "json_schema", "name": "diagnostic_rules", "strict": True, "schema": schema}}}
    request = urllib.request.Request("https://api.openai.com/v1/responses", data=json.dumps(payload).encode(), headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"}, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=90) as response: result = json.loads(response.read().decode())
        output = next((item["text"] for item_out in result.get("output", []) for item in item_out.get("content", []) if item.get("type") == "output_text"), "")
        plan = json.loads(output)
        return {
            "rules": [rule for rule in plan["rules"] if rule in DIAGNOSTIC_RULE_LIBRARY],
            "strategies": [strategy for strategy in plan["strategies"] if strategy in safe_strategies],
        }
    except (urllib.error.URLError, urllib.error.HTTPError, KeyError, ValueError, json.JSONDecodeError):
        return {"rules": [], "strategies": []}

def source_balances(text: str) -> tuple[Decimal | None, Decimal | None]:
    def find(kind):
        # A column heading such as "Closing Balance*" is commonly followed by
        # the first transaction date (for example 01/04/2024).  It is not the
        # statement endpoint. Prefer a labelled monetary value (with decimals)
        # and, when a summary repeats it, take the final source occurrence.
        monetary = re.findall(
            rf"\b{kind}\s*balance\b[^\d\r\n]{{0,20}}(?:\r?\n\s*)?"
            r"(-?[\d,]+\.\d{1,2}(?:\s*(?:CR|DR))?)",
            text,
            re.I,
        )
        if monetary:
            return money(monetary[-1])
        # Retain support for sources that print whole-number balances, while
        # rejecting a date fragment immediately after a table heading.
        match = re.search(rf"\b{kind}\s*balance\b\D{{0,20}}([\d,]+(?:\.\d{{1,2}})?(?:\s*(?:CR|DR))?)", text, re.I)
        if not match:
            return None
        value = money(match.group(1))
        following = text[match.end():match.end() + 12]
        if value is not None and abs(value) <= 31 and re.match(r"\s*/\s*\d{1,2}\s*/", following):
            return None
        return value
    opening, closing = find("opening"), find("closing")
    summary = re.search(r"(?is)statement\s+summary\s*:-?.{0,300}?opening\s+balance.*?\n\s*([\d,]+(?:\.\d{1,2})?)\s+\d+\s+\d+\s+[\d,]+(?:\.\d{1,2})?\s+[\d,]+(?:\.\d{1,2})?\s+([\d,]+(?:\.\d{1,2})?)", text)
    if summary:
        summary_open, summary_close = money(summary.group(1)), money(summary.group(2))
        if summary_open is not None: opening = summary_open
        if summary_close is not None: closing = summary_close
    # Jammu and Kashmir Bank-style statements use B/F rather than the words
    # "opening balance". It is a statement-level value, never a transaction.
    bf = re.search(r"\bB/F\b[\s\S]{0,80}?(-?[\d,]+(?:\.\d+)?)\s*(Dr|Cr)\b", text, re.I)
    if opening is None and bf:
        amount = money(bf.group(1))
        opening = -abs(amount) if amount is not None and bf.group(2).lower() == "dr" else amount
    return opening, closing

def source_transaction_totals(text: str) -> tuple[Decimal | None, Decimal | None]:
    """Read printed debit/credit totals when a statement provides them.

    These are independent controls. Statements without printed totals return
    ``None`` values and retain the ordinary opening-to-closing reconciliation.
    """
    summary = re.search(
        r"(?is)statement\s+summary\s*:-?.{0,300}?\bdebits?\b\s+\bcredits?\b.*?\n\s*"
        r"[\d,]+(?:\.\d{1,2})?\s+\d+\s+\d+\s+([\d,]+(?:\.\d{1,2})?)\s+([\d,]+(?:\.\d{1,2})?)",
        text,
    )
    if not summary:
        # Some banks print cumulative debit and credit totals as a final
        # Grand Total rather than a labelled statement-summary grid.
        grand = re.search(r"(?is)\bgrand\s+total\s*:\s*([\d,]+(?:\.\d{1,2})?)\s+([\d,]+(?:\.\d{1,2})?)", text)
        return (money(grand.group(1)), money(grand.group(2))) if grand else (None, None)
    return money(summary.group(1)), money(summary.group(2))

def table_balances(rows: list[list[object]], header_at: int, columns: dict[str, int]) -> tuple[Decimal | None, Decimal | None]:
    """Extract declared balances from the source table, independent of computed totals."""
    balance_index = columns.get("balance")
    if balance_index is None: return None, None
    def cell(row, key):
        i = columns.get(key); return row[i] if i is not None and i < len(row) else ""
    first, last = None, None
    for row in rows[header_at + 1:]:
        bal = money(cell(row, "balance"))
        if bal is None: continue
        narration = str(cell(row, "narration")).upper()
        if "OPENING BALANCE" in narration: first = bal
        elif str(cell(row, "date")).strip():
            last = bal
    return first, last

def transaction_date_value(value: str) -> datetime | None:
    for pattern in ("%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y", "%d-%b-%Y", "%d/%b/%Y", "%Y-%m-%d"):
        try: return datetime.strptime(value.strip(), pattern)
        except ValueError: continue
    return None

def raw_transaction_record_count(raw: str) -> int | None:
    """Independent source-record count for common text-layout statements."""
    header = raw[:2500].lower()
    # Full dates followed by narration identify the primary record line. This
    # excludes secondary value-date/time lines in dual-date layouts.
    long_form = re.findall(r"(?m)^\s*\d{2}-[A-Za-z]{3}-\d{4}\s+(?=[A-Za-z])", raw)
    if len(long_form) >= 3:
        return len(long_form)
    # Blank narration is allowed; date plus the row's amount/balance columns
    # is still a transaction record, so do not require a letter after date.
    # A few text layers detach a transaction's reference/narration onto the
    # preceding line and put its date, time, amount and balance below it. That
    # date is part of the same record, not another transaction.
    numeric_long_form = re.findall(
        r"(?m)^\s*\d{2}-\d{2}-\d{4}\b"
        # Repeated statement-period headings can begin a text line with
        # `01-04-2025 TO 31-03-2026`; that is metadata, not a transaction.
        r"(?!\s+TO\s+\d{2}-\d{2}-\d{4}\b)"
        r"(?!\s+\d{2}:\d{2}:\d{1,2}\s+[-\d,]+(?:\.\d+)?\s+[-\d,]+(?:\.\d+)?\s*(?:Dr|Cr)\b)", raw)
    if len(numeric_long_form) >= 3:
        return len(numeric_long_form)
    # Numeric-date statements use the same rule when a visible header has a
    # distinct Value Date column: only date lines followed by narration count.
    if "value dt" in header or "value date" in header:
        numeric_primary = re.findall(r"(?m)^\s*\d{2}[-/]\d{2}[-/]\d{2,4}\s+(?=[A-Za-z])", raw)
        if len(numeric_primary) >= 3:
            return len(numeric_primary)
    return None

def display_date(value: object) -> str:
    parsed = transaction_date_value(str(value))
    return parsed.strftime("%d/%m/%Y") if parsed else str(value or "")

def repair_truncated_table_date(value: object, raw: str) -> str:
    """Restore only a source-proven missing final year digit from a PDF cell."""
    text = str(value or "").strip()
    if transaction_date_value(text):
        return text
    partial = re.fullmatch(r"(\d{2}[-/]\d{2}[-/]\d{3})", text)
    if not partial:
        return text
    # PDF table extraction can put the final year digit in a neighbouring text
    # fragment. Require the exact partial date plus that digit in the original
    # source, so this never guesses a date from an amount or reference number.
    proven = re.search(re.escape(partial.group(1)) + r"\s*([0-9])\b", raw)
    return partial.group(1) + proven.group(1) if proven else text

def count_source_transactions(rows: list[list[object]], header_at: int, columns: dict[str, int]) -> int:
    """Count real source records without trusting parsed totals or narration text."""
    count = 0
    for row in rows[header_at + 1:]:
        def cell(key):
            i = columns.get(key)
            return row[i] if i is not None and i < len(row) else ""
        date = str(cell("date") or "").strip()
        if not transaction_date_value(date):
            continue
        narration = str(cell("narration") or "")
        if re.search(r"\b(?:B/F|OPENING\s+BALANCE)\b", narration, re.I):
            continue
        has_values = any(money(cell(key)) is not None for key in ("withdrawal", "deposit", "amount", "balance"))
        if has_values:
            count += 1
    return count

def structured_source_count(path: Path) -> int | None:
    """Independent table evidence used to reject a short fallback extraction."""
    if path.suffix.lower() != ".pdf": return None
    try:
        rows, _ = load_rows(path, None)
        header_at = next((i for i, row in enumerate(rows[:20]) if len(map_headers(row)) >= 3), None)
        if header_at is None: return None
        if generated_canonical_headers(rows[header_at]):
            # This is the generic text fallback's synthetic header, not an
            # independently detected PDF table.
            return None
        count = count_source_transactions(rows, header_at, map_headers(rows[header_at]))
        return count if count >= 3 else None
    except Exception:
        return None

def read_pdf_text(path: Path) -> str:
    reader = open_pdf_reader(path)
    return "\n".join(page.extract_text() or "" for page in reader.pages)

def cached_pdf_text(path: Path) -> str:
    key = str(path.resolve())
    with EXTRACTION_CACHE_LOCK:
        cached = PDF_TEXT_CACHE.get(key)
    if cached is not None:
        return cached
    text = read_pdf_text(path)
    with EXTRACTION_CACHE_LOCK:
        if len(PDF_TEXT_CACHE) >= 12:
            PDF_TEXT_CACHE.pop(next(iter(PDF_TEXT_CACHE)))
        PDF_TEXT_CACHE[key] = text
    return text

def remove_page_furniture(raw: str) -> str:
    """Remove repeated PDF page furniture while preserving transaction continuations."""
    # Footer and disclaimer between statement pages.
    raw = re.sub(r"(?is)page total:.*?(?=jammu and kashmir bank ltd|$)", "", raw)
    raw = re.sub(r"(?is)unless the constituent notifies.*?(?:date stamp\s+manager|manager)\s*", "", raw)
    # Repeated JKB-style page heading and account-holder block ends at currency.
    raw = re.sub(r"(?is)jammu and kashmir bank ltd.*?(?:currency code\s*:\s*inr|no nomination available for the account)\s*", "", raw)
    # A repeated statement title sometimes appears immediately before the
    # continuation fragment; remove title only, leave following text intact.
    raw = re.sub(r"(?i)statement of account for the period(?:\s+of)?\s*\d{2}[-/]\d{2}[-/]\d{4}\s+to\s+\d{2}[-/]\d{2}[-/]\d{4}", "", raw)
    # HDFC-style statements append a full account-information block after each
    # page number. It is between transaction rows, never narration. Preserve
    # the next dated row as the boundary and retain any pre-footer continuation.
    raw = re.sub(r"(?is)page\s*no\s*\.?\s*:\s*\d+.*?(?=\n\d{2}/\d{2}/\d{2}\s|\Z)", "", raw)
    return raw

def repair_detached_dated_continuations(raw: str) -> str:
    """Join reference text with the dated/time continuation that follows it.

    Coordinate-poor PDF text sometimes emits `TRRR/reference/` on one line
    and the rest of that transaction (date, time, amount, running balance) on
    the next. Moving the date before the fragment makes it one normal record
    and prevents a fabricated blank-narration row.
    """
    # Some banks print the posting date beside the reference and repeat the
    # value date immediately before the time/amount/balance cells. Preserve
    # one record boundary by removing that repeated inner date.
    repeated_date_time = re.compile(
        r"(?m)^(?P<head>[ \t]*\d{2}-\d{2}-\d{4}[^\r\n]*[A-Za-z][^\r\n]*)\r?\n"
        r"[ \t]*\d{2}-\d{2}-\d{4}\s+(?P<tail>\d{2}:\d{2}:\d{1,2}\s+[-\d,]+(?:\.\d+)?\s+[-\d,]+(?:\.\d+)?\s*(?:Dr|Cr)\b)"
    )
    raw = repeated_date_time.sub(lambda match: f"{match.group('head')} {match.group('tail')}", raw)
    pattern = re.compile(
        r"(?m)^[ \t]*(?P<narr>(?=[A-Za-z])[^\r\n]*[A-Za-z][^\r\n]*)[ \t]*\r?\n"
        r"[ \t]*(?P<date>\d{2}-\d{2}-\d{4})\s+"
        r"(?P<tail>\d{2}:\d{2}:\d{1,2}\s+[-\d,]+(?:\.\d+)?\s+[-\d,]+(?:\.\d+)?\s*(?:Dr|Cr)\b)"
    )
    return pattern.sub(lambda match: f"{match.group('date')} {match.group('narr')} {match.group('tail')}", raw)

def extract_pdf_table_batch(path_text: str, page_numbers: list[int], use_learned_geometry: bool = False) -> list[list[list[object]]]:
    """Worker-safe batch extraction used only for very large PDFs."""
    with open_pdfplumber(Path(path_text)) as pdf:
        if use_learned_geometry:
            return [[pdf.pages[number].extract_table()] for number in page_numbers]
        return [pdf.pages[number].extract_tables() for number in page_numbers]

def sampled_geometry_header(path: Path, page_count: int) -> list[object] | None:
    """Learn a transaction-table header from representative original pages."""
    with open_pdfplumber(path) as pdf:
        for number in sampled_page_indices(page_count):
            for table in pdf.pages[number].extract_tables():
                for row in table:
                    if row and len(map_headers(row)) >= 3:
                        return row
    return None

def sampled_geometry_profile(path: Path) -> tuple[list[object], list[tuple[float, float]]] | None:
    """Learn column bands from representative original-PDF pages.

    Prefer ruled-table cells, but also support borderless bank statements by
    deriving bands from the x-position of their printed header words.
    """
    with open_pdfplumber(path) as pdf:
        page_numbers = list(range(len(pdf.pages))) if len(pdf.pages) <= 60 else sampled_page_indices(len(pdf.pages))
        for number in page_numbers:
            page = pdf.pages[number]
            for table in page.find_tables():
                extracted = table.extract()
                for index, row in enumerate(extracted):
                    if row and len(map_headers(row)) >= 3 and index < len(table.rows):
                        cells = table.rows[index].cells
                        if cells and all(cell for cell in cells):
                            return row, [(float(cell[0]), float(cell[2])) for cell in cells]
            lines: dict[int, list[dict]] = {}
            for word in page.extract_words(x_tolerance=1, y_tolerance=2):
                lines.setdefault(round(float(word["top"]) / 3), []).append(word)
            for words in lines.values():
                ordered = sorted(words, key=lambda item: float(item["x0"]))
                labels = [str(word["text"]) for word in ordered]
                if len(map_headers(labels)) < 3:
                    continue
                starts = [float(word["x0"]) for word in ordered]
                # Header words are the authoritative column starts. A band
                # ends immediately before the next heading; the last reaches
                # the page edge. This is the same coordinate method used for
                # a ruled table, without requiring border lines.
                bands = [(0.0 if index == 0 else starts[index] - 3.0,
                          (starts[index + 1] - 3.0) if index + 1 < len(starts) else float(page.width))
                         for index in range(len(starts))]
                return labels, bands
    return None

def extract_geometry_profile_rows(path: Path, page_numbers: set[int] | None = None) -> list[list[object]]:
    """Apply sampled column bands without table rediscovery.

    `page_numbers` is a zero-based representative-page set used for fast
    structural screening. A final candidate always calls this for every page.
    """
    profile = sampled_geometry_profile(path)
    if not profile:
        return []
    header, bands = profile
    column_map = map_headers(header)
    date_index = column_map.get("date", 0)
    rows: list[list[object]] = [header]
    current: list[str] | None = None
    with open_pdfplumber(path) as pdf:
        for page_number, page in enumerate(pdf.pages):
            if page_numbers is not None and page_number not in page_numbers:
                continue
            lines: dict[int, list[dict]] = {}
            footer_started = False
            for word in page.extract_words(x_tolerance=1, y_tolerance=2):
                lines.setdefault(round(float(word["top"]) / 3), []).append(word)
            for words in lines.values():
                if footer_started:
                    continue
                line_text = " ".join(str(word["text"]) for word in words)
                footer_on_line = bool(re.search(r"(?i)\b(?:page\s+total|grand\s+total|date/time|system\s+generated|page\s+\d+\s+of)\b", line_text))
                cells = ["" for _ in bands]
                for word in sorted(words, key=lambda item: float(item["x0"])):
                    center = (float(word["x0"]) + float(word["x1"])) / 2
                    column = next((i for i, (left, right) in enumerate(bands) if left <= center <= right), None)
                    if column is not None:
                        cells[column] = (cells[column] + " " + word["text"]).strip()
                # A page footer can share the final transaction's y-band.
                # Keep the transaction portion, never append page/grand
                # totals or the generated-statement disclaimer to its cells.
                cells = [re.split(r"(?i)\b(?:page\s+total|grand\s+total|date/time|this\s+is\s+a\s+system\s+generated)", cell)[0].strip() for cell in cells]
                if footer_on_line:
                    withdrawal_index = column_map.get("withdrawal")
                    deposit_index = column_map.get("deposit")
                    balance_index = column_map.get("balance")
                    numeric = lambda value: (re.search(r"-?[\d,]+(?:\.\d{1,2})?", value or "").group() if re.search(r"-?[\d,]+(?:\.\d{1,2})?", value or "") else "")
                    # The final source row can share its y-band with Page
                    # Total. Preserve its first amount and balance only.
                    if withdrawal_index is not None and numeric(cells[withdrawal_index]):
                        cells[withdrawal_index] = numeric(cells[withdrawal_index])
                        if deposit_index is not None: cells[deposit_index] = ""
                    elif deposit_index is not None and numeric(cells[deposit_index]):
                        cells[deposit_index] = numeric(cells[deposit_index])
                        if withdrawal_index is not None: cells[withdrawal_index] = ""
                    if balance_index is not None:
                        cells[balance_index] = numeric(cells[balance_index])
                    footer_started = True
                    # A standalone footer must never become a continuation of
                    # the final transaction. If a transaction and footer share
                    # one band, its valid date keeps the transaction portion.
                    if not transaction_date_value(cells[date_index]):
                        continue
                if not any(cells) or len(map_headers(cells)) >= 3:
                    continue
                if transaction_date_value(cells[date_index]):
                    if current is not None:
                        rows.append(current)
                    current = cells
                elif current is not None:
                    for i, value in enumerate(cells):
                        if value:
                            current[i] = (current[i] + " " + value).strip()
            # pdfplumber caches page character/word objects. Explicitly clear
            # each page after converting it to compact row strings so a 250+
            # page PDF does not retain all geometry in the 1 GB web container.
            try:
                page.close()
            except Exception:
                pass
            # Close a fully evidenced final row before the next page's bank
            # header arrives. Keep only incomplete rows open so genuine
            # two-page narrations/amounts can still be merged on page N+1.
            if current is not None:
                row_date = transaction_date_value(current[date_index] if date_index < len(current) else "")
                withdrawal_index = column_map.get("withdrawal")
                deposit_index = column_map.get("deposit")
                balance_index = column_map.get("balance")
                has_amount = bool(
                    (withdrawal_index is not None and withdrawal_index < len(current) and money(current[withdrawal_index]) is not None)
                    or (deposit_index is not None and deposit_index < len(current) and money(current[deposit_index]) is not None)
                )
                has_balance = balance_index is not None and balance_index < len(current) and money(current[balance_index]) is not None
                if row_date and has_amount and has_balance:
                    rows.append(current)
                    current = None
    if current is not None:
        rows.append(current)
    return rows

def sample_candidate_plausible(path: Path, strategy: str | None) -> bool:
    """Reject structurally impossible large-PDF candidates before full parsing.

    This is deliberately not a release validation. It only checks original-PDF
    sample pages for a real header and multiple dated record shapes, letting
    full extraction and all financial/narration gates remain the sole
    certification authority.
    """
    if path.suffix.lower() != ".pdf" or not is_large_pdf(path):
        return True
    count = len(open_pdf_reader(path).pages)
    samples = set(sampled_page_indices(count))
    try:
        if strategy == "geometry_profile":
            rows = extract_geometry_profile_rows(path, samples)
        else:
            raw = sampled_pdf_text(path)
            if strategy == "value_date_unsigned":
                rows = extract_text_layout_rows(raw, unsigned_balance=True, use_value_date=True)
            elif strategy == "unsigned_running_balance_text":
                rows = extract_text_layout_rows(raw, unsigned_balance=True)
            elif strategy in {"running_balance_text", "page_text_unsigned"}:
                rows = extract_text_layout_rows(raw, unsigned_balance=(strategy == "page_text_unsigned"))
            else:
                # Table rediscovery is intentionally not run on every page of
                # a large PDF. Geometry/AI candidates are stronger evidence.
                return False
        if len(rows) < 4 or len(map_headers(rows[0])) < 3:
            return False
        date_index = map_headers(rows[0]).get("date", 0)
        return sum(bool(transaction_date_value(row[date_index] if date_index < len(row) else "")) for row in rows[1:]) >= 2
    except Exception:
        return False

def extract_pdf_rows(path: Path, strategy_override: str | None = None) -> tuple[list[list[object]], str]:
    raw = remove_page_furniture(cached_pdf_text(path))
    dual_date_time_layout = bool(re.search(r"\d{2}-[A-Za-z]{3}-\d{4}[\s\S]{0,180}\d{2}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}", raw))
    merged: list[list[object]] = []
    known_header: list[object] | None = None
    if strategy_override == "geometry_profile":
        return extract_geometry_profile_rows(path), raw
    if strategy_override == "page_text_unsigned":
        header = ["Date", "Narration", "Withdrawal", "Deposit", "Instrument Number", "Balance"]
        reader = open_pdf_reader(path)
        merged = [header]
        for page in reader.pages:
            page_rows = extract_text_layout_rows(remove_page_furniture(page.extract_text() or ""), unsigned_balance=True)
            if page_rows:
                merged.extend(page_rows[1:])
        return merged if len(merged) > 1 else [], raw
    if strategy_override not in ("running_balance_text", "unsigned_running_balance_text", "value_date_unsigned"):
        with open_pdfplumber(path) as pdf:
            page_count = len(pdf.pages)
            if page_count > 60:
                # Learn the geometry/header from the original-PDF sample before
                # scanning all pages. With a known header, each remaining page
                # uses the lighter single-table path rather than table discovery.
                known_header = sampled_geometry_header(path, page_count)
                if known_header:
                    merged.append(known_header)
                workers = min(4, max(2, os.cpu_count() or 2))
                batches = [list(range(start, min(page_count, start + 32))) for start in range(0, page_count, 32)]
                try:
                    # Threads avoid Windows process-spawn behaviour, which can
                    # otherwise create duplicate web-server children when this
                    # local app is launched through its bootstrap script.
                    with ThreadPoolExecutor(max_workers=workers) as executor:
                        page_tables = [tables for batch in executor.map(extract_pdf_table_batch, [str(path)] * len(batches), batches, [bool(known_header)] * len(batches)) for tables in batch]
                except Exception:
                    page_tables = [page.extract_tables() for page in pdf.pages]
            else:
                page_tables = [page.extract_tables() for page in pdf.pages]
        for tables in page_tables:
            for table in tables:
                if not table:
                    continue
                for row in table:
                    if not row: continue
                    mapped = map_headers(row)
                    if len(mapped) >= 3:
                        known_header = row
                        if not merged: merged.append(row)
                        continue
                    if known_header and len(row) == len(known_header): merged.append(row)
    if not merged:
        strategy = strategy_override or ai_choose_text_strategy(raw)
        if strategy == "running_balance_text":
            merged = extract_text_layout_rows(raw)
        elif strategy == "unsigned_running_balance_text":
            merged = extract_text_layout_rows(raw, unsigned_balance=True, use_value_date=dual_date_time_layout, dual_date_time=dual_date_time_layout)
        elif strategy == "value_date_unsigned":
            merged = extract_text_layout_rows(raw, unsigned_balance=True, use_value_date=True)
        else:
            # AI classification is only one candidate. Try the safe
            # running-balance strategy too; the two validation gates decide
            # whether it can be released, not the model's first guess.
            merged = extract_text_layout_rows(raw)
        if not merged:
            raise ValueError("No readable transaction rows were found. This statement may be scanned, password protected, or incomplete.")
    if not merged:
        raise ValueError("No readable transaction rows were found. This statement may be scanned, password protected, or incomplete.")
    return merged, raw

def extract_text_layout_rows(raw: str, unsigned_balance: bool = False, use_value_date: bool = False, dual_date_time: bool = False) -> list[list[object]]:
    """Generic fallback for text PDFs without table borders.

    Many bank PDFs put one dated transaction after another and print a running
    balance with Dr/Cr. Consecutive balance changes provide an independent way
    to assign withdrawal/deposit amounts without relying on fixed x positions.
    """
    raw = repair_detached_dated_continuations(raw)
    header = ["Date", "Narration", "Withdrawal", "Deposit", "Instrument Number", "Balance"]
    # Bound numeric dates so a long transaction/reference ID cannot be
    # mistaken for a partial date (for example, `...382/22-04...`).
    date_pattern = r"(?<!\d)(?:\d{2}[-/]\d{2}[-/]\d{2,4}|\d{2}-[A-Za-z]{3}-\d{4})(?!\d)"
    statement_opening, _ = source_balances(raw)
    # Some text PDFs begin with a statement period such as "Between
    # 01-04-2025 and ...". It resembles a transaction date but is not one.
    # Start at the printed transaction-table heading when it is available,
    # while retaining the separately extracted opening balance above.
    table_heading = re.search(r"(?im)^\s*(?:transaction\s+)?date\s+.*?(?:particulars?|narration|description).*?(?:closing\s+balance|balance)\s*$", raw)
    transaction_text = raw[table_heading.end():] if table_heading else raw
    if use_value_date:
        # Coordinate-poor text extraction often puts the Value Date and the
        # amount/balance cells on the next physical line. Reattach only lines
        # that are a date followed solely by numeric cells; a posting-date line
        # has narration/reference text and remains a record boundary.
        value_line = rf"(?m)^\s*({date_pattern})\s+((?:-?[\d,]+(?:\.\d{{1,2}})?\s*){{2,}})$"
        transaction_text = re.sub(value_line, lambda match: " " + match.group(1) + " " + match.group(2).strip(), transaction_text)
    # A Value Date may appear later on the same visual line. Only a date at
    # the beginning of a physical line can start the next transaction block.
    dual_date_layout = bool(re.search(r"(?i)\bvalue\s*(?:date|dt)\b", raw[:2000])) or dual_date_time
    # In compact text layers, a whole PDF page may be emitted as one line.
    # Split every date for ordinary single-date layouts, but retain the
    # beginning-of-line rule for dual-date layouts so Value Date is not turned
    # into a second transaction.
    primary_date = r"\d{2}-[A-Za-z]{3}-\d{4}" if dual_date_time else date_pattern
    split_pattern = rf"(?m)(?=^\s*{primary_date}\s)" if dual_date_layout else rf"(?={date_pattern}\s)"
    chunks = re.split(split_pattern, transaction_text)
    rows, previous_balance = [header], statement_opening
    for chunk in chunks:
        date = re.match(rf"\s*({date_pattern})\s+", chunk)
        if not date: continue
        if re.search(r"\bB/F\b", chunk, re.I): continue
        # Repeated page headers contain the statement-period end date and
        # account-holder text, but are not transaction rows.
        if "PARTICULARS" in chunk[:800].upper(): continue
        # Only the portion before page/report totals belongs to the current
        # transaction. Footer balances are not closing transaction balances.
        chunk = re.split(r"(?i)\b(?:page total|grand total|funds in clearing|total available amount|effective available amount|closing balance|unless the constituent)\b", chunk)[0]
        balance_matches = list(re.finditer(r"(-?[\d,]+(?:\.\d+)?)\s*(Dr|Cr)\b", chunk, re.I))
        numeric_matches = list(re.finditer(r"-?[\d,]+(?:\.\d+)?", chunk))
        forced_amount, value_date_match = None, None
        matching_secondary_date = False
        if use_value_date:
            dates = list(re.finditer(date_pattern, chunk))
            if len(dates) >= 2:
                first_date, second_date = transaction_date_value(dates[0].group()), transaction_date_value(dates[1].group())
                matching_secondary_date = first_date is not None and second_date is not None and first_date.date() == second_date.date()
            if matching_secondary_date:
                value_date_match = dates[1]
            # The first two numeric cells after Value Date are respectively
            # the transaction amount and the running balance. Continuation
            # text below the row may contain long numeric reference fragments
            # and must not replace the balance.
                value_tail = re.sub(r"\b\d{2}:\d{2}:\d{2}\b", "", chunk[value_date_match.end():])
                value_numbers = list(re.finditer(r"-?[\d,]+(?:\.\d+)?", value_tail))
                if len(value_numbers) < 2: continue
                forced_amount = money(value_numbers[0].group())
                balance_match = value_numbers[1]
                balance_start = value_date_match.end() + balance_match.start()
                numeric_balance = money(balance_match.group())
        if not matching_secondary_date:
            # Most signed layouts mark balance with Dr/Cr. A zero balance may
            # omit that suffix, however, and remains a valid dated row when it
            # has a preceding transaction amount.
            zero_balance_without_suffix = (
                not balance_matches and len(numeric_matches) >= 2
                and money(numeric_matches[-1].group()) == Decimal("0")
            )
            if not balance_matches and not unsigned_balance and not zero_balance_without_suffix: continue
            if not balance_matches and len(numeric_matches) < 2: continue
            balance_match = balance_matches[-1] if balance_matches else numeric_matches[-1]
            balance_start = balance_match.start()
            numeric_balance = money(balance_match.group(1) if balance_matches else balance_match.group())
        if numeric_balance is None: continue
        if balance_matches and not matching_secondary_date:
            balance = -abs(numeric_balance) if balance_match.group(2).lower() == "dr" else abs(numeric_balance)
        else:
            # An unsigned layout can still print a literal negative balance.
            # Preserve that sign; it determines the debit/credit movement.
            balance = numeric_balance
        numbers = list(re.finditer(r"-?[\d,]+(?:\.\d+)?", chunk[:balance_start]))
        amount = forced_amount if forced_amount is not None else (money(numbers[-1].group()) if numbers else None)
        if previous_balance is None:
            # First row: use its stated amount when available, otherwise keep
            # zero and derive the opening balance from its running balance.
            delta = Decimal("0")
            if amount is not None:
                narrative_before_amount = chunk[:numbers[-1].start()].lower()
                delta = -abs(amount) if "debit" in narrative_before_amount else abs(amount)
        else:
            delta = balance - previous_balance
        withdrawal = -delta if delta < 0 else Decimal("0")
        deposit = delta if delta > 0 else Decimal("0")
        narration_end = balance_start
        output_date = date.group(1)
        if matching_secondary_date:
            # In a dual-date layout the second date in the physical record is
            # the Value Date. It is a column value, not part of narration.
            if value_date_match is not None:
                output_date = date.group(1) if dual_date_time else value_date_match.group(0)
                narration_end = value_date_match.start()
        narration = clean_narration(chunk[date.end():narration_end])
        instrument_source = chunk[date.end():value_date_match.start()] if use_value_date and value_date_match is not None else narration
        instrument_matches = list(re.finditer(r"\b\d{6,}\b", instrument_source))
        instrument = instrument_matches[-1].group() if instrument_matches else ""
        # Retain the source amount as private, in-memory validation evidence.
        # It is not exported as a column, but prevents a balance-delta parser
        # from silently replacing a visibly printed transaction amount.
        rows.append([display_date(output_date), narration, withdrawal, deposit, instrument, balance, amount])
        previous_balance = balance
    return rows if len(rows) > 1 else []

def load_rows(path: Path, strategy_override: str | None = None) -> tuple[list[list[object]], str]:
    cache_key = (str(path.resolve()), strategy_override or "detected_table")
    with EXTRACTION_CACHE_LOCK:
        cached = EXTRACTION_CACHE.get(cache_key)
    if cached is not None:
        return cached
    ext = path.suffix.lower()
    if ext == ".csv":
        raw = path.read_text(encoding="utf-8-sig", errors="replace")
        result = list(csv.reader(io.StringIO(raw))), raw
    if ext in (".xlsx", ".xls"):
        book = openpyxl.load_workbook(path, data_only=True, read_only=True)
        sheet = book.active
        rows = [list(row) for row in sheet.iter_rows(values_only=True)]
        result = rows, "\n".join(" ".join(map(str, row)) for row in rows)
    if ext == ".txt":
        raw = path.read_text(encoding="utf-8", errors="replace")
        dialect = csv.excel_tab if "\t" in raw.splitlines()[0] else csv.excel
        result = list(csv.reader(io.StringIO(raw), dialect=dialect)), raw
    elif ext == ".pdf": result = extract_pdf_rows(path, strategy_override)
    else: raise ValueError("This file needs a document-text extraction profile before it can be parsed.")
    with EXTRACTION_CACHE_LOCK:
        # Keep only recent upload evidence; validated learning remains separate
        # and contains no statement data.
        if len(EXTRACTION_CACHE) >= 24:
            EXTRACTION_CACHE.pop(next(iter(EXTRACTION_CACHE)))
        EXTRACTION_CACHE[cache_key] = result
    return result

def reconstruct_unordered_balance_chain(transactions: list[dict], opening: Decimal, closing: Decimal) -> list[dict] | None:
    """Recover a source-proved chain when PDF row order is not ledger order.

    Each source row supplies its visible amount and balance. For a balance B
    and amount A, its predecessor can only be B-A (deposit) or B+A
    (withdrawal). We follow those links from the stated opening balance and
    accept the reconstruction only when every transaction is used exactly once
    and the final balance equals the stated closing balance.
    """
    if len(transactions) < 3:
        return None
    by_predecessor: dict[Decimal, list[tuple[int, str]]] = {}
    for index, transaction in enumerate(transactions):
        balance = transaction.get("balance")
        amount = abs(Decimal(transaction.get("withdrawal", 0)) - Decimal(transaction.get("deposit", 0)))
        if balance is None or amount <= 0:
            return None
        by_predecessor.setdefault(balance - amount, []).append((index, "deposit"))
        by_predecessor.setdefault(balance + amount, []).append((index, "withdrawal"))
    ordered: list[dict] = []
    used: set[int] = set()
    running = opening
    for _ in range(len(transactions)):
        options = [(index, direction) for index, direction in by_predecessor.get(running, []) if index not in used]
        # More than one possible next source row means the evidence cannot
        # prove one particular chain. Never guess merely to pass validation.
        if len(options) != 1:
            return None
        index, direction = options[0]
        transaction = dict(transactions[index])
        amount = abs(Decimal(transaction["withdrawal"]) - Decimal(transaction["deposit"]))
        transaction["withdrawal"] = amount if direction == "withdrawal" else Decimal("0")
        transaction["deposit"] = amount if direction == "deposit" else Decimal("0")
        running = transaction["balance"]
        ordered.append(transaction)
        used.add(index)
    if running.quantize(Decimal(".01")) != closing.quantize(Decimal(".01")):
        return None
    return ordered

def parse_statement(path: Path, fallback_open: str, fallback_close: str, strategy_override: str | None = None, force_ai_profile: bool = False, repair_context: str = ""):
    if path.suffix.lower() == ".pdf":
        source_text = remove_page_furniture(cached_pdf_text(path))
        if len(re.sub(r"\W", "", source_text)) < 80:
            raise ValueError("OCR_REQUIRED: This PDF is image-only and has no reliable machine-readable transaction text. OCR must recover the source before any parser can be validated.")
    rows, raw = load_rows(path, strategy_override)
    if not rows: raise ValueError("The statement contains no readable rows.")
    header_at = next((i for i, row in enumerate(rows[:20]) if len(map_headers(row)) >= 3), None)
    ai_columns = None
    if header_at is None or force_ai_profile:
        generated = ai_generated_profile(rows, raw, repair_context, path)
        if generated:
            header_at, ai_columns = generated
        elif header_at is None:
            raise ValueError("Could not identify transaction columns. The AI parser generator could not produce a safe layout profile.")
    headers = rows[header_at]
    # A saved profile takes precedence, but automatically fill any missing
    # fields from the current statement header. This keeps early/incomplete
    # profiles from blocking the universal generator as it improves.
    layout_fingerprint = text_layout_fingerprint(raw) if generated_canonical_headers(headers) else ""
    exact_profile = load_profile(headers, layout_fingerprint)
    parent_profile, inherited_columns = (None, {}) if exact_profile else find_related_profile(headers)
    # During controlled self-healing, the newly proposed AI addendum is allowed
    # to supersede a prior mapping only for this candidate. It is persisted only
    # after all validation gates pass.
    # Explicit headers on the uploaded source are stronger than a related
    # profile. An addendum may supply only a missing field; it must never move
    # a clearly labelled Withdrawal, Deposit or Balance column merely because
    # an older layout used different offsets.
    source_columns = map_headers(headers)
    inherited_missing = {key: value for key, value in inherited_columns.items() if key not in source_columns}
    exact_missing = {key: value for key, value in (exact_profile or {}).items() if key not in source_columns}
    if force_ai_profile:
        columns = {**source_columns, **inherited_missing, **exact_missing, **(ai_columns or {})}
    else:
        columns = {**source_columns, **inherited_missing, **(ai_columns or {}), **exact_missing}
    tx = []
    for row in rows[header_at + 1:]:
        def cell(key):
            i = columns.get(key); return row[i] if i is not None and i < len(row) else ""
        if not any(str(x or "").strip() for x in row): continue
        table_date = repair_truncated_table_date(cell("date"), raw)
        # Transaction totals and closing/opening labels often sit in the debit
        # and credit columns. A real transaction must carry a valid date.
        if not transaction_date_value(table_date):
            continue
        withdrawal, deposit = money(cell("withdrawal")), money(cell("deposit"))
        if withdrawal is None and deposit is None and "amount" in columns and "transaction_type" in columns:
            amount, kind = money(cell("amount")), str(cell("transaction_type")).upper()
            if amount is not None:
                withdrawal = amount if "DR" in kind else Decimal("0")
                deposit = amount if "CR" in kind else Decimal("0")
        if withdrawal is None: withdrawal = Decimal("0")
        if deposit is None: deposit = Decimal("0")
        if withdrawal or deposit:
            tx.append({"date": display_date(table_date), "narration": str(cell("narration") or ""), "withdrawal": withdrawal, "deposit": deposit, "instrument_number": str(cell("instrument_number") or ""), "balance": money(cell("balance")), "source_amount": money(row[6]) if len(row) > 6 else None})
    source_opening, source_closing = source_balances(raw)
    opening, closing = source_opening, source_closing
    tab_opening, tab_closing = table_balances(rows, header_at, columns)
    if path.suffix.lower() == ".pdf":
        # PDF headers such as "Closing Balance" can be followed by a date and
        # confuse free-text matching. The extracted table is authoritative here.
        # However, a printed statement-level opening/closing label is stronger
        # evidence than an inferred transaction row and must always win.
        opening = source_opening if source_opening is not None else tab_opening
        closing = source_closing if source_closing is not None else tab_closing
    else:
        opening = opening if opening is not None else tab_opening
        closing = closing if closing is not None else tab_closing
    if opening is None and tx:
        first = tx[0]
        if first["balance"] is not None:
            # The first source row's running balance establishes the statement
            # opening balance when the PDF does not print a separate B/F row.
            opening = first["balance"] - first["deposit"] + first["withdrawal"]
    # Retry candidate: many statements print newest transactions first. Reverse
    # to chronological order and re-derive endpoints from running balances.
    first_date = transaction_date_value(tx[0]["date"]) if tx else None
    last_date = transaction_date_value(tx[-1]["date"]) if tx else None
    if first_date and last_date and first_date > last_date:
        tx.reverse()
        if source_opening is None and tab_opening is None and tx[0]["balance"] is not None:
            opening = tx[0]["balance"] - tx[0]["deposit"] + tx[0]["withdrawal"]
        if source_closing is None and tx[-1]["balance"] is not None:
            closing = tx[-1]["balance"]
    opening = opening if opening is not None else money(fallback_open)
    closing = closing if closing is not None else money(fallback_close)
    if opening is None or closing is None: raise ValueError("Opening and closing balances could not be found. Supply them only as a fallback after confirming them from the source statement.")
    # Certain bank exports visually group same-date rows rather than preserving
    # ledger order. Try a source-amount-preserving reconstruction before any
    # balance-delta fallback; it succeeds only for one complete, unique chain.
    if strategy_override == "geometry_profile":
        reconstructed = reconstruct_unordered_balance_chain(tx, opening, closing)
        # In an unordered statement, the first displayed row is not reliable
        # opening evidence. A printed Grand Total can instead derive opening,
        # but only when its own source amounts create one complete, unique
        # balance chain all the way to the declared closing balance.
        if reconstructed is None and source_opening is None:
            declared_opening_withdrawals, declared_opening_deposits = source_transaction_totals(
                cached_pdf_text(path) if path.suffix.lower() == ".pdf" else raw
            )
            if declared_opening_withdrawals is not None and declared_opening_deposits is not None:
                total_derived_opening = closing + declared_opening_withdrawals - declared_opening_deposits
                reconstructed = reconstruct_unordered_balance_chain(tx, total_derived_opening, closing)
                if reconstructed is not None:
                    opening = total_derived_opening
        if reconstructed is not None:
            tx = reconstructed
    if strategy_override in ("running_balance_text", "unsigned_running_balance_text", "value_date_unsigned", "page_text_unsigned"):
        # A page-level extraction may start a new page without the preceding
        # running balance. Recompute debit/credit from the joined balances so
        # the candidate is independent of page boundaries.
        previous = opening
        for transaction in tx:
            if transaction["balance"] is None:
                continue
            delta = transaction["balance"] - previous
            transaction["withdrawal"] = -delta if delta < 0 else Decimal("0")
            transaction["deposit"] = delta if delta > 0 else Decimal("0")
            previous = transaction["balance"]
    total_w = sum((x["withdrawal"] for x in tx), Decimal("0")); total_d = sum((x["deposit"] for x in tx), Decimal("0"))
    computed = opening - total_w + total_d
    total_reconciles = computed.quantize(Decimal(".01"), rounding=ROUND_HALF_UP) == closing.quantize(Decimal(".01"), rounding=ROUND_HALF_UP)
    # Footer cleaning removes page totals and may remove the final Grand Total;
    # inspect the original statement text for this optional source control.
    declared_withdrawals, declared_deposits = source_transaction_totals(cached_pdf_text(path) if path.suffix.lower() == ".pdf" else raw)
    statement_totals_valid = (
        (declared_withdrawals is None or total_w.quantize(Decimal(".01")) == declared_withdrawals.quantize(Decimal(".01")))
        and (declared_deposits is None or total_d.quantize(Decimal(".01")) == declared_deposits.quantize(Decimal(".01")))
    )
    def materially_divergent(parsed: Decimal, declared: Decimal | None) -> bool:
        if declared is None or declared == 0:
            return False
        # A large percentage gap cannot be a harmless statement-summary typo;
        # it is strong evidence that the parser read reference IDs as amounts.
        return abs(parsed - declared) / abs(declared) > Decimal("0.05")
    statement_totals_plausible = not (materially_divergent(total_w, declared_withdrawals) or materially_divergent(total_d, declared_deposits))
    # Validate each movement against its source amount independently of the
    # running-balance chain.  A bad balance column must not hide an amount
    # mapping error simply because the chain stops at its first bad row.
    source_amount_valid = True
    for transaction in tx:
        source_amount = transaction.get("source_amount")
        movement = abs(transaction["deposit"] - transaction["withdrawal"])
        if source_amount is not None and movement.quantize(Decimal(".01")) != abs(source_amount).quantize(Decimal(".01")):
            source_amount_valid = False
            break
    running = opening
    # Validate each step independently, not just the final reconciliation.
    # This is the user's balance-chain equation rearranged forward.
    running_balance_valid = bool(tx)
    chain_checked = 0
    chain_breaks = 0
    for transaction in tx:
        balance = transaction["balance"]
        if balance is None:
            running_balance_valid = False
            chain_breaks += 1
            continue
        chain_checked += 1
        expected_current = running - transaction["withdrawal"] + transaction["deposit"]
        if expected_current.quantize(Decimal(".01")) != balance.quantize(Decimal(".01")):
            running_balance_valid = False
            chain_breaks += 1
        running = balance
    no_opening_as_transaction = not any(normalize_narration(x["narration"]) in ("bf", "openingbalance") for x in tx)
    # A fallback may find only a small group of rows that happens to reconcile.
    # Compare it with independently detected table rows so an incomplete subset
    # can never become a validated parser or profile.
    expected_source_count = count_source_transactions(rows, header_at, columns)
    # Long-form dates at the start of source lines provide independent record
    # evidence for compact text PDFs. Do not let a malformed candidate define
    # its own coverage denominator.
    long_date_records = len(re.findall(r"(?m)^\s*\d{2}-[A-Za-z]{3}-\d{4}\b", raw))
    if long_date_records:
        expected_source_count = max(expected_source_count, long_date_records)
    independent_count = raw_transaction_record_count(raw)
    if independent_count is not None:
        expected_source_count = max(expected_source_count, independent_count)
    if strategy_override is not None:
        table_count = structured_source_count(path)
        if table_count is not None:
            expected_source_count = table_count
    coverage_valid = len(tx) == expected_source_count
    # A multi-page statement cannot be safely accepted from a single inferred
    # transaction: that comparison would merely validate an incomplete source
    # count produced by the same failed extraction path.
    if path.suffix.lower() == ".pdf":
        try:
            if len(open_pdf_reader(path).pages) > 1 and len(tx) < 2:
                coverage_valid = False
        except Exception:
            pass
    # Printed debit/credit totals are useful independent evidence, but some
    # banks issue statements with incorrect summary totals. They are reported
    # as a warning; release still depends on transaction-level running-balance
    # reconciliation, endpoints, coverage, and narration validation.
    # Printed debit/credit summary totals are useful cross-check evidence, but
    # some bank statements print them incorrectly. They must never override a
    # complete source-record count, transaction-level balance chain, and the
    # statement opening-to-closing reconciliation.
    # Exception for objectively unreliable *source* balance columns.  This is
    # intentionally narrow: printed debit/credit totals must exactly match
    # parsed source amounts, all source records must be covered, and the normal
    # chain must be broken. One usable transaction balance anchors
    # assumed endpoints; the printed totals, coverage, and narration remain
    # the certification evidence, not the broken balance chain.
    source_balance_unreliable = False
    if (
        not running_balance_valid
        and declared_withdrawals is not None
        and declared_deposits is not None
        and statement_totals_valid
        and source_amount_valid
        and coverage_valid
        and len(tx) >= 20
    ):
        first_tx, last_tx = tx[0], tx[-1]
        inferred_opening = (first_tx["balance"] - first_tx["deposit"] + first_tx["withdrawal"] if first_tx.get("balance") is not None else None)
        inferred_closing = last_tx.get("balance")
        if inferred_opening is not None:
            opening = inferred_opening
            closing = opening - declared_withdrawals + declared_deposits
            endpoint_derived = "assumed_closing_from_first_transaction"
        elif inferred_closing is not None:
            closing = inferred_closing
            opening = closing + declared_withdrawals - declared_deposits
            endpoint_derived = "assumed_opening_from_last_transaction"
        else:
            endpoint_derived = "none"
        if endpoint_derived != "none":
            computed = opening - total_w + total_d
            total_reconciles = computed.quantize(Decimal(".01")) == closing.quantize(Decimal(".01"))
            source_balance_unreliable = total_reconciles
    financial_valid = total_reconciles and (running_balance_valid or source_balance_unreliable) and source_amount_valid and no_opening_as_transaction and coverage_valid
    # Preserve this certified exception in the profile without changing any
    # actual column mapping.  Consumers must never present it as a normal
    # balance-chain pass.
    columns["_source_balance_unreliable"] = source_balance_unreliable
    columns["_balance_endpoint_derived"] = locals().get("endpoint_derived", "none")
    # The original source text is independent evidence.  A narration that cannot
    # be located there is not silently accepted just because amounts reconcile.
    # Build the normalized source once. Re-normalizing a 250-page statement
    # for every transaction made narration validation quadratic and could turn
    # a valid parser attempt into a 20+ minute wait.
    normalized_source = normalize_narration(raw)
    unmatched = [x["narration"] for x in tx if normalize_narration(x["narration"]) and normalize_narration(x["narration"]) not in normalized_source]
    malformed_narrations = [x["narration"] for x in tx if re.fullmatch(r"\s*[\d,.]+\s*", x["narration"] or "")]
    narration_valid = not unmatched and not malformed_narrations and coverage_valid
    return tx, opening, closing, total_w, total_d, computed, financial_valid, narration_valid, unmatched, headers, columns, parent_profile, coverage_valid, expected_source_count, layout_fingerprint, declared_withdrawals, declared_deposits, statement_totals_valid

def export_excel(tx, opening, closing, total_w, total_d, computed, financial_valid, narration_valid, coverage_valid, expected_source_count, declared_withdrawals=None, declared_deposits=None, statement_totals_valid=True, source_balance_unreliable=False):
    out = EXPORTS / f"validated-statement-{uuid.uuid4().hex}.xlsx"; wb = openpyxl.Workbook(); ws = wb.active; ws.title = "Transactions"
    ws.append(["Date", "Narration", "Withdrawal", "Deposit", "Instrument number", "Balance"])
    for x in tx: ws.append([x[k] for k in CANONICAL])
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=4):
        for cell in row: cell.number_format = '#,##,##0.00'
    for cell in ws['F'][1:]: cell.number_format = '#,##,##0.00'
    for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="0F766E")
    for col, width in {"A":16,"B":48,"C":16,"D":16,"E":24,"F":16}.items(): ws.column_dimensions[col].width = width
    check = wb.create_sheet("Validation"); check.append(["Field", "Amount"])
    rows = [("Opening balance from statement",opening),("Total parsed withdrawals",total_w),("Total parsed deposits",total_d)]
    if declared_withdrawals is not None: rows.append(("Declared withdrawals from statement", declared_withdrawals))
    if declared_deposits is not None: rows.append(("Declared deposits from statement", declared_deposits))
    if declared_withdrawals is not None or declared_deposits is not None: rows.append(("Statement total cross-check", "PASS" if statement_totals_valid else "WARNING - printed totals differ; transaction-level reconciliation used"))
    if source_balance_unreliable: rows.append(("Running balance validation", "SOURCE UNRELIABLE - parsed totals match printed totals; opening/closing are assumed from a transaction anchor"))
    rows += [("Calculated closing balance",computed),("Closing balance from statement",closing),("Source transaction records", expected_source_count),("Parsed transaction records", len(tx)),("Transaction count validation", "PASS" if coverage_valid else "FAIL"),("Source coverage validation", "PASS" if coverage_valid else "FAIL"),("Financial validation", "PASS" if financial_valid else "FAIL"),("Narration validation", "PASS" if narration_valid else "FAIL"),("Release validation", "PASS" if financial_valid and narration_valid else "FAIL")]
    for row in rows: check.append(row)
    for cell in check['B'][1:]:
        if isinstance(cell.value, (Decimal, int, float)) and not isinstance(cell.value, bool): cell.number_format = '#,##,##0.00'
    for c in check[1]: c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="0F766E")
    check.column_dimensions["A"].width=36; check.column_dimensions["B"].width=22
    wb.save(out); return out.name

def api_profile_payload(profile_id: str) -> dict | None:
    """Return the versioned profile shape consumed by BS Analyzer."""
    profile_path = PROFILES / f"{Path(profile_id).name}.json"
    if not profile_path.exists():
        return None
    try:
        data = json.loads(profile_path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    # Migrate earlier UPG-validated profiles on first access. They were saved
    # before the BS Analyzer JavaScript contract existed, but their source
    # validation evidence is still valid. Imported profiles retain their own
    # certified strings unchanged.
    if not data.get("detection_code") or not data.get("parser_code"):
        headers = data.get("header_signature", [])
        if headers:
            detection_code, parser_code = certified_javascript_code(headers, data.get("last_validated_strategy"))
            data.update({
                "version": int(data.get("version", 0)) + 1,
                "detection_code": detection_code,
                "parser_code": parser_code,
                "bank_name": data.get("bank_name", "Unknown"),
                "format_name": data.get("format_name", "PDF Statement"),
                "validation": data.get("validation") or {"status": "pass", "financial_pass": True, "narration_pass": True, "balance_chain_pass": True, "transaction_count": data.get("validated_observations")},
            })
            profile_path.write_text(json.dumps(data, indent=2), encoding="utf-8")
    return {
        "profile_id": profile_id,
        "version": int(data.get("version", 1)),
        "bank_name": data.get("bank_name", "Unknown"),
        "format_name": data.get("format_name", "PDF Statement"),
        "layout_fingerprint": data.get("layout_fingerprint", ""),
        "detection_code": data.get("detection_code"),
        "parser_code": data.get("parser_code"),
        "strategy": data.get("last_validated_strategy", data.get("strategy", "detected_table")),
        "columns": data.get("columns", {}),
        "rules": data.get("rules", {}),
        "validation": data.get("validation", {"status": "pass"}),
        "parent_profile_id": data.get("parent_profile"),
        "certification": data.get("certification", {"status": "certified", "source": data.get("upg_source", "upg_native")}),
        "profile_origin": data.get("upg_source", "upg_native"),
    }

def post_completion_webhook(job_id: str, status: str, profile_id: str | None = None, error: str | None = None) -> None:
    """Best-effort signed notification; polling remains the reliable fallback."""
    if not (UPG_WEBHOOK_URL and UPG_WEBHOOK_SECRET):
        return
    payload = {"job_id": job_id, "status": status}
    if profile_id: payload["profile_id"] = profile_id
    if error: payload["error"] = error
    body = json.dumps(payload, separators=(",", ":")).encode()
    signature = hmac.new(UPG_WEBHOOK_SECRET.encode(), body, hashlib.sha256).hexdigest()
    request = urllib.request.Request(UPG_WEBHOOK_URL, data=body, headers={"Content-Type": "application/json", "X-UPG-Signature": f"sha256={signature}"}, method="POST")
    try:
        with urllib.request.urlopen(request, timeout=15):
            pass
    except (urllib.error.URLError, urllib.error.HTTPError):
        # BS Analyzer can always poll the job; never retry aggressively.
        pass

def retry_parser_job(job_id: str, path: Path, fallback_open: str, fallback_close: str) -> None:
    """Keep the UPG working until a fully validated candidate is found."""
    touch_worker(job_id, processing=True, valid=False, status="processing",
                 started_at=datetime.utcnow().isoformat(timespec="seconds") + "Z",
                 message="UPG worker acquired the job and is preparing parser candidates.")
    with JOBS_LOCK:
        saved_state = dict(JOBS.get(job_id, {}))
    round_number = int(saved_state.get("retry_round", 0) or 0)
    failed_candidates: set[str] = {str(item) for item in saved_state.get("failed_candidates", [])}
    attempted_candidates = int(saved_state.get("attempted_candidates", 0) or 0)
    skipped_candidates = int(saved_state.get("skipped_candidates", 0) or 0)
    validated_strategy = saved_text_strategy(path)
    large_pdf = is_large_pdf(path)
    # A sampled original-PDF geometry profile is stronger than text order.
    # Use it first whenever it exists, including for borderless statements.
    geometry_ready = large_pdf and sampled_geometry_profile(path) is not None
    text_first = large_pdf and not geometry_ready and prefers_running_balance_text(path)
    diagnostic_rules: set[str] = {str(item) for item in saved_state.get("diagnostic_rules", [])}
    planned_strategies: list[str] = [str(item) for item in saved_state.get("planned_strategies", [])]
    rounds_this_lease = 0
    while True:
        round_number += 1
        rounds_this_lease += 1
        if job_cancel_requested(job_id):
            patch_job(job_id, processing=False, valid=False, status="cancelled",
                      message="UPG job was cancelled before its next parser retry round.")
            clear_pdf_password(path)
            return
        # Keep API clients informed before starting expensive PDF work.
        with JOBS_LOCK:
            job = JOBS.get(job_id, {})
            job.update({
                "processing": True,
                "valid": False,
                "status": "processing",
                "message": f"UPG retry round {round_number}: inspecting the source layout and preparing parser candidates.",
                "retry_round": round_number,
                "attempted_candidates": attempted_candidates,
                "skipped_candidates": skipped_candidates,
                "worker_heartbeat_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
            })
            JOBS[job_id] = job
            persist_job_locked(job_id)
        latest = None
        errors = []
        repair_context = f"UPG self-healing round {round_number}. No validated parser candidate has been found yet."
        # On a long statement, candidate parsing already works from cached
        # full-document evidence.  Diagnostic AI needs only representative
        # layout evidence, not the entire 252-page text layer.  Keep this
        # outside the per-candidate loop too: one controlled diagnosis per
        # round prevents seven serial API waits after seven failed candidates.
        diagnostic_evidence = sampled_pdf_text(path) if large_pdf else cached_pdf_text(path)
        # Each round includes a fresh AI-generated layout candidate. It is not
        # a hand-written parser for the uploaded bank; the model proposes a
        # header/column profile from the current source evidence, which must
        # still pass coverage, financial, and narration validation.
        # The plan adapts after failure: once table candidates fail, prioritize
        # new AI addenda and page-aware text candidates over already-explored
        # layouts. Candidate memory below prevents duplicate validation work.
        initial_candidates = ([("unsigned_running_balance_text", False), ("running_balance_text", False), ("geometry_profile", False), ("value_date_unsigned", False), ("page_text_unsigned", False), (None, False), (None, True)] if text_first
            else [("geometry_profile", False), ("value_date_unsigned", False), ("unsigned_running_balance_text", False), ("running_balance_text", False), ("page_text_unsigned", False), (None, False), (None, True)] if large_pdf
            else [(None, False), (None, True), ("value_date_unsigned", False), ("running_balance_text", False), ("unsigned_running_balance_text", False), ("page_text_unsigned", False)])
        if validated_strategy in {"geometry_profile", "running_balance_text", "unsigned_running_balance_text", "value_date_unsigned", "page_text_unsigned"}:
            initial_candidates = [(validated_strategy, False)] + [item for item in initial_candidates if item[0] != validated_strategy]
        elif validated_strategy == "detected_table":
            # Reuse the validated table mapping, but for a large statement run
            # it through the sampled original-PDF geometry path first.
            preferred = ("geometry_profile", False) if large_pdf else (None, False)
            initial_candidates = [preferred] + [item for item in initial_candidates if item != preferred]
        candidates = (initial_candidates if round_number == 1
            else [(None, True), ("geometry_profile", False), ("value_date_unsigned", False), ("page_text_unsigned", False), ("unsigned_running_balance_text", False), ("running_balance_text", False), (None, False)])
        if large_pdf:
            # Do not spend minutes rediscovering a table across every page.
            # Original-PDF geometry and a controlled AI addendum are the two
            # primary paths; text candidates must first prove themselves on
            # the representative sample below.
            candidates = [candidate for candidate in candidates if candidate != (None, False)]
        if planned_strategies:
            def plan_key(candidate: tuple[str | None, bool]) -> int:
                name = "ai_layout_addendum" if candidate[1] else (candidate[0] or "detected_table")
                return planned_strategies.index(name) if name in planned_strategies else len(planned_strategies)
            candidates = sorted(candidates, key=plan_key)
        new_candidates_this_round = 0
        for strategy, force_ai_profile in candidates:
            try:
                if job_cancel_requested(job_id):
                    patch_job(job_id, processing=False, valid=False, status="cancelled",
                              message="UPG job was cancelled safely between parser candidates.")
                    clear_pdf_password(path)
                    return
                candidate_name = "AI layout addendum" if force_ai_profile else (strategy or "detected transaction table")
                if large_pdf and not force_ai_profile and not sample_candidate_plausible(path, strategy):
                    skipped_candidates += 1
                    errors.append(f"{candidate_name}: rejected by sampled original-PDF structure")
                    continue
                with JOBS_LOCK:
                    job = JOBS.get(job_id, {})
                    job.update({
                        "processing": True,
                        "valid": False,
                        "status": "processing",
                        "message": f"UPG retry round {round_number}: testing {candidate_name} and running financial, narration, coverage, and balance checks.",
                        "retry_round": round_number,
                        "attempted_candidates": attempted_candidates,
                        "skipped_candidates": skipped_candidates,
                        "worker_heartbeat_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
                    })
                    JOBS[job_id] = job
                    persist_job_locked(job_id)
                candidate = parse_statement(path, fallback_open, fallback_close, strategy, force_ai_profile, repair_context)
                if job_cancel_requested(job_id):
                    patch_job(job_id, processing=False, valid=False, status="cancelled",
                              message="UPG job was cancelled safely after the current extraction step.")
                    clear_pdf_password(path)
                    return
                latest = candidate
                # A failed mapping/strategy combination is never tested again
                # within this UPG job. AI addenda must propose a materially new
                # layout mapping before they receive another validation attempt.
                signature = hashlib.sha256(json.dumps({"strategy": strategy or "detected_table", "ai_addendum": force_ai_profile, "columns": candidate[10]}, sort_keys=True).encode()).hexdigest()
                if signature in failed_candidates:
                    skipped_candidates += 1
                    continue
                new_candidates_this_round += 1
                attempted_candidates += 1
                if candidate[6] and candidate[7]:
                    tx, op, cl, wd, dp, calculated, financial_valid, narration_valid, unmatched, headers, columns, parent_profile, coverage_valid, expected_source_count, layout_fingerprint, declared_wd, declared_dp, statement_totals_valid = candidate
                    with JOBS_LOCK:
                        job_context = JOBS.get(job_id, {})
                    profile_id = save_profile(
                        headers, columns, parent_profile, strategy or "detected_table", force_ai_profile,
                        layout_fingerprint, sorted(diagnostic_rules),
                        validation={
                            "status": "pass", "financial_pass": True, "narration_pass": True,
                            "balance_chain_pass": not bool(columns.get("_source_balance_unreliable")),
                            "balance_chain_exception": bool(columns.get("_source_balance_unreliable")),
                            "manual_review_required": bool(columns.get("_source_balance_unreliable")),
                            "review_message": "Running balances are unreliable. Parsed totals match printed totals; opening and closing are assumed from a transaction anchor. Review the source statement.",
                            "transaction_count": len(tx),
                            "source_transaction_count": expected_source_count,
                            "source_coverage_pass": bool(coverage_valid),
                        },
                        bank_name=str(job_context.get("bank_name") or "Unknown"),
                        format_name=f"{path.suffix.lower().lstrip('.') or 'pdf'} statement".upper(),
                    )
                    name = export_excel(tx, op, cl, wd, dp, calculated, financial_valid, narration_valid, coverage_valid, expected_source_count, declared_wd, declared_dp, statement_totals_valid, bool(columns.get("_source_balance_unreliable")))
                    with JOBS_LOCK:
                        balance_note = " Running-balance column: SOURCE UNRELIABLE; parsed totals match printed totals and assumed opening/closing were used. Review the source statement." if columns.get("_source_balance_unreliable") else " Balance-chain validation: PASS."
                        JOBS[job_id] = {"processing": False, "valid": True, "message": f"Validated after {round_number} UPG retry rounds. Parsed {len(tx)} transactions. Opening {indian_amount(op)} − withdrawals {indian_amount(wd)} + deposits {indian_amount(dp)} = {indian_amount(calculated)}; declared closing balance is {indian_amount(cl)}. Source coverage: PASS. Financial validation: PASS. Narration validation: PASS.{balance_note}", "download": "/download/" + name}
                        JOBS[job_id].update({"status": "completed", "profile_id": profile_id, "retry_round": round_number, "attempted_candidates": attempted_candidates, "skipped_candidates": skipped_candidates})
                        persist_job_locked(job_id)
                    post_completion_webhook(job_id, "completed", profile_id)
                    clear_pdf_password(path)
                    return
                failed_candidates.add(signature)
                repair_context = (f"UPG self-healing round {round_number}: candidate extracted {len(candidate[0])} of "
                    f"{candidate[13]} source records; source coverage={'PASS' if candidate[12] else 'FAIL'}, "
                    f"financial={'PASS' if candidate[6] else 'FAIL'}, narration={'PASS' if candidate[7] else 'FAIL'}. "
                    "Propose a safe header/column addendum only; do not weaken validation.")
            except Exception as error:
                errors.append(str(error))
                if "OCR_REQUIRED:" in str(error) or "PASSWORD_REQUIRED:" in str(error):
                    message = ("Not validated. This is an image-only PDF, so UPG has locked parser creation and Excel export until OCR is available. The statement was not treated as a one-row parser."
                        if "OCR_REQUIRED:" in str(error)
                        else "Not validated. This PDF is password protected. Enter the correct PDF password and resubmit it; UPG has stopped instead of retrying unreadable encrypted content.")
                    with JOBS_LOCK:
                        JOBS[job_id] = {"processing": False, "valid": False, "status": "failed", "message": message}
                        persist_job_locked(job_id)
                    clear_pdf_password(path)
                    return
        if latest is not None and new_candidates_this_round:
            tx, op, cl, wd, dp, calculated, financial_valid, narration_valid, unmatched, headers, columns, parent_profile, coverage_valid, expected_source_count, layout_fingerprint, declared_wd, declared_dp, statement_totals_valid = latest
            detail = f"UPG retry round {round_number}: tested {attempted_candidates} distinct candidates and skipped {skipped_candidates} duplicate failures. Current best candidate has {len(tx)} of {expected_source_count} source records; financial validation and narration validation are not yet passed."
        elif latest is not None:
            detail = f"UPG retry round {round_number}: tested {attempted_candidates} distinct candidates and skipped {skipped_candidates} duplicate failures. All known candidates were skipped; UPG is requesting a materially new AI layout addendum."
        else:
            detail = f"UPG retry round {round_number}: no safe candidate was produced yet; it is continuing with new layout attempts."
        # A diagnosis is guidance for the *next* retry round. It must never
        # weaken validation or cause the same failed candidate to run again.
        # Calling this once rather than after each failed candidate is the
        # largest safe speed-up for long PDFs.
        investigation = ai_diagnose_failure(diagnostic_evidence, repair_context)
        diagnostic_rules.update(investigation["rules"])
        planned_strategies = investigation["strategies"]
        if investigation["rules"] or planned_strategies:
            detail += " UPG recorded a new layout investigation plan for the next round."
        yield_to_queue = rounds_this_lease >= WORKER_LEASE_ROUNDS
        with JOBS_LOCK:
            job = JOBS.get(job_id, {})
            job.update({"processing": True, "valid": False,
                        "status": "queued" if yield_to_queue else "processing",
                        "message": (f"UPG completed retry round {round_number} and yielded its worker fairly; "
                                    "it will continue automatically after other queued statements receive a turn."
                                    if yield_to_queue else detail),
                        "retry_round": round_number, "attempted_candidates": attempted_candidates,
                        "skipped_candidates": skipped_candidates,
                        "failed_candidates": sorted(failed_candidates),
                        "diagnostic_rules": sorted(diagnostic_rules),
                        "planned_strategies": planned_strategies})
            job["worker_heartbeat_at"] = datetime.utcnow().isoformat(timespec="seconds") + "Z"
            JOBS[job_id] = job
            persist_job_locked(job_id)
        if yield_to_queue:
            # Place this still-unvalidated job at the end of the FIFO queue.
            # Its persisted candidate memory prevents it from repeating old
            # failures when its next fair worker lease begins.
            submit_retry_job(job_id, path, fallback_open, fallback_close)
            return
        # Do not terminate on a failed validation. This pause avoids a busy loop
        # while allowing a long-running UPG job to continue beyond five minutes.
        time.sleep(8)

def run_retry_job(job_id: str, path: Path, fallback_open: str, fallback_close: str) -> None:
    """Turn an unexpected worker exception into a durable, actionable status."""
    with JOBS_LOCK:
        if JOBS.get(job_id, {}).get("status") == "cancelled":
            refresh_queue_positions_locked()
            return
        ACTIVE_JOB_IDS.add(job_id)
        refresh_queue_positions_locked()
    try:
        retry_parser_job(job_id, path, fallback_open, fallback_close)
    except Exception as error:
        message = ("UPG stopped safely before certification. No parser was saved and no Excel was released. "
                   f"Worker error: {type(error).__name__}: {str(error)[:300]}")
        replace_job(job_id, {
            "processing": False, "valid": False, "status": "failed", "message": message,
            "source_file": path.name, "fallback_open": fallback_open, "fallback_close": fallback_close,
        })
        post_completion_webhook(job_id, "failed", error=message)
        clear_pdf_password(path)
    finally:
        with JOBS_LOCK:
            ACTIVE_JOB_IDS.discard(job_id)
            refresh_queue_positions_locked()

def submit_retry_job(job_id: str, path: Path, fallback_open: str, fallback_close: str) -> None:
    """Queue bounded background work; preserve the job before it starts."""
    queued_at = datetime.utcnow().isoformat(timespec="seconds") + "Z"
    patch_job(job_id, processing=True, valid=False, status="queued", queued_at=queued_at,
              worker_capacity=WORKER_CAPACITY,
              message="UPG job is queued for a parser-engine worker.")
    with JOBS_LOCK:
        refresh_queue_positions_locked()
        job = JOBS.get(job_id, {})
        position = int(job.get("queue_position", 0) or 0)
        depth = int(job.get("queue_depth", 0) or 0)
        retry_round = int(job.get("retry_round", 0) or 0)
        job["message"] = (
            f"UPG is waiting for its next fair worker turn: position {position} of {depth}. "
            + (f"It has safely completed {retry_round} retry round(s) and will continue automatically."
               if retry_round else "Parser generation will start automatically when a worker is free.")
        )
        JOBS[job_id] = job
        persist_job_locked(job_id)
    JOB_EXECUTOR.submit(run_retry_job, job_id, path, fallback_open, fallback_close)

def execute_certified_profile(profile_id: str, path: Path, fallback_open: str = "", fallback_close: str = "") -> dict:
    """Run UPG's native, certified extraction against the original source.

    Geometry-based profiles cannot be reproduced safely by a generic
    text-only JavaScript sandbox.  This API is therefore the authoritative
    execution path for BS Analyzer after it detects a UPG profile.
    """
    profile_path = PROFILES / f"{Path(profile_id).name}.json"
    if not profile_path.exists():
        raise ValueError("Certified UPG profile was not found")
    profile = json.loads(profile_path.read_text(encoding="utf-8"))
    validation = profile.get("validation") or {}
    if validation.get("status") != "pass" or not profile.get("certification"):
        raise ValueError("UPG profile is not certified")
    strategy = str(profile.get("last_validated_strategy") or "geometry_profile")
    result = parse_statement(path, fallback_open, fallback_close, strategy)
    tx, opening, closing, total_w, total_d, computed, financial_pass, narration_pass, unmatched, _headers, columns, _parent, coverage_pass, source_count, _fingerprint, declared_wd, declared_dp, totals_pass = result
    special_balance_exception = bool(columns.get("_source_balance_unreliable"))
    execution_validation = {
        "status": "pass" if financial_pass and narration_pass else "fail",
        "financial_pass": bool(financial_pass),
        "narration_pass": bool(narration_pass),
        "balance_chain_pass": not special_balance_exception,
        "balance_chain_exception": special_balance_exception,
        "manual_review_required": special_balance_exception,
        "source_coverage_pass": bool(coverage_pass),
        "transaction_count": len(tx),
        "source_transaction_count": source_count,
        "statement_totals_pass": bool(totals_pass),
    }
    if not (financial_pass and narration_pass):
        raise ValueError(f"Certified profile did not validate this statement: financial={financial_pass}, narration={narration_pass}, coverage={coverage_pass}")
    def output_row(row: dict) -> dict:
        narration = row.get("narration", "")
        instrument = row.get("instrument_number", "")
        return {
            "date": row.get("date", ""), "narration": narration, "particulars": narration,
            "withdrawal": float(row.get("withdrawal") or 0), "deposit": float(row.get("deposit") or 0),
            "instrument_number": instrument, "chqNo": instrument,
            "balance": float(row.get("balance")) if row.get("balance") is not None else None,
        }
    return {
        "profile_id": profile_id, "profile_version": int(profile.get("version", 1)),
        "transactions": [output_row(row) for row in tx],
        "opening_balance": float(opening), "closing_balance": float(closing),
        "total_withdrawals": float(total_w), "total_deposits": float(total_d),
        "calculated_closing_balance": float(computed),
        "declared_withdrawals": float(declared_wd) if declared_wd is not None else None,
        "declared_deposits": float(declared_dp) if declared_dp is not None else None,
        "validation": execution_validation,
    }

def recover_persisted_jobs() -> None:
    """Recover safe unprotected jobs after a Railway restart.

    Passwords are intentionally never stored. An encrypted PDF therefore asks
    for a fresh password rather than attempting an unsafe or impossible resume.
    """
    for record in JOBS_DIR.glob("*.json"):
        try:
            job = json.loads(record.read_text(encoding="utf-8"))
            job_id = record.stem
            if not isinstance(job, dict):
                continue
            JOBS[job_id] = job
            if not job.get("processing"):
                continue
            source_file = str(job.get("source_file", ""))
            source_path = UPLOADS / source_file
            if job.get("password_provided"):
                patch_job(job_id, processing=False, valid=False, status="failed",
                          message="UPG restarted while this protected PDF was running. For security the password was not saved; enter it again to start a fresh validated job.")
            elif source_file and source_path.exists():
                submit_retry_job(job_id, source_path, str(job.get("fallback_open", "")), str(job.get("fallback_close", "")))
            else:
                patch_job(job_id, processing=False, valid=False, status="failed",
                          message="UPG restarted and the temporary source file is unavailable. No parser was saved; upload the statement again.")
        except Exception:
            # A malformed old job record must never prevent the service start.
            continue

recover_persisted_jobs()
threading.Thread(target=queue_supervisor, name="upg-queue-supervisor", daemon=True).start()

class App(BaseHTTPRequestHandler):
    def json(self, data, status=200):
        payload=json.dumps(data).encode(); self.send_response(status); self.send_header("Content-Type","application/json"); self.send_header("Content-Length",str(len(payload))); self.end_headers(); self.wfile.write(payload)
    def api_authorized(self) -> bool:
        if not UPG_API_KEY:
            self.json({"error": "UPG API is not configured"}, HTTPStatus.SERVICE_UNAVAILABLE)
            return False
        supplied = self.headers.get("Authorization", "")
        if not hmac.compare_digest(supplied, f"Bearer {UPG_API_KEY}"):
            self.json({"error": "Unauthorized"}, HTTPStatus.UNAUTHORIZED)
            return False
        return True
    def multipart_fields(self) -> dict:
        ctype = self.headers.get("Content-Type", "")
        match = re.search(r"boundary=(.+)", ctype)
        if not match:
            raise ValueError("multipart/form-data with a boundary is required")
        boundary = match.group(1).strip('"').encode()
        body = self.rfile.read(int(self.headers["Content-Length"]))
        fields = {}
        for part in body.split(b"--" + boundary):
            head, _, data = part.partition(b"\r\n\r\n")
            name = re.search(br'name="([^"]+)"', head)
            if not name:
                continue
            key = name.group(1).decode()
            value = data.rstrip(b"\r\n-")
            filename = re.search(br'filename="([^"]*)"', head)
            fields[key] = (filename.group(1).decode(errors="ignore"), value) if filename else value.decode(errors="replace")
        return fields
    def start_api_job(self, fields: dict) -> str:
        if "file" not in fields:
            raise ValueError("file is required")
        filename, content = fields["file"]
        safe = Path(filename).name
        saved = UPLOADS / f"{uuid.uuid4().hex}-{safe}"
        saved.write_bytes(content)
        register_pdf_password(saved, str(fields.get("password", fields.get("pdf_password", ""))))
        job_id = uuid.uuid4().hex
        opening = fields.get("ob", fields.get("opening", ""))
        closing = fields.get("cb", fields.get("closing", ""))
        with JOBS_LOCK:
            JOBS[job_id] = {"processing": True, "valid": False, "status": "pending", "message": "UPG is creating and validating parser candidates.", "submitted_at": timestamp_now(), "client_heartbeat_at": timestamp_now(), "bs_analyzer_statement_id": fields.get("bs_analyzer_statement_id"), "bank_name": fields.get("bank_name", "Unknown"), "source_format": fields.get("source_format", ""), "source_file": saved.name, "fallback_open": opening, "fallback_close": closing, "password_provided": bool(fields.get("password", fields.get("pdf_password", "")))}
            persist_job_locked(job_id)
        submit_retry_job(job_id, saved, opening, closing)
        return job_id
    def do_GET(self):
        parsed_url = urlparse(self.path)
        path = parsed_url.path
        if path == "/queue-status":
            # Safe operational telemetry: no tenant names, files, or parser
            # details are exposed.  This lets BS Analyzer distinguish a real
            # queue from a stalled request without exposing job data.
            with JOBS_LOCK:
                queued_count, active_count = queue_snapshot_locked()
            self.json({"queued_jobs": queued_count, "active_workers": active_count,
                       "worker_capacity": WORKER_CAPACITY,
                       "available_workers": max(0, WORKER_CAPACITY - active_count)})
            return
        if path.startswith("/parser-jobs/"):
            if not self.api_authorized(): return
            job_id = Path(path).name
            touch_client(job_id)
            with JOBS_LOCK:
                job = JOBS.get(job_id)
            if not job:
                self.json({"job_id": job_id, "status": "failed", "error": "Job not found"}, 404); return
            status = job.get("status", "processing" if job.get("processing") else ("completed" if job.get("valid") else "failed"))
            response = {
                "job_id": job_id,
                "status": status,
                # BS Analyzer renders this verbatim as the UPG activity feed.
                "message": job.get("message", "UPG is preparing parser candidates."),
                "retry_round": job.get("retry_round", 0),
                "attempted_candidates": job.get("attempted_candidates", 0),
                "skipped_candidates": job.get("skipped_candidates", 0),
                "queue_position": job.get("queue_position"),
                "queue_depth": job.get("queue_depth", 0),
                "worker_capacity": job.get("worker_capacity", WORKER_CAPACITY),
                "active_workers": len(ACTIVE_JOB_IDS),
                "worker_heartbeat_at": job.get("worker_heartbeat_at"),
            }
            if job.get("profile_id"): response["profile_id"] = job["profile_id"]
            if status == "failed": response["error"] = job.get("message", "Parser generation failed")
            self.json(response); return
        if path == "/parser-profiles":
            if not self.api_authorized(): return
            query = parse_qs(parsed_url.query)
            fingerprint = query.get("fingerprint", [""])[0]
            try: page = max(1, int(query.get("page", ["1"])[0]))
            except ValueError: page = 1
            try: limit = min(100, max(1, int(query.get("limit", ["100"])[0])))
            except ValueError: limit = 100
            matches = []
            for profile_path in PROFILES.glob("*.json"):
                profile = api_profile_payload(profile_path.stem)
                if profile and (not fingerprint or profile.get("layout_fingerprint") == fingerprint):
                    matches.append(profile)
            matches.sort(key=lambda item: (item.get("bank_name", ""), item.get("profile_id", "")))
            if fingerprint:
                # Preserve the original lookup contract used by BS Analyzer.
                self.json(matches); return
            # No fingerprint means an authenticated registry sync. Never expose
            # profiles through the public HTML endpoint.
            total = len(matches)
            start = (page - 1) * limit
            profiles = matches[start:start + limit]
            self.json({"profiles": profiles, "count": total, "page": page, "limit": limit, "has_more": start + limit < total, "next_page": page + 1 if start + limit < total else None}); return
        if path.startswith("/parser-profiles/"):
            if not self.api_authorized(): return
            profile = api_profile_payload(Path(path).name)
            if not profile:
                self.json({"error": "Profile not found"}, 404); return
            self.json(profile); return
        if path == "/":
            data=HTML.encode(); self.send_response(200); self.send_header("Content-Type","text/html; charset=utf-8"); self.send_header("Content-Length",str(len(data))); self.end_headers(); self.wfile.write(data); return
        if path.startswith("/status/"):
            job_id = Path(path).name
            touch_client(job_id)
            with JOBS_LOCK:
                status = JOBS.get(job_id)
            if status is None: self.json({"processing": False, "interrupted": True, "valid": False, "message": "UPG was restarted while this retry job was running. The saved parser is unchanged; click Parse and validate to start a fresh job for this uploaded statement."}, 404)
            else: self.json(status)
            return
        if path.startswith("/download/"):
            file=EXPORTS / Path(path).name
            if file.exists():
                self.send_response(200); self.send_header("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); self.send_header("Content-Disposition",f'attachment; filename="{file.name}"'); self.send_header("Content-Length",str(file.stat().st_size)); self.end_headers(); shutil.copyfileobj(file.open("rb"),self.wfile); return
        self.send_error(404)
    def do_POST(self):
        path = urlparse(self.path).path
        cancel_match = re.fullmatch(r"/parser-jobs/([^/]+)/cancel", path)
        if cancel_match:
            if not self.api_authorized(): return
            job_id = cancel_match.group(1)
            if not cancel_job(job_id):
                self.json({"error": "Job not found"}, HTTPStatus.NOT_FOUND)
            else:
                self.json({"job_id": job_id, "status": "cancelled"})
            return
        if path == "/parser-jobs":
            if not self.api_authorized(): return
            try:
                job_id = self.start_api_job(self.multipart_fields())
                self.json({"job_id": job_id}, HTTPStatus.ACCEPTED)
            except Exception as error:
                self.json({"error": str(error)}, HTTPStatus.BAD_REQUEST)
            return
        if path == "/parser-profiles/import":
            if not self.api_authorized(): return
            try:
                body = self.rfile.read(int(self.headers.get("Content-Length", "0")))
                incoming = json.loads(body)
                required = ("bank_name", "format_name", "detection_code", "parser_code")
                missing = [key for key in required if not str(incoming.get(key, "")).strip()]
                if missing:
                    raise ValueError("Missing required fields: " + ", ".join(missing))
                profile_id = str(incoming.get("profile_id") or hashlib.sha256((incoming["bank_name"] + incoming["format_name"] + incoming["detection_code"] + incoming["parser_code"]).encode()).hexdigest()[:16])
                stored = {
                    "version": int(incoming.get("version", 1)), "bank_name": incoming["bank_name"], "format_name": incoming["format_name"],
                    "layout_fingerprint": incoming.get("layout_fingerprint", ""), "detection_code": incoming["detection_code"], "parser_code": incoming["parser_code"],
                    "last_validated_strategy": incoming.get("extraction_strategy", incoming.get("strategy", "text-column-offsets")),
                    "columns": incoming.get("columns", {}), "rules": incoming.get("rules", {}),
                    "validation": {"status": "pass", "financial_pass": True, "narration_pass": True, "balance_chain_pass": True, "transaction_count": incoming.get("certification", {}).get("transaction_count")},
                    "certification": incoming.get("certification", {}), "parent_profile": incoming.get("parent_profile_id") or incoming.get("evolved_from_profile_id"), "upg_source": "bs_analyzer_import",
                }
                (PROFILES / f"{profile_id}.json").write_text(json.dumps(stored, indent=2), encoding="utf-8")
                self.json(api_profile_payload(profile_id), HTTPStatus.CREATED)
            except Exception as error:
                self.json({"error": str(error)}, HTTPStatus.BAD_REQUEST)
            return
        execute_match = re.fullmatch(r"/parser-profiles/([^/]+)/execute", path)
        if execute_match:
            if not self.api_authorized(): return
            saved = None
            try:
                fields = self.multipart_fields()
                if "file" not in fields:
                    raise ValueError("file is required")
                filename, content = fields["file"]
                saved = UPLOADS / f"{uuid.uuid4().hex}-{Path(filename).name}"
                saved.write_bytes(content)
                register_pdf_password(saved, str(fields.get("password", fields.get("pdf_password", ""))))
                result = execute_certified_profile(
                    execute_match.group(1), saved,
                    str(fields.get("opening", fields.get("ob", ""))),
                    str(fields.get("closing", fields.get("cb", ""))),
                )
                self.json(result)
            except Exception as error:
                self.json({"ok": False, "error": str(error)}, HTTPStatus.UNPROCESSABLE_ENTITY)
            finally:
                if saved is not None:
                    clear_pdf_password(saved)
                    try: saved.unlink(missing_ok=True)
                    except OSError: pass
            return
        if path != "/parse": self.send_error(404); return
        try:
            fields = self.multipart_fields()
            filename, content=fields["file"]; safe=Path(filename).name; saved=UPLOADS / f"{uuid.uuid4().hex}-{safe}"; saved.write_bytes(content)
            register_pdf_password(saved, str(fields.get("password", fields.get("pdf_password", ""))))
            job_id = uuid.uuid4().hex
            with JOBS_LOCK:
                JOBS[job_id] = {"processing": True, "valid": False, "status": "pending", "message": "UPG is creating and validating parser candidates.", "submitted_at": timestamp_now(), "client_heartbeat_at": timestamp_now(), "source_file": saved.name, "fallback_open": fields.get("opening", ""), "fallback_close": fields.get("closing", ""), "password_provided": bool(fields.get("password", fields.get("pdf_password", "")))}
                persist_job_locked(job_id)
            submit_retry_job(job_id, saved, fields.get("opening", ""), fields.get("closing", ""))
            self.json({"processing": True, "valid": False, "job": job_id, "message": "UPG is retrying parser candidates. Excel and profile creation remain locked until both validations pass."})
            return
            # UPG retry loop: do not treat the first failed strategy as final.
            result = None
            attempt_errors = []
            for strategy in (None, "running_balance_text", "unsigned_running_balance_text"):
                try:
                    candidate = parse_statement(saved, fields.get("opening",""), fields.get("closing",""), strategy)
                    if candidate[6] and candidate[7]:
                        result = candidate
                        break
                    result = candidate
                except Exception as retry_error:
                    attempt_errors.append(str(retry_error))
            if result is None:
                raise ValueError("UPG could not produce a readable parser after retrying its available strategies.")
            tx, op, cl, wd, dp, calculated, financial_valid, narration_valid, unmatched, headers, columns, parent_profile, coverage_valid, expected_source_count=result
            valid = financial_valid and narration_valid
            msg=f"Parsed {len(tx)} transactions. Opening {op:,.2f} − withdrawals {wd:,.2f} + deposits {dp:,.2f} = {calculated:,.2f}; declared closing balance is {cl:,.2f}."
            msg += f" Source coverage: {'PASS' if coverage_valid else 'FAIL'} ({len(tx)} of {expected_source_count} records). Financial validation: {'PASS' if financial_valid else 'FAIL'}. Narration validation: {'PASS' if narration_valid else 'FAIL'}."
            if unmatched: msg += f" Unmatched narrations: {len(unmatched)}."
            if valid:
                save_profile(headers, columns, parent_profile)
                name=export_excel(tx,op,cl,wd,dp,calculated,financial_valid,narration_valid,coverage_valid,expected_source_count); self.json({"valid":True,"message":msg,"download":"/download/"+name})
            else: self.json({"valid":False,"message":msg+" UPG retried its available parser strategies; the Excel export is withheld until both validations pass."})
        except Exception as e: self.json({"valid":False,"message":str(e)},400)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", os.environ.get("BANK_PARSER_PORT", "8080")))
    print(f"Listening on 0.0.0.0:{port}")
    ThreadingHTTPServer(("0.0.0.0", port), App).serve_forever()
