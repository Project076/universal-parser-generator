"""Specialist-worker contracts for the Universal Parser Generator.

The workers in this module are logical, deterministic pipeline specialists;
they are not fifty independent OS processes.  Each worker owns one immutable
pipeline step, reads only the certified knowledge scopes relevant to that
step, and can request an additive addendum from the supervisor when blocked.
"""
from __future__ import annotations

from typing import Mapping, Sequence


# One primary knowledge domain per immutable pipeline step.  Keeping this map
# explicit makes an accidental unscoped "give the model everything" change
# fail during application startup and in tests.
STEP_DOMAINS: dict[int, str] = {
    1: "source", 2: "source", 3: "source",
    4: "headers", 5: "geometry", 6: "record_boundaries",
    7: "narration", 8: "furniture", 9: "amounts", 10: "balances",
    11: "dates", 12: "furniture", 13: "classification",
    14: "coverage", 15: "transaction_count", 16: "financial",
    17: "balances", 18: "ledger_order", 19: "endpoints",
    20: "planning", 21: "planning", 22: "candidate_memory",
    23: "capability_composition", 24: "evidence", 25: "geometry",
    26: "layout_reuse", 27: "agent_context", 28: "compatibility",
    29: "learning", 30: "activation", 31: "export", 32: "code",
    33: "layout_reuse", 34: "execution", 35: "execution",
    36: "addenda", 37: "import", 38: "addenda", 39: "regression",
    40: "revision", 41: "lineage", 42: "provenance",
    43: "capability_composition", 44: "receipts", 45: "learning",
    46: "drift", 47: "repair_routing", 48: "evidence",
    49: "budget", 50: "certification",
}


DOMAIN_LIBRARY_SCOPES: dict[str, tuple[str, ...]] = {
    "source": ("source_readers", "native_structure", "source_shape"),
    "headers": ("header_semantics", "header_synonyms"),
    "geometry": ("measured_geometry", "column_roles", "source_cells"),
    "record_boundaries": ("transaction_boundaries", "page_continuations"),
    "narration": ("narration_boundaries", "continuation_merge", "amount_exclusion"),
    "furniture": ("headers_footers", "page_totals", "summaries", "bf_rows"),
    "amounts": ("money_tokens", "indian_punctuation", "source_amounts"),
    "balances": ("balance_tokens", "balance_direction", "balance_chain", "wrong_running_balance"),
    "dates": ("date_columns", "value_date", "date_normalization"),
    "classification": ("debit_credit", "balance_delta", "explicit_columns"),
    "coverage": ("source_coverage", "narration_coverage"),
    "transaction_count": ("date_count", "amount_count", "balance_count"),
    "financial": ("opening_closing", "movement_totals", "reconciliation"),
    "ledger_order": ("source_order", "page_order", "reverse_order"),
    "endpoints": ("opening_balance", "closing_balance", "summary_endpoints"),
    "planning": ("parser_plans", "source_strategy"),
    "candidate_memory": ("failed_candidates", "negative_knowledge"),
    "capability_composition": ("certified_capabilities", "composition_safety"),
    "evidence": ("measured_evidence", "candidate_scorecards"),
    "layout_reuse": ("layout_fingerprints", "certified_profiles"),
    "agent_context": ("scoped_lessons", "blocked_step_evidence"),
    "compatibility": ("additive_compatibility", "legacy_regression"),
    "learning": ("certified_lessons", "challenge_history"),
    "activation": ("activation_guards",),
    "export": ("certified_export",),
    "code": ("certified_code", "code_integrity"),
    "execution": ("remote_execution", "execution_failures"),
    "addenda": ("versioned_addenda", "parent_profiles"),
    "import": ("certified_import",),
    "regression": ("old_profile_replay", "addendum_regression"),
    "revision": ("atomic_revision",),
    "lineage": ("profile_lineage",),
    "provenance": ("capability_provenance",),
    "receipts": ("application_receipts",),
    "drift": ("capability_drift",),
    "repair_routing": ("blocked_step_routes", "versioned_addenda"),
    "budget": ("ai_budget", "retry_budget", "cache"),
    "certification": ("all_release_gates", "independent_audit"),
}


def build_specialist_worker_registry(manifest: Mapping[int, str]) -> dict[int, dict[str, object]]:
    """Build and validate one immutable specialist contract per step."""
    expected = set(range(1, 51))
    if set(manifest) != expected or set(STEP_DOMAINS) != expected:
        raise ValueError("Specialist registry requires exactly steps 1 through 50")
    registry: dict[int, dict[str, object]] = {}
    for number in sorted(manifest):
        name = str(manifest[number])
        domain = STEP_DOMAINS[number]
        scopes = DOMAIN_LIBRARY_SCOPES.get(domain)
        if not scopes:
            raise ValueError(f"No certified-knowledge scope for step {number}: {name}")
        registry[number] = {
            "worker_id": f"W{number:02d}_{name}",
            "step_number": number,
            "step_name": name,
            "domain": domain,
            "library_scopes": list(scopes),
            "learning_policy": "versioned_addendum_only",
            "on_blocked": "escalate_to_supervisor_with_measured_evidence",
            "replay_from_step": number,
            "may_overwrite_certified_learning": False,
        }
    if len({str(worker["worker_id"]) for worker in registry.values()}) != 50:
        raise ValueError("Every specialist worker must have a unique identity")
    return registry


def build_blocked_step_escalation(
    worker: Mapping[str, object],
    failure_type: str,
    rule_modules: Sequence[str] = (),
    rule_groups: Sequence[str] = (),
    upstream_steps: Sequence[str] = (),
) -> dict[str, object]:
    """Create the bounded agent-supervisor handoff for one blocked step."""
    step_number = int(worker["step_number"])
    return {
        "schema_version": 1,
        "event": "specialist_step_blocked",
        "worker_id": str(worker["worker_id"]),
        "blocked_step": f"S{step_number:02d}_{worker['step_name']}",
        "domain": str(worker["domain"]),
        "failure_type": str(failure_type),
        "certified_library_scopes": list(worker.get("library_scopes", [])),
        "target_rule_modules": list(dict.fromkeys(str(item) for item in rule_modules if str(item)))[:8],
        "target_rule_groups": list(dict.fromkeys(str(item) for item in rule_groups if str(item)))[:6],
        "upstream_steps_to_recheck": list(dict.fromkeys(str(item) for item in upstream_steps if str(item)))[:6],
        "supervisor_task": "propose_one_measured_addendum_for_this_worker",
        "learning_policy": "versioned_addendum_only",
        "replay_from_step": step_number,
        "forbidden": [
            "overwrite_certified_rule",
            "replace_certified_parser",
            "copy_foreign_geometry",
            "invent_source_values",
            "weaken_certification_gates",
            "advance_to_unrelated_step",
        ],
        "release_policy": "full_independent_certification_required",
    }
