"""Assessment layer: map extracted evidence onto the EUC review schema.

The scanner (:mod:`excel_xray.scan`) is the *evidence* layer. This module turns
that evidence into the reviewer-facing fields of the EUC (End-User Computing)
assessment — file-level and tab-level — matching the target schema.

Every field is a :class:`Field` carrying not just a value but its ``basis``, so
the output is honest about what is grounded in the file versus what still needs a
model, a human, or the wider corpus to answer:

* ``extracted``    — read directly from the workbook.
* ``derived``      — a heuristic over extracted metrics (with evidence).
* ``needs_llm``    — a narrative/judgement call for the model step.
* ``needs_human``  — not knowable from the file (e.g. usage frequency).
* ``needs_corpus`` — needs the whole folder of workbooks (e.g. duplication).

Step 1 implements the file-level ``extracted`` and ``derived`` fields; the rest
are returned with the correct basis and a ``null`` value, ready for later steps.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import asdict, dataclass, field
from collections import Counter

from .scan import WorkbookXray
from . import review

# --------------------------------------------------------------------- Field

Basis = str  # one of: extracted | derived | needs_llm | needs_human | needs_corpus


@dataclass
class Field:
    """One assessment field: its value, where it came from, and why."""

    value: object = None
    basis: Basis = "needs_human"
    confidence: float | None = None
    evidence: list[str] = field(default_factory=list)

    @classmethod
    def extracted(cls, value, evidence=None) -> "Field":
        return cls(value=value, basis="extracted", confidence=1.0,
                   evidence=evidence or [])

    @classmethod
    def derived(cls, value, confidence, evidence) -> "Field":
        return cls(value=value, basis="derived", confidence=round(confidence, 2),
                   evidence=list(evidence))

    @classmethod
    def pending(cls, basis: Basis, note: str = "") -> "Field":
        return cls(value=None, basis=basis, evidence=[note] if note else [])


# ------------------------------------------------------------- schema shapes


@dataclass
class FileAssessment:
    """File-level summary — the top block of the target schema."""

    # Fact Assessment
    file_id: Field = field(default_factory=Field)
    file_name: Field = field(default_factory=Field)
    scan_status: Field = field(default_factory=Field)
    scan_error: Field = field(default_factory=Field)
    sheet_count_total: Field = field(default_factory=Field)
    sheet_count_hidden: Field = field(default_factory=Field)
    business_area_process: Field = field(default_factory=Field)
    process: Field = field(default_factory=Field)
    sub_process: Field = field(default_factory=Field)
    purpose_of_file: Field = field(default_factory=Field)
    key_output_outcome: Field = field(default_factory=Field)
    complexity: Field = field(default_factory=Field)
    key_inputs: Field = field(default_factory=Field)
    source_system: Field = field(default_factory=Field)
    key_outputs: Field = field(default_factory=Field)
    usage_frequency: Field = field(default_factory=Field)
    completion_timeline: Field = field(default_factory=Field)
    euc_preparer: Field = field(default_factory=Field)
    output_recipient: Field = field(default_factory=Field)
    # Key AI Finding / Observation
    potential_duplication: Field = field(default_factory=Field)
    similar_duplicate_files: Field = field(default_factory=Field)
    potential_simplification: Field = field(default_factory=Field)
    potential_consolidation: Field = field(default_factory=Field)
    potential_automation: Field = field(default_factory=Field)
    potential_retirement: Field = field(default_factory=Field)
    # Workbook logic / Automation
    logic_type: Field = field(default_factory=Field)
    logic_types: Field = field(default_factory=Field)
    key_calculations_logic: Field = field(default_factory=Field)
    reconciliation_logic: Field = field(default_factory=Field)
    manual_intervention: Field = field(default_factory=Field)
    macros_vba_external_links: Field = field(default_factory=Field)


@dataclass
class TabAssessment:
    """Tab-level detail — the bottom block of the target schema (Step 2)."""

    tab_name: Field = field(default_factory=Field)
    tab_visibility: Field = field(default_factory=Field)
    tab_category: Field = field(default_factory=Field)
    tab_roles: Field = field(default_factory=Field)
    tab_purpose_description: Field = field(default_factory=Field)
    tab_information_analysis: Field = field(default_factory=Field)
    key_calculation_transformation_logic: Field = field(default_factory=Field)
    upstream_dependencies: Field = field(default_factory=Field)
    downstream_dependencies: Field = field(default_factory=Field)
    human_validation_required: Field = field(default_factory=Field)
    validation_reason: Field = field(default_factory=Field)


@dataclass
class Assessment:
    file: FileAssessment
    tabs: list[TabAssessment] = field(default_factory=list)
    error_summary: list[dict] = field(default_factory=list)
    hidden_groups: list[dict] = field(default_factory=list)
    input_groups: list[dict] = field(default_factory=list)


# ---------------------------------------------------------- workbook metrics


def _totals(wx: WorkbookXray) -> dict:
    """Roll per-sheet numbers up to the workbook, for the file-level fields."""
    cells = sum(s.populated_cells for s in wx.sheets)
    formulas = sum(s.formula_profile.get("total", 0) for s in wx.sheets)
    distinct = sum(s.formula_profile.get("distinct_skeletons", 0) for s in wx.sheets)
    functions: Counter = Counter()
    for s in wx.sheets:
        for fn, n in s.formula_profile.get("top_functions", []):
            functions[fn] += n
    cross_sheet = sum(s.formula_profile.get("cross_sheet_count", 0) for s in wx.sheets)
    hardcoded = sum(s.formula_profile.get("hardcoded_literal_count", 0) for s in wx.sheets)
    volatile = sum(s.formula_profile.get("volatile_count", 0) for s in wx.sheets)
    return {
        "cells": cells,
        "formulas": formulas,
        "distinct": distinct,
        "functions": functions,
        "cross_sheet": cross_sheet,
        "hardcoded": hardcoded,
        "volatile": volatile,
        "non_formula_cells": cells - formulas,
        "linked_workbooks": len(wx.external_links),
        "connections": len(wx.connections),
    }


# ------------------------------------------------------- deterministic fields


def _complexity(wx: WorkbookXray, t: dict) -> Field:
    """High / Medium / Low from data volume, calculation logic and linkage."""
    ev: list[str] = []
    score = 0

    if t["cells"] > 20_000:
        score += 2; ev.append(f"{t['cells']:,} populated cells (large)")
    elif t["cells"] > 2_000:
        score += 1; ev.append(f"{t['cells']:,} populated cells (moderate)")
    else:
        ev.append(f"{t['cells']:,} populated cells (small)")

    if t["formulas"] > 5_000 or t["distinct"] > 60:
        score += 2; ev.append(f"{t['formulas']:,} formulas / {t['distinct']} distinct calcs (heavy logic)")
    elif t["formulas"] > 200 or t["distinct"] > 15:
        score += 1; ev.append(f"{t['formulas']:,} formulas / {t['distinct']} distinct calcs (moderate logic)")
    elif t["formulas"]:
        ev.append(f"{t['formulas']:,} formulas / {t['distinct']} distinct calcs (light logic)")

    links = t["linked_workbooks"] + t["connections"]
    if links > 3:
        score += 2; ev.append(f"{links} external link(s)/connection(s)")
    elif links:
        score += 1; ev.append(f"{links} external link(s)/connection(s)")

    if len(wx.sheets) > 12:
        score += 1; ev.append(f"{len(wx.sheets)} sheets")
    if wx.has_vba:
        score += 2; ev.append("contains VBA macros")
    if wx.has_power_query:
        score += 1; ev.append("uses Power Query")

    verdict = "High" if score >= 5 else "Medium" if score >= 2 else "Low"
    # confidence rises as the score sits clearly inside a band, not on a boundary.
    conf = 0.6 + 0.1 * min(3, abs(score - 3.5))
    return Field.derived(verdict, min(conf, 0.9), ev)


def _logic_type(wx: WorkbookXray, t: dict) -> Field:
    """Reconciliation / Calculation / Data Transformation / Manual Input / Reporting."""
    fns = t["functions"]
    ev: list[str] = []
    scores: Counter = Counter()

    lookup = sum(fns[f] for f in ("VLOOKUP", "HLOOKUP", "XLOOKUP", "MATCH", "INDEX"))
    agg = sum(fns[f] for f in ("SUM", "SUMIF", "SUMIFS", "AVERAGE", "COUNT", "COUNTIF", "SUBTOTAL"))
    textfn = sum(fns[f] for f in ("TEXT", "CONCATENATE", "CONCAT", "LEFT", "RIGHT", "MID", "TRIM", "SUBSTITUTE"))
    cond = sum(fns[f] for f in ("IF", "IFS", "IFERROR"))

    if lookup:
        scores["Data Transformation"] += lookup
        ev.append(f"{lookup} lookup/match call(s) — retrieval or mapping, not proof of reconciliation")
    reads, _ = _dependency_maps(wx)
    if any(review.reconciliation(s, reads.get(s.name, set())) for s in wx.sheets):
        scores["Reconciliation"] += max(1, lookup + cond * 0.5)
        ev.append("reconciliation/variance labels detected")
    if agg:
        scores["Calculation"] += agg
        ev.append(f"{agg} aggregation call(s)")
    if textfn:
        scores["Data Transformation"] += textfn
        ev.append(f"{textfn} text-manipulation call(s)")
    if wx.has_power_query:
        scores["Data Transformation"] += 5
        ev.append("Power Query present")
    if wx.pivot_cache_sources:
        scores["Reporting"] += 3
        ev.append(f"{len(wx.pivot_cache_sources)} pivot source(s) — reporting")

    if any(review._matches(s.name, r"manual entry|manual input|manual adjustment") for s in wx.sheets):
        scores["Manual Input"] += 4
        ev.append("tab name suggests manual capture; actual workflow requires confirmation")

    if t["formulas"] and not scores:
        scores["Calculation"] += 1
        ev.append("worksheet formulas perform arithmetic or reference-based calculations")

    if not scores:
        return Field.derived("Other", 0.4, ev or ["no dominant logic signal"])
    primary = scores.most_common(1)[0][0]
    total = sum(scores.values())
    conf = 0.5 + 0.4 * (scores[primary] / total)
    ranked = [k for k, _ in scores.most_common()]
    return Field.derived(primary, min(conf, 0.9), ev + [f"ranked: {', '.join(ranked)}"])


def _key_inputs(wx: WorkbookXray) -> Field:
    inputs: list[str] = []
    inputs += [f"linked workbook: {review.short_source(x)}" for x in wx.external_links]
    inputs += [f"connection: {c.get('name') or c.get('type') or 'unnamed'}"
               for c in wx.connections]
    inputs += [f"pivot source: {p}" for p in wx.pivot_cache_sources]
    if inputs:
        return Field.extracted(inputs)
    return Field(value=None, basis="needs_llm",
                 evidence=["no external links/connections declared; embedded sources and their provenance require tab-level review"])


def _source_system(wx: WorkbookXray) -> Field:
    # Never expose connection strings: they can contain servers, usernames,
    # client folders, URLs or credentials. A generic workbook label such as
    # "tb" is an input-purpose hint, not proof of the originating system.
    systems = sorted({
        review.short_source(str(c.get("name") or c.get("description")).strip())
        for c in wx.connections
        if (c.get("name") or c.get("description"))
        and str(c.get("name") or c.get("description")).strip().lower()
        not in {"connection", "query", "external data", "tb", "trial balance"}
    })
    if systems:
        return Field.derived(systems, 0.65,
            ["Names from formal Excel connection metadata; confirm the underlying provider/system with the owner"])
    return Field(value="Not established — owner confirmation required",
                 basis="needs_human",
                 evidence=["No reliable formal system identifier was found; workbook labels such as 'tb' are not treated as source systems"])


def _euc_preparer(wx: WorkbookXray) -> Field:
    creator = (wx.core_props or {}).get("creator")
    last = (wx.core_props or {}).get("last_modified_by")
    names = [n for n in {creator, last} if n]
    if names:
        return Field(value=names, basis="derived", confidence=0.5,
                     evidence=["from document metadata (author/last-modified-by); "
                               "may reflect a template author, not the preparer"])
    return Field.pending("needs_human", "no author recorded in document metadata")


def _key_calculations(wx: WorkbookXray, t: dict) -> Field:
    if not t["formulas"]:
        return Field.extracted([], ["workbook has no formulas"])
    skeletons: Counter = Counter()
    for s in wx.sheets:
        for sk, n in s.formula_profile.get("top_skeletons", []):
            skeletons[sk] += n
    top = [f"{sk}  (×{n})" for sk, n in skeletons.most_common(8)]
    top_fns = [f"{fn}×{n}" for fn, n in t["functions"].most_common(8)]
    return Field.extracted(
        {"top_functions": top_fns, "top_formula_shapes": top},
        [f"{t['formulas']:,} formulas reduce to {t['distinct']} distinct shapes"],
    )


def _manual_intervention(wx: WorkbookXray, t: dict) -> Field:
    ev = [f"{t['non_formula_cells']:,} stored cells; these may be imported, pasted, labels or manual entries"]
    if t["hardcoded"]:
        ev.append(f"{t['hardcoded']} formula(s) contain hardcoded numeric constants; these may be embedded assumptions or overrides, not confirmed manual intervention")
    return Field(value="Not established — confirm which inputs are keyed, pasted, adjusted or overridden, by whom and how often.",
                 basis="needs_human", evidence=ev)


def _macros_links(wx: WorkbookXray) -> Field:
    external_refs = sum(s.formula_profile.get("external_count", 0) for s in wx.sheets)
    distinct_external = len({review.short_source(x).casefold() for x in wx.external_links})
    formal = len(wx.connections)
    ref_word = "reference" if external_refs == 1 else "references"
    book_word = "workbook" if distinct_external == 1 else "workbooks"
    summary = (
        f"VBA: {'present; use case requires owner confirmation' if wx.has_vba else 'none detected'}. "
        f"Power Query: {'present; refresh steps and business use require owner confirmation' if wx.has_power_query else 'none detected'}. "
        f"Formal data connections: {formal if formal else 'none detected'}. "
        f"External workbook links: {external_refs} {ref_word} across "
        f"{distinct_external} distinct external {book_word}."
    )
    return Field.extracted({
        "summary": summary,
        "vba": "present" if wx.has_vba else "none detected",
        "power_query": "present" if wx.has_power_query else "none detected",
        "formal_data_connections": formal,
        "external_workbook_link_references": external_refs,
        "distinct_external_workbooks": distinct_external,
    }, ["Mechanisms are reported separately; no use case is inferred from presence alone"])


# ------------------------------------------------- AI findings (Step 3)

_RECON_KEYWORDS = (
    "recon", "variance", "diff", "difference", "tolerance", "exception",
    "ageing", "aging", "aged", "unmatched", "tie-out", "tie out", "balance",
)
_SUPERSEDED_NAMES = (
    "old", "backup", "copy", " v1", "v1.", "draft", "archive", "tmp", "temp", "bak",
)


def _months_since(iso: str | None) -> int | None:
    """Whole months between an ISO timestamp and now; None if unparseable."""
    if not iso:
        return None
    try:
        s = str(iso).replace("Z", "+00:00")
        when = _dt.datetime.fromisoformat(s)
        if when.tzinfo is not None:
            when = when.replace(tzinfo=None)
    except (ValueError, TypeError):
        return None
    days = (_dt.datetime.now() - when).days
    return max(0, days // 30)


def _recon_signals(wx: WorkbookXray, t: dict) -> list[str]:
    """Shared reconciliation-pattern detection, reused by two fields."""
    fns = t["functions"]
    signals: list[str] = []
    matching = sum(fns.get(f, 0) for f in ("VLOOKUP", "HLOOKUP", "XLOOKUP", "MATCH", "INDEX"))
    if matching:
        signals.append(f"matching via {matching} lookup/match call(s)")
    if fns.get("ABS"):
        signals.append(f"{fns['ABS']} ABS() call(s) — variance/tolerance comparison")

    skeletons = [sk for s in wx.sheets for sk, _ in s.formula_profile.get("top_skeletons", [])]
    tol = [sk for sk in skeletons if ("<" in sk or ">" in sk) and "N" in sk and "IF" in sk.upper()]
    if tol:
        signals.append(f"tolerance/threshold test, e.g. {tol[0]}")

    headers = [h.lower() for s in wx.sheets for r in s.regions for h in r.headers if h]
    kw = sorted({k for h in headers for k in _RECON_KEYWORDS if k in h})
    if kw:
        signals.append("reconciliation-labelled header(s): " + ", ".join(kw))
    return signals


def _reconciliation_logic(wx: WorkbookXray, t: dict) -> Field:
    reads, _ = _dependency_maps(wx)
    reconciliations = []
    candidates = []
    for s in wx.sheets:
        item = review.reconciliation(s, reads[s.name])
        if item:
            reconciliations.append(item)
        elif candidate := review.reconciliation_candidate(s, reads[s.name]):
            candidates.append(candidate)
    status = "confirmed" if reconciliations else (
        "candidates_requiring_review" if candidates else "none_identified")
    value = {
        "count": len(reconciliations),
        "status": status,
        "reconciliations": reconciliations,
        "candidates_requiring_review": candidates,
        "evidence": [
            "A confirmed reconciliation must compare at least two observed sources and contain agreement, difference or exception evidence.",
            "Lookup, IF and variance formulas alone are retained as review candidates and are not counted.",
        ],
    }
    confidence = 0.75 if reconciliations else 0.6
    return Field.derived(value, confidence, value["evidence"])


def _simplification(wx: WorkbookXray, t: dict) -> Field:
    ops: list[str] = []
    if t["hardcoded"] > 5:
        ops.append(f"{t['hardcoded']} hardcoded numbers buried in formulas — lift to a "
                   "named inputs/assumptions block")
    if t["formulas"] > 50 and t["distinct"]:
        comp = t["formulas"] / t["distinct"]
        if comp < 3:
            ops.append(f"low formula reuse ({comp:.1f} per distinct shape) — many "
                       "one-off formulas to standardise")
    if t["volatile"]:
        ops.append(f"{t['volatile']} volatile function(s) (OFFSET/INDIRECT/…) — replace "
                   "with structured references/tables")
    errors = sum(len(s.error_cells) for s in wx.sheets)
    if errors:
        ops.append(f"{errors} cached error cell(s) — fix or remove broken logic")

    verdict = "Yes" if len(ops) >= 2 else "Possibly" if ops else "No"
    conf = 0.55 + 0.1 * min(3, len(ops))
    return Field.derived(
        {"verdict": verdict, "opportunities": ops or ["no obvious simplification signals"]},
        min(conf, 0.85), [f"{len(ops)} simplification signal(s)"],
    )


def _automation(wx: WorkbookXray, t: dict) -> Field:
    drivers: list[str] = []
    against: list[str] = []

    comp = t["formulas"] / t["distinct"] if t["distinct"] else 0
    if comp >= 5:
        drivers.append(f"systematic formulas ({comp:.0f}x reuse) — rules are regular and codifiable")
    elif 0 < comp < 2:
        against.append("many bespoke one-off formulas — logic is not uniform")
    if wx.connections or wx.external_links:
        drivers.append("existing data connections/links — source is already systemised")
    reads, _ = _dependency_maps(wx)
    if any(review.reconciliation(s, reads.get(s.name, set())) for s in wx.sheets):
        drivers.append("confirmed reconciliation structure — comparison rules may be suitable for controlled automation")
    if wx.has_vba:
        drivers.append("already partly automated via VBA/macros")

    verdict = "Candidate — workflow confirmation required" if drivers else "Not established"
    conf = 0.55 + 0.1 * min(3, len(drivers))
    return Field.derived(
        {"verdict": verdict, "drivers": drivers or ["no strong automation drivers"],
         "against": against,
         "workflow_confirmation": "Confirm the actual repeated step, source availability, frequency, exception handling and approval controls before estimating automation benefit."},
        min(conf, 0.85), [f"{len(drivers)} driver(s), {len(against)} counter-signal(s)"],
    )


def _retirement(wx: WorkbookXray) -> Field:
    signals: list[str] = []
    months = _months_since((wx.core_props or {}).get("modified") or wx.fs_modified)
    if months is not None and months > 12:
        signals.append(f"not modified in ~{months} months (stale)")
    name = wx.filename.lower()
    if any(k in name for k in _SUPERSEDED_NAMES):
        signals.append("filename suggests a superseded/backup copy")
    return Field(value={"verdict": "Not established — owner confirmation required",
                        "signals": signals,
                        "confirmation": "Confirm active usage, recipients, replacement coverage and retention requirements. Hidden sheets and cached errors are not retirement evidence."},
                 basis="needs_human", evidence=signals + ["file age and naming are context only, not proof of disuse"])


_BUSINESS_AREA = {
    "Calculation": "Calculation / modelling",
    "Reconciliation": "Reconciliation / control",
    "Reconciliation / Control": "Reconciliation / control",
    "Data Transformation": "Data preparation / transformation",
    "Reporting": "Reporting / MI",
    "Manual Input": "Manual data capture",
    "Other": "General-purpose / other",
}


def _business_area(logic_type: Field) -> Field:
    """Best-fit process classification, mapped from the detected logic type."""
    area = _BUSINESS_AREA.get(logic_type.value, "General-purpose / other")
    return Field.derived(area, logic_type.confidence or 0.5,
                         [f"mapped from logic type '{logic_type.value}'"])


# ------------------------------------------------------------------- entry


def assess_file(wx: WorkbookXray) -> FileAssessment:
    """Step 1: populate the file-level fields we can ground in the evidence."""
    t = _totals(wx)
    fa = FileAssessment()

    # Fact Assessment — extracted / derived
    fa.file_id = Field.extracted(wx.sha256[:12], ["content hash (stable across renames)"])
    fa.file_name = Field.extracted(wx.filename)
    fa.scan_status = Field.extracted(wx.parse_status)
    incomplete = [review.safe_message(x, wx.path) for x in wx.warnings]
    incomplete += [review.safe_message(n, wx.path) for s in wx.sheets for n in s.notes]
    if wx.parse_status == "full":
        scan_error = "N/A"
    elif incomplete:
        scan_error = "; ".join(dict.fromkeys(incomplete))
    elif wx.parse_status == "failed":
        scan_error = "Workbook scan failed; no technical error detail was retained."
    else:
        scan_error = "Workbook scan was partial; some content could not be read."
    fa.scan_error = Field.extracted(scan_error)
    fa.sheet_count_total = Field.extracted(len(wx.sheets))
    fa.sheet_count_hidden = Field.extracted(sum(s.state != "visible" for s in wx.sheets))
    fa.complexity = _complexity(wx, t)
    fa.key_inputs = _key_inputs(wx)
    fa.source_system = _source_system(wx)
    fa.euc_preparer = _euc_preparer(wx)

    # Fact Assessment — deferred (narrative fields filled by the assessor step)
    fa.purpose_of_file = Field.pending("needs_llm", "narrative from structure + headers")
    fa.key_output_outcome = Field.pending("needs_llm", "business outcome the model supports")
    fa.key_outputs = Field.pending("needs_llm", "name outputs from terminal/reporting tabs")
    fa.process = Field.pending("needs_llm", "broader end-to-end process supported by workbook evidence")
    fa.sub_process = Field.pending("needs_llm", "intermediate business activity supported by workbook evidence")
    fa.usage_frequency = Field.pending("needs_human", "How often is this process run: daily, monthly, quarterly or ad hoc?")
    fa.completion_timeline = Field.pending("needs_human", "What is the completion deadline relative to period-end, and how long does preparation take?")
    fa.output_recipient = Field.pending("needs_human", "Which team, person or downstream process receives the final deliverable?")

    # Key AI Finding / Observation — heuristic (Step 3); corpus ones deferred
    fa.potential_duplication = Field(value="Requires corpus analysis", basis="needs_corpus",
                                     evidence=["needs comparison with other workbooks"])
    fa.similar_duplicate_files = Field(value="Requires corpus analysis", basis="needs_corpus",
                                       evidence=["needs comparison with other workbooks"])
    fa.potential_consolidation = Field(value="Requires corpus analysis", basis="needs_corpus",
                                       evidence=["needs comparison with other workbooks"])
    fa.potential_simplification = _simplification(wx, t)
    fa.potential_automation = _automation(wx, t)
    fa.potential_retirement = _retirement(wx)

    # Workbook logic / Automation — extracted / derived
    fa.logic_type = _logic_type(wx, t)
    fa.business_area_process = _business_area(fa.logic_type)
    fa.key_calculations_logic = _key_calculations(wx, t)
    fa.manual_intervention = _manual_intervention(wx, t)
    fa.macros_vba_external_links = _macros_links(wx)
    fa.reconciliation_logic = _reconciliation_logic(wx, t)

    return fa


# ------------------------------------------------------- tab-level (Step 2)

def _sheet_ratio(s) -> float:
    return s.formula_profile.get("total", 0) / max(1, s.populated_cells)


def _dependency_maps(wx: WorkbookXray) -> tuple[dict, dict]:
    """Return (reads, read_by) over in-workbook cross-sheet references.

    ``reads[name]``   = set of sheets whose cells this tab's formulas reference.
    ``read_by[name]`` = set of sheets that reference this tab (its downstream).
    """
    names = {s.name for s in wx.sheets}
    canonical = {name.casefold(): name for name in names}
    reads: dict[str, set] = {}
    read_by: dict[str, set] = {n: set() for n in names}
    for s in wx.sheets:
        refs = {
            canonical[r.casefold()] for r in s.formula_profile.get("referenced_sheets", {})
            if r.casefold() in canonical and r.casefold() != s.name.casefold()
        }
        reads[s.name] = refs
        for r in refs:
            read_by.setdefault(r, set()).add(s.name)
    return reads, read_by


def _tab_category(s, wx, has_downstream: bool) -> Field:
    detected = review.roles(s, has_downstream)
    if detected == ["Uncertain"]:
        return Field(value="Uncertain", basis="needs_llm", confidence=0.3,
                     evidence=["No strong tab-role signal; confirm purpose with the owner"])
    return Field.derived(detected[0], 0.65 if len(detected) > 1 else 0.75,
                         ["Role candidates from labels, regions, formulas and dependencies",
                          "Detected roles: " + ", ".join(detected)])


def _tab_information(s, has_upstream: bool, has_downstream: bool) -> Field:
    """data input / intermediate workings / output / supporting."""
    ratio = _sheet_ratio(s)
    kinds = {r.kind for r in s.regions}
    if s.populated_cells and ratio < 0.1:
        val = "data input" if has_downstream else "supporting"
    elif has_downstream and (has_upstream or ratio > 0.2):
        val = "intermediate workings"
    elif ratio > 0.1 and not has_downstream:
        val = "output"
    elif kinds <= {"notes", "title", "key_value", "unknown"}:
        val = "supporting"
    else:
        val = "intermediate workings"
    ev = [f"formula ratio {ratio:.0%}",
          f"upstream={'yes' if has_upstream else 'no'}, "
          f"downstream={'yes' if has_downstream else 'no'}"]
    return Field.derived(val, 0.6, ev)


def _tab_key_calc(s) -> Field:
    fp = s.formula_profile
    return Field.derived(review.calculation(s), 0.65,
        [f"{fp.get('total', 0)} formulas, {fp.get('distinct_skeletons', 0)} distinct shapes; workload is not proof of business complexity"])


def _tab_upstream(s, wx, reads_set: set) -> Field:
    sources = s.formula_profile.get("external_sources", [])
    return Field.derived({
        "within_workbook": sorted(reads_set),
        "external_files": sorted({review.short_source(x) for x in sources}),
        "unresolved_external": bool(s.formula_profile.get("unresolved_external_indices") or
                                    (s.formula_profile.get("external_count") and not sources)),
        "scope": "Observed formula references only; connections, named ranges and dynamic references may be unresolved."
    }, 0.75, ["External paths shortened to filenames; unresolved targets are not assigned to every tab"])


def _tab_downstream(s, read_by_set: set) -> Field:
    in_wb = [f"sheet: {r}" for r in sorted(read_by_set)]
    return Field(
        value={"in_workbook": in_wb or ["none within this workbook"],
               "other_files": None},
        basis="derived" if in_wb else "extracted",
        confidence=1.0,
        evidence=["observed formula references; dynamic, named and VBA dependencies may be absent; cross-file consumers require additional evidence"],
    )


def _human_validation(s, wx) -> tuple[Field, Field]:
    _, read_by = _dependency_maps(wx)
    detected = review.roles(s, bool(read_by[s.name]))
    reasons = review.review_questions(s, detected[0], detected)
    return (
        Field.derived("Y" if reasons else "N", 0.7,
                      [f"{len(reasons)} targeted review question(s); no blanket flag for VBA or hidden status"]),
        Field.derived(reasons or ["No targeted question; still subject to normal review sampling."],
                      0.7, ["A clean tab is not proof of correctness"]),
    )


def assess_tab(s, wx, reads_set, read_by_set) -> TabAssessment:
    ta = TabAssessment()
    has_up = bool(reads_set)
    has_down = bool(read_by_set)
    ta.tab_name = Field.extracted(s.name,
                                  ["hidden sheet" for _ in [1] if s.state != "visible"])
    ta.tab_visibility = Field.extracted({"visible": "Visible", "hidden": "Hidden", "veryHidden": "Very hidden"}.get(s.state, s.state))
    ta.tab_category = _tab_category(s, wx, has_down)
    ta.tab_roles = Field.derived(review.roles(s, has_down), 0.65, ["A tab can contain multiple functional roles; primary role retained for comparisons"])
    ta.tab_information_analysis = _tab_information(s, has_up, has_down)
    category = ta.tab_category.value
    info = {"Output": "output", "Input": "data input", "Mapping": "supporting",
            "Calculation": "intermediate workings", "Control Check": "intermediate workings"}.get(category)
    if info:
        ta.tab_information_analysis = Field.derived(info, 0.65, ["Primary role and observed dependencies; a report can also feed downstream tabs"])
    ta.key_calculation_transformation_logic = _tab_key_calc(s)
    ta.upstream_dependencies = _tab_upstream(s, wx, reads_set)
    ta.downstream_dependencies = _tab_downstream(s, read_by_set)
    ta.human_validation_required, ta.validation_reason = _human_validation(s, wx)
    ta.tab_purpose_description = Field.pending(
        "needs_llm", "narrative purpose from headers + category + calculations")
    return ta


def _apply_narrative(a: Assessment, narr, basis: str, label: str) -> None:
    """Write an assessor's :class:`~excel_xray.narrative.Narrative` onto the
    narrative fields, tagging each with the assessor's basis."""
    ev = [f"{basis} by {label}"]

    def put(fld, value):
        if value is not None:
            fld.value = value
            fld.basis = basis
            fld.confidence = 0.7 if basis == "inferred" else 0.4
            fld.evidence = ev + (
                [] if basis == "inferred"
                else ["offline template — enable the LLM assessor for a considered answer"]
            )

    put(a.file.purpose_of_file, narr.purpose_of_file)
    put(a.file.key_output_outcome, narr.key_output_outcome)
    eligible_outputs = review.output_names(a.tabs)
    proposed_outputs = narr.key_outputs or []
    final_outputs = [
        item for item in proposed_outputs
        if any(name.casefold() in str(item).casefold() for name in eligible_outputs)
    ]
    if not final_outputs:
        final_outputs = sorted(eligible_outputs) or [
            "no final business deliverable established — owner confirmation required"
        ]
    put(a.file.key_outputs, final_outputs)
    a.file.key_outputs.evidence.append(
        "Limited to report/deliverable tabs; inputs, mappings, helpers and intermediate workings are excluded"
    )
    put(a.file.process, narr.process)
    put(a.file.sub_process, narr.sub_process)
    for fld in (a.file.process, a.file.sub_process):
        if fld.value is None:
            fld.value = "Not established — owner confirmation required"
            fld.basis = "needs_human"
            fld.confidence = None
            fld.evidence = ["The assessor did not establish this field from workbook evidence"]
    # Opportunity details are built before the narrative step. Fill their
    # process context now that the offline or model assessor has supplied it.
    for fld, key in ((a.file.potential_simplification, "details"),
                     (a.file.potential_automation, "details")):
        if isinstance(fld.value, dict):
            for item in fld.value.get(key, []):
                item["process"] = a.file.process.value
                item["sub_process"] = a.file.sub_process.value
    for ta in a.tabs:
        put(ta.tab_purpose_description, narr.tabs.get(ta.tab_name.value))


def _business_review(a, wx, reads, read_by):
    """Keep technical evidence but lead with business activities and uncertainty."""
    fa = a.file
    a.input_groups = review.input_groups(wx, a.tabs, read_by)
    fa.key_inputs = Field.derived(a.input_groups, 0.65,
        ["Purposes inferred from source labels; consumers listed only when observed; essentiality requires owner confirmation"])
    a.error_summary = review.error_summary(wx, a.tabs, read_by)
    a.hidden_groups = review.hidden_groups(wx, a.tabs, read_by)
    canonical_logic = {"Reconciliation": "Reconciliation / Control"}
    detected = [canonical_logic.get(fa.logic_type.value, fa.logic_type.value)] if fa.logic_type.value != "Other" else []
    role_logic = {"Calculation": "Calculation", "Output": "Reporting",
                  "Mapping": "Data Transformation", "Control Check": "Reconciliation / Control"}
    for t in a.tabs:
        for role in t.tab_roles.value:
            if role in role_logic:
                detected.append(role_logic[role])
    for s in wx.sheets:
        functions = dict(s.formula_profile.get("top_functions", []))
        if any(functions.get(f) for f in ("LEFT", "RIGHT", "MID", "TRIM", "TEXT", "XLOOKUP", "VLOOKUP")):
            detected.append("Data Transformation")
        if review._matches(s.name, r"manual entry|manual input|manual adjustment"):
            detected.append("Manual Input")
    if wx.has_power_query:
        detected.append("Data Transformation")
    detected = list(dict.fromkeys(detected)) or ["Other"]
    fa.logic_types = Field.derived(detected, 0.65, ["Multiple activities can coexist; name-based manual capture is a candidate, not an observed workflow"])
    fa.business_area_process = Field.derived(
        "; ".join(_BUSINESS_AREA.get(x, x) for x in detected), 0.65,
        ["Combined process candidates from tab roles and formula operations"])
    calculation_sheets = sorted((s for s in wx.sheets if s.formula_profile.get("total")),
                               key=lambda s: -s.formula_profile.get("total", 0))
    operations = [review.calculation(s)["summary"] for s in calculation_sheets]
    existing = {}
    existing["summary"] = "\n".join(operations[:5]) or "No worksheet calculations detected."
    if len(operations) > 5:
        existing["summary"] += f"\nShowing the five largest formula workloads; all {len(operations)} tab operations appear in Calculation steps."
    existing["steps"] = [{"tab": s.name, "inputs": sorted(reads[s.name]),
                           "operation": review.calculation(s)["summary"],
                           "potential_outputs": sorted(review.reachable(s.name, read_by) & review.output_names(a.tabs))}
                          for s in wx.sheets if s.formula_profile.get("total")]
    fa.key_calculations_logic = Field.derived(existing, 0.65,
        ["Business operation summaries inferred from formulas and labels; formula counts and patterns remain in the technical appendix"])
    opportunities, opportunity_details, automation_steps, automation_details = [], [], [], []
    unknown_process = "Not established — owner confirmation required"
    for s in wx.sheets:
        fp = s.formula_profile
        activity = review.calculation(s)["summary"]
        if fp.get("hardcoded_literal_count", 0) > 5:
            action = "Confirm which hardcoded embedded constants are business assumptions or overrides before centralising them."
            opportunities.append(f"{s.name}: {action} Activity: {activity}")
            opportunity_details.append({"process": unknown_process, "sub_process": unknown_process,
                "worksheets": [s.name], "observed_evidence": "Hardcoded numeric constants occur in formulas",
                "candidate_action": action, "confirmation_required": "Confirm ownership, meaning, approval and permitted override process."})
        if fp.get("volatile_count"):
            action = "Review recalculation-dependent logic for a more stable design."
            opportunities.append(f"{s.name}: {action} Activity: {activity}")
            opportunity_details.append({"process": unknown_process, "sub_process": unknown_process,
                "worksheets": [s.name], "observed_evidence": "Volatile functions are present",
                "candidate_action": action, "confirmation_required": "Confirm intended recalculation behaviour and control ownership."})
        count, distinct = fp.get("total", 0), fp.get("distinct_skeletons", 0)
        if count > 50 and distinct and count / distinct < 3:
            action = "Review one-off calculation rules for standardisation."
            opportunities.append(f"{s.name}: {action} Activity: {activity}")
            opportunity_details.append({"process": unknown_process, "sub_process": unknown_process,
                "worksheets": [s.name], "observed_evidence": "Low formula-pattern reuse",
                "candidate_action": action, "confirmation_required": "Confirm whether the variants represent valid business exceptions."})
        if count and distinct and count / distinct >= 5:
            action = "Candidate — workflow confirmation required: assess repeated calculation rules for controlled automation."
            automation_steps.append(f"{s.name}: {action} Inputs: {', '.join(sorted(reads[s.name])) or 'local inputs (confirm source)'}. Activity: {activity}")
            automation_details.append({"process": unknown_process, "sub_process": unknown_process,
                "worksheets": [s.name], "observed_evidence": "Repeated formula patterns",
                "candidate_action": action,
                "confirmation_required": "Confirm the actual repeated manual step, source availability, frequency, exceptions, approvals and controls."})
    for source in a.input_groups:
        if source["source_type"] not in {"External workbooks", "Formal data connections"}:
            continue
        if not source["consumers"]:
            continue
        action = "Candidate — workflow confirmation required: assess controlled source ingestion and refresh."
        automation_steps.append(
            f"{', '.join(source['consumers'])}: {action} Observed source: {source['source']}."
        )
        automation_details.append({
            "process": unknown_process, "sub_process": unknown_process,
            "worksheets": source["consumers"],
            "observed_evidence": f"Observed {source['source_type'].lower()} dependency: {source['source']}",
            "candidate_action": action,
            "confirmation_required": "Confirm the actual repeated manual step, source availability, refresh frequency, exceptions, approvals and controls.",
        })
    for rec in fa.reconciliation_logic.value.get("reconciliations", []):
        action = "Candidate — workflow confirmation required: assess controlled matching and exception workflow automation."
        automation_steps.append(f"{rec['worksheet']}: {action}")
        automation_details.append({
            "process": unknown_process, "sub_process": unknown_process,
            "worksheets": [rec["worksheet"]],
            "observed_evidence": rec["overview"],
            "candidate_action": action,
            "confirmation_required": "Confirm matching rules, tolerance, exception ownership, frequency, approvals and controls.",
        })
    fa.potential_simplification = Field.derived(
        {"verdict": "Candidate — workflow confirmation required" if opportunities else "No supported candidate identified",
         "opportunities": opportunities, "details": opportunity_details}, 0.65,
        ["Opportunities tied to named activities; cached errors are in Error summary, not treated as retirement evidence"])
    au = fa.potential_automation.value
    au["steps"] = automation_steps
    au["details"] = automation_details
    if not automation_steps:
        au["steps"] = ["No repeated manual step established. Ask the owner to describe source extraction, copying, adjustments and approvals."]


def assess(wx: WorkbookXray, assessor=None) -> Assessment:
    """Full assessment: deterministic fields (Steps 1-3) plus the narrative
    fields (Step 4) filled by ``assessor`` — the offline stub by default, so the
    pipeline stays network-free unless a model-backed assessor is passed."""
    from .narrative import OfflineAssessor, build_bundle

    reads, read_by = _dependency_maps(wx)
    tabs = [
        assess_tab(s, wx, reads.get(s.name, set()), read_by.get(s.name, set()))
        for s in sorted(wx.sheets, key=lambda s: (s.state != "visible", s.position))
    ]
    a = Assessment(file=assess_file(wx), tabs=tabs)
    _business_review(a, wx, reads, read_by)

    assessor = assessor or OfflineAssessor()
    narr = assessor.narrate(build_bundle(a, wx))
    _apply_narrative(a, narr, assessor.basis, assessor.label)
    return a


def to_dict(a: Assessment) -> dict:
    return asdict(a)
