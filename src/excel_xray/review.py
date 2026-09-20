"""Business-facing interpretations of structural evidence, not usage claims.

Rules use names, detected headers, formula profiles and observed sheet edges.
Names are hints, and dependency reachability indicates potential impact only.
"""

from collections import defaultdict
import re
from urllib.parse import unquote, urlsplit

from .util import range_boundaries


def short_source(value):
    """Display filenames for Windows, POSIX and URL paths without query strings."""
    text = unquote(str(value)).replace("\\", "/")
    if "://" in text:
        text = urlsplit(text).path
    return text.rstrip("/").rsplit("/", 1)[-1].split("?", 1)[0].split("#", 1)[0]


def safe_message(value, source_path=None):
    """Remove machine-local paths and full URLs from client-facing text."""
    text = str(value)
    if source_path:
        text = text.replace(str(source_path), short_source(source_path))
    text = re.sub(r"https?://[^\s;,]+", lambda m: short_source(m.group(0)), text)
    path_pattern = r"(?<!\w)(?:[A-Za-z]:[\\/]|/Users/|/home/|/Volumes/|/mnt/|\\\\)[^\s;,]+"
    return re.sub(path_pattern, lambda m: short_source(m.group(0)), text)


def labels(s):
    return " ".join([s.name] + [h for r in s.regions for h in r.headers if h]).lower()


def _matches(text, pattern):
    return bool(re.search(pattern, text.lower().replace("_", " ")))


def roles(s, has_downstream):
    name = s.name.lower().replace("_", " ")
    fp = s.formula_profile
    total = fp.get("total", 0)
    ratio = total / max(1, s.populated_cells)
    output = _matches(name, r"\b(output|report|summary|dashboard|results?|mi|pack)\b|cash flow|balance sheet|profit.*loss|financial statement|fs[ -]*form")
    mapping = _matches(name, r"\b(coa|map|mapping|lookup|reference|xref)\b|chart of accounts")
    input_named = _matches(name, r"\b(input|inputs|data|raw|extract|source|feed|import|assumptions?|tb)\b|trial balance")
    control = _matches(name, r"\b(check|control|recon|reconciliation|tie.out|proof|variance)\b")
    validation = _matches(name, r"\b(validation|review|qa|signoff|sign.off|approval)\b")
    working = _matches(name, r"\b(calc|calculation|working|workings|model|engine|compute)\b")
    found = []
    # Explicit source/mapping labels avoid calling a static COA an output.
    if mapping:
        found.append("Mapping")
    elif input_named:
        found.append("Input")
    elif output:
        found.append("Output")
    if control:
        found.append("Control Check")
    if validation:
        found.append("Validation")
    if total and (working or ratio > 0.4 or any(r.kind == "calculation" for r in s.regions)):
        found.append("Calculation")
    if not found and s.populated_cells > 10 and ratio < 0.1 and has_downstream:
        found.append("Input")
    if not found and total and not has_downstream and s.state == "visible" and fp.get("cross_sheet_count"):
        found.append("Output")
    if not found and total:
        found.append("Calculation")
    return list(dict.fromkeys(found)) or ["Uncertain"]


def calculation(s):
    fp = s.formula_profile
    total, distinct = fp.get("total", 0), fp.get("distinct_skeletons", 0)
    fns = dict(fp.get("top_functions", []))
    actions = []
    if any(fns.get(f) for f in ("SUM", "SUMIF", "SUMIFS", "SUBTOTAL", "AVERAGE")):
        actions.append("aggregate amounts into totals or summaries")
    if any(fns.get(f) for f in ("VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH", "LOOKUP")):
        actions.append("match or retrieve records using lookup rules")
    if any(fns.get(f) for f in ("LEFT", "RIGHT", "MID", "TRIM", "SUBSTITUTE", "TEXT", "CONCAT")):
        actions.append("standardise or reshape text fields")
    if any(fns.get(f) for f in ("IF", "IFS", "IFERROR", "IFNA", "ABS")):
        actions.append("apply conditions or flag exceptions")
    if total and not actions:
        actions.append("calculate amounts or carry referenced values forward")
    heavy = total > 5000 or distinct > 60
    complexity = "Heavy" if heavy else "Moderate" if total > 200 or distinct > 15 else "Light"
    headers = list(dict.fromkeys(h for r in s.regions for h in r.headers if h))[:4]
    summary = (f"{complexity} formula workload on {s.name}: " + "; ".join(actions) + ".") if total else f"{s.name}: stores data or presentation content; no worksheet formulas detected."
    if headers:
        summary += " Covers " + ", ".join(headers) + "."
    return {"summary": summary, "workload": complexity if total else "No formulas",
            "top_functions": [f"{fn}×{n}" for fn, n in fp.get("top_functions", [])[:8]],
            "top_formula_shapes": [f"{sk} (×{n})" for sk, n in fp.get("top_skeletons", [])[:6]]}


def reachable(start, read_by):
    seen, queue = {start}, [start]
    while queue:
        for name in sorted(read_by.get(queue.pop(), ())):
            if name not in seen:
                seen.add(name)
                queue.append(name)
    return seen


def output_names(tabs):
    """Return conservative final-deliverable candidates.

    A terminal formula tab is not enough.  The primary role must be Output and
    its name must carry reporting/deliverable evidence; generic working,
    calculation and helper tabs stay in tab-level detail only.
    """
    output_pattern = r"\b(output|report|summary|dashboard|results?|mi|pack)\b|cash flow|balance sheet|profit.*loss|financial statement|fs[ -]*form"
    excluded = r"\b(working|workings|calc|calculation|map|mapping|lookup|reference|helper|support|staging|raw|input|data)\b"
    return {
        t.tab_name.value for t in tabs
        if t.tab_category.value == "Output"
        and _matches(t.tab_name.value, output_pattern)
        and not _matches(t.tab_name.value, excluded)
    }


def business_purpose(text):
    for pattern, purpose in [
        (r"trial balance|\btb\b", "Trial balance"),
        (r"claim", "Claims data"),
        (r"exchange|\bfx\b|currency rate", "Exchange rates"),
        (r"prior|previous|last year", "Prior-period data"),
        (r"adjustment|journal|manual booking", "Adjustments / journals"),
        (r"\bcoa\b|chart of accounts|mapping", "Account mapping"),
        (r"premium|policy", "Policy / premium data"),
        (r"assumption|parameter", "Assumptions / parameters"),
        (r"tax", "Tax schedules"),
        (r"reserve|reserving|ibnr", "Reserve calculations"),
        (r"cash flow|financial statement|balance sheet|fs[ _-]*form", "Financial reporting"),
    ]:
        if _matches(text, pattern):
            return purpose
    return "Other input — purpose to confirm"


def input_groups(wx, tabs, read_by):
    groups = []
    by_name = {s.name: s for s in wx.sheets}
    for t in tabs:
        if t.tab_category.value not in ("Input", "Mapping"):
            continue
        name = t.tab_name.value
        groups.append({"purpose": business_purpose(labels(by_name[name])),
                       "source": name, "source_type": "In-workbook tabs",
                       "consumers": sorted(read_by.get(name, ())),
                       "association": "Observed worksheet references" if read_by.get(name) else "No direct formula consumer observed",
                       "observed": bool(read_by.get(name)),
                       "reference_count": len(read_by.get(name, ())),
                       "essential": "Essentiality requires owner confirmation"})

    # OOXML can list the same external workbook more than once and may encode
    # spaces or retain a machine-local path.  Group on the safe display name;
    # keep only the count and observed consumers in the client-facing model.
    external = defaultdict(list)
    for path in wx.external_links:
        external[short_source(path).casefold()].append(path)
    for raw_paths in external.values():
        source = short_source(raw_paths[0]) or "Unresolved external workbook"
        raw_keys = {str(x) for x in raw_paths}
        consumers = sorted({
            s.name for s in wx.sheets
            if raw_keys & set(map(str, s.formula_profile.get("external_sources", [])))
            or source.casefold() in {
                short_source(x).casefold() for x in s.formula_profile.get("external_sources", [])
            }
        })
        formula_refs = sum(
            s.formula_profile.get("external_count", 0) for s in wx.sheets
            if s.name in consumers
        )
        groups.append({"purpose": business_purpose(source), "source": source,
                       "source_type": "External workbooks", "consumers": consumers,
                       "association": "Directly observed indexed external reference" if consumers else "Consumer worksheet unresolved",
                       "observed": bool(consumers),
                       "reference_count": max(len(raw_paths), formula_refs),
                       "essential": "Essentiality requires owner confirmation"})
    for c in wx.connections:
        source = short_source(c.get("name") or c.get("description") or "Unnamed connection")
        groups.append({"purpose": business_purpose(source), "source": source,
                       "source_type": "Formal data connections", "consumers": [],
                       "association": "Directly observed in Excel connection metadata; consumer worksheet unresolved",
                       "observed": True, "reference_count": 1,
                       "essential": "Essentiality and business use require owner confirmation"})
    for source in wx.pivot_cache_sources:
        display = short_source(source) if ("/" in str(source) or "\\" in str(source) or "://" in str(source)) else unquote(str(source)).split("?", 1)[0].split("#", 1)[0]
        groups.append({"purpose": business_purpose(source),
                       "source": display, "source_type": "Pivot sources",
                       "consumers": [], "association": "Directly observed in pivot metadata; consumer worksheet unresolved",
                       "observed": True, "reference_count": 1,
                       "essential": "Essentiality and refresh ownership require owner confirmation"})
    for s in wx.sheets:
        unresolved = s.formula_profile.get("unresolved_external_indices", [])
        if unresolved or (s.formula_profile.get("external_count") and
                          not s.formula_profile.get("external_sources")):
            groups.append({
                "purpose": "Other input — purpose to confirm",
                "source": "Unresolved external source",
                "source_type": "Unresolved sources",
                "consumers": [s.name],
                "association": "External reference observed, but the source target could not be resolved",
                "observed": False,
                "reference_count": max(len(unresolved), s.formula_profile.get("external_count", 0), 1),
                "essential": "Source identity and essentiality require owner confirmation",
            })

    # Deduplicate identical safe sources while preserving consumer evidence and
    # reference counts. This also prevents repeated links from bloating output.
    merged = {}
    for item in groups:
        key = (item["source_type"], item["source"].casefold())
        if key not in merged:
            merged[key] = item
            continue
        current = merged[key]
        current["consumers"] = sorted(set(current["consumers"]) | set(item["consumers"]))
        current["reference_count"] += item["reference_count"]
        current["observed"] = current["observed"] or item["observed"]
    return sorted(merged.values(), key=lambda x: (x["source_type"], x["purpose"], x["source"]))


def error_summary(wx, tabs, read_by):
    cats = {t.tab_name.value: t.tab_category.value for t in tabs}
    outputs = output_names(tabs)
    result = []
    for s in sorted(wx.sheets, key=lambda x: (x.state != "visible", x.position)):
        grouped = defaultdict(list)
        for error in s.error_cells:
            location, _, kind = error.rpartition(" ")
            coord = location.rsplit("!", 1)[-1]
            area = "Final-output candidate" if s.name in outputs else "Input" if cats[s.name] in ("Input", "Mapping") else "Working / supporting"
            region = "Unresolved region"
            try:
                col, row, _, _ = range_boundaries(coord)
                matches = [r for r in s.regions if r.top <= row <= r.bottom and r.left <= col <= r.right]
                if matches:
                    region = min(matches, key=lambda r: (r.bottom-r.top+1)*(r.right-r.left+1)).ref
            except ValueError:
                pass
            grouped[(kind or "Unknown error", area, region)].append(coord)
        potential = sorted(reachable(s.name, read_by) & outputs)
        for (kind, area, region), cells in sorted(grouped.items()):
            result.append({"tab": s.name, "type": kind, "count": len(cells),
                           "area": area, "region": region, "cells": cells,
                           "potential_outputs": potential,
                           "impact": "Potential sheet-level dependency; cell-level impact not proven" if potential else "No output path detected; business impact remains unconfirmed",
                           "recalculation": "Recalculate in Excel with source links refreshed, then confirm whether the error persists and affects the deliverable."})
    return result


def hidden_groups(wx, tabs, read_by):
    by_name = {t.tab_name.value: t for t in tabs}
    outputs = output_names(tabs)
    groups = defaultdict(list)
    for s in wx.sheets:
        if s.state != "visible":
            role = by_name[s.name].tab_category.value
            purpose = business_purpose(labels(s))
            groups[(purpose, role)].append(s.name)
    return [{"purpose": role if purpose.startswith("Other input") else purpose + " / " + role,
             "count": len(names), "tabs": names,
             "supports": sorted(set().union(*(reachable(n, read_by) & outputs for n in names))),
             "explanation": {"Calculation": "Hidden calculation workings", "Input": "Hidden source data or assumptions",
                             "Mapping": "Hidden reference or mapping tables", "Output": "Hidden report or deliverable candidates",
                             "Control Check": "Hidden reconciliation or control checks", "Validation": "Hidden review/support checks"}.get(role, "Purpose requires owner confirmation"),
             "basis": "derived"} for (purpose, role), names in sorted(groups.items())]


def review_questions(s, category, roles_value):
    """Prioritise consequential uncertainty, not every technical observation."""
    fp = s.formula_profile
    questions = []
    if s.error_cells:
        kinds = sorted({e.rsplit(" ", 1)[-1] for e in s.error_cells})
        questions.append(f"After refreshing sources and recalculating {s.name}, do the {', '.join(kinds)} errors remain, and which business deliverables change?")
    if fp.get("external_count"):
        questions.append(f"Are the external inputs used by {s.name} current, complete and approved for this reporting period?")
    if fp.get("volatile_count"):
        questions.append(f"Does recalculating {s.name} change the reporting date, selected records or reported amounts as intended?")
    if fp.get("hardcoded_literal_count", 0) >= 5 and category in ("Calculation", "Output", "Control Check"):
        questions.append(f"Are the fixed numbers in {s.name}'s formulas approved assumptions or ordinary calculation constants, and who maintains them?")
    if category == "Uncertain" or len(roles_value) > 1:
        questions.append(f"Which parts of {s.name} are final business deliverables, and which are inputs or intermediate workings?")
    elif category == "Output" and not _matches(s.name, r"report|output|summary|dashboard|pack|cash flow|balance sheet|financial|fs[ _-]*form"):
        questions.append(f"Is {s.name} a business deliverable or an intermediate working paper?")
    # Layout uncertainty alone does not trigger review unless it dominates the sheet.
    low = sum(r.populated_cells for r in s.regions if r.detect_confidence < 0.7)
    if low > max(20, s.populated_cells * 0.5):
        questions.append(f"Are the main input, calculation and output blocks on {s.name} identified correctly?")
    return questions


def reconciliation(s, reads):
    signal = _matches(labels(s), r"\brecon|variance|difference|unmatched|tie.out|tolerance|exception")
    # Confirmation needs two observed comparison sources as well as language
    # that shows agreement/difference testing. Function names alone are not
    # reconciliation evidence.
    if not signal or len(reads) < 2:
        return None
    headers = list(dict.fromkeys(h for r in s.regions for h in r.headers if h))
    keys = [h for h in headers if _matches(h, r"\bid\b|account|policy|claim|reference|code")]
    sources = sorted(reads)
    tolerance_headers = [h for h in headers if _matches(h, r"tolerance|threshold")]
    exception_headers = [h for h in headers if _matches(h, r"variance|difference|exception|unmatched")]
    matching = ("Candidate matching field(s) from headers: " + ", ".join(keys[:8])) if keys else "Not established"
    tolerance = ("Tolerance field observed: " + ", ".join(tolerance_headers[:4])) if tolerance_headers else "Not established"
    exception = ("Difference/exception field(s) observed: " + ", ".join(exception_headers[:4])) if exception_headers else "Not established"
    evidence = [
        f"{s.name} directly references comparison sources: {', '.join(sources)}",
        "Worksheet name or headers contain reconciliation/difference evidence",
    ]
    return {
        "overview": f"{sources[0]} compared with {sources[1]} on {s.name}",
        "source_a": sources[0],
        "source_b": sources[1],
        "matching_criteria": matching,
        "tolerance": tolerance,
        "exception_logic": exception,
        "evidence": evidence,
        "confidence": 0.78 if keys else 0.68,
        "worksheet": s.name,
    }


def reconciliation_candidate(s, reads):
    """Return ambiguous comparison signals without counting a reconciliation."""
    fp = s.formula_profile
    functions = dict(fp.get("top_functions", []))
    lookups = [f for f in ("VLOOKUP", "HLOOKUP", "XLOOKUP", "INDEX", "MATCH", "LOOKUP")
               if functions.get(f)]
    label_signal = _matches(labels(s), r"\brecon|variance|difference|unmatched|tie.out|tolerance|exception")
    if not (lookups or label_signal):
        return None
    observed = []
    if lookups:
        observed.append("lookup/match functions: " + ", ".join(lookups))
    if label_signal:
        observed.append("reconciliation, variance or exception label")
    if reads:
        observed.append("observed worksheet source(s): " + ", ".join(sorted(reads)))
    return {
        "worksheet": s.name,
        "observed_signal": "; ".join(observed),
        "confirmation_required": "Confirm both comparison sources, the matching key, agreement test, tolerance and exception handling.",
    }
