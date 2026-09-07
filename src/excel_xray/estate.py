"""Estate comparison (Step 7): compare EUCs across a folder.

Step 5 flagged look-alikes from formula shapes + headers with one blended score.
This goes further: it fingerprints each workbook on four independent signals —

* **formula shapes**  (normalised R1C1 skeletons): the calculation *method*,
* **input signature** (input-tab headers, source systems, connections, links):
  what it *consumes*,
* **output signature** (output-tab headers): the *deliverable* it produces,
* **dependency topology** (tab-role counts + cross-sheet edge count): the
  *pipeline shape*,

then classifies each pair's *relationship* rather than emitting a single number,
because the signals answer different questions:

* **Duplicate** — same method, inputs and outputs.
* **Same output, different method** — same deliverable, different formulas →
  consolidate.
* **Overlapping logic** — shared calculation shapes → extract a reusable piece.
* **Shared source** — same inputs, different work → common data lineage.

Finally it clusters the estate (connected components over linked pairs) so a
review sees families of related EUCs, not just a list of pairs.
"""

from __future__ import annotations

import math
from collections import Counter
from dataclasses import dataclass, field

from .assessment import _dependency_maps
from .scan import WorkbookXray

# Relationship thresholds (per-signal, all in [0, 1]).
SK_DUP, OUT_DUP, IN_DUP = 0.80, 0.60, 0.50
OUT_SAME = 0.60
SK_OVERLAP = 0.40
IN_SHARED = 0.60
LINK_OVERALL = 0.35  # a pair at/above this overall is "linked" for clustering

REL_COLOR = {
    "Duplicate": "#C0392B",
    "Same output, different method": "#B4531F",
    "Overlapping logic": "#2D6CA2",
    "Shared source": "#6B4E9E",
    "Related": "#7A8290",
    "Unrelated": "#C9CFD8",
}
_TOPO_KEYS = ("Input", "Calculation", "Mapping", "Control Check",
              "Validation", "Output", "edges")


@dataclass
class EucFingerprint:
    file_id: str
    file_name: str
    logic_type: str
    skeletons: set = field(default_factory=set)
    input_sig: set = field(default_factory=set)
    output_sig: set = field(default_factory=set)
    topology: dict = field(default_factory=dict)
    headers: set = field(default_factory=set)
    functions: set = field(default_factory=set)


@dataclass
class EstateResult:
    fingerprints: list
    pairs: list        # (i, j, comparison dict) for linked pairs, strongest first
    clusters: list     # list of index lists (size >= 2), largest first
    singletons: list   # indices with no links


def _norm_headers(headers) -> set:
    out: set = set()
    for h in headers:
        for part in (h or "").split("::"):
            p = part.strip().lower()
            if len(p) >= 2 and not p.isdigit():
                out.add(p)
    return out


def fingerprint_euc(wx: WorkbookXray, assessment) -> EucFingerprint:
    """Build the four-signal fingerprint from a workbook + its assessment."""
    sheet_headers = {
        s.name: [h for r in s.regions for h in r.headers if h] for s in wx.sheets
    }
    skeletons: set = set()
    functions: set = set()
    all_headers: set = set()
    for s in wx.sheets:
        for sk, _ in s.formula_profile.get("top_skeletons", []):
            skeletons.add(sk)
        for fn, _ in s.formula_profile.get("top_functions", []):
            functions.add(fn)
        all_headers |= _norm_headers(sheet_headers[s.name])

    input_sig: set = set()
    output_sig: set = set()
    cat_counts: Counter = Counter()
    for ta in assessment.tabs:
        cat = ta.tab_category.value
        info = ta.tab_information_analysis.value
        cat_counts[cat] += 1
        hs = _norm_headers(sheet_headers.get(ta.tab_name.value, []))
        if cat == "Input" or info == "data input":
            input_sig |= hs
        if cat == "Output" or info == "output":
            output_sig |= hs

    src = assessment.file.source_system.value
    if isinstance(src, list):
        for s in src:
            input_sig.add("src:" + str(s).lower()[:60])
    for x in wx.external_links:
        input_sig.add("ext:" + str(x).lower())
    for c in wx.connections:
        nm = c.get("name") or c.get("type")
        if nm:
            input_sig.add("conn:" + str(nm).lower())

    reads, _ = _dependency_maps(wx)
    topology = {k: cat_counts.get(k, 0) for k in _TOPO_KEYS if k != "edges"}
    topology["edges"] = sum(len(v) for v in reads.values())

    return EucFingerprint(
        file_id=assessment.file.file_id.value,
        file_name=wx.filename,
        logic_type=assessment.file.logic_type.value,
        skeletons=skeletons, input_sig=input_sig, output_sig=output_sig,
        topology=topology, headers=all_headers, functions=functions,
    )


def _jaccard(a: set, b: set) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def _cosine(a: dict, b: dict) -> float:
    keys = set(a) | set(b)
    dot = sum(a.get(k, 0) * b.get(k, 0) for k in keys)
    na = math.sqrt(sum(v * v for v in a.values()))
    nb = math.sqrt(sum(v * v for v in b.values()))
    return dot / (na * nb) if na and nb else 0.0


def _relationship(sk: float, inp: float, out: float, overall: float) -> str:
    if sk >= SK_DUP and out >= OUT_DUP and inp >= IN_DUP:
        return "Duplicate"
    if out >= OUT_SAME and sk < 0.5:
        return "Same output, different method"
    if sk >= SK_OVERLAP:
        return "Overlapping logic"
    if inp >= IN_SHARED and out < 0.4 and sk < 0.4:
        return "Shared source"
    if overall >= LINK_OVERALL:
        return "Related"
    return "Unrelated"


def compare(a: EucFingerprint, b: EucFingerprint) -> dict:
    """Per-signal similarity + a typed relationship for one pair."""
    sk = _jaccard(a.skeletons, b.skeletons)
    if not a.skeletons and not b.skeletons:  # data workbooks: lean on headers
        sk = _jaccard(a.headers, b.headers)
    inp = _jaccard(a.input_sig, b.input_sig)
    out = _jaccard(a.output_sig, b.output_sig)
    topo = _cosine(a.topology, b.topology)
    overall = round(0.45 * sk + 0.25 * out + 0.20 * inp + 0.10 * topo, 3)
    rel = _relationship(sk, inp, out, overall)
    return {
        "skeleton": round(sk, 3), "input": round(inp, 3), "output": round(out, 3),
        "topology": round(topo, 3), "overall": overall,
        "relationship": rel, "linked": rel != "Unrelated",
    }


def _clusters(n: int, linked_pairs: list) -> list:
    parent = list(range(n))

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for i, j in linked_pairs:
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[rj] = ri

    groups: dict = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(i)
    return list(groups.values())


def build_estate(named_assessments: list) -> EstateResult:
    """Compare every pair in ``[(wx, assessment), ...]`` and cluster the estate."""
    fps = [fingerprint_euc(wx, a) for wx, a in named_assessments]
    n = len(fps)
    pairs = []
    linked = []
    for i in range(n):
        for j in range(i + 1, n):
            c = compare(fps[i], fps[j])
            if c["linked"]:
                pairs.append((i, j, c))
                linked.append((i, j))
    pairs.sort(key=lambda p: p[2]["overall"], reverse=True)

    groups = _clusters(n, linked)
    clusters = sorted((g for g in groups if len(g) > 1), key=len, reverse=True)
    singletons = [g[0] for g in groups if len(g) == 1]
    return EstateResult(fingerprints=fps, pairs=pairs,
                        clusters=clusters, singletons=singletons)
