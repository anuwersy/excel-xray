"""Landscape layer (Step 7): a portfolio view across a folder of workbooks.

The corpus layer (:mod:`excel_xray.corpus`) answers, per file, *"is this a
duplicate?"*. This layer rolls those pairwise comparisons up into a **landscape**:
it clusters look-alike workbooks and gives each file one portfolio recommendation
— **Keep**, **Consolidate** or **Remove** — carrying a similarity score and the
evidence behind it.

It only *informs*. Nothing here edits, moves, renames or deletes a workbook.

Honesty carries over from the assessment layer. Whether a file is actually *used*
is not knowable from the file (that stays ``needs_human``); this layer's Remove
verdicts rest on **file-intrinsic** signals — a near-duplicate of another
workbook, or retirement signals (stale, superseded name, broken logic) — and
every verdict says so and asks for business confirmation before anything is
acted on.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import asdict, dataclass, field

from .corpus import (
    DUP_THRESHOLD,
    SIMILAR_THRESHOLD,
    SKELETON_DUP,
    Fingerprint,
    fingerprint,
    similarity,
)

# Recommendation verdicts.
KEEP = "Keep"
CONSOLIDATE = "Consolidate"
REMOVE = "Remove"

# Filename fragments that mark a superseded/working copy (mirrors the assessment
# layer's retirement heuristic).
_SUPERSEDED_NAMES = (
    "old", "backup", "copy", " v1", "v1.", "draft", "archive", "tmp", "temp", "bak",
)

_COMPLEXITY_RANK = {"High": 3, "Medium": 2, "Low": 1}


# --------------------------------------------------------------- schema shapes


@dataclass
class LandscapeEntry:
    """One workbook's place in the landscape and its portfolio recommendation."""

    file_id: str
    file_name: str
    path: str
    logic_type: str
    complexity: str
    recommendation: str
    confidence: float
    rationale: list[str] = field(default_factory=list)
    cluster_id: int | None = None
    is_cluster_primary: bool = False
    top_match: dict | None = None
    retirement_signals: list[str] = field(default_factory=list)


@dataclass
class Cluster:
    """A group of workbooks that look alike (connected by similarity edges)."""

    cluster_id: int
    member_ids: list[str]
    member_names: list[str]
    primary_id: str
    primary_name: str
    shared_logic_type: str | None
    cohesion: float          # mean pairwise overall similarity within the cluster
    has_duplicates: bool     # at least one pair is a near-duplicate


@dataclass
class Landscape:
    entries: list[LandscapeEntry] = field(default_factory=list)
    clusters: list[Cluster] = field(default_factory=list)
    summary: dict = field(default_factory=dict)


# ------------------------------------------------------ per-file metadata


@dataclass
class _Meta:
    """The file-intrinsic facts the ranking and Remove logic need, pulled from
    the scan + assessment so the core algorithm stays unit-testable."""

    path: str = ""
    modified: _dt.datetime = _dt.datetime.min
    cells: int = 0
    complexity: str = "Low"
    complexity_rank: int = 1
    superseded: bool = False
    retirement_signals: list[str] = field(default_factory=list)


def _parse_dt(iso) -> _dt.datetime:
    if not iso:
        return _dt.datetime.min
    try:
        s = str(iso).replace("Z", "+00:00")
        when = _dt.datetime.fromisoformat(s)
        if when.tzinfo is not None:
            when = when.replace(tzinfo=None)
        return when
    except (ValueError, TypeError):
        return _dt.datetime.min


def _meta_from(wx, assessment) -> _Meta:
    name = wx.filename.lower()
    complexity = assessment.file.complexity.value or "Low"
    ret = assessment.file.potential_retirement.value
    signals = ret.get("signals", []) if isinstance(ret, dict) else []
    return _Meta(
        path=wx.path,
        modified=_parse_dt((wx.core_props or {}).get("modified") or wx.fs_modified),
        cells=sum(s.populated_cells for s in wx.sheets),
        complexity=complexity,
        complexity_rank=_COMPLEXITY_RANK.get(complexity, 1),
        superseded=any(k in name for k in _SUPERSEDED_NAMES),
        retirement_signals=list(signals),
    )


# ------------------------------------------------------------- core algorithm


def _is_dup(sim: dict) -> bool:
    """A pair close enough to treat as materially the same workbook."""
    return sim["overall"] >= DUP_THRESHOLD or sim["skeleton"] >= SKELETON_DUP


def _components(n: int, edges: list[tuple[int, int]]) -> list[list[int]]:
    """Connected components over ``n`` nodes given similarity edges (union-find).

    Singletons are returned as their own one-element component, preserving order.
    """
    parent = list(range(n))

    def find(x: int) -> int:
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for a, b in edges:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[max(ra, rb)] = min(ra, rb)

    groups: dict[int, list[int]] = {}
    for i in range(n):
        groups.setdefault(find(i), []).append(i)
    return [groups[r] for r in sorted(groups)]


def _pick_primary(members: list[int], metas: list[_Meta], fps: list[Fingerprint]) -> int:
    """The 'keep' representative of a cluster: freshest, then largest, then most
    complex, then a non-superseded name, then a stable id tiebreak."""
    def key(i: int):
        m = metas[i]
        return (
            m.modified,               # most recently modified wins
            m.cells,                  # then the most complete
            m.complexity_rank,        # then the richest logic
            0 if m.superseded else 1,  # a superseded-looking name loses ties
            fps[i].file_id,           # deterministic final tiebreak
        )
    return max(members, key=key)


def _confidence(base: float, bump: float = 0.0) -> float:
    return round(min(0.85, base + bump), 2)


def build_landscape(fps: list[Fingerprint], metas: list[_Meta]) -> Landscape:
    """Cluster the fingerprints and derive one recommendation per file.

    Pure over its inputs (fingerprints + per-file metadata) so it can be tested
    without a live workbook.
    """
    n = len(fps)
    # Pairwise similarity, computed once.
    sims: dict[tuple[int, int], dict] = {}
    edges: list[tuple[int, int]] = []
    for i in range(n):
        for j in range(i + 1, n):
            s = similarity(fps[i], fps[j])
            sims[(i, j)] = s
            if s["overall"] >= SIMILAR_THRESHOLD or s["skeleton"] >= SKELETON_DUP:
                edges.append((i, j))

    def sim(i: int, j: int) -> dict:
        return sims[(i, j)] if i < j else sims[(j, i)]

    comps = _components(n, edges)

    # Assign cluster ids only to genuine (>1 member) clusters, in a stable order.
    cluster_id_of: dict[int, int] = {}
    clusters: list[Cluster] = []
    primary_of: dict[int, int] = {}
    for comp in comps:
        if len(comp) < 2:
            continue
        cid = len(clusters) + 1
        primary = _pick_primary(comp, metas, fps)
        for i in comp:
            cluster_id_of[i] = cid
            primary_of[i] = primary
        pairs = [sim(a, b)["overall"] for x, a in enumerate(comp) for b in comp[x + 1:]]
        cohesion = round(sum(pairs) / len(pairs), 3) if pairs else 0.0
        has_dups = any(_is_dup(sim(a, b)) for x, a in enumerate(comp) for b in comp[x + 1:])
        logic_types = {fps[i].logic_type for i in comp}
        clusters.append(Cluster(
            cluster_id=cid,
            member_ids=[fps[i].file_id for i in comp],
            member_names=[fps[i].file_name for i in comp],
            primary_id=fps[primary].file_id,
            primary_name=fps[primary].file_name,
            shared_logic_type=next(iter(logic_types)) if len(logic_types) == 1 else None,
            cohesion=cohesion,
            has_duplicates=has_dups,
        ))

    entries: list[LandscapeEntry] = []
    for i in range(n):
        entries.append(_entry_for(i, fps, metas, sim, cluster_id_of, primary_of, n))

    summary = _summarise(entries, clusters, n)
    return Landscape(entries=entries, clusters=clusters, summary=summary)


def _top_match(i: int, fps, sim, n: int) -> dict | None:
    best_j, best = None, None
    for j in range(n):
        if j == i:
            continue
        s = sim(i, j)
        if best is None or s["overall"] > best["overall"]:
            best, best_j = s, j
    if best_j is None:
        return None
    relation = ("duplicate" if _is_dup(best)
                else "similar" if best["overall"] >= SIMILAR_THRESHOLD
                else "distinct")
    return {
        "file": fps[best_j].file_name,
        "file_id": fps[best_j].file_id,
        "similarity": best["overall"],
        "signals": best,
        "relation": relation,
    }


def _entry_for(i, fps, metas, sim, cluster_id_of, primary_of, n) -> LandscapeEntry:
    fp, m = fps[i], metas[i]
    cid = cluster_id_of.get(i)
    top = _top_match(i, fps, sim, n)
    rationale: list[str] = []

    if cid is not None:
        primary = primary_of[i]
        is_primary = primary == i
        if is_primary:
            dups = [j for j in range(n) if j != i and primary_of.get(j) == i
                    and _is_dup(sim(i, j))]
            sims_members = [j for j in range(n) if j != i and cluster_id_of.get(j) == cid]
            if dups:
                rec, conf = KEEP, _confidence(0.7)
                names = ", ".join(fps[j].file_name for j in dups)
                rationale.append(
                    f"canonical of a group of {len(sims_members) + 1}; "
                    f"{len(dups)} near-duplicate(s) point here: {names}")
                rationale.append("kept as the freshest/most complete copy of the group")
            else:
                rec = CONSOLIDATE
                conf = _confidence(0.55, 0.1 * min(3, len(sims_members)))
                rationale.append(
                    f"one of {len(sims_members) + 1} similar workbook(s) — merge the "
                    "group into a single maintained file")
        else:
            primary_sim = sim(i, primary)
            if _is_dup(primary_sim):
                rec = REMOVE
                conf = _confidence(0.6, 0.25 * max(0.0, primary_sim["overall"] - DUP_THRESHOLD))
                rationale.append(
                    f"near-duplicate of '{fps[primary].file_name}' "
                    f"(similarity {primary_sim['overall']:.0%}) — redundant copy")
                rationale.append("confirm the canonical copy covers this file before removing")
            else:
                rec, conf = CONSOLIDATE, _confidence(0.55, 0.05)
                rationale.append(
                    f"similar to '{fps[primary].file_name}' "
                    f"(similarity {primary_sim['overall']:.0%}) and its group — "
                    "merge candidate")
        if fp.logic_type:
            rationale.append(f"logic type: {fp.logic_type}")
    else:
        # Standalone workbook — nothing else looks like it.
        signals = m.retirement_signals
        strong = len(signals) >= 2 or any(
            ("superseded" in s or "broken" in s or "abandoned" in s) for s in signals)
        if strong:
            rec, conf = REMOVE, _confidence(0.5)
            rationale.append("retirement signals present: " + "; ".join(signals))
            rationale.append("not a duplicate of any other file — Remove rests on "
                             "retirement signals; confirm it is genuinely unused")
        else:
            rec, conf = KEEP, _confidence(0.6)
            if signals:
                rationale.append("minor retirement signal(s): " + "; ".join(signals)
                                 + " — monitor")
            rationale.append("distinct from the rest of the folder — no duplication "
                             "or consolidation opportunity")

    return LandscapeEntry(
        file_id=fp.file_id,
        file_name=fp.file_name,
        path=m.path,
        logic_type=fp.logic_type,
        complexity=m.complexity,
        recommendation=rec,
        confidence=conf,
        rationale=rationale,
        cluster_id=cid,
        is_cluster_primary=(cid is not None and primary_of.get(i) == i),
        top_match=top,
        retirement_signals=list(m.retirement_signals),
    )


def _summarise(entries: list[LandscapeEntry], clusters: list[Cluster], n: int) -> dict:
    counts = {KEEP: 0, CONSOLIDATE: 0, REMOVE: 0}
    for e in entries:
        counts[e.recommendation] = counts.get(e.recommendation, 0) + 1
    # A cluster whose members merely look alike is a consolidation group; a
    # cluster containing near-duplicates keeps one copy and Removes the rest.
    dup_clusters = sum(1 for c in clusters if c.has_duplicates)
    return {
        "workbooks": n,
        "recommendations": counts,
        "clusters": len(clusters),
        "duplicate_clusters": dup_clusters,
        "estimated_removable": counts.get(REMOVE, 0),
        "estimated_consolidation_groups": sum(1 for c in clusters if not c.has_duplicates),
        "note": ("Whether a workbook is actually used is not knowable from the file "
                 "(needs_human). Remove verdicts flag file-intrinsic redundancy or "
                 "retirement signals only, and must be confirmed against business "
                 "usage before anything is acted on. This layer never changes a file."),
        "thresholds": {
            "similar": SIMILAR_THRESHOLD,
            "duplicate": DUP_THRESHOLD,
            "skeleton_duplicate": SKELETON_DUP,
        },
    }


# --------------------------------------------------------------------- entry


def assess_landscape(workbooks: list, assessments: list) -> Landscape:
    """Build the landscape from already-computed assessments.

    ``workbooks`` and ``assessments`` are aligned lists (same order), as produced
    by :func:`excel_xray.corpus.assess_corpus`.
    """
    fps = [fingerprint(wx, a) for wx, a in zip(workbooks, assessments)]
    metas = [_meta_from(wx, a) for wx, a in zip(workbooks, assessments)]
    return build_landscape(fps, metas)


def to_dict(landscape: Landscape) -> dict:
    return {
        "summary": landscape.summary,
        "clusters": [asdict(c) for c in landscape.clusters],
        "entries": [asdict(e) for e in landscape.entries],
    }
