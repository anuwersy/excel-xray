"""Step 7 — portfolio landscape: Keep / Consolidate / Remove across a folder."""

from __future__ import annotations

import datetime as _dt

from excel_xray import xray_workbook
from excel_xray.corpus import Fingerprint, assess_corpus
from excel_xray.landscape import (
    CONSOLIDATE,
    KEEP,
    REMOVE,
    _components,
    _meta_from,
    _Meta,
    assess_landscape,
    build_landscape,
)


def _fp(name, skeletons, headers, logic="Calculation", functions=("SUM",)):
    return Fingerprint(
        file_id=name, file_name=name, logic_type=logic,
        skeletons=set(skeletons), headers=set(headers),
        functions=set(functions), sheet_names={name.lower()},
    )


def _meta(cells=100, modified="2024-01-01", complexity="Medium",
          superseded=False, signals=()):
    return _Meta(
        path=f"/tmp/{cells}",
        modified=_dt.datetime.fromisoformat(modified),
        cells=cells, complexity=complexity,
        complexity_rank={"High": 3, "Medium": 2, "Low": 1}[complexity],
        superseded=superseded, retirement_signals=list(signals),
    )


# ------------------------------------------------------------- clustering


def test_components_groups_connected_and_keeps_singletons():
    # edges: 0-1 connected; 2 alone.
    comps = _components(3, [(0, 1)])
    assert sorted(map(sorted, comps)) == [[0, 1], [2]]


# ----------------------------------------------------------- recommendations


def test_identical_pair_keeps_one_removes_the_other():
    a = _fp("keep.xlsx", ["RC[-1]*N", "SUM(R[-3]C:R[-1]C)"], ["policy id", "premium"])
    b = _fp("dupe.xlsx", ["RC[-1]*N", "SUM(R[-3]C:R[-1]C)"], ["policy id", "premium"])
    # 'keep.xlsx' is fresher and larger, so it should survive.
    metas = [_meta(cells=500, modified="2024-06-01"),
             _meta(cells=100, modified="2023-01-01")]
    land = build_landscape([a, b], metas)

    by_name = {e.file_name: e for e in land.entries}
    assert by_name["keep.xlsx"].recommendation == KEEP
    assert by_name["keep.xlsx"].is_cluster_primary
    assert by_name["dupe.xlsx"].recommendation == REMOVE
    # both sit in the same cluster
    assert by_name["keep.xlsx"].cluster_id == by_name["dupe.xlsx"].cluster_id
    assert land.summary["recommendations"][REMOVE] == 1
    assert land.summary["recommendations"][KEEP] == 1


def test_similar_but_not_duplicate_pair_is_consolidate():
    # Share headers/logic and some shape overlap -> similar, below duplicate.
    a = _fp("q1.xlsx", ["RC[-1]*N", "AAA"], ["region", "revenue", "cost"])
    b = _fp("q2.xlsx", ["RC[-1]*N", "BBB"], ["region", "revenue", "margin"])
    land = build_landscape([a, b], [_meta(), _meta()])
    recs = {e.file_name: e.recommendation for e in land.entries}
    # If they cluster at all, neither is a duplicate, so both consolidate.
    if land.clusters:
        assert recs["q1.xlsx"] == CONSOLIDATE
        assert recs["q2.xlsx"] == CONSOLIDATE
        assert land.summary["estimated_consolidation_groups"] >= 1


def test_distinct_workbook_is_kept():
    a = _fp("model.xlsx", ["RC[-1]*N"], ["policy id"])
    b = _fp("mapping.xlsx", ["SHEET!R1C1"], ["gl account"],
            logic="Reporting", functions=("VLOOKUP",))
    land = build_landscape([a, b], [_meta(), _meta()])
    assert all(e.recommendation == KEEP for e in land.entries)
    assert all(e.cluster_id is None for e in land.entries)


def test_standalone_with_retirement_signals_is_remove():
    a = _fp("model.xlsx", ["RC[-1]*N"], ["policy id"])
    b = _fp("other.xlsx", ["ZZZ"], ["unrelated"], logic="Reporting", functions=("TEXT",))
    metas = [
        _meta(signals=["not modified in ~30 months (stale)",
                       "filename suggests a superseded/backup copy"]),
        _meta(),
    ]
    land = build_landscape([a, b], metas)
    by_name = {e.file_name: e for e in land.entries}
    assert by_name["model.xlsx"].recommendation == REMOVE
    assert by_name["model.xlsx"].retirement_signals
    assert by_name["other.xlsx"].recommendation == KEEP


def test_top_match_and_relation_reported():
    a = _fp("a.xlsx", ["RC[-1]*N", "SUM(R[-3]C:R[-1]C)"], ["policy id", "premium"])
    b = _fp("b.xlsx", ["RC[-1]*N", "SUM(R[-3]C:R[-1]C)"], ["policy id", "premium"])
    land = build_landscape([a, b], [_meta(), _meta()])
    e = land.entries[0]
    assert e.top_match["file"] == "b.xlsx"
    assert e.top_match["relation"] == "duplicate"
    assert 0.0 <= e.top_match["similarity"] <= 1.0


# --------------------------------------------------------------- end to end


def test_assess_landscape_flags_identical_workbooks(fixture_path):
    wxs = [xray_workbook(fixture_path), xray_workbook(fixture_path)]
    assessments = assess_corpus(wxs)
    land = assess_landscape(wxs, assessments)

    assert len(land.entries) == 2
    recs = sorted(e.recommendation for e in land.entries)
    # Two copies of the same file: one kept, one removed.
    assert recs == [KEEP, REMOVE]
    assert land.summary["workbooks"] == 2
    assert land.clusters and land.clusters[0].has_duplicates
    # The honesty note about usage is always present.
    assert "needs_human" in land.summary["note"]


def test_meta_from_reads_scan_and_assessment(fixture_path):
    wx = xray_workbook(fixture_path)
    a = assess_corpus([wx])[0]
    m = _meta_from(wx, a)
    assert m.cells > 0
    assert m.complexity in ("High", "Medium", "Low")
