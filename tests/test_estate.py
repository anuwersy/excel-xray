"""Step 7 — estate comparison: relationship typing and clustering."""

from __future__ import annotations

from excel_xray import assess
from excel_xray.estate import (
    EucFingerprint,
    _clusters,
    build_estate,
    compare,
)


def _fp(name, *, skel=(), inp=(), out=(), topo=None, funcs=()):
    return EucFingerprint(
        file_id=name, file_name=name, logic_type="Calculation",
        skeletons=set(skel), input_sig=set(inp), output_sig=set(out),
        topology=topo or {"Input": 1, "Calculation": 1, "Output": 1, "edges": 2},
        headers=set(inp) | set(out), functions=set(funcs),
    )


def test_identical_is_duplicate():
    a = _fp("a", skel=["s1", "s2", "s3"], inp=["prem", "lob"], out=["region", "total"])
    b = _fp("b", skel=["s1", "s2", "s3"], inp=["prem", "lob"], out=["region", "total"])
    assert compare(a, b)["relationship"] == "Duplicate"


def test_same_output_different_method():
    a = _fp("a", skel=["x1", "x2"], out=["region", "q1", "q2", "fy"])
    b = _fp("b", skel=["y1", "y2"], out=["region", "q1", "q2", "fy"])  # disjoint shapes
    c = compare(a, b)
    assert c["skeleton"] == 0.0 and c["output"] >= 0.6
    assert c["relationship"] == "Same output, different method"


def test_overlapping_logic():
    # 2 shared skeletons of 4 union = 0.5 (>= 0.40), but outputs differ.
    a = _fp("a", skel=["s1", "s2", "s3"], out=["oa"])
    b = _fp("b", skel=["s1", "s2", "s4"], out=["ob"])
    assert compare(a, b)["relationship"] == "Overlapping logic"


def test_shared_source():
    a = _fp("a", skel=["a1", "a2"], inp=["feed_x", "col1", "col2"], out=["oa1", "oa2"])
    b = _fp("b", skel=["b1", "b2"], inp=["feed_x", "col1", "col2"], out=["ob1", "ob2"])
    c = compare(a, b)
    assert c["input"] >= 0.6 and c["skeleton"] < 0.4 and c["output"] < 0.4
    assert c["relationship"] == "Shared source"


def test_unrelated():
    a = _fp("a", skel=["a1"], inp=["ia"], out=["oa"])
    b = _fp("b", skel=["b1"], inp=["ib"], out=["ob"])
    c = compare(a, b)
    assert c["relationship"] == "Unrelated" and not c["linked"]


def test_clusters_connected_components():
    # 0-1 linked, 1-2 linked (transitive family), 3 alone, 4-5 linked
    groups = _clusters(6, [(0, 1), (1, 2), (4, 5)])
    sizes = sorted(len(g) for g in groups)
    assert sizes == [1, 2, 3]  # {3}, {4,5}, {0,1,2}


def test_build_estate_on_real_fixture(xray):
    a = assess(xray)
    # Two identical copies -> one Duplicate family, no singletons.
    estate = build_estate([(xray, a), (xray, a)])
    assert len(estate.clusters) == 1
    assert len(estate.clusters[0]) == 2
    assert estate.singletons == []
    assert estate.pairs[0][2]["relationship"] == "Duplicate"


def test_estate_report_renders(xray):
    from excel_xray.estate_report import build_estate_report
    a = assess(xray)
    html = build_estate_report(build_estate([(xray, a), (xray, a)]))
    assert "EUC estate comparison" in html
    assert "Duplicate" in html
