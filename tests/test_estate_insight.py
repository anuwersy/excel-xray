"""Step 8 — estate insight layer (families + recommendations)."""

from __future__ import annotations

from excel_xray import assess
from excel_xray.estate import build_estate
from excel_xray.estate_insight import (
    EstateInsight,
    build_estate_bundle,
    generate_estate_insight,
)


def _duplicate_estate(xray):
    a = assess(xray)
    return build_estate([(xray, a), (xray, a)])


def test_offline_insight_recommends_consolidation(xray):
    estate = _duplicate_estate(xray)
    ins = generate_estate_insight(estate)  # offline default
    assert ins.basis == "drafted"
    assert len(ins.families) == 1
    fam = ins.families[0]
    assert fam.basis == "drafted"
    assert "consolidate" in fam.recommended_action.lower()
    assert fam.summary and fam.rationale


def test_estate_summary_and_opportunities(xray):
    ins = generate_estate_insight(_duplicate_estate(xray))
    assert "family" in ins.estate_summary.lower()
    assert ins.top_opportunities  # at least one ranked action


def test_bundle_is_value_free(xray):
    import json
    bundle = build_estate_bundle(_duplicate_estate(xray))
    # A distinctive cached figure must not leak into the estate bundle.
    assert "59937" not in json.dumps(bundle, default=str)
    assert bundle["family_count"] == 1


def test_custom_assessor_marks_inferred(xray):
    class FakeLLM:
        basis = "inferred"
        label = "fake"

        def insight(self, bundle):
            from excel_xray.estate_insight import FamilyInsight
            fams = [FamilyInsight(summary="s", recommended_action="Keep one",
                                  rationale="r", basis="inferred")
                    for _ in bundle["families"]]
            return EstateInsight(families=fams, estate_summary="overview",
                                 top_opportunities=["do X"], basis="inferred")

    ins = generate_estate_insight(_duplicate_estate(xray), FakeLLM())
    assert ins.basis == "inferred"
    assert ins.families[0].recommended_action == "Keep one"


def test_report_includes_insight(xray):
    from excel_xray.estate_report import build_estate_report
    estate = _duplicate_estate(xray)
    html = build_estate_report(estate, generate_estate_insight(estate))
    assert "Estate insight" in html
    assert "Top opportunities" in html
