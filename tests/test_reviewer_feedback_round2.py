"""Regression coverage for the second business-review feedback round."""

from copy import deepcopy
import json
import zipfile

from excel_xray import assess, write_excel_report
from excel_xray.regions import Region
from excel_xray.scan import SheetXray
from excel_xray.tabular import FILE_FIELDS, file_rows, fmt_value, to_csv


def sheet(name, *, position=0, formulas=0, reads=(), functions=(), headers=()):
    return SheetXray(
        name=name, position=position, state="visible", dimension="A1:D20",
        max_row=20, max_col=4, populated_cells=80, density=1.0,
        formula_profile={
            "total": formulas, "distinct_skeletons": 1 if formulas else 0,
            "cross_sheet_count": formulas if reads else 0,
            "external_count": 0, "hardcoded_literal_count": 0,
            "volatile_count": 0, "referenced_sheets": dict.fromkeys(reads, 1),
            "external_sources": [], "unresolved_external_indices": [],
            "top_functions": list(functions),
            "top_skeletons": [("RC[-1]*N", formulas)] if formulas else [],
        },
        regions=[Region(kind="calculation" if formulas else "data_table",
                        origin="detected", top=1, left=1, bottom=20, right=4,
                        headers=list(headers), populated_cells=80,
                        detect_confidence=0.9)],
    )


def test_scan_status_error_and_sheet_labels(xray):
    full = assess(xray).file
    assert full.scan_status.value == "full"
    assert full.scan_error.value == "N/A"
    labels = [label for _, _, label in FILE_FIELDS]
    assert "No. of Sheets - Total" in labels
    assert "No. of Sheets - Hidden" in labels

    partial_wx = deepcopy(xray)
    partial_wx.parse_status = "partial"
    partial_wx.sheets[0].notes.append("row scan capped at 5")
    partial = assess(partial_wx).file
    assert partial.scan_status.value == "partial"
    assert "row scan capped" in partial.scan_error.value

    failed_wx = deepcopy(xray)
    failed_wx.parse_status = "failed"
    failed_wx.warnings = [f"corrupt workbook relationship part in {xray.path}"]
    failed = assess(failed_wx).file
    assert failed.scan_status.value == "failed"
    assert "corrupt workbook" in failed.scan_error.value
    assert xray.path not in failed.scan_error.value


def test_client_outputs_exclude_raw_metadata(xray, tmp_path):
    a = assess(xray)
    labels = {r["field"] for r in file_rows(a)}
    assert not {"Size Bytes", "SHA-256", "Source Path"} & labels
    path = tmp_path / "review.xlsx"
    write_excel_report(xray, path, a)
    with zipfile.ZipFile(path) as zf:
        xml = b"\n".join(zf.read(n) for n in zf.namelist() if n.endswith(".xml"))
    assert xray.path.encode() not in xml
    assert xray.sha256.encode() not in xml
    assert b"Source path" not in xml and b"SHA-256" not in xml and b"Size Bytes" not in xml


def test_process_fields_and_single_workbook_findings(xray):
    fa = assess(xray).file
    assert fa.process.value
    assert fa.sub_process.value
    assert fa.potential_duplication.value == "Requires corpus analysis"
    assert fa.potential_consolidation.value == "Requires corpus analysis"
    assert fa.potential_retirement.value["verdict"] == "Not established — owner confirmation required"


def test_key_inputs_are_grouped_decoded_deduplicated_and_safe(xray):
    wx = deepcopy(xray)
    wx.external_links = [
        "file:///C:/Users/jacob/client/claims%20feed.xlsx?token=secret",
        r"C:\Users\jacob\client\claims feed.xlsx",
    ]
    wx.sheets[0].formula_profile.update(
        external_count=3,
        external_sources=[wx.external_links[0]],
        unresolved_external_indices=[],
    )
    wx.sheets[1].formula_profile.update(
        external_count=1, external_sources=[], unresolved_external_indices=[9]
    )
    a = assess(wx)
    external = [x for x in a.input_groups if x["source_type"] == "External workbooks"]
    assert len(external) == 1
    assert external[0]["source"] == "claims feed.xlsx"
    assert external[0]["reference_count"] >= 2
    assert external[0]["consumers"] == [wx.sheets[0].name]
    assert any(x["source_type"] == "Unresolved sources" for x in a.input_groups)
    rendered = fmt_value(a.file.key_inputs.value)
    assert all(group in rendered for group in (
        "In-workbook tabs", "External workbooks", "Formal data connections",
        "Pivot sources", "Unresolved sources",
    ))
    assert "Users" not in rendered and "token=secret" not in rendered and "%20" not in rendered


def test_tb_is_input_purpose_not_a_source_system(xray):
    wx = deepcopy(xray)
    wx.sheets = [sheet("TB", headers=["Account", "Balance"])]
    wx.connections = []
    a = assess(wx)
    assert any(x["purpose"] == "Trial balance" for x in a.input_groups)
    assert a.file.source_system.value == "Not established — owner confirmation required"


def test_key_outputs_exclude_workings_and_helpers(xray):
    wx = deepcopy(xray)
    wx.sheets = [
        sheet("Raw data", position=0),
        sheet("Working", position=1, formulas=20, reads=["Raw data"]),
        sheet("Management report", position=2, formulas=20, reads=["Working"]),
        sheet("Mapping", position=3),
    ]
    outputs = assess(wx).file.key_outputs.value
    joined = " ".join(outputs)
    assert "Management report" in joined
    assert "Working" not in joined and "Mapping" not in joined and "Raw data" not in joined

    from excel_xray.narrative import Narrative

    class OverinclusiveAssessor:
        basis = "inferred"
        label = "test"

        def narrate(self, bundle):
            return Narrative(
                purpose_of_file="Test purpose", key_output_outcome="Test outcome",
                key_outputs=["Working schedule", "Management report deliverable"],
                process="Not established — owner confirmation required",
                sub_process="Not established — owner confirmation required",
                tabs={t["name"]: "Test tab purpose" for t in bundle["tabs"]},
            )

    model_outputs = assess(wx, assessor=OverinclusiveAssessor()).file.key_outputs.value
    assert model_outputs == ["Management report deliverable"]


def test_logic_calculations_reconciliation_and_manual_intervention(xray):
    wx = deepcopy(xray)
    wx.sheets = [sheet("Lookup", formulas=70, functions=[("VLOOKUP", 70)])]
    lookup = assess(wx).file
    assert lookup.reconciliation_logic.value["count"] == 0
    assert lookup.reconciliation_logic.value["candidates_requiring_review"]
    assert lookup.manual_intervention.basis == "needs_human"
    assert "Not established" in lookup.manual_intervention.value
    assert "top_functions" not in lookup.key_calculations_logic.value
    assert "VLOOKUP" not in fmt_value(lookup.key_calculations_logic.value)

    wx.sheets += [
        sheet("Trial balance", position=1),
        sheet("Ledger", position=2),
        sheet("GL Reconciliation", position=3, formulas=90,
              reads=["Trial balance", "Ledger"],
              headers=["Account code", "Variance", "Tolerance"]),
    ]
    result = assess(wx).file.reconciliation_logic.value
    assert result["count"] == 1
    confirmed = result["reconciliations"][0]
    assert {"overview", "source_a", "source_b", "matching_criteria", "tolerance",
            "exception_logic", "evidence", "confidence"} <= confirmed.keys()
    assert confirmed["source_a"] != confirmed["source_b"]


def test_logic_types_are_unique_and_connections_are_separate(xray):
    wx = deepcopy(xray)
    wx.external_links = ["file:///C:/private/external.xlsx"]
    wx.sheets[0].formula_profile.update(external_count=17, external_sources=wx.external_links)
    wx.connections = []
    fa = assess(wx).file
    labels = [label for _, _, label in FILE_FIELDS]
    assert labels.count("Logic Types") == 1
    assert "Logic Type" not in labels
    mechanisms = fa.macros_vba_external_links.value
    assert mechanisms["formal_data_connections"] == 0
    assert mechanisms["external_workbook_link_references"] == 17
    assert "Formal data connections: none detected" in mechanisms["summary"]
    assert "External workbook links: 17 references" in mechanisms["summary"]


def test_opportunities_carry_process_context(xray):
    fa = assess(xray).file
    for field in (fa.potential_simplification, fa.potential_automation):
        for item in field.value.get("details", []):
            assert {"process", "sub_process", "worksheets", "observed_evidence",
                    "candidate_action", "confirmation_required"} <= item.keys()
            assert item["process"] == fa.process.value
            assert item["sub_process"] == fa.sub_process.value


def test_assessment_json_contains_no_source_path(xray):
    payload = json.dumps(assess(xray), default=lambda obj: obj.__dict__)
    assert xray.path not in payload
    assert xray.sha256 not in payload


def test_csv_shortens_caller_supplied_file_path(xray, tmp_path):
    path = tmp_path / "assessment.csv"
    to_csv([(xray.path, assess(xray))], path)
    text = path.read_text()
    assert xray.path not in text
    assert xray.filename in text
