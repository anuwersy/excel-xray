"""Regression cases for the business-review feedback, independent of LLM calls."""

from copy import deepcopy
import json
import zipfile

from lxml import etree as ET

from excel_xray import assess, xray_workbook, write_excel_report
from excel_xray.assessment import to_dict
from excel_xray.formulas import analyse
from excel_xray.report import build_report
from excel_xray.scan import SheetXray
from excel_xray.regions import Region
from excel_xray.review import short_source
from excel_xray.tabular import fmt_value


def sheet(name, *, position=0, hidden=False, formulas=0, reads=(), functions=(), errors=(), headers=()):
    return SheetXray(name=name, position=position, state="hidden" if hidden else "visible",
        dimension="A1:D30", max_row=30, max_col=4, populated_cells=100, density=0.83,
        formula_profile={"total": formulas, "distinct_skeletons": 1 if formulas else 0,
                         "cross_sheet_count": formulas if reads else 0,
                         "compression": formulas, "external_count": 0,
                         "hardcoded_literal_count": 0, "volatile_count": 0,
                         "referenced_sheets": dict.fromkeys(reads, 1),
                         "top_functions": list(functions), "top_skeletons": [("RC[-1]*N", formulas)] if formulas else []},
        error_cells=[f"{name}!{e}" for e in errors],
        regions=[Region(kind="calculation" if formulas else "data_table", origin="detected",
                        top=1, left=1, bottom=30, right=4, headers=list(headers),
                        populated_cells=100, detect_confidence=0.9)])


def model(xray):
    wx = deepcopy(xray)
    wx.sheets = [
        sheet("Claims workings", hidden=True, formulas=70, reads=["Claims data"], errors=["B4 #REF!", "C5 #REF!", "D6 #DIV/0!"]),
        sheet("Claims data", position=1, headers=["Claim ID", "Paid amount"]),
        sheet("Management report", position=2, formulas=80, reads=["Claims workings"], functions=[("SUM", 80)]),
        sheet("Archive workings", position=3, hidden=True, formulas=50, reads=["Management report"]),
        sheet("COA", position=4, headers=["Account code", "Description"]),
    ]
    wx.has_vba = True
    wx.external_links = []
    wx.connections = []
    wx.pivot_cache_sources = []
    return wx


def test_visibility_order_multi_roles_and_output_with_downstream(xray):
    a = assess(model(xray))
    assert [t.tab_visibility.value for t in a.tabs] == ["Visible"] * 3 + ["Hidden"] * 2
    tabs = {t.tab_name.value: t for t in a.tabs}
    report = tabs["Management report"]
    assert report.tab_category.value == "Output"
    assert report.tab_information_analysis.value == "output"
    assert report.tab_roles.value == ["Output", "Calculation"]
    assert tabs["COA"].tab_category.value == "Mapping"
    assert tabs["COA"].tab_information_analysis.value == "supporting"
    assert all("COA" not in output for output in a.file.key_outputs.value)
    assert set(a.file.logic_types.value) >= {"Calculation", "Reporting", "Data Transformation"}


def test_errors_grouped_by_type_area_region_and_potential_outputs(xray):
    a = assess(model(xray))
    assert sum(r["count"] for r in a.error_summary) == 3
    ref = next(r for r in a.error_summary if r["type"] == "#REF!")
    assert ref["count"] == 2 and ref["region"] == "A1:D30"
    assert ref["area"] == "Working / supporting"
    assert ref["potential_outputs"] == ["Management report"]
    assert "not proven" in ref["impact"]
    assert "Recalculate" in ref["recalculation"]
    html = build_report(model(xray), a)
    assert "2 of 5 sheets are hidden" in html
    assert "#REF!: 2" in html and "Error summary" in html
    assert "Claims data / Calculation" in html


def test_error_in_final_output_and_no_output_path(xray):
    wx = model(xray)
    wx.sheets[2].error_cells = ["Management report!A2 #N/A"]
    wx.sheets[4].error_cells = ["COA!B3 #N/A"]
    rows = assess(wx).error_summary
    assert next(r for r in rows if r["tab"] == "Management report")["area"] == "Final-output candidate"
    assert next(r for r in rows if r["tab"] == "COA")["potential_outputs"] == []


def test_error_reachability_handles_cycles_and_case_insensitive_refs(xray):
    wx = model(xray)
    wx.sheets[0].formula_profile["referenced_sheets"] = {"archive workings": 1}
    wx.sheets[2].formula_profile["referenced_sheets"] = {"CLAIMS WORKINGS": 1}
    # Claims workings -> Management report -> Archive workings -> Claims workings.
    a = assess(wx)
    assert all(r["potential_outputs"] == ["Management report"] for r in a.error_summary)


def test_review_is_not_blanket_for_vba_or_hidden_tabs(xray):
    a = assess(model(xray))
    tabs = {t.tab_name.value: t for t in a.tabs}
    assert tabs["Claims data"].human_validation_required.value == "N"
    assert tabs["Archive workings"].human_validation_required.value == "N"
    assert tabs["Claims workings"].human_validation_required.value == "Y"
    assert all(q.endswith("?") for q in tabs["Claims workings"].validation_reason.value)
    assert a.file.potential_retirement.basis == "needs_human"
    assert not any("hidden" in signal or "error" in signal for signal in a.file.potential_retirement.value["signals"])
    assert a.file.manual_intervention.basis == "needs_human"
    assert "manually entered" not in fmt_value(a.file.potential_automation.value)


def test_inputs_grouped_and_consumers_grounded(xray):
    wx = model(xray)
    wx.external_links = ["file:///C:/finance/claims%20feed.xlsx", r"C:\finance\fx.xlsx"]
    wx.sheets[0].formula_profile.update(external_count=3, external_sources=[wx.external_links[0]])
    wx.sheets[2].formula_profile.update(external_count=1, external_sources=[], unresolved_external_indices=[7])
    a = assess(wx)
    groups = {g["source"]: g for g in a.input_groups}
    assert groups["claims feed.xlsx"]["purpose"] == "Claims data"
    assert groups["claims feed.xlsx"]["consumers"] == ["Claims workings"]
    assert groups["fx.xlsx"]["consumers"] == []
    assert groups["Claims data"]["consumers"] == ["Claims workings"]
    tabs = {t.tab_name.value: t for t in a.tabs}
    assert tabs["Management report"].upstream_dependencies.value["unresolved_external"]
    assert "C:/" not in fmt_value(a.file.key_inputs.value)
    assert "finance" not in fmt_value(tabs["Claims workings"].upstream_dependencies.value)
    assert short_source("https://example.com/a/fx.xlsx?token=secret") == "fx.xlsx"


def test_lookup_only_is_not_reconciliation_and_keys_are_candidates(xray):
    wx = deepcopy(xray)
    wx.sheets = [sheet("Mapping", formulas=70, functions=[("VLOOKUP", 70)])]
    a = assess(wx)
    assert a.file.reconciliation_logic.value["count"] == 0
    assert a.file.reconciliation_logic.value["status"] == "candidates_requiring_review"
    assert a.file.logic_type.value == "Data Transformation"
    wx.sheets += [sheet("Trial balance", position=1), sheet("Ledger", position=2),
                  sheet("Reconciliation", position=3, formulas=90,
                        reads=["Trial balance", "Ledger"], headers=["Account code", "Variance"])]
    a = assess(wx)
    result = a.file.reconciliation_logic.value
    assert result["count"] == 1
    r = result["reconciliations"][0]
    assert [r["source_a"], r["source_b"]] == ["Ledger", "Trial balance"]
    assert "Account code" in r["matching_criteria"]
    assert r["tolerance"] == "Not established"
    assert "Reconciliation / Control" in a.file.logic_types.value


def test_stored_cells_do_not_prove_manual_input(xray):
    wx = deepcopy(xray)
    wx.sheets = [sheet("Raw data")]
    a = assess(wx)
    assert "Manual Input" not in a.file.logic_types.value
    assert a.file.manual_intervention.basis == "needs_human"
    assert a.file.potential_automation.value["verdict"] == "Not established"


def test_review_detail_exports_and_unknowns_remain_explicit(xray, tmp_path):
    wx = model(xray)
    a = assess(wx)
    payload = to_dict(a)
    assert {"error_summary", "input_groups", "hidden_groups"} <= payload.keys()
    assert payload["file"]["usage_frequency"]["value"] is None
    assert "How often" in payload["file"]["usage_frequency"]["evidence"][0]
    path = tmp_path / "review.xlsx"
    write_excel_report(wx, path, a)
    with zipfile.ZipFile(path) as zf:
        workbook = zf.read("xl/workbook.xml").decode()
        assert all(name in workbook for name in ("Error summary", "Input sources", "Hidden sheet groups", "Calculation steps"))
        assert "2 of 5 sheets are hidden" in zf.read("xl/worksheets/sheet7.xml").decode()
    assert "59937" not in json.dumps(payload)


def test_external_index_parsing():
    for formula in ("=[2]Claims!A1", "='[2]Claims data'!A1"):
        facts = analyse(formula, 1, 1)
        assert facts.external_indices == {2}
        assert facts.external_refs == 1 and not facts.referenced_sheets
    assert not analyse('="[2]Claims!A1"', 1, 1).external_indices


def test_external_reference_order_and_consumer_mapping(fixture_path, tmp_path):
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
    with zipfile.ZipFile(fixture_path) as zf:
        parts = {n: zf.read(n) for n in zf.namelist()}
    wb = ET.fromstring(parts["xl/workbook.xml"])
    refs = ET.SubElement(wb, f"{{{main}}}externalReferences")
    relationships = ET.fromstring(parts["xl/_rels/workbook.xml.rels"])
    for index, rid, source in ((1, "extZ", "../fx.xlsx"), (2, "extA", "file:///C:/finance/claims.xlsx")):
        ET.SubElement(refs, f"{{{main}}}externalReference", {f"{{{rel}}}id": rid})
        ET.SubElement(relationships, f"{{{pkg}}}Relationship", Id=rid, Type=f"{rel}/externalLink", Target=f"externalLinks/externalLink{index}.xml")
        parts[f"xl/externalLinks/externalLink{index}.xml"] = f'<externalLink xmlns="{main}" xmlns:r="{rel}"><externalBook r:id="book"/></externalLink>'.encode()
        parts[f"xl/externalLinks/_rels/externalLink{index}.xml.rels"] = f'<Relationships xmlns="{pkg}"><Relationship Id="book" Type="{rel}/externalLinkPath" Target="{source}" TargetMode="External"/></Relationships>'.encode()
    parts["xl/workbook.xml"] = ET.tostring(wb)
    parts["xl/_rels/workbook.xml.rels"] = ET.tostring(relationships)
    data = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    row = ET.SubElement(data.find(f"{{{main}}}sheetData"), f"{{{main}}}row", r="999")
    cell = ET.SubElement(row, f"{{{main}}}c", r="A999")
    ET.SubElement(cell, f"{{{main}}}f").text = "'[2]Claims data'!A1"
    ET.SubElement(cell, f"{{{main}}}v").text = "1"
    parts["xl/worksheets/sheet1.xml"] = ET.tostring(data)
    path = tmp_path / "external.xlsx"
    with zipfile.ZipFile(path, "w") as zf:
        for n, content in parts.items():
            zf.writestr(n, content)
    wx = xray_workbook(str(path))
    assert wx.external_links == ["../fx.xlsx", "file:///C:/finance/claims.xlsx"]
    assert wx.sheets[0].formula_profile["external_sources"] == ["file:///C:/finance/claims.xlsx"]
    assert all(not s.formula_profile["external_sources"] for s in wx.sheets[1:])


def test_scan_cap_is_partial(fixture_path):
    assert xray_workbook(fixture_path, max_rows=5).parse_status == "partial"


def test_large_workbook_keeps_detail_out_of_summary(xray, tmp_path):
    wx = deepcopy(xray)
    wx.sheets = [sheet(f"Calculation {i}", position=i, hidden=i >= 29,
                       formulas=70, functions=[("SUM", 70)]) for i in range(74)]
    for s in wx.sheets:
        s.formula_profile["hardcoded_literal_count"] = 20
    a = assess(wx)
    assert sum(r["count"] for r in a.hidden_groups) == 45
    assert len(a.file.key_calculations_logic.value["steps"]) == 74
    assert "all 74 tab operations" in a.file.key_calculations_logic.value["summary"]
    assert len(a.file.potential_simplification.value["opportunities"]) == 74
    assert "69 further candidates" in fmt_value(a.file.potential_simplification.value)
    path = tmp_path / "large.xlsx"
    write_excel_report(wx, path, a)
    main = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(path) as zf:
        steps = ET.fromstring(zf.read("xl/worksheets/sheet11.xml"))
        assert len(steps.findall("s:sheetData/s:row", main)) == 75
        hidden = ET.fromstring(zf.read("xl/worksheets/sheet9.xml"))
        assert len(hidden.findall("s:sheetData/s:row", main)) == 46
        assert "45 of 74 sheets are hidden" in zf.read("xl/worksheets/sheet7.xml").decode()
