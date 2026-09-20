"""Excel exports preserve assessment evidence and do not modify inputs."""

import copy
import csv
import hashlib
import shutil
import sys
import xml.etree.ElementTree as ET
import zipfile

import pytest

from excel_xray import assess, build_estate, generate_estate_insight
from excel_xray.cli import collect, main
from excel_xray.excel_report import is_excel_report, write_excel_report, write_estate_excel_report
from excel_xray.tabular import FILE_FIELDS, TAB_FIELDS

NS = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_sheet(path, index):
    with zipfile.ZipFile(path) as zf:
        assert zf.testzip() is None
        root = ET.fromstring(zf.read(f"xl/worksheets/sheet{index}.xml"))
    rows = []
    for row in root.findall("s:sheetData/s:row", NS):
        values = {}
        for c in row:
            value = c.findtext("s:is/s:t", namespaces=NS) if c.get("t") == "inlineStr" else c.findtext("s:v", namespaces=NS)
            values[c.get("r")] = value
        rows.append(values)
    return root, rows


def run_cli(monkeypatch, *args):
    monkeypatch.setattr(sys, "argv", ["excel-xray", *map(str, args)])
    return main()


def test_excel_assessments_and_numeric_confidence(xray, tmp_path):
    path = tmp_path / "report.xlsx"
    a = assess(xray)
    before = hashlib.sha256(open(xray.path, "rb").read()).hexdigest()
    write_excel_report(xray, path, a)
    root, rows = read_sheet(path, 1)
    assert len(rows) == len(FILE_FIELDS) + 1
    assert rows[1]["C2"] == a.file.file_id.value
    complexity = next(row for row in rows if "Complexity" in row.values())
    assert float(next(v for k, v in complexity.items() if k.startswith("E"))) == a.file.complexity.confidence
    assert root.find("s:sheetViews/s:sheetView/s:pane", NS).get("state") == "frozen"
    assert root.find("s:autoFilter", NS) is not None
    assert len(read_sheet(path, 2)[1]) == len(a.tabs) * len(TAB_FIELDS) + 1
    assert is_excel_report(path)
    assert hashlib.sha256(open(xray.path, "rb").read()).hexdigest() == before


def test_untrusted_text_stays_text(xray, tmp_path):
    a = assess(xray)
    a.file.purpose_of_file.value = '=HYPERLINK("https://example.com","click")'
    a.file.key_output_outcome.value = "資料 & <review>\x00"
    path = tmp_path / "report.xlsx"
    write_excel_report(xray, path, a)
    root, rows = read_sheet(path, 1)
    assert not root.findall(".//s:f", NS)
    assert any(a.file.purpose_of_file.value in row.values() for row in rows)
    assert any("資料 & <review>" in row.values() for row in rows)


def test_refuse_source_and_unrelated_workbook_overwrites(xray, tmp_path):
    a = assess(xray)
    with pytest.raises(ValueError, match="source workbook"):
        write_excel_report(xray, xray.path, a)
    other = tmp_path / "existing.xlsx"
    shutil.copyfile(xray.path, other)
    with pytest.raises(ValueError, match="non-report"):
        write_excel_report(xray, other, a)


def test_oversized_text_leaves_existing_report_intact(xray, tmp_path):
    a = assess(xray)
    path = tmp_path / "report.xlsx"
    write_excel_report(xray, path, a)
    original = path.read_bytes()
    a.file.purpose_of_file.value = "x" * 32768
    with pytest.raises(ValueError, match="character limit"):
        write_excel_report(xray, path, a)
    assert path.read_bytes() == original


def test_estate_export_preserves_pairs_and_insight(xray, tmp_path):
    other = copy.deepcopy(xray)
    other.filename = "second.xlsx"
    estate = build_estate([(xray, assess(xray)), (other, assess(other))])
    insight = generate_estate_insight(estate)
    path = tmp_path / "estate.xlsx"
    write_estate_excel_report(estate, path, insight)
    _, rows = read_sheet(path, 2)
    assert len(rows) == len(estate.pairs) + 1
    assert rows[1]["B2"] == "second.xlsx"
    assert float(rows[1]["H2"]) == estate.pairs[0][2]["overall"]
    assert any(insight.families[0].recommended_action in r.values() for r in read_sheet(path, 3)[1])


def test_default_cli_and_rescan_excludes_reports(fixture_path, tmp_path, monkeypatch):
    source = tmp_path / "model.xlsx"
    shutil.copyfile(fixture_path, source)
    assert run_cli(monkeypatch, tmp_path) == 0
    assert is_excel_report(tmp_path / "xray_model.xlsx")
    assert collect(str(tmp_path)) == [str(source)]
    assert run_cli(monkeypatch, tmp_path) == 0
    assert not (tmp_path / "xray_xray_model.xlsx").exists()


def test_folder_estate_and_same_stem_outputs(fixture_path, tmp_path, monkeypatch):
    for folder in ("a", "b"):
        (tmp_path / folder).mkdir()
        shutil.copyfile(fixture_path, tmp_path / folder / "model.xlsx")
    out = tmp_path / "reports"
    assert run_cli(monkeypatch, tmp_path, "--estate", "-o", out) == 0
    assert len(list(out.glob("xray_model_*.xlsx"))) == 2
    assert is_excel_report(out / "estate.xlsx")
    assert len(collect(str(tmp_path))) == 2


def test_html_and_json_modes_remain_available(fixture_path, tmp_path, monkeypatch, capsys):
    assert run_cli(monkeypatch, fixture_path, "--format", "html", "-o", tmp_path) == 0
    assert (tmp_path / "xray_messy_reserving_model.html").exists()
    assert not list(tmp_path.glob("*.xlsx"))
    capsys.readouterr()
    assert run_cli(monkeypatch, fixture_path, "--json", "-o", tmp_path) == 0
    assert '"filename": "messy_reserving_model.xlsx"' in capsys.readouterr().out


def test_json_rejects_silently_ignored_exports(fixture_path, monkeypatch):
    with pytest.raises(SystemExit) as error:
        run_cli(monkeypatch, fixture_path, "--json", "--estate")
    assert error.value.code == 2


def test_failed_scan_is_exported_to_excel_csv_and_assessment_json(tmp_path, monkeypatch, capsys):
    bad = tmp_path / "broken.xlsx"
    bad.write_bytes(b"not an OOXML workbook")

    assert run_cli(monkeypatch, bad, "--assess") == 0
    payload = __import__("json").loads(capsys.readouterr().out)
    assert payload["file"]["scan_status"]["value"] == "failed"
    assert "not a zip container" in payload["file"]["scan_error"]["value"]

    out = tmp_path / "reports"
    csv_path = out / "assessment.csv"
    assert run_cli(monkeypatch, bad, "-o", out, "--csv", csv_path) == 0
    report = out / "xray_broken.xlsx"
    assert is_excel_report(report)
    _, rows = read_sheet(report, 1)
    status = next(row for row in rows if "Scan Status" in row.values())
    error = next(row for row in rows if "Scan Error" in row.values())
    assert "failed" in status.values()
    assert any("not a zip container" in value for value in error.values())
    with open(csv_path, newline="", encoding="utf-8") as fh:
        csv_rows = list(csv.DictReader(fh))
    assert any(r["Field"] == "Scan Status" and r["Value"] == "failed" for r in csv_rows)
