"""Native Excel reports using the existing lxml dependency and OOXML.

All extracted text is stored as text, never as executable Excel formulas.
Report workbooks carry a marker so folder scans do not ingest their own output.
"""

from __future__ import annotations

import math
import os
import tempfile
import zipfile
from pathlib import Path

from lxml import etree as ET

from .tabular import file_rows, tab_rows
from .review import short_source
from .util import get_column_letter

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT = "http://schemas.openxmlformats.org/package/2006/content-types"
MARKER = "excel-xray-report.txt"


def is_excel_report(path) -> bool:
    try:
        with zipfile.ZipFile(path) as zf:
            return zf.read(MARKER) == b"Excel X-ray report v1\n"
    except (OSError, KeyError, zipfile.BadZipFile):
        return False


def _element(tag, **attrs):
    return ET.Element(f"{{{MAIN}}}{tag}", nsmap={None: MAIN},
                      **{k: str(v) for k, v in attrs.items()})


def _child(parent, tag, **attrs):
    el = _element(tag, **attrs)
    parent.append(el)
    return el


def _xml(el):
    return ET.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def _text(value):
    # XML 1.0 rejects control characters; preserve all permitted Unicode.
    value = str(value)
    value = "".join(c for c in value if c in "\t\n\r" or
                    0x20 <= ord(c) <= 0xD7FF or 0xE000 <= ord(c) <= 0xFFFD or
                    0x10000 <= ord(c) <= 0x10FFFF)
    if len(value.encode("utf-16-le")) // 2 > 32767:
        raise ValueError("Excel report cell exceeds the 32,767-character limit")
    return value


def _styles():
    root = _element("styleSheet")
    fonts = _child(root, "fonts", count=2)
    for bold in (False, True):
        font = _child(fonts, "font")
        _child(font, "sz", val=11)
        _child(font, "name", val="Calibri")
        _child(font, "color", rgb="FFFFFFFF" if bold else "FF172B4D")
        if bold:
            _child(font, "b")
    fills = _child(root, "fills", count=3)
    for pattern in ("none", "gray125", "solid"):
        fill = _child(fills, "fill")
        p = _child(fill, "patternFill", patternType=pattern)
        if pattern == "solid":
            _child(p, "fgColor", rgb="FF17365D")
    borders = _child(root, "borders", count=1)
    _child(borders, "border")
    base = _child(root, "cellStyleXfs", count=1)
    _child(base, "xf", numFmtId=0, fontId=0, fillId=0, borderId=0)
    xfs = _child(root, "cellXfs", count=3)
    for font, fill, num in ((0, 0, 0), (1, 2, 0), (0, 0, 10)):
        xf = _child(xfs, "xf", numFmtId=num, fontId=font, fillId=fill,
                    borderId=0, xfId=0, applyAlignment=1, applyNumberFormat=1)
        _child(xf, "alignment", vertical="top", wrapText=1)
    styles = _child(root, "cellStyles", count=1)
    _child(styles, "cellStyle", name="Normal", xfId=0, builtinId=0)
    return root


def _sheet(headers, rows, widths, percentages=()):
    root = _element("worksheet")
    views = _child(root, "sheetViews")
    view = _child(views, "sheetView", workbookViewId=0, showGridLines=0)
    _child(view, "pane", ySplit=1, topLeftCell="A2", activePane="bottomLeft", state="frozen")
    _child(view, "selection", pane="bottomLeft", activeCell="A2", sqref="A2")
    _child(root, "sheetFormatPr", defaultRowHeight=30)
    cols = _child(root, "cols")
    for i, width in enumerate(widths, 1):
        _child(cols, "col", min=i, max=i, width=width, customWidth=1)
    data = _child(root, "sheetData")
    count = 0
    for r, values in enumerate([headers, *rows], 1):
        if r > 1_048_576:
            raise ValueError("Excel report exceeds the worksheet row limit")
        count = r
        height = max((sum(max(1, math.ceil(len(line) / max(1, widths[c] - 2)))
                          for line in str(v or "").split("\n"))
                      for c, v in enumerate(values)), default=1)
        row = _child(data, "row", r=r, ht=min(409, max(30, 15 * height + 6)), customHeight=1)
        for c, value in enumerate(values, 1):
            if value is None:
                continue
            style = 1 if r == 1 else 2 if c in percentages else 0
            cell = _child(row, "c", r=f"{get_column_letter(c)}{r}", s=style)
            if isinstance(value, bool):
                cell.set("t", "b")
                _child(cell, "v").text = str(int(value))
            elif isinstance(value, (int, float)) and math.isfinite(value):
                _child(cell, "v").text = str(value)
            else:
                cell.set("t", "inlineStr")
                text = _child(_child(cell, "is"), "t")
                text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                text.text = _text(value)
    _child(root, "autoFilter", ref=f"A1:{get_column_letter(len(headers))}{count}")
    return root


def _write(path, sheets, source_paths=()):
    dest = Path(path)
    if dest.suffix.lower() != ".xlsx":
        raise ValueError("Excel report path must end in .xlsx")
    if any(dest.resolve() == Path(p).resolve() or
           (dest.exists() and os.path.samefile(dest, p)) for p in source_paths):
        raise ValueError("Excel report cannot overwrite a source workbook")
    if dest.exists() and not is_excel_report(dest):
        raise ValueError(f"Refusing to overwrite a non-report workbook: {dest}")
    dest.parent.mkdir(parents=True, exist_ok=True)
    wb = _element("workbook")
    sheet_list = _child(wb, "sheets")
    relationships = ET.Element(f"{{{REL}}}Relationships", nsmap={None: REL})
    types = ET.Element(f"{{{CONTENT}}}Types", nsmap={None: CONTENT})
    for ext, typ in (("rels", "application/vnd.openxmlformats-package.relationships+xml"),
                     ("xml", "application/xml"), ("txt", "text/plain")):
        ET.SubElement(types, f"{{{CONTENT}}}Default", Extension=ext, ContentType=typ)
    parts = {"xl/styles.xml": _xml(_styles())}
    for i, (name, headers, rows, widths, percentages) in enumerate(sheets, 1):
        el = _child(sheet_list, "sheet", name=name, sheetId=i)
        el.set(f"{{{DOC_REL}}}id", f"rId{i}")
        ET.SubElement(relationships, f"{{{REL}}}Relationship", Id=f"rId{i}",
                      Type=f"{DOC_REL}/worksheet", Target=f"worksheets/sheet{i}.xml")
        parts[f"xl/worksheets/sheet{i}.xml"] = _xml(_sheet(headers, rows, widths, percentages))
        ET.SubElement(types, f"{{{CONTENT}}}Override", PartName=f"/xl/worksheets/sheet{i}.xml",
                      ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    ET.SubElement(relationships, f"{{{REL}}}Relationship", Id="styles",
                  Type=f"{DOC_REL}/styles", Target="styles.xml")
    for name, typ in (("workbook", "sheet.main"), ("styles", "styles")):
        ET.SubElement(types, f"{{{CONTENT}}}Override", PartName=f"/xl/{name}.xml",
                      ContentType=f"application/vnd.openxmlformats-officedocument.spreadsheetml.{typ}+xml")
    root_rels = ET.Element(f"{{{REL}}}Relationships", nsmap={None: REL})
    ET.SubElement(root_rels, f"{{{REL}}}Relationship", Id="rId1",
                  Type=f"{DOC_REL}/officeDocument", Target="xl/workbook.xml")
    parts.update({"[Content_Types].xml": _xml(types), "_rels/.rels": _xml(root_rels),
                  "xl/workbook.xml": _xml(wb), "xl/_rels/workbook.xml.rels": _xml(relationships),
                  MARKER: b"Excel X-ray report v1\n"})
    # Build first and replace atomically so errors leave an existing report intact.
    fd, tmp = tempfile.mkstemp(dir=dest.parent, suffix=".tmp")
    try:
        with os.fdopen(fd, "wb") as fh, zipfile.ZipFile(fh, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, content in parts.items():
                zf.writestr(name, content)
        os.replace(tmp, dest)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    return str(dest)


def write_excel_report(wx, path, assessment):
    """Write a filterable, editable report for one scanned workbook."""
    fr, tr = file_rows(assessment), tab_rows(assessment)
    summary = [[r[k] for k in ("type", "field", "value", "basis", "confidence", "evidence")]
               + [None, None] for r in fr]
    tabs = [[r[k] for k in ("tab", "type", "field", "value", "basis", "confidence", "evidence")]
            + [None, None] for r in tr]
    sheets, regions, formulas, warnings = [], [], [], []
    for s in sorted(wx.sheets, key=lambda s: (s.state != "visible", s.position)):
        sheets.append([s.name, s.state, s.populated_cells, s.max_row, s.max_col,
                       len(s.regions), s.formula_profile.get("total", 0),
                       s.formula_profile.get("distinct_skeletons", 0), s.density])
        for r in s.regions:
            regions.append([s.name, r.ref, r.kind, r.origin, "\n".join(r.headers),
                            r.n_data_rows, r.formula_cells, r.detect_confidence,
                            "\n".join(r.evidence)])
        for shape, count in s.formula_profile.get("top_skeletons", []):
            formulas.append([s.name, shape, count])
        warnings += [[s.name, "Scan note", n] for n in s.notes]
        warnings += [[s.name, "Cached Excel error", e] for e in s.error_cells]
    warnings += [["Workbook", "Scan warning", w] for w in wx.warnings]
    metadata = [["Source file", wx.filename], ["File ID", assessment.file.file_id.value],
                ["Scan Status", assessment.file.scan_status.value],
                ["Scan Error", assessment.file.scan_error.value],
                ["Source modified", wx.fs_modified], ["VBA present", wx.has_vba],
                ["Hidden sheets", f"{sum(s.state != 'visible' for s in wx.sheets)} of {len(wx.sheets)} sheets are hidden"],
                ["Cached errors", sum(len(s.error_cells) for s in wx.sheets)],
                ["Power Query present", wx.has_power_query],
                ["External links", "\n".join(short_source(x) for x in wx.external_links)],
                ["Formula coverage", "Top 15 formula patterns per source sheet; formulas are not recalculated."],
                ["Review", "Use Reviewer value and Reviewer notes to record amendments. Regenerating replaces these edits."],
                ["Interpretation", "Derived scores are heuristics; drafted and inferred text require review."],
                ["Privacy", "Headers and assessment metadata may be sensitive. Local paths and full file hashes are excluded."]]

    opportunity_rows = []
    for kind, fld in (("Simplification", assessment.file.potential_simplification),
                      ("Automation", assessment.file.potential_automation)):
        details = fld.value.get("details", []) if isinstance(fld.value, dict) else []
        for item in details:
            opportunity_rows.append([
                kind, item.get("process"), item.get("sub_process"),
                ", ".join(item.get("worksheets", [])), item.get("observed_evidence"),
                item.get("candidate_action"), item.get("confirmation_required"), fld.basis,
            ])
    return _write(path, [
        ("File assessment", ["Section", "Field", "Value", "Basis", "Confidence", "Evidence", "Reviewer value", "Reviewer notes"], summary, [28, 32, 65, 18, 14, 65, 40, 45], (5,)),
        ("Tab assessments", ["Source tab", "Section", "Field", "Value", "Basis", "Confidence", "Evidence", "Reviewer value", "Reviewer notes"], tabs, [25, 28, 35, 65, 18, 14, 65, 40, 45], (6,)),
        ("Sheet inventory", ["Source tab", "Visibility", "Populated cells", "Rows", "Columns", "Regions", "Formulas", "Distinct patterns", "Density"], sheets, [28, 18, 20, 14, 14, 14, 16, 20, 14], (9,)),
        ("Regions", ["Source tab", "Range", "Kind", "Origin", "Headers", "Data rows", "Formula cells", "Confidence", "Evidence"], regions, [25, 20, 20, 18, 60, 15, 18, 14, 65], (8,)),
        ("Formula patterns", ["Source tab", "Formula pattern (text)", "Occurrences"], formulas, [28, 100, 18], ()),
        ("Warnings", ["Source tab", "Type", "Detail"], warnings, [28, 25, 100], ()),
        ("Report info", ["Field", "Value"], metadata, [28, 100], ()),
        ("Error summary", ["Source tab", "Error type", "Count", "Area", "Region", "Example cells (first 12)", "Potential affected outputs", "Impact certainty", "Recalculation / review"],
         [[r["tab"], r["type"], r["count"], r["area"], r["region"], ", ".join(r["cells"][:12]), ", ".join(r["potential_outputs"]) or "No output path observed", r["impact"], r["recalculation"]] for r in assessment.error_summary],
         [28, 18, 12, 28, 20, 40, 45, 65, 75], ()),
        ("Hidden sheet groups", ["Purpose candidate", "Group size", "Hidden tab", "Potential supported outputs", "Explanation", "Basis"],
         [[r["purpose"], r["count"], name, ", ".join(r["supports"]) or "No output path observed", r["explanation"], r["basis"]] for r in assessment.hidden_groups for name in r["tabs"]],
         [28, 12, 55, 55, 60, 18], ()),
        ("Input sources", ["Group", "Business purpose candidate", "Source", "Reference count", "Consumer tabs", "Relationship", "Directly observed?", "Essentiality / confirmation"],
         [[r["source_type"], r["purpose"], r["source"], r.get("reference_count", 1), ", ".join(r["consumers"]) or "Not resolved", r["association"], "Yes" if r.get("observed") else "No / unresolved", r["essential"]] for r in assessment.input_groups],
         [28, 35, 50, 16, 55, 65, 20, 45], ()),
        ("Calculation steps", ["Source tab", "Observed input tabs", "Principal operation", "Potential business outputs"],
         [[r["tab"], ", ".join(r["inputs"]) or "No cross-tab reference observed", r["operation"], ", ".join(r["potential_outputs"]) or "No output path observed"] for r in assessment.file.key_calculations_logic.value.get("steps", [])],
         [30, 45, 100, 50], ()),
        ("Review opportunities", ["Type", "Process", "Sub-Process", "Relevant worksheets", "Observed evidence", "Candidate action", "Confirmation required", "Basis"],
         opportunity_rows, [22, 34, 34, 40, 55, 70, 75, 18], ()),
    ], [wx.path])


def write_estate_excel_report(estate, path, insight, source_paths=()):
    """Write folder relationships, family membership and recommendations."""
    fps = estate.fingerprints
    pairs = [[fps[i].file_name, fps[j].file_name, c["relationship"],
              c["skeleton"], c["input"], c["output"], c["topology"], c["overall"]]
             for i, j, c in estate.pairs]
    families = []
    for num, (group, item) in enumerate(zip(estate.clusters, insight.families), 1):
        for i in group:
            families.append([num, fps[i].file_name, fps[i].file_id, item.summary,
                             item.recommended_action, item.rationale, item.basis])
    summary = [["Workbooks", len(fps)], ["Related families", len(estate.clusters)],
               ["Unrelated workbooks", len(estate.singletons)],
               ["Summary", insight.estate_summary], ["Basis", insight.basis]]
    summary += [["Opportunity", x] for x in insight.top_opportunities]
    summary.append(["Interpretation", "Similarity is structural, not proof of identical data or financial correctness. Only linked pairs are listed."])
    return _write(path, [
        ("Estate summary", ["Field", "Value"], summary, [28, 100], ()),
        ("Linked pairs", ["Workbook A", "Workbook B", "Relationship", "Formula similarity", "Input similarity", "Output similarity", "Topology similarity", "Overall similarity"], pairs, [32, 32, 35, 22, 20, 20, 22, 22], (4, 5, 6, 7, 8)),
        ("Families", ["Family", "Workbook", "File ID", "Summary", "Recommended action", "Rationale", "Basis"], families, [12, 32, 24, 65, 65, 65, 18], ()),
        ("Unrelated workbooks", ["Workbook", "File ID"], [[fps[i].file_name, fps[i].file_id] for i in estate.singletons], [40, 30], ()),
    ], source_paths)
