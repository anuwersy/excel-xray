"""Excel X-ray — structural scanner and explainer for financial workbooks.

Reads structure straight from the OOXML package (openpyxl silently drops merged
ranges, tables, connections, external links and pivot caches) and cells in a
single ``lxml`` pass (formula and cached value together). Detects hand-built
table regions, normalises formulas to counted skeletons, and explains the
result as a self-contained HTML report or JSON.
"""

from __future__ import annotations

from .assessment import Assessment, FileAssessment, TabAssessment, assess
from .corpus import Fingerprint, assess_corpus, fingerprint, similarity
from .estate import EstateResult, build_estate, compare, fingerprint_euc
from .estate_insight import (
    ClaudeEstateAssessor,
    EstateInsight,
    OfflineEstateAssessor,
    generate_estate_insight,
)
from .estate_report import build_estate_report, write_estate_csv, write_estate_report
from .narrative import ClaudeAssessor, Narrative, OfflineAssessor
from .ooxml import WorkbookStructure, read_structure
from .report import build_report, write_report
from .scan import (
    SheetXray,
    UnreadableWorkbook,
    WorkbookXray,
    to_json,
    triage,
    xray_workbook,
)

__all__ = [
    "xray_workbook",
    "to_json",
    "triage",
    "WorkbookXray",
    "SheetXray",
    "UnreadableWorkbook",
    "read_structure",
    "WorkbookStructure",
    "build_report",
    "write_report",
    "assess",
    "Assessment",
    "FileAssessment",
    "TabAssessment",
    "OfflineAssessor",
    "ClaudeAssessor",
    "Narrative",
    "assess_corpus",
    "fingerprint",
    "similarity",
    "Fingerprint",
    "build_estate",
    "compare",
    "fingerprint_euc",
    "EstateResult",
    "build_estate_report",
    "write_estate_report",
    "write_estate_csv",
    "generate_estate_insight",
    "EstateInsight",
    "OfflineEstateAssessor",
    "ClaudeEstateAssessor",
    "main",
]

__version__ = "0.1.0"


def main() -> None:
    """Console-script entry point."""
    from .cli import main as _main

    raise SystemExit(_main())
