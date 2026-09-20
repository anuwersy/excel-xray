#!/usr/bin/env python3
"""Excel X-ray command line.

  excel-xray FILE.xlsx                 Excel report next to the file
  excel-xray FOLDER -o out/            every workbook in a folder
  excel-xray FILE.xlsx --json          machine-readable JSON to stdout

Reads only. Never writes to, moves or renames a source file.
"""

from __future__ import annotations

import argparse
import hashlib
import os
import sys
import time
import traceback

import json

from .assessment import assess
from .assessment import to_dict as assessment_to_dict
from .report import write_report
from .excel_report import is_excel_report, write_excel_report, write_estate_excel_report
from .scan import UnreadableWorkbook, failed_workbook, to_json, xray_workbook

EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
SKIP_PREFIX = ("~$", ".")


def _narrative_assessor(args):
    """Build the --llm narrative assessor for the chosen --provider, or None."""
    if not args.llm:
        return None
    max_tokens = int(os.environ.get("LLM_MAX_TOKENS") or 2000)
    if args.provider == "openai":
        from .narrative import OpenAIAssessor
        return OpenAIAssessor(model=args.model, max_tokens=max_tokens,
                               azure_endpoint=args.azure_endpoint,
                               api_version=args.api_version)
    from .narrative import ClaudeAssessor
    return ClaudeAssessor(model=args.model, max_tokens=max_tokens)


def _estate_assessor(args):
    """Build the --llm estate-insight assessor for the chosen --provider, or None."""
    if not args.llm:
        return None
    if args.provider == "openai":
        from .estate_insight import OpenAIEstateAssessor
        return OpenAIEstateAssessor(model=args.model,
                                     azure_endpoint=args.azure_endpoint,
                                     api_version=args.api_version)
    from .estate_insight import ClaudeEstateAssessor
    return ClaudeEstateAssessor(model=args.model)


def collect(target: str) -> list[str]:
    if os.path.isfile(target):
        return [] if is_excel_report(target) else [target]
    out = []
    for root, _dirs, files in os.walk(target):
        for f in files:
            if f.startswith(SKIP_PREFIX):
                continue
            if os.path.splitext(f)[1].lower() in EXTS:
                path = os.path.join(root, f)
                if not is_excel_report(path):
                    out.append(path)
    return sorted(out)


def main() -> int:
    ap = argparse.ArgumentParser(
        prog="excel-xray", description="Excel X-ray structural scanner"
    )
    ap.add_argument("target", help="workbook or folder")
    ap.add_argument("-o", "--out", default=None, help="output directory")
    ap.add_argument("--format", choices=("xlsx", "html"), default="xlsx",
                    help="report format (default: xlsx)")
    output_mode = ap.add_mutually_exclusive_group()
    output_mode.add_argument("--json", action="store_true", help="emit scan JSON to stdout")
    output_mode.add_argument("--assess", action="store_true",
                    help="emit the EUC assessment JSON to stdout")
    ap.add_argument("--llm", action="store_true",
                    help="use a model assessor for narrative fields "
                         "(needs a credential for the chosen --provider)")
    ap.add_argument("--provider", choices=["claude", "openai"], default=None,
                    help="LLM backend for --llm: 'claude' (Anthropic, needs the "
                         "'anthropic' package) or 'openai' (GPT via the public "
                         "OpenAI API, or an Azure OpenAI deployment if "
                         "--azure-endpoint/$AZURE_OPENAI_ENDPOINT is set; needs "
                         "the 'openai' package). Default: claude, unless an "
                         "Azure endpoint is given, which implies openai.")
    ap.add_argument("--model", default=None,
                    help="model id for --llm. Claude default: $ANTHROPIC_MODEL "
                         "or claude-opus-5. OpenAI default: $OPENAI_MODEL or "
                         "gpt-4o; on Azure this is the deployment name "
                         "(else $AZURE_OPENAI_DEPLOYMENT).")
    ap.add_argument("--azure-endpoint", default=None,
                    help="Azure OpenAI resource endpoint, e.g. "
                         "https://<resource>.openai.azure.com "
                         "(else $AZURE_OPENAI_ENDPOINT). Implies --provider openai.")
    ap.add_argument("--api-version", default=None,
                    help="Azure OpenAI api-version (else $AZURE_OPENAI_API_VERSION, "
                         "default 2024-10-21)")
    ap.add_argument("--csv", default=None, metavar="PATH",
                    help="also write the EUC assessment as a CSV table")
    ap.add_argument("--estate", action="store_true",
                    help="compare workbooks across the folder: write an estate "
                         "report (estate.xlsx by default) to the out dir")
    ap.add_argument("--max-rows", type=int, default=200_000)
    args = ap.parse_args()
    if args.json and (args.csv or args.estate):
        ap.error("--json cannot be combined with --csv or --estate; use --assess instead")

    if args.llm:
        # Load a local .env so ANTHROPIC_API_KEY / OPENAI_API_KEY /
        # AZURE_OPENAI_* can live there. Best-effort: python-dotenv ships
        # with the [llm] extra, so this is a no-op otherwise.
        try:
            from dotenv import load_dotenv
            load_dotenv()
        except ImportError:
            pass

    args.azure_endpoint = args.azure_endpoint or os.environ.get("AZURE_OPENAI_ENDPOINT")
    args.api_version = args.api_version or os.environ.get("AZURE_OPENAI_API_VERSION")
    # An Azure endpoint only makes sense for the openai provider.
    args.provider = args.provider or ("openai" if args.azure_endpoint else "claude")

    # Resolve model / token budget: explicit flag wins, then the environment
    # (loaded from .env above), then a per-provider default.
    if args.provider == "openai":
        args.model = (args.model or os.environ.get("AZURE_OPENAI_DEPLOYMENT")
                      or os.environ.get("OPENAI_MODEL") or "gpt-4o")
    else:
        args.model = args.model or os.environ.get("ANTHROPIC_MODEL") or "claude-opus-5"

    paths = collect(args.target)
    if not paths:
        print(f"no workbooks found under {args.target}", file=sys.stderr)
        return 2

    outdir = args.out or (args.target if os.path.isdir(args.target)
                          else os.path.dirname(os.path.abspath(args.target)))
    os.makedirs(outdir, exist_ok=True)

    ok = partial = failed = 0
    reasons: dict[str, list[str]] = {}
    batch: list = []  # (path, wx) for assessment-aware output modes
    timings: dict[str, float] = {}  # path -> extraction seconds, for the report line below
    for p in paths:
        t0 = time.perf_counter()
        try:
            wx = xray_workbook(p, max_rows=args.max_rows)
        except UnreadableWorkbook as e:
            elapsed = time.perf_counter() - t0
            failed += 1
            reasons.setdefault(e.category, []).append(os.path.basename(p))
            print(f"SKIPPED  {os.path.basename(p):46} {e}  ({elapsed:.2f}s)", file=sys.stderr)
            wx = failed_workbook(p, str(e))
            timings[p] = elapsed
            if args.json:
                print(to_json(wx))
            else:
                batch.append((p, wx))
            continue
        except Exception as e:  # noqa: BLE001 - report and keep going over a corpus
            elapsed = time.perf_counter() - t0
            failed += 1
            reasons.setdefault("unexpected", []).append(os.path.basename(p))
            print(f"FAILED   {os.path.basename(p):46} {type(e).__name__}: {e}  "
                  f"({elapsed:.2f}s)", file=sys.stderr)
            if os.environ.get("XRAY_DEBUG"):
                traceback.print_exc()
            wx = failed_workbook(p, f"{type(e).__name__}: {e}")
            timings[p] = elapsed
            if args.json:
                print(to_json(wx))
            else:
                batch.append((p, wx))
            continue
        elapsed = time.perf_counter() - t0
        timings[p] = elapsed

        ok += wx.parse_status == "full"
        partial += wx.parse_status == "partial"
        print(f"EXTRACT  {os.path.basename(p):46} {elapsed:6.2f}s", file=sys.stderr)
        if args.json:  # raw scan, no assessment
            print(to_json(wx))
        else:
            batch.append((p, wx))

    # Assessment-aware modes (Excel/HTML, --assess, --csv). One corpus pass so
    # duplication/consolidation see the whole set.
    assessments = None
    if batch and (args.assess or args.csv or args.estate or not args.json):
        assessor = _narrative_assessor(args)
        wxs = [wx for _, wx in batch]
        if len(wxs) > 1:
            from .corpus import assess_corpus
            assessments = assess_corpus(wxs, assessor)
        else:
            assessments = [assess(wxs[0], assessor)]

    if args.assess and assessments is not None:
        payload = [assessment_to_dict(a) for a in assessments]
        print(json.dumps(payload if len(payload) != 1 else payload[0],
                         indent=2, default=str))
    elif not args.json and assessments is not None:
        for (p, wx), a in zip(batch, assessments):
            name = os.path.splitext(os.path.basename(p))[0]
            # Recursive scans may contain several workbooks with the same stem.
            if sum(os.path.splitext(os.path.basename(q))[0].casefold() == name.casefold()
                   for q in paths) > 1:
                name += "_" + hashlib.sha256(os.path.abspath(p).encode()).hexdigest()[:12]
            dest = os.path.join(outdir, f"xray_{name}.{args.format}")
            if args.format == "xlsx":
                if any(os.path.realpath(dest) == os.path.realpath(q) for q in paths):
                    ap.error(f"report destination is a source workbook: {dest}")
                write_excel_report(wx, dest, a)
            else:
                write_report(wx, dest, a)
            regions = sum(len(s.regions) for s in wx.sheets)
            low = sum(1 for s in wx.sheets for r in s.regions
                      if r.detect_confidence < 0.70)
            print(f"{wx.parse_status:8} {os.path.basename(p):46} "
                  f"{len(wx.sheets):3} sheets  {regions:3} regions  "
                  f"{low:2} low-conf  {timings.get(p, 0):5.2f}s  -> {dest}")

    if args.csv and assessments is not None:
        from .tabular import to_csv
        named = [(wx.filename, a) for (_, wx), a in zip(batch, assessments)]
        to_csv(named, args.csv)
        print(f"wrote {args.csv} ({len(named)} workbook(s))", file=sys.stderr)

    if args.estate and assessments is not None:
        if len(assessments) < 2:
            print("--estate needs more than one workbook to compare", file=sys.stderr)
        else:
            from .estate import build_estate
            from .estate_insight import generate_estate_insight
            from .estate_report import write_estate_csv, write_estate_report
            pairs_in = [(wx, a) for (_, wx), a in zip(batch, assessments)
                        if wx.parse_status != "failed"]
            if len(pairs_in) < 2:
                print("--estate needs at least two successfully read workbooks to compare",
                      file=sys.stderr)
            else:
                estate = build_estate(pairs_in)
                insight_assessor = _estate_assessor(args)
                insight = generate_estate_insight(estate, insight_assessor)
                html_path = os.path.join(outdir, "estate.html")
                csv_path = os.path.join(outdir, "estate_pairs.csv")
                if args.format == "xlsx":
                    report_path = os.path.join(outdir, "estate.xlsx")
                    write_estate_excel_report(estate, report_path, insight,
                                              [wx.path for wx, _ in pairs_in])
                else:
                    report_path = html_path
                    write_estate_report(estate, html_path, insight)
                    write_estate_csv(estate, csv_path)
                print(f"{len(estate.fingerprints)} workbooks  "
                      f"{len(estate.clusters)} families  {len(estate.pairs)} linked pairs"
                      f"  -> {report_path}", file=sys.stderr)

    total = len(paths)
    print(f"\ncoverage: {ok}/{total} full, {partial} partial, {failed} unreadable",
          file=sys.stderr)
    for cat, files in sorted(reasons.items()):
        print(f"  {cat:12} {len(files):3}  {', '.join(files[:4])}"
              + (" ..." if len(files) > 4 else ""), file=sys.stderr)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
