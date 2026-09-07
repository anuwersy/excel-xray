"""Standalone HTML + CSV for the estate comparison (Step 7)."""

from __future__ import annotations

import csv
import html

from .estate import REL_COLOR, EstateResult

_CSS = """
*{box-sizing:border-box}
body{margin:0;background:#F7F8FA;color:#10192B;
  font:15px/1.55 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
.wrap{max-width:1180px;margin:0 auto;padding:32px 24px 80px}
.eyebrow{font:600 11px/1 ui-monospace,monospace;letter-spacing:.18em;
  text-transform:uppercase;color:#5B6577;margin-bottom:10px}
h1{margin:0;font-size:30px;letter-spacing:-.02em;font-weight:650}
.vitals{display:flex;flex-wrap:wrap;gap:0;margin:22px 0 8px;border:1px solid #D5DAE3;
  border-radius:3px;overflow:hidden;background:#fff}
.vital{flex:1 1 130px;padding:12px 14px;border-right:1px solid #D5DAE3}
.vital:last-child{border-right:0}
.vital b{display:block;font:600 22px/1.1 ui-monospace,monospace}
.vital span{display:block;font-size:11px;letter-spacing:.08em;text-transform:uppercase;
  color:#5B6577;margin-top:5px}
h2{font-size:17px;margin:30px 0 6px}
.cluster{background:#fff;border:1px solid #D5DAE3;border-radius:4px;margin:14px 0;
  padding:12px 16px}
.cluster h3{margin:0 0 8px;font-size:15px}
.files{font-family:ui-monospace,monospace;font-size:12.5px;color:#374151;margin:0 0 10px}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{text-align:left;font:600 10px/1.3 ui-monospace,monospace;letter-spacing:.06em;
  text-transform:uppercase;color:#5B6577;padding:6px 8px;border-bottom:1px solid #D5DAE3}
td{padding:6px 8px;border-bottom:1px solid #EEF1F5;vertical-align:middle}
.rel{display:inline-block;font:600 10px/1.4 ui-monospace,monospace;padding:2px 7px;
  border-radius:2px;color:#fff;white-space:nowrap}
.bars{display:flex;gap:8px;flex-wrap:wrap;font:11px/1.4 ui-monospace,monospace;color:#5B6577}
.bar{display:inline-flex;align-items:center;gap:4px}
.bar i{display:inline-block;height:7px;border-radius:1px;background:#2D6CA2;min-width:1px}
.bar u{font-style:normal;width:26px;color:#9AA3B2}
.matrix{overflow-x:auto;margin-top:8px}
table.mx td,table.mx th{text-align:center;padding:3px 5px;border:1px solid #EEF1F5;
  font-size:11px}
table.mx td.lbl,table.mx th.lbl{text-align:left;white-space:nowrap;font-weight:600;
  font-family:ui-monospace,monospace}
.legend{display:flex;flex-wrap:wrap;gap:14px;margin:16px 0;font-size:12px;color:#5B6577}
.legend span{display:flex;align-items:center;gap:6px}
.dot{width:10px;height:10px;border-radius:2px;display:inline-block}
footer{margin-top:30px;padding-top:16px;border-top:1px solid #D5DAE3;
  font-size:12.5px;color:#5B6577}
"""


def _esc(s) -> str:
    return html.escape(str(s if s is not None else ""))


def _rel_badge(rel: str) -> str:
    return f"<span class='rel' style='background:{REL_COLOR.get(rel, '#888')}'>{_esc(rel)}</span>"


def _bars(c: dict) -> str:
    out = ["<div class='bars'>"]
    for key in ("skeleton", "input", "output", "topology"):
        v = c[key]
        out.append(f"<span class='bar'>{key[:4]}<i style='width:{int(v*40)}px'></i>"
                   f"<u>{v:.2f}</u></span>")
    return "".join(out) + "</div>"


def build_estate_report(estate: EstateResult) -> str:
    fps = estate.fingerprints
    P = []
    A = P.append
    A(f"<!doctype html><meta charset='utf-8'>"
      f"<meta name='viewport' content='width=device-width,initial-scale=1'>"
      f"<title>Estate comparison</title><style>{_CSS}</style><div class='wrap'>")
    A("<div class='eyebrow'>Excel X-ray &middot; estate comparison</div>"
      "<h1>EUC estate comparison</h1>")

    A("<div class='vitals'>")
    for val, label in [
        (len(fps), "workbooks"),
        (len(estate.clusters), "clusters"),
        (len(estate.pairs), "linked pairs"),
        (sum(1 for _i, _j, c in estate.pairs if c["relationship"] == "Duplicate"),
         "duplicates"),
        (len(estate.singletons), "unrelated"),
    ]:
        A(f"<div class='vital'><b>{val}</b><span>{label}</span></div>")
    A("</div>")

    A("<div class='legend'>")
    for rel, col in REL_COLOR.items():
        if rel == "Unrelated":
            continue
        A(f"<span><i class='dot' style='background:{col}'></i>{_esc(rel)}</span>")
    A("</div>")

    # ---- Clusters -------------------------------------------------------
    if estate.clusters:
        A("<h2>Related families</h2>")
        for gi, group in enumerate(estate.clusters, 1):
            names = [fps[i].file_name for i in group]
            A(f"<div class='cluster'><h3>Family {gi} &middot; {len(group)} workbooks</h3>"
              f"<div class='files'>{_esc('  •  '.join(names))}</div>")
            inside = [(i, j, c) for i, j, c in estate.pairs
                      if i in group and j in group]
            A("<table><tr><th>Workbook A</th><th>Workbook B</th>"
              "<th>Relationship</th><th>Overall</th><th>Signals</th></tr>")
            for i, j, c in inside:
                A(f"<tr><td>{_esc(fps[i].file_name)}</td>"
                  f"<td>{_esc(fps[j].file_name)}</td>"
                  f"<td>{_rel_badge(c['relationship'])}</td>"
                  f"<td>{c['overall']:.2f}</td>"
                  f"<td>{_bars(c)}</td></tr>")
            A("</table></div>")
    else:
        A("<h2>Related families</h2><p>No workbooks are linked above threshold.</p>")

    # ---- Similarity matrix (only when small enough to read) --------------
    n = len(fps)
    if 2 <= n <= 40:
        best = {}
        for i, j, c in estate.pairs:
            best[(i, j)] = best[(j, i)] = c["overall"]
        A("<h2>Similarity matrix</h2><div class='matrix'><table class='mx'><tr><th class='lbl'></th>")
        for k in range(n):
            A(f"<th>{k+1}</th>")
        A("</tr>")
        for i in range(n):
            A(f"<tr><td class='lbl'>{i+1}. {_esc(fps[i].file_name)}</td>")
            for j in range(n):
                if i == j:
                    A("<td style='background:#10192B;color:#fff'>–</td>")
                else:
                    v = best.get((i, j), 0.0)
                    shade = f"rgba(45,108,162,{min(v,1.0):.2f})" if v else "#fff"
                    txt = f"{v:.2f}" if v else ""
                    A(f"<td style='background:{shade}'>{txt}</td>")
            A("</tr>")
        A("</table></div>")

    if estate.singletons:
        names = ", ".join(fps[i].file_name for i in estate.singletons)
        A(f"<h2>Unrelated workbooks</h2><div class='files'>{_esc(names)}</div>")

    A("<footer>Similarity is computed from formula shapes, input signatures, "
      "output signatures and dependency topology — never cell values. Relationship "
      "labels are heuristic; treat them as leads for a reviewer, not conclusions."
      "</footer></div>")
    return "".join(P)


def write_estate_report(estate: EstateResult, path: str) -> str:
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(build_estate_report(estate))
    return path


def write_estate_csv(estate: EstateResult, path: str) -> str:
    fps = estate.fingerprints
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["Workbook A", "Workbook B", "Relationship", "Overall",
                    "Formula shapes", "Inputs", "Outputs", "Topology"])
        for i, j, c in estate.pairs:
            w.writerow([fps[i].file_name, fps[j].file_name, c["relationship"],
                        c["overall"], c["skeleton"], c["input"], c["output"],
                        c["topology"]])
    return path
