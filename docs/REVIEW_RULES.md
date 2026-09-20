# Business-review assessment rules

These rules implement the comments in the three supplied review screenshots.
They improve deterministic interpretation and the optional narrative prompt;
they do not train or fine-tune a model. The screenshot's sample counts and
business facts are not hardcoded into the scanner.

## Header and error review

- A successful assessment reports **Scan Status = full** and **Scan Error =
  N/A**. Partial and failed scans retain a concise explanation of unread or
  failed content. Cached Excel formula errors remain a separate diagnostic.
- Report hidden sheets as **N of M sheets are hidden**, with purpose groups,
  member tabs and potentially supported outputs. Hidden status does not imply
  obsolescence, weak controls or retirement suitability.
- Summarise cached errors by type, tab, containing region and area. Identify
  potentially affected output tabs through transitive worksheet dependencies.
  Keep individual locations in Warnings and JSON for traceability.
- An error in a source tab does not establish that every downstream output cell
  is incorrect. Dependency reachability is a possible exposure, not a proven
  cell-level impact. Recalculate in Excel with refreshed sources to confirm
  whether cached errors persist and affect the deliverable.
- A row-capped scan is labelled partial. Missing/dynamic references, defined
  names, macro dependencies and other unsupported formula constructs may leave
  gaps in the dependency graph.

## Tab-level review

- Add Tab Visibility. Analyse and present visible tabs first; retain the
  original order within visible and hidden groups. Very-hidden status remains
  distinguishable in the assessment and inventory.
- Keep a primary category for compatibility with estate comparisons and add
  All Detected Tab Roles. Explicit mapping/source labels take precedence over
  generic terminal-sheet signals. Report/summary tabs can remain outputs even
  when other tabs use them. Formula-heavy reporting tabs can have both Output
  and Calculation roles.
- COA / chart of accounts is a mapping input, not a final deliverable. A sheet
  with no downstream references is not, by itself, a business output.
- Describe aggregation, lookup/matching, data preparation and conditional logic
  in plain language. Formula workload is Light, Moderate (>200 formulas or >15
  distinct shapes), or Heavy (>5,000 formulas or >60 shapes). These thresholds
  describe technical workload, not a validated financial-risk rating.
- Split upstream dependencies into observed in-workbook references and external
  filenames. Resolve indexed external links in workbook external-reference
  order. Unresolved sources and connection/pivot consumers are explicitly
  unknown rather than assigned to every sheet. Full paths remain only in the
  underlying raw scan metadata where originally available.
- Ask targeted questions about persistent errors, source freshness,
  recalculation, potentially significant fixed constants and ambiguous roles.
  Workbook-wide VBA, hidden status, isolated weak region detections or stored
  cells do not automatically flag every tab for human validation.

## File-level review

- **Business Area / Process** consolidates high-level functional tab roles.
  **Process** describes the broader end-to-end business process and
  **Sub-Process** the intermediate outcome/activity. If structural evidence is
  weak, both use **Not established — owner confirmation required**.
- The client-facing summary contains one deduplicated **Logic Types** field.
  Its taxonomy is Calculation, Reporting, Data Transformation,
  Reconciliation / Control, Manual Input and Other. A primary logic type may
  remain internal for compatibility.
- Group inputs into In-workbook tabs, External workbooks, Formal data
  connections, Pivot sources and Unresolved sources. Within those groups,
  classify candidate business purpose: trial balance, claims, FX,
  prior-period data, adjustments, mapping, policy/premium data, assumptions,
  tax, reserves and financial reporting. Label inferred purposes as derived;
  list providing files/tabs and consumers where observed. Decode and shorten
  source names, deduplicate them and retain reference counts. Essentiality
  needs owner confirmation. A `tb` label means Trial balance input; it does not
  establish the source system.
- Separate final deliverable candidates from inputs and intermediate workings.
  Do not promote every terminal working tab to a final output.
- Tie simplification and automation candidates to a named tab, its operation
  and observed inputs. Error remediation is separate from simplification.
  Stored values could be imports, labels, pasted data or manual entry; they do
  not demonstrate repetitive manual work. Ask about the actual process step,
  recurrence, exceptions and approvals before claiming automation savings.
- Explain principal calculation steps in plain business language. Function
  frequencies and formula skeletons stay in the technical appendix and are not
  repeated in the client-facing calculation summary.
- A confirmed reconciliation requires an observed comparison of at least two
  sources plus agreement, difference or exception evidence. Lookup, IF,
  variance and roll-forward formulas alone are not reconciliations. Ambiguous
  signals remain candidates requiring review and are excluded from the
  confirmed count. Each confirmed item identifies both sources, matching
  criteria, tolerance, exception logic, evidence and confidence; unsupported
  details remain **Not established**.
- Retain Manual Intervention as **needs_human**, rather than a misleading
  High/Medium/Low score based on non-formula cell counts.
- Report VBA, Power Query, formal connections and external workbook links as
  separate mechanisms. The scanner does not inspect
  VBA code or execute macros, so business use cases, triggers and affected
  outputs require confirmation. Formal connections may be absent while
  external workbook formula links are present.
- Usage frequency, deadline and recipient remain **needs_human** with specific
  reviewer questions and the existing editable Excel reviewer columns. More
  sample workbooks alone cannot establish these operational facts.
- Duplication/consolidation still use deterministic folder-wide comparison.
  State comparison coverage and the structural nature of matches; one file
  leaves these fields **needs_corpus**, not a false negative. Candidate
  duplication does not authorise retiring files without business validation.

## Report and API changes

Excel and HTML include **Error summary**, **Hidden sheet groups**, **Input
sources**, **Calculation steps** and **Review opportunities**. Long file-level
lists are summarised with explicit pointers to these full detail tables.
JSON assessments also expose
`error_summary`, `hidden_groups` and `input_groups`.

The tab schema adds `tab_visibility` and `tab_roles`; the file schema includes
`scan_status`, `scan_error`, `sheet_count_total`, `sheet_count_hidden`,
`process`, `sub_process` and `logic_types`. Client-facing reports omit file
size, full SHA-256 and raw source path; the short File ID remains. Raw scan JSON
continues to retain technical metadata for diagnostics. `upstream_dependencies.value` is an object with
`within_workbook`, `external_files`, `unresolved_external` and `scope`, replacing
the old flat list. Review verdicts deliberately use candidate/unknown language
where the previous version overstated certainty.

Validation uses the repository fixtures and additional regression cases for
mixed-role tabs, hidden support logic, transitive error exposure, external-link
ordering, partial scans and unsupported manual-effort claims. The underlying
74-tab management-account workbook was not supplied; the screenshots alone
cannot establish accuracy on that workbook or on a broader production corpus.
