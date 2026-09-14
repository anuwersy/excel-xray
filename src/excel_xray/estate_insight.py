"""Estate insight layer (Step 8): interpret the deterministic comparison.

The similarity scores, relationship types and clusters from :mod:`excel_xray.estate`
stay deterministic — reproducible and auditable. This layer sits *on top* and adds
the judgement the numbers can't express: what each family of related EUCs is, and
what to do about it (consolidate / keep one / extract shared logic / align source).

Same shape as the file-level narrative: an ``EstateAssessor`` interface with an
offline template (basis ``drafted``) and opt-in model-backed implementations —
Claude, or GPT via the public OpenAI API / an Azure OpenAI deployment (basis
``inferred``). Only fingerprint metadata — file names, headers, relationship
types, scores — is sent to the model; never cell values.
"""

from __future__ import annotations

import json
import os
from collections import Counter
from dataclasses import dataclass, field

from .estate import EstateResult


@dataclass
class FamilyInsight:
    summary: str = ""
    recommended_action: str = ""
    rationale: str = ""
    basis: str = "drafted"


@dataclass
class EstateInsight:
    families: list = field(default_factory=list)  # aligned with estate.clusters
    estate_summary: str = ""
    top_opportunities: list = field(default_factory=list)
    basis: str = "drafted"


# --------------------------------------------------------------- evidence bundle


def build_estate_bundle(estate: EstateResult) -> dict:
    """Value-free evidence for the insight step, one entry per family."""
    fps = estate.fingerprints
    families = []
    for group in estate.clusters:
        pairs = [c for i, j, c in estate.pairs if i in group and j in group]
        rels = Counter(c["relationship"] for c in pairs)

        def avg(k):
            return round(sum(c[k] for c in pairs) / len(pairs), 2) if pairs else 0.0

        fam = [fps[i] for i in group]
        shared_out = set.intersection(*[f.output_sig for f in fam]) if fam else set()
        shared_in = set.intersection(*[f.input_sig for f in fam]) if fam else set()
        families.append({
            "files": [fps[i].file_name for i in group],
            "logic_types": sorted({fps[i].logic_type for i in group}),
            "relationships": dict(rels),
            "dominant_relationship": rels.most_common(1)[0][0] if rels else "Related",
            "avg_overall": avg("overall"),
            "avg_skeleton": avg("skeleton"),
            "avg_output": avg("output"),
            "avg_input": avg("input"),
            "shared_outputs": sorted(shared_out)[:12],
            "shared_inputs": sorted(shared_in)[:12],
        })
    return {
        "workbook_count": len(fps),
        "family_count": len(estate.clusters),
        "unrelated_count": len(estate.singletons),
        "families": families,
    }


# --------------------------------------------------------------- offline assessor

_ACTION = {
    "Duplicate": "Consolidate to one master and retire the duplicates.",
    "Same output, different method":
        "Consolidate — the same deliverable is produced by different logic.",
    "Overlapping logic":
        "Extract the shared calculation into one reusable component.",
    "Shared source":
        "Align on a single source extract; consider one shared control.",
    "Related": "Review together — related activity.",
}


class OfflineEstateAssessor:
    """Templated insight from the deterministic family evidence. No network."""

    basis = "drafted"
    label = "offline template"

    def insight(self, bundle: dict) -> EstateInsight:
        families = []
        for fam in bundle["families"]:
            n = len(fam["files"])
            rel = fam["dominant_relationship"]
            logic = ", ".join(fam["logic_types"]).lower()
            summary = f"{n} {logic} workbooks with {rel.lower()} characteristics"
            if fam["shared_outputs"]:
                summary += "; shared outputs include " + ", ".join(fam["shared_outputs"][:5])
            elif fam["shared_inputs"]:
                summary += "; shared inputs include " + ", ".join(fam["shared_inputs"][:5])
            summary += "."
            rationale = (f"average across the family — formula shapes "
                         f"{fam['avg_skeleton']}, outputs {fam['avg_output']}, "
                         f"inputs {fam['avg_input']}.")
            families.append(FamilyInsight(
                summary=summary,
                recommended_action=_ACTION.get(rel, "Review together."),
                rationale=rationale, basis="drafted",
            ))

        estate_summary = (
            f"{bundle['workbook_count']} workbooks form {bundle['family_count']} "
            f"related family(ies); {bundle['unrelated_count']} are unrelated."
        )
        ranked = sorted(
            zip(bundle["families"], families),
            key=lambda fp: len(fp[0]["files"]) * fp[0]["avg_overall"], reverse=True,
        )
        top = []
        for fam, ins in ranked[:3]:
            lead = fam["files"][0]
            more = f" + {len(fam['files']) - 1} more" if len(fam["files"]) > 1 else ""
            top.append(f"{lead}{more}: {ins.recommended_action}")

        return EstateInsight(families=families, estate_summary=estate_summary,
                             top_opportunities=top, basis="drafted")


# --------------------------------------------------------------- claude assessor

_SYSTEM = (
    "You advise a financial-controls team reviewing a portfolio (estate) of "
    "End-User Computing spreadsheets. You are given a value-free summary of "
    "families of similar workbooks — file names, shared column headers, the "
    "deterministic relationship types and similarity scores between them. For each "
    "family, explain in one or two sentences what it appears to be and how the "
    "workbooks relate, and give a concrete recommended action (consolidate, keep "
    "one and retire the rest, extract shared logic, or align the data source). Do "
    "not invent figures, owners or systems. Reply with a single JSON object."
)

_INSTRUCTION = (
    "Return JSON with these keys:\n"
    '  "estate_summary": string — one or two sentences on the estate overall.\n'
    '  "top_opportunities": array of strings — the highest-value actions, ranked.\n'
    '  "families": array, one object per family in the same order as the input, '
    'each {"summary": string, "recommended_action": string, "rationale": string}.\n'
    "Evidence:\n"
)


class ClaudeEstateAssessor:
    basis = "inferred"

    def __init__(self, model: str = "claude-opus-5", max_tokens: int = 2000):
        self.model = model
        self.max_tokens = max_tokens
        self.label = f"Claude ({model})"

    def insight(self, bundle: dict) -> EstateInsight:
        try:
            import anthropic
        except ImportError as e:
            raise RuntimeError(
                "the Claude estate assessor needs the 'anthropic' package — "
                "install it with: uv add --optional llm anthropic"
            ) from e
        from .narrative import _parse_json

        client = anthropic.Anthropic()
        msg = client.messages.create(
            model=self.model, max_tokens=self.max_tokens, system=_SYSTEM,
            messages=[{"role": "user",
                       "content": _INSTRUCTION + json.dumps(bundle, default=str)}],
        )
        text = "".join(b.text for b in msg.content if b.type == "text")
        data = _parse_json(text)
        fam_out = data.get("families") or []
        families = []
        for i in range(len(bundle["families"])):
            f = fam_out[i] if i < len(fam_out) else {}
            families.append(FamilyInsight(
                summary=f.get("summary", ""),
                recommended_action=f.get("recommended_action", ""),
                rationale=f.get("rationale", ""), basis="inferred",
            ))
        return EstateInsight(
            families=families,
            estate_summary=data.get("estate_summary", ""),
            top_opportunities=data.get("top_opportunities") or [],
            basis="inferred",
        )


# --------------------------------------------------------- openai / azure assessor


class OpenAIEstateAssessor:
    """Opt-in GPT-backed estate insight. Same client selection as
    :class:`narrative.OpenAIAssessor` — public OpenAI API by default, or an
    Azure OpenAI deployment when given (or found via env) an endpoint."""

    basis = "inferred"

    def __init__(self, model: str = "gpt-4o", max_tokens: int = 2000,
                 azure_endpoint: str | None = None, api_version: str | None = None):
        self.model = model
        self.max_tokens = max_tokens
        self.azure_endpoint = azure_endpoint or os.environ.get("AZURE_OPENAI_ENDPOINT")
        self.api_version = api_version or os.environ.get("AZURE_OPENAI_API_VERSION")
        self.label = (f"Azure OpenAI ({model})" if self.azure_endpoint
                      else f"OpenAI ({model})")

    def insight(self, bundle: dict) -> EstateInsight:
        from .narrative import _openai_complete
        data = _openai_complete(
            self.model, self.max_tokens, _SYSTEM,
            _INSTRUCTION + json.dumps(bundle, default=str),
            azure_endpoint=self.azure_endpoint, api_version=self.api_version,
        )
        fam_out = data.get("families") or []
        families = []
        for i in range(len(bundle["families"])):
            f = fam_out[i] if i < len(fam_out) else {}
            families.append(FamilyInsight(
                summary=f.get("summary", ""),
                recommended_action=f.get("recommended_action", ""),
                rationale=f.get("rationale", ""), basis="inferred",
            ))
        return EstateInsight(
            families=families,
            estate_summary=data.get("estate_summary", ""),
            top_opportunities=data.get("top_opportunities") or [],
            basis="inferred",
        )


def generate_estate_insight(estate: EstateResult, assessor=None) -> EstateInsight:
    """Interpret ``estate``. Offline template by default; pass a ClaudeEstateAssessor
    for model-written insight."""
    assessor = assessor or OfflineEstateAssessor()
    return assessor.insight(build_estate_bundle(estate))
