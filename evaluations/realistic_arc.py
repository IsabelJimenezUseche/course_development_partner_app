"""Shared harness for realistic professor workflows.

Two ideas hold this together.

**Faculty do not arrive with a specification.** In Co-design they arrive vague,
answer tersely, lose a room, push back on something too advanced, spot one bad
question, and then want something to hand a TA. An arc that sends nine tidy
requirements is not testing the product anyone will actually use.

**Auto is a different job, not a quieter one.** A professor who selects Auto is
delegating: they write one thorough brief with the constraints already in it,
walk away, and come back to review. So Auto gets its own short arc that front-loads
everything Co-design reveals over time. Same deliverables demanded, different
interaction shape. Running the identical nine turns in both modes would measure
neither honestly.

Hard pass/fail is reserved for things the app controls — deliverables exist in the
right format, mode discipline holds, professor-facing text is not an internal stub.
Judgments about teaching quality are *recorded* as observations, because substring
proxies for pedagogy fail as phrasing varies, and a suite that chases them drifts
green while measuring less.
"""

from __future__ import annotations

import json
import re
import statistics
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

from fastapi.testclient import TestClient

RESULTS = Path(__file__).resolve().parent / "results"
MODES = ("Co-design", "Auto")

# Language that belongs to the machinery, never in front of a professor.
INTERNAL_TERMS = (
    "skill.md",
    "state_file",
    "artifact_spec",
    "json schema",
    "validator command",
    "project-state filename",
    "tool call",
)
# A reply that is only one of these is a receipt, not an answer.
STUB_PATTERNS = (
    r"^prepared\s+\*\*.+\*\*\s+for artifact generation\.?$",
    r"^artifact specification[\s–—-]",
    r"^saved `[^`]+` to the project's design state.*$",
    r"^updated the project's design state:?$",
)
PROFESSOR_QUESTION_PATTERNS = (
    "would you like",
    "which approach",
    "please provide",
    "what additional information",
    "could you clarify",
    "let me know",
)


@dataclass(frozen=True)
class Probe:
    """A recorded observation. Never a pass/fail assertion."""

    name: str
    pattern: str
    from_turn: int = 0
    expect_absent: bool = False


@dataclass(frozen=True)
class Scenario:
    name: str
    project: dict
    codesign_arc: tuple[tuple[str, str], ...]
    auto_arc: tuple[tuple[str, str], ...]
    probes: tuple[Probe, ...] = field(default_factory=tuple)
    min_office_files: int = 2
    required_formats: tuple[str, ...] = ("docx",)
    # Phrases the assistant must never produce for this domain. Asserting the absence
    # of a dangerous claim is a sound hard check: it fails only when something genuinely
    # bad happened, unlike a presence check on preferred phrasing.
    forbidden_claims: tuple[str, ...] = field(default_factory=tuple)


def is_stub(text: str) -> bool:
    stripped = text.strip().lower()
    return any(re.match(p, stripped, re.IGNORECASE) for p in STUB_PATTERNS)


# A question mark is not a question. Counting glyphs scored a reply that asked three
# numbered questions as five, because one question was compound and another carried a
# "?" inside a parenthetical example. Count the things a professor would count.
_PARENTHETICAL = re.compile(r"\([^)]*\)")
_QUOTED = re.compile(r"[\"“”][^\"“”]*[\"“”]|'[^']{4,}'")


def count_questions(text: str) -> int:
    """Count questions actually put to the educator.

    One line that ends in `?` is one question however many clauses it joins.
    Examples in parentheses or quotes are illustrations of an answer, not new asks.
    """
    if not text:
        return 0
    asked = 0
    for line in text.splitlines():
        stripped = _QUOTED.sub(" ", _PARENTHETICAL.sub(" ", line)).strip()
        if not stripped:
            continue
        # Bare interrogatives inside a sentence do not count; the line has to ask.
        if stripped.rstrip("*_` ").endswith("?"):
            asked += 1
    return asked


def _send(client: TestClient, project_id: str, message: str) -> dict:
    response = client.post(
        "/api/chat",
        json={"project_id": project_id, "messages": [{"role": "user", "content": message}]},
    )
    if response.status_code != 200:
        return {"error": f"HTTP {response.status_code}", "detail": response.text[:300]}
    return response.json()


def _answer_decision(client: TestClient, project_id: str, response: dict) -> dict | None:
    """Play the professor: take the recommended option so the cycle continues."""
    decision = response.get("decision")
    if not decision or not decision.get("options"):
        return None
    option = decision["options"][0]
    follow_up = client.post(
        "/api/chat",
        json={
            "project_id": project_id,
            "messages": [{"role": "user", "content": option["value"]}],
            "decision_trace": {
                "origin_message_id": (response.get("assistant_message") or {}).get("id"),
                "question": decision["question"][:500],
                "selected_label": option["label"][:200],
                "selected_value": option["value"][:2000],
            },
        },
    )
    return follow_up.json() if follow_up.status_code == 200 else None


def run_scenario(client: TestClient, scenario: Scenario, mode: str) -> dict:
    arc = scenario.codesign_arc if mode == "Co-design" else scenario.auto_arc
    created = client.post(
        "/api/projects", json={**scenario.project, "mode": mode}
    )
    created.raise_for_status()
    project_id = created.json()["project"]["id"]
    turns: list[dict] = []

    for phase, prompt in arc:
        response = _send(client, project_id, prompt)
        if "error" in response:
            turns.append({"phase": phase, "prompt": prompt, **response})
            continue
        record = {
            "phase": phase,
            "prompt": prompt,
            "reply": response.get("content") or "",
            "reply_chars": len(response.get("content") or ""),
            "asked_decision": bool(response.get("decision")),
            "questions_asked": count_questions(response.get("content") or ""),
            "artifact": (response.get("artifact") or {}).get("title"),
            "artifact_format": (response.get("artifact") or {}).get("file_format"),
            "tool_calls": [t["tool"] for t in (response.get("skill_tool_calls") or [])],
            "state_file": (response.get("state_file") or {}).get("file"),
        }
        # In Co-design a decision card is a stop; a real professor answers it.
        if mode == "Co-design":
            answered = _answer_decision(client, project_id, response)
            if answered:
                record["decision_answered"] = True
                record["reply"] += "\n" + (answered.get("content") or "")
        turns.append(record)

    workspace = client.get(f"/api/projects/{project_id}").json()
    state = client.get(f"/api/projects/{project_id}/state").json()
    validation = client.post(
        f"/api/projects/{project_id}/validate", params={"design_profile": "produce"}
    ).json()

    replies = [t.get("reply", "") for t in turns if t.get("reply")]
    visible = "\n".join(replies).lower()
    file_artifacts = [a for a in workspace.get("artifacts", []) if a.get("has_file")]
    formats = {a.get("file_format") for a in file_artifacts}

    checks = {
        "completed_whole_arc": len([t for t in turns if "error" not in t]) == len(arc),
        "produced_office_files": len(file_artifacts) >= scenario.min_office_files,
        "no_internal_terms": not any(term in visible for term in INTERNAL_TERMS),
        "no_raw_contract_fences": not re.search(
            r"```(?:json|artifact_spec|state_file)\b", visible
        ),
        "no_stub_only_replies": not any(is_stub(r) for r in replies),
    }
    for required in scenario.required_formats:
        checks[f"produced_{required}_file"] = required in formats
    offending_claims = [
        c for c in scenario.forbidden_claims if re.search(c, visible, re.I)
    ]
    if scenario.forbidden_claims:
        checks["made_no_forbidden_claim"] = not offending_claims

    if mode == "Auto":
        checks["auto_never_asked"] = not any(t.get("asked_decision") for t in turns)
        checks["auto_no_question_requests"] = not any(
            p in visible for p in PROFESSOR_QUESTION_PATTERNS
        )
        # Delegation means the brief is answered in full, not in fragments.
        checks["auto_front_loaded_substantive"] = (
            turns[0].get("reply_chars", 0) > 400 if turns else False
        )
    else:
        checks["codesign_engaged_early"] = any(
            t.get("questions_asked", 0) > 0 or t.get("asked_decision") for t in turns[:3]
        )
        checks["codesign_bounded_questions"] = all(
            t.get("questions_asked", 0) <= 3 for t in turns
        )

    observations = {
        "median_reply_chars": int(statistics.median([len(r) for r in replies] or [0])),
        "shortest_reply_chars": min([len(r) for r in replies] or [0]),
        "stub_replies": sum(1 for r in replies if is_stub(r)),
        "artifacts_created": len(file_artifacts),
        "artifact_formats": sorted(f for f in formats if f),
        "state_files": [e["file"] for e in state.get("state_files", [])],
        "validator_status": validation.get("status"),
        "tool_calls_total": sum(len(t.get("tool_calls", [])) for t in turns),
        "decision_cards_shown": sum(1 for t in turns if t.get("asked_decision")),
        "forbidden_claims_matched": offending_claims,
    }
    for probe in scenario.probes:
        window = "\n".join(t.get("reply", "") for t in turns[probe.from_turn :]).lower()
        hit = bool(re.search(probe.pattern, window, re.IGNORECASE))
        observations[probe.name] = (not hit) if probe.expect_absent else hit

    return {
        "scenario": scenario.name,
        "mode": mode,
        "project_id": project_id,
        "arc_length": len(arc),
        "turns": turns,
        "checks": checks,
        "observations": observations,
        "status": "pass" if all(checks.values()) else "fail",
    }


def summarize(scenario: Scenario, results: list[dict]) -> str:
    lines = [
        f"# {scenario.name}",
        "",
        f"Generated {datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC.",
        "",
        "Co-design gets a conversational arc where constraints emerge over time. Auto gets a "
        "delegation brief with the same constraints stated up front. Same deliverables "
        "demanded; different interaction shape.",
        "",
        "| Mode | Turns | Status | Artifacts | Decision cards | Stub replies | Tool calls |",
        "|---|---|---|---|---|---|---|",
    ]
    for r in results:
        o = r["observations"]
        lines.append(
            f"| {r['mode']} | {r['arc_length']} | **{r['status']}** | "
            f"{o['artifacts_created']} | {o['decision_cards_shown']} | "
            f"{o['stub_replies']} | {o['tool_calls_total']} |"
        )
    for r in results:
        lines += ["", f"## {r['mode']}", "", "### Checks", ""]
        for name, ok in r["checks"].items():
            lines.append(f"- {'PASS' if ok else 'FAIL'} — `{name}`")
        lines += ["", "### Observations (recorded, not asserted)", ""]
        for name, value in r["observations"].items():
            lines.append(f"- `{name}`: {value}")
        lines += ["", "### Arc", "", "| Turn | Reply chars | Artifact | Decision |", "|---|---|---|---|"]
        for t in r["turns"]:
            lines.append(
                f"| {t['phase']} | {t.get('reply_chars', 0)} | "
                f"{t.get('artifact') or '—'} | {'yes' if t.get('asked_decision') else '—'} |"
            )
    return "\n".join(lines) + "\n"


def main(scenario: Scenario, slug: str, argv: list[str] | None = None) -> int:
    import argparse
    import sys

    from server import app

    parser = argparse.ArgumentParser(description=scenario.name)
    parser.add_argument("--mode", choices=MODES, help="run only one mode")
    parser.add_argument("--keep", action="store_true", help="leave the projects in place")
    args = parser.parse_args(argv if argv is not None else sys.argv[1:])

    modes = [args.mode] if args.mode else list(MODES)
    client = TestClient(app)
    results = [run_scenario(client, scenario, mode) for mode in modes]

    if not args.keep:
        for r in results:
            client.delete(f"/api/projects/{r['project_id']}")

    RESULTS.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    (RESULTS / f"{slug}-{stamp}.json").write_text(
        json.dumps(results, indent=2), encoding="utf-8"
    )
    report = summarize(scenario, results)
    (RESULTS / f"{slug}-{stamp}.md").write_text(report, encoding="utf-8")
    print(report)
    return 0 if all(r["status"] == "pass" for r in results) else 1
