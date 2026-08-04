"""Prepare ECE 437 as a real, visible project in Co-design mode.

Unlike `run_vlsi_course_eval.py`, which drives a hidden `eval-` project in Rapid mode,
this creates an ordinary project that shows up in the app's project switcher and works
the way the skill's default mode actually behaves: short cycles that stop at a decision.

Co-design changes the mechanics. The app only extracts a `state_file` when the response
carries no decision (server.py), so each request may have to be answered before any file
is written. This driver plays the professor: when a decision card comes back it selects
the recommended option, replays it with a `decision_trace`, and continues the cycle.

Usage:
    python evaluations/run_vlsi_codesign_project.py
    python evaluations/run_vlsi_codesign_project.py --delete   # tear the project down after
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path

import httpx

APP_URL = "http://127.0.0.1:8001"
EVAL_DIR = Path(__file__).resolve().parent
FIXTURE = EVAL_DIR / "fixtures" / "vlsi-introduction.txt"
RESULT_DIR = EVAL_DIR / "results"
MAX_DECISION_ROUNDS = 3

PROJECT = {
    "name": "ECE 437 — Introduction to VLSI Design",
    "course_name": "ECE 437 Introduction to VLSI Design",
    "level": "Undergraduate",
    "class_time": "50 minutes",
    "outcome": (
        "Given a static CMOS gate specification and a load, students size the gate chain "
        "using logical effort and justify the sizing against delay and energy constraints."
    ),
    "mode": "Co-design",
    "notes": (
        "Legacy source is the prior-offering departmental record. Documented misconceptions: "
        "pull-up/pull-down duality, delay as a fixed per-gate constant, dynamic energy versus "
        "static leakage, setup versus hold. Prior exams over-sampled recall and never required "
        "a justified sizing or power decision. No identifiable student information."
    ),
}

CYCLES = (
    (
        "establish",
        "Let's begin. I am preparing ECE 437 Introduction to VLSI Design and the attached "
        "prior-offering record is my legacy source. Create the complete course-design-brief.md "
        "portable state file from the supplied template. Answer every field: use the legacy "
        "record where it speaks, and write what is still unsupplied rather than leaving a field "
        "blank or inventing authority. Emit it as a state_file block.",
        "course-design-brief.md",
    ),
    (
        "design",
        "Now create the complete alignment-map.md for the first three weeks: CMOS gate "
        "construction and duality, delay and logical effort, and dynamic versus static power. "
        "Give every outcome an ID of the form LO-1, a cognitive-demand token from the controlled "
        "list, evidence of learning, a learning mechanism, an activity, and a feedback or "
        "assessment column. Use ASCII hyphens only. Emit it as a state_file block.",
        "alignment-map.md",
    ),
    (
        "assessment",
        "Now create the complete assessment-blueprint.md for the midterm over those three weeks. "
        "Item IDs must look like M-1 (letter, hyphen, number). Separate multiple outcome IDs with "
        "semicolons, never commas. Use plain ASCII hyphens in every column heading. The legacy "
        "record says prior exams over-sampled recall and never required a justified decision, so "
        "make the sampled demand match the outcomes. Emit it as a state_file block.",
        "assessment-blueprint.md",
    ),
    (
        "validation",
        "Now create the complete project-index.md listing every state file this project holds, "
        "each with schema version, authority, status, and last-updated date. Mark anything still "
        "missing as a gap rather than omitting it. Emit it as a state_file block.",
        "project-index.md",
    ),
)


def _now() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def run(delete: bool) -> dict:
    started = _now()
    cycles = []

    with httpx.Client(base_url=APP_URL, timeout=300.0) as client:
        model_id = client.get("/api/config").json().get("model_id")
        created = client.post("/api/projects", json=PROJECT)
        created.raise_for_status()
        project = created.json()["project"]
        project_id = project["id"]

        upload = client.post(
            f"/api/projects/{project_id}/sources",
            files={"files": (FIXTURE.name, FIXTURE.read_bytes(), "text/plain")},
            data={"data_classification_ack": "true"},
        )
        upload.raise_for_status()
        source_id = upload.json()["sources"][0]["source_id"]

        for profile, prompt, expected in CYCLES:
            record = {
                "profile": profile,
                "expected_state_file": expected,
                "decisions_answered": [],
                "requests": 0,
            }
            message = prompt
            decision_trace = None

            for _ in range(MAX_DECISION_ROUNDS + 1):
                body = {
                    "project_id": project_id,
                    "skill_profile": profile,
                    "messages": [{"role": "user", "content": message}],
                }
                if decision_trace:
                    body["decision_trace"] = decision_trace
                response = client.post("/api/chat", json=body)
                record["requests"] += 1
                if response.status_code != 200:
                    record["error"] = f"HTTP {response.status_code}: {response.text[:200]}"
                    break
                payload = response.json()
                decision = payload.get("decision")
                if decision:
                    # Play the professor: take the recommended option, which the contract
                    # puts first, and continue the cycle rather than stopping.
                    option = decision["options"][0]
                    record["decisions_answered"].append(
                        {"question": decision["question"], "selected": option["label"]}
                    )
                    decision_trace = {
                        "origin_message_id": (payload.get("assistant_message") or {}).get("id"),
                        "question": decision["question"][:500],
                        "selected_label": option["label"][:200],
                        "selected_value": option["value"][:2000],
                    }
                    message = option["value"]
                    continue
                state_file = payload.get("state_file") or {}
                record["state_file_written"] = state_file.get("file")
                record["validator_status"] = (state_file.get("validation") or {}).get("status")
                record["validator_findings"] = (state_file.get("validation") or {}).get(
                    "findings", []
                )
                record["response_chars"] = len(payload.get("content") or "")
                record["sources_used"] = payload.get("sources_used", [])
                break
            cycles.append(record)

        state = client.get(f"/api/projects/{project_id}/state").json()
        validation = client.post(
            f"/api/projects/{project_id}/validate", params={"design_profile": "establish"}
        ).json()
        workspace = client.get(f"/api/projects/{project_id}").json()

        if delete:
            client.delete(f"/api/projects/{project_id}")

    return {
        "run": "vlsi_codesign_visible_project",
        "started": started,
        "finished": _now(),
        "model_id": model_id,
        "project_id": project_id,
        "project_name": PROJECT["name"],
        "mode": PROJECT["mode"],
        "source_id": source_id,
        "cycles": cycles,
        "state_files": state.get("state_files", []),
        "validation": validation,
        "messages": len(workspace.get("messages", [])),
        "deleted": delete,
    }


def summarize(result: dict) -> str:
    lines = [
        "# ECE 437 Introduction to VLSI — Co-design project run",
        "",
        f"- Project: **{result['project_name']}** (`{result['project_id']}`)",
        f"- Mode: {result['mode']} · Model: {result['model_id']}",
        f"- Source: {result['source_id']} · Saved messages: {result['messages']}",
        "",
        "| Cycle | Requests | Decisions answered | State file | Validator |",
        "|---|---|---|---|---|",
    ]
    for cycle in result["cycles"]:
        lines.append(
            f"| {cycle['profile']} | {cycle['requests']} | "
            f"{len(cycle['decisions_answered'])} | "
            f"{cycle.get('state_file_written') or cycle.get('error', '—')} | "
            f"{cycle.get('validator_status') or '—'} |"
        )
    validation = result["validation"]
    state_lines = [
        f"- `{e['file']}` ({e['bytes']} bytes)" for e in result["state_files"]
    ] or ["- none"]
    lines += [
        "",
        "## Portable state",
        "",
        *state_lines,
        "",
        f"## Validators — overall **{validation.get('status')}**",
        "",
    ]
    for check in validation.get("checks", []):
        lines.append(f"- `{check['script']}` → {check['status']}")
        for finding in check["findings"][:6]:
            lines.append(f"  - {finding['level']}: {finding['message']}")
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--delete", action="store_true", help="remove the project afterwards")
    args = parser.parse_args()

    result = run(args.delete)
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    (RESULT_DIR / f"vlsi-codesign-{result['started']}.json").write_text(
        json.dumps(result, indent=2), encoding="utf-8"
    )
    report = summarize(result)
    (RESULT_DIR / f"vlsi-codesign-{result['started']}.md").write_text(report, encoding="utf-8")
    print(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
