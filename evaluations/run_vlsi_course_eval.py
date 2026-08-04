"""Drive the app through a full course-preparation workflow and record what it produced.

Scenario: a professor preparing an Introduction to VLSI Design course works through the
app the way the UI would drive it — create the project, upload the legacy course record,
run the design turns, then run the skill's own validators over whatever portable state
came out the other side.

Companion to `compare_vlsi_paths.py`, which contrasts this run with the same workflow
carried out by applying the skill directly.

Usage:
    python evaluations/run_vlsi_course_eval.py            # against a running app
    python evaluations/run_vlsi_course_eval.py --keep     # leave the project in place
"""

from __future__ import annotations

import argparse
import json
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

import httpx

APP_URL = "http://127.0.0.1:8001"
EVAL_DIR = Path(__file__).resolve().parent
FIXTURE_DIR = EVAL_DIR / "fixtures"
RESULT_DIR = EVAL_DIR / "results"
FIXTURE = "vlsi-introduction.txt"

COURSE = {
    "course_name": "ECE 437 Introduction to VLSI Design",
    "level": "Undergraduate",
    "class_time": "50 minutes",
    "outcome": (
        "Given a static CMOS gate specification and a load, students size the gate chain "
        "using logical effort and justify the sizing against delay and energy constraints."
    ),
    # Rapid keeps the run non-interactive: one drafting pass per turn, assumptions visible,
    # consolidated faculty review at the end rather than a decision block mid-run.
    "mode": "Rapid",
}


@dataclass(frozen=True)
class Turn:
    """One thing the professor asks for, and what the response has to contain."""

    name: str
    skill_profile: str
    prompt: str
    expect_state_file: str | None = None
    required_terms: tuple[str, ...] = field(default_factory=tuple)


TURNS = (
    Turn(
        name="establish_design_brief",
        skill_profile="establish",
        prompt=(
            "Rapid mode. I am preparing ECE 437 Introduction to VLSI Design for next term. "
            "Use the attached prior-offering record as the legacy source. Create the complete "
            "course-design-brief.md portable state file from the supplied template. Record the "
            "course context, the prior-knowledge evidence actually available, the documented "
            "student misconceptions, the constraints, and the open decisions the course owner "
            "must resolve. Emit it as a state_file block."
        ),
        expect_state_file="course-design-brief.md",
        required_terms=("outcome", "misconception", "constraint"),
    ),
    Turn(
        name="alignment_map",
        skill_profile="design",
        prompt=(
            "Rapid mode. Create the complete alignment-map.md portable state file for the first "
            "three weeks of ECE 437: CMOS gate construction and network duality, delay and "
            "logical effort, and dynamic versus static power. For every outcome record a "
            "cognitive-demand token from the controlled list, the evidence of learning, the "
            "learning mechanism, the activity, and the feedback or assessment. Emit it as a "
            "state_file block."
        ),
        expect_state_file="alignment-map.md",
        required_terms=("LO-", "analyze", "evidence"),
    ),
    Turn(
        name="assessment_blueprint",
        skill_profile="assessment",
        prompt=(
            "Rapid mode. Create the complete assessment-blueprint.md portable state file for the "
            "ECE 437 midterm covering those three weeks. The legacy record says the prior exams "
            "over-sampled recall of gate structures and never required a justified sizing or "
            "power decision, so make the sampled cognitive demand match the outcomes. Declare "
            "the assessed-outcome scope explicitly. Emit it as a state_file block."
        ),
        expect_state_file="assessment-blueprint.md",
        required_terms=("outcome", "demand", "item"),
    ),
    Turn(
        name="lesson_artifact",
        skill_profile="artifact",
        prompt=(
            "Rapid mode. Produce a complete 50-minute active-learning worksheet for the logical-"
            "effort session. Students must size a three-stage gate chain driving a stated load "
            "and justify the sizing against a delay budget. Target the documented misconception "
            "that delay is a fixed per-gate constant. Include timing, ordered student tasks, "
            "instructor checkpoints, a debrief, and an instructor-review list. Cite the project "
            "source ID for any claim taken from the legacy record."
        ),
        required_terms=("minute", "student", "instructor", "debrief"),
    ),
    Turn(
        name="project_index",
        skill_profile="validation",
        prompt=(
            "Rapid mode. Create the complete project-index.md portable state file listing every "
            "state file this project now holds, each with its schema version, authority, status, "
            "and last-updated date. Mark anything still missing as a gap rather than omitting it. "
            "Emit it as a state_file block."
        ),
        expect_state_file="project-index.md",
        required_terms=("state file", "status"),
    ),
)


def _now() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")


def _grade_turn(turn: Turn, payload: dict, state_body: str = "") -> dict:
    """Record what actually came back, without asserting the content is pedagogically good."""
    state_file = payload.get("state_file") or {}
    validation = state_file.get("validation") or {}
    # When the whole reply was the state-file payload the app substitutes a short
    # confirmation line, so the terms have to be sought in the written file instead.
    content = ((payload.get("content") or "") + "\n" + state_body).lower()
    return {
        "turn": turn.name,
        "skill_profile_requested": turn.skill_profile,
        "skill_profile_used": (payload.get("skill_runtime") or {}).get("profile"),
        "loaded_files": (payload.get("skill_runtime") or {}).get("loaded_files", []),
        "loaded_assets": (payload.get("skill_runtime") or {}).get("loaded_assets", []),
        "response_chars": len(payload.get("content") or ""),
        "sources_used": payload.get("sources_used", []),
        "missing_terms": [t for t in turn.required_terms if t.lower() not in content],
        "expected_state_file": turn.expect_state_file,
        "state_file_written": state_file.get("file"),
        "state_file_matched": bool(
            turn.expect_state_file and state_file.get("file") == turn.expect_state_file
        ),
        "validator": validation.get("script"),
        "validator_status": validation.get("status"),
        "validator_findings": validation.get("findings", []),
        "artifact": (payload.get("artifact") or {}).get("title"),
        "artifact_error": payload.get("artifact_tool_error"),
    }


def run(keep: bool) -> dict:
    project_id = f"eval-vlsi-{uuid4().hex[:8]}"
    started = _now()
    transcript = []

    with httpx.Client(base_url=APP_URL, timeout=300.0) as client:
        health = client.get("/api/config")
        health.raise_for_status()
        model_id = health.json().get("model_id")

        # Uploading creates the project under our own `eval-` id (and marks it hidden),
        # matching how run_artifact_evals.py drives the app.
        upload = client.post(
            f"/api/projects/{project_id}/sources",
            files={"files": (FIXTURE, (FIXTURE_DIR / FIXTURE).read_bytes(), "text/plain")},
            data={"data_classification_ack": "true"},
        )
        upload.raise_for_status()
        upload_data = upload.json()
        if upload_data["errors"] or not upload_data["sources"]:
            raise RuntimeError(f"source upload failed: {upload_data['errors']}")
        source_id = upload_data["sources"][0]["source_id"]
        client.patch(f"/api/projects/{project_id}", json=COURSE).raise_for_status()

        for turn in TURNS:
            # gpt-oss intermittently returns no text content at all (HTTP 502 from the
            # app). Retry once so a transient upstream blip is not scored as a failure,
            # and record the retry so the flake rate stays visible.
            attempts = []
            payload = None
            for attempt in range(2):
                response = client.post(
                    "/api/chat",
                    json={
                        "project_id": project_id,
                        "skill_profile": turn.skill_profile,
                        "messages": [{"role": "user", "content": turn.prompt}],
                    },
                )
                attempts.append(response.status_code)
                if response.status_code == 200:
                    payload = response.json()
                    break
            if payload is None:
                transcript.append(
                    {
                        "turn": turn.name,
                        "error": response.text[:500],
                        "status": response.status_code,
                        "attempts": attempts,
                    }
                )
                continue
            written = (payload.get("state_file") or {}).get("file")
            state_body = ""
            if written:
                body = client.get(f"/api/projects/{project_id}/state/{written}")
                if body.status_code == 200:
                    state_body = body.json()["content"]
            record = _grade_turn(turn, payload, state_body)
            record["attempts"] = attempts
            record["content"] = payload.get("content")
            transcript.append(record)

        state = client.get(f"/api/projects/{project_id}/state").json()
        validation = client.post(
            f"/api/projects/{project_id}/validate", params={"design_profile": "produce"}
        ).json()
        workspace = client.get(f"/api/projects/{project_id}").json()

        state_contents = {}
        for entry in state.get("state_files", []):
            body = client.get(f"/api/projects/{project_id}/state/{entry['file']}")
            if body.status_code == 200:
                state_contents[entry["file"]] = body.json()["content"]

        if not keep:
            client.delete(f"/api/projects/{project_id}")

    return {
        "run": "vlsi_course_preparation_via_app",
        "started": started,
        "finished": _now(),
        "model_id": model_id,
        "project_id": project_id,
        "source_id": source_id,
        "course": COURSE,
        "turns": transcript,
        "state_files": state.get("state_files", []),
        "state_contents": state_contents,
        "validation": validation,
        "artifacts": [a.get("title") for a in workspace.get("artifacts", [])],
    }


def summarize(result: dict) -> str:
    lines = [
        "# Introduction to VLSI — course preparation through the app",
        "",
        f"- Run: {result['started']} → {result['finished']}",
        f"- Model: {result['model_id']}",
        f"- Course: {result['course']['course_name']} ({result['course']['mode']} mode)",
        f"- Source: {result['source_id']}",
        "",
        "## Turns",
        "",
        "| Turn | Profile | Assets | State file | Validator | Missing terms | Attempts |",
        "|---|---|---|---|---|---|---|",
    ]
    for turn in result["turns"]:
        if "error" in turn:
            lines.append(
                f"| {turn['turn']} | — | — | — | HTTP {turn['status']} | — | "
                f"{turn.get('attempts', [])} |"
            )
            continue
        lines.append(
            f"| {turn['turn']} | {turn['skill_profile_used']} | "
            f"{len(turn['loaded_assets'])} | {turn['state_file_written'] or '—'} | "
            f"{turn['validator_status'] or '—'} | "
            f"{', '.join(turn['missing_terms']) or 'none'} | "
            f"{len(turn.get('attempts', []))} |"
        )
    validation = result["validation"]
    state_lines = [
        f"- `{entry['file']}` ({entry['bytes']} bytes)" for entry in result["state_files"]
    ] or ["- none"]
    lines += [
        "",
        "## Portable state produced",
        "",
        *state_lines,
        "",
        "## Validator report",
        "",
        f"- Overall: **{validation.get('status')}** across {len(validation.get('checks', []))} checks",
    ]
    for check in validation.get("checks", []):
        lines.append(f"- `{check['script']}` → {check['status']} (exit {check['exit_code']})")
        for finding in check["findings"][:6]:
            lines.append(f"  - {finding['level']}: {finding['message']}")
    lines += ["", f"- Scope: {validation.get('scope_note', '')}", ""]
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--keep", action="store_true", help="leave the eval project in place")
    args = parser.parse_args()

    result = run(args.keep)
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    stamp = result["started"]
    (RESULT_DIR / f"vlsi-course-app-{stamp}.json").write_text(
        json.dumps(result, indent=2), encoding="utf-8"
    )
    report = summarize(result)
    (RESULT_DIR / f"vlsi-course-app-{stamp}.md").write_text(report, encoding="utf-8")
    print(report)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
