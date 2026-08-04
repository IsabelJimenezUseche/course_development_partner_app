"""Run a complete professor-centered VLSI Co-design workflow against the live app.

Like the Python workflow, this test uses only information supplied by the simulated
professor in the project and conversation. It does not upload an internal course record.
The workflow produces a lesson and Word quiz, revises that quiz, and closes with a handoff.
A single successful response is never scored as a passing run.
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import httpx

sys.path.insert(0, str(Path(__file__).resolve().parent))
from realistic_arc import count_questions  # noqa: E402


APP_URL = "http://127.0.0.1:8001"
ROOT = Path(__file__).resolve().parents[1]
RESULTS = ROOT / "evaluations" / "results" / "vlsi-complete-codesign-20260803"
INTERNAL_TERMS = (
    "skill.md",
    "state_file",
    "artifact_spec",
    "json schema",
    "validator command",
    "project-state filename",
)


def compact(content: str) -> str:
    return re.sub(r"[\s\u00a0\u202f]+", " ", content).lower()


def has_process_specific_number(content: str) -> bool:
    return bool(
        re.search(r"\b\d+(?:\.\d+)?\s*(?:nm|ff|ps)\b", content, re.IGNORECASE)
    )


def lesson_artifact_is_technically_ready(response: dict) -> bool:
    artifact = response.get("artifact") or {}
    content = compact(artifact.get("content", ""))
    invalid_claim = re.search(
        r"(?:increasing (?:the )?(?:gate )?size reduces g|"
        r"g (?:decreases|reduces|scales) (?:with|as)[^.;]{0,30}width|"
        r"p (?:scales|increases) (?:linearly )?(?:with|as)[^.;]{0,30}width|"
        r"p\s*(?:∝|=)\s*w)",
        content,
    )
    return bool(
        artifact.get("has_file")
        and content
        and "g" in content
        and "p" in content
        and "h" in content
        and invalid_claim is None
        and not has_process_specific_number(content)
    )


def quiz_artifact_has_correct_key(response: dict) -> bool:
    artifact = response.get("artifact") or {}
    content = compact(artifact.get("content", ""))
    return bool(
        artifact.get("has_file")
        and "design a" in content
        and "design b" in content
        and "answer key" in content
        and re.search(r"(?:total[^\n]{0,40})?6(?:\.0+)?\b", content)
        and re.search(r"(?:total[^\n]{0,40})?5(?:\.0+)?\b", content)
        and re.search(
            r"(?:design|option|preferred|choose|select)[^\n]{0,30}\bb\b", content
        )
    )


def create_project(client: httpx.Client) -> str:
    response = client.post(
        "/api/projects",
        json={
            "name": "ECE 437 VLSI — complete professor Co-design workflow",
            "course_name": "ECE 437 Introduction to VLSI Design",
            "level": "Undergraduate, third or fourth year",
            "class_time": "75 minutes",
            "outcome": (
                "Students compare transistor-sizing choices and justify one using load, "
                "delay, and logical-effort reasoning."
            ),
            "mode": "Co-design",
            "notes": (
                "Professor-supplied test context only. Do not invent process-design-kit "
                "values, design rules, or current technology parameters."
            ),
        },
    )
    response.raise_for_status()
    return response.json()["project"]["id"]


def send(
    client: httpx.Client,
    project_id: str,
    prompt: str,
    response_to_decision: dict | None = None,
) -> tuple[dict, list[int]]:
    body: dict = {
        "project_id": project_id,
        "skill_profile": "auto",
        "messages": [{"role": "user", "content": prompt}],
    }
    if response_to_decision and response_to_decision.get("decision"):
        decision = response_to_decision["decision"]
        selected = decision["options"][0]
        body["display_content"] = f"Selected: {selected['label']}"
        body["decision_trace"] = {
            "origin_message_id": response_to_decision["assistant_message"]["id"],
            "question": decision["question"],
            "selected_label": selected["label"],
            "selected_value": selected["value"],
        }

    attempts: list[int] = []
    for _ in range(2):
        response = client.post("/api/chat", json=body)
        attempts.append(response.status_code)
        if response.status_code == 200:
            return response.json(), attempts
    raise RuntimeError(f"Chat failed after {attempts}: {response.text[:500]}")


def response_record(phase: str, prompt: str, response: dict, attempts: list[int]) -> dict:
    artifact = response.get("artifact") or {}
    return {
        "phase": phase,
        "prompt": prompt,
        "response": response,
        "attempts": attempts,
        # A "?" glyph is not a question: compound asks and parenthetical
        # examples inflated this and failed replies that obeyed the limit.
        "visible_question_count": count_questions(response.get("content", ""))
        + (1 if response.get("decision") else 0),
        "artifact_file_created": bool(artifact.get("has_file")),
    }


def main() -> int:
    RESULTS.mkdir(parents=True, exist_ok=True)
    turns: list[dict] = []
    with httpx.Client(base_url=APP_URL, timeout=300.0) as client:
        project_id = create_project(client)

        lesson_start = "I want to teach about transistor sizing and delay in VLSI."
        response, attempts = send(client, project_id, lesson_start)
        turns.append(response_record("lesson_start", lesson_start, response, attempts))

        lesson_context = (
            "These are third- and fourth-year ECE students who know digital logic and basic "
            "electronic circuits and can read transistor schematics. They tend to treat delay "
            "as a fixed value for each gate. I have 75 minutes with 28 students working in pairs. "
            "By the end, they should compare two sizing choices and justify one using load, delay, "
            "and logical-effort reasoning. Use normalized examples only—no process-specific values."
        )
        response, attempts = send(client, project_id, lesson_context, response)
        turns.append(response_record("lesson_context", lesson_context, response, attempts))

        lesson_production = (
            "Create the practical 75-minute lesson now. Include a prediction before calculation, "
            "a guided sizing comparison, pair work, and a quick way for me to see whether students "
            "still think delay is fixed per gate. Use d = gh + p and explain that widening a gate "
            "reduces its resistance but increases the input capacitance seen by the previous stage. "
            "Do not simplify an isolated gate's RC product to 1/W or 1/W squared."
        )
        response, attempts = send(client, project_id, lesson_production, response)
        turns.append(
            response_record("lesson_production", lesson_production, response, attempts)
        )
        if response.get("decision"):
            lesson_choice = (
                "Use the recommended teaching approach and finish the lesson with the context "
                "and source evidence already supplied."
            )
            response, attempts = send(client, project_id, lesson_choice, response)
            turns.append(
                response_record(
                    "lesson_decision_answer", lesson_choice, response, attempts
                )
            )

        lesson_revision = (
            "Before I use the lesson, correct its normalized logical-effort model. For scaled "
            "copies of the same gate topology, keep g and p constant in this classroom model. "
            "Sizing changes input capacitance and therefore h = Cload/Cin, while also changing "
            "the load seen by the previous stage. Do not say that increasing width reduces g or "
            "that p scales linearly with width. Use only normalized W0, C0, and delay quantities—"
            "do not introduce fF, ps, nm, or any process value. Recheck the example arithmetic "
            "and create the corrected lesson as a downloadable Word document."
        )
        lesson_response, attempts = send(
            client, project_id, lesson_revision, response
        )
        turns.append(
            response_record(
                "lesson_technical_revision", lesson_revision, lesson_response, attempts
            )
        )
        if not lesson_artifact_is_technically_ready(lesson_response):
            lesson_retry = (
                "The lesson file is missing or still violates the technical assumptions. Recreate "
                "the complete downloadable Word lesson now. For a scaled copy of one topology, "
                "keep g and p constant; put sizing effects in Cin, h = Cload/Cin, and the load on "
                "the previous stage. Use only W0, C0, and normalized delay—no fF, ps, or nm."
            )
            lesson_response, attempts = send(
                client, project_id, lesson_retry, lesson_response
            )
            turns.append(
                response_record(
                    "lesson_technical_revision_file",
                    lesson_retry,
                    lesson_response,
                    attempts,
                )
            )

        quiz_start = "I also want a short quiz about logical effort and transistor sizing."
        quiz_response, attempts = send(client, project_id, quiz_start)
        turns.append(response_record("quiz_start", quiz_start, quiz_response, attempts))

        quiz_context = (
            "Make it a 15-minute practice quiz for feedback, not a grade. Students should compare "
            "two gate-chain sizing choices and justify which better addresses a stated load and "
            "delay goal. They may use a calculator and the course formula sheet. Keep all quantities "
            "normalized and do not invent current process or design-rule values."
        )
        quiz_response, attempts = send(
            client, project_id, quiz_context, quiz_response
        )
        turns.append(response_record("quiz_context", quiz_context, quiz_response, attempts))

        quiz_production = (
            "Create the ready-to-use Word quiz now with three questions, reasoning space, and a "
            "concise instructor answer key. Use only the course and student information I have "
            "provided in this conversation."
        )
        quiz_response, attempts = send(
            client, project_id, quiz_production, quiz_response
        )
        turns.append(
            response_record("quiz_production", quiz_production, quiz_response, attempts)
        )
        if not (quiz_response.get("artifact") or {}).get("has_file"):
            retry = (
                "Complete the downloadable Word quiz now using the agreed formative format and "
                "normalized quantities. Make safe presentation assumptions instead of asking "
                "another question."
            )
            quiz_response, attempts = send(client, project_id, retry, quiz_response)
            turns.append(
                response_record("quiz_production_retry", retry, quiz_response, attempts)
            )

        revision = (
            "Revise the downloadable Word VLSI quiz: replace Question 2 with this normalized "
            "two-stage comparison. Design A has W1 = W0 and W2 = 2W0; Design B has W1 = 2W0 "
            "and W2 = 4W0; both drive an external load of 4C0. Use g = 1 and p = 1 per stage, "
            "with h1 = W2/W1 and h2 = 4W0/W2. Ask students to calculate total delay, compare "
            "internal capacitance as a simple dynamic-energy proxy, and select and justify one. "
            "Independently verify every answer-key calculation and do not claim process-specific data."
        )
        revision_response, attempts = send(client, project_id, revision, quiz_response)
        turns.append(response_record("quiz_revision", revision, revision_response, attempts))
        if revision_response.get("decision"):
            revision_choice = "Use the recommended option and complete the revised Word quiz now."
            revision_response, attempts = send(
                client, project_id, revision_choice, revision_response
            )
            turns.append(
                response_record(
                    "quiz_revision_decision_answer",
                    revision_choice,
                    revision_response,
                    attempts,
                )
            )
        if not quiz_artifact_has_correct_key(revision_response):
            revision_retry = (
                "The revised Word file is missing content or its key cannot be verified. Recreate "
                "the complete three-question downloadable Word quiz now, including the full question "
                "text and answer key. In Question 2, Design A uses W1=W0 and W2=2W0, giving h1=2, "
                "h2=2, and total delay 6. Design B uses W1=2W0 and W2=4W0, giving h1=2, h2=1, "
                "and total delay 5. State that B is faster but has the higher capacitance/energy "
                "proxy. Use no process-specific values and do not ask another question."
            )
            revision_response, attempts = send(
                client, project_id, revision_retry, revision_response
            )
            turns.append(
                response_record(
                    "quiz_revision_file", revision_retry, revision_response, attempts
                )
            )

        handoff = (
            "Briefly summarize what is ready for class and what still requires my technical "
            "verification. Use a short list, not an implementation table."
        )
        handoff_response, attempts = send(client, project_id, handoff)
        turns.append(response_record("handoff", handoff, handoff_response, attempts))

        workspace_response = client.get(f"/api/projects/{project_id}")
        workspace_response.raise_for_status()
        workspace = workspace_response.json()

    assistant_visible = "\n".join(
        turn["response"].get("content", "") for turn in turns
    )
    lower = assistant_visible.lower()
    final_deliverable_content = "\n".join(
        [
            (lesson_response.get("artifact") or {}).get("content", ""),
            (revision_response.get("artifact") or {}).get("content", ""),
        ]
    )
    # The professor-facing chat is deliberately short; the teaching lives in the
    # deliverables. Checks about what the lesson *teaches* have to read both, or they
    # test the summary rather than the material.
    taught_lower = (assistant_visible + "\n" + final_deliverable_content).lower()
    file_artifacts = [
        artifact for artifact in workspace.get("artifacts", []) if artifact.get("has_file")
    ]
    word_artifacts = [
        artifact
        for artifact in file_artifacts
        if artifact.get("file_format") == "docx"
    ]
    correct_delay_key = quiz_artifact_has_correct_key(revision_response)
    lesson_questions = turns[0]["visible_question_count"]
    quiz_questions = next(
        turn["visible_question_count"] for turn in turns if turn["phase"] == "quiz_start"
    )
    checks = {
        "multi_turn": len(turns) >= 7,
        "lesson_start_asks_at_most_three": 1 <= lesson_questions <= 3,
        "quiz_start_asks_at_most_three": quiz_questions <= 3,
        "no_internal_terms": not any(term in lower for term in INTERNAL_TERMS),
        "no_external_source_used": all(
            not (turn["response"].get("sources_used") or []) for turn in turns
        ),
        "professor_misconception_addressed": "fixed" in taught_lower
        and "delay" in taught_lower,
        "office_quiz_created": bool(word_artifacts),
        "revision_completed": bool(
            (revision_response.get("artifact") or {}).get("has_file")
        ),
        "revised_answer_key_calculations_correct": correct_delay_key,
        "lesson_technical_revision_completed": bool(
            (lesson_response.get("artifact") or {}).get("has_file")
        ),
        "lesson_technical_model_correct": lesson_artifact_is_technically_ready(
            lesson_response
        ),
        "no_incomplete_payload_notice": "withheld an incomplete" not in lower,
        "no_raw_contract_fences": not re.search(
            r"```(?:markdown|json|artifact_spec|state_file)\b", lower
        ),
        "no_unverified_process_numbers": not has_process_specific_number(
            final_deliverable_content
        ),
    }
    result = {
        "test": "Complete professor-centered ECE 437 VLSI Co-design workflow",
        "project_id": project_id,
        "source_id": None,
        "mode": "Co-design",
        "turns": turns,
        "workspace": workspace,
        "checks": checks,
        "status": "pass" if all(checks.values()) else "fail",
    }
    output = RESULTS / "workflow-record.json"
    output.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "output": str(output),
                "status": result["status"],
                "project_id": project_id,
                "source_id": None,
                "turns": len(turns),
                "file_artifacts": len(file_artifacts),
                "checks": checks,
            },
            indent=2,
        )
    )
    return 0 if result["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
