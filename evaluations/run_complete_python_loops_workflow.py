from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from fastapi.testclient import TestClient


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from server import app  # noqa: E402


RESULTS = ROOT / "evaluations" / "results" / "python-loops-complete-workflow-20260803"
INTERNAL_TERMS = (
    "skill.md",
    "state_file",
    "artifact_spec",
    "json schema",
    "validator command",
    "project-state filename",
)
PROFESSOR_QUESTION_PATTERNS = (
    "would you like",
    "which approach",
    "please provide",
    "what additional information",
    "select the information",
)


def create_project(client: TestClient, name: str, mode: str) -> str:
    response = client.post(
        "/api/projects",
        json={
            "name": name,
            "course_name": "Introduction to Python",
            "level": "Undergraduate",
            "class_time": "75 minutes",
            "outcome": "Students will trace and write loops and explain what remains true as a loop repeats.",
            "mode": mode,
            "notes": "Professor-centered end-to-end workflow test.",
        },
    )
    response.raise_for_status()
    return response.json()["project"]["id"]


def send(
    client: TestClient,
    project_id: str,
    prompt: str,
    response_to_decision: dict | None = None,
) -> dict:
    payload: dict = {
        "project_id": project_id,
        "skill_profile": "auto",
        "messages": [{"role": "user", "content": prompt}],
    }
    if response_to_decision and response_to_decision.get("decision"):
        decision = response_to_decision["decision"]
        selected = decision["options"][0]
        payload["display_content"] = f"Selected: {selected['label']}"
        payload["decision_trace"] = {
            "origin_message_id": response_to_decision["assistant_message"]["id"],
            "question": decision["question"],
            "selected_label": selected["label"],
            "selected_value": selected["value"],
        }
    response = client.post("/api/chat", json=payload)
    response.raise_for_status()
    return response.json()


def response_record(phase: str, prompt: str, response: dict) -> dict:
    artifact = response.get("artifact") or {}
    return {
        "phase": phase,
        "prompt": prompt,
        "response": response,
        "visible_question_count": len(re.findall(r"\?", response.get("content", "")))
        + (1 if response.get("decision") else 0),
        "artifact_file_created": bool(artifact.get("has_file")),
    }


def run_codesign(client: TestClient) -> dict:
    project_id = create_project(
        client, "Python loops — complete professor Co-design workflow", "Co-design"
    )
    turns: list[dict] = []

    intro = "I want to teach about loops."
    response = send(client, project_id, intro)
    turns.append(response_record("lesson_start", intro, response))

    context = (
        "This is an introductory Python class. Students know variables, conditionals, and lists, "
        "but they have not used loops. I have 75 minutes with 24 students in a computer lab. "
        "By the end, they should trace and write a loop and explain what remains true as it repeats."
    )
    response = send(client, project_id, context, response)
    turns.append(response_record("lesson_context", context, response))

    lesson_request = (
        "Please create the practical lesson plan now. Include a prediction before students run "
        "code, guided practice, and a quick way for me to see who is still confused."
    )
    response = send(client, project_id, lesson_request, response)
    turns.append(response_record("lesson_production", lesson_request, response))
    if response.get("decision"):
        lesson_choice = (
            "Use the recommended teaching approach and finish the 75-minute lesson plan with the "
            "context I already supplied."
        )
        response = send(client, project_id, lesson_choice, response)
        turns.append(response_record("lesson_decision_answer", lesson_choice, response))

    quiz_start = "I want to create a quiz about invariants inside a loop."
    quiz_response = send(client, project_id, quiz_start)
    turns.append(response_record("quiz_start", quiz_start, quiz_response))

    quiz_context = (
        "Make it a 10-minute practice quiz for feedback, not a grade. Students should identify an "
        "invariant and explain why it remains true. They may run the code, but may not use notes or AI."
    )
    quiz_response = send(client, project_id, quiz_context, quiz_response)
    turns.append(response_record("quiz_context", quiz_context, quiz_response))

    quiz_production = (
        "Create the ready-to-use Word quiz now. Include three questions, space for reasoning, and "
        "a concise instructor answer key. Use the information already provided."
    )
    quiz_response = send(client, project_id, quiz_production, quiz_response)
    turns.append(response_record("quiz_production", quiz_production, quiz_response))
    if not (quiz_response.get("artifact") or {}).get("has_file"):
        retry = (
            "Complete the downloadable Word quiz now using the agreed practice format. Make safe "
            "reversible assumptions for presentation details instead of asking another question."
        )
        quiz_response = send(client, project_id, retry, quiz_response)
        turns.append(response_record("quiz_production_retry", retry, quiz_response))

    revision = (
        "Revise the quiz: make Question 2 use a while loop with an index, and make its answer key "
        "explicitly explain initialization, preservation, and what the invariant means at the end."
    )
    revision_response = send(client, project_id, revision, quiz_response)
    turns.append(response_record("quiz_revision", revision, revision_response))
    if revision_response.get("decision"):
        revision_followup = "Use the recommended option and complete that revised Word quiz now."
        revision_response = send(
            client, project_id, revision_followup, revision_response
        )
        turns.append(
            response_record("quiz_revision_decision_answer", revision_followup, revision_response)
        )

    handoff = (
        "Briefly summarize what is ready for class and what I still need to review. Use a short list, "
        "not an implementation table."
    )
    handoff_response = send(client, project_id, handoff)
    turns.append(response_record("handoff", handoff, handoff_response))

    workspace = client.get(f"/api/projects/{project_id}").json()
    file_artifacts = [item for item in workspace["artifacts"] if item.get("has_file")]
    visible = "\n".join(turn["response"].get("content", "") for turn in turns).lower()
    start_questions = turns[0]["visible_question_count"]
    quiz_questions = next(
        turn["visible_question_count"] for turn in turns if turn["phase"] == "quiz_start"
    )
    checks = {
        "multi_turn": len(turns) >= 7,
        "lesson_start_asks_at_most_three": 1 <= start_questions <= 3,
        "quiz_start_asks_at_most_three": quiz_questions <= 3,
        "no_internal_terms": not any(term in visible for term in INTERNAL_TERMS),
        "office_quiz_created": bool(file_artifacts),
        "revision_completed": bool(
            (revision_response.get("artifact") or {}).get("has_file")
        ),
        "no_raw_contract_fences": not re.search(
            r"```(?:markdown|json|artifact_spec|state_file)\b", visible
        ),
    }
    return {
        "project_id": project_id,
        "mode": "Co-design",
        "turns": turns,
        "workspace": workspace,
        "checks": checks,
        "status": "pass" if all(checks.values()) else "fail",
    }


def run_auto(client: TestClient) -> dict:
    project_id = create_project(client, "Python loops — complete Auto workflow", "Auto")
    turns: list[dict] = []
    initial = (
        "Create a 75-minute introductory Python lesson on loops for 24 students in a computer lab. "
        "They know variables, conditionals, and lists. Include prediction, guided practice, and a "
        "10-minute formative loop-invariant quiz with three reasoning questions and an answer key. "
        "Create the quiz as a downloadable Word worksheet."
    )
    response = send(client, project_id, initial)
    turns.append(response_record("auto_production", initial, response))
    if not (response.get("artifact") or {}).get("has_file"):
        completion = (
            "Finish the downloadable Word quiz using your stated assumptions and the supplied class "
            "context. Do not ask me a question."
        )
        response = send(client, project_id, completion)
        turns.append(response_record("auto_production_completion", completion, response))

    revision = (
        "Revise Question 2 to use a while loop with an index and update the answer key to explain "
        "initialization, preservation, and the invariant at termination."
    )
    revision_response = send(client, project_id, revision)
    turns.append(response_record("auto_revision", revision, revision_response))
    workspace = client.get(f"/api/projects/{project_id}").json()
    visible = "\n".join(turn["response"].get("content", "") for turn in turns).lower()
    asks_professor = any(pattern in visible for pattern in PROFESSOR_QUESTION_PATTERNS)
    checks = {
        "multi_turn": len(turns) >= 2,
        "no_decision_cards": all(not turn["response"].get("decision") for turn in turns),
        "no_professor_question_request": not asks_professor,
        "office_artifact_created": any(
            item.get("has_file") for item in workspace.get("artifacts", [])
        ),
        "revision_completed": bool(
            (revision_response.get("artifact") or {}).get("has_file")
        ),
        "no_raw_contract_fences": not re.search(
            r"```(?:markdown|json|artifact_spec|state_file)\b", visible
        ),
    }
    return {
        "project_id": project_id,
        "mode": "Auto",
        "turns": turns,
        "workspace": workspace,
        "checks": checks,
        "status": "pass" if all(checks.values()) else "fail",
    }


def main() -> int:
    RESULTS.mkdir(parents=True, exist_ok=True)
    client = TestClient(app)
    codesign = run_codesign(client)
    auto = run_auto(client)
    result = {
        "test": "Complete professor-centered Introduction to Python loops workflow",
        "codesign": codesign,
        "auto": auto,
        "status": "pass" if codesign["status"] == auto["status"] == "pass" else "fail",
    }
    output = RESULTS / "workflow-record.json"
    output.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "output": str(output),
                "status": result["status"],
                "codesign": {
                    "project_id": codesign["project_id"],
                    "turns": len(codesign["turns"]),
                    "checks": codesign["checks"],
                },
                "auto": {
                    "project_id": auto["project_id"],
                    "turns": len(auto["turns"]),
                    "checks": auto["checks"],
                },
            },
            indent=2,
        )
    )
    return 0 if result["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
