from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from fastapi.testclient import TestClient


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from server import app  # noqa: E402


RESULTS = ROOT / "evaluations" / "results" / "realistic-professor-codesign-20260803"
FORBIDDEN_PROFESSOR_TERMS = (
    "markdown",
    "skill.md",
    "skill structure",
    "schema",
    "state_file",
    "artifact_spec",
    "json block",
    "validator",
    "project index",
)


def create_project(client: TestClient) -> str:
    response = client.post(
        "/api/projects",
        json={
            "name": "Realistic professor Co-design — Python loops",
            "course_name": "Introduction to Python",
            "level": "Undergraduate",
            "class_time": "75 minutes",
            "outcome": (
                "Students should be able to use a loop to process a list and explain "
                "what remains true each time the loop repeats."
            ),
            "mode": "Co-design",
            "notes": (
                "There are 24 students in a computer lab. They already know variables, "
                "lists, and if statements, but this is their first lesson on loops."
            ),
        },
    )
    response.raise_for_status()
    return response.json()["project"]["id"]


def chat(client: TestClient, project_id: str, message: str, **extra: object) -> dict:
    payload = {
        "project_id": project_id,
        "skill_profile": "auto",
        "messages": [{"role": "user", "content": message}],
        **extra,
    }
    response = client.post("/api/chat", json=payload)
    response.raise_for_status()
    return response.json()


def answer_first_decision(client: TestClient, project_id: str, response: dict) -> dict | None:
    decision = response.get("decision")
    if not decision:
        return None
    selected = decision["options"][0]
    plain_label = re.sub(r"\s*\(Recommended\)\s*", "", selected["label"]).strip().lower()
    display = f"Let’s use {plain_label}."
    return chat(
        client,
        project_id,
        (
            f"{display} Please turn that into a practical 75-minute class plan. "
            "Include something students predict before they run code, time to practice, "
            "and a quick way for me to see who is still confused."
        ),
        display_content=display,
        decision_trace={
            "origin_message_id": response["assistant_message"]["id"],
            "question": decision["question"],
            "selected_label": selected["label"],
            "selected_value": selected["value"],
        },
    )


def professor_language_audit(responses: list[dict]) -> dict:
    visible = "\n\n".join(response.get("content", "") for response in responses)
    for response in responses:
        decision = response.get("decision")
        if decision:
            visible += "\n" + decision["question"]
            visible += "\n" + "\n".join(option["label"] for option in decision["options"])
    lower = visible.lower()
    term_hits = [term for term in FORBIDDEN_PROFESSOR_TERMS if term in lower]
    questions = re.findall(r"[^\n?.!]{4,}\?", visible)
    return {
        "status": "pass" if not term_hits else "fail",
        "forbidden_term_hits": term_hits,
        "visible_question_count": len(questions),
        "questions": [question.strip() for question in questions],
        "criterion": (
            "Professor-facing language must stay about students, content, evidence, time, "
            "difficulty, feedback, assessment use, and classroom constraints—not implementation contracts."
        ),
    }


def main() -> int:
    RESULTS.mkdir(parents=True, exist_ok=True)
    client = TestClient(app)
    project_id = create_project(client)

    lesson_prompt = (
        "I want to teach loops in my introductory Python class next week. "
        "My students know variables, lists, and if statements, but they have not used loops before. "
        "I have 75 minutes with 24 students in a computer lab. Help me plan the lesson."
    )
    lesson_response = chat(client, project_id, lesson_prompt)
    lesson_followup = answer_first_decision(client, project_id, lesson_response)

    quiz_prompt = (
        "I also want a short quiz about loop invariants—the idea that something stays true "
        "each time the loop repeats. It should take about 10 minutes. I want students to "
        "explain their reasoning, not just choose an answer."
    )
    quiz_response = chat(client, project_id, quiz_prompt)

    responses = [lesson_response]
    if lesson_followup:
        responses.append(lesson_followup)
    responses.append(quiz_response)
    audit = professor_language_audit(responses)

    trace_response = client.get(f"/api/projects/{project_id}/trace")
    trace_response.raise_for_status()
    project_response = client.get(f"/api/projects/{project_id}")
    project_response.raise_for_status()

    record = {
        "project_id": project_id,
        "mode": "Co-design",
        "professor_prompts": {
            "lesson": lesson_prompt,
            "quiz": quiz_prompt,
        },
        "turns": {
            "lesson_response": lesson_response,
            "lesson_followup": lesson_followup,
            "quiz_response": quiz_response,
        },
        "professor_language_audit": audit,
        "trace": trace_response.json(),
        "project": project_response.json(),
    }
    (RESULTS / "conversation-record.json").write_text(
        json.dumps(record, indent=2), encoding="utf-8"
    )

    summary = {
        "project_id": project_id,
        "lesson_profile": lesson_response["skill_runtime"]["profile"],
        "lesson_decision": lesson_response.get("decision"),
        "lesson_followup_profile": (
            lesson_followup["skill_runtime"]["profile"] if lesson_followup else None
        ),
        "quiz_profile": quiz_response["skill_runtime"]["profile"],
        "quiz_decision": quiz_response.get("decision"),
        "professor_language_audit": audit,
        "trace_summary": trace_response.json()["summary"],
    }
    print(json.dumps(summary, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
