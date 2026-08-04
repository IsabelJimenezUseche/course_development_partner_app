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
INTERNAL_TERMS = (
    "markdown",
    "skill.md",
    "schema",
    "state_file",
    "artifact_spec",
    "json",
    "validator",
    "project-index",
)


def make_project(client: TestClient, name: str, outcome: str) -> str:
    response = client.post(
        "/api/projects",
        json={
            "name": name,
            "course_name": "Introduction to Python",
            "level": "Undergraduate",
            "class_time": "50 minutes",
            "outcome": outcome,
            "mode": "Co-design",
            "notes": "No other teaching context has been supplied yet.",
        },
    )
    response.raise_for_status()
    return response.json()["project"]["id"]


def ask(client: TestClient, project_id: str, prompt: str) -> dict:
    response = client.post(
        "/api/chat",
        json={
            "project_id": project_id,
            "skill_profile": "auto",
            "messages": [{"role": "user", "content": prompt}],
        },
    )
    response.raise_for_status()
    return response.json()


def visible_questions(response: dict) -> list[str]:
    questions: list[str] = []
    decision = response.get("decision")
    if decision:
        questions.append(decision["question"].strip())
    content = response.get("content", "")
    for line in content.splitlines():
        if "?" not in line:
            continue
        candidate = line[: line.index("?") + 1]
        cleaned = re.sub(r"<[^>]+>|[*_`]", "", candidate)
        cleaned = re.sub(r"^\s*(?:\||\d+[.)-]?)\s*", "", cleaned).strip(" |")
        if len(cleaned) >= 8 and cleaned not in questions:
            questions.append(cleaned)
    return questions


def audit_probe(kind: str, response: dict) -> dict:
    content = response.get("content", "")
    decision = response.get("decision")
    visible = content
    if decision:
        visible += "\n" + decision["question"]
        visible += "\n" + "\n".join(option["label"] for option in decision["options"])
    lower = visible.lower()
    term_hits = [term for term in INTERNAL_TERMS if term in lower]
    questions = visible_questions(response)
    produced_output = bool(response.get("artifact") or response.get("state_file"))

    if kind == "lesson":
        consequential_concepts = (
            "students",
            "already know",
            "prior",
            "time",
            "minutes",
            "class",
            "goal",
            "able to",
            "struggle",
        )
        expected = (
            "At least one question or explicit assumption should establish the learning goal, "
            "students' starting point, or feasible class context before a finished lesson is produced."
        )
    else:
        consequential_concepts = (
            "graded",
            "grade",
            "practice",
            "formative",
            "feedback",
            "minutes",
            "time",
            "resources",
            "notes",
            "students",
            "already know",
        )
        expected = (
            "Before consequential scoring is designed, ask whether the quiz is practice or graded "
            "and establish the intended evidence, time, and allowed resources; a clearly labeled "
            "formative draft may proceed with reversible assumptions."
        )

    question_text = " ".join(questions).lower()
    content_has_assumption = "assum" in lower or "provisional" in lower
    relevant_question = any(term in question_text for term in consequential_concepts)
    asks_reasonably = 1 <= len(questions) <= 3 and relevant_question
    if produced_output:
        interaction_ok = asks_reasonably or content_has_assumption
    else:
        interaction_ok = asks_reasonably
    status = "pass" if not term_hits and interaction_ok else "fail"
    return {
        "status": status,
        "questions": questions,
        "question_count": len(questions),
        "relevant_question_detected": relevant_question,
        "produced_artifact_or_state": produced_output,
        "visible_assumption_or_provisional_label": content_has_assumption,
        "internal_term_hits": term_hits,
        "expected_behavior": expected,
    }


def reaudit_existing() -> int:
    output = RESULTS / "missing-data-interaction-probes.json"
    result = json.loads(output.read_text(encoding="utf-8"))
    result["lesson_probe"]["audit"] = audit_probe(
        "lesson", result["lesson_probe"]["response"]
    )
    result["quiz_probe"]["audit"] = audit_probe(
        "quiz", result["quiz_probe"]["response"]
    )
    output.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "output": str(output),
                "lesson": result["lesson_probe"]["audit"],
                "quiz": result["quiz_probe"]["audit"],
            },
            indent=2,
        )
    )
    return 0


def main() -> int:
    RESULTS.mkdir(parents=True, exist_ok=True)
    client = TestClient(app)

    lesson_project = make_project(
        client,
        "Professor interaction probe — loops lesson",
        "Students should understand loops.",
    )
    lesson_prompt = "I want to teach about loops."
    lesson_response = ask(client, lesson_project, lesson_prompt)

    quiz_project = make_project(
        client,
        "Professor interaction probe — loop-invariant quiz",
        "Students should understand what stays true while a loop runs.",
    )
    quiz_prompt = "I want to create a quiz about invariants inside a loop."
    quiz_response = ask(client, quiz_project, quiz_prompt)

    result = {
        "test_purpose": (
            "Check whether Co-design requests genuinely necessary professor knowledge in "
            "plain teaching language while keeping implementation details hidden."
        ),
        "lesson_probe": {
            "project_id": lesson_project,
            "prompt": lesson_prompt,
            "response": lesson_response,
            "audit": audit_probe("lesson", lesson_response),
        },
        "quiz_probe": {
            "project_id": quiz_project,
            "prompt": quiz_prompt,
            "response": quiz_response,
            "audit": audit_probe("quiz", quiz_response),
        },
        "interaction_standard": {
            "ask_from_professor": [
                "What students should be able to do",
                "What students already know or where they struggle",
                "Available class time, setting, and essential constraints",
                "For a quiz: practice versus graded use, evidence desired, time, and allowed resources",
                "Grading or policy choices only when they become consequential",
            ],
            "safe_to_assume_and_label": [
                "Example context or dataset",
                "A reversible activity format",
                "Draft wording and visual organization",
                "A formative, zero-point draft when explicitly labeled provisional",
            ],
            "never_ask_professor": [
                "Markdown formatting",
                "SKILL routes or files",
                "JSON schemas or fenced blocks",
                "Validator commands",
                "Project-state filenames",
            ],
        },
    }
    output = RESULTS / "missing-data-interaction-probes.json"
    output.write_text(json.dumps(result, indent=2), encoding="utf-8")
    print(
        json.dumps(
            {
                "output": str(output),
                "lesson": result["lesson_probe"]["audit"],
                "quiz": result["quiz_probe"]["audit"],
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    if "--reaudit" in sys.argv:
        raise SystemExit(reaudit_existing())
    raise SystemExit(main())
