from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from server import (
    _create_tool_artifact,
    _database_connection,
    _ensure_project,
    _get_project_workspace,
    _materialize_artifact_file,
    _persist_exchange,
    _utc_now,
    get_settings,
)


PROJECT_ID = "project-office-artifact-demo"


SPECS = [
    {
        "kind": "slides",
        "title": "Active Learning Strategy Comparison",
        "subtitle": "A faculty discussion deck for choosing an instructional approach",
        "sections": [
            {
                "heading": "Begin with the learning evidence",
                "body": "Choose the strategy that gives students the clearest opportunity to practice the intended performance.",
                "bullets": [
                    "Name the observable learning outcome",
                    "Identify acceptable evidence of learning",
                    "Check feasibility for the available class time",
                ],
            },
            {
                "heading": "Compare two workable approaches",
                "body": "Use common criteria instead of selecting a strategy by familiarity alone.",
                "bullets": [
                    "Case-based discussion: interpretation, judgment, and justification",
                    "Structured peer instruction: explanation, feedback, and revision",
                    "Instructor review remains the disciplinary quality gate",
                ],
            },
            {
                "heading": "Make the decision traceable",
                "body": "Record the chosen strategy, rationale, assumptions, and the evidence that would cause the instructor to revise the choice.",
                "bullets": [
                    "Decision: working instructional strategy",
                    "Evidence: connection to the outcome",
                    "Open item: what must still be confirmed",
                ],
            },
        ],
        "source_ids": [],
    },
    {
        "kind": "document",
        "title": "Instructor Facilitation Guide",
        "subtitle": "Active-learning strategy comparison",
        "sections": [
            {
                "heading": "Purpose",
                "body": "Use this guide to facilitate a short, evidence-centered comparison of two instructional strategies without treating model output as disciplinary authority.",
                "bullets": [
                    "Keep the learning outcome visible throughout the discussion",
                    "Ask participants to distinguish confirmed information from assumptions",
                    "Record the final decision and the reason for it",
                ],
            },
            {
                "heading": "Facilitation sequence",
                "body": "Move from the desired student performance to evidence, activity mechanics, access considerations, and implementation risk.",
                "checklist": [
                    "The strategy creates observable evidence of the learning outcome",
                    "Students receive a meaningful opportunity to practice and revise",
                    "Materials and participation routes are accessible",
                    "The plan fits the actual class time and staffing",
                ],
                "table": {
                    "headers": ["Decision lens", "Instructor check"],
                    "rows": [
                        ["Alignment", "What student action demonstrates the outcome?"],
                        ["Evidence", "What work product or performance will be reviewed?"],
                        ["Feasibility", "What constraints could prevent implementation?"],
                    ],
                },
            },
        ],
        "source_ids": [],
    },
    {
        "kind": "worksheet",
        "title": "Student Evidence Comparison Worksheet",
        "subtitle": "Compare, justify, and reflect",
        "sections": [
            {
                "heading": "Identify the two approaches",
                "body": "Write the names of the approaches you are comparing and describe the decision context in one sentence.",
                "prompts": [
                    "What are the two approaches?",
                    "What decision must be made?",
                ],
                "response_lines": 2,
            },
            {
                "heading": "Compare the evidence",
                "body": "Use the same criteria for both approaches. Distinguish evidence from assumptions.",
                "table": {
                    "headers": ["Criterion", "Approach A", "Approach B"],
                    "rows": [
                        ["Alignment to outcome", "", ""],
                        ["Quality of evidence", "", ""],
                        ["Feasibility", "", ""],
                    ],
                },
                "prompts": ["Which comparison is best supported by evidence?"],
                "response_lines": 3,
            },
            {
                "heading": "Make and qualify your recommendation",
                "checklist": [
                    "I stated a clear recommendation",
                    "I connected the recommendation to evidence",
                    "I identified an assumption or limitation",
                ],
                "prompts": ["What do you recommend, and what evidence supports it?"],
                "response_lines": 4,
            },
        ],
        "source_ids": [],
    },
]


def main() -> None:
    settings = get_settings()
    _ensure_project(settings, PROJECT_ID)
    now = _utc_now()
    with _database_connection(settings) as connection:
        connection.execute(
            """
            UPDATE projects
            SET name = ?, course_name = ?, level = ?, class_time = ?, outcome = ?, mode = ?, notes = ?, updated_at = ?
            WHERE id = ?
            """,
            (
                "Office Artifact Demo",
                "Faculty Design Studio",
                "Professional",
                "75 minutes",
                "Compare two instructional strategies and justify a recommendation using explicit evidence.",
                "Guided",
                "Demonstration project containing locally generated PowerPoint and Word artifacts.",
                now,
                PROJECT_ID,
            ),
        )

    with _database_connection(settings) as connection:
        connection.execute("DELETE FROM messages WHERE project_id = ?", (PROJECT_ID,))

    _, decision_message, _ = _persist_exchange(
        settings,
        PROJECT_ID,
        "Help me choose the instructional strategy that should anchor this 75-minute workshop.",
        "Choose the working strategy. The selection will be recorded in the project trace and applied consistently across the slide deck, instructor guide, and student worksheet.",
        {
            "decision": {
                "question": "Which strategy should anchor the 75-minute workshop?",
                "options": [
                    {
                        "label": "Case-based discussion (Recommended)",
                        "description": "Students compare evidence, make a judgment, and justify a recommendation.",
                        "value": "case-based discussion",
                    },
                    {
                        "label": "Structured peer instruction",
                        "description": "Students explain an initial choice, receive peer feedback, and revise.",
                        "value": "structured peer instruction",
                    },
                ],
            },
            "skill_runtime": {
                "profile": "design",
                "loaded_files": ["SKILL.md", "references/interaction-protocol.md"],
                "fingerprint": "local-office-demo-decision",
            },
            "sources_used": [],
        },
        "design",
    )
    _, artifact_message, _ = _persist_exchange(
        settings,
        PROJECT_ID,
        "Use case-based discussion as the working strategy.",
        "Case-based discussion is now the recorded working decision. The artifact family applies the same evidence-centered rationale across the PowerPoint, instructor guide, and student worksheet.",
        {
            "skill_runtime": {
                "profile": "artifact",
                "loaded_files": ["SKILL.md", "references/artifact-patterns.md"],
                "fingerprint": "local-office-demo-artifacts",
            },
            "sources_used": [],
        },
        "design",
        "Selected: Case-based discussion (Recommended)\nStudents compare evidence, make a judgment, and justify a recommendation.",
        {
            "origin_message_id": decision_message["id"],
            "question": "Which strategy should anchor the 75-minute workshop?",
            "selected_label": "Case-based discussion (Recommended)",
            "selected_value": "case-based discussion",
        },
    )

    existing_by_title = {
        artifact["title"]: artifact for artifact in _get_project_workspace(settings, PROJECT_ID)["artifacts"]
    }
    created = []
    refreshed = []
    for spec in SPECS:
        if spec["title"] in existing_by_title:
            refreshed.append(
                _materialize_artifact_file(
                    settings,
                    PROJECT_ID,
                    existing_by_title[spec["title"]]["id"],
                    spec,
                )
            )
            continue
        created.append(_create_tool_artifact(settings, PROJECT_ID, spec))

    with _database_connection(settings) as connection:
        connection.execute(
            "UPDATE artifacts SET message_id = ? WHERE project_id = ?",
            (artifact_message["id"], PROJECT_ID),
        )
        connection.execute(
            "UPDATE projects SET updated_at = ? WHERE id = ?", (_utc_now(), PROJECT_ID)
        )
    final_workspace = _get_project_workspace(settings, PROJECT_ID)
    print(
        json.dumps(
            {
                "project_id": PROJECT_ID,
                "project_name": final_workspace["project"]["name"],
                "artifact_count": len(final_workspace["artifacts"]),
                "created": [artifact["title"] for artifact in created],
                "refreshed": [artifact["title"] for artifact in refreshed],
                "formats": [artifact["file_format"] for artifact in final_workspace["artifacts"]],
            },
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
