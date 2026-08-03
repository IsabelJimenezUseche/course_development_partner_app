import socket
import zipfile
from io import BytesIO

from docx import Document
from fastapi.testclient import TestClient
from pptx import Presentation
import server

from server import (
    _apply_auto_decision,
    _collaboration_mode_instruction,
    _extract_decision,
    _extract_artifact_spec,
    _infer_skill_profile,
    _hashed_text_vector,
    _load_skill_runtime,
    _persist_exchange,
    _render_markdown,
    _vector_cosine,
    app,
    find_available_port,
    get_settings,
    ChatMessage,
)


client = TestClient(app)


def test_local_vector_retrieval_favors_relevant_document_chunk():
    query = _hashed_text_vector("scaled dot product attention queries keys values")
    relevant = _hashed_text_vector(
        "Scaled dot product attention compares queries with keys and combines the values."
    )
    distractor = _hashed_text_vector(
        "The course registration calendar lists deadlines for adding and dropping classes."
    )

    assert _vector_cosine(query, relevant) > _vector_cosine(query, distractor)


def test_health_does_not_expose_api_key(monkeypatch):
    secret = "test-secret-that-must-not-appear"
    monkeypatch.setenv("PURDUE_GENAI_API_KEY", secret)

    response = client.get("/health")

    assert response.status_code == 200
    assert response.json()["api_key_configured"] is True
    assert secret not in response.text


def test_safe_config_does_not_expose_api_key(monkeypatch):
    secret = "another-test-secret"
    monkeypatch.setenv("PURDUE_GENAI_API_KEY", secret)

    response = client.get("/api/config")

    assert response.status_code == 200
    assert response.json()["api_key_configured"] is True
    assert secret not in response.text


def test_root_serves_faculty_ui():
    response = client.get("/")

    assert response.status_code == 200
    assert "Course Development Partner" in response.text
    assert "Design decisions that stay connected" in response.text
    assert "Generated artifacts" in response.text


def test_skill_runtime_loads_actual_skill_files():
    response = client.get("/api/skill", params={"profile": "assessment"})

    assert response.status_code == 200
    runtime = response.json()
    assert runtime["name"] == "course-development-partner"
    assert runtime["loaded_files"] == [
        "SKILL.md",
        "references/assessment-quality.md",
        "references/artifact-patterns.md",
    ]
    assert len(runtime["fingerprint"]) == 16


def test_skill_runtime_adds_limited_model_reliability_checks():
    prompt, runtime = _load_skill_runtime(get_settings(), "artifact")

    assert "Reliability overlay for this limited-capability model" in prompt
    assert "displayed block totals equal the stated available time" in prompt
    assert "Cross-check every disciplinary example" in prompt
    assert runtime["reliability_overlay"]["target"] == "gpt-oss-120b"


def test_upload_extracts_and_lists_text_source(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))
    project_id = "testproject01"

    upload_response = client.post(
        f"/api/projects/{project_id}/sources",
        data={"data_classification_ack": "true"},
        files={
            "files": (
                "course-notes.txt",
                b"Students compare two reaction-rate models and justify a selection.",
                "text/plain",
            )
        },
    )

    assert upload_response.status_code == 200
    result = upload_response.json()
    assert result["errors"] == []
    assert result["sources"][0]["source_id"].startswith("SRC-")
    assert result["sources"][0]["filename"] == "course-notes.txt"
    assert result["sources"][0]["vector_index"]["algorithm"] == "local-hashed-term-cosine-v1"
    assert result["sources"][0]["vector_index"]["dimensions"] == 384
    vector_path = (
        tmp_path
        / "projects"
        / project_id
        / "sources"
        / result["sources"][0]["source_id"]
        / "vectors.json"
    )
    assert vector_path.is_file()

    list_response = client.get(f"/api/projects/{project_id}/sources")
    assert list_response.status_code == 200
    assert len(list_response.json()["sources"]) == 1

    source_id = result["sources"][0]["source_id"]
    preview_response = client.get(f"/api/projects/{project_id}/sources/{source_id}/preview")
    assert preview_response.status_code == 200
    assert "reaction-rate models" in preview_response.json()["preview_text"]
    assert preview_response.json()["vector_chunks"][0]["main_idea"]
    assert "reaction-rate" in preview_response.json()["vector_chunks"][0]["keywords"]
    assert preview_response.json()["vector_metadata"]["metadata_algorithm"] == "extractive-main-idea-v3"

    download_response = client.get(f"/api/projects/{project_id}/sources/{source_id}/download")
    assert download_response.status_code == 200
    assert download_response.content == b"Students compare two reaction-rate models and justify a selection."
    assert "course-notes.txt" in download_response.headers["content-disposition"]


def test_upload_requires_data_confirmation(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))

    response = client.post(
        "/api/projects/testproject02/sources",
        data={"data_classification_ack": "false"},
        files={"files": ("notes.txt", b"Public course notes", "text/plain")},
    )

    assert response.status_code == 400
    assert "FERPA" in response.json()["detail"]


def test_chat_requires_api_key(monkeypatch):
    monkeypatch.setenv("PURDUE_GENAI_API_KEY", "")
    monkeypatch.setenv("PURDUE_GENAI_MODEL_ID", "gpt-oss-120b")

    response = client.post(
        "/api/chat",
        json={"messages": [{"role": "user", "content": "Hello"}]},
    )

    assert response.status_code == 503
    assert "PURDUE_GENAI_API_KEY" in response.json()["detail"]


def test_port_fallback_skips_occupied_port():
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as occupied:
        occupied.bind(("127.0.0.1", 0))
        occupied.listen(1)
        occupied_port = occupied.getsockname()[1]

        selected = find_available_port("127.0.0.1", occupied_port, 2)

    assert selected == occupied_port + 1


def test_project_workspace_persists_context(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))
    create_response = client.post(
        "/api/projects",
        json={
            "name": "CHE 205 redesign",
            "course_name": "Reaction engineering",
            "outcome": "Analyze rate data.",
            "mode": "Guided",
        },
    )

    assert create_response.status_code == 201
    project_id = create_response.json()["project"]["id"]

    update_response = client.patch(
        f"/api/projects/{project_id}",
        json={"class_time": "75 minutes", "notes": "Teams of four"},
    )
    workspace_response = client.get(f"/api/projects/{project_id}")

    assert update_response.status_code == 200
    assert workspace_response.status_code == 200
    workspace = workspace_response.json()
    assert workspace["project"]["name"] == "CHE 205 redesign"
    assert workspace["project"]["class_time"] == "75 minutes"
    assert workspace["project"]["notes"] == "Teams of four"
    assert workspace["messages"] == []
    assert workspace["artifacts"] == []


def test_markdown_renderer_supports_tables_and_sanitizes_html():
    rendered = _render_markdown(
        "# Artifact\n\n| Outcome | Evidence |\n|---|---|\n| Analyze | Explanation |"
        "\n\n<script>alert('unsafe')</script>"
    )

    assert "<h1>Artifact</h1>" in rendered
    assert "<table>" in rendered
    assert "<script>" not in rendered


def test_structured_decision_is_removed_and_parsed():
    content, decision = _extract_decision(
        "Choose an evidence pattern.\n\n```decision\n"
        '{"question":"Which pattern?","options":['
        '{"label":"Performance (Recommended)","description":"Authentic evidence",'
        '"value":"Use a performance task"},'
        '{"label":"Written explanation","description":"Fast to review",'
        '"value":"Use a written explanation"}]}\n```'
    )

    assert content == "Choose an evidence pattern."
    assert decision["question"] == "Which pattern?"
    assert len(decision["options"]) == 2


def test_auto_mode_accepts_project_setting_and_resolves_choice_without_modal(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))
    project_response = client.post("/api/projects", json={"name": "Auto project", "mode": "Auto"})
    assert project_response.status_code == 201
    project = project_response.json()["project"]
    assert project["mode"] == "Auto"
    assert "do not ask the educator a question" in _collaboration_mode_instruction("Auto")

    content, interactive_decision = _extract_decision(
        "Select a path.\n\n```decision\n"
        '{"question":"Which path?","options":['
        '{"label":"A (Recommended)","description":"Best aligned","value":"a"},'
        '{"label":"B","description":"Alternative","value":"b"}]}'
        "\n```"
    )
    resolved, remaining_decision, auto_decision = _apply_auto_decision(content, interactive_decision)
    assert remaining_decision is None
    assert auto_decision["selected_label"] == "A (Recommended)"
    assert "Auto decision recorded" in resolved
    assert "Best aligned" in resolved


def test_question_list_falls_back_to_choice_dialog():
    content, decision = _extract_decision(
        "Please choose:\n"
        "1. **Which strategies should students compare?** "
        "(e.g., case-based learning vs. lecture, problem-based learning vs. flipped classroom)"
    )

    assert content.startswith("Please choose")
    assert decision["question"] == "Which strategies should students compare?"
    assert decision["options"][0]["label"].endswith("(Recommended)")
    assert len(decision["options"]) == 3


def test_generic_json_fence_is_accepted_for_decision_schema():
    content, decision = _extract_decision(
        "Choose a direction.\n```json\n"
        '{"question":"Which direction?","options":['
        '{"label":"A (Recommended)","description":"Use A","value":"A"},'
        '{"label":"B","description":"Use B","value":"B"}]}\n```'
    )

    assert content == "Choose a direction."
    assert decision["question"] == "Which direction?"


def test_explicit_artifact_request_wins_over_review_keyword():
    profile = _infer_skill_profile(
        [ChatMessage(role="user", content="Create an artifact with an instructor review note")]
    )

    assert profile == "artifact"


def test_open_decision_is_not_saved_as_finished_artifact(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))
    project_id = "project-decision-test"

    _, assistant, artifact = _persist_exchange(
        get_settings(),
        project_id,
        "Create a worksheet artifact",
        "Which worksheet pattern should we use?",
        {
            "decision": {
                "question": "Which pattern?",
                "options": [
                    {"label": "A", "description": "A", "value": "A"},
                    {"label": "B", "description": "B", "value": "B"},
                ],
            }
        },
        "artifact",
    )

    assert assistant["decision"] is not None
    assert artifact is None
    assert client.get(f"/api/projects/{project_id}").json()["artifacts"] == []


def test_specific_message_download_includes_skill_route(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))
    project_id = "project-message-download"
    _, assistant, _ = _persist_exchange(
        get_settings(),
        project_id,
        "Draft a table",
        "## Comparison\n\n| A | B |\n|---|---|\n| 1 | 2 |",
        {
            "skill_runtime": {
                "profile": "artifact",
                "loaded_files": ["SKILL.md", "references/artifact-patterns.md"],
                "fingerprint": "testfingerprint01",
            },
            "sources_used": [],
        },
        "artifact",
    )

    response = client.get(
        f"/api/projects/{project_id}/messages/{assistant['id']}/download?format=markdown"
    )

    assert response.status_code == 200
    assert "SKILL profile: artifact" in response.text
    assert "references/artifact-patterns.md" in response.text
    assert "## Comparison" in response.text


def test_skill_trace_links_decision_to_selected_answer(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))
    project_id = "project-trace-test"
    _, decision_message, _ = _persist_exchange(
        get_settings(),
        project_id,
        "Help me choose",
        "Choose an evidence pattern.",
        {
            "decision": {
                "question": "Which evidence pattern?",
                "options": [
                    {"label": "Performance", "description": "Authentic", "value": "performance"},
                    {"label": "Explanation", "description": "Written", "value": "explanation"},
                ],
            },
            "skill_runtime": {
                "profile": "design",
                "loaded_files": ["SKILL.md", "references/design-workflow.md"],
                "fingerprint": "tracefingerprint1",
            },
        },
        "design",
    )
    _persist_exchange(
        get_settings(),
        project_id,
        "Use a performance task",
        "The performance task is now the working evidence pattern.",
        {
            "skill_runtime": {
                "profile": "design",
                "loaded_files": ["SKILL.md", "references/design-workflow.md"],
                "fingerprint": "tracefingerprint1",
            }
        },
        "design",
        "Selected: Performance",
        {
            "origin_message_id": decision_message["id"],
            "question": "Which evidence pattern?",
            "selected_label": "Performance",
            "selected_value": "performance",
        },
    )

    response = client.get(f"/api/projects/{project_id}/trace")

    assert response.status_code == 200
    trace = response.json()
    assert trace["summary"]["decisions"] == 1
    assert trace["summary"]["answers"] == 1
    answer_event = next(event for event in trace["events"] if event["type"] == "answer")
    assert answer_event["decision_trace"]["origin_message_id"] == decision_message["id"]
    assert answer_event["decision_trace"]["selected_label"] == "Performance"


def test_project_export_downloads_markdown(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))
    project = client.post("/api/projects", json={"name": "Export test"}).json()["project"]

    response = client.get(f"/api/projects/{project['id']}/export?format=markdown")

    assert response.status_code == 200
    assert response.headers["content-type"].startswith("text/markdown")
    assert "attachment" in response.headers["content-disposition"]
    assert "# Export test" in response.text


def test_artifact_spec_contract_is_extracted_and_hidden_from_chat():
    content = """Your worksheet is ready.

```artifact_spec
{"kind":"worksheet","title":"Evidence Check","subtitle":"Practice","sections":[{"heading":"Claim","body":"","bullets":[],"prompts":["What is the strongest evidence?"],"checklist":[],"response_lines":3,"table":null}],"source_ids":[]}
```"""

    cleaned, spec = _extract_artifact_spec(content)

    assert cleaned == "Your worksheet is ready."
    assert spec["kind"] == "worksheet"
    assert spec["sections"][0]["prompts"] == ["What is the strongest evidence?"]


def test_local_artifact_tools_create_slides_document_and_worksheet(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))
    project = client.post(
        "/api/projects",
        json={"name": "Office tool test", "course_name": "EDCI  Test Studio"},
    ).json()["project"]
    expected_formats = {"slides": "pptx", "document": "docx", "worksheet": "docx"}

    for kind, file_format in expected_formats.items():
        response = client.post(
            f"/api/projects/{project['id']}/artifact-tools/generate",
            json={
                "kind": kind,
                "title": f"{kind.title()} smoke test",
                "subtitle": "Deterministic artifact harness",
                "sections": [
                    {
                        "heading": "Compare the evidence",
                        "body": "Use the supplied criteria to make an instructor-reviewed choice.",
                        "bullets": ["Identify the claim", "Connect evidence to the outcome"],
                        "prompts": ["Which evidence is most convincing?"] if kind == "worksheet" else [],
                        "checklist": ["Evidence is source-grounded"],
                        "response_lines": 2,
                        "table": {
                            "headers": ["Criterion", "Evidence"],
                            "rows": [["Alignment", "Directly supports the outcome"]],
                        },
                    }
                ],
                "source_ids": [],
            },
        )

        assert response.status_code == 200, response.text
        artifact = response.json()["artifact"]
        assert artifact["has_file"] is True
        assert artifact["file_format"] == file_format
        assert artifact["tool_trace"]["schema_validated"] is True
        assert "Compare the evidence" in artifact["html"]
        assert "Directly supports the outcome" in artifact["html"]
        download = client.get(
            f"/api/projects/{project['id']}/artifacts/{artifact['id']}/download?format=office"
        )
        assert download.status_code == 200
        assert len(download.content) > 1_000
        if kind == "slides":
            deck = Presentation(BytesIO(download.content))
            assert len(deck.slides) == 2
            assert "Slides smoke test" in deck.slides[0].shapes[0].text
            assert "[Sources]" in deck.slides[1].notes_slide.notes_text_frame.text
        else:
            document = Document(BytesIO(download.content))
            document_text = "\n".join(paragraph.text for paragraph in document.paragraphs)
            assert f"{kind.title()} smoke test" in document_text
            assert "Compare the evidence" in document_text
            if kind == "worksheet":
                assert "Which evidence is most convincing?" in document_text

    project_package = client.get(f"/api/projects/{project['id']}/export?format=zip")
    assert project_package.status_code == 200
    assert project_package.headers["content-type"].startswith("application/zip")
    with zipfile.ZipFile(BytesIO(project_package.content)) as archive:
        names = archive.namelist()
        assert "conversation.md" in names
        assert "project.json" in names
        assert any(name.endswith(".pptx") for name in names)
        assert sum(name.endswith(".docx") for name in names) == 2


def test_legacy_decision_json_is_hidden_when_workspace_reloads(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))
    project = client.post("/api/projects", json={"name": "Legacy decision"}).json()["project"]
    _persist_exchange(
        get_settings(),
        project["id"],
        "Help me choose",
        "Choose a direction.\n\n```json\n{\"question\":\"Which direction?\",\"options\":[{\"label\":\"A\",\"description\":\"First\",\"value\":\"a\"},{\"label\":\"B\",\"description\":\"Second\",\"value\":\"b\"}]}\n```",
        {"skill_runtime": {"profile": "design", "loaded_files": ["SKILL.md"], "fingerprint": "legacy"}},
        "design",
    )

    workspace = client.get(f"/api/projects/{project['id']}").json()
    assistant = workspace["messages"][-1]
    assert assistant["content"] == "Choose a direction."
    assert assistant["decision"]["question"] == "Which direction?"
    assert "```json" not in assistant["html"]


def test_artifact_tool_rejects_invented_source_ids(monkeypatch, tmp_path):
    monkeypatch.setenv("APP_DATA_DIR", str(tmp_path))
    project = client.post("/api/projects", json={"name": "Source guard"}).json()["project"]

    response = client.post(
        f"/api/projects/{project['id']}/artifact-tools/generate",
        json={
            "kind": "document",
            "title": "Unsafe source test",
            "sections": [{"heading": "Claim", "body": "A claim"}],
            "source_ids": ["SRC-FFFFFFFFFFFF"],
        },
    )

    assert response.status_code == 400
    assert "unknown project sources" in response.json()["detail"]


def test_settings_masks_api_key_and_updates_selected_env(monkeypatch, tmp_path):
    env_paths = {
        "home": tmp_path / "home.env",
        "app": tmp_path / "app.env",
        "current": tmp_path / "current.env",
    }
    monkeypatch.setattr(server, "_environment_paths", lambda: env_paths)
    monkeypatch.setenv("PURDUE_GENAI_API_KEY", "must-not-be-returned")

    get_response = client.get("/api/settings")
    patch_response = client.patch(
        "/api/settings",
        json={
            "target": "current",
            "values": {
                "PURDUE_GENAI_MODEL_ID": "gpt-oss:120b",
                "APP_PORT": "8001",
            },
        },
    )

    assert get_response.status_code == 200
    assert "must-not-be-returned" not in get_response.text
    api_key = next(
        field for field in get_response.json()["fields"]
        if field["name"] == "PURDUE_GENAI_API_KEY"
    )
    assert api_key["configured"] is True
    assert api_key["value"] == ""
    assert patch_response.status_code == 200
    assert patch_response.json()["restart_required"] is True
    assert "PURDUE_GENAI_MODEL_ID='gpt-oss:120b'" in env_paths["current"].read_text()
