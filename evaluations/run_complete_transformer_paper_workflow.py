"""Run a complete source-grounded Transformer paper Co-design workflow.

The workflow downloads the official NeurIPS 2017 paper at runtime, uploads it through
the same source endpoint used by the browser, and then simulates a professor planning a
lesson, producing a Word guide and PowerPoint, creating and revising a Word quiz, and
requesting a handoff. A source upload by itself is not a pass: the source must appear in
retrieval traces and Office-artifact traces.
"""

from __future__ import annotations

import json
import os
import re
import tempfile
from io import BytesIO
import sys
from pathlib import Path
from zipfile import ZipFile, is_zipfile

import httpx

sys.path.insert(0, str(Path(__file__).resolve().parent))
from realistic_arc import count_questions  # noqa: E402
from pypdf import PdfReader


APP_URL = os.getenv("APP_URL", "http://127.0.0.1:8001").rstrip("/")
PAPER_URL = (
    "https://proceedings.neurips.cc/paper_files/paper/2017/file/"
    "3f5ee243547dee91fbd053c1c4a845aa-Paper.pdf"
)
ROOT = Path(__file__).resolve().parents[1]
RESULTS = ROOT / "evaluations" / "results" / "transformer-paper-workflow-20260803"
INTERNAL_TERMS = (
    "skill.md",
    "state_file",
    "artifact_spec",
    "json schema",
    "validator command",
    "project-state filename",
)
LATER_MODEL_PATTERN = re.compile(
    r"\b(?:bert|gpt(?:-\d+)?|chatgpt|vision transformer)\b", re.IGNORECASE
)


def compact(content: str) -> str:
    return re.sub(r"[\s\u00a0\u202f]+", " ", content).lower()


def technical_text(content: str) -> str:
    return (
        compact(content)
        .replace("₁", "1")
        .replace("₂", "2")
        .replace("₃", "3")
        .replace("→", "->")
    )


def has_incorrect_paper_location(content: str) -> bool:
    text = technical_text(content)
    positional_as_equation_one = bool(
        re.search(
            r"(?:positional|sinusoidal)[^.\n]{0,100}(?:eq(?:uation)?\.?\s*\(?1\)?|"
            r"eq(?:uation)?\.?\s*\(?1\)?[^.\n]{0,100}(?:positional|sinusoidal))",
            text,
        )
    )
    figure_two_as_mask = bool(
        re.search(
            r"(?:mask[^.\n]{0,100}figure\s*2|figure\s*2[^.\n]{0,100}mask)",
            text,
        )
    )
    attention_as_equation_two = bool(
        re.search(
            r"(?:scaled\s+dot-product|attention)[^.\n]{0,100}"
            r"eq(?:uation)?\.?\s*\(?2\)?",
            text,
        )
    )
    return positional_as_equation_one or figure_two_as_mask or attention_as_equation_two


def has_unsupported_paper_claim(content: str) -> bool:
    text = technical_text(content)
    return "proof-of-concept" in text or bool(
        re.search(
            r"learned (?:positional )?embeddings[^.\n]{0,100}"
            r"(?:cannot|can't|do not|don't) extrapolate",
            text,
        )
    )


def has_incorrect_transformer_instruction(content: str) -> bool:
    text = technical_text(content)
    if re.search(r"three\s+sub-?layers?\s+(?:per|in(?: the)?)\s+encoder", text):
        return True
    if "zeros on and above the diagonal" in text:
        return True
    # Table 1 is the complexity comparison, not an ablation. Fire only when a sentence
    # actually conflates them: a sentence that draws the distinction ("Table 1 is a
    # theoretical comparison, not an experimental ablation") is correct teaching and
    # must not be scored as an error. Mirrors the negation guard on multiply_mask below.
    for sentence in re.split(r"(?<=[.!?])\s+|\n+", text):
        conflates_table_one_with_ablation = re.search(
            r"ablation[^.\n]{0,100}table\s*1|table\s*1[^.\n]{0,100}ablation", sentence
        )
        if conflates_table_one_with_ablation and not re.search(
            r"(?:not|without|never|rather than|instead of)[^.\n]{0,40}ablation", sentence
        ):
            return True
    binary_mask_misused_as_additive = (
        re.search(r"mask matrix\s+m[^.\n]{0,100}(?:1.?s|ones)[^.\n]{0,100}(?:0.?s|zeros)", text)
        and re.search(r"add\s+m[^.\n]{0,100}(?:negative|−∞|-∞)", text)
    )
    if binary_mask_misused_as_additive:
        return True
    for sentence in re.split(r"(?<=[.!?])\s+|\n+", text):
        multiply_mask = re.search(
            r"(?:multiply[^.\n]{0,60}(?:mask|\bm\b)|"
            r"(?:mask|\bm\b)[^.\n]{0,60}multiply)",
            sentence,
        )
        if multiply_mask and not any(
            marker in sentence for marker in ("do not multiply", "not multiplied")
        ):
            return True
    return False


def has_later_model_claim(content: str) -> bool:
    """Allow an explicit scope boundary, but reject teaching claims about successors."""
    for sentence in re.split(r"(?<=[.!?])\s+|\n+", content):
        if not LATER_MODEL_PATTERN.search(sentence):
            continue
        lower = compact(sentence)
        boundary_markers = (
            "does not discuss",
            "do not discuss",
            "not discussed",
            "not covered",
            "not in the paper",
            "outside the paper",
            "out of scope",
        )
        if not any(marker in lower for marker in boundary_markers):
            return True
    return False


def create_project(client: httpx.Client) -> str:
    response = client.post(
        "/api/projects",
        json={
            "name": "Transformer architecture — complete paper-grounded Co-design workflow",
            "course_name": "Neural Networks and Sequence Modeling",
            "level": "Senior undergraduate and first-year graduate",
            "class_time": "90 minutes",
            "outcome": (
                "Students trace information through the original Transformer and explain "
                "how its attention, masking, and positional information support sequence modeling."
            ),
            "mode": "Co-design",
            "notes": (
                "Use the uploaded NeurIPS 2017 paper as the authority. Distinguish claims "
                "reported in the paper from later Transformer developments."
            ),
        },
    )
    response.raise_for_status()
    return response.json()["project"]["id"]


def download_paper(destination: Path) -> dict:
    with httpx.Client(follow_redirects=True, timeout=120.0) as client:
        response = client.get(PAPER_URL)
        response.raise_for_status()
    content_type = response.headers.get("content-type", "")
    if not response.content.startswith(b"%PDF"):
        raise RuntimeError("The NeurIPS URL did not return a PDF")
    destination.write_bytes(response.content)
    pages = len(PdfReader(BytesIO(response.content)).pages)
    return {
        "url": PAPER_URL,
        "filename": destination.name,
        "bytes": len(response.content),
        "pages": pages,
        "content_type": content_type,
    }


def upload_paper(client: httpx.Client, project_id: str, paper_path: Path) -> dict:
    with paper_path.open("rb") as paper:
        response = client.post(
            f"/api/projects/{project_id}/sources",
            data={"data_classification_ack": "true"},
            files={"files": (paper_path.name, paper, "application/pdf")},
        )
    response.raise_for_status()
    payload = response.json()
    if payload.get("errors") or len(payload.get("sources", [])) != 1:
        raise RuntimeError(f"Paper upload failed: {payload}")
    return payload["sources"][0]


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
        "sources_used": response.get("sources_used") or [],
    }


def artifact_source_ids(response: dict) -> list[str]:
    artifact = response.get("artifact") or {}
    return (artifact.get("tool_trace") or {}).get("source_ids") or []


def lesson_artifact_ready(response: dict, source_id: str) -> bool:
    artifact = response.get("artifact") or {}
    content = compact(artifact.get("content", ""))
    required = ("encoder", "decoder", "attention", "quer", "key", "value", "position")
    return bool(
        artifact.get("has_file")
        and artifact.get("file_format") == "docx"
        and source_id in artifact_source_ids(response)
        and all(term in content for term in required)
        and not has_incorrect_paper_location(artifact.get("content", ""))
        and not has_unsupported_paper_claim(artifact.get("content", ""))
        and not has_incorrect_transformer_instruction(artifact.get("content", ""))
    )


def slides_artifact_ready(response: dict, source_id: str) -> bool:
    artifact = response.get("artifact") or {}
    content = compact(artifact.get("content", ""))
    required = ("encoder", "decoder", "attention", "mask", "position", "evidence")
    return bool(
        artifact.get("has_file")
        and artifact.get("file_format") == "pptx"
        and source_id in artifact_source_ids(response)
        and all(term in content for term in required)
        and not has_incorrect_paper_location(artifact.get("content", ""))
        and not has_unsupported_paper_claim(artifact.get("content", ""))
        and not has_incorrect_transformer_instruction(artifact.get("content", ""))
    )


def revised_quiz_ready(response: dict, source_id: str) -> bool:
    artifact = response.get("artifact") or {}
    content = compact(artifact.get("content", ""))
    technical = technical_text(artifact.get("content", ""))
    masking = "mask" in content and ("future" in content or "subsequent" in content)
    correct_prohibited_links = all(
        re.search(
            rf"(?:{left}\s*->\s*{right}[^.\n]{{0,80}}prohibited|"
            rf"prohibited[^.\n]{{0,240}}{left}\s*->\s*{right})",
            technical,
        )
        for left, right in (("t1", "t2"), ("t1", "t3"), ("t2", "t3"))
    )
    # Accept either order, as the prohibited-link check above already does. A quiz that
    # labels the group first ("Allowed links: T2 -> T1 and T3 -> T2") is answering
    # correctly; requiring the label to trail the link rejected correct work.
    correct_allowed_links = all(
        re.search(
            rf"(?:{left}\s*->\s*{right}[^.\n]{{0,80}}(?:allowed|permitted)|"
            rf"(?:allowed|permitted)[^.\n]{{0,240}}{left}\s*->\s*{right})",
            technical,
        )
        for left, right in (("t2", "t1"), ("t3", "t2"))
    )
    attention_calculation_complete = (
        "1.34" in technical
        and "0.99" in technical
        and "0.67" in technical
        and "0.33" in technical
        and re.search(r"q[^.\n]{0,20}k1[^.\n]{0,20}=\s*1", technical)
        and re.search(r"q[^.\n]{0,20}k2[^.\n]{0,20}=\s*0", technical)
        and (
            "equation 1" in technical
            or "section 3.2.1" in technical
            or "section 3.2" in technical
        )
    )
    bounded_positional_claim = (
        "nearly identical" in technical
        and "may" in technical
        and "extrapolat" in technical
    )
    return bool(
        artifact.get("has_file")
        and artifact.get("file_format") == "docx"
        and source_id in artifact_source_ids(response)
        and "question 2" in content
        and "answer key" in content
        and "encoder" in content
        and "decoder" in content
        and masking
        and correct_prohibited_links
        and correct_allowed_links
        and attention_calculation_complete
        and bounded_positional_claim
        and not has_incorrect_paper_location(artifact.get("content", ""))
        and not has_unsupported_paper_claim(artifact.get("content", ""))
    )


def office_file_valid(client: httpx.Client, project_id: str, artifact: dict) -> bool:
    if not artifact.get("has_file"):
        return False
    response = client.get(
        f"/api/projects/{project_id}/artifacts/{artifact['id']}/download",
        params={"format": "office"},
    )
    if response.status_code != 200 or not is_zipfile(BytesIO(response.content)):
        return False
    required_prefix = "ppt/" if artifact.get("file_format") == "pptx" else "word/"
    with ZipFile(BytesIO(response.content)) as archive:
        return archive.testzip() is None and any(
            name.startswith(required_prefix) for name in archive.namelist()
        )


def main() -> int:
    RESULTS.mkdir(parents=True, exist_ok=True)
    turns: list[dict] = []
    with tempfile.TemporaryDirectory(prefix="transformer-paper-") as temp_dir:
        paper_path = Path(temp_dir) / "attention-is-all-you-need-neurips-2017.pdf"
        paper = download_paper(paper_path)

        with httpx.Client(base_url=APP_URL, timeout=300.0) as client:
            project_id = create_project(client)
            source = upload_paper(client, project_id, paper_path)
            source_id = source["source_id"]

            preview_response = client.get(
                f"/api/projects/{project_id}/sources/{source_id}/preview"
            )
            preview_response.raise_for_status()
            preview = preview_response.json()

            lesson_start = (
                "I want to teach the Transformer architecture using the paper I uploaded."
            )
            response, attempts = send(client, project_id, lesson_start)
            turns.append(response_record("lesson_start", lesson_start, response, attempts))

            lesson_context = (
                "These are senior undergraduates and first-year graduate students. They know "
                "linear algebra, softmax, backpropagation, and encoder-decoder RNNs, but they "
                "have not studied attention. I have 90 minutes with 30 students working in pairs. "
                "By the end, they should trace information through the original encoder and decoder, "
                "explain queries, keys, values, scaling, decoder masking, and positional encoding, "
                "and distinguish what the 2017 paper demonstrates from later claims."
            )
            response, attempts = send(client, project_id, lesson_context, response)
            turns.append(response_record("lesson_context", lesson_context, response, attempts))

            lesson_production = (
                "Create the complete 90-minute lesson as a downloadable Word instructor guide. "
                "Use the uploaded paper as the factual authority and cite it. Include a prediction "
                "before explanation, a guided trace of Figure 1, a small scaled dot-product attention "
                "calculation, pair work on masking and positional information, and an evidence check "
                "that separates reported results from interpretation. Cite decoder masking to "
                "Section 3.2.3 and Figure 1, not Figure 2. The positional-encoding formulas are in "
                "Section 3.5 and are not Equation 1. The encoder layer has two sublayers and the "
                "decoder layer has three. Add the causal mask to attention logits before softmax; "
                "do not multiply by the mask. A three-token allowed-region matrix has ones on and "
                "below the diagonal. Do not introduce later models."
            )
            lesson_response, attempts = send(
                client, project_id, lesson_production, response
            )
            turns.append(
                response_record(
                    "lesson_production", lesson_production, lesson_response, attempts
                )
            )
            lesson_retry = (
                "Complete the downloadable Word instructor guide now. It must explicitly cover "
                "the encoder, decoder, queries, keys, values, scaled attention, masking, and "
                "positional encoding, and it must identify the uploaded paper as its source. "
                "Use Section 3.2.3 and Figure 1 for decoder masking; Figure 2 is the attention "
                "mechanism diagram. Use Section 3.5 for positional encoding and do not label its "
                "formulas Equation 1. The encoder has two sublayers and the decoder has three. "
                "Add the causal mask to attention logits before softmax; do not multiply by it, "
                "and show allowed links on and below the diagonal. Table 1 compares theoretical "
                "complexity, sequential operations, and path length; it is not an ablation. Use "
                "only claims supported by the paper and ask no question."
            )
            for retry_index in range(1, 3):
                if lesson_artifact_ready(lesson_response, source_id):
                    break
                lesson_response, attempts = send(
                    client, project_id, lesson_retry, lesson_response
                )
                turns.append(
                    response_record(
                        f"lesson_production_retry_{retry_index}",
                        lesson_retry,
                        lesson_response,
                        attempts,
                    )
                )

            slides_request = (
                "Create a downloadable PowerPoint slide deck for this 90-minute lesson. Ground it "
                "in the uploaded paper and identify that source in the artifact. Include the "
                "original encoder-decoder architecture, scaled dot-product and multi-head attention, "
                "the correct causal mask, positional encoding, a worked pair prompt, and one slide "
                "that separates the paper's reported evidence from interpretation. Keep the encoder's "
                "two sublayers distinct from the decoder's three, add masks before softmax, cite "
                "Equation 1 for attention, Section 3.2.3 for masking, and Section 3.5 for position. "
                "If showing a binary allowed matrix, convert blocked zeros to an additive −∞ bias "
                "before softmax; do not add the 1/0 matrix itself. Describe Table 1 as a theoretical "
                "comparison, not an ablation."
            )
            slides_response, attempts = send(
                client, project_id, slides_request, lesson_response
            )
            turns.append(
                response_record("slides_production", slides_request, slides_response, attempts)
            )
            slides_retry = (
                "Complete the downloadable PowerPoint now with the full source-grounded slide "
                "content. It must cover encoder, decoder, attention, masking, positional encoding, "
                "and evidence boundaries; use the uploaded paper in the artifact source trace. "
                "Use Equation 1 for scaled attention, Section 3.2.3 and Figure 1 for masking, and "
                "Section 3.5 for positional encoding. The encoder has two sublayers and decoder "
                "three. Add an additive 0/−∞ mask to logits before softmax; do not add a binary "
                "1/0 matrix. Table 1 is a theoretical comparison, not an ablation. Ask no question."
            )
            for retry_index in range(1, 3):
                if slides_artifact_ready(slides_response, source_id):
                    break
                slides_response, attempts = send(
                    client, project_id, slides_retry, slides_response
                )
                turns.append(
                    response_record(
                        f"slides_production_retry_{retry_index}",
                        slides_retry,
                        slides_response,
                        attempts,
                    )
                )

            quiz_start = (
                "I also want a short quiz on how the original Transformer architecture works."
            )
            quiz_response, attempts = send(client, project_id, quiz_start)
            turns.append(response_record("quiz_start", quiz_start, quiz_response, attempts))

            quiz_context = (
                "Make it a 15-minute practice quiz for feedback, not a grade. Students may use the "
                "uploaded paper and a calculator. Use three questions that require tracing or "
                "explanation rather than recall, and include a concise instructor answer key with "
                "paper locations supporting the answers."
            )
            quiz_response, attempts = send(
                client, project_id, quiz_context, quiz_response
            )
            turns.append(response_record("quiz_context", quiz_context, quiz_response, attempts))

            quiz_production = (
                "Create the complete ready-to-use quiz as a downloadable Word document now. Ground "
                "all architecture claims in the uploaded paper, cite that source in the artifact, "
                "include reasoning space, and do not introduce later Transformer variants."
            )
            quiz_response, attempts = send(
                client, project_id, quiz_production, quiz_response
            )
            turns.append(
                response_record("quiz_production", quiz_production, quiz_response, attempts)
            )

            revision = (
                "Revise the downloadable Word quiz. Make Question 2 compare encoder self-attention "
                "with masked decoder self-attention using a short token sequence. Students should "
                "identify which attention links are prohibited and explain why. Update the answer "
                "key and cite Section 3.2.3. In the key, explicitly state that T1 → T2, T1 → T3, "
                "and T2 → T3 are prohibited in masked decoder self-attention; T2 → T1 and T3 → T2 "
                "are allowed. Keep a worked attention calculation whose output is approximately "
                "[1.34, 0.99]: use q=[1,0], k1=[1,1], k2=[0,2], v1=[2,0], v2=[0,3]; "
                "the dot products are 1 and 0 and the softmax weights are approximately 0.67 and "
                "0.33. Identify scaled dot-product attention as Equation 1. For positional "
                "encoding, accurately say the learned and sinusoidal versions produced nearly "
                "identical results and the authors chose sinusoids because they may allow "
                "extrapolation. Keep all three full questions."
            )
            revision_response, attempts = send(
                client, project_id, revision, quiz_response
            )
            turns.append(
                response_record("quiz_revision", revision, revision_response, attempts)
            )
            revision_retry = (
                "The revised file is missing content or source attribution. Recreate the complete "
                "three-question downloadable Word quiz now. Question 2 must compare encoder "
                "self-attention with masked decoder self-attention, prohibit attention to future "
                "positions, and explain the autoregressive reason. The answer key must literally "
                "state that T1 → T2, T1 → T3, and T2 → T3 are prohibited, while T2 → T1 and "
                "T3 → T2 are allowed. The worked attention answer must finish at approximately "
                "[1.34, 0.99] using q=[1,0], k1=[1,1], k2=[0,2], v1=[2,0], v2=[0,3]. "
                "Show q·k1=1, q·k2=0, and softmax weights about 0.67 and 0.33; cite Equation 1 "
                "or Section 3.2.1. Cite Section 3.2.3 for masking. "
                "For Section 3.5, say learned and sinusoidal positional encodings produced nearly "
                "identical results and the authors chose sinusoids because they may allow "
                "extrapolation; do not say learned embeddings cannot extrapolate. Identify the "
                "uploaded paper as the source and do not ask another question."
            )
            for retry_index in range(1, 3):
                if revised_quiz_ready(revision_response, source_id):
                    break
                revision_response, attempts = send(
                    client, project_id, revision_retry, revision_response
                )
                turns.append(
                    response_record(
                        f"quiz_revision_retry_{retry_index}",
                        revision_retry,
                        revision_response,
                        attempts,
                    )
                )

            handoff = (
                "Briefly summarize what is ready for class, which claims come directly from the "
                "uploaded paper, and what I still need to review. Use a short list."
            )
            handoff_response, attempts = send(client, project_id, handoff)
            turns.append(response_record("handoff", handoff, handoff_response, attempts))

            workspace_response = client.get(f"/api/projects/{project_id}")
            workspace_response.raise_for_status()
            workspace = workspace_response.json()
            office_checks = {
                artifact["id"]: office_file_valid(client, project_id, artifact)
                for artifact in workspace.get("artifacts", [])
                if artifact.get("has_file")
            }

    visible = "\n".join(turn["response"].get("content", "") for turn in turns)
    deliverables = "\n".join(
        [
            (lesson_response.get("artifact") or {}).get("content", ""),
            (slides_response.get("artifact") or {}).get("content", ""),
            (revision_response.get("artifact") or {}).get("content", ""),
        ]
    )
    lower = compact(f"{visible}\n{deliverables}")
    lesson_questions = turns[0]["visible_question_count"]
    quiz_questions = next(
        turn["visible_question_count"] for turn in turns if turn["phase"] == "quiz_start"
    )
    source_grounded_phases = {
        turn["phase"]: source_id in turn["sources_used"]
        for turn in turns
        if turn["phase"]
        in {
            "lesson_production",
            "lesson_production_retry_1",
            "lesson_production_retry_2",
            "quiz_production",
            "quiz_revision",
            "quiz_revision_retry_1",
            "quiz_revision_retry_2",
            "slides_production",
            "slides_production_retry_1",
            "slides_production_retry_2",
            "handoff",
        }
    }
    preview_text = compact(preview.get("preview_text", ""))
    checks = {
        "paper_downloaded_from_requested_url": paper["url"] == PAPER_URL
        and paper["bytes"] > 100_000
        and paper["pages"] == 11,
        "paper_uploaded_and_extracted": source.get("status") == "ready"
        and source.get("character_count", 0) > 20_000
        and "attention is all you need" in preview_text
        and "transformer" in preview_text,
        "multi_turn": len(turns) >= 9,
        "lesson_start_asks_at_most_three": 0 <= lesson_questions <= 3,
        "quiz_start_asks_at_most_three": quiz_questions <= 3,
        "major_turns_use_uploaded_source": bool(source_grounded_phases)
        and all(source_grounded_phases.values()),
        "lesson_word_file_is_source_grounded": lesson_artifact_ready(
            lesson_response, source_id
        ),
        "slides_file_is_source_grounded": slides_artifact_ready(
            slides_response, source_id
        ),
        "revised_quiz_word_file_is_source_grounded": revised_quiz_ready(
            revision_response, source_id
        ),
        "office_files_are_valid": bool(office_checks) and all(office_checks.values()),
        "paper_architecture_covered": all(
            term in lower
            for term in ("encoder", "decoder", "quer", "key", "value", "mask", "position")
        ),
        "no_later_model_claims": not has_later_model_claim(
            f"{visible}\n{deliverables}"
        ),
        "no_internal_terms": not any(term in lower for term in INTERNAL_TERMS),
        "no_raw_contract_fences": not re.search(
            r"```(?:markdown|json|artifact_spec|state_file)\b", lower
        ),
        "no_incomplete_payload_notice": "withheld an incomplete" not in lower,
    }
    result = {
        "test": "Complete paper-grounded Transformer architecture Co-design workflow",
        "project_id": project_id,
        "source_id": source_id,
        "paper": paper,
        "source": source,
        "source_preview": {
            "preview_truncated": preview.get("preview_truncated"),
            "vector_metadata": preview.get("vector_metadata"),
            "vector_chunks": len(preview.get("vector_chunks", [])),
        },
        "mode": "Co-design",
        "turns": turns,
        "source_grounded_phases": source_grounded_phases,
        "workspace": workspace,
        "office_checks": office_checks,
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
                "source_id": source_id,
                "turns": len(turns),
                "file_artifacts": len(office_checks),
                "source_grounded_phases": source_grounded_phases,
                "checks": checks,
            },
            indent=2,
        )
    )
    return 0 if result["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
