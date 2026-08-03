#!/usr/bin/env python3
"""Create and validate a persistent Attention Is All You Need demo project."""

from __future__ import annotations

import argparse
import json
import os
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
from typing import Optional
from zipfile import ZipFile, is_zipfile

import httpx
from pypdf import PdfReader


APP_URL = os.getenv("APP_URL", "http://127.0.0.1:8001").rstrip("/")
PAPER_URL = "https://arxiv.org/pdf/1706.03762"
RESULT_DIR = Path(__file__).resolve().parent / "results"


def download_paper(client: httpx.Client, destination: Path) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    response = client.get(PAPER_URL, follow_redirects=True, timeout=120.0)
    response.raise_for_status()
    destination.write_bytes(response.content)
    return destination


def artifact_specs(source_id: str) -> list[dict]:
    shared_sources = [source_id]
    return [
        {
            "kind": "slides",
            "title": "Attention Is All You Need: Transformer Foundations",
            "subtitle": "Graduate seminar mini-lesson grounded in Vaswani et al. (2017)",
            "sections": [
                {
                    "heading": "Learning objectives",
                    "body": "By the end of the session, learners can explain the Transformer design and evaluate evidence for its original machine-translation results.",
                    "bullets": [
                        "Trace information through encoder and decoder stacks",
                        "Explain scaled dot-product and multi-head attention",
                        "Compare attention-only processing with recurrent processing",
                    ],
                    "prompts": [], "checklist": [], "response_lines": 2, "table": None,
                },
                {
                    "heading": "The architectural shift",
                    "body": "The paper replaces recurrence and convolution with attention mechanisms, enabling more parallel computation during training.",
                    "bullets": [
                        "Encoder and decoder use repeated attention and feed-forward blocks",
                        "Residual connections and layer normalization wrap each sublayer",
                        "Positional encodings supply sequence-order information",
                    ],
                    "prompts": [], "checklist": [], "response_lines": 2, "table": None,
                },
                {
                    "heading": "Scaled dot-product attention",
                    "body": "Attention(Q, K, V) = softmax(QK^T / sqrt(d_k))V. Scaling moderates large dot products before the softmax.",
                    "bullets": [
                        "Queries determine what information is sought",
                        "Keys support relevance comparisons",
                        "Values provide the information combined in the output",
                        "Multiple heads learn distinct representation subspaces",
                    ],
                    "prompts": [], "checklist": [], "response_lines": 2, "table": None,
                },
                {
                    "heading": "Evidence reported in the paper",
                    "body": "The authors evaluate the architecture primarily on WMT 2014 translation and also test English constituency parsing.",
                    "bullets": [], "prompts": [], "checklist": [], "response_lines": 2,
                    "table": {
                        "headers": ["Task", "Reported result", "Interpretation boundary"],
                        "rows": [
                            ["English-to-German", "28.4 BLEU", "Original WMT 2014 setup"],
                            ["English-to-French", "41.8 BLEU", "Single-model result"],
                            ["Training", "3.5 days on eight GPUs", "Hardware- and implementation-dependent"],
                        ],
                    },
                },
                {
                    "heading": "Discussion and transfer",
                    "body": "Treat the paper as evidence for a specific architecture and evaluation context, not as proof that attention is universally sufficient.",
                    "bullets": [
                        "Which design choice most improves parallelism?",
                        "Which claim is directly supported by the reported experiments?",
                        "What additional evidence would be needed for a new domain?",
                    ],
                    "prompts": [], "checklist": ["Separate reported findings from later interpretations"], "response_lines": 2, "table": None,
                },
            ],
            "source_ids": shared_sources,
        },
        {
            "kind": "document",
            "title": "Instructor Guide: Teaching the Transformer Paper",
            "subtitle": "A source-grounded 90-minute graduate seminar plan",
            "sections": [
                {
                    "heading": "Purpose and preparation",
                    "body": "Use the paper to teach architectural reasoning: learners connect a stated limitation of recurrent sequence models to a specific Transformer design choice.",
                    "bullets": ["Assign Sections 1-3 before class", "Ask learners to annotate Figure 1", "Require claims to point to a paper section, equation, table, or figure"],
                    "prompts": [], "checklist": ["Paper is available to every learner", "Source ID appears in distributed materials"], "response_lines": 2, "table": None,
                },
                {
                    "heading": "90-minute sequence",
                    "body": "The sequence moves from individual interpretation to collaborative model tracing and evidence evaluation.",
                    "bullets": [], "prompts": [], "checklist": [], "response_lines": 2,
                    "table": {
                        "headers": ["Minutes", "Activity", "Evidence of learning"],
                        "rows": [
                            ["0-10", "Retrieval prompt", "Individual architecture sketch"],
                            ["10-30", "Guided Figure 1 trace", "Annotated encoder-decoder path"],
                            ["30-50", "Attention equation worked example", "Correct Q/K/V interpretation"],
                            ["50-70", "Evidence audit", "Claim-evidence-boundary table"],
                            ["70-85", "Transfer critique", "Defensible recommendation"],
                            ["85-90", "Exit ticket", "One supported claim and one limitation"],
                        ],
                    },
                },
                {
                    "heading": "Facilitation notes",
                    "body": "Keep attention focused on what the 2017 paper actually reports. Later applications of Transformers require separate sources.",
                    "bullets": ["Distinguish architecture components from training results", "Ask for evidence before accepting broad claims", "Treat BLEU scores as task- and setup-specific"],
                    "prompts": ["What does the paper establish?", "What remains an inference or an open validation need?"], "checklist": [], "response_lines": 3, "table": None,
                },
            ],
            "source_ids": shared_sources,
        },
        {
            "kind": "worksheet",
            "title": "Transformer Paper Evidence Worksheet",
            "subtitle": "Trace the architecture, evaluate claims, and identify limitations",
            "sections": [
                {
                    "heading": "Architecture trace",
                    "body": "Use Figure 1 and Sections 3.1-3.5 of the paper.",
                    "bullets": [],
                    "prompts": [
                        "Trace one token representation through the encoder stack.",
                        "Explain the roles of queries, keys, and values in your own words.",
                        "Why is positional information added to the input representations?",
                    ],
                    "checklist": ["Names the relevant sublayers", "Connects each explanation to the paper"], "response_lines": 4, "table": None,
                },
                {
                    "heading": "Claim-evidence-boundary audit",
                    "body": "Record only evidence present in the supplied paper.",
                    "bullets": [], "prompts": [], "checklist": [], "response_lines": 3,
                    "table": {
                        "headers": ["Claim", "Paper evidence", "Boundary or caveat"],
                        "rows": [
                            ["Attention improves parallelization", "", ""],
                            ["The model performs strongly on translation", "", ""],
                            ["The architecture transfers beyond translation", "", ""],
                        ],
                    },
                },
                {
                    "heading": "Exit ticket",
                    "body": "Write a recommendation that distinguishes reported evidence from your interpretation.",
                    "bullets": [],
                    "prompts": ["Which Transformer design choice is best supported by the paper's argument and results?", "What additional test would you require before applying the claim in a different domain?"],
                    "checklist": ["Cites a paper location", "States one limitation", "Avoids an unsupported universal claim"], "response_lines": 5, "table": None,
                },
            ],
            "source_ids": shared_sources,
        },
    ]


def validate_office_file(kind: str, content: bytes) -> dict:
    if len(content) <= 1_000 or not is_zipfile(BytesIO(content)):
        raise RuntimeError(f"Generated {kind} artifact is not a valid Office file")
    with ZipFile(BytesIO(content)) as archive:
        names = archive.namelist()
    required_prefix = "ppt/" if kind == "slides" else "word/"
    if not any(name.startswith(required_prefix) for name in names):
        raise RuntimeError(f"Generated {kind} artifact is missing {required_prefix} content")
    return {"bytes": len(content), "office_prefix": required_prefix, "valid": True}


def run(pdf_path: Path, existing_project_id: Optional[str] = None) -> dict:
    with httpx.Client(base_url=APP_URL, timeout=240.0) as app_client:
        health = app_client.get("/health")
        health.raise_for_status()
        health_data = health.json()
        if not all(health_data.get(key) for key in ("model_configured", "api_key_configured", "skill_loaded")):
            raise RuntimeError(f"App is not ready: {health_data}")

        if existing_project_id:
            workspace = app_client.get(f"/api/projects/{existing_project_id}")
            workspace.raise_for_status()
            workspace_data = workspace.json()
            project = workspace_data["project"]
            existing_artifacts = workspace_data.get("artifacts", [])
            sources_response = app_client.get(f"/api/projects/{project['id']}/sources")
            sources_response.raise_for_status()
            sources = sources_response.json()["sources"]
            if not sources:
                raise RuntimeError("The existing project has no uploaded source")
            source = sources[0]
        else:
            existing_artifacts = []
            project_response = app_client.post(
                "/api/projects",
                json={
                    "name": "Attention Is All You Need - Source-Grounded Demo",
                    "course_name": "Graduate Seminar: Transformer Foundations",
                    "level": "Graduate",
                    "class_time": "90 minutes",
                    "outcome": "Explain the Transformer architecture and evaluate claims using evidence and limitations from Vaswani et al. (2017).",
                    "mode": "Auto",
                    "notes": "Use only the uploaded arXiv paper for factual claims. Clearly separate reported results from later interpretations.",
                },
            )
            project_response.raise_for_status()
            project = project_response.json()["project"]

            with pdf_path.open("rb") as paper:
                upload = app_client.post(
                    f"/api/projects/{project['id']}/sources",
                    data={"data_classification_ack": "true"},
                    files={"files": (pdf_path.name, paper, "application/pdf")},
                )
            upload.raise_for_status()
            upload_data = upload.json()
            if upload_data["errors"] or len(upload_data["sources"]) != 1:
                raise RuntimeError(f"Source upload failed: {upload_data}")
            source = upload_data["sources"][0]

        prompt = (
            "Using only the uploaded paper, return exactly five bullets of no more than 25 words each: "
            "motivation, encoder-decoder design, attention equation, reported results, and one limitation. "
            "Cite the supplied source ID in every bullet. No question and no artifact specification."
        )
        chat_error = None
        try:
            chat = app_client.post(
                "/api/chat",
                json={
                    "project_id": project["id"],
                    "skill_profile": "establish",
                    "display_content": "Create a five-bullet source-grounded brief from the uploaded Transformer paper.",
                    "messages": [{"role": "user", "content": prompt}],
                },
            )
            chat.raise_for_status()
            chat_data = chat.json()
            chat_text = chat_data.get("content", "")
            chat_checks = {
                "request_completed": True,
                "source_used": source["source_id"] in chat_data.get("sources_used", []),
                "skill_loaded": "SKILL.md" in chat_data.get("skill_runtime", {}).get("loaded_files", []),
                "raw_reasoning_filtered": "<think>" not in chat_text.lower(),
                "transformer_content": all(term in chat_text.lower() for term in ("attention", "encoder", "decoder")),
            }
        except httpx.HTTPStatusError as exc:
            chat_error = f"{exc.response.status_code}: {exc.response.text[:300]}"
            chat_checks = {
                "request_completed": False,
                "source_used": False,
                "skill_loaded": False,
                "raw_reasoning_filtered": True,
                "transformer_content": False,
            }

        artifacts = []
        artifacts_to_validate = existing_artifacts
        if not artifacts_to_validate:
            artifacts_to_validate = []
            for spec in artifact_specs(source["source_id"]):
                generated = app_client.post(
                    f"/api/projects/{project['id']}/artifact-tools/generate", json=spec
                )
                generated.raise_for_status()
                artifacts_to_validate.append(generated.json()["artifact"])

        for artifact in artifacts_to_validate:
            if not artifact.get("has_file"):
                continue
            office = app_client.get(
                f"/api/projects/{project['id']}/artifacts/{artifact['id']}/download",
                params={"format": "office"},
            )
            office.raise_for_status()
            validation = validate_office_file(artifact["kind"], office.content)
            artifacts.append(
                {
                    "id": artifact["id"],
                    "kind": artifact["kind"],
                    "title": artifact["title"],
                    "file_format": artifact["file_format"],
                    "source_ids": artifact.get("tool_trace", {}).get("source_ids", []),
                    "validation": validation,
                }
            )

        package = app_client.get(f"/api/projects/{project['id']}/export", params={"format": "zip"})
        package.raise_for_status()
        with ZipFile(BytesIO(package.content)) as archive:
            package_names = archive.namelist()
        package_checks = {
            "conversation": "conversation.md" in package_names,
            "project_metadata": "project.json" in package_names,
            "slides": any(name.endswith(".pptx") for name in package_names),
            "documents": sum(name.endswith(".docx") for name in package_names) == 2,
        }
        if not all(package_checks.values()):
            raise RuntimeError(f"Project export checks failed: {package_checks}")

    report = {
        "created_at": datetime.now(timezone.utc).isoformat(),
        "project": project,
        "project_url": f"{APP_URL}/?project={project['id']}",
        "paper": {
            "url": PAPER_URL,
            "filename": pdf_path.name,
            "pages": len(PdfReader(str(pdf_path)).pages),
            "bytes": pdf_path.stat().st_size,
            "source_id": source["source_id"],
        },
        "chat_checks": chat_checks,
        "chat_error": chat_error,
        "artifacts": artifacts,
        "package_checks": package_checks,
    }
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    report_path = RESULT_DIR / f"attention-paper-demo-{timestamp}.json"
    report_path.write_text(json.dumps(report, indent=2), encoding="utf-8")
    report["report_path"] = str(report_path)
    return report


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", type=Path, default=Path("tmp/pdfs/attention-is-all-you-need.pdf"))
    parser.add_argument("--project-id")
    args = parser.parse_args()
    pdf_path = args.pdf.resolve()
    if not pdf_path.exists():
        with httpx.Client() as download_client:
            download_paper(download_client, pdf_path)
    print(json.dumps(run(pdf_path, args.project_id), indent=2))


if __name__ == "__main__":
    main()
