from __future__ import annotations

import json
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

import httpx


APP_URL = "http://127.0.0.1:8001"
EVAL_DIR = Path(__file__).resolve().parent
FIXTURE_DIR = EVAL_DIR / "fixtures"
RESULT_DIR = EVAL_DIR / "results"


@dataclass(frozen=True)
class Scenario:
    name: str
    fixture: str
    skill_profile: str
    prompt: str
    required_terms: tuple[str, ...]


SCENARIOS = (
    Scenario(
        name="stem_active_learning_worksheet",
        fixture="stem-reaction-engineering.txt",
        skill_profile="artifact",
        prompt=(
            "Rapid mode. Create a complete 50-minute active-learning worksheet and a concise "
            "instructor guide for undergraduate chemical engineering. Learning outcome: Given "
            "experimental rate observations, students will distinguish a rate law from a proposed "
            "reaction mechanism and justify what additional evidence is needed. Include timing, "
            "ordered student tasks, instructor checkpoints, scaffold fading, a debrief, and an "
            "instructor-review list. Cite the project source ID for source-grounded claims."
        ),
        required_terms=("learning outcome", "student", "instructor", "debrief", "time"),
    ),
    Scenario(
        name="psychology_conceptual_change_activity",
        fixture="psychology-conditioning.txt",
        skill_profile="artifact",
        prompt=(
            "Rapid mode. Create a 45-minute conceptual-change activity for an undergraduate "
            "psychology class. Students should classify UCS, UCR, CS, and CR in unfamiliar cases "
            "and distinguish acquisition, extinction, spontaneous recovery, generalization, and "
            "discrimination. Include an elicitation task, contrasting cases, feedback, transfer, "
            "a debrief, and instructor notes. Cite the project source ID for factual claims."
        ),
        required_terms=("learning outcome", "contrast", "feedback", "transfer", "debrief"),
    ),
    Scenario(
        name="writing_assessment_rubric",
        fixture="writing-position-argument.txt",
        skill_profile="assessment",
        prompt=(
            "Rapid mode. Draft a provisional analytic rubric architecture for a college position "
            "argument. Use a balanced, objective-first orientation but label it provisional pending "
            "instructor approval. Include criteria, observable evidence, performance-level logic, "
            "suggested weights, likely construct-irrelevant barriers, scorer guidance, and a "
            "calibration plan using de-identified responses. Cite the project source ID."
        ),
        required_terms=("criteria", "evidence", "weight", "provisional", "calibration"),
    ),
)


def contains_all(text: str, terms: tuple[str, ...]) -> tuple[bool, list[str]]:
    lowered = text.lower()
    missing = [term for term in terms if term.lower() not in lowered]
    return not missing, missing


def normalize_source_identifier(text: str) -> str:
    """Normalize typographic dash substitutions without weakening exact ID matching."""
    return text.translate(
        str.maketrans(
            {
                "‐": "-",
                "‑": "-",
                "‒": "-",
                "–": "-",
                "—": "-",
                "−": "-",
            }
        )
    )


def run() -> int:
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    report_rows = []
    saved_results = []

    with httpx.Client(base_url=APP_URL, timeout=240.0) as client:
        health = client.get("/health")
        health.raise_for_status()
        health_data = health.json()
        if not all(
            health_data.get(key)
            for key in ("model_configured", "api_key_configured", "skill_loaded")
        ):
            raise RuntimeError(f"App is not ready for live evaluation: {health_data}")

        for scenario in SCENARIOS:
            project_id = f"eval-{scenario.name[:24]}-{uuid4().hex[:8]}"
            fixture_path = FIXTURE_DIR / scenario.fixture
            with fixture_path.open("rb") as source_file:
                upload = client.post(
                    f"/api/projects/{project_id}/sources",
                    data={"data_classification_ack": "true"},
                    files={"files": (fixture_path.name, source_file, "text/plain")},
                )
            upload.raise_for_status()
            upload_data = upload.json()
            if upload_data["errors"] or not upload_data["sources"]:
                raise RuntimeError(f"Source upload failed for {scenario.name}: {upload_data}")
            source_id = upload_data["sources"][0]["source_id"]

            response = client.post(
                "/api/chat",
                json={
                    "project_id": project_id,
                    "skill_profile": scenario.skill_profile,
                    "messages": [{"role": "user", "content": scenario.prompt}],
                },
            )
            response.raise_for_status()
            result = response.json()
            artifact = result["content"]
            terms_pass, missing_terms = contains_all(artifact, scenario.required_terms)
            cited_source_ids = normalize_source_identifier(artifact)
            source_pass = (
                source_id in cited_source_ids
                and source_id in result.get("sources_used", [])
            )
            skill_pass = (
                result.get("skill_runtime", {}).get("profile") == scenario.skill_profile
                and "SKILL.md" in result.get("skill_runtime", {}).get("loaded_files", [])
            )
            reasoning_pass = "<think>" not in artifact.lower()
            unsupported_claim_pass = not any(
                phrase in artifact.lower()
                for phrase in (
                    "fully validated",
                    "ada compliant",
                    "wcag compliant",
                    "guaranteed to improve learning",
                )
            )
            checks = {
                "required_structure": terms_pass,
                "source_grounding": source_pass,
                "actual_skill_route": skill_pass,
                "raw_reasoning_filtered": reasoning_pass,
                "no_unsupported_quality_claim": unsupported_claim_pass,
            }
            passed = sum(checks.values())
            report_rows.append(
                {
                    "scenario": scenario.name,
                    "score": f"{passed}/{len(checks)}",
                    "missing_terms": missing_terms,
                    "checks": checks,
                    "source_id": source_id,
                    "skill_profile": result["skill_runtime"]["profile"],
                    "skill_fingerprint": result["skill_runtime"]["fingerprint"],
                }
            )
            saved_results.append(
                {
                    "scenario": scenario.name,
                    "project_id": project_id,
                    "source": upload_data["sources"][0],
                    "prompt": scenario.prompt,
                    "result": result,
                    "checks": checks,
                    "missing_terms": missing_terms,
                }
            )

    json_path = RESULT_DIR / f"artifact-eval-{timestamp}.json"
    json_path.write_text(json.dumps(saved_results, indent=2), encoding="utf-8")

    markdown_lines = [
        "# Course Development Partner live artifact evaluation",
        "",
        f"- Run: {timestamp}",
        f"- App: {APP_URL}",
        f"- Cases: {len(SCENARIOS)}",
        "",
        "| Scenario | Score | Skill profile | Source grounding | Missing required terms |",
        "|---|---:|---|---|---|",
    ]
    for row in report_rows:
        markdown_lines.append(
            f"| {row['scenario']} | {row['score']} | {row['skill_profile']} | "
            f"{'pass' if row['checks']['source_grounding'] else 'fail'} | "
            f"{', '.join(row['missing_terms']) or 'none'} |"
        )
    markdown_lines.extend(
        [
            "",
            "## Deterministic checks",
            "",
            "Each case checks required artifact vocabulary, exact source-ID grounding, actual "
            "SKILL.md routing, removal of raw reasoning tags, and absence of unsupported validation "
            "or compliance claims. These checks do not establish pedagogical or disciplinary quality; "
            "the generated artifacts still require qualified instructor review.",
            "",
            f"Raw outputs: `{json_path.name}`",
        ]
    )
    report_path = RESULT_DIR / f"artifact-eval-{timestamp}.md"
    report_path.write_text("\n".join(markdown_lines) + "\n", encoding="utf-8")
    print(report_path)
    print(json_path)
    for row in report_rows:
        print(f"{row['scenario']}: {row['score']}")
    return 0 if all(row["score"] == "5/5" for row in report_rows) else 1


if __name__ == "__main__":
    raise SystemExit(run())
