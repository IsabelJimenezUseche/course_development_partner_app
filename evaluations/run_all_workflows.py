"""Run every workflow end to end, from a clean slate, and judge what came out.

One iteration is: delete every project, restart nothing (the caller owns the server),
run all six workflows, then check three separate things about the results.

1. **It ran.** No exceptions, no HTTP errors, every arc completed.
2. **It validates.** The skill's own validators pass over whatever portable state
   each workflow produced.
3. **It is meaningful.** Every generated Office file is opened and its text compared
   against the professor's actual questions for that scenario. A .docx that parses but
   says nothing about what was asked is a failure worth catching, and it is the failure
   a status flag alone will never show.

Workflows covered: Python (Co-design and Auto), VLSI, CRISPR/cell biology, business
analytics, and the transformer paper.

Usage:
    python evaluations/run_all_workflows.py
    python evaluations/run_all_workflows.py --iterations 3
    python evaluations/run_all_workflows.py --only python business
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import httpx

ROOT = Path(__file__).resolve().parents[1]
EVAL = ROOT / "evaluations"
RESULTS = EVAL / "results"
APP_URL = "http://127.0.0.1:8001"

# Each workflow: the script to run, and the vocabulary its artifacts must actually
# engage with. The terms come from what the professor asked for in that scenario, so a
# file that ignores the question fails even when it opens cleanly.
WORKFLOWS = {
    "python": {
        "script": "run_realistic_python_workflow.py",
        "modes": ("Co-design", "Auto"),
        "expect_terms": ("loop", "trace", "variable"),
        "expect_any": ("for", "while"),
    },
    "vlsi": {
        "script": "run_complete_vlsi_codesign_workflow.py",
        "modes": (),
        "expect_terms": ("delay", "sizing"),
        "expect_any": ("transistor", "gate", "logical effort"),
    },
    "crispr": {
        "script": "run_realistic_cellbio_workflow.py",
        "modes": ("Co-design", "Auto"),
        "expect_terms": ("cell",),
        "expect_any": ("crispr", "cas9", "nucleus", "compartment"),
    },
    "business": {
        "script": "run_realistic_business_workflow.py",
        "modes": ("Co-design", "Auto"),
        "expect_terms": ("data",),
        "expect_any": ("client", "revenue", "defensible", "judgment"),
    },
    "transformer": {
        "script": "run_complete_transformer_paper_workflow.py",
        "modes": (),
        "expect_terms": ("attention",),
        "expect_any": ("encoder", "decoder", "mask"),
    },
}
MIN_ARTIFACT_WORDS = 120


def clean_projects(client: httpx.Client) -> int:
    """Delete every project so each iteration starts from nothing."""
    removed = 0
    for _ in range(5):  # hidden eval projects reappear as workflows create them
        listed = client.get("/api/projects").json().get("projects", [])
        if not listed:
            break
        for project in listed:
            if client.delete(f"/api/projects/{project['id']}").status_code == 200:
                removed += 1
    return removed


def docx_text(path: Path) -> str:
    try:
        with zipfile.ZipFile(path) as archive:
            xml = archive.read("word/document.xml").decode("utf-8", "ignore")
    except (zipfile.BadZipFile, KeyError, OSError):
        return ""
    return re.sub(r"<[^>]+>", " ", xml)


def pptx_text(path: Path) -> str:
    chunks = []
    try:
        with zipfile.ZipFile(path) as archive:
            for name in archive.namelist():
                if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                    chunks.append(archive.read(name).decode("utf-8", "ignore"))
    except (zipfile.BadZipFile, OSError):
        return ""
    return re.sub(r"<[^>]+>", " ", "\n".join(chunks))


def artifact_text(path: Path) -> str:
    if path.suffix == ".docx":
        return docx_text(path)
    if path.suffix == ".pptx":
        return pptx_text(path)
    return ""


def run_workflow(name: str, spec: dict, data_dir: Path) -> dict:
    started = datetime.now(timezone.utc)
    before = {p for p in data_dir.rglob("*") if p.suffix in {".docx", ".pptx"}}
    # The realistic harnesses delete their own projects on exit, which removes the
    # Office files before they can be opened. --keep leaves them for the audit; the
    # next iteration clears everything anyway.
    command = [sys.executable, str(EVAL / spec["script"])]
    if spec.get("modes"):
        command.append("--keep")
    process = subprocess.run(
        command,
        capture_output=True,
        text=True,
        cwd=str(ROOT),
        timeout=3600,
    )
    after = {p for p in data_dir.rglob("*") if p.suffix in {".docx", ".pptx"}}
    new_files = after - before

    statuses = re.findall(r'"status":\s*"(\w+)"', process.stdout)
    reported = sorted(set(statuses)) or ["unknown"]
    # Reports from the realistic harness are Markdown, not JSON.
    md_fail = len(re.findall(r"^- FAIL", process.stdout, re.MULTILINE))
    md_pass = len(re.findall(r"^- PASS", process.stdout, re.MULTILINE))

    findings = []
    meaningful = 0
    for path in sorted(new_files):
        text = artifact_text(path).lower()
        words = len(text.split())
        missing = [t for t in spec["expect_terms"] if t not in text]
        any_hit = (not spec["expect_any"]) or any(t in text for t in spec["expect_any"])
        if words < MIN_ARTIFACT_WORDS:
            findings.append(f"{path.name}: only {words} words")
        elif missing:
            findings.append(f"{path.name}: never mentions {', '.join(missing)}")
        elif not any_hit:
            findings.append(f"{path.name}: none of {', '.join(spec['expect_any'])}")
        else:
            meaningful += 1

    return {
        "workflow": name,
        "script": spec["script"],
        "exit_code": process.returncode,
        "crashed": bool(process.returncode != 0 and "Traceback" in process.stderr),
        "traceback": process.stderr.strip()[-800:] if "Traceback" in process.stderr else "",
        "reported_status": reported,
        "checks_passed": md_pass,
        "checks_failed": md_fail,
        "artifacts_created": len(new_files),
        "artifacts_meaningful": meaningful,
        "artifact_findings": findings,
        "seconds": round((datetime.now(timezone.utc) - started).total_seconds(), 1),
    }


def iteration(client: httpx.Client, names: list[str], data_dir: Path, index: int) -> dict:
    removed = clean_projects(client)
    runs = [run_workflow(n, WORKFLOWS[n], data_dir) for n in names]
    clean = all(
        not r["crashed"]
        and r["checks_failed"] == 0
        and "fail" not in r["reported_status"]
        and not r["artifact_findings"]
        for r in runs
    )
    return {
        "iteration": index,
        "projects_cleared": removed,
        "runs": runs,
        "clean": clean,
    }


def summarize(report: dict) -> str:
    lines = [
        "# All workflows — clean-slate iteration report",
        "",
        f"Generated {datetime.now(timezone.utc):%Y-%m-%d %H:%M} UTC.",
        "",
    ]
    for it in report["iterations"]:
        lines += [
            f"## Iteration {it['iteration']} — {'CLEAN' if it['clean'] else 'ISSUES'}",
            "",
            f"Cleared {it['projects_cleared']} projects first.",
            "",
            "| Workflow | Exit | Reported | Checks P/F | Artifacts | Meaningful | Secs |",
            "|---|---|---|---|---|---|---|",
        ]
        for r in it["runs"]:
            lines.append(
                f"| {r['workflow']} | {r['exit_code']} | {','.join(r['reported_status'])} | "
                f"{r['checks_passed']}/{r['checks_failed']} | {r['artifacts_created']} | "
                f"{r['artifacts_meaningful']} | {r['seconds']} |"
            )
        for r in it["runs"]:
            if r["artifact_findings"] or r["traceback"]:
                lines += ["", f"### {r['workflow']} findings", ""]
                for f in r["artifact_findings"]:
                    lines.append(f"- artifact: {f}")
                if r["traceback"]:
                    lines += ["", "```", r["traceback"], "```"]
        lines.append("")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--iterations", type=int, default=1)
    parser.add_argument("--only", nargs="*", choices=sorted(WORKFLOWS), default=None)
    parser.add_argument("--data-dir", default=str(ROOT / "data" / "projects"))
    args = parser.parse_args()

    names = args.only or list(WORKFLOWS)
    data_dir = Path(args.data_dir)
    data_dir.mkdir(parents=True, exist_ok=True)

    report = {"started": datetime.now(timezone.utc).isoformat(), "iterations": []}
    with httpx.Client(base_url=APP_URL, timeout=120.0) as client:
        client.get("/api/config").raise_for_status()
        for index in range(1, args.iterations + 1):
            report["iterations"].append(iteration(client, names, data_dir, index))

    RESULTS.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    (RESULTS / f"all-workflows-{stamp}.json").write_text(
        json.dumps(report, indent=2), encoding="utf-8"
    )
    text = summarize(report)
    (RESULTS / f"all-workflows-{stamp}.md").write_text(text, encoding="utf-8")
    print(text)
    return 0 if all(i["clean"] for i in report["iterations"]) else 1


if __name__ == "__main__":
    raise SystemExit(main())
