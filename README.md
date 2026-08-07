# Course Development Partner — local workspace

A local web application that runs the [Course Development Partner](https://github.com/denphi/course_development_partner)
Agent Skill as a tool-using agent, so a course design survives between sessions and its
teaching artifacts get checked before anyone teaches from them.

The skill is the method and it is portable. **This app is optional.** It exists because a
design conversation in a chat window disappears, produces no editable files, and nothing
in it verifies that the quiz you generated actually assesses the outcome you wrote down.

## Two ways to use the method

**Skill only — about two minutes, no setup.** Copy the `course-development-partner`
folder into your agent's skills directory and start a conversation. You get the full
design workflow, the state-file templates, and the validators as scripts you can run by
hand. Most people should start here.

**This app — for persistence and verification.** Adds a project workspace, uploaded
source material the model actually cites, generated Word and PowerPoint files, and
validators the model runs on itself. Needs Python and a Purdue GenAI key.

## Quickstart

```bash
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

cp .env.example .env          # then set PURDUE_GENAI_API_KEY and COURSE_SKILL_DIR
.venv/bin/python server.py    # http://127.0.0.1:8001
```

`COURSE_SKILL_DIR` must point at the installed skill — the directory holding `SKILL.md`,
`references/`, `assets/`, and `scripts/`. If port 8001 is taken the server picks the next
free one and prints it.

Requires Python 3.10+. Everything runs locally except the model call.

## What the app adds

**The model checks its own work.** It has five tools — `read_skill_template`,
`list_state_files`, `read_state_file`, `write_state_file`, `run_validators`. Writing a
state file runs that file's validator and returns the findings, and the model is expected
to fix what comes back before replying. In testing it repaired malformed identifiers and
comma-separated lists unaided, on its own second attempt.

**Six validators** ship with the skill and run as ordinary subprocesses:

```
validate_design_state.py        validate_course_curriculum_map.py
validate_alignment_map.py       validate_artifact_manifest.py
validate_assessment_blueprint.py validate_project.py
```

They share one contract — exit `0` clean, `1` a hard error, `2` gaps still to fill — and
accept `--json` for programmatic callers. A gap is a design step not taken yet; an error
is a malformed file. The distinction is deliberate.

**Deterministic Office output.** Word and PowerPoint files are built locally by
`artifact_tools.py` from a validated specification. The model never manipulates binary
formats.

**Source grounding.** Uploaded PDF, DOCX, PPTX, XLSX, TXT, MD, and CSV files are chunked
and retrieved locally, and responses cite the source IDs they drew on.

## Configuration

Set in `.env` (see `.env.example`). `.env` is gitignored — keep keys out of the repo.

| Variable | Purpose |
|---|---|
| `COURSE_SKILL_DIR` | Path to the installed skill. Required. |
| `PURDUE_GENAI_API_KEY` | Model access. Required. |
| `PURDUE_GENAI_MODEL_ID` | Defaults to `gpt-oss:120b`. |
| `PURDUE_GENAI_MAX_TOKENS` | Keep at 6000 or above; a state file does not fit in less. |
| `APP_WORK_DIR` | Base for relative paths, so project data can live outside the install. |
| `APP_DATA_DIR` | Where projects, sources, and artifacts are stored. |

## Tests and evaluations

```bash
.venv/bin/python -m pytest -q                      # unit tests, app + installed skill
.venv/bin/python evaluations/run_all_workflows.py  # end-to-end, needs the server running
```

The evaluation harness drives five realistic professor workflows — intro Python, VLSI,
cell biology and CRISPR, business analytics, and a source-grounded paper lesson — in both
Co-design and Auto modes. It clears all projects, runs each workflow, validates the
portable state, then opens every generated Office file and checks it engages what the
professor actually asked for. `--iterations 3` is worth using: these runs vary.

## Known limits

- Tuned for `gpt-oss:120b` through Purdue GenAI. Other models will need the prompt
  contracts revisited.
- The Purdue endpoint rate-limits. A throttled evaluation run currently reports as a
  failure rather than a skip, so read red results carefully.
- Validators check structure, not teaching quality. Passing means a file is well formed
  for its declared scope — never that the material is good, accessible, or correct.
- The educator remains the disciplinary and pedagogical authority. Hazard-bearing content
  is always routed to the responsible safety owner.
