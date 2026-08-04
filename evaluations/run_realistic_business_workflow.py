"""Coding for a business school — realistic professor arcs.

The audience is the whole point. These are MBA and undergraduate business students in a
required analytics course. They are not going to become engineers, they are quietly
anxious about programming, and they will abandon the material the moment it feels like a
computer science class. What they actually need is to read a colleague's analysis script
well enough to challenge it, and to automate something they currently do by hand in a
spreadsheet.

So the design target is judgment, not syntax: given a short pandas script that produces a
revenue figure, decide whether you would present that number to a client. That is a
business skill expressed in code, and it is assessable.

The arc carries two pressures a business school actually applies. Around turn five the
professor raises the grading system, which the skill treats as the course owner's
decision rather than something to infer. Around turn six they ask about permitted AI
use, which is a live policy question in every analytics classroom right now.

Usage:
    python evaluations/run_realistic_business_workflow.py
    python evaluations/run_realistic_business_workflow.py --mode Auto --keep
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from realistic_arc import Probe, Scenario, main  # noqa: E402

PROJECT = {
    "name": "Coding for business students — realistic professor arc",
    "course_name": "BUS 315 Business Analytics",
    "level": "Undergraduate",
    "class_time": "90 minutes",
    "outcome": (
        "Students read a short data-analysis script, identify an assumption that would "
        "change the reported figure, and justify whether they would present that figure "
        "to a client."
    ),
    "notes": (
        "Required analytics course. Mixed MBA and senior undergraduate business students. "
        "Spreadsheet-fluent, programming-anxious, mostly non-technical majors. They will "
        "disengage if it reads as a computer science course. No identifiable student "
        "information."
    ),
}

CODESIGN_ARC = (
    (
        "vague_opening",
        "I teach the required analytics course in our business school and the coding week is "
        "where I lose them. They are spreadsheet people. Half of them decide in the first "
        "twenty minutes that this is not for them. I need something better for a week Monday.",
    ),
    (
        "terse_answer",
        "35 students, mix of MBAs and senior undergrads, 90 minutes, laptops. Almost none have "
        "programmed. They are sharp about business judgment and nervous about syntax.",
    ),
    (
        "lesson_request",
        "Build the session. I do not want them writing code from scratch — I want them reading "
        "a short script that computes a revenue number and deciding whether they would put "
        "that number in front of a client. Timings, what I say, what they do.",
    ),
    (
        "constraint_change",
        "One thing I forgot: our lab machines cannot install anything, so it has to run in a "
        "browser notebook with pandas already available. Adjust anything that assumed local "
        "installs.",
    ),
    (
        "grading_question",
        "How should this be graded? It is 15 percent of the course and I let students resubmit "
        "once. Tell me what that means for the design.",
    ),
    (
        "ai_policy",
        "Students will absolutely paste this into an AI assistant. I would rather design for "
        "that than pretend it will not happen. What should my policy be and how does the "
        "assessment change?",
    ),
    (
        "quiz_request",
        "Give me a short Word handout — three items where they inspect a snippet and judge "
        "whether the reported figure is defensible, with an answer key I can grade from.",
    ),
    (
        "artifact_pushback",
        "Item 2 is really testing whether they know pandas syntax, not whether they can judge "
        "the number. Rewrite it so a spreadsheet person can answer it correctly.",
    ),
    (
        "delegate",
        "My TA is a finance PhD student who has never taught coding. One page for him: the two "
        "places students will freeze, and what to say instead of fixing it for them.",
    ),
    (
        "handoff",
        "What is ready for Monday and what still needs a decision from me?",
    ),
)

# Delegation: the constraints Co-design surfaces over ten turns, stated once.
AUTO_ARC = (
    (
        "delegation_brief",
        "Take this one over completely — do not ask me questions, just build it and state your "
        "assumptions. BUS 315 Business Analytics, 35 students, mixed MBA and senior "
        "undergraduate business majors, 90 minutes with laptops. Almost none have programmed; "
        "they are spreadsheet-fluent and anxious about syntax, and they disengage fast if it "
        "feels like a computer science course. Design it so they read a short pandas script "
        "that computes a revenue figure and decide whether they would present that number to "
        "a client — reading and judging, not writing from scratch. Lab machines cannot install "
        "anything, so assume a browser notebook with pandas available. The work is 15 percent "
        "of the course grade with one permitted resubmission, and I want the design to assume "
        "students will use AI assistants rather than pretend they will not. Deliver three "
        "things: the session with timings and what I say versus what they do; a Word handout "
        "with three items where they inspect a snippet and judge whether the reported figure "
        "is defensible, answerable by a spreadsheet person and not a syntax quiz, with a "
        "gradeable answer key; and a one-page guide for a finance PhD TA who has never taught "
        "coding, covering the two places students freeze and what to say instead of fixing it.",
    ),
    (
        "review_pass",
        "Reviewing the handout: any item that really tests pandas syntax rather than business "
        "judgment needs rewriting so a spreadsheet person can answer it. Make the key explain "
        "why the figure is or is not defensible rather than just marking right and wrong. "
        "Regenerate the Word file.",
    ),
    (
        "handoff",
        "What is ready for Monday and what still needs a decision from me?",
    ),
)

SCENARIO = Scenario(
    name="Coding for business students — realistic professor arc",
    project=PROJECT,
    codesign_arc=CODESIGN_ARC,
    auto_arc=AUTO_ARC,
    probes=(
        Probe("framed_around_business_judgment", r"client|stakeholder|defensib|judgment"),
        Probe("kept_reading_over_writing", r"read|inspect|interpret|trace"),
        Probe("honored_no_install_constraint", r"notebook|colab|browser|cloud", from_turn=3),
        # The skill treats the grading system as the owner's decision, not an inference.
        Probe("engaged_resubmission_policy", r"resubmit|resubmission|revision", from_turn=4),
        Probe("addressed_ai_use", r"\bAI\b|assistant|generative", from_turn=5),
        Probe("produced_ta_guidance", r"\bTA\b|teaching assistant"),
        # Business students, not CS students: heavy CS framing is a design smell.
        Probe(
            "avoided_cs_course_framing",
            r"big[- ]O|time complexity|recursion|data structure",
            expect_absent=True,
        ),
    ),
    # It may recommend a grading approach. It may not assert the owner's policy for them.
    forbidden_claims=(
        r"i have (?:set|decided) (?:your|the) (?:grade|grading|weight)",
        r"this (?:is|will be) worth \d+ percent of (?:their|the) final grade\b(?!.*confirm)",
    ),
    min_office_files=2,
    required_formats=("docx",),
)

if __name__ == "__main__":
    raise SystemExit(main(SCENARIO, "realistic-business"))
