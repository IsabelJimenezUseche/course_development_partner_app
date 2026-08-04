"""Intro Python loops — a realistic professor arc in Co-design, a delegation brief in Auto.

Co-design: the professor arrives vague, answers tersely, loses the room to a
double-booking, pushes back on material that is too advanced, finds one ambiguous quiz
question, and then wants something for the TA.

Auto: the same professor, but delegating. One thorough brief with every constraint
already in it, then a review pass and a handoff check. Same deliverables demanded.

Usage:
    python evaluations/run_realistic_python_workflow.py
    python evaluations/run_realistic_python_workflow.py --mode Auto --keep
"""

from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from realistic_arc import Probe, Scenario, main  # noqa: E402

PROJECT = {
    "name": "Intro Python loops — realistic professor arc",
    "course_name": "Introduction to Programming",
    "level": "Undergraduate",
    "class_time": "75 minutes",
    "outcome": (
        "Students trace the value of each variable through a loop and predict its state "
        "at termination."
    ),
    "notes": "First-year students, week 3, computer lab, 24 students.",
}

CODESIGN_ARC = (
    (
        "vague_opening",
        "My first-years fall apart on loops. Week 3 of intro Python, 24 of them, in a lab. "
        "I need something for next Tuesday.",
    ),
    (
        "terse_answer",
        "They know variables, conditionals, lists. 75 minutes. Mostly they can't trace what "
        "the variable is doing on each pass.",
    ),
    (
        "lesson_request",
        "Good. Build me the actual lesson — timings, what I say, what they do.",
    ),
    (
        "constraint_change",
        "Change of plan: the room got double-booked and I only have 50 minutes now. Rework it "
        "to fit, and tell me what you cut.",
    ),
    (
        "pushback",
        "Drop the list comprehensions — that is week 7 material and it will lose them. Keep it "
        "to explicit for and while loops.",
    ),
    (
        "quiz_request",
        "Now a short formative quiz I can hand out on paper — three questions on tracing a "
        "loop, with an answer key. Word file please.",
    ),
    (
        "artifact_pushback",
        "Question 2 is ambiguous — two of the answers are defensible as written. Rewrite it so "
        "there is exactly one correct answer, and update the key.",
    ),
    (
        "delegate",
        "My TA runs the lab section. Give her a one-page guide: what to watch for, the two "
        "mistakes they will make, and how to respond.",
    ),
    (
        "handoff",
        "What is ready to use on Tuesday and what still needs me?",
    ),
)

# Delegation: everything Co-design discovers over nine turns is stated once, up front.
AUTO_ARC = (
    (
        "delegation_brief",
        "I am handing this to you completely — do not ask me questions, just build it. "
        "Week 3 intro Python, 24 first-years in a computer lab, 50 minutes exactly. They "
        "know variables, conditionals, and lists. Their problem is tracing what a variable "
        "holds on each pass of a loop. Stay on explicit for and while loops — list "
        "comprehensions are week 7 and will lose them. I need three things: the lesson with "
        "timings and what I say versus what they do; a three-question formative quiz on "
        "tracing a loop as a Word file with an answer key, each question having exactly one "
        "defensible answer; and a one-page guide for my TA covering what to watch for, the "
        "two mistakes they will make, and how to respond. State any assumption you make.",
    ),
    (
        "review_pass",
        "I have looked at the quiz. Tighten any question where more than one answer could be "
        "defended, and make the answer key explain the reasoning rather than just giving the "
        "value. Regenerate the Word file.",
    ),
    (
        "handoff",
        "What is ready to use on Tuesday and what still needs me?",
    ),
)

SCENARIO = Scenario(
    name="Intro Python loops — realistic professor arc",
    project=PROJECT,
    codesign_arc=CODESIGN_ARC,
    auto_arc=AUTO_ARC,
    probes=(
        # Co-design reveals the 50-minute change at turn 3; Auto states it at turn 0.
        Probe("honored_50_minute_constraint", r"50[\s‑-]*(?:minute|min)", from_turn=3),
        Probe(
            "still_referenced_75_minutes",
            r"75[\s‑-]*(?:minute|min)",
            from_turn=3,
        ),
        Probe("avoided_comprehensions_after_pushback", r"comprehension", from_turn=5,
              expect_absent=True),
        Probe("produced_ta_guidance", r"\bTA\b|teaching assistant", from_turn=0),
    ),
    min_office_files=2,
    required_formats=("docx",),
)

if __name__ == "__main__":
    raise SystemExit(main(SCENARIO, "realistic-python"))
