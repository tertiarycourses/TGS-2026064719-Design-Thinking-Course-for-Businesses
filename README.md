# Design Thinking Course for Businesses

**WSQ Course Code:** TGS-2026064719
**TSC:** Design Thinking Practices · DSN-ACE-3014-1.1
**Duration:** 1 day · 8 course hours (385 min instruction + 95 min assessment)
**Conducted by:** Tertiary Infotech Academy Pte Ltd (UEN 201200696W)
**Version:** v15.0 · 10 August 2026

Courseware for the WSQ *Design Thinking Course for Businesses*, generated from a
single source so the slide deck, Lesson Plan, Learner Guide and activity files
can never drift apart.

Course page: <https://www.tertiarycourses.com.sg/casl-design-thinking-course-for-businesses.html>

---

## Artifacts

| Artifact | File |
|---|---|
| Trainer slide deck (89 slides) | [`courseware/Design Thinking Course for Businesses-v15.0.pptx`](courseware/) |
| Learner slides (PDF) | `courseware/Design Thinking Course for Businesses-v15.0.pdf` |
| Lesson Plan | `courseware/LP-Design Thinking Course for Businesses.docx` / `.pdf` |
| Learner Guide | `courseware/LG-Design Thinking Course for Businesses.docx` / `.pdf` |
| Learner Guide (Markdown mirror) | [`LG-Design Thinking Course for Businesses.md`](LG-Design%20Thinking%20Course%20for%20Businesses.md) |
| Hands-on activities | [`labs/`](labs/) |

> The assessment instruments (WA + PP question papers and answer keys) are
> **confidential** and are distributed via Google Drive / the LMS only. They are
> never committed to this repository.

---

## Learning outcomes

- **LO1** Understand key concepts of design thinking to communicate the design outcomes
- **LO2** Apply design thinking to generate new ideas for the organization
- **LO3** Uncover opportunities of design thinking
- **LO4** Implement plans to embed various stages of design thinking across the organization
- **LO5** Execute the design concept through prototypes
- **LO6** Utilize metrics to measure the outcomes of design ideas and solutions

## Topics

| # | Topic | Skills mapped | Activities |
|---|---|---|---|
| 1 | Key Concepts & Principles of Design Thinking | A6, K1, K2, K3, A1, K6 | 1–2 |
| 2 | Applications of Design Thinking | A2, K5 | 3 |
| 3 | Action Phases of Design Thinking | A4, K4, A1 | 4 |
| 4 | Methodologies and Visual Tools for Design Thinking | A5, K7, A3 | 5–7 |

## Hands-on activities

All activities run in the browser — no installation required.

| # | Activity | Ed-tool |
|---|---|---|
| 1 | Design a Vase — Reframing the Brief | [designthinking](https://alfredang.github.io/designthinking/) |
| 2 | Design Thinking Readiness — Mindset & Culture Self-Assessment | [digitaltransformation](https://alfredang.github.io/digitaltransformation/) |
| 3 | Uncover Design Thinking Opportunities in Your Business | [designthinking](https://alfredang.github.io/designthinking/) |
| 4 | The Wallet Project — A Full Design Thinking Cycle | [designthinking](https://alfredang.github.io/designthinking/) |
| 5 | Empathy Map & Persona — The Gift-Giving Experience | [designthinking](https://alfredang.github.io/designthinking/) |
| 6 | POV, How Might We & Brainstorming | [designthinking](https://alfredang.github.io/designthinking/) |
| 7 | Prototype, Test & Measure the Outcome | [designthinking](https://alfredang.github.io/designthinking/) · [raci](https://alfredang.github.io/raci/) · [scrum](https://alfredang.github.io/scrum/) · [bcm](https://alfredang.github.io/bcm/) |

### Ed-tools

| Tool | Link |
|---|---|
| Design Thinking Toolkit | <https://alfredang.github.io/designthinking/> |
| RACI Matrix | <https://alfredang.github.io/raci/> |
| Scrum Planner | <https://alfredang.github.io/scrum/> |
| Digital Transformation Canvas | <https://alfredang.github.io/digitaltransformation/> |
| Business Continuity Management | <https://alfredang.github.io/bcm/> |

---

## Assessment

| Instrument | Format | Duration |
|---|---|---|
| Practical Performance (PP) | Design-thinking tasks using the ed-tools, open book | 70 min |
| Oral Questioning (OQ) | Assessor asks each learner verbally, open book | 20 min |

The instruments are a revision of the approved papers on file — same instrument
types, same question and task counts (4 tasks / 7 questions), same K and A code
mapping and the same timings — with the content rewritten against this build's
slides and activities. The OQ is delivered **verbally**: the assessor asks each
question and records the learner's response; there is no written paper for the
learner to complete.

A minimum of **75% attendance** is required to be eligible for assessment and
funding, and learners must be assessed **Competent** in both instruments.

---

## Rebuilding the courseware

Everything is generated from `course_data.py` + `data_domain1..4.py` in
`.claude/skills/courseware-build/build/`. Edit the content there — never edit the
generated DOCX/PPTX by hand — then run:

```bash
bash build.sh
```

This regenerates the labs, deck, Lesson Plan and Learner Guide, renders all
PDFs, injects page-numbered Tables of Contents, and reconciles the LP's slide
references against the final deck.

**Requirements:** Python 3 with `python-pptx`, `python-docx` and `pypdf`, plus
LibreOffice (`soffice`) on the PATH for PDF rendering.

### Design rules enforced by the generators

- The slide deck is **highly visual** — tile grids, chevron flows, profile cards,
  canvas mock-ups, stat bands and quadrant matrices. It deliberately carries
  **no step-by-step instruction slides**; those live in the Learner Guide and `labs/`.
- Every slide footer carries this course's title and TGS code.
- Admin block order: Digital Attendance → Trainer profiles → Ground Rules →
  Download Material → Ed-tools → Lesson Plan → Learning Outcomes →
  **Briefing before Assessment** → Assessment → Assessment Flow.
- Closing block: Assessment → Assessment Flow → Digital Attendance (TRAQOM) →
  Recommended Courses → Support → Thank You.
- The Lesson Plan asserts that each day totals exactly 480 minutes.

---

© 2026 Tertiary Infotech Academy Pte Ltd. All rights reserved.
