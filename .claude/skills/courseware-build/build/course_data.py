"""
SINGLE SOURCE OF TRUTH — Design Thinking Course for Businesses (TGS-2026064719).

Every artifact (PPT, LP, LG, LG.md, labs index) is generated from this file plus
data_domain1..4.py, so the deck, lesson plan, learner guide and labs can never
drift apart.

Content is carried over from the approved master deck
"WSQ - Master Trainer Slides - Design Thinking Course for Businesses - v14.pptx"
(see courseware/reference/), restructured to the current Tertiary Infotech WSQ
house standard and extended with hands-on labs built on the browser-based
ed-tools at https://alfredang.github.io/.
"""

# ------------------------------------------------------------------ metadata
TITLE        = "Design Thinking Course for Businesses"
SHORT_TITLE  = "Design Thinking Course for Businesses"   # used in output filenames
COURSE_CODE  = "TGS-2026064719"
TSC_TITLE    = "Design Thinking Practices"
TSC_CODE     = "DSN-ACE-3014-1.1"
VERSION      = "v15.0"
VERSION_DATE = "10 August 2026"
ORG          = "Tertiary Infotech Academy Pte Ltd"
UEN          = "UEN: 201200696W"
TRAINER      = "Dr Alfred Ang"
DAYS         = 1

# ------------------------------------------------------------------ ed-tools (browser-based, no install)
EDTOOLS = [
    ("Design Thinking Toolkit", "https://alfredang.github.io/designthinking/",
     "Empathy Map, Persona, Journey Map, POV, HMW, Brainstorm, Prototype and Test canvases — the primary tool for this course."),
    ("RACI Matrix", "https://alfredang.github.io/raci/",
     "Assign Responsible / Accountable / Consulted / Informed roles so a design idea has a clear owner when it moves to delivery."),
    ("Scrum Planner", "https://alfredang.github.io/scrum/",
     "Turn validated design concepts into a prioritised product backlog and sprint plan."),
    ("Digital Transformation Canvas", "https://alfredang.github.io/digitaltransformation/",
     "Position a design initiative within the organisation's wider transformation roadmap."),
    ("Business Continuity Management", "https://alfredang.github.io/bcm/",
     "Stress-test a design solution for operational risk and continuity before rollout."),
]

# ------------------------------------------------------------------ outcomes
LEARNING_OUTCOMES = [
    "LO1: Understand key concepts of design thinking to communicate the design outcomes",
    "LO2: Apply design thinking to generate new ideas for the organization",
    "LO3: Uncover opportunities of design thinking",
    "LO4: Implement plans to embed various stages of design thinking across the organization",
    "LO5: Execute the design concept through prototypes",
    "LO6: Utilize metrics to measure the outcomes of design ideas and solutions",
]

# ------------------------------------------------------------------ TSC abilities & knowledge (Skills Framework for Design)
TSC_ABILITIES = [
    "A1: Apply design thinking methodologies to define design problems and generate new ideas for the organization",
    "A2: Uncover opportunities for applying design thinking across the organization",
    "A3: Utilize metrics to benchmark and measure outcomes of design ideas and solutions",
    "A4: Implement plans to embed design thinking across the organization",
    "A5: Facilitate the development of design concepts through prototypes and visual tools",
    "A6: Communicate design outcomes and their value to stakeholders",
]
TSC_KNOWLEDGE = [
    "K1: Key concepts and principles of design thinking",
    "K2: Importance and business value of design thinking",
    "K3: Traits and mindset of a design thinker",
    "K4: Action phases of design thinking and the activities within each phase",
    "K5: Use cases and applications of design thinking across industries",
    "K6: Approaches to applying design thinking within an organization",
    "K7: Methodologies, visual tools and metrics used at each action phase",
]

# ------------------------------------------------------------------ topics (= skills domains)
TOPICS = [
    dict(num=1, code="01",
         title="Key Concepts & Principles of Design Thinking",
         subtitle="Concept · Importance · Traits of a design thinker · Applying design thinking for the organisation",
         weighting="A6, K1, K2, K3, A1, K6",
         concepts=[
            ("What design thinking is", "A human-centred problem-solving approach that aims to improve people's experience — as much a mindset as a process."),
            ("Why it matters", "Design-led organisations report higher market share, stronger competitive advantage and more loyal customers."),
            ("Not just for designers", "Anyone can think like a designer to create products and services that improve the customer experience."),
            ("Analytical vs creative", "Design thinking deliberately pairs left-brain analysis with right-brain creativity."),
            ("The design thinker's mindset", "Think users first, ask the right questions, sketch, commit to explore, prototype to test."),
            ("Desirability · Viability · Feasibility", "Organisational design thinking lives at the intersection of user desire, business viability and technical feasibility."),
         ]),
    dict(num=2, code="02",
         title="Applications of Design Thinking",
         subtitle="Use cases across industries · Uncovering opportunities for applying design thinking",
         weighting="A2, K5",
         concepts=[
            ("Airbnb", "Built a culture of experimentation and design-led iteration to move from failing start-up to a billion-dollar business."),
            ("IBM", "Applies design thinking at scale to complex teams, problems and organisations through IBM Design Thinking."),
            ("Bank of America", "Partnered with IDEO to research real savers, producing the 'Keep the Change' account."),
            ("Uber Eats", "Designs its food-delivery service around observed customer, courier and restaurant journeys."),
            ("Healthcare", "Stanford's Hasso Plattner Institute of Design redesigned the emergency-room patient experience."),
            ("Public & social sector", "Clean Team (Ghana in-home toilets), Golden Gate Regional Center and The Good Kitchen redesigned public services."),
         ]),
    dict(num=3, code="03",
         title="Action Phases of Design Thinking",
         subtitle="The Stanford d.school five-stage model · Empathize · Define · Ideate · Prototype · Test",
         weighting="A4, K4, A1",
         concepts=[
            ("Empathize", "Understand the user's experience, situation and emotion by observing, engaging and immersing — without judging."),
            ("Define", "Synthesise the findings into a meaningful, actionable problem statement (user + need + insight)."),
            ("Ideate", "Go wide: generate a large variety and quantity of ideas before narrowing to the strongest few."),
            ("Prototype", "Build to think — a cheap, fast, low-fidelity artefact that makes the idea tangible."),
            ("Test", "Put the prototype in users' hands, gather feedback, and refine or reframe."),
            ("Iterative, not linear", "The five phases loop — insights from Test routinely send you back to Empathize or Define."),
         ]),
    dict(num=4, code="04",
         title="Methodologies and Visual Tools for Design Thinking",
         subtitle="Empathize · Define · Ideate · Prototype · Test methods, tools and metrics",
         weighting="A5, K7, A3",
         concepts=[
            ("Empathize tools", "Empathy Map (Says / Thinks / Does / Feels), Persona Mapping, Journey Mapping, interviews and bodystorming."),
            ("Define tools", "Point of View (POV) statements, 'How Might We' questions, Why-How laddering and the Business Model Canvas."),
            ("Ideate tools", "Brainstorming, mind mapping, sketchstorm, analogies and gamestorming; then dot voting or a Now-Wow-How matrix to select."),
            ("Prototype tools", "Low- and high-fidelity prototypes: paper sketches, storyboards, wireframes, cardboard and clay models."),
            ("Test tools", "User test scripts, feedback grids, evaluation matrices and A/B comparison of alternative prototypes."),
            ("Measuring outcomes", "Traditional KPIs, customer feedback (CSAT/NPS), design-thinking activity metrics and business-value/novelty mapping."),
         ]),
]

# ------------------------------------------------------------------ day themes (8 training hours)
DAY_THEMES = {
    1: "Design Thinking Concepts, Applications, Action Phases, Visual Tools & Assessment",
}

# ------------------------------------------------------------------ timetable (single source)
# The Lesson Plan computes its schedule from durations; these strings are the SAME
# window expressed for the deck's Lesson Plan slide, so the two cannot disagree.
DAY_START     = "9:30am"
LUNCH_START   = "12:25pm"
LUNCH_END     = "1:25pm"
DAY_END       = "6:30pm"

# ------------------------------------------------------------------ assessment
# Instruments and durations mirror the approved assessment plan for this course:
#   Practical Performance (PP) — 70 mins
#   Oral Questioning (OQ)      — 20 mins
# The written instrument is ORAL QUESTIONING, not a written SAQ paper.
ASSESSMENT = dict(
    written="Oral Questioning (OQ) — the assessor asks each learner the questions verbally, 20 minutes, open book.",
    practical="Practical Performance (PP) — design-thinking tasks using the course ed-tools, 70 minutes, open book.",
    note="A minimum of 75% attendance is required to be eligible for assessment and funding.",
)
OQ_MINUTES = 20
PP_MINUTES = 70
BRIEFING_MINUTES = 5
ASSESSMENT_MINUTES = OQ_MINUTES + PP_MINUTES + BRIEFING_MINUTES   # 95

# ------------------------------------------------------------------ ROI / business value data (from master deck)
ROI_STATS = [
    ("41%", "higher market share"),
    ("46%", "competitive advantage overall"),
    ("50%", "more loyal customers"),
    ("70%", "digital experiences beat competitors"),
]
ROI_SOURCE = "Source: Design Management Institute — dmi.org/page/2015DVIandOTW"
