"""Topic 4 — Methodologies and Visual Tools for Design Thinking. Labs 5-7."""

DOMAIN4 = [
    dict(
        num=5, topic=4,
        title="Empathy Map & Persona — The Gift-Giving Experience",
        objective="LO2 / A1, A5, K7 — build an empathy map and persona to capture what a user says, thinks, does and feels",
        desc=("Redesign the gift-giving experience for a partner. You start by building empathy: interview "
              "them about the last gift they gave, then plot what they said, thought, did and felt into a "
              "four-quadrant empathy map and distil it into a persona."),
        build="A completed four-quadrant empathy map and a one-page persona for your partner.",
        services="Design Thinking Toolkit — Empathy Map & Persona canvases (alfredang.github.io/designthinking)",
        edtool="https://alfredang.github.io/designthinking/",
        steps=[
            ("Open https://alfredang.github.io/designthinking/ and select the Empathy Map canvas.", ""),
            ("Interview your partner for 5 minutes about the last gift they gave: what they did, how they chose it, and how it felt.", ""),
            ("SAYS — record direct quotes, verbatim. Do not paraphrase; the exact words carry the insight.", ""),
            ("THINKS — record what they were thinking but did not say out loud, such as worries about cost or judgement.", ""),
            ("DOES — record the observable actions: where they shopped, how long it took, who they asked.", ""),
            ("FEELS — record the emotions and their intensity: anxious, rushed, proud, relieved.", ""),
            ("Look for contradictions between quadrants — a positive quote next to a negative feeling is where the real insight hides.", ""),
            ("Switch to the Persona canvas and distil the map into a named persona with goals, frustrations and one defining quote.", ""),
            ("Share your empathy map with your partner and ask them to correct anything you got wrong.", ""),
        ],
        test=("All four quadrants are populated with your partner's own words, you have identified at least one "
              "contradiction between quadrants, and your partner confirms the persona sounds like them."),
    ),
    dict(
        num=6, topic=4,
        title="POV, How Might We & Brainstorming",
        objective="LO2, LO3 / A1, A2, A3, A4, K7 — convert empathy findings into a POV problem statement, generate a high volume of ideas and select the strongest",
        desc=("Take the empathy map from the previous lab and convert it into a sharp Point of View problem "
              "statement, reframe it as a set of 'How Might We' questions, then run a timed brainstorm and "
              "converge on the best ideas using dot voting."),
        build="A validated POV statement, three HMW questions, 20+ generated ideas and a shortlist of three.",
        services="Design Thinking Toolkit — POV, HMW and Brainstorm canvases (alfredang.github.io/designthinking)",
        edtool="https://alfredang.github.io/designthinking/",
        steps=[
            ("Open the POV canvas at https://alfredang.github.io/designthinking/.", ""),
            ("Write your POV using the template: [USER] needs a way to [USER'S NEED] because [INSIGHT].", ""),
            ("Check the quality: the need must be a verb, and the insight must be a discovery — not a restatement of the need.", ""),
            ("List any underlying assumptions you are making, so they can be tested later.", ""),
            ("Reframe the POV into three 'How Might We' questions — broad enough to allow many answers, narrow enough to give direction.", ""),
            ("Warm-up brainstorm: in 2 minutes, list as many creative uses for a paperclip as you can. Aim for quantity, defer judgement.", ""),
            ("Run the real brainstorm on your best HMW for 5 minutes. Rules: go for volume, build on others' ideas, one conversation at a time, no criticism, encourage wild ideas.", ""),
            ("Target at least 20 ideas. Remember: 10 ideas beat 3, and 200 beat 50.", ""),
            ("Converge — use dot voting or a Now-Wow-How matrix to select the three strongest ideas.", ""),
            ("Record why each shortlisted idea was chosen, so the rationale survives to the prototype stage.", ""),
        ],
        test=("Your POV has a verb-based need and a genuine insight, you generated 20 or more ideas, and you "
              "have three shortlisted ideas each with a stated reason for selection."),
    ),
    dict(
        num=7, topic=4,
        title="Prototype, Test & Measure the Outcome",
        objective="LO5, LO6 / A3, A5, A6, K7 — build a low-fidelity prototype, test it with a user, define the metrics that prove it worked, and communicate the outcome and its value to stakeholders",
        desc=("Turn your shortlisted idea into a low-fidelity prototype, test it with a real user, then define "
              "how you would measure whether it succeeded — moving from 'we like this idea' to 'here is the "
              "evidence it works'. Finish by handing the concept over for delivery with a RACI matrix and a "
              "sprint plan."),
        build="A tested prototype, a completed evaluation matrix, an ROI/metrics plan and a RACI handover.",
        services="Design Thinking Toolkit (alfredang.github.io/designthinking), RACI Matrix (alfredang.github.io/raci), Scrum Planner (alfredang.github.io/scrum), BCM (alfredang.github.io/bcm)",
        edtool="https://alfredang.github.io/designthinking/",
        steps=[
            ("Open the Prototype canvas at https://alfredang.github.io/designthinking/ and pick your top idea from Lab 6.", ""),
            ("Decide the fidelity: low-fidelity (sketch, storyboard, paper model) is correct at this stage. Just start building — do not polish.", ""),
            ("Build the prototype in 10 minutes. Remember what you are testing for, and build with the user in mind.", ""),
            ("Write down your top assumption — the one thing that must be true for this idea to work.", ""),
            ("Design the cheapest, fastest test that could disprove that assumption. Define where, when and with whom you will run it.", ""),
            ("Test with a partner: show, don't tell. Let them use it, ask them to think aloud, and note every point of confusion.", ""),
            ("Record feedback in an evaluation matrix, comparing your prototype against at least one alternative so the user can express a preference.", ""),
            ("Define your success metrics across four categories: traditional KPIs, customer feedback (CSAT/NPS), design-thinking activity metrics and quick results.", ""),
            ("Plot your concept on the business-value versus novelty grid — valuable and novel is the target quadrant.", ""),
            ("Open https://alfredang.github.io/raci/ and assign Responsible, Accountable, Consulted and Informed roles for taking the concept forward.", ""),
            ("Open https://alfredang.github.io/scrum/ and convert the concept into a prioritised backlog with a first sprint goal.", ""),
            ("Open https://alfredang.github.io/bcm/ and note one operational risk the solution introduces, plus its mitigation.", ""),
            ("COMMUNICATE THE OUTCOME — prepare a 2-minute pitch to your stakeholder (the budget holder). State the original problem, the insight you discovered, the prototype, what the test showed, and the metrics that will prove success.", ""),
            ("Lead with the insight, not the solution — a stakeholder funds a problem they now understand. Contrast your evidence with the obvious solution they would otherwise have paid for.", ""),
            ("Deliver the pitch to your partner. Ask them: would you fund this? If not, what evidence is missing? Note their answer — that gap is your next test.", ""),
        ],
        test=("Your prototype was changed by user feedback, you have named the assumption you tested and the "
              "metrics that would prove success, the concept has an owner in the RACI matrix and a first "
              "sprint goal, and you have pitched the outcome and its value to your stakeholder leading with "
              "the insight rather than the solution."),
    ),
]
