#!/usr/bin/env python3
"""Generate labs/lab-NN-*.md and labs/README.md from the same single source
(course_data.py + data_domainN.py) so the activity files never drift from the
deck, Lesson Plan and Learner Guide."""
import os, sys, re
HERE=os.path.dirname(os.path.abspath(__file__)); sys.path.insert(0,HERE)
import course_data as C
from data_domain1 import DOMAIN1; from data_domain2 import DOMAIN2
from data_domain3 import DOMAIN3; from data_domain4 import DOMAIN4
ACT=DOMAIN1+DOMAIN2+DOMAIN3+DOMAIN4
TOPICS={t["num"]:t for t in C.TOPICS}

def _find_repo(start):
    env=os.environ.get("COURSE_REPO")
    if env and os.path.isdir(env): return env
    d=start
    for _ in range(8):
        d=os.path.dirname(d)
        if os.path.isdir(os.path.join(d,"courseware")) and os.path.isdir(os.path.join(d,"labs")): return d
    return os.path.dirname(os.path.dirname(HERE))
REPO=_find_repo(HERE)
LABS=os.path.join(REPO,"labs"); os.makedirs(LABS,exist_ok=True)

def slug(t):
    s=re.sub(r"[^a-z0-9]+","-",t.lower()).strip("-")
    return re.sub(r"-+","-",s)

for old in os.listdir(LABS):
    if old.startswith("lab-") and old.endswith(".md"): os.remove(os.path.join(LABS,old))

for a in ACT:
    tp=TOPICS[a["topic"]]
    fn=f"lab-{a['num']:02d}-{slug(a['title'])}.md"
    out=[]
    out.append(f"# Activity {a['num']} — {a['title']}")
    out.append("")
    out.append(f"**Course:** {C.TITLE} ({C.COURSE_CODE})  ")
    out.append(f"**Topic {tp['code']}:** {tp['title']}  ")
    out.append(f"**Objective:** {a['objective']}  ")
    out.append(f"**Ed-tool:** {a.get('edtool','—')}")
    out.append("")
    out.append("## Goal")
    out.append("")
    out.append(a["desc"])
    out.append("")
    out.append("## What you'll produce")
    out.append("")
    out.append(a["build"])
    out.append("")
    out.append(f"**Tools:** {a['services']}")
    out.append("")
    out.append("## Steps")
    out.append("")
    for i,(instr,cmd) in enumerate(a["steps"],1):
        out.append(f"{i}. {instr}")
        if cmd: out += ["", "   ```", f"   {cmd}", "   ```", ""]
    out.append("")
    out.append("## You're done when")
    out.append("")
    out.append(a["test"])
    out.append("")
    out.append("## What you learned")
    out.append("")
    for c,_d in tp["concepts"][:3]:
        out.append(f"- {c}")
    out.append("")
    out.append("---")
    out.append("")
    out.append(f"© 2026 {C.ORG}. All rights reserved.")
    with open(os.path.join(LABS,fn),"w") as f: f.write("\n".join(out))
    print("Saved",fn)

# ---- index
idx=[f"# {C.TITLE} — Hands-On Activities",""]
idx.append(f"**WSQ Course Code:** {C.COURSE_CODE}  |  **TSC:** {C.TSC_TITLE} ({C.TSC_CODE})  |  **Version {C.VERSION}**")
idx.append("")
idx.append(f"{len(ACT)} facilitated activities across {len(C.TOPICS)} topics. Every activity runs in the browser — "
           "no software installation is required.")
idx.append("")
idx.append("## Ed-tools used")
idx.append("")
idx.append("| Ed-tool | Link | Used for |")
idx.append("|---|---|---|")
for nm,url,desc in C.EDTOOLS:
    idx.append(f"| {nm} | [{url.replace('https://','')}]({url}) | {desc} |")
idx.append("")
idx.append("## Activities")
idx.append("")
idx.append("| # | Activity | Topic | Ed-tool |")
idx.append("|---|---|---|---|")
for a in ACT:
    fn=f"lab-{a['num']:02d}-{slug(a['title'])}.md"
    tp=TOPICS[a["topic"]]
    idx.append(f"| {a['num']} | [{a['title']}]({fn}) | {tp['code']} — {tp['title']} | [{a.get('edtool','').replace('https://','').rstrip('/')}]({a.get('edtool','')}) |")
idx.append("")
idx.append("---")
idx.append("")
idx.append(f"© 2026 {C.ORG}. All rights reserved.")
with open(os.path.join(LABS,"README.md"),"w") as f: f.write("\n".join(idx))
print("Saved labs/README.md")
