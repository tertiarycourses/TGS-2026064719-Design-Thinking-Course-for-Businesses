#!/usr/bin/env bash
# Single-command aligned build of the WSQ courseware from the single source
# (.claude/skills/courseware-build/build/course_data.py + data_domainN.py).
#
# Produces in courseware/: the slide deck PPT + PDF, the Lesson Plan and Learner
# Guide as DOCX + PDF (with page-numbered Tables of Contents), plus the labs/
# activity files and the LG Markdown mirror.
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
export COURSE_REPO="$HERE"
B="$HERE/.claude/skills/courseware-build/build"
CW="$HERE/courseware"
SOFFICE="${SOFFICE:-soffice}"

TITLE="Design Thinking Course for Businesses"
VER="$(python3 -c "import sys;sys.path.insert(0,'$B');import course_data as C;print(C.VERSION)")"

echo "==> Generate labs, PPT, LP and LG from the single source"
python3 "$B/build_labs.py"
python3 "$B/build_slides.py"
python3 "$B/build_lesson_plan.py"     # reads the built deck for slide references
python3 "$B/build_learner_guide.py"

PPT="$CW/$TITLE-$VER.pptx"
LP="$CW/LP-$TITLE.docx"
LG="$CW/LG-$TITLE.docx"

echo "==> Render PDFs (pass 1)"
"$SOFFICE" --headless --convert-to pdf --outdir "$CW" "$PPT" >/dev/null 2>&1
"$SOFFICE" --headless --convert-to pdf --outdir "$CW" "$LP"  >/dev/null 2>&1
"$SOFFICE" --headless --convert-to pdf --outdir "$CW" "$LG"  >/dev/null 2>&1

echo "==> Inject page-numbered Table of Contents (LP + LG)"
python3 "$B/inject_toc.py" "$LP" "${LP%.docx}.pdf" 2
python3 "$B/inject_toc.py" "$LG" "${LG%.docx}.pdf" 2

echo "==> Render PDFs (pass 2 — with built TOC)"
"$SOFFICE" --headless --convert-to pdf --outdir "$CW" "$LP" >/dev/null 2>&1
"$SOFFICE" --headless --convert-to pdf --outdir "$CW" "$LG" >/dev/null 2>&1

# The LP embeds slide references read from the deck; rebuild it once more so the
# references reflect the final deck, then re-render.
echo "==> Reconcile LP slide references against the final deck"
python3 "$B/build_lesson_plan.py"
"$SOFFICE" --headless --convert-to pdf --outdir "$CW" "$LP" >/dev/null 2>&1
python3 "$B/inject_toc.py" "$LP" "${LP%.docx}.pdf" 2
"$SOFFICE" --headless --convert-to pdf --outdir "$CW" "$LP" >/dev/null 2>&1

echo "==> Done. Artifacts in courseware/:"
ls -1 "$CW"/*.pptx "$CW"/*.docx "$CW"/*.pdf
