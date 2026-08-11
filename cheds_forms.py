"""
cheds_forms.py
The registry of the 10 CHEDS forms this app knows how to build an
import-ready sheet for. Shared by both app_streamlit.py (web app) and
app_cli.py (menu-driven script) so the list only has to be maintained once.

Each entry: display name (shown in the dropdown/menu) -> mapping config
JSON path, relative to this file's folder (see mappings/).

TO ADD AN 11TH FORM LATER:
  1. Build/verify its mapping config the usual way (auto-map-ds + build,
     or by hand) and drop the finished .json into mappings/.
  2. Add one line below.
That's the only code change needed -- both apps pick it up automatically.
"""

FORMS = {
    "1. Program Learning Outcomes":  "mappings/Program_learning_outcomes_mapping.json",
    "2. Course Learning Outcomes":   "mappings/Course_Learning_outcomes_mapping.json",
    "3. Program Skills":             "mappings/Program_Skills_mapping.json",
    "4. Students - Research":        "mappings/Student_Research_mapping.json",
    "5. Research Impact":            "mappings/Research_Impact_mapping.json",
    "6. Institute - R&D / GERD":     "mappings/Institute_R_D_mapping.json",
    "7. Students - SOD Applicants":  "mappings/Students_SOD_Applicants_mapping.json",
    "8. Institute - Financials":     "mappings/Institute_Financials_mapping.json",
    "9. Graduate Licensure":         "mappings/Graduate_Licensure_mapping.json",
    "10. Graduates":                 "mappings/Student_Graduates_mapping.json",
    "11. Institute - Academic Programs": "mappings/Institute_Academic_Program_mapping.json",
}

# Forms known to have NO live push-to-CHEDS workflow in the .ds export yet
# (see each mapping's own "_form_note" for details) -- both apps show a
# heads-up for these rather than pretending they're fully verified.
UNVERIFIED_NO_WORKFLOW = {
    "6. Institute - R&D / GERD",
    "8. Institute - Financials",
}
