"""
cheds_forms.py
Registry of the CHEDS forms this app can build an import-ready sheet for.
Shared by app_streamlit.py and app_cli.py.

Each entry: display name -> mapping config JSON path (relative to this folder).
To add a form: drop its .json into mappings/ and add one line below.
"""

FORMS = {
    "1. Program Learning Outcomes":       "mappings/Program_learning_outcomes_mapping.json",
    "2. Course Learning Outcomes":        "mappings/Course_Learning_outcomes_mapping.json",
    "3. Program Skills":                  "mappings/Program_Skills_mapping.json",
    "4. Students - Research":             "mappings/Student_Research_mapping.json",
    "5. Research Impact":                 "mappings/Research_Impact_mapping.json",
    "6. Institute - R&D / GERD":          "mappings/Institute_R_D_mapping.json",
    "7. Students - SOD Applicants":       "mappings/Students_SOD_Applicants_mapping.json",
    "8. Institute - Financials":          "mappings/Institute_Financials_mapping.json",
    "9. Graduate Licensure":              "mappings/Graduate_Licensure_mapping.json",
    "10. Graduates":                      "mappings/Student_Graduates_mapping.json",
    "11. Institute - Academic Programs":  "mappings/Institute_Academic_Program_mapping.json",
    "12. Applicants - Basic Details":     "mappings/Applicants_Basic_Details_mapping.json",
    "13. Applicants - Academic Proficiency": "mappings/Applicants_Academic_Proficiency_mapping.json",
    "14. Employee - Basic Details":       "mappings/Employee_Basic_Details_mapping.json",
    "15. Students - Enrollments":         "mappings/Students_Enrollments_mapping.json",
    "16. Students - Attrition":           "mappings/Student_Attrition_mapping.json",
}

# Forms with NO live push-to-CHEDS workflow in the .ds export yet --
# both apps show a heads-up for these.
UNVERIFIED_NO_WORKFLOW = {
    "6. Institute - R&D / GERD",
    "8. Institute - Financials",
}