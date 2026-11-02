"""
cheds_forms.py
Registry of the CHEDS forms this app can build an import-ready sheet for.
Shared by app_streamlit.py and app_cli.py.

Each entry: display name -> mapping config JSON path (relative to this folder).
To add a form: drop its .json into mappings/ and add one line below.

2026-11-02 -- HEDB 202601 pass. Every mapping was re-checked field by field
against the HEDB 202601 dictionary ("Database Field Name" column) and the
Zoho_To_CHEDS_ (35).ds export. 19 datasets that previously had no mapping were
added (entries 23-41), and the three mappings that still carried fields the
Ministry removed in 202601 (Graduates, Graduate Licensure, SOD Applicants) were
rebuilt. Institute - Financials was rebuilt as well: its old raw_column names
were Zoho field names (IF_*), not the HEDB export headers (Finance_*).
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
    "17. Students - SOD":                 "mappings/Students_SOD_mapping.json",
    "18. Students - Scholarship":         "mappings/Student_Scholarship_mapping.json",
    "19. Institute - Research & Innovation Unit": "mappings/Institute_Research_Innovation_Unit_mapping.json",
    "20. Institute - Research Projects":  "mappings/Institute_Research_Projects_mapping.json",
    "21. Employee - Workload":            "mappings/Employee_Workload_mapping.json",
    "22. Students - Internship":          "mappings/Student_Internship_mapping.json",
    # --- added 2026-11-02 for HEDB 202601 ---
    "23. Courses":                        "mappings/Institute_Courses_mapping.json",
    "24. Course Faculty":                 "mappings/Institute_Course_Faculty_mapping.json",
    "25. Employee - Salary":              "mappings/Employee_Salary_mapping.json",
    "26. Staff - Elementary Occupation":  "mappings/Staff_Elementary_Occupation_Staff_mapping.json",
    "27. Microcredential - Enrollments":  "mappings/Microcredential_Enrollments_mapping.json",
    "28. Microcredential - Graduates":    "mappings/Microcredential_Graduates_mapping.json",
    "29. OBF Self Report":                "mappings/OBF_Self_Report_mapping.json",
    "30. Institute - Events":             "mappings/Institute_Events_mapping.json",
    "31. Institute - Operations":         "mappings/Institute_Operations_mapping.json",
    "32. Institute - Overview":           "mappings/Institute_Overview_mapping.json",
    "33. Institute - Leadership Contact Information": "mappings/Institute_Leadership_Contact_Information_mapping.json",
    "34. Institute - Partnerships":       "mappings/Institute_Partnerships_mapping.json",
    "35. Institute - Employers":          "mappings/Institute_Employers_mapping.json",
    "36. Institute - Patents":            "mappings/Institute_Patents_mapping.json",
    "37. Institute - Intellectual Property": "mappings/Institute_IP_mapping.json",
    "38. Institute - Publications":       "mappings/Institute_Publications_mapping.json",
    "39. Institute - R&D Classified Expenditures": "mappings/Institute_R_D_Classified_Expenditures_mapping.json",
    "40. Institute - Startups & Spinoffs": "mappings/Institute_Startup_Spinoffs_mapping.json",
}

# Forms with NO live push-to-CHEDS workflow in the .ds export yet --
# both apps show a heads-up for these.
UNVERIFIED_NO_WORKFLOW = {
    # 2026-10-22: "6. Institute - R&D / GERD" removed from this set -- the .ds
    # export (v32) now carries button.Institute_R_D_Function plus its form
    # workflows, so it does have a live push path.
    #
    # 2026-11-02: "8. Institute - Financials" removed as well -- v35 carries
    # button.Institute_Financials_Function / _Delete_Function and the form
    # workflow Push_Api_Production36. Every form in FORMS above now has a
    # production push function in the .ds export, so this set is empty.
}

# Datasets in the HEDB 202601 dictionary that deliberately have NO entry above.
#   Student - Funding, Student - Funding Planning -- confirmed out of scope for SIUD.
#   Students - Licensures                         -- no Zoho form exists yet.
#   Institute - Surveys                           -- the dictionary sheet carries no
#                                                    Database Field Name column, so
#                                                    the export headers are unknown.
OUT_OF_SCOPE = {
    "Student - Funding",
    "Student - Funding Planning",
    "Students - Licensures",
    "Institute - Surveys",
}
