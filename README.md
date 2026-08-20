# CHEDS Import Tool (SIUD)

Converts raw HEDB-format data exports into Zoho-Creator-ready import
sheets, using per-module mapping configs verified field-by-field against
`Zoho_To_CHEDS.ds` and `HEDB_.xlsx`.

## What's here

| File | Purpose |
|---|---|
| `app_streamlit.py` | The web app -- run this for the live deployment |
| `app_cli.py` | Menu-driven terminal version of the same thing |
| `cheds_forms.py` | Registry of which modules have a finished mapping -- add one line per new module |
| `cheds_import_tool.py` | The actual transform engine (`build`, `report`, `auto-map-ds`, etc.) |
| `mappings/*.json` | One mapping config per CHEDS module, hand-verified |
| `HEDB_.xlsx` | The institution's data dictionary + lookup value lists -- must stay in this folder |

**Currently done: 15 of ~44 modules.** See `cheds_forms.py`'s `FORMS` dict for the exact list.

Note on `Students_Enrollments`: this mapping was built from the form's field
definitions in `Zoho_To_CHEDS.ds` but the actual push/Deluge script for this
module hasn't been reviewed yet (unlike Applicants/Employee, where the push
script was checked directly). Worth a quick sanity pass once that script is
available -- see the mapping file's `_form_note` for details.

## Running locally

```bash
pip install -r requirements.txt
streamlit run app_streamlit.py
```

## First-time GitHub setup

```bash
cd cheds_repo               # this folder
git init
git add .
git commit -m "Initial commit: 14 verified CHEDS mappings"
git branch -M main
git remote add origin <your-repo-url>
git push -u origin main
```

## Deploying (Streamlit Community Cloud)

1. Go to share.streamlit.io, sign in, "New app".
2. Point it at this repo, branch `main`, main file `app_streamlit.py`.
3. Deploy. It reinstalls from `requirements.txt` automatically.
4. **Access control**: since this handles institution-specific and
   Emirates-ID-adjacent data, set the app to restrict viewers (Streamlit
   Cloud: App settings -> Sharing -> restrict by email) rather than leaving
   it fully public.
5. Every future `git push` to `main` auto-redeploys -- no manual restart needed.

## Adding a new module's mapping (the recurring workflow)

1. Get the new `mappings/<Module>_mapping.json` (built and verified with Claude).
2. Drop it into `mappings/`.
3. Add one line to the `FORMS` dict in `cheds_forms.py`.
4. `git add mappings/<file> cheds_forms.py && git commit -m "Add <Module> mapping" && git push`
5. Done -- the live app picks it up within the deploy cycle, no other code changes needed.

## A note on HEDB_.xlsx

`app_streamlit.py` expects `HEDB_.xlsx` to live directly in this folder
(`HEDB_PATH = APP_DIR / "HEDB_.xlsx"`). If CHEDS/MOHESR issues an updated
version of the data dictionary, replace this file and push -- don't rename it.
