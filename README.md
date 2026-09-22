# F.A.S.T. Rescue — Course Roster Builder

Automatically generates print-ready Word course rosters from an Excel export. Upload a `.xlsx` file and download filled `.docx` rosters — one per course — with participant names sorted, numbered, and routed to the correct template and columns.

---

## Quick Start

```bash
pip install python-docx openpyxl
python build_rosters.py EXPORT.xlsx --templates templates --out rosters
```

Or run the web app:

```bash
pip install streamlit
streamlit run app.py
```

---

## Templates (7)

| Template | Triggered by course name containing | Layout |
|---|---|---|
| In-Class First Aid | `in-class` or `in class` | Portrait, 4-col |
| Blended First Aid | `blended` | Portrait, 5-col (Online Part 1) |
| Recertification First Aid | `recert` (without blended/in-class) | Portrait, 4-col |
| Working At Heights | `working at heights` | Portrait, 5-col (WAHF/WAHR) |
| Lift Truck | `lift truck` or `forklift` | Landscape, 16-col |
| EWP | `ewp`, `elevated work platform`, or `elevating work platform` | Landscape, 16-col |
| General | anything else (fallback) | Portrait, 4-col |

Template files must have `Template` somewhere in the filename to be detected.

The three First Aid templates have their In-Class and Blended filenames cross-named in the originals. The script identifies them by layout (presence of an "Online Part 1" column and "Manual" in Course Material), not filename.

---

## What It Does Automatically

- Detects course type from the Name column and picks the correct template
- Handles multi-colon company strings by anchoring on the 8-digit `YYYYMMDD` date
- Strips account-number prefixes and suffixes from company names
- Cleans course names — strips trailing city, always ends with `Training`
- Strips company name and contact name from the location field
- Formats contact as `Name – phone`
- Formats dates: single (`May 6, 2026`), two-day (`May 28 & 29, 2026`), multi-session (`May 2, 3 & 5, 2026`)
- Formats times as `8:30 am to 4:00 pm`
- Parses participant lists in any export format (one-per-line, multi-line with emails, bulleted/grouped)
- Strips emails, fixes `Last, First` → `First Last`, collapses double spaces
- Sorts participant names A–Z (optional)
- Adds a **Sr. No.** column (or fills the existing one on landscape templates)
- Sizes the table to `max(20, participants + 5)` rows — always at least 20, always 5+ blank rows
- Sets consistent row heights (`atLeast`) — rows expand only when a name wraps to a second line
- Skips cancelled courses by default
- Title-cases filenames with minor-word exceptions (`of`, `the`, `and`, etc.)
- Strips trailing dots from filenames (`Ltd.` → no double-dot before `.docx`)

---

## Instructor Blanking

If the instructor name contains **Noeline** or **Frank Keegan**, the Instructor field on the roster is left blank and the filename uses `Instructor` as a placeholder.

---

## Same-Day Duplicate Handling

When the same company has multiple courses on the same date, filenames are disambiguated:

| What differs | Suffix added to filename |
|---|---|
| Time slot | `Session 1`, `Session 2`, ... |
| Course name | Course name appended |
| Location | Location appended |
| Nothing (identical) | `Session 1`, `Session 2`, ... |

---

## Working at Heights — 3 Variants

| Course name contains | Participant source | WAHF/WAHR column |
|---|---|---|
| `working at heights` (no recert/refresher) | Participant Names only | All rows = **WAHF** |
| `working at heights` + `recert`/`refresher` (no `full`) | Recertification Participants only | All rows = **WAHR** |
| `working at heights` + `full` + `recert` (mix) | Both columns | **WAHF** from Participant Names, **WAHR** from Recertification Participants |

---

## First Aid + Recert Combos

| Course type | Template used | How recert participants are marked |
|---|---|---|
| Blended + Recert | Blended (5-col) | **Online Part 1** column = `Recert` for recert participants |
| In-Class + Recert | In-Class (4-col) | **Instructor Notes** column = `Recert` for recert participants |
| Recert only | Recert (4-col) | All participants from Recertification column, no special marking |

---

## Output Filename Format

```
Instructor - First date - Company [- Suffix].docx
```

- First letters capitalised, minor words (`of`, `the`, `and`) lowercase mid-name
- Trailing dots stripped
- Missing instructor → `Instructor`
- Missing company → segment omitted
- Suffix added only for same-day duplicates (see above)

---

## CLI Options

| Flag | Description |
|---|---|
| `--templates PATH` | Folder with blank template `.docx` files (default: `.`) |
| `--out PATH` | Output folder (default: `rosters`) |
| `--keep-order` | Preserve export order instead of A–Z |
| `--include-cancelled` | Include cancelled courses |

**Example:**

```bash
python build_rosters.py "May Export.xlsx" --templates templates --out "May Rosters"
```

---

## Excel Export Format

The script expects a single sheet with these column headers (case-insensitive):

| Column | Description |
|---|---|
| Name | Full course entry — account IDs, colons, YYYYMMDD date, course name, city |
| Customer | Company name (may include account prefix/suffix) |
| Course Location | Address (may be prefixed with contact or company name) |
| Training Contact | Contact name (may include account prefix) |
| Training Contact Phone | Phone number(s) |
| Course Date | Start date |
| Course End Date | End date |
| Other Date | Middle session date (for multi-session courses) |
| Time (from) | Start time |
| Time (To) | End time |
| Instructors | Instructor name |
| Participant Names | Newline-separated participant list |
| Recertification Participants | Same format — used for recert/refresher courses |

---

## Repository Structure

```
roster-app/
├── app.py                       # Streamlit web UI
├── build_rosters.py             # Core logic (also runs as CLI)
├── requirements.txt             # Python dependencies
├── README.md
├── templates/                   # Blank Word templates
│   ├── Blended_First_Aid_Template.docx
│   ├── In_Class_First_Aid_Template.docx
│   ├── Recert_First_Aid_Template.docx
│   ├── Working_At_Heights_Template.docx
│   ├── Lift_Truck_Template.docx
│   ├── EWP_Template.docx
│   └── General_Template.docx
├── assets/
│   └── logo.png
└── .streamlit/
    └── config.toml              # Theme and server settings
```

---

## Deploying to Streamlit Community Cloud (Free)

1. Push this repo to GitHub (public for the free tier).
2. Go to [share.streamlit.io](https://share.streamlit.io), sign in with GitHub.
3. Click **Create app → From an existing repo**, select this repo, set main file to `app.py`, deploy.
4. Optional — add a password: in the app's **Settings → Secrets**, add:
   ```toml
   password = "your-password-here"
   ```

The free tier sleeps inactive apps — first visit after a quiet period takes ~30 seconds to wake up.

### Updating

Push changes to GitHub. Streamlit Cloud redeploys automatically within a minute.

---

## Adding a New Template

1. Add the blank `.docx` file to `templates/` — include `Template` in the filename.
2. Add a keyword check in `detect_course_type()` in `build_rosters.py` to return a new type string.
3. Add a filename-keyword match in `classify_template()` to map to the new file.
4. If the template has a pre-existing Sr. No. column or multi-row headers, no extra work — the script detects and handles both automatically.

---

## Dependencies

| Package | Purpose |
|---|---|
| `python-docx` | Reading and writing Word `.docx` files |
| `openpyxl` | Reading Excel `.xlsx` files |
| `streamlit` | Web interface (optional — CLI works without it) |

```bash
pip install python-docx openpyxl streamlit
```
