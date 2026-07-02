# Vision Report Transformer · Report Analyzer

A zero-backend web app + small Python tooling to transform, summarize and export equipment rental reports. Pick a report type in the browser, upload the raw logsheet (Excel/.xls/.csv), preview interactive KPIs & charts, and download a clean, styled Excel export. All browser processing runs locally; server-side features (optional) use Supabase for lightweight auth / uploads.

Key goals:
- Produce a formula-accurate Summary report that mirrors a reference workbook.
- Keep UX simple: pick report → upload → preview → export.
- Provide small CLI tools to prepare and manipulate reports offline.

---

## Features

- Frontend (index.html)
  - Select report type (Rental / Stock).
  - Drag & drop or choose Excel (.xlsx / .xls) files.
  - Interactive KPIs, charts and table rendered in-browser (Chart.js + SheetJS).
  - Styled Excel export (ExcelJS), PDF/CSV export.
  - Optional integration with Supabase to store recent uploads and lightweight RBAC.

- Python utilities
  - build_summary_report.py — regenerate the two-sheet workbook ("Master" + "Summery report") from raw Master logsheet while preserving formulas, styles, formats and injecting cached values for dynamic-array formulas.
  - merge_report.py — collapse split deployment rows into the 21-column Summery report layout.
  - combine_deployments.py — merge rows that represent the same deployment across months into a single continuous block.
  - etl_transform.py — ETL that converts raw daily logs into the monthly PNM template format, preserving template formulas.

- Supabase helpers
  - Database setup SQL (supabase_setup.sql) to create users, uploads, report history, RBAC, and helper functions for registering/logging in, listing uploads, and admin functions.

---

## Stack

- Languages: HTML/CSS/JS (frontend), Python 3 (utilities)
- Frontend: Plain static site (index.html) using:
  - SheetJS (XLSX) to read Excel in-browser
  - ExcelJS to write styled Excel exports
  - Chart.js for charts
  - html2canvas + jsPDF for PDF export
  - Supabase JS client (optional)
- Python libs (used by CLI scripts):
  - pandas
  - openpyxl
- Utilities: LibreOffice (headless) is used by scripts to convert legacy .xls to .xlsx when required.

---

## Repository layout

Top-level tree (important files only):

```
README.md                         <- (this file)
MERGE-README.md
SUMMARY-REPORT-README.md
SETUP-SUPABASE.md                 <- supabase deployment notes
index.html                        <- single-file frontend app (UI + Supabase config)
images/                           <- logo / icons used by the frontend
vercel.json                       <- Vercel static deployment config
supabase_setup.sql                <- DB schema + functions for Supabase
build_summary_report.py           <- build Summary report (preserve formulas + inject cached values)
merge_report.py                   <- collapse split rows into Summery report layout
combine_deployments.py            <- combine deployments across months
etl_transform.py                  <- ETL from raw dump -> PNM template
```

How it fits together:
- index.html is a self-contained static UI that reads Excel workbooks in the browser, runs a chosen processor (implemented inside the JS), and renders KPIs/charts/tables. When enabled, uploads and history use Supabase.
- The Python scripts are CLI tools for batch or offline processing and for producing the same Summery report layout programmatically. build_summary_report.py is careful to preserve Excel styling and formulas and injects cached values for compatibility across viewers.

---

## Quick start — Browser (fastest)

1. Clone or download this repository.
2. Open index.html in a modern browser (double-click or serve it from a static server).
3. On the page:
   - Choose "Rental Report" or "Stock Report".
   - Upload your Excel file (.xlsx). For large legacy .xls files, open in Excel and Save As .xlsx first (recommended).
   - Explore KPIs / charts, then use the Export button to download Excel / CSV / PDF.

Note: index.html contains a SUPABASE_URL and SUPABASE_ANON key (used only if you enable the Supabase-backed features). Edit the COMPANY block in index.html to match your organization before public deployment.

---

## Quick start — CLI (Python)

Prereqs:
- Python 3.8+
- pip install pandas openpyxl
- LibreOffice (optional; required for converting .xls to .xlsx headlessly)

Examples:

- Build the Summary report (creates Master + Summery report from raw Master)
```bash
python build_summary_report.py <new_master.xlsx|.xls|.csv> [reference.xlsx] [output.xlsx]
# example
python build_summary_report.py may_logsheet.xlsx final_fine.xlsx Summary-Report-Output.xlsx
```

- Merge split deployment rows into the 21-column Summery layout:
```bash
python merge_report.py <input.xlsx|.csv|.tsv> [output.xlsx]
# example
python merge_report.py raw-report.xlsx Merged-Summary-Report.xlsx
```

- Combine deployments across months (collapse same deployment identity into 1 row):
```bash
python combine_deployments.py <input.xlsx|.csv|.tsv> [output.xlsx]
```

- Run the ETL to convert a raw dump into the PNM template (preserves template formulas):
```bash
python etl_transform.py <dump.xls|.xlsx> <TEMPLATE-PNM.xlsx> <output.xlsx>
```

Scripts notes:
- The Python scripts expect certain column names in the input. See SUMMARY-REPORT-README.md for examples and expected column names.
- For legacy .xls inputs, the scripts call LibreOffice headless to convert to .xlsx. If you cannot install LibreOffice, convert .xls to .xlsx manually before running.

---

## Supabase (optional server-side features)

To enable uploads, recent-history and simple RBAC backed by Supabase:

1. Create a Supabase project.
2. In Supabase SQL editor, run the contents of `supabase_setup.sql`. This sets up:
   - app_users table with registration/login functions
   - uploads and report_history tables
   - helper functions (register_user, login_user, whoami, add_upload, list_uploads, delete_upload, admin_* helpers)
   - a public Storage bucket `report-files` (public-read by default)
3. Update the SUPABASE_URL and SUPABASE_ANON variables in `index.html` if you used a custom project.
4. Deploy the static site (see below). The frontend will call the Supabase functions to list/upload files and manage user sessions.

Security notes:
- The frontend uses the Supabase anon key for public client functionality only (designed to be safe for the described flows). Do NOT expose service_role keys in the client.
- The SQL defines owner-scoped policies: users can delete only their own uploads; admin functions require an approved admin.

See SETUP-SUPABASE.md for additional operational notes.

---

## Deploying (Vercel or any static host)

Vercel (static, no build):
- Install/vercel or import the repo at vercel.com:
  - Framework: Other
  - Build command: none
  - Output directory: .
Vercel will serve index.html as a static site. vercel.json is included for recommended headers / settings.

Other hosts:
- Upload index.html and the images/ folder to any static host (S3 + CloudFront, Netlify, GitHub Pages, etc.).

Before public deployment:
- Edit the COMPANY block near the top of index.html to set company name, address, email, website, LinkedIn, and APP_VERSION.

---

## Development notes

- UI logic is all in index.html (single-file). To extend report types, add a new processor object to the REPORTS registry in the frontend JS (processor must implement process(workbook), kpis(view), charts(view), columns, tableColumns, filters, searchKeys, exportXlsx(view)).
- The Python scripts are intentionally conservative: they preserve template formulas & styles and inject cached values for formulas that headless LibreOffice cannot evaluate.
- For large legacy .xls files, Save As .xlsx in Excel for best results before uploading (some very large .xls files may not be read reliably).

---

## Contributing

- Open an issue describing the change or enhancement.
- For code changes, fork, make a branch, and open a PR with tests or a short demonstration (screencast / sample input + output).
- If you add behaviour that requires backend changes (DB schema or storage), update SETUP-SUPABASE.md and supabase_setup.sql accordingly.

---

## License & authors

- Add a LICENSE file to this repo (none included by default). If you want me to add a permissive default (MIT) or another license, say which and I can prepare the file.
- Author: Vision Infra / repository maintainer

---

## Troubleshooting / FAQ

- Large .xls won't upload or parse? Save as .xlsx first in Excel (or use LibreOffice to convert).
- Excel formulas show errors in LibreOffice? The scripts compute cached values and inject them so Excel viewers still display correct numbers; Excel will recalculate on open.
- Want per-file private storage? Switch the Supabase bucket to private and add an Edge Function to return short-lived signed URLs after validating session tokens.

---

## Try asking
- How do I add a new report type in the frontend so it appears in the dashboard?
- Can you add a CLI command that runs the full pipeline: ETL -> combine -> build_summary_report -> upload to Supabase?
- Can you create a small sample dataset and a recorded example (input .xlsx → exported Summery-Report-Output.xlsx) to use for end-to-end tests?
