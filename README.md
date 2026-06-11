# Vision Infra — Report Analyzer (multi-report)

A zero-backend web app. Pick a report type, upload its Excel, get an interactive
dashboard and a clean Excel export. All processing runs locally in the browser.

## Workflow
1. Open the app → choose a report type (Rental Report / Stock Report).
2. Upload the Excel file for that report.
3. The matching processor runs and renders KPIs, charts and a table.
4. Download the Excel report (rules specific to that report type), or CSV / PDF.

## Architecture (easy to extend)
All report logic lives in the `REPORTS` registry. Each processor is a self-contained
object: `process(workbook)`, `kpis(view)`, `charts(view)`, `columns`, `tableColumns`,
`filters`, `searchKeys`, `exportXlsx(view)`. The shell is report-agnostic — to add a
new report type, register one more object in `REPORTS`; nothing else changes.

- **Rental Report** — aggregates the daily logsheet by Asset × Site × Month.
- **Stock Report** — generic preview/passthrough today; plug custom rules into its
  processor when the spec is ready.

## Rental — calculations (validated against the source Summery report)
- Equipment Count = unique assets (replaces "Deployments").
- Total Days / Free / Transit / Breakdown / Idle / **Working Days** = distinct dates per status.
- Deployed = Total − Free − Transit ; Billing = Idle + Working.
- **Worked Hours** = SUM of ALL running hours (every status) — matches SUMIFS in the sheet.
- **Avg Utilisation** = Σ Worked Hours ÷ Σ Available Hours, where Available = (260/30) × Billing Days.

Validated on the sample: Equipment Count 484, Worked Hours 116,406, Working Days 7,933,
Avg Utilisation 107.97%, with 0 per-row mismatches vs the Summery report sheet.

## Rental — Excel export columns (exactly these 21, in order)
Sr No, Asset Code, Make, Model, YOM, Department, Site Code, Client Name, Location,
Month/Year, Start Date, Close Date, Total Days, Free Days, In Transit Days,
Deployed Days, Breakdown Days, Idle Days, Working Days, Billing Days, Worked Hours.

## Deploy to Vercel (static, no build)
`npm i -g vercel` then `vercel` in this folder, or import the repo (framework: Other,
no build command, output dir `.`). Vercel serves index.html automatically.

## .xls note
Large legacy .xls files read best if you Save As .xlsx in Excel first.
