# Summary Report Generator — automation

Regenerates the two-sheet workbook (`Master` + `Summery report`) from a new raw
logsheet, inheriting the **exact** formulas, layout, styles, column widths, number
formats, conditional formatting and auto-filter of the reference `final_fine.xlsx`.

## Run it
```bash
python build_summary_report.py <new_master.xlsx|.xls|.csv> [reference.xlsx] [output.xlsx]
# example:
python build_summary_report.py may_logsheet.xlsx final_fine.xlsx Summary-Report-Output.xlsx
```
- **new_master** – your fresh raw data. Must contain the Master columns
  (`equipmentnumber`, `sitecode`, `date`, `status`, `running_hrs`, `make`, `model`,
  `fleettype`, `party_name`, `sitelocation`, …).
- **reference** – the template whose look & formulas are inherited (default `final_fine.xlsx`).
- **output** – the file to write.

## What it produces
- **Master** sheet = your new raw data (styled like the reference).
- **Summery report** sheet = one row per distinct **(Asset Code × Site Code)**, fully
  formula-driven off `Master`. The formulas written are byte-for-byte the reference's:

| Col | Formula |
|----|---------|
| A Sr No | `=ROW()-1` |
| C/D/F Make/Model/Dept | `=IFERROR(INDEX(Master!I:I,MATCH($B2,Master!G:G,0)),"")` (I/J/L) |
| J Month/Year | `=TEXT(K2,"mmm-yyyy")` |
| K/L Start/Close | `=IFERROR(MINIFS/MAXIFS(Master!B:B,Master!G:G,$B2,Master!E:E,$G2),"")` |
| M,N,O,Q,R,S day counts | `=IFERROR(ROWS(UNIQUE(FILTER(Master!$B$2:$B$N,(asset)*(site)*(status)))),0)` |
| P Deployed | `=M2-N2-O2` |
| T Billing | `=R2+S2` |
| U Worked Hours | `=SUMIFS(Master!T:T,Master!G:G,$B2,Master!E:E,$G2)` |

Click any of those cells in Excel and the formula bar shows the formula. Because the
sheet references `Master`, editing the raw data and recalculating updates everything.

## Two technical notes
1. **No pivot tables.** Despite the brief, the reference file contains **no pivot tables,
   charts or Excel Tables** — the summary is built entirely with worksheet formulas
   (verified by inspecting the file's internal parts). So there are no pivots to
   replicate; the formula engine above *is* the dynamic summary. If you'd like a real
   PivotTable added (e.g. worked-hours by department/client), that can be built on top.
2. **Cached values are injected.** The day-count columns use Excel-365 dynamic-array
   functions (`UNIQUE`/`FILTER`) that headless LibreOffice can't evaluate. To keep the
   file correct in *every* viewer, the script computes the values in Python and writes
   them as cached results **alongside** the formulas (validated to match the reference
   to the cent across all 582 rows). Excel still recalculates live on open.
