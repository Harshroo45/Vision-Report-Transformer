# Combining split deployment rows

A machine that moves between sites in a month (or whose deployment got logged in
pieces) shows up as several rows. You want ONE row per **Asset × Site × Month**.

There are two tools depending on what you start from.

## 1. summary_from_dump.py  (recommended — regenerates from the raw logsheet)
Produces the correct WORKED HOURS. Your expected numbers (PHALTAN May = 24,
YARD = 2) are the **total running hours of every status**, which only the raw
logsheet has — the older output counted "Working" hours only (22).

```bash
python summary_from_dump.py logsheet-rental-report.xls  Summary-Report-Merged.xlsx
```
Grouping: (Asset Code, Site Code, Month). Per group: START = earliest, CLOSE =
latest, day-columns = distinct dates per status, WORKED HOURS = sum of all
running hours. DEPLOYED (=Total-Free-Transit), Billing (=Idle+Working) and Sr No
are live formulas. Output = the 21-column "Summery report" layout.

## 2. merge_report.py  (when you only have the 32-column output, not the dump)
Collapses duplicate rows already in an output file by SUMMING the day/hour
columns and taking min START / max CLOSE. Use this if your file already has the
split rows with the hours you want. (Note: it sums the WORKED HOURS that are
present, so it will not recover the running-hours-on-a-non-working-day case —
for that, use tool #1.)

```bash
python merge_report.py your_output.xlsx  Merged-Summary-Report.xlsx
```

Both keep Month separate, so PARNER-April and PARNER-May stay as two rows.
